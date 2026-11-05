#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""symmetric_dup_names.py —— 同一主体多份决定书的文件名「对称化」（2026-11-02）

问题（用户口径）：同一目录下同一主体的多份决定书，例如
    厦门市鑫鼎盛控股有限公司.pdf            ← 无后缀
    厦门市鑫鼎盛控股有限公司（2025-12-02）.pdf  ← 带日期
读者无法判断「哪一份是谁」，而命名规范要求「同名多份必带可区分信息」。

做法：
  1. 扫 01/02/03 分类目录，按 (目录, 主体基名) 分组，取「≥2 个 PDF」的组；
  2. 每个文件用台账「案例文件」列反查所属行 → 取该行「处罚时间」（落款日，权威值）；
  3. 统一改名为 `基名（YYYY-MM-DD）.pdf`；若同组出现同日碰撞，追加措施标签（取自该行「处罚措施」前两段）；
  4. 台账「案例文件」列同步替换；05_重点典型案例 里同名副本一并改名。

用法：
    python scripts/symmetric_dup_names.py --root <库根>            # 干跑，只打印
    python scripts/symmetric_dup_names.py --root <库根> --apply    # 落盘（自动备份台账）

踩坑提示：干跑输出必须人眼过一遍「旧名 → 新名」再看 --apply；不要管道接 head，会 SIGPIPE 掉保存。
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ledger_bak import backup_ledger  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

CAT_COLS = ("01", "02", "03")
TAIL_PAREN = re.compile(r"（[^（）]*）$")
MEASURE_TAGS = ("撤销登记", "公开谴责", "暂停备案", "纪律处分", "行政处罚", "市场禁入",
                "没收违法所得", "罚款", "警示函", "责令改正", "监管谈话", "限制交易",
                "出具警示函", "责令增加内部合规检查次数", "责令定期报告")


def as_date(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v or "").strip()
    m = re.search(r"(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})", s)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return ""


def measure_tag(measures):
    t = str(measures or "")
    hits = [k for k in MEASURE_TAGS if k in t]
    # 「出具警示函」与「警示函」去重，只留更长的一个
    hits = [h for h in hits if h == "出具警示函" or "出具警示函" not in hits]
    if not hits:
        return ""
    return "·".join(sorted(hits, key=len, reverse=True)[:2])


def split_files(cell):
    return [x.strip() for x in re.split(r"[；;]", str(cell or "")) if x.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--apply", action="store_true", help="落盘（默认只干跑）")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    ledger = None
    for f in os.listdir(root):
        if f.startswith("04_") and f.endswith(".xlsx") and not f.startswith("~$"):
            ledger = os.path.join(root, f)
            break
    if not ledger:
        sys.exit("找不到 04_ 台账 xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active
    hdr = [str(c.value or "") for c in ws[1]]
    C_NAME, C_DATE, C_MEAS, C_FILES = 3, 1, 4, 11  # 0-based：D/处罚时间、E/处罚措施、K/案例文件
    assert hdr[C_FILES] == "案例文件" and hdr[C_DATE] == "处罚时间", hdr[:12]

    # 台账索引：文件名 → [(row, date, measures)]
    idx = {}
    for r in range(2, ws.max_row + 1):
        subj = ws.cell(r, C_NAME + 1).value
        if not subj:
            continue
        d = as_date(ws.cell(r, C_DATE + 1).value)
        m = ws.cell(r, C_MEAS + 1).value
        for fn in split_files(ws.cell(r, C_FILES + 1).value):
            idx.setdefault(fn, []).append((r, d, m))

    # 分组扫描
    plan = []   # (目录, 旧名, 新名, 行号)
    skipped = []
    for d1 in sorted(os.listdir(root)):
        if not os.path.isdir(os.path.join(root, d1)) or d1[:2] not in CAT_COLS:
            continue
        for d2 in sorted(os.listdir(os.path.join(root, d1))):
            p2 = os.path.join(root, d1, d2)
            if not os.path.isdir(p2):
                continue
            groups = {}
            for f in sorted(os.listdir(p2)):
                if not f.lower().endswith(".pdf"):
                    continue
                stem = f[:-4]
                base = TAIL_PAREN.sub("", stem) or stem
                groups.setdefault(base, []).append(f)
            for base, fs in groups.items():
                if len(fs) < 2:
                    continue
                # 组内已全部带后缀且互不相同 → 跳过（已经对称）
                if all(TAIL_PAREN.search(f[:-4]) for f in fs) and len({f[:-4] for f in fs}) == len(fs):
                    continue
                cand = []
                for f in fs:
                    hits = idx.get(f, [])
                    rows = {h[0] for h in hits}
                    if len(hits) != 1 or len(rows) != 1:
                        skipped.append((d1, f, "台账反查不唯一：%d 条" % len(hits)))
                        cand = None
                        break
                    r, d, m = hits[0]
                    if not d:
                        skipped.append((d1, f, "台账处罚时间缺失"))
                        cand = None
                        break
                    cand.append((p2, f, base, d, m, r))
                if not cand:
                    continue
                # 同日碰撞 → 追加措施标签
                seen = {}
                for item in cand:
                    seen.setdefault(item[3], []).append(item)
                finals = []
                for d, items in seen.items():
                    if len(items) == 1:
                        finals.append(items[0] + ("%s（%s）" % (items[0][2], d),))
                    else:
                        for it in items:
                            tag = measure_tag(it[4])
                            nm = "%s（%s%s）" % (it[2], d, ("·" + tag) if tag else "")
                            finals.append(it + (nm,))
                names = [x[-1] for x in finals]
                if len(set(names)) != len(names):
                    skipped.append((d1, base, "组内仍重名：%s" % "、".join(names)))
                    continue
                for p2x, f, b, d, m, r, nm in finals:
                    if nm + ".pdf" == f:
                        continue
                    plan.append((p2x, f, nm + ".pdf", r))

    print("待改名 %d 个：" % len(plan))
    for p2, old, new, r in plan:
        print("  行%-4d %s\n          %s\n       → %s" % (r, os.path.basename(p2), old, new))
    if skipped:
        print("\n跳过 %d 个（需人工判）：" % len(skipped))
        for d1, f, why in skipped:
            print("  %s / %s —— %s" % (d1[:6], f[:50], why))
    if not a.apply:
        print("\n[干跑] 未改动任何文件；确认无误后加 --apply")
        return

    backup_ledger(root, ledger, "symname")

    # 05 副本按「旧名 → 新名」映射同步
    map05 = {old: new for _p, old, new, _r in plan}
    typ_root = None
    for d in os.listdir(root):
        if d.startswith("05_"):
            typ_root = os.path.join(root, d)
    synced05 = 0
    if typ_root:
        for dp, _dn, fns in os.walk(typ_root):
            for f in fns:
                if f in map05:
                    src = os.path.join(dp, f)
                    dst = os.path.join(dp, map05[f])
                    if os.path.exists(dst):
                        print("  ! 05 目标已存在，跳过：%s" % dst)
                        continue
                    os.rename(src, dst)
                    synced05 += 1

    done = 0
    for p2, old, new, r in plan:
        src, dst = os.path.join(p2, old), os.path.join(p2, new)
        if not os.path.isfile(src):
            print("  ! 源文件不存在：%s" % src)
            continue
        if os.path.exists(dst):
            print("  ! 目标已存在，跳过：%s" % dst)
            continue
        os.rename(src, dst)
        # 台账 K 列按「；」分项**精确匹配**替换。
        # 2026-11-02 踩坑：整串 str.replace 会把「林劲-厦门市鑫鼎盛控股有限公司.pdf」里
        # 的同名尾段一起改掉（人名文件被误加日期后缀），A1 立刻报孤儿 1 / 缺失 1。
        items = [x.strip() for x in str(ws.cell(r, C_FILES + 1).value or "").split("；")]
        ws.cell(r, C_FILES + 1).value = "；".join(new if x == old else x for x in items)
        done += 1
    wb.save(ledger)
    print("改名 %d 个，05 副本同步 %d 个，台账已保存" % (done, synced05))


if __name__ == "__main__":
    main()
