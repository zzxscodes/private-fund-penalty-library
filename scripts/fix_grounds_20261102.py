#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_grounds_20261102.py —— F 列（核心处罚事由）统一形态 + 修内容（2026-11-02）

两条根因（都已修在源头，这里负责把存量掰回来）：
  ① 旧抽取器 `(?:经查)[^：:]*[:：]` 吃到「统一社会信用代码：」的冒号，把主体名和
     左括号切掉 → 台账出现「91310109…E ）在管理…」（5 行）。
  ② 部分行的 F 取自证监局**网页摘要**（以「…等问题」结尾，信息被压缩）而不是决定书正文
     （16 行）。行28 就是典型：网页版漏了「损害部分投资者利益」。

处理顺序（逐行）：
  1. 命中「信用代码残片 / 网页导语」→ 读决定书原文重抽（优先）
  2. 整段未分点（>40 字）→ 读原文重抽；原文取不到则按原文顿号拆并列项
  3. 统一走 ledger_style.normalize_grounds()：剥引导句、剥「等问题」尾巴、补句号、
     「你公司」→ 主体名、压空格
  4. 已分点的行也过一遍（补句末标点 121 行、剥尾巴）

保护（宁可不动，不可改错）：
  - 重抽值过短（< 旧值 60%）→ 判为抽取截断，回退旧值
  - 重抽值含「我局决定 / 违反了 / 现决定」→ 判为越界，回退旧值
  - 重抽后分点数少于旧的 → 判为退化，回退旧值
  - 任何跳过的行都打印原因，可逐条复核

用法：
  python scripts/fix_grounds_20261102.py --root <库根>              # 干跑，打印对照
  python scripts/fix_grounds_20261102.py --root <库根> --apply      # 落盘
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import extract_core as EC  # noqa: E402
import ledger_style as LS  # noqa: E402
from audit_cells import check_cell  # noqa: E402
from ledger_bak import backup_ledger  # noqa: E402
from probe_grounds import pdf_text_join, text_layer_ok  # noqa: E402

IDX = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|[（(]\d+[）)])", re.M)
OVERREACH = re.compile(r"我局决定|本局决定|现决定|违反了《|以上事实|上述事实")


def n_items(s):
    return len(IDX.findall(s or ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--show", type=int, default=6)
    ap.add_argument("--only", default="", help="只处理这些行号，逗号分隔")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]
    only = {int(x) for x in args.only.split(",") if x.strip()} if args.only else None

    n_redraw = n_norm = n_skip = n_same = 0
    shown = 0
    skip_reasons = []

    for i in range(2, ws.max_row + 1):
        if not ws.cell(i, 1).value:
            continue
        if only and i not in only:
            continue
        cell = ws.cell(i, 6)
        old = cell.value
        if not isinstance(old, str) or not old.strip():
            continue
        old = old.strip()
        codes = check_cell(5, old, i)
        subject = str(ws.cell(i, 4).value or "")
        source = str(ws.cell(i, 11).value or "")

        need_pdf = (any(c in ("X8信用代码残片", "X22网页导语句", "X23叙述引导词开头", "X7括号不配对",
                              "X25事由越界", "X21疑似多条未拆分", "X26网页外壳混入")
                        for c in codes)
                    or (not IDX.search(old) and len(old) > 40))

        base = old
        redrawn = False
        if need_pdf:
            names = [x.strip() for x in re.split(r"[；;]", str(ws.cell(i, 12).value or "")) if x.strip()]
            paths = [x.strip() for x in re.split(r"[；;]", str(ws.cell(i, 13).value or "")) if x.strip()]
            d = paths[0] if paths else ""
            orgs = [n for n in names if "-" not in n]
            target = (orgs or names or [""])[0]
            p = os.path.join(root, d, target) if target else ""
            if p and os.path.isfile(p) and text_layer_ok(p):
                try:
                    t = pdf_text_join(p)
                    e = EC.extract_amac(t) if source == "中基协" else EC.extract_csrc(t)
                    cand = (e.get("grounds") or "").strip()
                except Exception as ex:
                    cand = ""
                    skip_reasons.append((i, "抽取异常：%s" % ex))
                if cand:
                    reason = ""
                    if len(cand) < 0.6 * len(old):
                        reason = "重抽过短(%d<%d)" % (len(cand), int(0.6 * len(old)))
                    elif OVERREACH.search(cand):
                        reason = "重抽越界"
                    elif n_items(cand) < n_items(old):
                        reason = "分点退化(%d<%d)" % (n_items(cand), n_items(old))
                    if reason:
                        # 旧值本身串味（多条未拆/越界）时，重抽变短是修复而非截断 → 采用
                        if cand and len(cand) >= 60 and any(
                                c in codes for c in ("X21疑似多条未拆分", "X25事由越界")):
                            base = cand
                            redrawn = True
                        else:
                            skip_reasons.append((i, reason + "，保留旧值"))
                    else:
                        base = cand
                        redrawn = True
            else:
                skip_reasons.append((i, "无文字层/文件缺失，按旧值归一"))

        new = LS.normalize_grounds(base, subject, split_parallel=True).strip()

        if not new or len(new) < 12:
            skip_reasons.append((i, "归一后过短，跳过"))
            continue
        if new == old:
            n_same += 1
            continue

        # 变化过大又没有原文支撑 → 拦下（只拦「骤短」，不拦「分点减少」——
        # 越界截断本来就会砍掉多余条目，那是正确行为）
        if (not redrawn and len(old) > 30 and len(new) < 0.45 * len(old)
                and "X8信用代码残片" not in codes and "X25事由越界" not in codes
                and "X21疑似多条未拆分" not in codes and "X26网页外壳混入" not in codes):
            skip_reasons.append((i, "归一后骤短(%d→%d)，跳过" % (len(old), len(new))))
            continue

        if redrawn:
            n_redraw += 1
        else:
            n_norm += 1
        if shown < args.show:
            print("\n行%-4d %s" % (i, subject[:30]))
            print("  形态 %s" % (",".join(codes) or "（仅形态归一）"))
            print("  旧 %s" % old[:120].replace("\n", " ⏎ "))
            print("  新 %s" % new[:120].replace("\n", " ⏎ "))
            shown += 1
        if args.apply:
            cell.value = new

    print("\n改写 %d 行（其中按原文重抽 %d）｜未变 %d｜跳过 %d"
          % (n_redraw + n_norm, n_redraw, n_same, len(skip_reasons)))
    if skip_reasons:
        print("\n【跳过明细】前 20 条")
        for i, r in skip_reasons[:20]:
            print("  行%-4d %s" % (i, r))

    if args.apply:
        backup_ledger(root, ledger, "grounds_20261102")
        wb.save(ledger)
        print("\n已保存（备份在 _案例库维护留痕/台账备份/）")
    else:
        print("\n[干跑] 未落盘；加 --apply 执行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
