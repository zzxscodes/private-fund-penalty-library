#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""plan_names.py —— 案例文件改名计划器（2026-11-02）

把「命名规则 R1~R7」编译成一份可审阅的改名计划（old → new + 理由），
默认只打印（dry-run），加 `--apply` 才落盘：
  ① 磁盘改名（同名目标已存在则跳过并报冲突，绝不覆盖）
  ② 台账 K 列按「；」分项**精确匹配**替换（不做 str.replace，避免带改同前缀的其他文件）
  ③ 05 典型副本按 basename 同步（R7：副本名必须与 K 列同名）

命名规则（唯一口径，SKILL.md 0.5.13）：
  R1 机构决定书    `机构全称.pdf`
  R2 人员决定书    `姓名-机构全称.pdf`
  R3 复核双件      `…（原决定）.pdf` / `…（复核决定）.pdf`
  R4 同主体多案    **库内**该主体（机构全称 / 姓名+机构）对应 ≥2 个案例行 → 每份都带 `（落款日）`
                   判据是「案例行数」不是「文件数」：只要同一主体在库里有两案，
                   文件名就必须能自证属于哪一案（浙江优策 4 行即是此例）
  R5 唯一性        同目录不得同名；不得出现「（N）」下载序号
  R6 不多余        库内只有 1 个案子的主体 → 不带日期后缀（复核双件除外）
  R7 副本一致      05 典型副本与 K 列同名

R3 与 R4 冲突时（既复核双件又多案）：R3 优先，语义比日期更重要。

用法：
  python scripts/plan_names.py --root <库根>            # 只看计划
  python scripts/plan_names.py --root <库根> --apply    # 落盘
"""
import argparse
import os
import re
import shutil
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

from subject_clean import strip_role_prefix  # noqa: E402

TAIL = re.compile(r"[（(]([^（）()]*)[）)]$")
DATE = re.compile(r"^20\d{2}-\d{2}-\d{2}$")
SEQ = re.compile(r"^\d+$")
SEM = ("原决定", "复核决定")
K_COL = 12   # L 列 → 1-based 12（案例文件）
L_COL = 13


def split_multi(v):
    if v is None:
        return []
    return [p.strip() for p in re.split(r"[；;]", str(v or "")) if p.strip()]


def stem_tail(fn):
    """(主干, 后缀)。只剥「日期 / 原决定 / 复核决定 / 纯数字序号」。"""
    s = fn[:-4] if fn.lower().endswith(".pdf") else fn
    m = TAIL.search(s)
    if not m:
        return s, ""
    inner = m.group(1).strip()
    if DATE.match(inner) or inner in SEM or SEQ.match(inner):
        return s[:m.start()], inner
    return s, ""


def find_ledger(root):
    return next(os.path.join(root, f) for f in os.listdir(root)
                if f.startswith("04_") and f.endswith(".xlsx") and not f.startswith("~$"))


def load_rows(ws):
    rows = []
    for r in range(2, ws.max_row + 1):
        raw_subj = str(ws.cell(r, 4).value or "").strip()
        if not raw_subj:
            continue
        segs = [x.strip() for x in re.split(r"[；;]", raw_subj) if x.strip()]
        org = strip_role_prefix(segs[0])
        persons = []
        for s in segs[1:]:
            persons += [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]
        rows.append({
            "row": r,
            "decided": str(ws.cell(r, 2).value or "")[:10],
            "org": org,
            "org_raw": segs[0],
            "persons": persons,
            "measures": str(ws.cell(r, 5).value or ""),
            "files": split_multi(ws.cell(r, K_COL).value),
            "dirs": split_multi(ws.cell(r, L_COL).value) or [""],
            "typical": str(ws.cell(r, 15).value or "").strip(),
        })
    return rows


def index(rows):
    org_rows, person_rows = defaultdict(set), defaultdict(set)
    for x in rows:
        org_rows[x["org"]].add(x["row"])
        for p in x["persons"]:
            person_rows[(p, x["org"])].add(x["row"])
    return org_rows, person_rows


def classify_file(fn, rec):
    """→ (kind, subject, tail)。kind ∈ org / person / unknown。"""
    stem, tail = stem_tail(fn)
    org = rec["org"]
    # R8 合并文书（「对你公司、李厚文予以书面警示」这类一份文书同时处理机构与人员）：
    #   整行只有 1 份文书、且措施列有「管理人：」段（机构确实被处罚），却只有人员件命名
    #   → 该件其实是机构+人员合并文书，按机构全称命名。
    #   反例守卫：个人案（如「实际负责X公司」的从业人员、机构未被处罚）措施列**没有**
    #   「管理人：」段，仍按人员件命名，不得误改成机构名。
    merged_org = False
    if len(rec["files"]) == 1 and re.search(r"(?:^|\n)\s*(?:管理人|机构)（?", rec["measures"]):
        merged_org = True
    # 人员件必须整体等于「姓名-机构全称」
    for p in rec["persons"]:
        if stem == "%s-%s" % (p, org):
            return ("org", org, tail) if merged_org else ("person", p, tail)
    if stem == org:
        return "org", org, tail
    # 兜底：机构名被角色前缀污染的旧文件名（「尉亲-曾登记为X公司」「刘凯歌-实际负责X公司」）
    for p in rec["persons"]:
        pre = p + "-"
        if stem.startswith(pre) and strip_role_prefix(stem[len(pre):]) == org:
            return ("org", org, tail) if merged_org else ("person", p, tail)
    if strip_role_prefix(stem) == org:
        return "org", org, tail
    # 兜底：破折号粘连（「寿宁投资管理（上海-寿宁投资管理（上海）有限公司」）
    #   判据=主干以「-机构全称」结尾，前缀不是合法姓名
    if stem.endswith("-" + org) and len(stem) > len(org) + 1:
        return "org", org, tail
    return "unknown", stem, tail


def plan(rows):
    org_rows, person_rows = index(rows)
    items, unknown, keeps = [], [], []
    for rec in rows:
        org = rec["org"]
        # 同一行内同主体文件数 → 判断是否为复核双件
        groups = defaultdict(list)
        for f in rec["files"]:
            kind, subj, tail = classify_file(f, rec)
            if kind == "unknown":
                unknown.append((rec["row"], f, org))
                continue
            groups[(kind, subj)].append((f, tail))

        for (kind, subj), fl in groups.items():
            pair = len(fl) >= 2                       # 复核双件：同主体 ≥2 份
            sem_tails = [t for _f, t in fl if t in SEM]
            multi = (len(org_rows[org]) >= 2) if kind == "org" \
                else (len(person_rows[(subj, org)]) >= 2)
            for f, tail in fl:
                base = org if kind == "org" else "%s-%s" % (subj, org)
                if pair and sem_tails:
                    # R3 优先：保留语义后缀（原决定/复核决定）
                    if tail in SEM:
                        want, why = "%s（%s）" % (base, tail), "R3 复核双件"
                    else:
                        # 同主体多份却只有部分标了语义 → 交给人工，不猜
                        unknown.append((rec["row"], f, "复核双件未标语义"))
                        continue
                elif tail == "复核决定":
                    # 只有复核件、原决定未取得（原决定多在窗口外）——标注是**事实**，必须保留
                    want, why = "%s（复核决定）" % base, "R3 仅存复核件"
                elif tail == "原决定":
                    # 孤立的「原决定」：该主体没有复核件，这个后缀会让人误以为被复核过 → 去掉
                    want, why = base, "R6 孤立原决定去语义"
                elif pair:
                    if tail and DATE.match(tail):
                        want, why = "%s（%s）" % (base, tail), "R5 同案多份留日期"
                    else:
                        want, why = base, "R1/R2 规范化"
                elif multi:                             # R4
                    want, why = "%s（%s）" % (base, rec["decided"]), "R4 同主体多案加日期"
                else:                                   # R6
                    want, why = base, "R6 单案去日期"
                new = want + ".pdf"
                if new == f:
                    keeps.append((rec["row"], f))
                else:
                    items.append((rec["row"], f, new, why))
    return items, unknown, keeps


def apply_plan(root, ws, items, log):
    """磁盘改名 + K 列精确替换 + 05 副本同步。返回 (done, skipped)。"""
    typ_dirs = [os.path.join(root, d) for d in os.listdir(root) if d.startswith("05_")]
    # basename → 05 副本绝对路径
    copies = defaultdict(list)
    for td in typ_dirs:
        for dp, _dn, fns in os.walk(td):
            for fn in fns:
                copies[fn].append(os.path.join(dp, fn))

    dir_of = {}
    for r in range(2, ws.max_row + 1):
        dir_of[r] = split_multi(ws.cell(r, L_COL).value) or [""]

    done, skipped = [], []
    for row, old, new, why in items:
        dirs = dir_of.get(row) or [""]
        if len(dirs) != 1:
            skipped.append((row, old, "L 列路径不唯一：%d 条" % len(dirs)))
            continue
        d = dirs[0]
        if not os.path.isabs(d):
            # 台账「存放路径」列存的是**相对库根**的分类目录（如 02_各地证监局…/02_登记备案…）
            d = os.path.join(root, d)
        sp, dp = os.path.join(d, old), os.path.join(d, new)
        if not os.path.isfile(sp):
            skipped.append((row, old, "源文件不存在"))
            continue
        if os.path.exists(dp):
            skipped.append((row, old, "目标同名已存在"))
            continue
        try:
            os.rename(sp, dp)
        except OSError as e:
            skipped.append((row, old, "rename 失败：%s" % e))
            continue
        # K 列精确替换（按「；」分项）
        toks = split_multi(ws.cell(row, K_COL).value)
        if old in toks:
            toks[toks.index(old)] = new
            ws.cell(row, K_COL).value = "；".join(toks)
        # 05 副本同步（R7）
        for cp in copies.get(old, []):
            cd, cf = os.path.dirname(cp), os.path.join(os.path.dirname(cp), new)
            if os.path.exists(cf):
                log.append("  · 05 副本目标已存在，跳过：%s" % cf)
                continue
            try:
                shutil.move(cp, cf)
                log.append("  · 05 副本同步：%s → %s" % (old, new))
            except OSError as e:
                log.append("  · 05 副本改名失败 %s：%s" % (cp, e))
        done.append((row, old, new, why))
    return done, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    ledger = find_ledger(root)
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    rows = load_rows(ws)
    items, unknown, keeps = plan(rows)
    print("== 改名计划：%d 行 / 建议改名 %d 份 / 已合规 %d 份 / 无法判定 %d 份 =="
          % (len(rows), len(items), len(keeps), len(unknown)))
    bywhy = defaultdict(list)
    for row, old, new, why in items:
        bywhy[why].append((row, old, new))
    for why in sorted(bywhy):
        print("\n【%s】%d 份" % (why, len(bywhy[why])))
        for row, old, new in bywhy[why]:
            print("  行%-4d %s" % (row, old))
            print("       → %s" % new)
    if unknown:
        print("\n【无法判定（需人工）】%d 份" % len(unknown))
        for row, f, hint in unknown:
            print("  行%-4d %s  ｜%s" % (row, f, hint))

    if not a.apply:
        print("\n（dry-run，未改动任何文件；加 --apply 落盘）")
        return
    if not items:
        print("\n无需改动。")
        return
    from ledger_bak import backup_ledger
    b = backup_ledger(root, ledger, "plan_names")
    print("\n台账已备份：%s" % b)
    log = []
    done, skipped = apply_plan(root, ws, items, log)
    wb.save(ledger)
    for line in log:
        print(line)
    print("\n已改名 %d 份，跳过 %d 份" % (len(done), len(skipped)))
    for row, old, r in skipped:
        print("  跳过 行%-4d %s ｜%s" % (row, old, r))


if __name__ == "__main__":
    main()
