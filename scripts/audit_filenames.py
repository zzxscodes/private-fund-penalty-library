#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""audit_filenames.py —— 案例文件命名体检（2026-11-02，规则见 references/ledger-schema.md）

命名规则（唯一口径）：
  R1 机构决定书：`机构全称.pdf`
  R2 人员决定书：`姓名-机构全称.pdf`（姓名必须在台账 D 列人名单内）
  R3 复核双件  ：同一主体同一案内同时有原决定与复核决定 → `…（原决定）.pdf` / `…（复核决定）.pdf`
                 孤立的「（复核决定）」保留——原决定多数在窗口外未取得，标注是**事实**；
                 孤立的「（原决定）」去掉——该主体没有复核件，后缀会让人误以为被复核过。
  R4 同主体多案：该主体在**库内**有 ≥2 个案例（不同落款日）→ 每份文件都带 `（YYYY-MM-DD）`
                 （不依赖同目录：两案可能分属不同分类目录/不同信源）
  R5 唯一性    ：同目录不得同名；不得出现「（N）」下载序号
  R6 不多余    ：该主体库内只有 1 个案例 → 不带日期后缀（复核双件除外）
  R8 合并文书  ：一份文书同时处理机构与人员 → 按**机构全称**命名（措施列有「管理人：」段
                 却只有人员件命名 = 命名错，plan_names.py 会自动纠正）

关键实现细节（踩过的坑）：
  ① 剥尾部括号时**绝不能**剥「（有限合伙）」——那是主体名的一部分，
     否则 63 行会假报「机构件缺失」（实测）。
  ② 判「多案」要看**库内同名机构的行数**，不是同目录文件数；
     否则跨目录的第二案（如中基协 + 证监局各一案）会被判成「多余后缀」（实测 66 条假报）。
  ③ 「人员没有独立文件」绝大多数是**正常**的：一份联名决定书/合并文书的抬头写
     「关于对胡瀚阳、卞桂玲、孙曼采取…」，一份文书覆盖多人。判据是
     **链接数 vs 文件数**：相等 → 正常（E5' 观察项）；链接多 → 真缺件（E5 问题）。
     实测全库 50 条「无独立文件」全部链接数=文件数，0 缺件。

输出问题类别：
  E1 文件名主体不在台账主体列（张冠李戴 / 漏登人员）—— 最严重
  E2 机构件缺失（措施列有「管理人：」却找不到机构件）
  E3 机构件后缀不符规则（该带日期没带 / 不该带却带了 / 复核双件没标语义）
  E4 人员件后缀不符规则
  E5 人员件缺失（链接数 > 文件数，真缺件）
  E5' 联名/合并文书：多人共用 1 份（正常观察，不需处理）
  E6 同一目录同名 / 残留（N）

用法：python scripts/audit_filenames.py --root <库根> [--out report.txt]
"""
import argparse
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

TAIL = re.compile(r"[（(]([^（）()]*)[）)]$")
DATE = re.compile(r"^20\d{2}-\d{2}-\d{2}$")
SEQ = re.compile(r"^\d+$")
SEM = ("原决定", "复核决定")


def split_multi(v):
    if v is None:
        return []
    return [p.strip() for p in re.split(r"[；;]", str(v)) if p.strip()]


def stem_tail(fn):
    """(主干, 后缀)。只剥「日期 / 原决定 / 复核决定 / 纯数字序号」，不碰（有限合伙）等主体名组成部分。"""
    s = fn[:-4] if fn.lower().endswith(".pdf") else fn
    m = TAIL.search(s)
    if not m:
        return s, ""
    inner = m.group(1).strip()
    if DATE.match(inner) or inner in SEM or SEQ.match(inner):
        return s[:m.start()], inner
    return s, ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--out", default="")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    ledger = next(os.path.join(root, f) for f in os.listdir(root)
                  if f.startswith("04_") and f.endswith(".xlsx") and not f.startswith("~$"))
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    rows, org_rows, person_rows = [], defaultdict(set), defaultdict(set)
    for r in range(2, ws.max_row + 1):
        subj = str(ws.cell(r, 4).value or "").strip()
        if not subj:
            continue
        segs = [x.strip() for x in re.split(r"[；;]", subj) if x.strip()]
        org = segs[0]
        persons = []
        for s in segs[1:]:
            persons += [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]
        rec = {"row": r, "org": org, "persons": persons,
               "decided": str(ws.cell(r, 2).value or "")[:10],
               "measures": str(ws.cell(r, 5).value or ""),
               "files": split_multi(ws.cell(r, 12).value),
               "urls": split_multi(ws.cell(r, 14).value),
               "dirs": split_multi(ws.cell(r, 13).value) or [""]}
        rows.append(rec)
        org_rows[org].add(r)
        for p in persons:
            person_rows[(p, org)].add(r)

    E = defaultdict(list)
    for rec in rows:
        r, org = rec["row"], rec["org"]
        multi_case = len(org_rows[org]) >= 2
        stems = {}
        for f in rec["files"]:
            st, tl = stem_tail(f)
            stems[f] = (st, tl)
            if SEQ.match(tl):
                E["E6 残留下载序号"].append((r, f))
            if st == org or any(st == "%s-%s" % (p, org) for p in rec["persons"]):
                continue
            E["E1 文件名主体不在台账"].append((r, f, org[:30]))

        # E2/E3 机构件
        org_files = [(f, tl) for f, (st, tl) in stems.items() if st == org]
        need_org = "管理人：" in rec["measures"] or "机构：" in rec["measures"]
        if not org_files and need_org:
            E["E2 机构件缺失"].append((r, org[:36], "；".join(rec["files"])[:50]))
        tails = [tl for _f, tl in org_files]
        if len(org_files) >= 2 and set(tails) & set(SEM) and len(set(tails) & set(SEM)) != len(org_files):
            E["E3 复核双件未成对标语义"].append((r, "｜".join(f for f, _t in org_files)[:70]))
        elif len(org_files) == 1:
            tl = org_files[0][1]
            if tl == "原决定":
                E["E3 单份却标（原决定）"].append((r, org_files[0][0]))
            elif multi_case and not (DATE.match(tl) or tl in SEM):
                E["E3 同机构多案但机构件无日期"].append((r, org, org_files[0][0]))
            elif not multi_case and DATE.match(tl):
                E["E3 单案机构件多了日期"].append((r, org_files[0][0]))
        elif len(org_files) >= 2 and not set(tails) & set(SEM):
            if len([t for t in tails if DATE.match(t)]) != len(org_files):
                E["E3 同案多份机构件后缀不齐"].append((r, "｜".join(f for f, _t in org_files)[:70]))

        # E4/E5 人员件
        n_url = len(rec["urls"])
        n_file = len(rec["files"])
        for p in rec["persons"]:
            pf = [(f, tl) for f, (st, tl) in stems.items() if st == "%s-%s" % (p, org)]
            if not pf:
                if "高管（%s）" % p in rec["measures"] or "人员：" in rec["measures"]:
                    # 2026-11-02：链接数 > 文件数 才是真缺件；相等说明「一份文书处理多个
                    # 当事人」（联名决定书 / 机构与人员合并文书），属正常，不进问题清单。
                    if n_url > n_file:
                        E["E5 人员件缺失（链接多于文件）"].append((r, p, org[:26]))
                    else:
                        E["E5' 联名/合并文书：多人共用 1 份（正常观察）"].append((r, p, org[:22]))
                continue
            pmulti = len(person_rows[(p, org)]) >= 2
            tails = [tl for _f, tl in pf]
            if len(pf) == 1:
                tl = pf[0][1]
                if tl == "原决定":
                    E["E4 单份却标（原决定）"].append((r, pf[0][0]))
                elif pmulti and not (DATE.match(tl) or tl in SEM):
                    E["E4 同一人多案但无日期"].append((r, pf[0][0]))
                elif not pmulti and DATE.match(tl):
                    E["E4 单人单案多了日期"].append((r, pf[0][0]))
            elif set(tails) & set(SEM) and len(set(tails) & set(SEM)) != len(pf):
                E["E4 复核双件未成对标语义"].append((r, "｜".join(f for f, _t in pf)[:70]))

    # 首行必须同时给出「问题条数」与「观察条数」：run_selfcheck 靠这行取数，
    # 旧写法只有「N 类问题」→ 汇总表永远显示「?」（2026-11-03 实测）
    obs_kinds = [k for k in E if "正常观察" in k]
    n_obs = sum(len(E[k]) for k in obs_kinds)
    n_prob = sum(len(E[k]) for k in E if k not in obs_kinds)
    lines = ["== 命名体检：%d 行 / 问题 %d 条｜观察 %d 条 ==" % (len(rows), n_prob, n_obs)]
    for k in sorted(E):
        lines.append("\n【%s】%d 条" % (k, len(E[k])))
        for x in E[k][:12]:
            lines.append("   " + " ｜ ".join(str(v) for v in x))
        if len(E[k]) > 12:
            lines.append("   …（其余 %d 条）" % (len(E[k]) - 12))
    txt = "\n".join(lines)
    print(txt)
    if a.out:
        open(a.out, "w", encoding="utf-8").write(txt)
        print("\n报告已写：%s" % a.out)


if __name__ == "__main__":
    main()
