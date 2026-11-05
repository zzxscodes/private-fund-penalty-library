#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""normalize_subjects.py —— 台账 D 列（被处罚机构/人）规范化体检与修复（2026-11-02）

查三类问题：

  N1 机构名前角色短语污染 —— 「曾登记为X公司」「实际负责X公司」。
     危害不只是难看：机构名一被污染，同一家中基协/证监局的多份文书就不再归为同一主体，
     「同机构多案加日期」「同案拆行判重」全部失效（浙江优策 4 行曾被拆成 3 个"不同机构"）。
     修复：subject_clean.strip_role_prefix（单一事实源）。

  N2 人员名在整行正文（E~J 列）中从不出现 —— **只报不改**。
     实测 6 条里 5 条是「联名决定书」：决定书抬头写「关于对胡瀚阳、卞桂玲、孙曼采取…」，
     但台账措施列只抽到第一条「高管（胡瀚阳）：对你们采取…」，后两人在台账文本里
     自然不出现。这类是**正常**的，绝不能自动删人。
     真正要处理的两类：
       · 提取碎片（杭州复投「新享」：措施/事由/抬头都没有，且平台只有 3 份文书）→ 删
       · 漏登（马鞍山哲方抬头上是「乔徽、李皓、李磊」，台账只写乔徽、李皓）→ 补
     所以本脚本只列清单，判定交人工。

  N3 机构名无组织形式后缀（纯人名当机构）—— 只报不改，多为「只处罚从业人员、无机构主体」
     的合法个人案件（魏锋、张守军），文件名就是「姓名.pdf」，不是错误。

用法：
  python scripts/normalize_subjects.py --root <库根>            # 体检
  python scripts/normalize_subjects.py --root <库根> --apply    # 只修 N1
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

from subject_clean import strip_role_prefix  # noqa: E402

C_SUBJECT = 4  # 1-based
BODY_COLS = range(5, 11)  # E~J：措施/事由/证据/依据/申辩/违规类型


def split_multi(v):
    return [p.strip() for p in re.split(r"[；;]", str(v or "")) if p.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    ledger = next(os.path.join(root, f) for f in os.listdir(root)
                  if f.startswith("04_") and f.endswith(".xlsx") and not f.startswith("~$"))
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    n1, n2, n3 = [], [], []
    for r in range(2, ws.max_row + 1):
        subj = str(ws.cell(r, C_SUBJECT).value or "").strip()
        if not subj:
            continue
        segs = split_multi(subj)
        org = segs[0]
        fixed = strip_role_prefix(org)
        if fixed != org:
            n1.append((r, org, fixed))
        persons = []
        for s in segs[1:]:
            persons += [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]
        body = "".join(str(ws.cell(r, c).value or "") for c in BODY_COLS)
        for p in persons:
            if p and p not in body:
                n2.append((r, org[:26], p, str(ws.cell(r, 5).value or "")[:36].replace("\n", " ")))
        if not any(t in org for t in ("公司", "合伙企业", "企业", "中心", "基金", "协会",
                                      "银行", "证券", "信托", "集团", "事务所", "研究院")):
            n3.append((r, org, "、".join(persons)[:30]))

    print("== D 列规范化体检 ==")
    print("\n【N1 机构名带角色前缀】%d 行" % len(n1))
    for r, o, f in n1:
        print("  行%-4d %s → %s" % (r, o, f))
    print("\n【N2 人员名在整行正文中不出现（只报不改，需照文书抬头判定）】%d 条" % len(n2))
    for r, o, p, m in n2:
        print("  行%-4d %-28s ｜%-6s ｜措施=%s" % (r, o, p, m))
    print("\n【N3 机构段是纯人名（个人案，通常正常）】%d 行" % len(n3))
    for r, o, p in n3:
        print("  行%-4d %-16s ｜%s" % (r, o, p))

    if not a.apply:
        print("\n（体检模式；--apply 只修 N1）")
        return
    if not n1:
        print("\nN1 无待修项。")
        return
    from ledger_bak import backup_ledger
    print("\n台账已备份：%s" % backup_ledger(root, ledger, "normalize_subjects"))
    for r, o, f in n1:
        segs = split_multi(ws.cell(r, C_SUBJECT).value)
        segs[0] = f
        ws.cell(r, C_SUBJECT).value = "；".join(segs)
        print("  已修 行%-4d → %s" % (r, f))
    wb.save(ledger)
    print("\n完成，修 %d 行" % len(n1))


if __name__ == "__main__":
    main()
