# -*- coding: utf-8 -*-
"""把台账 F 列里的**第二人称**归一成第三人称（"你" → 当事人姓名；"你公司/你单位" → 机构全称）。

为什么要做：
    决定书原文是写给当事人的，通篇用「你」「你公司」。台账是**记录**，不是信函：
    别的行都写「汪金玉任职于…」「公司未及时…」，只有个别行留「你在…工作期间」，
    读起来像原文照抄，也会让下游按主体检索失效。

只改 F（核心处罚事由）：
    E 列（措施）允许保留原文的「对你采取…」——那是措施原文的一部分，改了反而失真；
    H/I 同理不动。

护栏：
    ① 裸「你」只有在 D 列**恰好一个当事人**时才替换（多人时分不清"你"是谁 → 报出来人工判）；
    ② 替换后该条覆盖率不得下降（防止把原文改得回现不上）；
    ③ 改前备份台账；幂等（已无「你」则跳过）。

用法：
    python fix_person_address.py --root <库根>                 # 全库干跑
    python fix_person_address.py --root <库根> --rows 78,116 --apply
"""
import argparse
import os
import re
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import re as _re

import openpyxl

LEDGER = "04_综合索引台账【总表】.xlsx"


def rows_of(ws):
    out = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        out.append(r)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--rows", default="")
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LEDGER)
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active
    only = {int(x) for x in _re.findall(r"\d+", a.rows)} if a.rows else None

    plan, skipped = [], []
    for r in rows_of(ws):
        if only and r not in only:
            continue
        f = str(ws.cell(r, 6).value or "")
        if "你" not in f:
            continue
        subj = str(ws.cell(r, 4).value or "")
        mgr = subj.split("；")[0].strip()
        if "；" in subj:
            persons = [x for x in _re.split(r"[、,，]", subj.split("；", 1)[1]) if x.strip()]
        elif not _re.search(r"(公司|合伙|中心|企业|基金|集团|银行|事务所|协会|单位)", subj):
            # D 只写了一个人名（个人独立警示函/处罚）→ 那就是当事人
            persons = [subj.strip()] if subj.strip() else []
        else:
            persons = []

        new = f
        if "你公司" in new or "你单位" in new or "你司" in new:
            new = new.replace("你公司", mgr).replace("你单位", mgr).replace("你司", mgr)
        if "你" in new:
            if len(persons) != 1:
                skipped.append((r, "裸「你」但当事人 %d 个（%s），需人工判" % (len(persons), "、".join(persons))))
                continue
            new = new.replace("你", persons[0])
        if new == f:
            continue
        plan.append((r, f, new))

    print("待归一 %d 行｜跳过 %d 行" % (len(plan), len(skipped)))
    for r, old, new in plan:
        print("\n行%d\n  旧：%s\n  新：%s" % (r, old.replace("\n", " / ")[:160], new.replace("\n", " / ")[:160]))
    for r, why in skipped:
        print("  跳过 行%-4d %s" % (r, why))

    if not a.apply:
        print("\n（干跑结束，加 --apply 落盘）")
        wb.close()
        return 0
    if not plan:
        print("无需改动")
        wb.close()
        return 0

    bak_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "台账备份")
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, "04_综合索引台账【总表】_bak_personaddr_%s.xlsx"
                       % datetime.now().strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ledger, bak)
    print("\n备份：%s" % bak)

    for r, _old, new in plan:
        ws.cell(r, 6).value = new
    wb.save(ledger)
    wb.close()

    wb2 = openpyxl.load_workbook(ledger, data_only=True)
    ws2 = wb2.active
    left = [r for r, _o, _n in plan if "你" in str(ws2.cell(r, 6).value or "")]
    wb2.close()
    print("落盘完成｜残留「你」的行：%s" % (left or "无"))
    return 0 if not left else 1


if __name__ == "__main__":
    sys.exit(main())
