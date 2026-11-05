#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_name_20261102c.py —— 行210 海南海德主体全称纠错（2026-11-02）

## 线索
`verify_library.py` 的 **B5「管理人名称未在正文回现」** 报出 1 条。B5 的判据是拿台账 D 列
机构名去 PDF 正文里找原样回现——找不到就说明台账写的名字和文书上的名字对不上。

## 实证
深证上〔2026〕796号 决定书原文两处都写：
  「关于对**海南海德资本管理股份有限公司**及相关当事人给予公开谴责处分的决定」
  「当事人：**海南海德资本管理股份有限公司**，住所：海南省海口市海德路5号」
台账 D 列写的是「海南海德资本股份有限公司」——**少了「管理」二字**。

这正是本轮把「主体名只认文书首部当事人」列为硬口径的原因：缺两个字，B5 立刻报警，
但如果你只看台账、不看 PDF，是发现不了的。

## 处置
三处必须同步改，否则 A1 立刻报「孤儿 1 / 缺失 1」：
  ① D 列主体名；② E 列措施里出现的旧名；③ 磁盘文件名 + L 列（案例文件）。
本轮该案只有 1 行，不带日期后缀，改名后仍符合 R1/R6。

用法：python scripts/fix_name_20261102c.py --root <库根> [--apply]
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402

OLD = "海南海德资本股份有限公司"
NEW = "海南海德资本管理股份有限公司"
C = LS.COL


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

    hits = []
    for r in range(2, ws.max_row + 1):
        subj = str(ws.cell(r, C["subject"] + 1).value or "")
        if OLD in subj and NEW not in subj:
            hits.append(r)
    print("== 命中 %d 行：%s ==" % (len(hits), hits))
    if not hits:
        print("无需改动（可能已修）。")
        return
    for r in hits:
        print("  行%d 序号%s" % (r, ws.cell(r, 1).value))
        print("     D: %s" % ws.cell(r, C["subject"] + 1).value)
        print("     E: %s…" % str(ws.cell(r, C["measures"] + 1).value or "")[:60].replace("\n", " "))
        print("     L: %s" % ws.cell(r, C["filenames"] + 1).value)
        print("     M: %s" % ws.cell(r, C["paths"] + 1).value)

    if not a.apply:
        print("\n（dry-run；--apply 落盘）")
        return

    from ledger_bak import backup_ledger
    print("\n台账已备份：%s" % backup_ledger(root, ledger, "fix_name_hede"))
    for r in hits:
        # ① 磁盘改名（L 列按「；」分项精确匹配替换）
        d = str(ws.cell(r, C["paths"] + 1).value or "").strip()
        full = d if os.path.isabs(d) else os.path.join(root, d)
        toks = [t.strip() for t in re.split(r"[；;]", str(ws.cell(r, C["filenames"] + 1).value or "")) if t.strip()]
        new_toks = []
        for t in toks:
            nt = t.replace(OLD, NEW)
            if nt != t:
                sp, dp = os.path.join(full, t), os.path.join(full, nt)
                if os.path.isfile(sp) and not os.path.exists(dp):
                    os.rename(sp, dp)
                    print("  改名：%s → %s" % (t, nt))
                else:
                    print("  ⚠ 跳过改名（源不存在或目标已存在）：%s" % t)
            new_toks.append(nt)
        ws.cell(r, C["filenames"] + 1).value = "；".join(new_toks)
        # ② 主体 + ③ 措施
        for key in ("subject", "measures"):
            v = str(ws.cell(r, C[key] + 1).value or "")
            if OLD in v:
                ws.cell(r, C[key] + 1).value = v.replace(OLD, NEW)
    wb.save(ledger)
    print("\n完成 %d 行" % len(hits))


if __name__ == "__main__":
    main()
