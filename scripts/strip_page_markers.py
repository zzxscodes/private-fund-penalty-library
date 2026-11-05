#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""strip_page_markers.py —— 清掉正文列里混进的原文页码标记（2026-11-02）

## 现象
OCR 抓证监局/交易所 PDF 时，页脚的「— 2 —」「– 3 –」会被当成正文吸进
处罚措施 / 核心处罚事由 / 核心处罚依据，于是台账里出现：

    核心处罚事由：— 2 — 2021 年3 月，海德股份开展不良资产收购重组业务…

读者第一眼看到的是乱码。实测 5 处（E 措施 2、F 事由 3）。

## 判据
只认**长破折号**形态：`— 2 —` / `– 3 –` / `第 2 页`。
**不认半角连字符**——「2026-08-15」里的 `-08-` 长得一样，一并处理会把日期改坏
（第一版正则就是 `[—–-]`，直接报出 9 条假报，其中 4 条是日期）。

## 用法
  python scripts/strip_page_markers.py --root <库根>            # 只看
  python scripts/strip_page_markers.py --root <库根> --apply
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402

COLS = ("measures", "grounds", "evidence", "legal_basis", "defense")
MARK = re.compile(r"[—–]\s*\d{1,3}\s*[—–]|第\s*\d{1,3}\s*页")
# 标记常与前后文字连成一行，删掉标记本身即可，顺手收掉行首/断行处的多余空白
CLEAN = re.compile(r"[ \t\u3000]*[—–]\s*\d{1,3}\s*[—–][ \t\u3000]*")


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
        for key in COLS:
            cell = ws.cell(r, LS.COL[key] + 1)
            v = str(cell.value or "")
            if not MARK.search(v):
                continue
            new = CLEAN.sub("", v).strip()
            new = re.sub(r"\n{3,}", "\n\n", new)
            hits.append((r, key, MARK.findall(v), new))
    print("== 页码标记：命中 %d 处 ==" % len(hits))
    for r, key, marks, new in hits:
        print("  行%-4d %-12s %s" % (r, key, marks))
    if not hits:
        return
    if not a.apply:
        print("\n（dry-run；--apply 落盘）")
        return
    from ledger_bak import backup_ledger
    print("\n台账已备份：%s" % backup_ledger(root, ledger, "strip_page_marks"))
    for r, key, _m, new in hits:
        ws.cell(r, LS.COL[key] + 1).value = new
    wb.save(ledger)
    print("\n已清理 %d 处" % len(hits))


if __name__ == "__main__":
    main()
