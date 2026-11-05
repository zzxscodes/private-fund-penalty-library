#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""normalize_paths.py —— 存放路径列去重（2026-11-02 用户口径）

用户口径：「同一个案例的多个决定书文件存放路径肯定是同一个的，不要按文件数量重复写。」

历史写法：一行多份文书时，L 列把同一个目录写了 N 遍（`A；A；A`），
既冗余，也让「K/L/M 一一对应」的直觉校验变得别扭。

本脚本把 L 列统一成 **1 条目录**；若一行内确实出现多个不同目录（不该发生），
只告警不合并（等人工判定，避免把错误的行改成"看起来正常"）。

用法：
    python scripts/normalize_paths.py --root <库根>            # 干跑
    python scripts/normalize_paths.py --root <库根> --apply    # 落盘（自动备份台账）
"""
import argparse
import os
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
from ledger_bak import backup_ledger  # noqa: E402


def split_multi(v):
    import re
    if v is None:
        return []
    return [p.strip() for p in re.split(r"[；;]", str(v)) if p.strip()]


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
    C_FILES, C_PATHS = 12, 13        # 1-based：K 案例文件 / L 存放路径

    dup, weird = [], []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 4).value:
            continue
        fns = split_multi(ws.cell(r, C_FILES).value)
        pts = split_multi(ws.cell(r, C_PATHS).value)
        if len(pts) <= 1:
            continue
        if len(set(pts)) == 1:
            dup.append((r, len(fns), pts[0]))
        else:
            weird.append((r, fns, pts))

    print("== 同一路径重复填写（可安全去重）%d 行 ==" % len(dup))
    for r, n, p in dup[:12]:
        print("  行%-4d 文件 %d 份 → 路径只留 1 条：%s" % (r, n, p[:56]))
    if len(dup) > 12:
        print("  …（其余 %d 行同型）" % (len(dup) - 12))
    if weird:
        print("\n== 一行内出现多个不同目录 %d 行（不自动改，需人工判） ==" % len(weird))
        for r, fns, pts in weird[:10]:
            print("  行%-4d 文件 %s" % (r, fns[:3]))
            print("        路径 %s" % pts)

    if not a.apply:
        print("\n[干跑] 未改动；确认后加 --apply")
        return
    backup_ledger(root, ledger, "normalize_paths")
    n = 0
    for r, _n, p in dup:
        ws.cell(r, C_PATHS).value = p
        n += 1
    wb.save(ledger)
    print("已去重 %d 行，台账保存" % n)


if __name__ == "__main__":
    main()
