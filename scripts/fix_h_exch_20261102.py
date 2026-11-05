#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_h_exch_20261102.py —— 交易所 H 列（核心处罚依据）定点修复

问题（2026-11-02 逐格体检发现）：
  行73、行242 的 H 列把**正文**混进了依据列，规则名还是 OCR 残缺的：
    行73  ：《上海证券交易所首次公开发行证券发行与三条、第六条、第七条等规定。
           鉴于上述违规事实和情节，根据《实施细则》第七十四条、第七十五条；《实施细则》第七十四条、第七十五条
    行242 ：《上海证券交易所首次公开发行证券发行与（））第三条…（同上）
  同一批案子（行72/74/335/336）在库内的标准写法是：
    《上海证券交易所首次公开发行证券发行与承销业务实施细则》第三条、第六条、第七条；
    《实施细则》第七十四条、第七十五条
  法规清单第 41 条收录的也是这个名字 → 以库内一致写法为准，不另造名称。

做法：带前置断言（old_has），只改这两行；改完校验「书名号配对 + 无正文词 + 无重复段」。

用法：python scripts/fix_h_exch_20261102.py --root <库根> [--apply]
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

from ledger_bak import backup_ledger  # noqa: E402

CANON = ("《上海证券交易所首次公开发行证券发行与承销业务实施细则》第三条、第六条、第七条；"
         "《实施细则》第七十四条、第七十五条")

FIX = [
    (73, "发行与三条"),
    (242, "发行与（））"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]

    done = 0
    for row, must in FIX:
        cell = ws.cell(row, 8)  # H 列（1-based）
        old = str(cell.value or "")
        if must not in old:
            print("× 行%-4d 前置断言失败（找不到 %r），跳过" % (row, must))
            continue
        print("行%-4d\n  旧：%s\n  新：%s" % (row, old[:120], CANON))
        done += 1
        if args.apply:
            cell.value = CANON

    # 校验：所有引用了该规则的 H 格，书名号必须配对、不得含正文词、不得重复条款段
    bad = []
    for i in range(2, ws.max_row + 1):
        h = str(ws.cell(i, 8).value or "")
        if "首次公开发行证券发行与" not in h:
            continue
        if h.count("《") != h.count("》"):
            bad.append((i, "书名号不配对"))
        if re.search(r"等规定|鉴于|情节", h):
            bad.append((i, "混入正文"))
        if h.count("第七十四条") > 1:
            bad.append((i, "条款段重复"))

    print("\n修复 %d 行｜校验残留 %d 处" % (done, len(bad)))
    for i, r in bad:
        print("  行%-4d %s" % (i, r))

    if args.apply:
        backup_ledger(root, ledger, "h_exch_20261102")
        wb.save(ledger)
        print("已保存（备份在 _案例库维护留痕/台账备份/）")
    else:
        print("[干跑] 未落盘；加 --apply")
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
