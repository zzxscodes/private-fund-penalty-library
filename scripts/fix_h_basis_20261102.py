#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_h_basis_20261102.py —— H 列（核心处罚依据）漏抽条款定点补齐

背景（2026-11-02 逐格体检 X31）：
  4 行的 H 列只有规则名、没有条款号。逐份读原文后确认**不是原文没写，
  而是台账抽取时把「第X条」丢了**：

  行26  原文：违反了《私募投资基金监督管理暂行办法》(证监会令第105号…)第四条第一款的规定，
              按照《暂行办法》第三十三条的规定
        台账：《私募投资基金监督管理暂行办法》
  行206 原文：违反了《深圳证券交易所公司债券上市规则（2022年修订）》…第1.5条、第3.1.1条、
              第3.2.4条，《…（2023年修订）》…第1.5条、第3.1.1条、第3.2.2条；李厚文另违反
              《债券上市规则（2022年修订）》第1.5条、第1.6条第二款…
        台账：《深圳证券交易所公司债券上市规则（2022年修订）》；《债券上市规则（2022年修订）》
  行207 原文：违反了《深圳证券交易所公司债券上市规则（2023年修订）》第1.5条、第3.2.2条的规定
        台账：只有规则名
  行321 原文：违反了《上海证券交易所交易规则》第1.4条，《上海证券交易所科创板股票异常交易
              实时监控细则》第三条，《上海证券交易所程序化交易管理实施细则》第四条等规定
        台账：《上海证券交易所交易规则》

原则：**照原文录，不替原文归纳**；只把「，」并列改成「；」以符合台账一格多依据的分隔约定。

用法：python scripts/fix_h_basis_20261102.py --root <库根> [--apply]
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

from ledger_bak import backup_ledger  # noqa: E402

FIX = [
    (26, "《私募投资基金监督管理暂行办法》",
     "《私募投资基金监督管理暂行办法》第四条第一款；《暂行办法》第三十三条"),
    (206, "《深圳证券交易所公司债券上市规则（2022年修订）》",
     "《深圳证券交易所公司债券上市规则（2022年修订）》第1.5条、第3.1.1条、第3.2.4条；"
     "《深圳证券交易所公司债券上市规则（2023年修订）》第1.5条、第3.1.1条、第3.2.2条；"
     "《债券上市规则（2022年修订）》第1.5条、第1.6条第二款；"
     "《债券上市规则（2023年修订）》第1.5条、第1.6条第二款"),
    (207, "《深圳证券交易所公司债券上市规则（2023年修订）》",
     "《深圳证券交易所公司债券上市规则（2023年修订）》第1.5条、第3.2.2条"),
    (321, "《上海证券交易所交易规则》",
     "《上海证券交易所交易规则》第1.4条；《上海证券交易所科创板股票异常交易实时监控细则》第三条；"
     "《上海证券交易所程序化交易管理实施细则》第四条"),
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
    for row, must, new in FIX:
        cell = ws.cell(row, 8)
        old = str(cell.value or "")
        if must not in old:
            print("× 行%-4d 前置断言失败（现值不含 %r），跳过" % (row, must[:40]))
            continue
        print("行%-4d\n  旧：%s\n  新：%s" % (row, old, new))
        done += 1
        if args.apply:
            cell.value = new

    bad = []
    for i in range(2, ws.max_row + 1):
        h = str(ws.cell(i, 8).value or "")
        if not h:
            bad.append((i, "空"))
            continue
        if h.count("《") != h.count("》"):
            bad.append((i, "书名号不配对"))
        if re.search(r"等规定|鉴于|情节|，$", h):
            bad.append((i, "混入正文"))
        if not re.search(r"第[一二三四五六七八九十百零\d]+(?:\.[\d]+)*条|公告|办法》|条例》", h):
            bad.append((i, "无条款号"))
    print("\n修复 %d 行｜全表 H 列残留 %d 处" % (done, len(bad)))
    for i, r in bad[:10]:
        print("  行%-4d %s" % (i, r))

    if args.apply:
        backup_ledger(root, ledger, "h_basis_20261102")
        wb.save(ledger)
        print("已保存（备份在 _案例库维护留痕/台账备份/）")
    else:
        print("[干跑] 未落盘；加 --apply")
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
