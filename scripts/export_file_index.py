#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""export_file_index.py —— 导出台账 ↔ 磁盘文件逐份对照清单（人工检验用，2026-11-02）

为什么需要：verify_library.py 只回答「数量对不对」，audit_filenames.py 只回答
「有没有违规」；但人工检验需要的是**一张能逐行勾选的表**——台账第 N 行的
这几份文件，在磁盘上到底叫什么、在哪个目录、命名属于哪一形态。

输出列（1 行 = 1 份决定书，按台账序号升序）：
  台账序号 / 机构来源 / 被处罚机构人 / 处罚时间 / 分类目录 / 文件名
  / 命名形态 / 是否在磁盘 / 是否 05 副本

用法：
  python scripts/export_file_index.py --root <库根> [--out <csv 路径>]
默认输出到 <库父目录>/_案例库维护留痕/台账-文件逐份对照_<yyyymmdd>.csv
（UTF-8 BOM，Excel 直接双击打开不乱码）
"""
import argparse
import csv
import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402

DATE_TAIL = re.compile(r"（(20\d{2}-\d{2}-\d{2})）$")
SEM_TAIL = re.compile(r"（(原决定|复核决定)）$")


def split_multi(v):
    if v is None:
        return []
    return [p.strip() for p in re.split(r"[；;]", str(v)) if p.strip()]


def shape_of(name, subject):
    """判命名形态（只做分类，不判对错——对错交给 audit_filenames.py）"""
    stem = os.path.splitext(name)[0]
    sem = SEM_TAIL.search(stem)
    if sem:
        base = SEM_TAIL.sub("", stem)
        kind = sem.group(1)
        dated = bool(DATE_TAIL.search(base))
        who = "人员件" if "-" in base else "机构件"
        return f"{who}·{kind}" + ("·带日期" if dated else "")
    base = DATE_TAIL.sub("", stem)
    dated = bool(DATE_TAIL.search(stem))
    who = "人员件" if "-" in base else "机构件"
    return who + ("·带日期" if dated else "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--ledger", default="")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "台账.xlsx")
    if not os.path.isfile(ledger):
        cands = [f for f in os.listdir(root)
                 if f.endswith(".xlsx") and ("台账" in f or "总表" in f)
                 and not f.startswith("~$")]
        if not cands:
            print("× 找不到台账 xlsx")
            return 2
        ledger = os.path.join(root, cands[0])

    wb = openpyxl.load_workbook(ledger, data_only=True)
    ws = wb.worksheets[0]

    out = args.out or os.path.join(
        LS.maint_dir(root),
        "台账-文件逐份对照_%s.csv" % datetime.date.today().strftime("%Y%m%d"))
    os.makedirs(os.path.dirname(out), exist_ok=True)

    rows, n_file, n_missing = [], 0, 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r is None or all(v in (None, "") for v in r):
            continue
        seq = r[LS.COL["seq"]]
        if seq in (None, ""):
            continue
        source = (r[LS.COL["source"]] or "").strip()
        subject = (r[LS.COL["subject"]] or "").strip()
        decided = r[LS.COL["decided"]]
        if isinstance(decided, (datetime.datetime, datetime.date)):
            decided = decided.strftime("%Y-%m-%d")
        else:
            decided = str(decided or "").strip()
        names = split_multi(r[LS.COL["filenames"]])
        dirs = split_multi(r[LS.COL["paths"]])
        # M 列单值口径：正常只有 1 条；若历史数据仍多值，按序对应（不崩）
        for i, name in enumerate(names):
            n_file += 1
            d = dirs[i] if i < len(dirs) else (dirs[0] if dirs else "")
            full = name if os.path.isabs(name) else os.path.join(root, d, name)
            exists = os.path.isfile(full)
            if not exists:
                n_missing += 1
            rows.append([
                seq, source, subject, decided, d, name,
                shape_of(name, subject),
                "√" if exists else "× 缺失",
                "是" if d.startswith("05") else "",
            ])

    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["台账序号", "机构来源", "被处罚机构人", "处罚时间",
                    "分类目录", "文件名", "命名形态", "是否在磁盘", "05副本"])
        w.writerows(rows)

    print("== 逐份对照清单 ==")
    print("  台账案例 %d 行 → 文件 %d 份（磁盘缺失 %d）" % (
        len({r[0] for r in rows}), n_file, n_missing))
    print("  输出：%s" % out)
    return 0 if n_missing == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
