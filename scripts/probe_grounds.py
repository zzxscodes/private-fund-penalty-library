#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""probe_grounds.py —— F 列「旧值 vs 决定书原文重抽」对照探针（2026-11-02，只读）

用途：在批量改 F 列之前，先把「旧值 / 重抽值 / 原文有没有分点」并排打出来，人眼过一遍。
**不写台账**，只产出 CSV 供审阅。

关键实现：pdf_text_join()
  fitz 按行提取 PDF 文本，中文行尾会带 \\n，直接 join 会把「交由」变成「交 由」、
  「勤勉」变成「勤 勉」——这就是台账里 F/H 列大量半角空格的来源。
  这里把「中文行尾 + 中文行首」的换行直接删掉（不加空格），但：
    ① 行尾是「。；」不删（段落边界）
    ② 下一行以枚举标记（一、/（一）/1.）开头不删（分条边界）

用法：
  python scripts/probe_grounds.py --root <库根> [--all] [--out x.csv]
    --all   不只测有问题的行，全库 508 行都跑（慢，扫描件要 OCR）
"""
import argparse
import csv
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import extract_core as EC  # noqa: E402
import ledger_style as LS  # noqa: E402
from audit_cells import check_cell  # noqa: E402

JOIN = re.compile(
    r"(?<=[\u4e00-\u9fa5，、：；])\n"
    r"(?![一二三四五六七八九十]+、|[（(]\d+[）)]|\d+[.、])"
    r"(?=[\u4e00-\u9fa5（《“])")
ENUM_ANY = re.compile(r"(?:^|\n)\s*(?:[一二三四五六七八九十]+、|[（(][一二三四五六七八九十\d]+[）)]|\d+[.、])")


def pdf_text_join(path):
    import fitz
    doc = fitz.open(path)
    raw = "\n".join(pg.get_text() for pg in doc)
    doc.close()
    return JOIN.sub("", raw)


def text_layer_ok(path):
    try:
        import fitz
        doc = fitz.open(path)
        n = sum(len(pg.get_text().strip()) for pg in doc)
        doc.close()
        return n > 300
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--out", default="")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    ws = openpyxl.load_workbook(ledger, data_only=True).worksheets[0]
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if r[0]:
            rows.append((i, r))

    out = args.out or os.path.join(LS.maint_dir(root), "F列重抽对照_20261102.csv")
    os.makedirs(os.path.dirname(out), exist_ok=True)

    recs, n_scan = [], 0
    for i, r in rows:
        old = "" if r[5] is None else str(r[5])
        codes = check_cell(5, old, i)
        if not args.all and not codes:
            continue
        names = [x.strip() for x in re.split(r"[；;]", str(r[11] or "")) if x.strip()]
        paths = [x.strip() for x in re.split(r"[；;]", str(r[12] or "")) if x.strip()]
        d = paths[0] if paths else ""
        # 取第一份机构件的 PDF（文件名最短且不含「-」的优先）
        org_names = [n for n in names if "-" not in n] or names
        target = org_names[0] if org_names else ""
        p = os.path.join(root, d, target) if target else ""
        new, src, has_enum, layered = "", "", "", ""
        if p and os.path.isfile(p):
            layered = "有" if text_layer_ok(p) else "无"
            if layered == "无":
                n_scan += 1
                src = "（扫描件，未 OCR）"
            else:
                t = pdf_text_join(p)
                source = "中基协" if str(r[10]) == "中基协" else "证监局"
                try:
                    e = EC.extract_amac(t) if source == "中基协" else EC.extract_csrc(t)
                    new = e.get("grounds", "")
                except Exception as ex:
                    src = "抽取异常：%s" % ex
                m = re.search(r"(?:经查|存在以下)", t)
                seg = t[m.start():m.start() + 1200] if m else t[:1200]
                has_enum = "有" if ENUM_ANY.search(seg) else "无"
        recs.append([i, r[0], str(r[3] or "")[:40], str(r[10] or ""), ",".join(codes),
                     layered, has_enum, old, new, src])

    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["行", "序号", "主体", "来源", "命中形态", "文字层", "原文有分点",
                    "旧F", "重抽F", "备注"])
        w.writerows(recs)

    same = sum(1 for x in recs if x[7].strip() == x[8].strip() and x[8])
    diff = sum(1 for x in recs if x[8] and x[7].strip() != x[8].strip())
    noenum = sum(1 for x in recs if x[6] == "无")
    print("== F 列重抽对照 ==")
    print("  待查 %d 行｜原文无分点 %d｜重抽与旧值相同 %d｜不同 %d｜扫描件未 OCR %d"
          % (len(recs), noenum, same, diff, n_scan))
    print("  输出：%s" % out)
    for x in recs[:6]:
        print("\n  行%s %s [%s]" % (x[0], x[2], x[4]))
        print("    旧：%s" % x[7][:110])
        print("    新：%s" % x[8][:110])
    return 0


if __name__ == "__main__":
    sys.exit(main())
