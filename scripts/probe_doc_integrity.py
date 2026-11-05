#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""probe_doc_integrity.py —— 决定书「必备要素」完整性探针（2026-11-02）

缘起（行86 上海忠石）：
  台账 F 列是「被停职后超越授权操作基金产品证券账户…」（个人文书的话），可机构文书里
  根本没有事实段 —— 打开一看，PDF 第 1 页开头就是「根据《私募投资基金监督管理暂行办法》
  第三十三条的规定，现决定对你公司采取责令改正…」，**「经查/存在以下」那段事实被裁掉了**。
  这是网页另存 PDF 时首屏被固定导航遮挡的典型后果：文件存在、能打开、有文字层、
  台账非空，所有常规断言（A1 行数、B 非空、C12 形态）全部通过，**但证据链断了一截**。

判据（行政/自律文书的最小结构）：
  · 事实要素：`经查` / `存在以下` / `存在下列` / `当事人` + `违法|违规|违反`
  · 决定要素：`决定` + `采取|给予|作出|处分|处罚` 或 `予以` / `责令` / `警示函`
  缺任一 → 记「要素缺失」，按缺失项分类。

只读，不改任何文件。输出 CSV 供人工核，重抓需联网。

用法：
  python scripts/probe_doc_integrity.py --root <库根> [--csv out.csv] [--limit N]
"""
import argparse
import csv
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

# 判据必须建在「去空白后的整串」上：PDF 折行会把「违\n反了」拆成两半，
# 直接匹配原文会漏（实测 46 份「要素缺失」里绝大多数是这条造成的假报）。
FACT = re.compile(r"经查|存在以下|存在下列|当事人|检查中发现|现场检查|核查发现|发现你")
VIOL = re.compile(r"违法|违规|违反|不符合|未按")
DECIDE = re.compile(r"决定|予以|责令|警示函|监管措施|纪律处分|处分")


def text_of(fp):
    try:
        import fitz
        d = fitz.open(fp)
        t = "".join(p.get_text() for p in d)
        d.close()
        return t
    except Exception:
        return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--csv", default="")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    dirs = [d for d in os.listdir(root) if d[:2] in ("01", "02", "03")]
    rows, n_scan, n_ok = [], 0, 0
    for d in sorted(dirs):
        for dp, _, fs in os.walk(os.path.join(root, d)):
            for f in fs:
                if not f.lower().endswith(".pdf"):
                    continue
                fp = os.path.join(dp, f)
                t = text_of(fp)
                if len(t.strip()) < 200:
                    n_scan += 1                     # 扫描件，本探针不判（需 OCR）
                    continue
                t = re.sub(r"\s+", "", t)      # 去空白后再判（折行断字）
                miss = []
                if not FACT.search(t):
                    miss.append("无事实段")
                if not VIOL.search(t):
                    miss.append("无违规表述")
                if not DECIDE.search(t):
                    miss.append("无决定段")
                if miss:
                    rows.append((d, f, "、".join(miss), len(t)))
                else:
                    n_ok += 1
                if args.limit and len(rows) >= args.limit:
                    break

    print("\n有文字层且要素齐全 %d 份｜扫描件跳过 %d 份｜要素缺失 %d 份"
          % (n_ok, n_scan, len(rows)))
    for d, f, why, n in rows[:30]:
        print("  %-8s %-52s %-14s %d字" % (d[:8], f[:50], why, n))
    if len(rows) > 30:
        print("  …其余 %d 份见 CSV" % (len(rows) - 30))
    if args.csv and rows:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["目录", "文件", "缺失要素", "文字量"])
            w.writerows(rows)
        print("\nCSV：%s" % args.csv)
    return 0


if __name__ == "__main__":
    sys.exit(main())
