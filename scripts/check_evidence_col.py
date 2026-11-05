# -*- coding: utf-8 -*-
"""G 列（核心证据材料）回现体检 —— 双向判据。

方向一「占位漏抽」：G 仍写「未列举证据材料」，但文书里明明有「上述违法事实，有…等证据证明」，
    → 说明抽取漏了。这是 2026-11-02 实测发现的问题（3 行，行106/213/402）。

方向二「内容串味」：G 写了具体清单，但清单项在文书里回现不了（命中率过低），
    → 说明串了别的案子的证据，或 OCR 写坏。

为什么单列一个探针：G 列 89% 是占位，靠 audit_cells（只查形态）和 check_fields_source
（只查 F 列）都盖不住；而占位本身是**诚实标注**——多数证监局文书的证据句只出现在
行政处罚决定书里，行政监管措施决定书（警示函/责令改正）确实不写证据段。
所以这个探针的输出必须分两类看：
  · 「漏抽」= 真问题，必须补；
  · 「原文无证据句」= 合规占位，不动。

用法：
    python check_evidence_col.py --root <库根>                  # 汇总
    python check_evidence_col.py --root <库根> --detail          # 明细
    python check_evidence_col.py --root <库根> --ocr             # 无文字层的补 OCR（慢）
    python check_evidence_col.py --root <库根> --csv out.csv
"""
import argparse
import csv
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl

from check_fields_source import is_person_file, pick_inst_file, norm

LEDGER = "04_综合索引台账【总表】.xlsx"
PLACEHOLDER = "未列举证据材料"

# 证据句：主句式「上述/以上…事实，有…等证据证明/予以证明」
EV_SENT = re.compile(
    r"(?:上述|以上|前述)?(?:违法|违规|违规违法|违法违约)?"
    r"(?:事实|违法事实|违规事实)"
    r"[^。]{0,15}?[，,]\s*有([^。]{4,400}?)等?证据"
    r"(?:证明|予以证明|证实|佐证|为证|在案)"
)
# 兜底句式：「有…等证据证明」
EV_ALT = re.compile(r"有([^。]{6,300}?)等证据(?:证明|予以证明|证实|佐证|为证)")

DF = "、，,；;"


def split_items(clause):
    return [x.strip() for x in re.split(r"[%s]" % DF, clause) if len(x.strip()) >= 2]


def find_evid(txt):
    """返回 (证据项清单, 命中句式)，找不到返回 ([], "")"""
    for rx in (EV_SENT, EV_ALT):
        m = rx.search(txt)
        if m:
            return split_items(m.group(1)), rx.pattern[:6]
    return [], ""


def read_text(path, cache_dir=None, ocr=False):
    import fitz

    d = fitz.open(path)
    try:
        t = re.sub(r"\s+", "", "".join(p.get_text() for p in d))
    finally:
        d.close()
    if len(t) >= 200:
        return t, "text"
    if not ocr:
        return t, "notext"
    try:
        from ocrlib import ocr_text

        return re.sub(r"\s+", "", ocr_text(path, cache_dir=cache_dir)), "ocr"
    except Exception as e:  # OCR 引擎缺失不应中断全量体检
        return t, "ocr_err:%s" % e


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--detail", action="store_true")
    ap.add_argument("--ocr", action="store_true", help="对无文字层的文书补 OCR")
    ap.add_argument("--csv", default="")
    ap.add_argument("--rows", default="", help="只查指定行，如 106,213,402")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = a.ledger or os.path.join(root, LEDGER)
    only = {int(x) for x in re.findall(r"\d+", a.rows)} if a.rows else None
    cache_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")

    wb = openpyxl.load_workbook(ledger, data_only=True)
    ws = wb.active

    miss_fill, foreign, honest, notext = [], [], 0, []
    rows_out = []

    for r in range(2, ws.max_row + 1):
        if only and r not in only:
            continue
        g = str(ws.cell(r, 7).value or "").strip()
        rel = str(ws.cell(r, 13).value or "").strip()
        files = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        if not files:
            continue
        inst = pick_inst_file(files)
        # G 列只认机构文书；无机构文书时按「先到先填」退回人员文书
        prim = inst or (files[0] if len(files) == 1 else "")

        # 只有人员文书（一案多人共用/无机构文书）时 prim 为空，此时仍要逐份试，
        # 否则整行被误判成「无文字层」—— 行57/135/149 实测踩坑。
        txt, how = "", ""
        order = ([prim] if prim else []) + [f for f in files if f != prim]
        for f in order:
            p = os.path.join(root, rel, f)
            if not os.path.isfile(p):
                continue
            txt, how = read_text(p, cache_dir, a.ocr)
            if len(txt) >= 200 or how.startswith("ocr_err"):
                prim = f
                break
        if how == "notext" or not txt:
            notext.append((r, prim, how))
            continue

        items, _pat = find_evid(txt)
        src = "机构" if (inst and prim == inst) else "人员"

        if PLACEHOLDER in g or not g:
            if items:
                miss_fill.append((r, prim, src, "、".join(items)))
                rows_out.append([r, "漏抽", prim, "、".join(items)])
            else:
                honest += 1
        else:
            # 内容方向：清单项逐项回现
            gi = split_items(re.sub(r"等?证据$", "", g))
            bad = [it for it in gi if norm(it) not in norm(txt) and len(norm(it)) >= 3]
            if gi and len(bad) == len(gi):
                foreign.append((r, prim, src, g, bad))
                rows_out.append([r, "串味", prim, g])
            elif bad and len(bad) * 2 > len(gi):
                foreign.append((r, prim, src, g, bad))
                rows_out.append([r, "部分不匹配", prim, "／".join(bad)])

    wb.close()

    print("=" * 78)
    print("G 列（核心证据材料）回现体检")
    print("=" * 78)
    print("占位且原文确无证据句（合规）: %d" % honest)
    print("占位但原文有证据句（漏抽）  : %d" % len(miss_fill))
    print("有内容但回现不上（串味）    : %d" % len(foreign))
    print("无文字层未判定              : %d" % len(notext))

    if a.detail or miss_fill or foreign:
        if miss_fill:
            print("\n--- 漏抽（应补）---")
            for r, f, src, it in miss_fill:
                print(" 行%-4d [%s] %s\n        原文证据：%s" % (r, src, f[:42], it[:150]))
        if foreign:
            print("\n--- 串味（应核）---")
            for r, f, src, g, bad in foreign:
                print(" 行%-4d [%s] %s\n        G=%s\n        未回现：%s"
                      % (r, src, f[:42], g[:100], "／".join(bad)[:100]))
    if notext:
        print("\n--- 无文字层（加 --ocr 复判）---")
        for r, f, how in notext:
            print(" 行%-4d %s  %s" % (r, f[:50], how))

    if a.csv:
        with open(a.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["台账行", "类型", "文书", "明细"])
            w.writerows(rows_out)
        print("\nCSV 已写：%s" % a.csv)

    return 0 if not (miss_fill or foreign) else 1


if __name__ == "__main__":
    sys.exit(main())
