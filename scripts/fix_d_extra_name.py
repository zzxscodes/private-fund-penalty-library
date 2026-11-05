# -*- coding: utf-8 -*-
"""F9 定点消除：把 D 列里「无文书、无措施、也不在文书首部」的人名去掉。

为什么会有这种错：
    F 列正文常写「实际控制人黄巍利用基金财产为自己或他人牟利」，
    抽取器把正文里出现的人名当成了当事人 → 写进 D 列。
    这正是硬口径③「主体名只认文书首部当事人，正文顺取一律不要」的实例。

    行408（浙江优策 / 尉亲）：该行只有一份文书——中基协处分〔2025〕359 号，
    首部「当事人：尉亲」；黄巍的文书是 334 号，在行449。D 列写成「尉亲、黄巍」是串名。

护栏（三条全过才允许改）：
    ① 该人名在本行**没有任何文件**（否则可能是正常当事人，不该删）；
    ② 该人名在 E 列措施里**完全没出现**（联名决定书会给每个人写「高管（姓名）：…」）；
    ③ 该人名不在本行任一文书**首部**（文号 + 当事人段，到「根据《」之前）；扫描件走 OCR。

用法：
    python fix_d_extra_name.py --root <库根> --row 408 --name 黄巍          # 干跑
    python fix_d_extra_name.py --root <库根> --row 408 --name 黄巍 --apply  # 落盘
"""
import argparse
import os
import re
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl

LEDGER = "04_综合索引台账【总表】.xlsx"


def _read_text(path):
    """优先文字层，没有就 OCR（带缓存，缓存在库外）。"""
    import fitz
    d = fitz.open(path)
    try:
        t = "".join(p.get_text() for p in d)
    finally:
        d.close()
    if len(re.sub(r"\s+", "", t)) >= 200:
        return t
    try:
        import ocrlib
        return ocrlib.ocr_text(path, cache_dir=CACHE) or ""
    except Exception:
        return t


CACHE = ""


def head_of(txt):
    """文书首部：去掉空白后，取「根据《」之前的部分（文号 + 当事人）。"""
    g = re.sub(r"\s+", "", txt or "")
    m = re.search(r"根据《", g)
    return g[:m.start()] if m else g[:600]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--row", type=int, required=True)
    ap.add_argument("--name", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    global CACHE
    root = os.path.abspath(a.root)
    CACHE = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")
    ledger = os.path.join(root, LEDGER)

    wb = openpyxl.load_workbook(ledger)
    ws = wb.active
    r, nm = a.row, a.name

    subj = str(ws.cell(r, 4).value or "")
    meas = str(ws.cell(r, 5).value or "")
    rel = str(ws.cell(r, 13).value or "").strip()
    files = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]

    print("行%d  D=%s" % (r, subj))
    print("     L=%s" % "；".join(files))

    # 护栏 ①：本行没有这个人的文件
    has_file = any(re.sub(r"\.pdf$", "", f, flags=re.I).split("-")[0] == nm for f in files)
    if has_file:
        print("✗ 本行有「%s」的文件，可能是正常当事人，中止" % nm)
        wb.close()
        return 1
    # 护栏 ②：措施里也没提他
    if nm in meas:
        print("✗ E 列措施里出现「%s」（联名决定书形态），中止" % nm)
        wb.close()
        return 1
    # 护栏 ③：不在任何文书首部
    for f in files:
        p = os.path.join(root, rel, f)
        if not os.path.isfile(p):
            print("✗ 文件缺失，无法核对首部：%s" % f)
            wb.close()
            return 1
        h = head_of(_read_text(p))
        if nm in h:
            print("✗ 「%s」出现在《%s》首部，是当事人，中止" % (nm, f))
            wb.close()
            return 1
    print("✓ 三条护栏通过：无文件、无措施、不在任何文书首部")

    parts = [x for x in re.split(r"([、,，])", subj.split("；", 1)[1])] if "；" in subj else []
    if not parts or nm not in parts:
        print("✗ D 列里没有「%s」，无需处理" % nm)
        wb.close()
        return 0
    names = [x for x in re.split(r"[、,，]", subj.split("；", 1)[1]) if x.strip()]
    new_names = [x for x in names if x.strip() != nm]
    new_subj = "%s；%s" % (subj.split("；")[0].strip(), "、".join(new_names))
    print("新 D=%s" % new_subj)

    if not new_names:
        print("✗ 删完没人了，D 列会失真，中止")
        wb.close()
        return 1

    if not a.apply:
        print("\n（干跑结束，加 --apply 落盘）")
        wb.close()
        return 0

    bak_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "台账备份")
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, "04_综合索引台账【总表】_bak_f9_row%d_%s.xlsx"
                       % (r, datetime.now().strftime("%Y%m%d_%H%M%S")))
    shutil.copy2(ledger, bak)
    print("备份：%s" % bak)

    ws.cell(r, 4).value = new_subj
    wb.save(ledger)
    wb.close()

    wb2 = openpyxl.load_workbook(ledger, data_only=True)
    got = str(wb2.active.cell(r, 4).value or "")
    wb2.close()
    ok = got == new_subj
    print("落盘复验：%s（%s）" % ("OK" if ok else "✗ 未写入", got))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
