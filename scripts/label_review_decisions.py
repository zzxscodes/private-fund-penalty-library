# -*- coding: utf-8 -*-
"""同一案例挂了「原决定 + 复核决定」两份文书的，文件名标明是哪一份。

背景（2026-11-02 用户裁定）：「以日期为结尾的命名不具有明确意义」——
复核案的两个文件应叫「X（原决定）.pdf」「X（复核决定）.pdf」，一眼能分清。

判定（不看文件名，看文书自身）：
  - 首页出现「复核决定书 / 复核申请」→ 复核决定
  - 首页出现「纪律处分决定书 / 行政监管措施决定书 / 行政处罚决定书」→ 原决定
  - 同一个小案例里同时存在两类时，才加标签（只有一类时不加，避免无意义的后缀）

用法：
  python label_review_decisions.py --root <库根>            # 干跑
  python label_review_decisions.py --root <库根> --apply    # 落盘（备份台账 + 被改名 PDF）
"""
import os
import re
import sys
import shutil
import argparse
from datetime import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import ledger_style as LS  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

C_FILES = LS.COL["filenames"] + 1
C_PATHS = LS.COL["paths"] + 1
T05 = "05_重点典型案例"

_REVIEW = "复核决定"
_ORIG = "原决定"
_TAG = re.compile(r"（(?:原决定|复核决定)）(?=\.pdf$)")
_DATE_TAG = re.compile(r"（\d{4}-\d{2}-\d{2}）(?=\.pdf$)")


def _kind_from_text(t):
    h = re.sub(r"\s+", "", t or "")[:300]
    if "复核决定" in h or "复核申请" in h:
        return _REVIEW
    if re.search(r"(纪律处分决定书|行政监管措施决定书|行政处罚决定书|市场禁入决定书)", h):
        return _ORIG
    return None


def file_kind(path):
    try:
        import pymupdf as fitz
        d = fitz.open(path)
        try:
            t = "".join(p.get_text() for p in d)
        finally:
            d.close()
    except Exception:
        t = ""
    k = _kind_from_text(t)
    if k:
        return k
    try:  # 扫描件
        import ocrlib
        return _kind_from_text(ocrlib.ocr_text(path, cache_dir=None, pages=[0]))
    except Exception:
        return None


def _stem(name):
    """去掉 .pdf 与已带的（原决定）/（复核决定）/（日期）后缀（后缀在扩展名之前，故先剥扩展名）。"""
    s = name[:-4] if name.lower().endswith(".pdf") else name
    s = re.sub(r"（(?:原决定|复核决定)）$", "", s)
    s = re.sub(r"（\d{4}-\d{2}-\d{2}）$", "", s)
    return s


def main(argv=None):
    ap = argparse.ArgumentParser(description="复核案文件名标注（原决定 / 复核决定）")
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args(argv)

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LS.LEDGER_NAME)
    if not os.path.isfile(ledger):
        print("!! 未找到台账：", ledger)
        return 2

    wb = load_workbook(ledger, data_only=False)
    ws = wb[LS.SHEET_CASE]

    plan = []          # (row, idx, old, new, path)
    skipped = []
    C_DEF = LS.COL["defense"] + 1
    for r in range(2, ws.max_row + 1):
        files = [x.strip() for x in str(ws.cell(r, C_FILES).value or "").split("；") if x.strip()]
        if len(files) < 2:
            continue
        # 闸门：只有「真复核案」才值得逐份 OCR 判原/复核（否则全表 OCR 要跑几十分钟）。
        # 复核案特征：申辩列有【复核决定 标注（见 0.5.7），或文件名里已带「复核」。
        defense = str(ws.cell(r, C_DEF).value or "")
        if "【复核决定" not in defense and not any("复核" in f for f in files):
            continue
        paths = [x.strip() for x in str(ws.cell(r, C_PATHS).value or "").split("；")]
        kinds = []
        for i, f in enumerate(files):
            p = paths[i] if i < len(paths) else (paths[-1] if paths else "")
            fp = os.path.join(root, p, f)
            kinds.append(file_kind(fp) if os.path.isfile(fp) else None)
        if _REVIEW not in kinds or _ORIG not in kinds:
            continue
        for i, f in enumerate(files):
            k = kinds[i]
            if not k:
                skipped.append((r, f, "判不出原/复核"))
                continue
            new = f"{_stem(f)}（{k}）.pdf"
            if new == f:
                continue
            p = paths[i] if i < len(paths) else (paths[-1] if paths else "")
            plan.append((r, i, f, new, p))

    print("=" * 76)
    print(f"台账：{ledger}")
    print(f"复核案需标注：{len(plan)} 个文件｜判不出：{len(skipped)}")
    print("=" * 76)
    for r, _i, old, new, _p in plan:
        print(f"行{r}: {old}\n     → {new}")
    for r, f, why in skipped:
        print(f"  !! 行{r} {f} — {why}")

    if not a.apply:
        print("\n[干跑] 未落盘。确认无误后加 --apply。")
        return 0

    if not plan:
        print("\n无需改动。")
        return 0

    wd = LS.work_dir(root)
    bk = os.path.join(wd, "pdf_bak_reviewlabel_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(bk, exist_ok=True)
    shutil.copy2(ledger, os.path.join(wd, "ledger_bak_reviewlabel_%s.xlsx"
                                      % datetime.now().strftime("%Y%m%d_%H%M%S")))

    fail, done, ren05 = [], 0, 0
    for r in sorted({x[0] for x in plan}):
        files = [x.strip() for x in str(ws.cell(r, C_FILES).value or "").split("；") if x.strip()]
        for rr, i, old, new, p in [x for x in plan if x[0] == r]:
            src = os.path.join(root, p, old)
            dst = os.path.join(root, p, new)
            if not os.path.isfile(src):
                fail.append((r, old, "源文件不存在"))
                continue
            if os.path.exists(dst):
                fail.append((r, old, "目标已存在"))
                continue
            try:
                shutil.copy2(src, os.path.join(bk, p.replace("/", "_") + "__" + old))
                os.rename(src, dst)
                files[i] = new
                done += 1
            except Exception as e:  # pragma: no cover
                fail.append((r, old, str(e)))
        ws.cell(r, C_FILES).value = "；".join(files)

    for r in sorted({x[0] for x in plan}):
        for _rr, _i, old, new, _p in [x for x in plan if x[0] == r]:
            for d, _sub, fs in os.walk(os.path.join(root, T05)):
                if old in fs and new not in fs:
                    try:
                        os.rename(os.path.join(d, old), os.path.join(d, new))
                        ren05 += 1
                    except Exception as e:  # pragma: no cover
                        fail.append((r, "05:" + old, str(e)))

    wb.save(ledger)
    print(f"\n已标注 {done} 个（05 副本 {ren05} 个）｜失败 {len(fail)}")
    for r, f, why in fail:
        print(f"  !! 行{r} {f}: {why}")
    print("备份目录：", bk)
    return 0 if not fail else 1


if __name__ == "__main__":
    raise SystemExit(main())
