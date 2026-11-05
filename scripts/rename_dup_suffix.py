# -*- coding: utf-8 -*-
"""把文件名里的「（2）（3）…」序号后缀改成**落款日期**。

背景（2026-11-02 用户实拍）：同一机构有多份决定书（原处分 + 复核决定、多批次监管措施）时，
下载器为避免覆盖自动加「（2）」，落到台账/文件夹里就成了「台州沃源…（2）.pdf」——
既看不出是哪一份，又让人误以为是「其他分类」的副本。全库实测 47 个。

口径：
  - 新名 = 原名去掉（N）+「（YYYY-MM-DD）」，日期取**该 PDF 自己的落款日**
    （文字层取最后一页日期；扫描件走 sign_date.find_sign 的 OCR），
    不用台账行日期——同一行可能同时挂着原处分与复核两份，日期不同；
  - 目标已存在则加「-2」，绝不覆盖；
  - 同步改名 `05_重点典型案例` 里的同名副本；
  - 幂等：文件名已不含（N）即跳过。

用法：
  python rename_dup_suffix.py --root <库根>            # 干跑
  python rename_dup_suffix.py --root <库根> --apply    # 落盘（备份台账 + 被改名 PDF）
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

_DUP = re.compile(r"（(\d+)）(?=\.pdf$)", re.I)
T05 = "05_重点典型案例"


def pdf_sign_date(path):
    """返回该 PDF 的落款日 ISO 字符串；取不到返回 None。

    顺序：文字层末段 → OCR 末两页（扫描件实测占少数）→ sign_date.find_sign 精细定位。
    """
    import sign_date as SD
    txt = ""
    pages = 0
    try:
        import pymupdf as fitz
        d = fitz.open(path)
        try:
            pages = d.page_count
            txt = "".join(p.get_text() for p in d)
        finally:
            d.close()
    except Exception:
        txt = ""
    ds = SD.parse_dates(txt[-1500:] if txt else "", max_date=SD.today_iso())
    if ds:
        return max(ds)
    if pages:
        try:
            import ocrlib
            for pno in [pages - 1, pages - 2]:
                if pno < 0:
                    continue
                raw = ocrlib.ocr_text(path, cache_dir=None, pages=[pno]) or ""
                ds = SD.parse_dates(raw, max_date=SD.today_iso())
                if ds:
                    return max(ds)
        except Exception:
            pass
    try:
        dt, _plan, _raw = SD.find_sign(path)
        return dt
    except Exception:
        return None


def main(argv=None):
    ap = argparse.ArgumentParser(description="把（2）（3）…文件名后缀改为落款日期")
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

    todo, nodate = [], []
    for r in range(2, ws.max_row + 1):
        files = [x.strip() for x in str(ws.cell(r, C_FILES).value or "").split("；") if x.strip()]
        paths = [x.strip() for x in str(ws.cell(r, C_PATHS).value or "").split("；")]
        if not any(_DUP.search(f) for f in files):
            continue
        for i, f in enumerate(files):
            if not _DUP.search(f):
                continue
            p = paths[i] if i < len(paths) else (paths[-1] if paths else "")
            fp = os.path.join(root, p, f)
            if not os.path.isfile(fp):
                nodate.append((r, f, "文件不存在"))
                continue
            d = pdf_sign_date(fp)
            if not d:
                nodate.append((r, f, "取不到落款日"))
                continue
            base = _DUP.sub("", f)
            stem = base[:-4] if base.lower().endswith(".pdf") else base
            new = f"{stem}（{d}）.pdf"
            todo.append((r, i, f, new, p))

    print("=" * 76)
    print(f"台账：{ledger}")
    print(f"含（N）后缀待改名：{len(todo)} 个文件｜无法处理：{len(nodate)} 个")
    print("=" * 76)
    for r, _i, old, new, p in todo:
        print(f"行{r}: {old}\n     → {new}")
    if nodate:
        print("\n!! 无法改名（需人工）：")
        for r, f, why in nodate:
            print(f"  行{r}: {f} — {why}")

    if not a.apply:
        print("\n[干跑] 未落盘。确认无误后加 --apply。")
        return 0 if not nodate else 1

    wd = LS.work_dir(root)
    bk = os.path.join(wd, "pdf_bak_dupsuffix_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(bk, exist_ok=True)
    shutil.copy2(ledger, os.path.join(wd, "ledger_bak_dupsuffix_%s.xlsx"
                                      % datetime.now().strftime("%Y%m%d_%H%M%S")))

    by_row = {}
    for r, i, old, new, p in todo:
        by_row.setdefault(r, []).append((i, old, new, p))

    fail, done = [], 0
    for r, items in by_row.items():
        files = [x.strip() for x in str(ws.cell(r, C_FILES).value or "").split("；") if x.strip()]
        for i, old, new, p in items:
            src = os.path.join(root, p, old)
            dst = os.path.join(root, p, new)
            if not os.path.isfile(src):
                fail.append((r, old, "源文件不存在"))
                continue
            if os.path.exists(dst):
                stem = new[:-4]
                k = 2
                while os.path.exists(os.path.join(root, p, f"{stem}-{k}.pdf")):
                    k += 1
                new = f"{stem}-{k}.pdf"
                dst = os.path.join(root, p, new)
            try:
                shutil.copy2(src, os.path.join(bk, p.replace("/", "_") + "__" + old))
                os.rename(src, dst)
                files[i] = new
                done += 1
            except Exception as e:  # pragma: no cover
                fail.append((r, old, str(e)))
        ws.cell(r, C_FILES).value = "；".join(files)

    # 05 副本同步改名（同名副本）
    ren05 = 0
    for r, items in by_row.items():
        for _i, old, new, _p in items:
            for d, _sub, fs in os.walk(os.path.join(root, T05)):
                if old in fs and new not in fs:
                    try:
                        os.rename(os.path.join(d, old), os.path.join(d, new))
                        ren05 += 1
                    except Exception as e:  # pragma: no cover
                        fail.append((r, "05:" + old, str(e)))

    wb.save(ledger)
    print(f"\n已改名 {done} 个（05 副本 {ren05} 个）｜失败 {len(fail)}")
    for r, f, why in fail:
        print(f"  !! 行{r} {f}: {why}")
    print("备份目录：", bk)
    return 0 if not fail else 1


if __name__ == "__main__":
    raise SystemExit(main())
