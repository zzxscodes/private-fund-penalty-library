# -*- coding: utf-8 -*-
"""补全被截断的机构全称（如「杭州复投私募基金管理合伙企业」→「…（有限合伙）」）。

背景（2026-11-02 用户裁定）：部分证监局文书的 PDF 文件名只到字号主体，丢了「（有限合伙）」，
台账 D 列主体与 12 列文件名、实际 PDF 名随之被截断，读起来像另一家公司。

口径：
  - **以决定书正文全称为准**（正文「关于对 XXX（有限合伙）采取…」），不用文件名反推；
  - 只处理「结尾是 合伙企业/管理中心/中心/企业/基金 且无括注」的疑似截断，逐个回正文验证，认不出就跳过并报告；
  - 改名同时更新：D 列主体（首段）+ L 列文件名 + 05 典型副本文件名；PDF 内容不动；
  - 幂等：已含括注的不再处理；重复跑应全 0。

用法：
  python complete_org_names.py --root <库根>            # 干跑，只报告
  python complete_org_names.py --root <库根> --apply    # 落盘（备份台账 + 备份被改名 PDF）
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

C_SUBJ = LS.COL["subject"] + 1
C_FILES = LS.COL["filenames"] + 1
C_PATHS = LS.COL["paths"] + 1
C_URLS = LS.COL["urls"] + 1

# 疑似截断：以这些词收尾，且整体没有括注
_TAIL = re.compile(r"(合伙企业|管理中心|中心|企业|基金|事务所)$")
_SUFFIXES = ("（有限合伙）", "（普通合伙）", "（有限合伙企业）", "(有限合伙)", "(普通合伙)")


def _norm(s):
    return re.sub(r"[\s\u3000\u200b]", "", str(s or ""))


def read_pdf_text(path):
    """取正文；无文字层（扫描件）→ OCR 首页（实测 2 例：珠海通沛、深圳知行通达）。"""
    try:
        import pymupdf as fitz
    except Exception:  # pragma: no cover
        import fitz
    txt = ""
    try:
        doc = fitz.open(path)
        try:
            txt = "".join(p.get_text() for p in doc)
        finally:
            doc.close()
    except Exception:
        txt = ""
    if len(re.sub(r"\s", "", txt)) >= 60:
        return txt
    try:  # 扫描件兜底
        import ocrlib
        return ocrlib.ocr_text(path, cache_dir=None, pages=[0]) or txt
    except Exception:
        return txt


def find_full_name(text, org):
    """从正文里找机构全称：优先「org + 括注」，其次正文中同前缀的括注名。认不出返回 None。"""
    t = _norm(text)
    o = _norm(org)
    if not o:
        return None
    for suf in _SUFFIXES:
        if o + _norm(suf) in t:
            return org + "（" + re.sub(r"[（）()]", "", suf) + "）"
    # 正文里 else 形式：「杭州复投私募基金管理合伙企业 （有限合伙）」已由上面覆盖；
    # 兜底：同前缀 + 任意 2~8 字括注
    m = re.search(re.escape(o) + r"[（(]([^）)]{1,8})[）)]", t)
    if m and re.search(r"(合伙|中心|企业|基金)", m.group(1)):
        return org + "（" + m.group(1) + "）"
    return None


def main(argv=None):
    ap = argparse.ArgumentParser(description="补全被截断的机构全称（以决定书正文为准）")
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

    todo, unresolved = [], []
    for r in range(2, ws.max_row + 1):
        subj = str(ws.cell(r, C_SUBJ).value or "")
        org = subj.split("；")[0].strip()
        if not org or not _TAIL.search(org) or re.search(r"[（(]", org):
            continue
        paths = str(ws.cell(r, C_PATHS).value or "").split("；")
        files = str(ws.cell(r, C_FILES).value or "").split("；")
        full, src = None, ""
        for p, f in zip(paths, files):
            fp = os.path.join(root, p.strip(), f.strip())
            if not os.path.isfile(fp):
                continue
            full = find_full_name(read_pdf_text(fp), org)
            src = f
            if full:
                break
        if full:
            todo.append((r, org, full, files, paths, subj))
        else:
            unresolved.append((r, org, src))

    print("=" * 72)
    print(f"台账：{ledger}")
    print(f"疑似截断机构名：{len(todo) + len(unresolved)} 行｜可在正文确认全称：{len(todo)} 行｜认不出：{len(unresolved)} 行")
    print("=" * 72)
    for r, org, full, files, _paths, _subj in todo:
        print(f"行{r}: {org}  →  {full}")
        for f in files:
            if org in f:
                print(f"        文件：{f}  →  {f.replace(org, full)}")
    if unresolved:
        print("\n!! 以下行正文里没找到「（有限合伙）」式全称，需人工按 PDF 确认：")
        for r, org, src in unresolved:
            print(f"  行{r}: {org}（{src}）")

    if not a.apply:
        print("\n[干跑] 未落盘。确认无误后加 --apply。")
        return 0

    if not todo:
        print("\n无需改动。")
        return 0

    wd = LS.work_dir(root)
    bk = os.path.join(wd, "pdf_bak_orgname_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(bk, exist_ok=True)
    shutil.copy2(ledger, os.path.join(wd, "ledger_bak_orgname_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")))
    print("\n备份目录：", bk)

    os.makedirs(bk, exist_ok=True)
    fail = []
    done = 0
    for r, org, full, files, paths, subj in todo:
        new_files, ok = [], True
        for p, f in zip(paths, files):
            old_fp = os.path.join(root, p.strip(), f.strip())
            new_name = f.replace(org, full)
            new_fp = os.path.join(root, p.strip(), new_name.strip())
            if not os.path.isfile(old_fp):
                new_files.append(f)
                continue
            if os.path.abspath(old_fp) == os.path.abspath(new_fp):
                new_files.append(new_name)
                continue
            if os.path.isfile(new_fp):  # 目标已存在 → 不覆盖，记失败
                ok = False
                fail.append((r, f, "目标已存在"))
                new_files.append(f)
                continue
            try:
                shutil.copy2(old_fp, os.path.join(bk, f.strip()))  # 先备份再改名
                os.rename(old_fp, new_fp)
                new_files.append(new_name)
            except Exception as e:  # pragma: no cover
                ok = False
                fail.append((r, f, str(e)))
                new_files.append(f)
        if not ok:
            continue
        # 05 典型副本同步改名
        t05 = os.path.join(root, LS.TYPICAL_DIR) if hasattr(LS, "TYPICAL_DIR") else os.path.join(root, "05_重点典型案例")
        for d, _sub, fs in os.walk(t05):
            for f in fs:
                if f in [x.strip() for x in files] and org in f:
                    src_p = os.path.join(d, f)
                    dst_p = os.path.join(d, f.replace(org, full))
                    if not os.path.exists(dst_p):
                        try:
                            shutil.copy2(src_p, os.path.join(bk, "05_" + f))
                            os.rename(src_p, dst_p)
                        except Exception as e:  # pragma: no cover
                            fail.append((r, "05:" + f, str(e)))

        tail = subj[len(org):] if subj.startswith(org) else ("；" + subj.split("；", 1)[1] if "；" in subj else "")
        ws.cell(r, C_SUBJ).value = full + tail
        ws.cell(r, C_FILES).value = "；".join(x.strip() for x in new_files)
        done += 1

    wb.save(ledger)
    print(f"已补全并落盘：{done} 行｜失败 {len(fail)} 项")
    for r, f, why in fail:
        print(f"  !! 行{r} {f}: {why}")
    return 0 if not fail else 1


if __name__ == "__main__":
    raise SystemExit(main())
