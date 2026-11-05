#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_cells_20261102.py —— 逐格机械清理（空格/叠字/分号/括号），2026-11-02

范围（只做**不需读原文**就能确定对错的）：
  Y1 半角空格/全角空格/制表符污染  —— D~I 六列（F 192、H 504、G 13、E 少量）
  Y2 叠字                        —— E 列 2 处（信息信息披露 / 私募私募基金产品）
  Y3 措施列以分号结尾              —— E 列 3 处，分号改句号
  Y4 D 列机构名缺右括号            —— 4 处「（有限合伙」缺「）」，同步 L 列 + 磁盘文件名 + 05 副本

**不做**（必须读决定书原文才能改，交给 fix_f_from_pdf_20261102.py）：
  - 信用代码残片（5 处）：删了代码主语就丢了，只能按原文重抽
  - 网页导语句（17 处）、整段未分点（136 处）、叙述引导词开头（11 处）

空格清理的判据（为什么不是简单 `re.sub`）：
  两侧都是 ASCII 字母数字/括号（`fund A`、`AI APP`）才保留 1 个空格；
  其余（中文↔中文、中文↔数字、书名号↔第X条、引号↔汉字）一律删——
  中文排版里这些空格全是抽取副产物，删掉才干净。**换行 \\n 不动**（分条要用）。

用法：
  python scripts/fix_cells_20261102.py --root <库根>              # 干跑
  python scripts/fix_cells_20261102.py --root <库根> --apply      # 落盘
"""
import argparse
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402
from ledger_bak import backup_ledger  # noqa: E402

# 只清理长文本列：D主体 E措施 F事由 G证据 H依据 I申辩
# 注意：openpyxl 的 ws.cell(row, col) 是 **1-based**，D=4 … I=9
COLS = [4, 5, 6, 7, 8, 9]
COLNAME = {4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I"}

DUP_FIX = {
    "信息信息披露": "信息披露",
    "私募私募基金": "私募基金",
}


def tighten(s):
    """删掉中文语境里的空格，保留纯 ASCII 词间空格"""
    if not s:
        return s

    def repl(m):
        a = s[m.start() - 1] if m.start() > 0 else ""
        b = s[m.end()] if m.end() < len(s) else ""
        ascii_word = lambda ch: ch and ch.isascii() and (ch.isalnum() or ch in ")]}")
        if ascii_word(a) and ascii_word(b):
            return " "
        return ""

    s = re.sub(r"[ \t\u3000]+", repl, s)
    # 半角标点 → 全角（数字之间、时间 12:30 保留）
    def rp(m):
        ch = m.group(0)
        i = m.start()
        a = s[i - 1] if i else ""
        b = s[i + 1] if i + 1 < len(s) else ""
        if a.isdigit() and b.isdigit():
            return ch
        return {",": "，", ";": "；", ":": "："}[ch]
    return re.sub(r"[,;:]", rp, s)


def fix_dup(s):
    out = s
    for bad, good in DUP_FIX.items():
        out = out.replace(bad, good)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--show", type=int, default=6)
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]

    n_sp = n_dup = n_semi = 0
    shown = 0

    # ---- Y1 空格 + Y2 叠字 + Y3 分号 ----
    for i in range(2, ws.max_row + 1):
        if not ws.cell(i, 1).value:
            continue
        for c in COLS:
            cell = ws.cell(i, c)
            v = cell.value
            if not isinstance(v, str) or not v:
                continue
            new = tighten(v)
            if new != v:
                n_sp += 1
                if shown < args.show:
                    # 找出第一处差异上下文
                    k = next((j for j in range(min(len(v), len(new)))
                              if v[j] != new[j]), 0)
                    print("  [空格] 行%-4d %s：%r → %r" % (
                        i, COLNAME[c], v[max(0, k - 10):k + 18],
                        new[max(0, k - 10):k + 18]))
                    shown += 1
                if args.apply:
                    cell.value = new
                v = new

            nv = fix_dup(v)
            if nv != v:
                n_dup += 1
                m = next((v.find(k) for k in DUP_FIX if k in v), -1)
                ctx = v[max(0, m - 10):m + 14] if m >= 0 else v[:40]
                print("  [叠字] 行%-4d %s：…%s…" % (i, COLNAME[c], ctx))
                if args.apply:
                    cell.value = nv
                v = nv

            if c == 5 and v and re.search(r"[；;]\s*$", v):
                n_semi += 1
                nv = re.sub(r"[；;]\s*$", "。", v)
                print("  [分号] 行%-4d E：…%r → …%r" % (i, v[-30:], nv[-30:]))
                if args.apply:
                    cell.value = nv

    print("\nY1 空格 %d 格 ｜ Y2 叠字 %d 格 ｜ Y3 分号 %d 格" % (n_sp, n_dup, n_semi))

    # ---- Y4 D 列缺右括号（同步 L 列 + 磁盘 + 05） ----
    fixes = []
    for i in range(2, ws.max_row + 1):
        d = ws.cell(i, 4).value
        if not isinstance(d, str):
            continue
        # 「（有限合伙」后缺「）」：出现在结尾或紧跟「；」
        m = re.search(r"（有限合伙(?=$|[；;、])", d)
        if not m:
            continue
        new_d = d[:m.end()] + "）" + d[m.end():]
        fixes.append((i, d, new_d))

    print("\nY4 D 列缺右括号：%d 处" % len(fixes))
    n_renamed = 0
    for i, old_d, new_d in fixes:
        old_org = old_d.split("；")[0].split("、")[0].strip()
        if not old_org.endswith("（有限合伙"):
            old_org = old_d.split("；")[0].strip()
        new_org = old_org + "）"
        lcell = ws.cell(i, 12)
        lv = lcell.value if isinstance(lcell.value, str) else ""
        path = ws.cell(i, 13).value or ""
        path = str(path).split("；")[0].strip()
        print("  行%-4d D：%r → %r" % (i, old_d, new_d))
        # 该行每个含旧名的文件名（机构件 + 各人员件）都要改名
        targets = []
        for nm in [x.strip() for x in lv.split("；") if x.strip()]:
            if old_org + ".pdf" in nm:
                targets.append((nm, nm.replace(old_org + ".pdf", new_org + ".pdf")))
        for old_name, new_name in targets:
            for base in (path, "05_重点典型案例"):
                if os.path.isfile(os.path.join(root, base, old_name)):
                    print("       %s改名：%s → %s" % ("待" if not args.apply else "",
                                                  old_name, new_name))
        if not args.apply:
            continue
        ws.cell(i, 4).value = new_d
        # E 列同步（若出现旧机构名）
        ec = ws.cell(i, 5)
        if isinstance(ec.value, str) and old_org in ec.value:
            ec.value = ec.value.replace(old_org, new_org)
        if not lv:
            continue
        new_lv = lv.replace(old_org + ".pdf", new_org + ".pdf")
        if new_lv == lv:
            continue
        lcell.value = new_lv
        # 磁盘改名（M 列是相对库根路径，逐个文件名改，含人员件）
        for old_name, new_name in targets:
            for base in (path, "05_重点典型案例"):
                src = os.path.join(root, base, old_name)
                dst = os.path.join(root, base, new_name)
                if not os.path.isfile(src):
                    continue
                try:
                    if os.path.exists(dst):
                        print("     ! 目标已存在，跳过：%s" % dst)
                        continue
                    os.rename(src, dst)
                    n_renamed += 1
                except OSError as e:
                    print("     × 改名失败 %s：%s" % (src, e))

    print("Y4 磁盘改名 %d 个" % n_renamed)

    if args.apply:
        backup_ledger(root, ledger, "cells_20261102")
        wb.save(ledger)
        print("\n已保存台账（备份在 _案例库维护留痕/台账备份/）")
    else:
        print("\n[干跑] 未落盘；加 --apply 执行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
