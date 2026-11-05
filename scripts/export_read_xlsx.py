# -*- coding: utf-8 -*-
"""导出一份**在 Excel 里能用原生滚轮滚到底**的全文阅读表。

## 为什么是这个形态

用户要的动作是「鼠标停在内容上，滚轮往下，把这一条看完」。

在 Excel 里，**唯一 100% 能被滚轮滚动的，是工作表本身**。
批注框、ActiveX 文本框、UserForm 里的 TextBox 全都试过，都收不到滚轮
（2026-11-04 实测：批注 VML 里 `scroll` 出现 0 次；嵌入文本框滚的是表格；
UserForm 发滚轮后截图差异 19 像素，内容没动）。

所以把内容搬进**单元格**，让用户滚工作表——这是 Excel 的原生行为，
不需要宏、不需要控件、不看 Office 激活状态。

## 版式

一条案例占若干行：

    第 445 行 | 2025-12-12 | 深圳海峡资产管理有限公司；孙俊杰 | 证监局    ← 加粗+底色
    【处罚措施】…
    【核心处罚事由】
    一、未采取问卷调查…
    二、未按照合同约定…
    【核心处罚依据】
    【法条全文】
    …逐行原文…

- 法条全文本来就是「一行条款 + 一行条文」的硬换行，正好一行一格
- 长段落按估算行高撑开，保证不被截断（Excel 行高上限 409 磅，所以必须拆行）
- 首行冻结、A 列固定行号，滚到哪里都知道是哪一条

## 用法

    python export_read_xlsx.py --root <库根>
    python export_read_xlsx.py --root <库根> --out <输出.xlsx>
"""

from __future__ import annotations

import argparse
import math
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from export_basis_html import collect

OUT_NAME = "04c_全文阅读_Excel内可滚.xlsx"

# 列宽（字符数）——B 列正文 110，约合 55 个汉字
COL_W = {"A": 10, "B": 110, "C": 14}
CHARS_PER_LINE = 53          # 110 字符宽约放 53 个汉字
LINE_PT = 14.5               # 每行文字高度（磅）
MAX_ROW_PT = 400.0           # Excel 行高上限 409.5，留余量

F_TITLE = Font(name="Microsoft YaHei", size=11, bold=True, color="8A5A00")
F_HEAD = Font(name="Microsoft YaHei", size=10, bold=True, color="B42318")
F_BODY = Font(name="Microsoft YaHei", size=10, color="1F2328")
F_DIM = Font(name="Microsoft YaHei", size=9, color="6A737D")
FILL_TITLE = PatternFill("solid", fgColor="FFF6DA")
FILL_HEAD = PatternFill("solid", fgColor="F7F8FA")
AL_WRAP = Alignment(wrap_text=True, vertical="top")
AL_TOP = Alignment(vertical="top")
THIN = Side(style="thin", color="E3E7EC")
BORDER = Border(bottom=THIN)


def _h(text: str) -> float:
    """按字数估算需要的行高（磅）。超过上限就截到上限——宁可略挤，不可撑爆。"""
    n = max(1, math.ceil(len(text) / CHARS_PER_LINE))
    return min(MAX_ROW_PT, n * LINE_PT + 4)


def build(items: list[dict], out: str) -> tuple[int, int]:
    wb = Workbook()
    ws = wb.active
    ws.title = "全文阅读"
    for col, w in COL_W.items():
        ws.column_dimensions[col].width = w

    # 表头
    heads = ["台账行", "内容（鼠标停在任意位置，滚轮往下即可）", "来源"]
    for i, t in enumerate(heads, start=1):
        c = ws.cell(row=1, column=i, value=t)
        c.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="8A5A00")
        c.alignment = AL_TOP
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    r = 2
    n_case = 0
    for d in items:
        n_case += 1

        # —— 标题行：一条案例的开头，滚过去一眼能认出来 ——
        ws.cell(row=r, column=1, value="第 %s 行" % d["r"]).font = F_TITLE
        ws.cell(row=r, column=1).alignment = AL_TOP
        title = "%s　%s" % (d["d"] or "", d["s"] or "（无主体）")
        c = ws.cell(row=r, column=2, value=title)
        c.font = F_TITLE
        c.alignment = AL_WRAP
        c.fill = FILL_TITLE
        ws.cell(row=r, column=3, value=d["src"] or "").font = F_DIM
        for cc in range(1, 4):
            ws.cell(row=r, column=cc).fill = FILL_TITLE
            ws.cell(row=r, column=cc).border = BORDER
        ws.row_dimensions[r].height = _h(title)
        r += 1

        def block(label: str, body: str):
            """写一个小标题 + 正文（正文按原文换行拆格，保证每格不被截断）。"""
            nonlocal r
            if not body:
                return
            c = ws.cell(row=r, column=2, value=label)
            c.font = F_HEAD
            c.fill = FILL_HEAD
            c.alignment = AL_TOP
            ws.row_dimensions[r].height = 17
            r += 1
            for ln in str(body).split("\n"):
                ln = ln.strip()
                if not ln:
                    continue
                c = ws.cell(row=r, column=2, value=ln)
                c.font = F_BODY
                c.alignment = AL_WRAP
                ws.row_dimensions[r].height = _h(ln)
                r += 1

        block("【处罚措施】", d["m"])
        block("【核心处罚事由】", d["g"])
        block("【核心处罚依据】", d["h"])
        block("【法条全文】", d["text"])

        # 条目之间留一行空隙
        r += 1

    wb.save(out)
    wb.close()
    return n_case, r


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="导出 Excel 内可用原生滚轮阅读的全文表")
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--out", help="输出 xlsx（默认放 <库父>/_案例库维护留痕/）")
    a = ap.parse_args(argv)

    root = os.path.abspath(a.root)
    items, _n_over = collect(root)
    if not items:
        raise SystemExit("台账里没有任何批注，先跑 annotate_basis.py --apply")

    out = a.out or os.path.join(
        os.path.dirname(root), "_案例库维护留痕", OUT_NAME)
    os.makedirs(os.path.dirname(out), exist_ok=True)

    n_case, last_row = build(items, out)
    print("案例 %d 条｜表格 %d 行｜输出 %.1f KB" % (
        n_case, last_row - 1, os.path.getsize(out) / 1024))
    print("→ %s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
