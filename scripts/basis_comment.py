# -*- coding: utf-8 -*-
"""H 列批注文本生成器（basis_comment）

一套逻辑两处用：
  - `annotate_basis.py` 写批注；
  - `audit_basis_comments.py`（Z15）重算一遍与台账里的批注比对，防止 06 更新后批注变成旧文。

批注形态（2026-11-03 用户要求：**只放条文本身**，不要生成出来的说明文字）：

    《私募投资基金监督管理暂行办法》第二十三条第九项
    私募基金管理人……不得有以下行为：
    （一）将其固有财产或者他人财产混同于基金财产从事投资活动；
    ……

    《私募投资基金监督管理暂行办法》第三十三条
    私募基金管理人……违反法律、行政法规及本办法规定……

即：每条 = 一行「法规名 + 本行引用的条/款/项」+ 条文正文；条与条之间空一行。
不写抬头、不写编号、不写脚注、不加 ▶ 标记——要确认引的是哪一项，看标题行即可。
"""

from __future__ import annotations

from lawtext import AMBIGUOUS, art_sort_key, norm_name  # 提示用：条号越界 / 简称歧义

MAX_ART_CHARS = 6000       # 单条上限（Excel 批注上限 32767 字，这里只防病态输入）

# 批注框尺寸（像素）。Excel 的悬停框不能滚动，超出可视区的部分直接看不到，
# 所以宽度取「换到任何一台机器都不越界」，高度按内容行数估算、封顶在屏高以内。
#
# 一、宽度为什么是 960（2026-11-04 用户实拍反馈定案）
#   曾经取 1300px：那是按本机 2560×1440 / 100% 缩放定的（可用高约 1150px、
#   H 列左侧横向约 1987px）。换到笔记本上就露馅了——**框比可视区还宽**，
#   右边界跑到屏幕外，用户看到的是「一行老长、右边文字看不全」，很像没有自动换行。
#   先做了对照实验再动手：让 Excel 自己 AddComment 一条、存成 xlsx、解包看 VML，
#   原生 shape 与我们的结构**逐项相同**（`<v:textbox style='mso-direction-alt:auto'>`
#   里都是空 div，全文 `wrap` 出现 0 次、`WordWrap` 0 次；
#   `Comment.Shape.TextFrame.WordWrap` 直接报 AttributeError，`TextFrame2.WordWrap` 报
#   0x800A01A8）。**结论：批注的自动换行本来就在，且没有任何文件级开关能设**，
#   症状的根因是框宽，不是换行。
#   故取 960px：1366×768 笔记本上 Excel 可用宽约 1315px、A 列起点约 43px，
#   右边界 1003px 仍留 300px 余量。代价：每行容量 212→155 字，超过 1150px 高的
#   条目由 5 条变 6 条（实测，几乎无损），其余全部一屏可读。
#
# 二、为什么不干脆顶到 1987（本机上限）
#   试算过 1300/1450/1600/1750/1900/1980 六档：批注正文是「一行条款 + 一行正文」的
#   **硬换行**，行数由换行决定而非宽度，加宽到 1980 也只能救 5 行里的 1 行（208）。
#   宽度不是瓶颈，高度受屏高限制——真装不下的那几条只能换载体，
#   出口是库根的 `04_综合索引台账【总表】.html`（可滚动、可检索，见 SKILL 0.5.24）。
#
# 三、换机器要重算什么
#   在目标机取屏幕分辨率与缩放；可用高 ≈ 屏高 − 功能区 − 任务栏 − 标签栏。
#   宽度不要跟着屏宽放大——按最小目标机（1366 宽）取值才是最稳的。
BOX_W = 960
BOX_MAX_H = 1150


def _xiang_label(x: str) -> str:
    return "（%s）" % x if x else ""


def _units(s: str) -> int:
    """按显示宽度折算：汉字/全角算 2，ASCII 算 1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


FONT_MAX = 9.0        # Excel 批注默认字号（Tahoma 9）
FONT_MIN = 7.0        # 缩字号时的下限：再小就看不清

# 默认**不缩字号**（2026-11-03 用户回退）：批注正文统一 9pt，超长的让用户自己滚动看。
# 想把长条压进一屏，跑 `annotate_basis.py --shrink --apply` 临时打开，不要改默认值。
SHRINK = False


def _height(text: str, font: float) -> float:
    """给定字号估算框高（像素）。"""
    unitw = 6.0 * font / 9.0            # 1 显示单位占多少像素
    usable = (BOX_W - 28) / unitw       # 一行能放多少显示单位
    lh = 14.5 * font / 9.0              # 行高随字号等比缩放
    lines = 0
    for ln in text.split("\n"):
        u = _units(ln)
        lines += max(1, int((u + usable - 1) // usable))
    return (26 * font / 9.0 + lines * lh) * 1.08 + 14


def fit(text: str) -> tuple:
    """估算批注框 → (宽, 高, 字号)。超长自动降字号，直到塞进一屏。

    **为什么降字号**：Excel 的悬停框**不能滚动**（只读预览，平台限制），塞不下就是看不全。
    与其让用户右键「编辑批注」（要离开当前格、还多两步），不如把字缩小一点让它一屏排开。
    实测分布（508 行）：488 行 9pt 即可；20 行需要缩小；最低 7pt，其中 6 行 7pt 仍超框。
    """
    f = FONT_MAX
    h = _height(text, f)
    if SHRINK:
        while h > BOX_MAX_H and f > FONT_MIN:
            f = max(FONT_MIN, f - 0.5)
            h = _height(text, f)
    return BOX_W, max(90, min(BOX_MAX_H, int(h))), f


def fit_size(text: str) -> tuple:
    """向后兼容：只要 (宽, 高)。"""
    w, h, _ = fit(text)
    return w, h

def build(index, h: str) -> tuple:
    """→ (批注文本, 统计dict)。统计含 laws/arts/missing/inferred。"""
    items = index.parse_h(h)
    order, entries, missing = [], {}, []
    for it in items:
        law = it["law"]
        if not law:
            missing.append(it)
            continue
        if not it["arts"]:
            key = (law, "")
            if key not in entries:
                entries[key] = {"kuan": "", "xiang": [], "inferred": it["inferred"],
                                "raw": it["raw"]}
                order.append(key)
            continue
        for a in it["arts"]:
            key = (law, a["no"])
            if key not in entries:
                entries[key] = {"kuan": "", "xiang": [], "inferred": it["inferred"],
                                "raw": it["raw"]}
                order.append(key)
            if a["kuan"] and not entries[key]["kuan"]:
                entries[key]["kuan"] = a["kuan"]
            if a["xiang"] and a["xiang"] not in entries[key]["xiang"]:
                entries[key]["xiang"].append(a["xiang"])

    laws = sorted({k[0] for k in order})
    blocks, n_ok, n_miss = [], 0, 0
    for key in order:
        law, no = key
        meta = entries[key]
        arts = index.laws.get(law, {}).get("arts", {})
        if not no:
            blocks.append("《%s》（本行只写了法规名，未指明条款；全文见 06 目录）" % law)
            n_ok += 1
            continue

        # 标题行只写「法规名 + 条/款/项」，与 H 列的引用一一对应
        title = "《%s》%s" % (law, no)
        if meta["xiang"]:
            title += "第%s项" % "、".join(meta["xiang"])
        elif meta["kuan"]:
            title += "第%s款" % meta["kuan"]

        body = arts.get(no)
        if not body:
            nums = [k for k in (art_sort_key(x) for x in arts) if k]
            mx = max(nums)[0] if nums else 0
            k = art_sort_key(no)
            if k and mx and k[0] > mx:
                why = "06 库内该法规共 %d 条，无此条（引用条号超出该版本）" % len(arts)
            else:
                why = "06 库内该法规未见此条"
            blocks.append("%s（%s，无正文）" % (title, why))
            n_miss += 1
            continue

        meta_law = index.laws.get(law, {})
        if meta_law.get("reliable") is False:
            blocks.append("%s（06 收录为扫描件且条号乱序，无法可靠抽取，见原文）" % title)
            n_miss += 1
            continue

        text = body
        if len(text) > MAX_ART_CHARS:
            text = text[:MAX_ART_CHARS] + "\n……（条文过长，此处截断，全文见 06 目录）"
        blocks.append(title + "\n" + text)
        n_ok += 1

    for it in missing:
        raw = it["raw"]
        n = norm_name(raw)
        cands = [t for t in index.titles if n and t.endswith(n)] if (
            n in AMBIGUOUS or len(n) <= 4) else []
        if len(cands) > 1:
            blocks.append("《%s》（简称不唯一，可能是：%s；请补全称）"
                          % (raw, "／".join(cands[:3])))
        else:
            blocks.append("《%s》（06 目录未收录，无正文）" % raw)
        n_miss += 1

    stat = {"laws": len(laws), "arts": n_ok, "missing": n_miss,
            "inferred": sum(1 for k in order if entries[k]["inferred"])}
    return "\n\n".join(blocks), stat
