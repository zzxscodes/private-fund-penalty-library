# -*- coding: utf-8 -*-
"""给台账 H 列（核心处罚依据）写批注：鼠标悬停即显示所引法条的正文。

用法：
    python annotate_basis.py --root <库根>              # 干跑，只出报告
    python annotate_basis.py --root <库根> --apply      # 落盘（自动备份台账）

判据与护栏：
  - 只写批注，**不改任何单元格的值**；落盘后逐格比对备份，有一个字不同就报错退出；
  - 法条正文一律现取自 `06_以上各类所有案例的配套法规依据`，不内置任何抄本；
  - 解析不到的法规（06 未收录）如实写「未收录」，不编条文；
  - 批注文本由 `basis_comment.build()` 生成，`audit_basis_comments.py`（Z15）用同一函数复算比对。
"""

import argparse
import csv
import hashlib
import os
import shutil
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

LEDGER = "04_综合索引台账【总表】.xlsx"
REG = "06_以上各类所有案例的配套法规依据"


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def snapshot(ws):
    """(行, 列) → 值，用于「只加批注、不改值」的逐格比对。"""
    return {(r, c): ws.cell(r, c).value
            for r in range(1, ws.max_row + 1) for c in range(1, ws.max_column + 1)}


def _fmt_sz(f):
    """7.0 → "7"，7.5 → "7.5"（Excel 的 sz 写最短形式）。"""
    return ("%.1f" % f).rstrip("0").rstrip(".")


def set_comment_fonts(path, fonts):
    """把需要缩字号的批注正文改写成带 `rPr` 的 run（Excel 认这个字号）。

    `Comment` 没有字号属性，openpyxl 落盘只写 `<text><t>…</t></text>`，一律默认 9pt。
    长批注塞不进一屏时（悬停框**不能滚动**），缩字号是唯一办法。
    写法与 Excel 自己生成的一致：`<r><rPr><sz val="7"/><rFont val="Tahoma"/></rPr><t>…</t></r>`。
    实测 openpyxl 能把这种 run 原样读回（逐字相等），Z15 的逐字比对不受影响。
    """
    import re
    import shutil
    import zipfile

    if not fonts:
        return 0, {}

    def one(m):
        nonlocal n
        ref, attrs, inner = m.group(1), m.group(2), m.group(3)
        row = int(re.sub(r"\D", "", ref))
        f = fonts.get(row)
        if not f:
            return m.group(0)
        body = re.sub(r"^<text>|</text>$", "", inner.strip())
        body = re.sub(r"^<t[^>]*>", "", body)
        body = re.sub(r"</t>$", "", body)
        n += 1
        return ("<comment ref=\"%s\"%s><text><r><rPr><sz val=\"%s\"/>"
                "<rFont val=\"Tahoma\"/></rPr><t xml:space=\"preserve\">%s</t></r>"
                "</text></comment>") % (ref, attrs, _fmt_sz(f), body)

    tmp = path + ".fonttmp"
    n = 0
    pat = re.compile(r'<comment ref="([A-Z]+\d+)"([^>]*)>(.*?)</comment>', re.S)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/comments") and item.filename.endswith(".xml"):
                data = pat.sub(one, data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)

    dist = {}
    for f in fonts.values():
        dist[f] = dist.get(f, 0) + 1
    return n, dist


def _col_px(ws, c1):
    """列宽（字符）→ 像素。

    2026-11-04 修正：原来用 `w * 7 + 5`，比 Excel 实际少算约 14%——本库 A~G 实测
    横向 1987px（`w*8` 口径得 1952，`w*7+5` 只有 1743，差 14%）。
    锚点按小列宽算出来，Excel 一按真实列宽解释，框就被拉宽：
    设 BOX_W=960 实际显示 1060px，设 1300 时实际约 1435px。
    这也是「小屏上框超出屏幕」的隐藏放大器，改成 ×8 后设定的宽度与实际显示一致。
    """
    from openpyxl.utils import get_column_letter
    d = ws.column_dimensions.get(get_column_letter(c1))
    w = d.width if (d is not None and d.width) else 8.43
    return max(8, int(w * 8))


def _row_px(ws, r1):
    """行高（磅）→ 像素（1pt = 4/3 px）。"""
    d = ws.row_dimensions.get(r1)
    h = d.height if (d is not None and d.height) else 15.0
    return max(8, int(h * 4 / 3))


def _anchor(ws, col1, row1, w_px, h_px):
    """算出「左上角在本格、右下角盖住 w_px×h_px」的 VML 锚点。

    锚点格式（0 基）：`col, dx, row, dy, col2, dx2, row2, dy2`。
    按**实际列宽/行高**逐格累加，而不是假定 64×20 的默认格子——本台账列宽是调过的，
    用默认值算出来的锚点会把框拉宽或压窄。
    """
    c, x = col1, 0
    while x + _col_px(ws, c) < w_px and c < 16000:
        x += _col_px(ws, c)
        c += 1
    r, y = row1, 0
    while y + _row_px(ws, r) < h_px and r < 1000000:
        y += _row_px(ws, r)
        r += 1
    return "%d, 0, %d, 0, %d, %d, %d, %d" % (
        col1 - 1, row1 - 1, c - 1, max(0, int(w_px - x)), r - 1, max(0, int(h_px - y)))


def _anchor_left(ws, stop_col, row1, w_px, h_px, gap=6):
    """把框放在 `stop_col`（H 列）**左侧**，右边缘贴着该列左边界。

    为什么放左侧（2026-11-04 用户反馈「显示批注时点不到原格」）：
    框如果压在本行格子上，鼠标就落在框上，点不回 H 列那一格，也就收不起批注。
    放到 H 左边的 D~G 列上方，H 列始终露着，原格永远可点。
    """
    def cum(c):
        return sum(_col_px(ws, k) for k in range(1, c))

    right = cum(stop_col)                       # H 列左边界（像素）
    if w_px >= right:                           # 左边放不下就贴着最左
        left = 0
    else:
        left = max(0, right - int(w_px) - gap)
    c = 1
    while c < stop_col - 1 and sum(_col_px(ws, k) for k in range(1, c + 1)) <= left:
        c += 1
    dx1 = max(0, int(left - cum(c)))
    end_col = stop_col - 1                      # 右边缘落在 G 列
    dx2 = max(0, int(right - gap - cum(end_col)))
    r, y = row1, 0
    while y + _row_px(ws, r) < h_px and r < 1000000:
        y += _row_px(ws, r)
        r += 1
    return "%d, %d, %d, %d, %d, %d, %d, %d" % (
        c - 1, dx1, row1 - 1, 0, end_col - 1, dx2, r - 1, max(0, int(h_px - y)))


# H1 表头批注的说明书正文：**唯一真源**。
# 网页台账的入口路径写在这里，所以 `export_basis_html.py` 改了落点、或章节号有变，
# 都改这一处；单改 H1 用的补丁脚本也 import 它，避免两份文案各自漂移（2026-11-04）。
HEADER_HINT = (
    "鼠标停在本列任一格，即可查看该行引用的法条正文（取自 06 目录的法规原文）。\n"
    "若一屏没显示完：右键该格→「编辑批注」，可上下滚动看全文。\n"
    "只有 5 格确实装不下：第 89、208、212、400、444 行——\n"
    "双击本目录下的「04_综合索引台账【总表】.html」，搜行号即可看全文；\n"
    "它还能按来源、违规类型、涉案法规、地区、处罚措施交叉筛选。"
)


def fix_vml_size(path, ws=None):
    """把批注框改成 Excel 真正认的形态：pt 尺寸 + 锚点。

    **两件事缺一不可**，这是 2026-11-03 两次反馈叠加出来的完整根因：

    ① **单位是 pt 不是 px**：`Comment.width/height` 只是 openpyxl 的内存属性，落盘写成
       `<shape style="...width:880px...">`，Excel/WPS 认 pt，px 直接被忽略。
    ② **必须有 `<x:Anchor>`**：openpyxl 写了 `<x:SizeWithCells/>` 却**不写锚点**，
       Excel 打开时按自己的默认锚点重算尺寸，把 ①里改好的 pt 尺寸顶掉——
       症状是「连条文不多的格子也显示不全」（用户第二次反馈）。
       锚点算法与 xlsxwriter 一致：从本格左上角起，按实际列宽/行高累加到目标像素。

    必须放在 `wb.save()` **之后**：openpyxl 每次 save 都按默认值重生成 VML，先改会被覆盖。
    """
    import re
    import shutil
    import zipfile

    def conv(m):
        """`width:880px` → `width:660.0pt`。只换单位，不改数值关系（1px = 0.75pt）。"""
        nonlocal n_conv
        n_conv += 1
        return "%s:%.1fpt" % (m.group(1), float(m.group(2)) * 0.75)

    tmp = path + ".vmltmp"
    n_shape = n_conv = n_anchor = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".vml"):
                s = data.decode("utf-8")
                # 前缀可能是 ns0: 也可能是 v:，按正则数，别按字面量数
                n_shape += len(re.findall(r"<(?:\w+:)?shape\b", s))
                s = re.sub(r"(?<![\w-])(width|height):([\d.]+)px", conv, s)

                def one_shape(m):
                    nonlocal n_anchor
                    blk = m.group(0)
                    mw = re.search(r"(?<![\w-])width:([\d.]+)pt", blk)
                    mh = re.search(r"(?<![\w-])height:([\d.]+)pt", blk)
                    mr = re.search(r"<(?:\w+:)?Row>(\d+)<", blk)
                    mc = re.search(r"<(?:\w+:)?Column>(\d+)<", blk)
                    if ws is None or not (mw and mh and mr and mc):
                        return blk
                    # 固定在 H 列左侧；不再用本格当锚点，否则框会压住这一格
                    a = _anchor_left(ws, 8, int(mr.group(1)) + 1,
                                     float(mw.group(1)) / 0.75, float(mh.group(1)) / 0.75)
                    # 位置交给锚点，旧的行内边距会把框推偏
                    blk = re.sub(r"margin-left:[\d.]+pt;margin-top:[\d.]+pt;",
                                 "margin-left:0;margin-top:0;", blk)
                    pre = re.search(r"<((?:\w+:)?)ClientData", blk).group(1) or ""
                    tag = "<%sAnchor>%s</%sAnchor>" % (pre, a, pre)
                    if re.search(r"<(?:\w+:)?Anchor>", blk):
                        blk = re.sub(r"<(?:\w+:)?Anchor>[^<]*</(?:\w+:)?Anchor>", tag, blk)
                    else:
                        blk = re.sub(r"(<(?:\w+:)?AutoFill>)", lambda x: tag + x.group(1),
                                     blk, count=1)
                    n_anchor += 1
                    return blk

                s = re.sub(r"<(?:\w+:)?shape\b.*?</(?:\w+:)?shape>", one_shape, s, flags=re.S)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return n_shape, n_conv, n_anchor


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--ledger", default="", help="台账路径（默认 <root>/%s）" % LEDGER)
    ap.add_argument("--apply", action="store_true", help="落盘（默认只干跑）")
    ap.add_argument("--csv", default="", help="未解析清单导出 CSV")
    ap.add_argument("--no-ocr", action="store_true", help="不 OCR（扫描版法规会被跳过）")
    ap.add_argument("--shrink", action="store_true", help="长批注自动缩字号压进一屏（默认关闭：统一 9pt，超长的自己滚动）")
    args = ap.parse_args()

    import openpyxl
    from openpyxl.comments import Comment

    import basis_comment

    if args.shrink:
        basis_comment.SHRINK = True    # 只有显式要求才缩字号，默认全部 9pt
    import lawtext

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, LEDGER)
    reg = os.path.join(root, REG)
    work = os.path.join(os.path.dirname(root), "_案例库维护留痕")
    cache = os.path.join(work, "_law_index", "index.json")

    t0 = time.time()
    idx, hit = lawtext.load_or_build(reg, cache, ocr=not args.no_ocr, verbose=True)
    print("\n法规库：%d 部｜%s（%.1fs）"
          % (len(idx.laws), "缓存命中" if hit else "已重建", time.time() - t0))

    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]
    before = snapshot(ws)

    plans, unresolved, n_cell, n_art, n_miss, maxlen = [], [], 0, 0, 0, 0
    for r in range(2, ws.max_row + 1):
        h = str(ws.cell(r, 8).value or "")
        if not h.strip():
            continue
        text, stat = basis_comment.build(idx, h)
        if not text:
            continue
        n_cell += 1
        n_art += stat["arts"]
        n_miss += stat["missing"]
        maxlen = max(maxlen, len(text))
        plans.append((r, text))
        for it in idx.parse_h(h):
            if not it["law"]:
                unresolved.append([r, it["raw"], "法规未收录",
                                   "、".join(a["no"] for a in it["arts"])])
            else:
                for a in it["arts"]:
                    if a["no"] not in idx.laws[it["law"]]["arts"]:
                        unresolved.append([r, it["law"], "条号未找到", a["no"]])

    print("待写批注：%d 行｜条文引用 %d 处｜未解析 %d 处｜最长批注 %d 字"
          % (len(plans), n_art, len(unresolved), maxlen))
    if unresolved:
        print("\n未解析明细（前 15 条）：")
        for x in unresolved[:15]:
            print("  行%-4s %-40s %s %s" % (x[0], x[1][:40], x[2], x[3]))

    if args.csv and unresolved:
        os.makedirs(os.path.dirname(args.csv) or ".", exist_ok=True)
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["台账行", "法规名", "问题", "条号"])
            w.writerows(unresolved)
        print("明细已导出：%s" % args.csv)

    # 样例：打印前 2 行批注，干跑时供人工确认形态
    for r, text in plans[:2]:
        print("\n———— 行 %d 批注样例 ————\n%s" % (r, text[:900]))

    # 顶到高度上限的行：框装不下，只能「编辑批注」滚动看。必须显式列出，
    # 否则用户以为「都看全了」，实际这几行被裁（2026-11-03）。
    # 悬停框不能滚动 → 塞不下就自动降字号；降到下限仍塞不下的，如实列出来让用户滚动
    over, fonts = [], {}
    for r, t in plans:
        _w, _h, _f = basis_comment.fit(t)
        if _h >= basis_comment.BOX_MAX_H:
            over.append(r)
        if _f < basis_comment.FONT_MAX:
            fonts[r] = _f
    print("\n批注框：%d×≤%d px｜9pt 原样 %d 行｜缩字号 %d 行｜缩到下限仍超框 %d 行%s"
          % (basis_comment.BOX_W, basis_comment.BOX_MAX_H,
             len(plans) - len(fonts), len(fonts), len(over),
             ("：" + "、".join(str(x) for x in over)) if over else ""))

    if not args.apply:
        print("\n（干跑结束，未改动台账；加 --apply 落盘）")
        return 0

    bak_dir = os.path.join(work, "台账备份_%s" % time.strftime("%Y%m%d"))
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, "%s_bak_%s.xlsx" % (
        os.path.splitext(LEDGER)[0], time.strftime("%H%M%S")))
    shutil.copy2(ledger, bak)
    print("\n已备份：%s" % bak)

    for r, text in plans:
        cell = ws.cell(r, 8)
        cm = cell.comment or Comment(text, "案例库")
        cm.text = text
        cm.author = "案例库"
        # 按内容估算尺寸：Excel 的悬停框不会滚动，框小了就把法条裁掉
        _w, _h, _f = basis_comment.fit(text)
        cm.width, cm.height = _w, _h
        cell.comment = cm

    # 表头也挂一条说明：否则拿到台账的人不知道 H 列有悬停内容
    #
    # 2026-11-04 加「装不下的 5 格怎么办」：悬停批注框是只读浮层，**没有滚动条**
    # （VML 的 ClientData 里没有任何滚动字段，Comment.Shape 也没有 ScrollBars）。
    # 按 2560×1440 实测：框能开到 1300×1150，504/509 一屏可读，剩 5 格物理装不下
    # （加宽到 1980 也只救 1 格，因为法条正文是逐条硬换行）。这 5 格必须给出口，
    # 否则用户只知道「看不全」却不知道去哪看。
    hint = HEADER_HINT
    hc = ws.cell(1, 8).comment or Comment(hint, "案例库")
    hc.text = hint
    hc.author = "案例库"
    hc.width = 700
    hc.height = 190
    ws.cell(1, 8).comment = hc
    wb.save(ledger)
    print("已写入批注：%d 个单元格（另含表头说明 1 条）" % len(plans))

    # 字号必须写进 comments xml：Comment 没有字号属性，长批注只能靠缩字号塞进一屏
    n_font, dist = set_comment_fonts(ledger, fonts)
    if n_font:
        detail = "，".join("%spt×%d" % (_fmt_sz(k), v) for k, v in sorted(dist.items()))
        print("批注字号：%d 行已缩小（%s）" % (n_font, detail))

    # 尺寸必须单独改 VML：openpyxl 写成 px 且不给锚点，Excel 会按默认锚点重算，等于没设
    n_shape, n_conv, n_anchor = fix_vml_size(ledger, ws)
    print("批注框：%d 个 shape｜px→pt %d 处｜补锚点 %d 个" % (n_shape, n_conv, n_anchor))

    # ---- 落盘后自证：值一字未改 + 批注逐个可读 ----
    wb2 = openpyxl.load_workbook(ledger)
    ws2 = wb2.worksheets[0]
    after = snapshot(ws2)
    diff = [k for k in before if before[k] != after.get(k)]
    got = sum(1 for r in range(2, ws2.max_row + 1) if ws2.cell(r, 8).comment)
    print("\n自证：单元格值差异 %d 处（须 0）｜H 列批注 %d 个／应有 %d 个"
          % (len(diff), got, len(plans)))
    if diff:
        print("  ✗ 有 %d 格被改动，前 3 处：%s" % (len(diff), diff[:3]))
        return 1
    if got != len(plans):
        print("  ✗ 批注数量不符")
        return 1
    print("  ✓ 只加了批注，台账内容零改动")
    print("文件 sha256 变化：%s → %s" % (sha(bak)[:12], sha(ledger)[:12]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
