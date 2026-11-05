# -*- coding: utf-8 -*-
"""用系统 CJK 字体重建证监局决定书 PDF。

背景（2026-10-27 实测）：此前用 chrome --headless --print-to-pdf 在无字体环境下生成的
52 份证监局 PDF 只嵌入 Helvetica/WinAnsi，中文全部渲染为方块"·"，等于不可读。
本脚本改用 pymupdf + 系统字体（微软雅黑）直接排版，并保留完整可复制文字层。
"""
import os, re, sys, json, textwrap
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import pymupdf as fitz
from html import unescape as _unescape

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = r"C:/Windows/Fonts/msyh.ttc"
if not os.path.exists(FONT):
    FONT = r"C:/Windows/Fonts/simhei.ttf"
assert os.path.exists(FONT), "无可用中文字体"

def tidy(s):
    s = _unescape(str(s or ""))
    s = re.sub(r"\s+", " ", s)
    return re.sub(r"([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])", r"\1\2", s).strip()


PW, PH = 595, 842           # A4
ML, MR, MT, MB = 56, 56, 64, 56
FONTNAME = "cjk"


def esc(t):
    return t


def make_pdf(out, title, body, meta_lines):
    doc = fitz.open()
    page = doc.new_page(width=PW, height=PH)
    y = MT
    # 标题（加粗：用黑体近似）
    page.insert_textbox(fitz.Rect(ML, y, PW - MR, y + 90), title,
                        fontname=FONTNAME, fontfile=FONT, fontsize=15,
                        align=fitz.TEXT_ALIGN_CENTER)
    y += 34 + 16 * max(0, len(title) // 26)
    for ln in meta_lines:
        page.insert_textbox(fitz.Rect(ML, y, PW - MR, y + 22), ln,
                            fontname=FONTNAME, fontfile=FONT, fontsize=9.5,
                            align=fitz.TEXT_ALIGN_CENTER)
        y += 16
    y += 10
    page.insert_textbox(fitz.Rect(ML, y, PW - MR, y + 4), "", fontname=FONTNAME,
                        fontfile=FONT, fontsize=4)
    y += 6
    # 正文：按宽度折行（中文按字符数）
    fs, lh = 11, 18
    # 每行字数必须留足余量：按 fs*0.98 估宽会让 44 字×11pt=484pt 超出版心 483pt，
    # insert_textbox 装不下就直接【整行不渲染】→ 生成只有标题的空壳 PDF（2026-10-27 实测 3 份中招）。
    maxchars = int((PW - ML - MR) / fs) - 2
    buf = []
    for para in re.split(r"\n+", body):
        para = para.strip()
        if not para:
            buf.append("")
            continue
        buf += textwrap.wrap(para, width=maxchars, drop_whitespace=False) or [para]
    for line in buf:
        if y + lh > PH - MB:
            page = doc.new_page(width=PW, height=PH)
            y = MT
        rc = page.insert_textbox(fitz.Rect(ML, y, PW - MR, y + lh + 2), line,
                                 fontname=FONTNAME, fontfile=FONT, fontsize=fs)
        if rc < 0:      # 兜底：装不下就缩小字号重写一次，绝不允许整行丢失
            page.insert_textbox(fitz.Rect(ML, y, PW - MR, y + lh + 2), line,
                                fontname=FONTNAME, fontfile=FONT, fontsize=fs - 2)
        y += lh
    try:
        doc.subset_fonts()      # 字体子集化，否则整份 TTC 嵌入 → 单份 12MB
    except Exception as e:
        print("   [warn] subset failed:", e)
    doc.save(out, garbage=3, deflate=True)
    doc.close()
    return os.path.getsize(out)


def main():
    lib_root = sys.argv[1]
    csrc = json.load(open(os.path.join(HERE, "csrc.json"), encoding="utf-8"))
    by_url = {it["url"]: it for it in csrc["items"]}
    recs = json.load(open(os.path.join(HERE, "records_30d.json"), encoding="utf-8"))
    n = 0
    skipped = 0
    for r in recs:
        if r.get("source") != "证监局":
            continue
        files = r.get("files") or []
        urls = r.get("urls") or []
        for i, f in enumerate(files):
            u = urls[i] if i < len(urls) else (urls[-1] if urls else "")
            it = by_url.get(u)
            if not it:
                print("  [skip] 无源条目:", r.get("subject"), "|", f["name"])
                skipped += 1
                continue
            body = tidy(it.get("text") or "")
            title = it.get("title", "").strip()
            meta = []
            if it.get("wh"):
                meta.append(it["wh"])
            meta.append("处罚（决定）时间：%s    挂网发布日期：%s" %
                        (r.get("date_decided", ""), r.get("publish_date", "")))
            meta.append("原文链接：%s" % it.get("url", ""))
            dst = os.path.join(lib_root, r["source_dir"], r["category"], f["name"])
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            sz = make_pdf(dst, title, body, meta)
            n += 1
            print("  [%d] %s (%d bytes)" % (n, f["name"], sz))
    print("\n重建完成：%d 份证监局决定书 PDF（CJK 字体可正常显示）；跳过 %d" % (n, skipped))


if __name__ == "__main__":
    main()
