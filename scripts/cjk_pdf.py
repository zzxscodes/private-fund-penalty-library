# -*- coding: utf-8 -*-
"""纯文本 → 可检索中文 PDF（**零额外依赖**：只用 pymupdf + 系统中文字体）。

为什么要有这个模块（2026-10-27 实测，可迁移性硬伤）：
  旧实现 `fetch_csrc.make_pdf_from_text` 依赖 **reportlab**（STSong-Light CID 字体）。
  但 reportlab 不在托管 Python 的默认依赖里 —— 换一台机器 `ModuleNotFoundError`，
  而 except 分支把它吞掉 → **证监局决定书 PDF 一份都没生成**，台账行全是空文件，
  且表面上看不出失败（抓取统计照样"成功 45 条"）。

  ⇒ 改成本模块：pymupdf 已是本技能硬依赖（正文抽取要用），只要系统有中文字体即可工作。
    reportlab 只在它可用时优先用（排版更稳），不可用就走 pymupdf，两条路都保证
    **PDF 有完整可复制文字层**（verify B7 要能取回正文）。

字体查找顺序：
  ① 环境变量 CJK_FONT（显式指定，最稳）
  ② Windows：C:\\Windows\\Fonts 下的 msyh/simhei/simsun/simkai
  ③ macOS：/System/Library/Fonts、/Library/Fonts 下的 PingFang/STHeiti/Hiragino
  ④ Linux：/usr/share/fonts 下的 Noto Sans CJK / WQY
  都找不到 → 抛异常并给出可执行的修复建议，绝不静默产出乱码 PDF。
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

FONT_CANDIDATES = [
    # Windows
    r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\simkai.ttf",
    # macOS
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    # Linux
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
]
FONT_DIRS = [r"C:\Windows\Fonts", "/System/Library/Fonts", "/Library/Fonts",
             "/usr/share/fonts"]
_FONT_CACHE = [None]


def find_cjk_font():
    """返回可用的中文字体文件路径；找不到抛 RuntimeError（附修复建议）。"""
    if _FONT_CACHE[0]:
        return _FONT_CACHE[0]
    env = os.environ.get("CJK_FONT")
    cands = [env] if env else []
    cands += FONT_CANDIDATES
    for p in cands:
        if p and os.path.exists(p):
            _FONT_CACHE[0] = p
            return p
    # 目录兜底扫描
    for d in FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            low = fn.lower()
            if low.endswith((".ttf", ".ttc", ".otf")) and any(
                    k in low for k in ("msyh", "simhei", "simsun", "simkai", "pingfang",
                                       "stheiti", "hiragino", "noto", "wqy", "cjk")):
                _FONT_CACHE[0] = os.path.join(d, fn)
                return _FONT_CACHE[0]
    raise RuntimeError(
        "未找到中文字体。请在机器上装上任意 CJK 字体，或用环境变量显式指定：\n"
        "  Windows: set CJK_FONT=C:\\Windows\\Fonts\\msyh.ttc\n"
        "  Linux  : sudo apt install fonts-noto-cjk  然后 CJK_FONT=/usr/share/fonts/..."
        "/NotoSansCJK-Regular.ttc\n"
        "  macOS  : CJK_FONT=/System/Library/Fonts/PingFang.ttc")


def _fitz():
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    return fitz


def _wrap(text, cpl):
    """按显示宽度折行：CJK 计 2 宽、ASCII 计 1 宽。"""
    out = []
    for para in re.split(r"\n+", text or ""):
        para = para.rstrip()
        if not para:
            out.append("")
            continue
        line, w = "", 0
        for ch in para:
            cw = 2 if ord(ch) > 0x2E80 else 1
            if w + cw > cpl:
                out.append(line)
                line, w = "", 0
            line += ch
            w += cw
        if line:
            out.append(line)
    return out


def text_to_pdf(text, out_path, title="", meta_lines=None, font_path=None):
    """把纯文本渲染成 A4 PDF（正文可复制，中文不乱码）。返回输出路径。"""
    font_path = font_path or find_cjk_font()
    fitz = _fitz()
    PW, PH = 595, 842
    ML, MR, MT, MB = 56, 56, 64, 56
    doc = fitz.open()
    page = doc.new_page(width=PW, height=PH)
    y = MT

    def put(s, size, gap, bold=False):
        nonlocal y, page
        lines = _wrap(s, int((PW - ML - MR) / (size * 0.55)))
        for ln in lines:
            if y + gap > PH - MB:
                page = doc.new_page(width=PW, height=PH)
                y = MT
            # pymupdf 的 insert_text 对 CJK 需要 fontfile + fontname
            page.insert_text((ML, y), ln, fontname="cjk", fontfile=font_path,
                             fontsize=size)
            y += gap

    if title:
        put(title, 15, 22)
        y += 6
    for m in (meta_lines or []):
        put(m, 10, 16)
    if meta_lines:
        y += 6
    put(text or "", 11, 17)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    # 必须子集化（2026-10-27 实测）：不调用则整份 TTC/TTF 被**逐页重复嵌入**，
    # 单份决定书 PDF 达 **19.6 MB**（一年库几百份 ≈ 10GB）。子集化后 ~100KB。
    try:
        doc.subset_fonts()
    except Exception as e:
        print("[WARN] 字体子集化失败，PDF 体积可能异常：%s" % e, file=sys.stderr)
    doc.save(out_path, garbage=4, deflate=True)
    doc.close()
    return out_path


def selfcheck(out_path=None):
    """冒烟自检：生成一页中文 PDF 并**取回文字层**验证可读。"""
    import tempfile

    fitz = _fitz()
    out_path = out_path or os.path.join(tempfile.gettempdir(), "cjk_pdf_selftest.pdf")
    text_to_pdf("私募投资基金监督管理暂行办法第四条第一款。测试中文可检索性。",
                out_path, title="自检文书")
    d = fitz.open(out_path)
    got = "".join(p.get_text() for p in d)
    d.close()
    hit = ("私募" in got) and ("第四条" in got)
    size_kb = os.path.getsize(out_path) / 1024
    # 体积断言：忘做 subset_fonts 时这里会立刻炸（19MB vs 100KB），防止回归
    if size_kb > 2048:
        hit = False
        print("  [FAIL] PDF 体积 %.0f KB，疑似未做字体子集化" % size_kb)
    return hit, out_path, got[:80]


if __name__ == "__main__":
    ok, p, got = selfcheck()
    print("字体:", find_cjk_font())
    print("生成:", p)
    print("取回:", got)
    print("自检:", "PASS 中文文字层可检索" if ok else "FAIL 文字层不可读（字体问题）")
    sys.exit(0 if ok else 1)
