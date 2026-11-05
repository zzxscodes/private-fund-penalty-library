# -*- coding: utf-8 -*-
"""法规正文库（lawtext）

用途：把 `06_以上各类所有案例的配套法规依据` 里的 docx/pdf 拆成「第X条 → 条文正文」，
供两处使用：
  1) `annotate_basis.py` 给台账 H 列写批注（鼠标悬停即见法条原文）；
  2) `audit_basis_comments.py`（Z15）校验「H 列引用的每一条都能在 06 里找到正文」。

设计原则（对齐 SKILL 总原则①）：
  - **唯一事实源是 06 目录本身**，本模块只做解析与缓存，不另存一份法条文本；
  - 缓存落在库**外**（`<库父目录>/_案例库维护留痕/_law_index/`），库根保持干净；
  - 不硬编码任何本机路径，库根由调用方传入。

条文切分口径：
  - 以「第X条」「第X条之N」作为条起点，到下一条起点前为止（含其下的（一）（二）各项）；
  - docx 按段落保留换行（每段通常就是一「款」），pdf 先做行级去噪再拼接，
    避免 PDF 视觉折行把一句话切断；
  - 章/节标题（第一章、第五节）不单独成条，随上一条尾部或丢弃。
"""

from __future__ import annotations

import os
import re

# ---------------------------------------------------------------- 名称归一

_BOOK = re.compile(r"[《》\s\u3000]")
_YEARFIX = re.compile(r"[（(]\s*20\d\d\s*年\s*(?:修订|修正|年修订)\s*[）)]")


def norm_name(s: str) -> str:
    """法规名归一：去书名号/空白，去掉「（2025年修订）」这类版本后缀。

    注意：「（试行）」是名称的一部分（《基金募集机构投资者适当性管理实施指引（试行）》），
    不能一起剥掉，否则会和正式版混淆。
    """
    s = _BOOK.sub("", str(s or ""))
    s = _YEARFIX.sub("", s)
    return s.strip()


# 简称/别名 → 06 目录里的规范名（规范名 = 文件名去掉扩展名与版本后缀）
ALIAS = {
    # 私募基金监管的核心两部
    "私募办法": "私募投资基金监督管理暂行办法",
    "暂行办法": "私募投资基金监督管理暂行办法",
    "私募暂行办法": "私募投资基金监督管理暂行办法",
    "私募基金监管办法": "私募投资基金监督管理暂行办法",
    "私募管理办法": "私募投资基金监督管理暂行办法",
    "私募监管办法": "私募投资基金监督管理暂行办法",
    "私募投资基金监督管理暂行办法": "私募投资基金监督管理暂行办法",
    "私募条例": "私募投资基金监督管理条例",
    "管理条例": "私募投资基金监督管理条例",
    "私募投资基金监督管理条例": "私募投资基金监督管理条例",
    # 其他常见简称
    "若干规定": "关于加强私募投资基金监管的若干规定",
    "关于加强私募投资基金监督管理的若干规定": "关于加强私募投资基金监管的若干规定",
    "适当性管理办法": "证券期货投资者适当性管理办法",
    "适当性办法": "证券期货投资者适当性管理办法",
    "登记备案办法": "私募投资基金登记备案办法",
    "私募基金登备办法": "私募投资基金登记备案办法",
    "募集办法": "私募投资基金募集行为管理办法",
    "信披办法": "私募投资基金信息披露管理办法",
    "内控指引": "私募投资基金管理人内部控制指引",
    "基金法": "中华人民共和国证券投资基金法",
    "证券法": "中华人民共和国证券法",
    "纪律处分办法": "中国证券投资基金业协会自律管理和纪律处分措施实施办法",
    "实施细则": "上海证券交易所首次公开发行证券发行与承销业务实施细则",
    # 2026-11-04 补：网页台账做「按法规筛选」时暴露出的同义写法（raw 文本里的原样，
    # 不是我们造的简称）。纯新增 key，不改变任何既有映射的行为。
    "投资者适当性办法": "证券期货投资者适当性管理办法",
    "关于加强私募基金监管的若干规定": "关于加强私募投资基金监管的若干规定",
    # OCR 把破折号「——」认成「一一」，与第 1 号指引是同一部
    "私募基金管理人登记指引第1号一一基本经营要求": "私募基金管理人登记指引第1号",
}

# 这些写法本身不唯一，必须靠同一单元格里先出现的全称来定（「《条例》第三十一条」跟在
# 「《私募投资基金监督管理条例》第三十二条」后面，指的就是同一部）
AMBIGUOUS = {"条例", "管理办法", "实施办法", "暂行规定", "规定", "办法", "指引", "通知", "公告"}

# 06 目录里确实没有的（如实标注，不猜、不编条文）
KNOWN_MISSING = {
    "私募基金管理人登记指引第1号",
    "私募投资基金备案指引第2号",
    "私募投资基金管理人登记和基金备案办法（试行）",
    "上市公司规范运作指引",
    "深圳证券交易所公司债券上市规则",
    "关于规范面向公众开展的证券投资咨询业务行为若干问题的通知",
    "合伙协议",
}

_CJK_NUM = "一二三四五六七八九十百零〇两"
_ART_RE = re.compile(r"^第([%s]+)条(?:之([%s]+))?" % (_CJK_NUM, _CJK_NUM))
_ANY_ART_RE = re.compile(r"第([0-9]+(?:\.[0-9]+){1,2}|[%s]+)条(?:之([%s]+))?" % (_CJK_NUM, _CJK_NUM))
# 交易所规则用「1.4」「3.1.5」这种数字条号，正文里不带「第…条」二字
_DEC_ART_RE = re.compile(r"^(\d+(?:\.\d+){1,2})[\s\u3000]+")
_PAGENO = re.compile(r"^[\s\-—－_]*\d{1,4}[\s\-—－_]*$")
_HEADER = re.compile(r"^(第[%s]+[章节]|[（(][%s]+[）)]\s*$)" % (_CJK_NUM, _CJK_NUM))


def num2cn(n: int) -> str:
    """3 → 三，21 → 二十一，105 → 一百零五。仅用于条号。"""
    d = "零一二三四五六七八九"
    if n <= 10:
        return "十" if n == 10 else d[n]
    if n < 20:
        return "十" + d[n % 10]
    if n < 100:
        s = d[n // 10] + "十"
        return s + d[n % 10] if n % 10 else s
    if n < 1000:
        s = d[n // 100] + "百"
        r = n % 100
        if r == 0:
            return s
        if r < 10:
            return s + "零" + d[r]
        return s + (num2cn(r) if r >= 20 else "一" + num2cn(r) if r == 10 else "十" + d[r % 10])
    return str(n)


def docx_autonum_articles(path: str) -> list:
    """复原 Word 自动编号的「第N条」条号（转换后文字层里没有条号的情况）。

    判据是确定性的：`numbering.xml` 里该编号级别的 `lvlText` 就是「第%1条」，
    文档流里每个带该 numPr 的段落依次递增。**不是按段落顺序猜**——若 lvlText 不是
    「第%1条」或存在 restart，直接返回空，交给上层标注「条号缺失」，绝不硬凑。
    """
    import zipfile
    from xml.etree import ElementTree as ET

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        z = zipfile.ZipFile(path)
        numx = ET.fromstring(z.read("word/numbering.xml"))
        docx = ET.fromstring(z.read("word/document.xml"))
    except Exception:      # noqa: BLE001
        return []
    # numId -> 该级别 lvlText 是否「第%1条」；以及起始值
    art_num, start_over = {}, {}
    for num in numx.findall(W + "num"):
        nid = num.get(W + "numId")
        for ov in num.findall(W + "lvlOverride"):
            so = ov.find(W + "startOverride")
            if so is not None:
                start_over[nid] = int(so.get(W + "val"))
    for an in numx.findall(W + "abstractNum"):
        aid = an.get(W + "abstractNumId")
        for lvl in an.findall(W + "lvl"):
            if lvl.get(W + "ilvl") != "0":
                continue
            lt = lvl.find(W + "lvlText")
            st = lvl.find(W + "start")
            if lt is not None and "第%1条" in (lt.get(W + "val") or ""):
                start_over.setdefault("abs:" + aid, int(st.get(W + "val")) if st is not None else 1)
    for num in numx.findall(W + "num"):
        aid = num.find(W + "abstractNumId")
        if aid is not None and "abs:" + aid.get(W + "val") in start_over:
            art_num[num.get(W + "numId")] = start_over["abs:" + aid.get(W + "val")]
    if not art_num:
        return []
    body = docx.find(W + "body")
    if body is None:
        return []
    counter = dict(art_num)
    out = []
    for p in body.iter(W + "p"):
        ppr = p.find(W + "pPr")
        nid = None
        if ppr is not None:
            npr = ppr.find(W + "numPr")
            if npr is not None:
                nv = npr.find(W + "numId")
                nid = nv.get(W + "val") if nv is not None else None
        text = "".join(t.text or "" for t in p.iter(W + "t")).strip()
        if nid in counter:
            out.append((art_no_from_int(counter[nid]), text))
            counter[nid] += 1
    return out


def art_no_from_int(n: int) -> str:
    return "第" + num2cn(n) + "条"


def art_no(m) -> str:
    """把正则匹配转成规范条号「第二十三条」「第八条之一」。"""
    s = "第" + m.group(1) + "条"
    if m.group(2):
        s += "之" + m.group(2)
    return s


# ---------------------------------------------------------------- 抽取正文

def _art_start_count(t: str) -> int:
    n = 0
    for ln in str(t or "").splitlines():
        s = ln.strip()
        if _ART_RE.match(s) or _DEC_ART_RE.match(s):
            n += 1
    return n


def _ocr_image(img) -> str:
    import ocrlib
    r, _ = ocrlib._engine()(img)
    return " ".join(t[1] for t in r) if r else ""


def _ocr_page_columns(path: str, i: int, gutter: float = 0.5) -> str:
    """把一页按左右两栏切开分别 OCR。

    为什么需要：公报类 PDF（如《证券投资基金法》的全国人大公报版）是**双栏排版**，
    整页 OCR 会左右串行，条号顺序全乱——「第二条」后面直接接上右栏的目录。
    按栏切图后各栏内部顺序才正确。
    """
    import fitz
    import numpy as np
    d = fitz.open(path)
    try:
        pg = d[i]
        w, h = pg.rect.width, pg.rect.height
        out = []
        for k in range(2):
            x0 = 0 if k == 0 else w * gutter
            x1 = w * gutter if k == 0 else w
            clip = fitz.Rect(x0, 0, x1, h)
            zoom = 200 / 72.0
            if clip.width * zoom > 2000:
                zoom *= 2000 / (clip.width * zoom)
            pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if img.ndim == 3 and img.shape[2] >= 3:
                R = np.ascontiguousarray(img[:, :, 0])
                img = np.stack([R, R, R], axis=2)
            out.append(_ocr_image(img))
        return "\n".join(out)
    finally:
        d.close()


def _two_column_page(pg):
    """便宜的双栏判定 → 返回栏缝位置（占页宽比例），单栏返回 None。

    为什么不用「先整页 OCR 再看条号多少」：整页 OCR 一页要 48 秒（实测），
    分栏两张只要 4 秒。用像素投影先判栏，把扫描件法规的解析从十几分钟压到一两分钟。

    判据：在页面中线 ±8% 范围内逐列取「墨量最少的窄条」，若该窄条几乎无墨、
    且左右两侧都有正文，即为双栏。**必须逐列找最窄**——直接取 [0.46,0.54] 平均值，
    会把紧挨栏缝的那一列字算进来，双栏被判成单栏（实测踩过）。
    只看正文区：页眉「全国人民代表大会常务委员会公报…」横跨两栏，会把栏缝糊住。
    """
    import fitz
    import numpy as np
    pix = pg.get_pixmap(matrix=fitz.Matrix(1, 1))
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    g = img[:, :, 0] if img.ndim == 3 else img
    h, w = g.shape
    if w < 200:
        return None
    body = g[int(h * 0.15):int(h * 0.85)]
    ink = (body < 160).mean(axis=0)
    best_frac, best_val = None, 1.0
    for x in range(42, 59):
        c = int(w * x / 100.0)
        val = float(ink[max(0, c - 2):c + 3].mean())
        if val < best_val:
            best_frac, best_val = x / 100.0, val
    left = float(ink[int(w * 0.10):int(w * 0.40)].mean())
    right = float(ink[int(w * 0.60):int(w * 0.90)].mean())
    if best_val < 0.01 and left > 0.02 and right > 0.02:
        return best_frac
    return None


def _ocr_page(path: str, i: int) -> str:
    import fitz
    import numpy as np
    d = fitz.open(path)
    try:
        pg = d[i]
        gutter = _two_column_page(pg)
        if gutter:
            t = _ocr_page_columns(path, i, gutter)
        else:
            zoom = 200 / 72.0
            if pg.rect.width * zoom > 2000:
                zoom *= 2000 / (pg.rect.width * zoom)
            pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if img.ndim == 3 and img.shape[2] >= 3:
                R = np.ascontiguousarray(img[:, :, 0])
                img = np.stack([R, R, R], axis=2)
            t = _ocr_image(img)
            if _art_start_count(t) < 3:      # 整页读不到条首，再按栏试一次
                t2 = _ocr_page_columns(path, i)
                if _art_start_count(t2) > _art_start_count(t):
                    t = t2
    finally:
        d.close()
    return t


def read_text(path: str, ocr: bool = True, meta: dict = None) -> str:
    """取纯文本。

    docx 走 python-docx；PDF **逐页**判断——文字层不全的（前几页有、后面是扫描图）
    必须把缺的页补 OCR，否则会拿到半部法规还以为完整。
    """
    low = path.lower()
    if low.endswith(".docx"):
        try:
            from docx import Document
        except ImportError:      # pragma: no cover
            return ""
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    if low.endswith(".pdf"):
        import fitz
        d = fitz.open(path)
        try:
            pages = [pg.get_text() for pg in d]
            n = len(pages)
        finally:
            d.close()
        out, n_ocr = [], 0
        for i, t in enumerate(pages):
            if len(t.strip()) >= 40:
                out.append(t)
            elif ocr:
                out.append(_ocr_page(path, i))
                n_ocr += 1
            else:
                out.append("")
        if meta is not None:
            meta["pages"] = len(pages)
            meta["ocr_pages"] = n_ocr
        return "\n".join(out)
    return ""


def cn2num(s: str) -> int:
    d = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
         "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    unit = {"十": 10, "百": 100, "千": 1000}
    total, cur = 0, 0
    for ch in s:
        if ch in d:
            cur = d[ch]
        elif ch in unit:
            total += (cur or 1) * unit[ch]
            cur = 0
    return total + cur


def art_sort_key(no: str):
    """条号 → 可比较的键；认不出返回 None。"""
    m = re.match(r"^第([0-9]+(?:\.[0-9]+){0,2})条$", no or "")
    if m:
        return tuple(int(x) for x in m.group(1).split("."))
    m = re.match(r"^第([%s]+)条(?:之([%s]+))?$" % (_CJK_NUM, _CJK_NUM), no or "")
    if m:
        return (cn2num(m.group(1)), cn2num(m.group(2)) if m.group(2) else 0)
    return None


def clean_lines(text: str) -> list:
    out = []
    for ln in str(text or "").replace("\u3000", " ").splitlines():
        s = ln.strip()
        if not s:
            continue
        # 扫描件页眉页脚：公报版每页都有「全国人民代表大会常务委员会公报 2015·3」
        s = re.sub(r"全国人民代表大会常务委员会公报\s*\d{4}\s*[·.]?\s*\d*", "", s).strip()
        s = re.sub(r"[-~—－]{1,3}\s*\d{1,4}\s*[-—－~]{0,3}$", "", s).strip()
        if not s:
            continue
        if _PAGENO.match(s):          # 页码
            continue
        if len(s) <= 30 and _HEADER.match(s):   # 章/节标题
            continue
        # 只有几个字、没有任何标点/数字的孤立行，多是扫描件的书眉（如「行业自律」），
        # 留着会粘在条文末尾。带标点的短句（「予以公告。」）不受影响。
        if len(s) <= 8 and not re.search(r"[。，、；：？！（）()0-9]", s) and not s.startswith("第"):
            continue
        # 章标题在转换后常常丢了「第十章」，只剩「关于创业投资基金的特别规定」这种孤行
        if (len(s) <= 24 and not re.search(r"[。，、；：？！（）()0-9]", s)
                and (s.startswith("关于") or s.endswith(("总则", "附则", "规定", "管理", "规则", "办法")))):
            continue
        out.append(s)
    return out


def split_articles(text: str) -> list:
    """→ [(条号, [段...]), ...]，保持原文顺序。

    支持两种条号：中文「第二十三条」、「第二十三条之一」，以及交易所规则的
    「1.4」「3.1.5」（正文里不带「第…条」二字）。目录页（带 `....` 引导符）一律跳过。
    """
    # OCR 把整页文字拼成一行，条号夹在句中 —— 先在「。；：）」之后、条号之前补换行，
    # 否则一页只会被认成一条（首条之后全被吞进它的正文）。
    text = re.sub(r"(?<=[。；：！？）)])\s*(?=第[%s]+条)" % _CJK_NUM, "\n", str(text or ""))
    paras = clean_lines(text)
    arts, cur_no, cur = [], None, []

    def _new(no, rest):
        nonlocal cur_no, cur
        if cur_no:
            arts.append((cur_no, cur))
        cur_no, cur = no, ([rest] if rest else [])

    for s in paras:
        if re.search(r"\.{4,}", s):      # 目录行
            continue
        m = _ART_RE.match(s)
        if m:
            _new(art_no(m), s[m.end():].lstrip(" 　、.．:："))
            continue
        if _DEC_ART_RE.match(s):
            dm = _DEC_ART_RE.match(s)
            _new("第%s条" % dm.group(1), s[dm.end():].strip())
            continue
        if cur_no:
            cur.append(s)
    if cur_no:
        arts.append((cur_no, cur))
    return arts


def _despace(s: str) -> str:
    """去掉汉字之间被排版插入的单个空格（`为 了 规 范` → `为了规范`）。

    只处理「汉字 空格+ 汉字」与「汉字 空格+ （」两类，不动英文/数字之间的空格，
    否则会把「第 1.4 条」这类正常写法也揉在一起。
    """
    s = re.sub(r"(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5（(《“])", "", s)
    s = re.sub(r"(?<=[）)》”])\s+(?=[\u4e00-\u9fa5])", "", s)
    return s


def _reflow_fragments(ps: list) -> list:
    """把「一个字一段」的坏稿还原成正常段落。

    症状：某份 docx 由网页/PDF 转换而来，段内每个字后面都跟一个 `<w:br/>`，
    解析成段落就变成 371 行、每行一个字（`证/券/公/司/承/销/…`）。批注里显示为
    竖排单字，且行数暴涨把框顶到尺寸上限、内容被截断（2026-11-03 实测）。
    反例：`上海证券交易所首次公开发行证券发行与承销业务实施细则.docx`，
    13819 字里有 6950 个单字行。

    判据保守：只把**长度 ≤1**的段落并进前一段。真正的款/项都远长于 1 个字，
    所以不会把「（一）…」「（二）…」粘成一行；条号标题本身也是完整行，不受影响。
    """
    out = []
    for p in ps:
        if len(p) <= 1 and out:
            out[-1] += p
        else:
            out.append(p)
    return out


def tidy_article(paras: list, is_pdf: bool) -> str:
    """把一段条文整理成便于阅读的多行文本。

    PDF 的视觉折行会把一句话切断，所以先拼成一整段，再在「（一）」这类项号前断行；
    docx 段落本来就是完整的款，直接保留换行。
    """
    ps = [_despace(p.strip()) for p in paras if p.strip()]
    if not ps:
        return ""
    if not is_pdf:
        body = "\n".join(_reflow_fragments(ps))
    else:
        body = "".join(ps)
        body = re.sub(r"(?=[（(][%s]{1,3}[）)])" % _CJK_NUM, "\n", body)
    return body.strip()


def _is_ordered(arts: dict) -> bool:
    """条号是否基本递增——用来识别「文本抽取把条文顺序搞乱了」的坏样本。

    乱序的典型成因：双栏排版串行、扫描件整页 OCR 阅读顺序错。这类样本宁可在批注里
    标「抽取不可靠，请核对原文」，也不能把错的条文端给用户。
    """
    keys = [art_sort_key(k) for k in arts]
    if sum(1 for k in keys if k is None) > max(2, len(keys) * 0.2):
        return False
    seq = [k for k in keys if k is not None]
    if len(seq) < 3:
        return False
    inv = sum(1 for i in range(1, len(seq)) if seq[i] < seq[i - 1])
    return inv / max(1, len(seq) - 1) <= 0.05


class LawIndex:
    """06 目录 → {法规名: {条号: 正文}}。"""

    def __init__(self):
        self.laws = {}        # 规范名 -> {"file": 相对路径, "kind": docx/pdf, "arts": {条号: 文本}}
        self.titles = []      # 规范名列表（用于模糊匹配）

    # ---------- 构建 ----------
    def build(self, reg_root: str, ocr: bool = True, verbose: bool = False, records: dict = None):
        for dp, dn, fn in os.walk(reg_root):
            for f in sorted(fn):
                if not f.lower().endswith((".pdf", ".docx")):
                    continue
                name = norm_name(os.path.splitext(f)[0])
                p = os.path.join(dp, f)
                meta = {}
                rel = os.path.relpath(p, reg_root).replace(os.sep, "/")
                rec = (records or {}).get(rel)
                if rec is not None:
                    txt = rec.get("text", "")
                    meta = {"pages": rec.get("pages", 0), "ocr_pages": rec.get("ocr_pages", 0)}
                else:
                    try:
                        txt = read_text(p, ocr=ocr, meta=meta)
                    except Exception as e:      # noqa: BLE001
                        if verbose:
                            print("  ! 读取失败 %s：%s" % (f, e))
                        continue
                arts = split_articles(txt)
                if len(arts) < 5 and f.lower().endswith(".docx"):
                    # 文字层丢了条号（Word 自动编号）→ 从 numbering.xml 确定性复原，
                    # 判据是 lvlText == 「第%1条」，不是按段落顺序数
                    auto = docx_autonum_articles(p)
                    if len(auto) >= 5:
                        arts = auto
                        if verbose:
                            print("  · 自动编号复原 %d 条：%s" % (len(auto), f))
                if len(arts) < 5:           # 文字层缺失或不是条文式文件（公告/决定）
                    if verbose:
                        print("  ! 条文不足 %d 条，跳过：%s" % (len(arts), f))
                    continue
                is_pdf = f.lower().endswith(".pdf")
                d = {}
                for no, paras in arts:
                    t = tidy_article(paras, is_pdf)
                    if t:
                        d[no] = t
                self.laws[name] = {
                    "file": os.path.relpath(p, reg_root).replace(os.sep, "/"),
                    "arts": d,
                    # 文件名里若带「（2025年修订）」，记下版本年，供批注标注版本差异
                    "year": (re.search(r"[（(](20\d\d)\s*年", f) or [None, ""])[1],
                    "reliable": _is_ordered(d),
                    # 有页面是 OCR 出来的 → 批注里要提示「个别字可能有误」
                    "ocr": bool(meta.get("ocr_pages")),
                }
                self.titles.append(name)
                if verbose:
                    print("  · %-46s %3d 条" % (name[:46], len(d)))
        self.titles.sort(key=len, reverse=True)
        return self

    # ---------- 解析台账 H 列 ----------
    def resolve(self, raw: str, seen: list) -> tuple:
        """把《…》里的名字解析成 06 里的规范名。

        `seen` 是**同一单元格内此前的全称**列表（按出现顺序）。简称歧义时先回溯它，
        例如「《私募投资基金监督管理条例》第三十二条…；《条例》第三十一条」——
        《条例》必须落到前一句那部条例上，而不是顺延到紧挨着的另一部办法。

        返回 (规范名 or None, 是否靠简称推断)。
        """
        n = norm_name(raw)
        if not n:
            return None, False
        if n in self.laws:
            return n, False
        if n in ALIAS and ALIAS[n] in self.laws:
            return ALIAS[n], n != ALIAS[n]
        # ① 简称歧义：先在本文本格已出现的全称里找「包含这个简称」的
        if n in AMBIGUOUS or len(n) <= 4:
            cand = [t for t in reversed(seen) if n in t]
            if cand:
                return cand[0], True
            cand = [t for t in self.titles if t.endswith(n)]
            if len(cand) == 1:
                return cand[0], True
        # ② 全称但写法有差异（空格/括号）：唯一包含关系才认
        cand = [t for t in self.titles if n and (n in t or t in n)]
        if len(cand) == 1:
            return cand[0], True
        return None, False

    def parse_h(self, h: str) -> list:
        """把 H 列拆成 [{raw, law, inferred, arts:[{no, kuan, xiang}]}]。

        H 列一段形态如：《私募办法》第二十三条第九项、第三十三条
        《…》之间用「；」分隔，条款之间用「、」分隔，不能按「、」切（会把条切碎）。
        """
        h = str(h or "")
        if not h.strip():
            return []
        segs = []
        pos = 0
        for m in re.finditer(r"《([^》]{1,60})》", h):
            if segs == [] and m.start() > pos:
                # 第一条法规之前的零星文字（极少），忽略
                pass
            elif segs:
                segs[-1]["tail"] = h[pos:m.start()]
            segs.append({"raw": m.group(1), "tail": ""})
            pos = m.end()
        if segs:
            segs[-1]["tail"] = h[pos:]
        out, seen = [], []
        for s in segs:
            law, inferred = self.resolve(s["raw"], seen)
            if law and law not in seen:
                seen.append(law)
            arts = []
            for m in _ANY_ART_RE.finditer(s["tail"]):
                no = art_no(m)
                rest = s["tail"][m.end():m.end() + 24]
                # 先吃掉「款」，再在**剩下的部分**找「项」——
                # 「第六条第一款第一项」若直接在原串上匹配「项」会漏，
                # 因为串首是「第一款」（2026-11-03 实测：行212 的项全被吞）
                km = re.match(r"第([%s]+)款" % _CJK_NUM, rest)
                if km:
                    rest = rest[km.end():]
                xm = re.match(r"第([%s]+)项" % _CJK_NUM, rest)
                arts.append({"no": no, "kuan": km.group(1) if km else "",
                             "xiang": xm.group(1) if xm else ""})
            out.append({"raw": s["raw"], "law": law, "inferred": inferred, "arts": arts})
        return out


# ---------------------------------------------------------------- 缓存

# 解析器版本：改了切分/清洗/OCR 逻辑必须 +1，否则旧缓存永远命中，
# 出现「代码修好了、批注还是旧文本」的假象（2026-11-03 实测踩过）
_CACHE_VER = "3"


def dir_signature(reg_root: str) -> str:
    """06 目录指纹 = 源文件指纹 + 解析器版本。两者任一变化都重建缓存。"""
    import hashlib
    import json
    items = []
    for dp, dn, fn in os.walk(reg_root):
        for f in sorted(fn):
            if f.lower().endswith((".pdf", ".docx")):
                p = os.path.join(dp, f)
                st = os.stat(p)
                items.append([os.path.relpath(p, reg_root).replace(os.sep, "/"),
                              st.st_size, int(st.st_mtime)])
    items.sort()
    return hashlib.sha1((_CACHE_VER + json.dumps(items)).encode("utf-8")).hexdigest()


def load_texts(reg_root: str, texts_path: str, ocr: bool = True, verbose: bool = False) -> dict:
    """按文件粒度缓存**抽取出来的纯文本**（含 OCR 结果）。

    这一层与「解析结果」分开缓存的意义：改条号切分/清洗逻辑时只需重解析（秒级），
    不必把扫描件重新 OCR 一遍（几分钟）。文件没变就直接复用文本。
    """
    import json
    old = {}
    if texts_path and os.path.isfile(texts_path):
        try:
            with open(texts_path, encoding="utf-8") as fh:
                old = json.load(fh)
        except Exception:      # noqa: BLE001
            old = {}
    out, dirty = {}, False
    for dp, dn, fn in os.walk(reg_root):
        for f in sorted(fn):
            if not f.lower().endswith((".pdf", ".docx")):
                continue
            p = os.path.join(dp, f)
            rel = os.path.relpath(p, reg_root).replace(os.sep, "/")
            st = os.stat(p)
            sig = "%d|%d" % (st.st_size, int(st.st_mtime))
            e = old.get(rel)
            if isinstance(e, dict) and e.get("sig") == sig:
                out[rel] = e
                continue
            meta = {}
            try:
                txt = read_text(p, ocr=ocr, meta=meta)
            except Exception as ex:      # noqa: BLE001
                if verbose:
                    print("  ! 取文本失败 %s：%s" % (f, ex))
                txt = ""
            out[rel] = {"sig": sig, "text": txt,
                        "pages": meta.get("pages", 0), "ocr_pages": meta.get("ocr_pages", 0)}
            dirty = True
            if verbose:
                print("  · 取文本 %-44s %6d 字%s" % (
                    f[:44], len(txt),
                    "（OCR %d 页）" % meta.get("ocr_pages") if meta.get("ocr_pages") else ""))
    if texts_path and dirty:
        os.makedirs(os.path.dirname(texts_path), exist_ok=True)
        with open(texts_path, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False)
    return out


def load_or_build(reg_root: str, cache_path: str, ocr: bool = True, verbose: bool = False):
    """→ (LawIndex, 是否命中解析缓存)。文本层缓存独立，改解析逻辑不会重跑 OCR。"""
    import json
    texts_path = os.path.join(os.path.dirname(cache_path), "texts.json") if cache_path else ""
    recs = load_texts(reg_root, texts_path, ocr=ocr, verbose=verbose)
    sig = dir_signature(reg_root)
    if cache_path and os.path.isfile(cache_path):
        try:
            with open(cache_path, encoding="utf-8") as fh:
                data = json.load(fh)
            if data.get("sig") == sig:
                idx = LawIndex()
                idx.laws = data["laws"]
                idx.titles = sorted(idx.laws, key=len, reverse=True)
                return idx, True
        except Exception:      # noqa: BLE001  缓存坏了就重建，不中断
            pass
    idx = LawIndex().build(reg_root, ocr=ocr, verbose=verbose, records=recs)
    if cache_path:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        with open(cache_path, "w", encoding="utf-8") as fh:
            json.dump({"sig": sig, "laws": idx.laws}, fh, ensure_ascii=False)
    return idx, False
