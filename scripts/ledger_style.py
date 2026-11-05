# -*- coding: utf-8 -*-
"""台账格式冻结规格 v2（唯一真源）。

!!! 本文件定义的样式是「冻结标准」，任何一次更新都不得改动 !!!
背景：曾因每次更新都用新脚本重新生成台账，导致列宽、配色、字体、列顺序反复变化。
现在的规则是：
  * 台账文件已存在 → 只做「行追加」，沿用文件里已有的表头/列宽/配色/字体，
    新增数据行直接复制上一数据行的样式，绝不重建、绝不重排、绝不换风格。
  * 台账文件不存在 → 才用本文件的规格新建，且新建结果必须与既有库完全一致。

被 ledger.py 与 init_library.py 共同引用，保证「新建」与「追加」产出同一种排版。

v2（2026-10-26）：一行=一个案例（机构+高管合并）；14 列；无文号列；
新增挂网发布日期（排序键）/核心处罚事由/核心证据材料/核心处罚依据三核心列。
规格文档：references/ledger-format.md。
"""
import os
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

LEDGER_NAME = "04_综合索引台账【总表】.xlsx"

# ---------- 运维脚手架目录（唯一真源，禁止硬编码本机路径） ----------
# 案例库「根目录」只放交付物（01~06 分类目录 + 台账 + .md 记录），所有备份/缓存/日志/排除区/
# 归档区等运维留痕一律放在「库父目录/_案例库维护留痕/<库名>/」下，绝不写进库内或库父目录顶层。
# 2026-10-30 修正：此前 verify/run_update 用 os.path.join(root,"_work")、repair_attach/split 用
# root+"_work" 拼接，把 _work 写进了库内（或生成库名+_work 的兄弟目录），污染交付物。统一收敛到此。
MAINT_AREA = "_案例库维护留痕"


def maint_dir(root):
    """返回本库的运维主目录：<库父目录>/_案例库维护留痕/<库名>/（可写、可恢复、不在交付物内）。"""
    root = os.path.abspath(root)
    parent = os.path.dirname(root)
    libname = os.path.basename(root.rstrip("/\\")) or "library"
    return os.path.join(parent, MAINT_AREA, libname)


def work_dir(root):
    """返回 _work（备份/缓存/校验报告）目录，位于运维主目录下。"""
    return os.path.join(maint_dir(root), "_work")


def case_excluded_dir(root):
    """非私募误收排除区：运维主目录下 _case_excluded_non_pf/<来源>/。"""
    return os.path.join(maint_dir(root), "_case_excluded_non_pf")


def case_archive_dir(root):
    """窗口外移出归档区：运维主目录下 _case_archive_out_of_window。"""
    return os.path.join(maint_dir(root), "_case_archive_out_of_window")

# ---------- 工作表 ----------
SHEET_CASE = "综合索引台账"
SHEET_LAW = "配套法规依据清单"

# ---------- 表头（列顺序冻结，不得增删改序） ----------
CASE_HEADERS = [
    "序号", "处罚时间", "挂网发布日期", "被处罚机构/人",
    "处罚措施（机构+人员）", "核心处罚事由", "核心证据材料", "核心处罚依据",
    "申辩理由", "违规类型", "机构来源", "案例文件", "存放路径", "原文链接", "是否典型案例",
]
# 2026-10-27 用户定稿：法规清单**只保留能稳定取到的列**。
# 已移除「文号 / 发布机关 / 发布日期 / 施行日期 / 效力状态 / 留存形式」——
# 这 6 项无法保证每部法规都取到（网页型法规常缺文号与日期，官方 PDF 也无元数据层），
# 留着只会产出半空列。留存形式虽可由扩展名推出，用户一并要求去掉。
LAW_HEADERS = [
    "序号", "法规类别", "法规名称", "库内案例引用数", "文件名", "存放路径", "篇幅",
]

# ---------- 列宽（冻结） ----------
# 2026-10-27 实测修正：原宽度按拉丁字符单位定义，但台账内容多为中文
# （每汉字≈2 宽单位），导致机构名/法规名/URL 被截断。现按 CJK 修正：
# 中文列整体放大约 1.6~2 倍；日期列(12→16)容纳 "2026-10-23"；
# URL 列(42→60)容纳 csrc 典型链接(~55~62 字符)。
CASE_WIDTHS = [6, 16, 16, 42, 46, 64, 54, 54, 54, 34, 14, 54, 42, 60, 10]
LAW_WIDTHS = [6, 30, 50, 14, 48, 40, 10]

# ---------- 自动换行的列（1-based） ----------
# E 处罚措施 / F 核心处罚事由 / G 核心证据材料 / H 核心处罚依据 / I 申辩理由 /
# L 案例文件 / M 存放路径 / N 原文链接
CASE_WRAP_COLS = (5, 6, 7, 8, 9, 12, 13, 14)
# 法规清单：法规名称 / 文件名 / 存放路径
LAW_WRAP_COLS = (3, 5, 6)

# ---------- 列下标（0-based，供 ledger.py / verify_library.py 引用） ----------
# A=0 序号 | B=1 处罚时间 | C=2 挂网发布日期 | D=3 被处罚机构/人 | E=4 处罚措施
# F=5 核心处罚事由 | G=6 核心证据材料 | H=7 核心处罚依据 | I=8 申辩理由
# J=9 违规类型 | K=10 机构来源 | L=11 案例文件 | M=12 存放路径 | N=13 原文链接 | O=14 是否典型案例
COL = {
    "seq": 0, "decided": 1, "publish": 2, "subject": 3, "measures": 4,
    "grounds": 5, "evidence": 6, "legal_basis": 7, "defense": 8, "category": 9,
    "source": 10, "filenames": 11, "paths": 12, "urls": 13, "typical": 14,
}

# ---------- 配色与字体（冻结） ----------
HEADER_FILL = "305496"      # 表头底纹（深蓝）
HEADER_FONT_COLOR = "FFFFFF"
BORDER_COLOR = "BFBFBF"     # 细边框（浅灰）
BORDER_STYLE = "thin"
FREEZE = "A2"


def header_cell(cell):
    """套用冻结的表头样式"""
    cell.font = Font(bold=True, color=HEADER_FONT_COLOR)
    cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _border()


def data_cell(cell, wrap=False):
    """套用冻结的数据行样式（字体保持工作簿默认，只设边框与对齐）"""
    cell.border = _border()
    cell.alignment = Alignment(vertical="top", wrap_text=wrap)


def _border():
    s = Side(style=BORDER_STYLE, color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def new_workbook():
    """按冻结规格创建全新台账（仅在文件不存在时调用）"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_CASE
    _init_sheet(ws, CASE_HEADERS, CASE_WIDTHS, CASE_WRAP_COLS)
    return wb


def _init_sheet(ws, headers, widths, wrap_cols):
    from openpyxl.utils import get_column_letter

    ws.append(headers)
    for i in range(1, len(headers) + 1):
        header_cell(ws.cell(row=1, column=i))
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = FREEZE


def ensure_law_sheet(wb):
    """确保「配套法规依据清单」工作表存在（不改动已有工作表）"""
    from openpyxl.utils import get_column_letter

    if SHEET_LAW in wb.sheetnames:
        return wb[SHEET_LAW]
    ws = wb.create_sheet(SHEET_LAW)
    ws.append(LAW_HEADERS)
    for i in range(1, len(LAW_HEADERS) + 1):
        header_cell(ws.cell(row=1, column=i))
        ws.column_dimensions[get_column_letter(i)].width = LAW_WIDTHS[i - 1]
    ws.freeze_panes = FREEZE
    return ws


def append_row(ws, values, wrap_cols):
    """在末尾追加一行：沿用上一数据行的样式；没有上一行时用冻结规格。

    不触碰表头、列宽、配色、冻结窗格——即「只做内容增加」。
    """
    from copy import copy

    ws.append(values)
    r = ws.max_row
    prev = r - 1
    for i in range(1, len(values) + 1):
        dst = ws.cell(row=r, column=i)
        if prev >= 2:
            src = ws.cell(row=prev, column=i)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
        else:
            data_cell(dst, wrap=(i in wrap_cols))
    return r


def renumber(ws):
    """把 A 列序号重排为 1..n（不改动任何样式）。

    2026-10-08 实测修复：按 `max_row` 无差别编号会把尾部空白行也写上序号，
    造出「幽灵行」。⇒ 只对 **B..N 有内容** 的行编号；空白行整体清空（含序号）。
    返回实际数据行数。
    """
    n = 0
    for row in ws.iter_rows(min_row=2, max_col=len(CASE_HEADERS)):
        has = any(c.value is not None and str(c.value).strip() != "" for c in row[1:])
        if has:
            n += 1
            row[0].value = n
        else:
            for c in row:
                c.value = None
    return n


# ======================================================================
# 处罚措施 / 核心处罚事由 文本规范化（2026-11-02 用户裁定，唯一真源）
# ----------------------------------------------------------------------
# 用户口径：
#  1) 处罚措施按主体分行：每个「管理人：…」「高管（姓名）：…」各占一行；
#     多个高管必须相邻；高管组与管理人组按「谁先在原文出现」决定先后。
#  2) 复核决定：在处罚措施/申辩理由里另起一段（空一行），不与原处分挤在一格。
#  3) 核心处罚事由：剔除末尾断句残头（如「。上述行为」「。 上述情形」），
#     这类残头是抽取在「违反了/上述事实」处截断后多带出的后半句开头，非核心事由。
# ======================================================================
_PARTY_PREFIX = re.compile(r'(?:管理人|高管（[^）]*）)[：:]')


def _split_party_entries(body):
    """把 'A：…；B：…；C：…' 拆成 ['A：…；','B：…；','C：…']，保留每条尾部标点。
    仅在「；」后紧跟主体前缀处切分，句内「；」不拆（如「…；你应加强…」）。"""
    parts = re.split(r'(?<=；)[ \t　]*(?=(?:管理人|高管（[^）]*）)[：:])', body)
    return [p.strip() for p in parts if p.strip()]


def _order_party_entries(parts):
    """多高管相邻；高管组/管理人组按谁先出现排序；非主体条目附后、保持原序。"""
    staff = [p for p in parts if p.startswith("高管（")]
    mgr = [p for p in parts if p.startswith("管理人")]
    other = [p for p in parts if not p.startswith("高管（") and not p.startswith("管理人")]

    def first_idx(pred):
        for i, p in enumerate(parts):
            if pred(p):
                return i
        return 10 ** 9

    if staff and mgr:
        if first_idx(lambda p: p.startswith("高管（")) < first_idx(lambda p: p.startswith("管理人")):
            ordered = staff + mgr
        else:
            ordered = mgr + staff
    else:
        ordered = staff + mgr
    return ordered + other


def paths_cell(paths):
    """台账 L 列（存放路径）的唯一取值口径（2026-11-02 用户裁定）。

    一个案例的多份决定书**必然同目录**，所以 L 列只写 1 条；
    历史上「按文件数把同一目录写 N 遍」被用户判为冗余
    （「同一个案例的多个决定书文件存放路径肯定是同一个的，不要按文件数量重复写」）。
    写入方一律走本函数，读取方按「1 条广播到 N 个文件」处理（见 verify_library.load_ledger）。
    """
    ps = []
    for p in (paths or []):
        p = str(p or "").strip()
        if p and p not in ps:
            ps.append(p)
    if not ps:
        return ""
    if len(ps) > 1:
        # 一行内出现多个不同目录属异常（v2 一行=一个案例），保留全量交给校验报错，不静默合并
        return "；".join(ps)
    return ps[0]


def format_measures_list(arr):
    """把 ['管理人：…','高管（X）：…', …] 排成多行文本（多高管相邻、组间按列表先后）。"""
    items = [strip_page_noise(str(x).strip()) for x in (arr or []) if str(x).strip()]
    if not items:
        return ""
    return "\n".join(_order_party_entries(items))


def normalize_measures_text(text):
    """规范化存量单元格文本：多主体分行 + 复核决定空行另起。幂等，可重复运行。"""
    if not text:
        return text
    s = strip_page_noise(str(text))
    note = ""
    idx = s.find("【复核决定")
    if idx >= 0:
        note = s[idx:].strip()
        body = s[:idx].rstrip(" \t　")
    else:
        body = s.rstrip()

    parts = _split_party_entries(body)
    # 仅当切出的每一段都以主体前缀开头时才重排（否则保守保留原文，绝不吞内容）
    if len(parts) >= 2 and all(_PARTY_PREFIX.match(p) for p in parts):
        out = "\n".join(_order_party_entries(parts))
    else:
        out = body.strip()
    if note:
        out = out + "\n\n" + note
    return out


# 残头短语：上述行为 / 相关行为 / 该行为 …（抽取在「违反了/上述事实」处截断后多带出的后半句开头）。
# 注意：**不含「等X行为/等X问题」**——那是「A、B、C等行为」的枚举收尾，是完整句，不能删。
_FRAG = (r'(?:上述|以上|以下|相关|有关|该|其|这些|此类)'
         r'(?:行为|情形|事实|事项|问题|内容|情况|违规行为|违法行为|违法事实|违规事实)')
_FRAG_TAIL = re.compile(r'(' + _FRAG + r')[。；;]?\s*$')
# 主语残尾的「动词性结尾」白名单：若残尾以这些词收尾，说明是完整谓语，不能连同主语一起删
_VERB_TAIL = re.compile(r'(存在|有|构成|实施|从事|进行|发生|发现|出现|存在以下)$')


def clean_grounds_tail(text):
    """删除核心处罚事由末尾的断句残头，幂等。

    两类残头（2026-11-02 实测）：
      A. 「…不一致。上述行为」        → 残头前是句末标点，直接删残头短语；
      B. 「…。武汉金瑞麟上述行为」    → 残头前是「短主语」，属结论句开头，整段删到上一个句末标点；
      C. 「…无法支持最终报价结果上述行为」→ 无句末标点，仅删残头短语。
    误删防护：残尾若以「存在/有/构成…」等谓语收尾，说明是完整句，仅删残头短语、不删主语。
    """
    if not text:
        return text
    t = str(text).strip()
    m = _FRAG_TAIL.search(t)
    if not m:
        return t
    frag_start = m.start(1)
    # 枚举收尾保护：「A、B、C等相关事项」「…等行为」是完整句的枚举收尾，不是残头，绝不删。
    if t[:frag_start].rstrip().endswith("等"):
        return t
    head = t[:frag_start]
    p = max(head.rfind("。"), head.rfind("；"), head.rfind(";"))
    tail_before = (head[p + 1:] if p >= 0 else head).strip()
    # B 类：残头前是短主语（≤12 字、且不以谓语收尾）→ 连同主语一起删到上一个句末标点
    if p >= 0 and 0 < len(tail_before) <= 12 and not _VERB_TAIL.search(tail_before):
        return t[:p + 1].rstrip()
    # A / C 类：只删残头短语（连同紧随的标点）；但删完若落在谓语上（「公司存在」）则不删
    if _VERB_TAIL.search(head.rstrip()):
        return t
    return (t[:frag_start] + t[m.end():]).rstrip()


# ---- 正文页码噪声（2026-11-02）：PDF 分页会把词切断，如「监控账 -2- 户」「第一 - 2 - 款」
# 去掉页码连同两侧空白即可复原；用前后非数字断言避免误伤 2026-2-24 这类日期。
# 2026-11-02 补：证监局网页另存件里的页码是**长破折号**「—2—」（行122「部分—2—所管理」），
# 旧版只认半角连字符，正则漏掉。三种破折号一并收（日期区间「2025—2026」前有数字，断言挡住）。
_PAGE_NO = re.compile(r'\s*(?<![0-9])[-—–]\s*\d{1,2}\s*[-—–](?![0-9])\s*')


def strip_page_noise(text):
    """删除正文里的「- 2 -」式页码。幂等。"""
    if not text:
        return text
    return _PAGE_NO.sub("", str(text))


# ---- 事由条目收尾（2026-11-02，_grounds_csrc 与 fix_grounds_multi 共用） ----
# 症状：切掉「违反了《X》」之后，条目尾部会剩「上述行为分别。」「以上情形。」这类残头
#       —— 因为原文是「…。上述行为分别违反了《X》…」，「分别」打断了正则的连续匹配。
# 对策：把「上述/以上/前述 + 行为/情形/事实/问题 + 分别/均/还/也」这种**只剩主语的渣**
#       反复从尾部剥掉，直到剥不动为止。只剥尾部，不碰句中。
_TAIL_DROSS = re.compile(
    r"[，,。；;]?\s*(?:(?:上述|以上|前述)(?:行为|情形|事实|问题|事项)?(?:分别|均|还|也)?"
    r"|相关行为|上述相关行为|分别|均)"
    r"\s*[，,。；;]?\s*$")
# 说明：「以上行为严重违反了…」切走后会留下裸主体名（「…报销费用。国彤万和。」，行245 实测）。
# 这里**不加自动剥离规则**——「…由国彤万和负责」这类正常句尾也是「主体名+句号」，
# 自动剥离会把正文剪掉。正确做法是那一行**不自动改**（见 fix_grounds_multi 的护栏说明）。
# 括号配平：OCR/网页另存会出现「（有限合伙)(以下简称…」这种一半全角一半半角
_PAREN_FIX = re.compile(r"(?<=[\u4e00-\u9fa5）)])\(")
_PAREN_FIX2 = re.compile(r"\)(?=[\u4e00-\u9fa5（(])")
# 中文夹单个数字（「风险特2征」「部分—2—所管理」残留）—— 典型的页码/脚注噪声。
# **判据必须带量词白名单**：第一版只写「CJK+数字+CJK」，把「等9只」「的2名」「于5人」
# 「年3月」全报了（实测 12 条假线索）。数字后接量词/日期单位就是正常表述。
CJK_DIGIT_CJK = re.compile(
    r"(?<=[\u4e00-\u9fa5])\d{1,2}"
    r"(?![\d年月日号只名个人家条款项次位份倍期届户笔张元万亿%％])"
    r"(?=[\u4e00-\u9fa5])")


def tidy_grounds_item(s):
    """单条事由收尾：去页码 → 括号配平 → 剥尾部残头 → 保句末标点。幂等。"""
    if not s:
        return s
    t = strip_page_noise(str(s)).strip()
    t = _PAREN_FIX.sub("（", t)
    t = _PAREN_FIX2.sub("）", t)
    for _ in range(4):
        t2 = _TAIL_DROSS.sub("", t).strip()
        if t2 == t or not t2:
            break
        t = t2
    t = t.rstrip("，,；;、 ")
    if t and not t.endswith(("。", "；", ";")):
        t += "。"
    return t


# ---- 核心处罚事由：序号列表化（2026-11-02 用户裁定） ----
# 边界分两档：
#   软边界 = 行首 / 句末标点(。；;) / 换行 / 逗号冒号引号收尾 —— 只用于中文数字序号
#            （实测 6 例「存在以下问题，一是…」「存在以下情形: 一是…」「…备案” 二是…」）
#   硬边界 = 行首 / 句末标点 / 换行 —— 阿拉伯数字序号只认硬边界，
#            避免把「23.74」「第3.3节」这类正文数字误当序号
_GROUNDS_SOFT_BOUND = r'[。；;\n，,：:””"）)]'
_GROUNDS_MARK_SOFT = re.compile(
    r'(?:^|(?<=' + _GROUNDS_SOFT_BOUND + r'))[ \t　]*'
    r'('
    r'[一二三四五六七八九十]+[、是]'
    r'|[（(][一二三四五六七八九十\d]+[）)]'
    r'|第[一二三四五六七八九十\d]+[，,、]'
    r')'
)
_GROUNDS_MARK_NUM = re.compile(
    r'(?:^|(?<=[。；;\n]))[ \t　]*(\d+[.、](?![0-9]))'
)
# 无分隔符粘连：OCR 丢标点导致「…存在虚假记载三是兼营…」（实测 1 例）。
# 仅在已有枚举上下文（存在 一、/一是 这类真条目起点）时启用，且只认「二~十是」，
# 规避「统一是」「第（三）是」等正文误判。
_GROUNDS_MARK_STUCK = re.compile(r'(?<=[\u4e00-\u9fa5])(?<![第])(?=[二三四五六七八九十]+是)')
_CN_D = "零一二三四五六七八九"


def _cn_num(s):
    """阿拉伯数字→中文数字（1~99）；已是中文则原样返回。"""
    if not s.isdigit():
        return s
    n = int(s)
    if n < 10:
        return _CN_D[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + _CN_D[n - 10]
    if n < 100:
        return _CN_D[n // 10] + "十" + (_CN_D[n % 10] if n % 10 else "")
    return s


def _norm_one_marker(seg):
    """把单条开头的序号统一成「一、」形式（保留其后原空格；吃掉序号后紧跟的逗号）。"""
    body = None
    m = re.match(r'([一二三四五六七八九十]+)[、是]', seg)
    if m:
        body = m.group(1) + "、" + seg[m.end():]
    if body is None:
        m = re.match(r'[（(]([一二三四五六七八九十\d]+)[）)]', seg)
        if m:
            body = _cn_num(m.group(1)) + "、" + seg[m.end():]
    if body is None:
        m = re.match(r'第([一二三四五六七八九十\d]+)[，,、]', seg)
        if m:
            body = _cn_num(m.group(1)) + "、" + seg[m.end():]
    if body is None:
        m = re.match(r'(\d+)[.、]', seg)
        if m:
            body = _cn_num(m.group(1)) + "、" + seg[m.end():]
    if body is None:
        return seg
    # 「一是，2016年…」→「一、，2016年…」的重复标点：序号后直接跟逗号则吃掉
    return re.sub(r'^([一二三四五六七八九十百]+、)[ \t　]*[，,]\s*', r'\1', body)


def normalize_grounds_items(text):
    """核心处罚事由：每个序号点独立成行，并把「一是/（一）/1./第一，」统一成「一、」。幂等。"""
    if not text:
        return text
    t = str(text)
    ms = list(_GROUNDS_MARK_SOFT.finditer(t))
    ms += list(_GROUNDS_MARK_NUM.finditer(t))
    if not ms:
        return t
    ms.sort(key=lambda m: m.start())
    # 去重叠（软边界匹配可能吃掉硬边界同一位置）
    uniq, last = [], -1
    for m in ms:
        if m.start() > last:
            uniq.append(m)
            last = m.end() - 1
    ms = uniq
    # 枚举上下文：存在「行首/句末标点后的中文数字序号」才允许修无标点粘连
    has_ctx = any(_GROUNDS_MARK_SOFT.match(t, m.start()) and
                  (m.start() == 0 or t[m.start() - 1] in "。；;\n")
                  for m in ms)
    if has_ctx:
        stuck = [m for m in _GROUNDS_MARK_STUCK.finditer(t)
                 if all(not (s <= m.start() < e) for s, e in ((x.start(), x.end()) for x in ms))]
        ms = sorted(ms + stuck, key=lambda m: m.start())
    segs = []
    pre = t[:ms[0].start()].strip()
    if pre:  # 序号前的前言（如「我局发现你公司存在以下问题。」）独占首行
        segs.append(pre.rstrip("，,"))
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(t)
        seg = t[m.start():end].strip().rstrip("，,")
        if seg:
            segs.append(_norm_one_marker(seg))
    return "\n".join(segs)


# ---------------------------------------------------------------------------
# 核心处罚事由（F 列）规范化 —— 2026-11-02 新增，唯一口径
# 背景：用户实拍 F 列「一、…」形态不统一 + 内容有残片。根因有二：
#   ① 旧抽取器 `(?:经查)[^：:]*[:：]` 吃到「统一社会信用代码：」的冒号，
#      把主体名和左括号切掉，只剩「91310109…E ）在管理…」（实测 5 行）；
#   ② 部分行的 F 取自证监局**网页摘要**而非决定书正文（「…等问题」结尾），
#      信息有压缩（行28 就漏了「损害部分投资者利益」）。
# 本函数只做「形态归一」，不改写事实：剥引导句、剥并列尾巴、补句号、替换称呼。
# ---------------------------------------------------------------------------
# 证监局网页另存为 PDF 时，**整页**（页脚+导航+索引号）会被当正文抽进来：
#   「中国政府网主办单位：…京ICP备05035542号…联系我们|法律声明|当前位置：首页>…专题」
#   「繁體|安徽监管局当前位置：首页>政务信息>…互动交流专题」
# 特征词命中才剥，正文以「互动交流」「专题」收尾，其后即当事人与事实。
SHELL_MARK = re.compile(r"中国政府网|京ICP备|京公网安备|网站识别码|联系我们\s*\|?\s*法律声明|"
                        r"当前位置：首页|请输入关键字|主办单位：|版权所有：|证监局介绍")
SHELL_TAIL = ("互动交流", "专题")


def strip_web_shell(t):
    """命中网页外壳特征时，从最后一个外壳尾巴词之后取正文，并去掉「XXX：」称呼语"""
    if not t or not SHELL_MARK.search(t):
        return t
    pos = -1
    for w in SHELL_TAIL:
        i = t.rfind(w)
        if i > pos:
            pos = i + len(w)
    if pos <= 0:
        return t
    rest = t[pos:]
    if "：" not in rest[:80] and ":" not in rest[:80]:
        return t                      # 「专题」后面没有称呼语 → 不是外壳，别乱截
    rest = re.sub(r"^[|｜>》\s\u3000]*[^：:。]{2,40}[：:]\s*", "", rest)
    return rest.strip()


GROUNDS_LEAD = (
    r"^我局对你公司开展[^，。；]{0,50}检查中发现[，,]\s*",
    r"^我局对你(?:公司|单位|企业)开展[^，。；]{0,50}检查[，,]?\s*",
    r"^我局对[^，。；]{0,60}开展[^，。；]{0,40}检查中发现[，,]\s*",
    r"^我局于[^，。；]{0,40}对你[^，。；]{0,30}开展[^，。；]{0,40}[，,]\s*",
    r"^我局(?:在|于)?[^，。；]{0,30}(?:现场)?检查中发现[，,]\s*",
    r"^根据相关行政监管措施决定书[，,]\s*",
    r"^经查(?:明)?[，,]?\s*",
    r"^我局发现[，,]?\s*",
)
GROUNDS_LEAD2 = (
    # 2026-11-02 修：主体名不止「你公司/公司」，多数证监局文书直接写机构全称
    # （「东方融投（福州）投资管理有限公司在开展私募基金业务过程中，存在以下问题：」），
    # 旧正则只认「你公司|公司」前缀 → 整句引导语留在 F 列，还带出「：。」残标点（行65/155/179）。
    # 允许前置任意不含句号/分号/冒号的 80 字，末尾顺带吃掉「：。」这类残标点。
    r"^[^。；\n：:]{0,80}存在以下"
    r"(?:违法|违规|违反|不当|问题)?(?:事实|行为|情形|问题)?[：:]\s*[。；，]?\s*",
    r"^存在以下(?:违法|违规|违反|不当|问题)?(?:事实|行为|情形|问题)?[：:]\s*[。；，]?\s*",
)
GROUNDS_TAIL = re.compile(r"[，,]?\s*等(?:问题|行为|情形|事实|情况)?\s*[。.]?\s*$")
CALL_YOU = re.compile(r"你(?:公司|单位|企业)")
# 事由列的越界信号：证据/认定/决定/申辩——出现即说明抽过头了
GROUNDS_CUT = re.compile(
    r"以上事实[，,]?\s*有|上述事实[，,]?\s*有|上述违法事实|证据证明|证据予以确认|"
    r"我局认为|本局认为|本所认为|我会认为|协会认为|经审理查明|足以认定|"
    r"我局决定|本局决定|本所决定|现决定|决定对你|"
    r"当事人提出|申辩意见|陈述申辩|"
    r"(?<=[。；])\s*(?:上述行为|以上行为|上述|该行为)(?:违反|构成)|"
    # 2026-11-02：证监局文书常写「…未按规定履行职责，违反《X》第四条…的规定。根据《X》第三十三条
    # 的规定，我局决定…」——「根据《X》第Y条」是**处罚授权条款**（属 H 列），不是违规事实。
    # 行312/314/365 的 F 列因此多带了授权条款（实测 3 行）。
    r"(?<=[。；])\s*(?:根据|依据)\s*《[^》]{2,40}》\s*第[一二三四五六七八九十百零\d]+")


def normalize_grounds(text, subject="", split_parallel=False):
    """F 列形态归一。

    subject           —— D 列的机构名（「你公司」等称呼替换成它，让每格独立可读）
    split_parallel    —— 对「存在 A、B、C 等行为」这种多项并列**按原文顿号拆分**
                         并加「一、二、三」。只在原文本身是并列结构时生效，
                         拆出来的每一项都是**原词**，不重写。
    """
    if not text:
        return text
    t = str(text).strip()
    # 0) 剥网页外壳（整页导航/页脚/索引号被当成正文，2026-11-02 行215/462 实测）
    t = strip_web_shell(t)

    # 1) 剥引导句（可能叠加两层：先「我局发现」，再「存在以下问题：」）
    for _ in range(2):
        for pat in GROUNDS_LEAD + GROUNDS_LEAD2:
            t2 = re.sub(pat, "", t, count=1)
            if t2 != t:
                t = t2.strip()
                break
        else:
            break

    # 1b) 句首信用代码残片兜底：`91310118575837696N）在开展…`
    m_code = re.match(r"^\s*\d{15,18}[A-Z0-9]?[ \t]*[）)]\s*", t)
    if m_code:
        t = t[m_code.end():]
        org = (subject.split("；")[0].split("、")[0].strip() if subject else "")
        if org and re.match(r"^在", t):
            t = org + t

    # 1c) 越界截断：事由列不该含证据/认定/决定/申辩内容（那一份该进 G/H/I 列）
    m_cut = GROUNDS_CUT.search(t)
    if m_cut and m_cut.start() > 10:
        t = t[:m_cut.start()].rstrip("，,；;。、 ")

    # 2) 多项并列拆分（仅当原文是「存在 …、…、… 等行为/情形」且尚未分点）
    if split_parallel and not re.search(r"^\s*[一二三四五六七八九十]+、", t, re.M):
        t = _split_parallel_items(t)

    # 2b) 拆分后条目里残留的「等情况，/等情形，」→「，」（原文并列尾巴，不是内容）
    if re.search(r"^\s*[一二三四五六七八九十]+、", t, re.M):
        t = re.sub(r"等(?:情况|情形)[，,]", "，", t)

    # 3) 剥并列尾巴「等问题/等行为/等情形」（分点模式下只剥最后一条的尾巴）
    if re.search(r"^\s*[一二三四五六七八九十]+、", t, re.M):
        lines = [x.strip() for x in t.split("\n") if x.strip()]
        if lines:
            lines[-1] = GROUNDS_TAIL.sub("", lines[-1])
            t = "\n".join(lines)
    else:
        t = GROUNDS_TAIL.sub("", t)

    # 4) 称呼替换为主体名
    if subject:
        org = subject.split("；")[0].split("、")[0].strip()
        if org and len(org) >= 6:
            t = CALL_YOU.sub(org, t)

    # 5) 补句末标点（分点逐条补；整段只在末尾补）
    #    条末半角「;」统一为全角「；」（公文体例），末条统一为「。」
    if re.search(r"^\s*[一二三四五六七八九十]+、", t, re.M):
        lines = []
        for x in t.split("\n"):
            x = x.strip()
            if not x:
                continue
            if x[-1] not in "。；;！？":
                x = x.rstrip("，,、") + "。"
            lines.append(x)
        if lines:
            lines[:-1] = [re.sub(r";$", "；", x) for x in lines[:-1]]
            lines[-1] = re.sub(r"[；;]$", "。", lines[-1])
        t = "\n".join(lines)
    else:
        if t and t[-1] not in "。；;！？":
            t = t.rstrip("，,、") + "。"

    # 6) 压空格（复用逐格清理的同一判据，避免两套实现）
    return _tighten_zh(t)


BAD_WORD = re.compile(r"未|违|不|擅自|挪用|侵占|虚构|隐瞒|超|承诺|混同|让渡|委托|从事|"
                      r"缺失|缺位|不当|虚假|误导|利益输送|内幕|操纵|借用|出借|"
                      r"将|向|通过|使用|以基金|利用")
LIST_HINT = re.compile(r"例如|包括|等客观|等信息|等资料|等文件|等材料|等协议|等凭证")


def _split_parallel_items(t):
    """把「存在 A、B、C 等行为」/「A、B、C」按顿号拆成带序号的条目（原词不重写）

    两档判据，从严到宽：
      档1（强）：句中出现「存在」→ 取「存在」之后的并列体，≥2 项、≥2 项带违规动词即拆。
      档2（弱）：句中无「存在」→ 要求 ≥3 项、≥2 项带违规词、且无「例如/包括/等资料」
                 这类**举例**提示，避免把「行业资讯、股票涨跌幅等客观信息」拆坏。
    """
    m = re.search(r"存在(.+?)(?:等(?:问题|行为|情形|事实|情况)?)?[。.]?$", t, re.S)
    if m:
        body = m.group(1)
        parts = [p.strip() for p in body.split("、") if p.strip()]
        if _ok_split(parts):
            return _emit(parts)
        return t

    if LIST_HINT.search(t) or len(t) > 300:
        return t
    body = re.sub(r"^(?:你公司|你单位|你企业|[^，。；]{0,40}?公司)", "", t)
    parts = [p.strip() for p in body.split("、") if p.strip()]
    if len(parts) >= 3 and _ok_split(parts):
        # 第一项常带主语（「管理的基金对外投资后未完成确权登记」），原样保留
        return _emit(parts)
    return t


def _ok_split(parts):
    """拆分闸门：≥2 项、**违规项占比 ≥50%**、且每项平均 ≥6 字

    占比闸门是为了挡住「日期枚举」「举例枚举」：
      行210 海德「2021年末、2022年末、2023年末、2024年6月30日」→ 0/4 → 不拆
      行139 南京广丰「法定代表人、办公地址、从业人员」→ 2/5=40% → 不拆
      行91  恒邦「委托…、未审慎…、不公平…、未按…、从事…」→ 5/5 → 拆
    """
    n = len(parts)
    if n < 2:
        return False
    # 任一项内含逗号 → 说明是**叙述体被顿号割裂**，不是并列条目（行139/265/210/420 实测）
    if any(("，" in p or "," in p) for p in parts):
        return False
    hit = sum(1 for p in parts if BAD_WORD.search(p))
    avg = sum(len(p) for p in parts) / n
    return hit >= 2 and hit / n >= 0.5 and avg >= 6


def _emit(parts):
    CN = "一二三四五六七八九十"
    return "\n".join("%s、%s" % (CN[i] if i < 10 else str(i + 1), p)
                     for i, p in enumerate(parts))


FULLWIDTH_MAP = {ord(c): ord(chr(ord(c) - 0xFEE0))
                 for c in "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
                 "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"}


# ---------------------------------------------------------------- 证监局辖区（唯一真源）
# csrc.gov.cn 的一级路径段 → 中文名。原来只在 verify_library.BN 里有一份，
# 2026-11-04 网页台账要按「地区」筛选，改由本模块持有、verify_library 反向引用，
# 避免同一份映射写两遍（网站改名/新增辖区时只改这里）。
BUREAU_CN = {
    'anhui': '安徽', 'beijing': '北京', 'chongqing': '重庆', 'dalian': '大连', 'fujian': '福建',
    'gansu': '甘肃', 'guangdong': '广东', 'guangxi': '广西', 'guizhou': '贵州', 'hainan': '海南',
    'hebei': '河北', 'heilongjiang': '黑龙江', 'henan': '河南', 'hubei': '湖北', 'hunan': '湖南',
    'jiangsu': '江苏', 'jiangxi': '江西', 'jilin': '吉林', 'liaoning': '辽宁', 'neimenggu': '内蒙古',
    'ningbo': '宁波', 'qingdao': '青岛', 'shaanxi': '陕西', 'shandong': '山东', 'shanghai': '上海',
    'shanxi': '山西', 'shenzhen': '深圳', 'sichuan': '四川', 'tianjin': '天津', 'xiamen': '厦门',
    'xinjiang': '新疆', 'yunnan': '云南', 'zhejiang': '浙江', 'ningxia': '宁夏', 'qinghai': '青海',
    'xizang': '西藏',
}


def _tighten_zh(s):
    """删中文语境空格 + 全角英数转半角（与 fix_cells 同一判据）"""
    s = s.translate(FULLWIDTH_MAP)
    def repl(m):
        a = s[m.start() - 1] if m.start() > 0 else ""
        b = s[m.end()] if m.end() < len(s) else ""
        aw = lambda ch: ch and ch.isascii() and (ch.isalnum() or ch in ")]}")
        return " " if (aw(a) and aw(b)) else ""
    return re.sub(r"[ \t\u3000]+", repl, s)
