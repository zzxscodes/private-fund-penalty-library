# -*- coding: utf-8 -*-
"""v2 三核心列（核心处罚事由 / 核心证据材料 / 核心处罚依据）+ 处罚措施 的【规则化】抽取。

设计原则（用户 2026-10-26 硬性要求）：不再靠模型逐条猜测，而是依据各信源决定书
固定的文章结构，用规则 + 正则抽取。各信源结构已实测总结：

【中基协 纪律处分/复核决定书】
  一、基本事实
    经查，X 存在以下违规行为：一是…… 二是…… 三是……  ← 核心处罚事由（加粗引导句+说明）
    X 上述行为违反了《……》第三十条 ……  ← “违反了”之后条文（含条款号）= 核心处罚依据（被违反的法条）
    以上事实有 …… 等证据予以确认 ……  ← “以上事实有…等予以确认”之间 = 核心证据材料
  二、纪律处分决定 → 处罚措施

【证监局 行政监管措施 / 行政处罚】
  经查/经查明，X 存在以下违法（规）事实：一是…/一、…  ……   ← 到 “违反了/上述事实/上述违法事实/本局认为/我会认为” 之前 = 核心处罚事由
  ……违反(了)? 《……》第X条 ……  ← 被违反的法条（含条款号）= 核心处罚依据（违反条款）
  根据/依据 《……》的规定，我局/本局决定：……  ← 授权法条（含条款号）= 核心处罚依据（处罚授权）
  ……我局/本局决定对你公司采取……措施。  ← 处罚措施
  注：证监局仅约 9% 在正文写 “以上事实有…等证据”，其余无显式证据清单 → 诚实标注。

【交易所 监管措施决定书】 结构同证监局（经查明…存在以下…事实 → 违反 → 本所决定），复用。

输出：dict(grounds, evidence, legal_basis, measures)，均为字符串（多条用 ； 连接）。

=====================================================================
三层契约（用户 2026-10-26 硬性要求，贯穿本技能一切需求）
---------------------------------------------------------------------
【规则优先】 本文件只做「规则 + 正则」抽取，不调用模型，结果确定可回归。
【仅当规则无确定内容才交模型】 规则对某列返回空 = 该列「无确定内容」，升格为
    PENDING_MODEL 哨兵（"待模型确认"），由下游（主会话/子代理）模型读取全文兜底填值。
【模型也没找到就诚实标注】 模型兜底后仍无 → 落 ABSENT 默认标记（按列语义：核心处罚
    事由/证据=无；核心处罚依据=正文未载明处罚依据；处罚措施=不明确），禁止编造、禁止留空。

落地：extract() 仍是纯规则层；mark_pending() 把空列升格为待模型确认；finalize()
把待模型确认列收口为「模型填值 或 诚实标记」。三函数串联即完整契约。
=====================================================================
"""
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

CN = "零一二三四五六七八九十百千"
# 条款号（容忍 OCR/HTML 多余空格，如 “第十一 条”）
ART = re.compile(r"第\s*[" + CN + r"]+\s*条(?:\s*第\s*[" + CN + r"]+\s*款)?(?:\s*第\s*[" + CN + r"]+\s*项)?")
LAW = re.compile(r"《([^》]+)》([^《。；;]*)")
# 定位“真正处罚决定”条款用【强】动词（复议/诉讼模板里只出现“监督管理措施不停止执行”/“如不服本处罚决定”，
# 不含这些强动词，因此不会被误判为处罚决定）
MEASURE_STRONG = re.compile(
    r"认定为不适当人选|责令改正|公开谴责|警示函|采取|罚款|撤销|暂停|没收|责令|认定为"
)
# 抽取“处罚措施”展示时仍允许较宽动词
MEASURE_KW = re.compile(
    r"认定为不适当人选|责令改正|公开谴责|警示函|采取|罚款|处罚|措施|监督管理措施|撤销|暂停|没收|责令|认定为"
)
# 真正的处罚决定标记（排除 文末复议/诉讼 模板里的 “本决定(书)”）
DEC_CLAUSE = re.compile(r"(?:现决定|我局决定|本局决定|本所决定|本会决定|我会决定|协会决定|本委决定|决定)[:：]?")
# 用于去掉措施文本开头的决定标记
DEC_MARK = re.compile(r"^(?:现决定|我局决定|本局决定|本所决定|本会决定|我会决定|协会决定|本委决定|决定)[：:]?\s*")
# 事实段停用词（核心处罚事由 在此之后截断）
# 2026-11-02 补充「上述行为/以上行为/以下行为/上述情形…」：这些是结论句开头（上述行为违反了…），
# 原清单漏收导致抽取在「违反了」处截断后仍把「上述行为」带进核心事由（用户实拍残头）。
REASON_STOPS = (
    "违反了", "上述事实", "上述违法事实", "本局认为", "我会认为", "经复核",
    "上述行为", "以上行为", "以下行为",
    "上述情形", "以上情形", "以下情形",
    "上述问题", "上述事项",
)


def _trim_grounds(s):
    """核心处罚事由收尾：去页码噪声 + 删末尾断句残头 + 序号列表化（一是→一、、每条一行）。
    唯一真源在 ledger_style。"""
    try:
        import ledger_style as _ls
        return _ls.normalize_grounds_items(_ls.clean_grounds_tail(_ls.strip_page_noise(s)))
    except Exception:
        return s

# ----------------------------------------------------------------- 申辩理由（v2 新增列）
# 决定书「申辩」段落结构（实测范例）：
#   AMAC 纪律处分：二、申辩意见 → 当事人提出以下申辩意见：…… （止于 三、）
#   AMAC 复核决定：二、申辩意见 → 申请人向协会提交了复核申请书……具体申辩意见如下：……（止于 三、）
#   csrc 行政处罚：……提出如下陈述申辩意见：（一）……（二）……（止于 经复核/本局认为）
#   明示未申辩：   "当事人未提出陈述、申辩意见，也未要求听证" → 无
DEFENSE_NO = re.compile(r"未提出陈述、申辩|未提出陈述申辩|未提出申辩|未提出听证|未提出陈述")
DEFENSE_OPEN_AMAC = re.compile(r"[二三四]、\s*申辩(?:意见|情况)\s*[:：]?\s*(.*)", re.S)
DEFENSE_OPEN_CSRC = re.compile(
    r"(?:提出如下|提出以下|提出了)[^。；]{0,20}(?:陈述、申辩意见|申辩意见|陈述申辩意见|申辩)[：:]?\s*(.*)", re.S)
# 申辩段终止标记（下一段/复核意见/审理意见/本局认为）
DEFENSE_STOPS = re.compile(r"经复核|本局认为|我局认为|我会认为|协会认为|经审理|[三四五]、")


def _clean(s):
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s).strip()
    return s.strip("，。、；;,. ")


def _enum(block):
    """按 一、二、 或 一是/二是 拆分有序条目，保留引导词。

    2026-11-02 修：旧版只认「一、」，而中基协文书绝大多数写的是「一是…二是…」，
    导致整段被当成一条（实测 344 行只抽到第一段、408 行子项被提升为一级条目）。
    「一是」只在句号后认，避免「统一是」「第一是」这类正文误拆。

    2026-11-02 二次修：分隔符不能只认「。」——证监局大量文书写成
    「一、…；二、…」（分号并列），只认句号会让第二条并进第一条（实测 行168/217/241）。
    现在允许 。；; 换行 四种前导。注意「三项制度」这类不会误拆——必须紧跟「、」或「是」。
    """
    parts = re.split(
        r"(?<=[。；;\n])(?=[一二三四五六七八九十]、|[一二三四五六七八九十]+是[^，,])", block)
    out = [_clean(p) for p in parts if _clean(p)]
    if out:
        return out
    return [_clean(x) for x in re.split(r"[。；;]", block) if _clean(x)]


def _law_refs_from(text, pattern, group=1):
    """按 pattern 抽取《法规名》+ 紧随条款号；去重（按 法规名+条款）。"""
    out, seen = [], set()
    for m in re.finditer(pattern, text, re.S):
        name = m.group(group).strip()
        tail = m.group(group + 1) if (m.lastindex or 0) >= group + 1 else ""
        arts = ART.findall(re.sub(r"\s+", "", tail))   # 去空格后再抽条款号
        key = (name, tuple(arts))
        if key not in seen and name:
            seen.add(key)
            s = "《" + name + "》"
            if arts:
                s += " " + "、".join(arts)
            out.append(s)
    return out


def _legal_basis(text):
    parts = []
    # (a) 被违反的法条：违反(了)? 《X》…（到 的/，/。/；/构成/规定 止）。csrc 在 《 前常插空格 → 加 \s*
    parts += _law_refs_from(
        text, r"违反(?:了)?\s*《([^》]+)》(.*?)(?:的|，|,|。|；|$|构成|规定)"
    )
    # (b) 处罚授权法条：根据/依据/依照 《X》…决定/规定
    parts += _law_refs_from(
        text, r"(?:根据|依据|依照)\s*《([^》]+)》(.*?)(?:决定|规定)"
    )
    # 跨两份清单去重（按 完整 “法规名+条款”，保留同一部法不同条款）
    seen, res = set(), []
    for s in parts:
        if s not in seen:
            seen.add(s)
            res.append(s)
    return "；".join(res)


SPEC_DEC = re.compile(
    r"(?:现决定|我局决定|本局决定|本所决定|本会决定|我会决定|协会决定|本委决定)[:：]?"
)


def _measure_clause(text):
    """定位真正的“处罚决定”条款。

    规则：
    1) 优先用带主语的“X决定”标记（我局决定/本局决定/现决定/本所决定/我会决定…），
       可天然避开【标题】里的裸“决定…采取责令改正”与【文末复议/诉讼模板】里的“本决定(书)”。
    2) 只在“决定”之后 220 字内出现【强】措施动词时才算候选。
    3) 跳过“本决定(书)/处罚决定书”里的“决定”——那是文书自称，不是处罚决定动词。
    4) 取第一个通过的候选（真正的主处罚决定通常最早出现；缴款/复议模板在其后）。
    5) 若无带主语标记，才兜底用裸“决定”（同样排除文书自称）。
    """
    for pat in (SPEC_DEC, re.compile(r"决定")):
        for m in pat.finditer(text):
            pre = text[max(0, m.start() - 3):m.start()]
            if re.search(r"本决|处罚决定|决定书", pre):
                continue
            after = text[m.end():m.end() + 220]
            if MEASURE_STRONG.search(after):
                return m.end()
    return None


def _measures(text):
    end = _measure_clause(text)
    if end is None:
        # 兜底：部分文书（如“根据《…》的规定，对你公司及…采取…措施”）无 “X决定” 标记，
        # 直接定位 “对X采取/责令改正/出具警示函…” 等措施动词句。
        m = re.search(
            r"对[^。；]*?(?:采取|责令改正|出具警示函|公开谴责|罚款|撤销|暂停|没收|认定为不适当人选|监管谈话|责令"
            # 2026-10-29 补：交易所监管函/监管警示不用「决定…采取」句式，而是
            # 「鉴于上述行为，现对你公司予以书面警示」→ 旧动词白名单全不命中，
            # 处罚措施落成「不明确」（国厚资产两案实测）。
            r"|予以书面警示|予以监管警示|予以通报批评|予以公开谴责|予以纪律处分|予以警示|出具监管函)"
            r"[^。；]*",
            text,
        )
        if not m:
            return ""
        seg = re.sub(r"\s+", "", text[m.start():])
        m_stop = re.search(
            r"如不服|如果对本|复议与诉讼期间|申请行政复议|提出行政复议|汇交|将罚款|将罚没款"
            r"|上述当事人应自收到|当事人应自收到|应自收到本处罚决定书|自收到本处罚决定书"
            # 监管函在措施句后紧跟「请你公司…充分重视/认真学习」的提醒语，一并截掉
            r"|，请(?:你|公司|你公司|相关当事人)",
            seg,
        )
        seg = seg[:m_stop.start()] if m_stop else seg
        return _clean(seg)
    seg = re.sub(r"\s+", "", text[end:])     # 先去空格，保证 “将 罚 款” 等可被截断正则匹配
    # 在第一个“复议/诉讼/缴款”模板标记处截断（这些标记绝不会出现在真正的处罚措施句内）
    m_stop = re.search(
        r"如不服|如果对本|复议与诉讼期间|申请行政复议|提出行政复议|汇交|将罚款|将罚没款"
        r"|上述当事人应自收到|当事人应自收到|应自收到本处罚决定书|自收到本处罚决定书",
        seg,
    )
    seg = seg[:m_stop.start()] if m_stop else seg
    seg = DEC_MARK.sub("", seg)              # 去掉 “我局决定：” 等前缀
    return _clean(seg)


def _measures_amac(block):
    """AMAC 纪律处分决定段（已是 ‘协会决定’ 之后的正文，不含标题/复议模板）：

    截掉复议模板与落款页脚，去掉 “根据《…》相关规定，协会决定” 引导语与 “作出以下纪律处分：”，去空格后返回。
    AMAC 措施动词多样（撤销登记/取消会员资格/公开谴责/加入黑名单/限期改正/认定为不适当人选…），
    故此处不依赖动词白名单，直接取结构化段落。
    """
    ms = re.search(r"如不服|如对本|当事人如|申诉复核|复核期间|根据《实施办法》|中国证券投资基金业协会|抄送", block)
    seg = block[:ms.start()] if ms else block
    seg = re.sub(r"^[\s\S]*?协会决定", "", seg)      # 去掉 “根据《…》相关规定，协会决定” 引导语（含换行）
    seg = re.sub(r"作出以下纪律处分[：:]?\s*", "", seg)
    seg = re.sub(r"\s+", "", seg)
    return _clean(seg)


def _find_amac_block(text):
    """定位 AMAC 纪律处分措施所在段落：优先 ‘二/三、纪律处分决定’，否则直接找 ‘协会决定’。

    注意：复核决定书（一、纪律处分情况/二、申辩意见/三、复核意见）没有独立的
    ‘二、纪律处分决定’ 段，只能靠 ‘协会决定’ 兜底；且其后可能没有 ‘如不服’ 等模板，
    故此处直接抓到文末，再由 _measures_amac 截掉页脚。
    """
    m = re.search(r"[二三]、\s*纪律处分决定(.*)", text, re.S)
    if m:
        return m.group(1)
    m = re.search(r"协会决定[：:]?\s*(.*)", text, re.S)
    if m:
        return m.group(1)
    return ""


def _cut_defense(seg):
    """申辩段收尾：截到下一段标记，去空格后清洗。"""
    s = DEFENSE_STOPS.search(seg)
    seg = seg[:s.start()] if s else seg
    return _clean(re.sub(r"\s+", "", seg))


def _defense_amac(text):
    """AMAC 申辩理由：定位 '二/三/四、申辩意见|申辩情况'，去掉引导语，截到下一段。

    OCR 文本会在字间插空格（"申辩 意见"），故先去空格再匹配。
    """
    flat = re.sub(r"\s+", "", text or "")
    m = DEFENSE_OPEN_AMAC.search(flat)
    if not m:
        return ""
    seg = m.group(1)
    # 去掉 "当事人提出以下申辩意见：" / "申请人向协会提交了…如下：" 等引导语
    seg = re.sub(r"^[^：:]{0,60}[：:]", "", seg, count=1)
    return _cut_defense(seg)


def _defense_csrc(text):
    """证监局/交易所 申辩理由：优先取实质申辩段落；没有 → 空（收口写「无」）。

    csrc 内联正文插空格（"陈述 申辩意见"），故先去空格再匹配。
    注意：不能先用 DEFENSE_NO（"未提出陈述、申辩"）整体判否——多当事人文书中
    常常是「一部分人提出申辩、另一部分人未提出」（如〔2026〕4号：新动力/毛凤丽/毛军亮
    均提出申辩），整体判否会把真实存在的申辩段落抹掉。故顺序必须是：
    【先找实质申辩段落 → 找不到才判为无】。
    """
    flat = re.sub(r"\s+", "", text or "")
    m = DEFENSE_OPEN_CSRC.search(flat)
    if m:
        return _cut_defense(m.group(1))
    # 无实质申辩段落（含明示"未提出陈述、申辩意见"的情形）→ 空，由契约收口为「无」
    if DEFENSE_NO.search(flat):
        return ""
    return ""


# ----------------------------------------------------------------- AMAC
def extract_amac(text):
    res = {"grounds": "", "evidence": "", "legal_basis": "", "measures": "", "defense": ""}
    # 事实段标题：多数是「一、基本事实」，广东逸信等写「一、事先告知情况」；
    # OCR 常在字间插空格，故处处容忍 \s*（实测「存在以下违 规行为」断行，导致整段漏抓）。
    bf_m = re.search(r"一、\s*(?:基本事实|事先告知情况|违规事实|违法事实|事实与理由)(.*?)(二、|$)",
                     text, re.S)
    bf = bf_m.group(1) if bf_m else text

    start = 0
    m_open = re.search(r"存在以下\s*(?:违法违规|违规|违法|违反|不当|问题)?\s*"
                       r"(?:事实|行为|情形|问题)\s*[:：]", bf)
    if m_open:
        start = m_open.end()
    # 事由终点（2026-11-02 修）：旧版遇第一个「违反了」就截断，而中基协文书多条违规时
    # 每条末尾都带「相关行为违反了《…》」，导致只抽到第一条（行36 丢了整条「委托他人
    # 行使投资管理职责」；行344 丢了「未及时更新登记信息」）。
    # 正确终点是**全部条目列完之后**的总结句 —— 证据句或「以上/上述行为违反」（最后一次）。
    stops = []
    for pat in ("以上事实有", "上述事实有", "以下事实有", "以上行为有", "上述行为有",
                "上述违法事实", "上述违法违规事实", "上述违规事实"):
        m = re.search(pat, bf[start:])
        if m:
            stops.append(m.start())
    viol = [m.start() for m in re.finditer(r"(?:以上|上述)行为违反了?", bf[start:])]
    if viol:
        # 出现多次 = 每条各自总结，取最后一次；只出现一次 = 全篇总结句
        stops.append(viol[-1] if len(viol) >= 2 else viol[0])
    end = min(stops) if stops else (re.search(r"违反了", bf[start:]).start()
                                    if re.search(r"违反了", bf[start:]) else len(bf) - start)
    reason = bf[start:start + end]
    reason = re.sub(r"^(?:经查明|经查)[^：:]*[：:]", "", reason)
    res["grounds"] = _trim_grounds("；".join(_enum(reason)))

    res["legal_basis"] = _legal_basis(bf)

    m_e1 = re.search(r"以上事实有|上述事实有|以下事实有|以上行为有|上述行为有|"
                     r"上述违法事实[，,]?\s*有|上述违法违规事实[，,]?\s*有", bf)
    m_e2 = re.search(r"等予以确认|予以确认|等证据证明|证据证明", bf)
    if m_e1 and m_e2 and m_e2.start() > m_e1.start():
        res["evidence"] = _clean(bf[m_e1.end():m_e2.start()])
    else:
        # AMAC 少数文书不以 “以上事实有…等证据” 句式列明 → 诚实标注，避免 B10 留空
        res["evidence"] = "未列举证据材料"

    res["defense"] = _defense_amac(text)

    m_dec = re.search(r"[二三]、\s*纪律处分决定(.*)", text, re.S)
    if m_dec:
        block = m_dec.group(1)
        m_end = re.search(r"如不服|如对本|当事人如|申诉复核|根据《实施办法》第三十九条", block)
        block = block[:m_end.start()] if m_end else block
        res["measures"] = _measures_amac(block)
    else:
        res["measures"] = _measures_amac(_find_amac_block(text))
    return res


# ----------------------------------------------------------------- 证监局 / 交易所
# 决定段起点：事实段到此为止。**不能**用「违反了」当终点（见 _grounds_csrc 注释）。
DEC_BOUND = re.compile(r"(?:我局|本局|本会|本所|协会|本委|现)\s*决定|(?:^|[。；])决定(?:对|给|采|予)")
# 条目内部的「违规定性句」起点，两档：
#   A 档（带主语前缀「上述/以上…行为/情形」）—— 这是**结论句**，切，且允许「不符合」；
#   B 档（裸动词）—— 只认「违反了/未恪守/未依照」：
#     · **不能**加「未按」——「一、未按规定披露信息」的「未按」是事实本身；
#     · **不能**加裸「不符合」——「资产证明文件不符合《私募办法》对合格投资者的认定标准」
#       是事实的一部分，切了会丢信息（行168/179 实测，旧版把两条都削成半句）。
ITEM_CUT_A = re.compile(
    r"[，,。；;]?\s*(?:上述|以上|前述)(?:行为|情形|事实|问题)?[^。]{0,8}?"
    r"(?:违反了?|不符合|未恪守|未依照)")
ITEM_CUT_B = re.compile(r"[，,。；;]?\s*(?:违反了?|未恪守|未依照)")


def _item_cut_pos(s):
    """返回条目内应截断的位置（两档取更早者），无则 -1。"""
    ps = [m.start() for m in (ITEM_CUT_A.search(s), ITEM_CUT_B.search(s)) if m]
    return min(ps) if ps else -1
# 引导句前言：「你公司在开展私募基金业务活动中，存在以下问题：」
LEAD_IN = re.compile(r"存在以下|开展私募基金业务|经查|现将|如下|问题[:：]\s*$")
MARK = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|[一二三四五六七八九十]+是|\d+[.、])")
# MARK 是「行首」判据，切前言时要用「任意位置」版本（页眉粘住的正文不在行首）
MARK_ANY = re.compile(r"(?:[一二三四五六七八九十]+、|[一二三四五六七八九十]+是|\d+[.、])")


def _grounds_csrc(text):
    """证监局/交易所「核心处罚事由」的唯一实现（抽取器与修复脚本共用，禁止第二套）。

    2026-11-02 根因修复：旧版把 REASON_STOPS（含「违反了」）当终点取 min，
    于是遇到「一、事实。违反了《X》第Y条。二是…」就在**第一条的定性句**处整段截断，
    后续条目全部丢失（实测 85 行，行330 最典型：F 只有「一是…咨询服务。」，
    H 列却抽到了「二是」的法条——说明抽取器看得到第二条，只是没写进 F）。

    新口径（顺序不能颠倒）：
      ① 先按枚举标记**分条**（「一、」/「一是」/「1.」）——分条必须在剔尾之前，
         否则又回到「在第一条定性句处截断」的老路；
      ② 再**逐条**剔除尾部的违规定性句（依据归 H 列，与库里 240 个多条目行的写法一致）；
      ③ 丢掉引导句前言（只在不带序号且有 ≥2 条时才丢，防单条案例误删全文）；
      ④ 序号交给 normalize_grounds_items 统一成「一、」并一行一条。
    """
    start = 0
    m_list = re.search(r"存在以下(?:违法|违规|违反|不当|问题)?(?:事实|行为|情形)[:：]", text)
    if m_list:
        start = m_list.end()
    else:
        m_k = re.search(r"(?:经查明|经查)[，,]?", text)
        if m_k:
            start = m_k.end()

    seg_end = len(text) - start
    m_dec = DEC_BOUND.search(text[start:])
    if m_dec:
        seg_end = m_dec.start()
    seg = text[start:start + seg_end]
    # 删掉「（…统一社会信用代码：XXXX）」这类括注；主体名在 D 列已有，事由只留行为。
    seg = re.sub(r"[（(][^（）()]{0,80}?统一社会信用代码[:：]?[0-9A-Za-z]{15,20}[）)]", "", seg)

    items = _enum(seg)
    if not items:
        return ""

    # ③ 引导句前言：不带序号的那一条，且后面还有带序号的条目 → 只切掉前言本身。
    #    **不能整条丢**——网页另存 PDF 时页眉/导航会粘在第一条前面（行2 实测：
    #    「繁體|宁波监管局当前位置：…一是登记备案信息不准确」是一条，整条丢就把第一条也丢了）。
    if len(items) > 1 and not MARK.match(items[0]):
        m0 = MARK_ANY.search(items[0])
        if m0:
            items[0] = items[0][m0.start():]
        else:
            items = items[1:]

    cleaned = []
    for it in items:
        m = _item_cut_pos(it)
        it2 = it[:m].strip() if m >= 0 else it
        it2 = re.sub(r"^\s*[^。：:]{0,12}[：:]\s*", "", it2) if not MARK.match(it2) else it2
        it2 = it2.strip().rstrip("，,；;")
        if not it2:
            it2 = it.strip()
        try:
            import ledger_style as _ls
            it2 = _ls.tidy_grounds_item(it2)
        except Exception:
            if it2 and not it2.endswith(("。", "；", ";")):
                it2 += "。"
        cleaned.append(it2)
    # 丢掉空的悬空序号条目（「二、」后没内容）——中基协决定书的段落标题被切碎时常出现，
    # 留着会变成「二。」这种残行（行45/58/92 实测）。
    cleaned = [c for c in cleaned
               if len(re.sub(r"^\s*[一二三四五六七八九十]+、", "", c).strip("。；;，, ")) >= 4]
    # 注意：**不要**在这里再按关键词（「开展私募基金业务」「经查」…）过滤短条目 ——
    # 那是第二套前言识别，会把正文条目当引导句删掉（行465 实测：
    # 「二、缺少开展私募基金业务必要的从业人员。」因含「开展私募基金业务」被误删）。
    # 前言已经在上面按「无序号」精确切掉了。
    return _trim_grounds("\n".join(cleaned))


def extract_csrc(text):
    res = {"grounds": "", "evidence": "", "legal_basis": "", "measures": "", "defense": ""}
    # 事由段落一律走 _grounds_csrc（唯一实现）；此处不再重复截断逻辑，
    # 否则「决定段终点」与「逐条剔尾」两套规则会各自漂移（§0.5.1 教训）。
    res["grounds"] = _grounds_csrc(text)

    res["legal_basis"] = _legal_basis(text)

    m_e1 = re.search(r"以上事实有|上述事实[，。].*?有|上述违法事实[，。].*?有", text)
    m_e2 = re.search(r"等证据|证据予以确认|予以确认", text)
    if m_e1 and m_e2 and m_e2.start() > m_e1.start():
        res["evidence"] = _clean(text[m_e1.end():m_e2.start()])
    else:
        # 证监局/交易所绝大多正文不逐项列明证据 → 诚实标注，避免 B10 留空
        res["evidence"] = "未列举证据材料"

    res["measures"] = _measures(text)
    res["defense"] = _defense_csrc(text)
    return res


# ----------------------------------------------------------------- 分发
def extract(text, source, person=False):
    if not text:
        return {"grounds": "", "evidence": "", "legal_basis": "", "measures": "", "defense": ""}
    text = text.replace("\r", " ")
    if source in ("amac", "amac_jg", "amac_ry"):
        return extract_amac(text)
    return extract_csrc(text)


# ----------------------------------------------------------------- 三层契约收口
# 规则层返回空 = 该列「无确定内容」，按契约升格为待模型确认哨兵。
PENDING_MODEL = "（待模型确认）"

# 模型兜底后仍无内容时，按列语义落的最终诚实标记（禁止留空、禁止编造）。
ABSENT_MARKERS = {
    "grounds": "无",                        # 核心处罚事由（正文几乎必载，极罕见无）
    "evidence": "无",                       # 核心证据材料（正文确无列示证据时）
    "legal_basis": "正文未载明处罚依据",     # 核心处罚依据
    "measures": "不明确",                    # 处罚措施（与 ledger-format.md 四.5 一致）
    "defense": "无",                         # 申辩理由：模型复核确认未申辩 → 写「无」
}


def mark_pending(rule_result):
    """规则穷尽后，四列中任何空字段升格为 PENDING_MODEL，触发下游模型兜底层。

    返回新 dict，原 rule_result 不动。
    """
    return {k: (v if v else PENDING_MODEL) for k, v in rule_result.items()}


def finalize(pending, model_filled=None, absent_markers=None):
    """模型兜底收口（三层契约第三层）。

    pending:        mark_pending() 的输出（含 PENDING_MODEL 哨兵）。
    model_filled:   可选 dict{列名: 模型填值}，仅用于覆盖 PENDING_MODEL 列。
    absent_markers: 可选 dict{列名: 最终标记}，缺省用 ABSENT_MARKERS。

    规则：
      - 非哨兵列 → 原样保留（规则确定的内容）；
      - 哨兵列 + 模型填了 → 用模型填值；
      - 哨兵列 + 模型也没填 → 落 absent_markers（无/待确认类诚实标记）。
    """
    model_filled = model_filled or {}
    absent = absent_markers or ABSENT_MARKERS
    out = {}
    for k, v in pending.items():
        if v == PENDING_MODEL:
            out[k] = model_filled.get(k) or absent.get(k, "无")
        else:
            out[k] = v
    return out


def extract_with_contract(text, source, person=False, model_filled=None):
    """一键走完契约：规则抽取 → 升格待模型确认 → 收口（模型填值 或 诚实标记）。

    这是对外主入口：调用方拿到的永远是「有确定内容 或 诚实标记」的值，绝不会出现
    裸空串（裸空串 = 漏标，会让下游误以为"已填"）。
    """
    rule = extract(text, source, person=person)
    return finalize(mark_pending(rule), model_filled=model_filled)


if __name__ == "__main__":
    import json
    src = sys.argv[1] if len(sys.argv) > 1 else "csrc"
    txt = sys.stdin.read()
    print(json.dumps(extract(txt, src), ensure_ascii=False, indent=2))
