# -*- coding: utf-8 -*-
"""v2 30天窗口案例库构建流水线（规则优先 + 契约兜底）。

输入：tmp_pilot/csrc.json（189 证监局）+ tmp_pilot/amac_30d.json（41 中基协）
输出：tmp_pilot/records_30d.json（v2 案例记录数组，机构+高管合并一行）

原则：
  * 三核心列 + 处罚措施文本 → extract_core（纯规则，三层契约兜底）；
  * 管理人名称 / 违规类型 / 处罚措施（管理人+高管）拆分 / 中基协机构人员配对 →
    规则启发式；规则无法确定处 → 诚实标注（管理人：不明确 / 高管：不明确），
    绝不编造。
"""
import os, re, json, sys, shutil
from html import unescape as _unescape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import subject_clean as _SC       # 碎片词表（NOISE_TOKENS）唯一事实源

# 证监局网页正文带 HTML 实体（&ensp; &nbsp; &amp; …，189 条里 169 条有）。不转义会把
# "&ensp;" 原样写进核心处罚事由/措施列、也写进 PDF，属明显脏数据（2026-10-27 实测 32 条中招）。
_CJK_SPACE = re.compile(r"(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])")


def tidy(s):
    """去 HTML 实体 + 合并汉字之间的多余空格（"二 是" → "二是"）。"""
    if not s:
        return s
    s = _unescape(str(s))
    s = re.sub(r"\s+", " ", s)
    s = _CJK_SPACE.sub("", s)
    return s.strip()
HERE = os.path.dirname(os.path.abspath(__file__))

import ssl as _ssl
_SSL = _ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = _ssl.CERT_NONE
_UA = {"User-Agent": "Mozilla/5.0"}
# 交易所决定书 PDF 落本地：复用 URL 哈希命名（与 fetch_exchange.cache_name 同规则），
# 带按域名 Referer + 节流 + 403 退避，避免重触发深交所静态域/北交所 WAF 频控。
def _exch_cache_name(url):
    """与 fetch_exchange.cache_name 完全同规则（sha1 前 20 位）。

    2026-10-30 修（爬取缺失根因）：原实现用【完整 40 位】sha1，而 `fetch_bse` 用【前 20 位】
    ——命名不一致，导致 build 永远命中不了 fetch 已下好的缓存、对北交所回退裸 HTTP →
    302 无限重定向 → 漏件（浙江浩坤昇发北交所决定书即此：PDF 明明已在
    `cache_exch/bse/016e432984d57e12f8ed.pdf`，build 却查不到、又下不动）。必须逐字一致。
    """
    import hashlib
    return hashlib.sha1((url or "").encode("utf-8")).hexdigest()[:20] + ".pdf"


def _exch_local(url, outdir):
    if not url:
        return ""
    if not url.startswith("http"):
        return url if os.path.exists(url) else ""
    sub = ("bse" if "bse.cn" in url else "sse" if "sse.com.cn" in url
           else "szse" if "szse.cn" in url else "")
    ref = {"bse": "https://www.bse.cn/", "sse": "https://www.sse.com.cn/",
           "szse": "https://www.szse.cn/"}.get(sub)
    cn = _exch_cache_name(url)
    p = os.path.join(outdir, cn)
    # 1) 复用 fetch_bse 缓存（fetch_exchange --cache-dir=<work>/cache_exch 下的 <sub>/<cn>）
    fetch_cache = os.path.join(os.path.dirname(outdir), "cache_exch", sub, cn) if sub else ""
    for cand in ([fetch_cache] if fetch_cache else []) + [p]:
        if os.path.exists(cand) and os.path.getsize(cand) >= 1024 and open(cand, "rb").read(5) == b"%PDF-":
            if cand != p:
                os.makedirs(outdir, exist_ok=True)
                shutil.copy2(cand, p)
            return p
    # 2) 裸 HTTP 下载（sse/szse 可行；bse 多数会 302）
    import urllib.request
    for attempt in range(4):
        try:
            req = urllib.request.Request(url, headers=(_UA | ({"Referer": ref} if ref else {})))
            data = urllib.request.urlopen(req, timeout=60, context=_SSL).read()
        except Exception:
            data = b""
        if data.startswith(b"%PDF") and len(data) >= 1024:
            os.makedirs(outdir, exist_ok=True)
            open(p, "wb").write(data)
            return p
        import time as _t
        _t.sleep(2 + attempt * 5)     # 403/异常 → 退避后重试
        data = b""
    # 3) 北交所反爬 302 → 浏览器会话兜底（先访问首页建会话，再取附件）
    if sub == "bse":
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as _pw:
                _b = _pw.chromium.launch(headless=True)
                _ctx = _b.new_context(user_agent=_UA.get("User-Agent", ""))
                _pg = _ctx.new_page()
                try:
                    _pg.goto("https://www.bse.cn/", timeout=45000, wait_until="domcontentloaded")
                except Exception:
                    pass
                _bd = _ctx.request.get(url, timeout=90000,
                                       headers={"Referer": "https://www.bse.cn/"}).body()
                _b.close()
            if _bd.startswith(b"%PDF") and len(_bd) >= 1024:
                os.makedirs(outdir, exist_ok=True)
                open(p, "wb").write(_bd)
                print("  [INFO] 北交所用浏览器会话补下附件成功：%s" % os.path.basename(p))
                return p
        except Exception as _e:
            print("  [WARN] 北交所浏览器兜底失败：%r" % (_e,))
    return ""
# 可迁移（2026-10-27）：一律用脚本自身目录，禁止写死盘符/用户名路径，
# 否则技能换机器后 import 直接失败（由 selftest_portable.py 守着）。
sys.path.insert(0, HERE)
import extract_core
# 分类 / 私募筛选：统一走 classify_filter（与 verify_library 同口径，单一事实源）
from classify_filter import classify, is_private_fund  # noqa: E402
# P3 名册闸门：文本判定 + AMAC 管理人名册（在册+已注销）复核，统一收录入口 decide()
from amac_lookup import decide, AmacRegistry, DEFAULT_DIR as _AMAC_REG_DIR  # noqa: E402

# 路径全部可在 main() 用命令行覆盖（滑动窗口流水线 run_update.py 会传各自工作目录）。
# 默认值仅保留给「脚本原地调试」场景。
CSRC_JSON = os.path.join(HERE, "csrc.json")
AMAC_JSON = os.path.join(HERE, "amac_30d.json")
EXCH_JSON = os.path.join(HERE, "exch.json")
OUT = os.path.join(HERE, "records_30d.json")
EXCLUDED_OUT = os.path.join(HERE, "excluded_30d.json")
AUDIT_OUT = os.path.join(HERE, "audit_amac_30d.json")


def _resolve_src(it):
    """决定书源文件路径：优先 local（抓取时已落盘），回退 pdf/src_pdf。

    坑（2026-10-27）：旧实现写死 `os.path.join(HERE, it["pdf"])`，而抓取产物里的
    pdf/local 多已是绝对路径 → join 后路径被整个替换成 HERE 下的相对路径，
    换工作目录即全部找不到文件（复制 0 份）。
    """
    for k in ("local", "pdf", "src_pdf"):
        v = it.get(k)
        if not v:
            continue
        if os.path.isabs(v) or os.path.exists(v):
            return v
        p = os.path.join(HERE, v)
        if os.path.exists(p):
            return p
        return v
    return ""
# 被剔除（非私募等）条目，供登记到案例库根目录「不收录清单.md」
EXCLUDED = []
AUDIT = []      # 中基协未出现"私募"字样的审计提示（不剔除，仅记录）
# P3 名册闸门：由 main() 加载 AMAC 管理人名册（在册+已注销）；None 时退化为纯文本判定
REGISTRY = None

# ----------------------------------------------------------------- 违规类型分类（六分法）：统一走 classify_filter.classify
# 分类口径 / 关键词 / 加权 / 置信度阈值全部在 classify_filter.py 单一维护，避免与 verify 漂移。


# ----------------------------------------------------------------- 实体抽取（机构 vs 人员）
# 机构名必须以这些后缀收尾（防止 "公司向基金业协会报送" 这类短语冒充机构名）
INST_SUFFIX = re.compile(r"(?:有限公司|有限责任公司|股份公司|股份有限公司|合伙企业|有限合伙|"
                         r"管理中心|事务所|公司|企业|集团)[）)]?$")

INST_KW = ["有限公司", "有限责任公司", "股份公司", "股份有限公司", "合伙企业", "私募", "基金",
           "资产", "投资", "证券", "资本", "财富", "企业", "集团", "管理中心", "公司"]

PERSON_TITLE = ["法定代表人", "董事长", "总经理", "总裁", "合规风控负责人", "风控负责人", "执行董事",
                "董事", "监事", "副总经理", "投资总监", "基金经理", "实际控制人", "总经理助理",
                "投资经理", "负责人", "董事长兼总经理", "执行董事兼总经理", "法定代表人和董事长",
                "总经理兼法定代表人"]


# 人名合法性过滤：title/当事人切词会把虚词、职务、机构词误当人名
# （实测误抓："兼董事"、"及新享"、"根据"），必须剔除，否则 D 列「被处罚机构/人」被污染。
_BAD_NAME = re.compile(r"^(?:根据|依据|及|兼|以及|同时|另外|上述|其中|本人|该公司|公司|当事人|"
                       r"以下|如下|期间|未|现|原|并|和|与|等|其|该|对|关于|决定|采取|予以|"
                       r"时任|现任|曾任|原任|登记为|现登记为)")
# 2026-10-27 补全：重建一年库时 D 列混进了「职能」「职责」「任职」「变更」「已于」等
# 非人名 token（它们既不在 _BAD_NAME 前缀里、也不在旧 _TITLE_WORD 里）。
# 补充一批「绝不可能出现在自然人姓名中」的语素，命中即否掉。
_TITLE_WORD = re.compile(r"(?:董事|监事|经理|总裁|负责人|高管|员工|人员|职务|法定代表人|"
                         r"实际控制|股东|协会|公司|基金|投资|资产|管理|部门|岗位|"
                         r"职能|职责|任职|变更|登记|备案|披露|报送|规定|办法|条例|细则|指引|"
                         r"监管|处罚|措施|决定|以下|上述|相关|部分|情况|问题|行为|事项|"
                         r"从业|资格|信息|报告|合同|投资者|募集|宣传|内控|风控|合规|经营|"
                         r"业务|产品|资金|财产|账户|材料|资料|文件|申请|复核|申辩|整改)")
# 切词残留的句首/句尾虚词：剥掉后再判是否像人名
# （实测："为濮萍"→"濮萍"、"鲍娇燕等"→"鲍娇燕"、"已于"→"于"(过短被否)）
_EDGE_FUNC = re.compile(r"^[为系即已及并或兼且又也本该此其]+|[等兼及]+$")
# 2026-10-29 新增：正文「合规风控负责人缺位」被 `负责人`+2字 模式切成"缺位"当人名，
# 污染台账 D 列与文件名（实测 景行天成 / 中诚新川 两案）。
# 这类「任职状态/占位」词绝不可能出现在自然人姓名中，整词命中即否掉。
_NON_PERSON = re.compile(r"^(?:缺位|空缺|空岗|不明|不明确|未知|无人|未设|未配|兼任|在任|离任|"
                         r"新任|到岗|在岗|专职|兼职|相关|有关|当事|所在|所属|任职|履职)$")


def clean_person(t):
    """剥掉首尾虚词后返回像人名的 token；不像人名则返回 None。

    调用点必须**写入本函数返回的值**，不能写原始 token——
    否则「为濮萍」「鲍娇燕等」「已于」这类带残留虚词的串会原样进 D 列。
    """
    s = _EDGE_FUNC.sub("", (t or "").strip())
    if not s or not (2 <= len(s) <= 4):
        return None
    if not re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", s):
        return None
    if _BAD_NAME.match(s) or _TITLE_WORD.search(s):
        return None
    if _NON_PERSON.match(s):
        return None
    # 2026-10-29：碎片词表统一从 subject_clean 取（唯一事实源）。
    # 旧版只拦「缺位/兼任/相关…」，漏掉「至今/职位/签章/离职/以前/之后/挂名/之一」，
    # 实测把 9 行正文碎片写进台账 D 列并拼出「职位-上海通赛投资有限公司.pdf」。
    if _SC.is_noise_token(s):
        return None
    if looks_inst(s):
        return None
    return s


def ok_person(t):
    """判断切出的 token 是否像自然人姓名（兼容旧调用）。"""
    return clean_person(t) is not None


def looks_inst(name):
    if not name or len(name) < 4:
        return False
    # 剔除被匿名化的机构（某资管公司 等）
    if re.fullmatch(r"某[一-龥]*?(?:公司|企业|合伙|资管|基金|资产|投资|证券|资本|财富|集团)", name):
        return False
    return any(k in name for k in INST_KW)


def normalize_inst(name):
    """把含职务/简称括号的实体归一为【机构全称】：
    例 "现任深圳信钰基金管理有限公司（以下简称深圳信钰）合规风控负责人" → "深圳信钰基金管理有限公司"。
    """
    if not name:
        return ""
    # 去掉任职状态前缀：现任/时任/曾任/原任/登记为/现登记为/原
    name = re.sub(r"^(?:现登记为|登记为|现任|时任|曾任|原任|原)", "", name)
    # 2026-10-29：正文里混进的「第N号」「content.shtml」「https://…」等非实体前缀，
    # 会整段带进主体名（实测 "第4号国厚资产管理股份有限公司" / "shtml景行天成（宁波）…"）。
    name = re.sub(r"^第?\s*\d+\s*号", "", name)
    name = re.sub(r"^[^\u4e00-\u9fa5（(]+", "", name)
    # 机构名可含全角/半角括号与数字字母（如「中投汇恒（北京）投资基金管理有限公司」），
    # 只匹配 [一-龥] 会被括号截断、只剩尾部「投资基金管理有限公司」。
    m = re.search(r"([\u4e00-\u9fa5（）()A-Za-z0-9]{2,40}?(?:有限公司|有限责任公司|股份公司|股份有限公司|合伙企业|管理中心|事务所|研究院|集团公司))", name)
    return m.group(1) if m else name


_TITLE_END = ("采取", "责令改正", "出具警示函", "行政处罚", "监管措施", "市场禁入",
              "公开谴责", "通报批评", "纪律处分", "注销", "决定", "予以", "实施",
              # 2026-10-29 补：交易所文书标题「关于对X的监管函 / 的监管警示 / 的决定」
              # 「的」不在旧终止集里 → 深交所债券业务中心监管函的主体**完全抽不到**，
              # 台账 D 列落成「不明确」、PDF 落成「不明确.pdf」（国厚资产两案实测）。
              "的监管函", "的监管警示", "的监管措施", "的通报批评", "的纪律处分决定", "的决定")


def title_entities(title):
    """从标题抽实体：'关于对A、B…' 或 '纪律处分决定书（NAME）' / '〔NAME〕'。"""
    m = re.search(r"关于对(.+?)(" + "|".join(_TITLE_END) + r")", title or "")
    if m:
        chunk = m.group(1)
        # 2026-10-29：切分只认「、，,」——按「和/及」切会把「深圳前海和凯创业投资
        # 有限公司」劈成「深圳前海」+「凯创业投资有限公司」，前者又被当成人名，
        # 于是台账主体退化成「凯创业投资有限公司；深圳前海」（行556 实测）。
        toks = [p.strip(" 、，,。（）()") for p in re.split(r"[、，,]+", chunk)
                if p.strip(" 、，,。（）()")]
        # 主体以机构后缀收尾：终止符未命中时 chunk 会拖出「…账户实施限制交易」等尾巴，
        # 逐个 token 截到机构后缀为止（2026-10-29）。
        out = []
        for t in toks:
            mi = re.search(r"^(.*?(?:有限公司|有限责任公司|股份公司|股份有限公司|合伙企业"
                           r"|有限合伙|管理中心|事务所|集团公司))", t)
            out.append(mi.group(1) if mi else t)
        return out
    m = re.search(r"纪律处分(?:复核)?决定书[（(【\[]\s*(.+?)\s*[）)】\]]", title or "")
    if m:
        return [m.group(1).strip()]
    return []


def entity_from_text(text, title=""):
    """返回 (manager_inst, persons_list)。

    注意：证监局内联正文/AMAC OCR 都会在字间插空格（"上 创 普 盛"），而 extract_core 抽出的
    处罚措施已去空格。为让 "manager in measures_text" 判定成立，实体抽取统一在【去空格】文本上做。
    """
    flat = re.sub(r"\s+", "", text or "")
    # 「（以下简称XX）」「（以下简称“XX”）」会把机构名从中间截断，导致抽取只拿到简称
    # （如 芜湖启晨 / 池州金桥，长度与后缀都不合规被丢弃）→ 先整段摘掉再匹配。
    flat = re.sub(r"[（(]\s*以下简称[^）)]*[）)]", "", flat)
    manager = ""
    persons = []
    # 1) 标题实体
    for e in title_entities(title):
        if looks_inst(e):
            if not manager:
                manager = e
        else:
            ce = clean_person(e)
            if ce and ce not in persons:
                persons.append(ce)
    # 2) 当事人 条款
    m = re.search(r"当事人[:：](.+?)(?:。|依据|根据|我局|本局|本所|决定)", flat)
    if m:
        chunk = m.group(1)
        toks = re.split(r"[，,、；;。]+", chunk)
        for t in toks:
            t = t.strip()
            if not t:
                continue
            if looks_inst(t) and not manager:
                manager = t
            else:
                ct = clean_person(t)
                if ct and ct not in persons:
                    persons.append(ct)
    # 2.5) 抬头送达行「机构、人名：经…」——交易所监管函的当事人只在这里出现一次。
    #      例：国厚资产管理股份有限公司、李厚文：经中国证监会安徽监管局…（2026-10-29 实测）
    #      说明：`负责人([\u4e00-\u9fa5]{2,3})(?![\u4e00-\u9fa5])` 因姓名后紧跟动词
    #      （「李厚文未能恪尽职守」）而回溯失败，抓不到人名，故单列此条。
    if not persons:
        hm = re.search(r"([一-龥（）()A-Za-z0-9、]{2,80})[：:](?:经|根据)"
                       r"(?:中国证券监督管理委员会|中国证监会|本所|我会|我局|查明|查)",
                       flat)
        if hm:
            for t in hm.group(1).split("、"):
                t = t.strip()
                if not t:
                    continue
                if looks_inst(t):
                    if not manager:
                        manager = normalize_inst(t)
                else:
                    cp = clean_person(t)
                    if cp and cp not in persons:
                        persons.append(cp)
    # 3) 违规主体：按优先级依次尝试（纯个人案中机构名只出现在正文）
    if not manager:
        pats = [
            # 2026-10-29 补（交易所文书专用，放最前）：
            # ① 标题/抬头式「关于对X的监管函」（深交所债券业务中心监管函主体全靠这条）
            # ② 债券发行人表述「X 作为公司债券…的发行人」
            # ③ 正文抬头送达行「X：你公司作为…」
            r"关于对([一-龥（）()]{2,40}?)(?:的(?:监管函|监管警示|决定|通报批评|纪律处分)|"
            r"(?:采取|予以|出具))",
            r"([一-龥（）()]{2,40}?)(?:作为公司债券|作为公司债券的发行人)",
            # 「你作为苏州中来光伏新材股份有限公司时任董事长」——「时任」在机构名之后，
            # 旧 pattern 靠「时任 X」抓不到（行82 拆分时实测）。
            r"你作为([一-龥（）()]{2,40}?(?:有限公司|有限责任公司|股份公司|股份有限公司"
            r"|合伙企业|有限合伙|管理中心|事务所))",
            # 「王军华为苏交科集团股份有限公司持股5%以上大股东」（纯个人案，机构只出现一次）
            r"为([一-龥（）()]{2,40}?(?:有限公司|有限责任公司|股份公司|股份有限公司|合伙企业"
            r"|有限合伙|管理中心|事务所))(?:持股|的股东|股东)",
            r"([一-龥（）()]{2,40}?(?:有限公司|有限责任公司|股份公司|股份有限公司|合伙企业"
            r"|有限合伙|管理中心|事务所))[：:]\s*(?:你公司|你们|你)",
            r"([一-龥]{2,40}?)[（(]统一社会信用代码",               # X（统一社会信用代码…）
            # 「我局对X（以下简称公司）开展…检查中发现」——纯个人案里机构只在这里出现一次，
            # 且后面不再重复，漏了这条会把「你作为公司向基金业协会报送的…」里的
            # 「公司向基金业协会报送」误当机构名（2026-10-27 孙勇案实测）。
            # 注意：终止集合里不能放"私募"（机构名里常带"私募"二字，会造成过早截断）
            r"我局(?:在)?对([一-龥（）()]{2,40}?)(?:（以下简称|开展|进行|检查)",
            r"([一-龥（）()]{2,40}?)(?:在开展|在从事|存在以下|存在如下|存在问题|存在的主要问题)",
            r"你(?:在)?([一-龥（）()]{2,40}?)(?:任职|工作|担任)",       # 你 在 X 任职期间
            # 中基协复核决定书：「申请人：徐铭琦…现任台州沃源私募基金管理有限公司基金经理」
            r"(?:现任|时任|曾任|任职于|就职于)([一-龥（）()]{2,40}?(?:有限公司|有限责任公司|"
            r"股份公司|股份有限公司|合伙企业|管理中心|事务所))",
            r"你作为([一-龥（）()]{2,40}?)(?:的|公司)",                # 你作为 X 的…
            # 「北京博晟投资管理有限公司所管理的5只私募基金产品…」（纯个人案，机构只在此出现）
            r"([一-龥（）()]{2,40}?(?:有限公司|有限责任公司|股份公司|股份有限公司|合伙企业"
            r"|有限合伙|管理中心|事务所))所(?:管理|发行|承销|销售)",
        ]
        for pat in pats:
            mm = re.search(pat, flat)
            g = mm.group(1).strip() if mm else ""
            if not g:
                continue
            # 去掉前导虚词（"我局发现""我局对"等）、「（以下称/以下简称 X）」简称括号，
            # 再取末尾以机构后缀结尾的一段 → 机构全称（2026-10-27 刘晓雪案实测：
            # "我局发现哈尔滨经云创业投资有限公司（以下称公司）" → "哈尔滨经云创业投资有限公司"）
            g = re.sub(r'^(?:我局发现|我局在|我局对|我局)', '', g)
            # 2026-10-29：交易所文书抬头常带「深圳证券交易所关于对」等前缀，
            # 不去掉会把整串「深圳证券交易所关于对X股份有限公司」当机构名。
            g = re.sub(r'^(?:(?:关于对|深圳证券交易所|上海证券交易所|北京证券交易所|'
                       r'全国中小企业股份转让系统))+', '', g)
            g = re.sub(r'[（(](?:以下称|以下简称)[^）)]*[）)]', '', g)
            m_tail = re.search(
                r'([一-龥（）()A-Za-z0-9]{2,60}(?:有限公司|有限责任公司|股份公司|'
                r'股份有限公司|合伙企业|有限合伙|管理中心|事务所|集团公司|企业))'
                r'(?:\s*[）)])?\s*$', g)
            if m_tail:
                g = m_tail.group(1)
            # 候选必须以机构后缀收尾，否则 "公司向基金业协会报送" 这类短语会冒充机构名
            if g and looks_inst(g) and INST_SUFFIX.search(g):
                manager = g
                break
    # 4) 被送达人 "NAME：经查" 开头（机构）
    if not manager:
        mm = re.search(r"^([一-龥]{2,40}?)[:：]经查", flat)
        if mm and looks_inst(mm.group(1)):
            manager = mm.group(1)
    # 正文职务+姓名 补人名
    for t in PERSON_TITLE:
        for mm in re.finditer(t + r"([一-龥]{2,3})(?![一-龥])", flat):
            nm = clean_person(mm.group(1))
            if nm and nm not in persons:
                persons.append(nm)
    # 被送达人（个人）开头 "NAME：经查"
    if not persons:
        mm = re.search(r"^\s*([一-龥]{2,4})\s*[:：]\s*经查", text)
        if mm and not looks_inst(mm.group(1)):
            persons.append(mm.group(1))
    return normalize_inst(manager), _merge_persons(persons)


def _merge_persons(persons):
    """合并"同一人的不同写法"：OCR/异体字导致「逄颖颖」「逢颖颖」被当成两个人。

    判定：等长且仅 1 个字不同 → 视为同一人，保留先出现的写法。
    """
    out = []
    for p in persons:
        dup = next((q for q in out if len(q) == len(p)
                    and sum(a != b for a, b in zip(q, p)) <= 1), None)
        if dup is None:
            out.append(p)
    return out


def split_measures(measures_text, manager, persons):
    if not measures_text or measures_text.startswith("（"):
        return []
    clauses = [c.strip(" ；;。") for c in re.split(r"[；;。]", measures_text) if c.strip(" ；;。")]
    mgr_bucket, person_bucket = [], {p: [] for p in persons}
    for c in clauses:
        hp = next((p for p in persons if p in c), None)
        if hp:
            person_bucket[hp].append(c)
        elif manager and manager in c:
            mgr_bucket.append(c)
        else:
            mgr_bucket.append(c)
    out = []
    mgr_mentioned = bool(manager) and manager in measures_text
    if mgr_bucket and mgr_mentioned:
        out.append("管理人：" + "；".join(mgr_bucket))
    for p in persons:
        if person_bucket[p]:
            out.append("高管（%s）：%s" % (p, "；".join(person_bucket[p])))
    # 机构未被提及（纯个人处罚）：把条款归到高管
    if not mgr_mentioned and persons and not any(person_bucket[p] for p in persons):
        out.append("高管（%s）：%s" % (persons[0], "；".join(mgr_bucket)))
    if not out:
        out.append("管理人：" + measures_text)
    return out


# ----------------------------------------------------------------- 私募相关性过滤
# 本库是【私募】监管处罚案例库。证监局平台会同时下发上市公司（股份公司）、券商营业部、
# 审计机构、内幕交易个人、证券投资顾问公司等与私募无关的文书，必须剔除，否则会误收
# （2026-10-27 用户指出：本版误收大量股份公司案例）。判定：正文需出现私募/中基协等关键词。
# 【强阳性】私募专指词：真私募监管文书必然点名"私募…"或援引《私募投资基金监督管理条例/暂行办法》。
# 只要出现即判私募，命中率极高（189 条证监局文书里 52 条为真私募，全部含"私募"）。
PF_STRONG = ["私募", "基金管理人登记", "管理人登记"]
# 【弱阳性】仅出现协会名，不足以定性：上市公司公告里也可能提到"中国证券投资基金业协会"
# （2026-10-27 实测：雄安新动力科技股份有限公司内幕交易/信披案只因正文出现"基金业协会"
#   被旧规则误收，实质引用《证券法》《上市公司信息披露管理办法》，与私募监管无关）。
PF_WEAK = ["中基协", "中国证券投资基金业协会", "基金业协会"]
# 机构名称像私募机构（用于弱阳性时的二次确认）
_PF_NAME = re.compile(r"(私募|基金管理|投资管理|资产管理|股权投资|创业投资|投资基金|资本管理|财富管理)")
# 明显不是私募机构的名称特征 → 唯一事实源在 classify_filter._NOT_PF_NAME
# （2026-10-29 收敛：本文件不再保留独立副本，避免与 classify_filter 两套不一致；
#  build_pipeline 已 `from classify_filter import is_private_fund`，剔除判定统一走 classify_filter）
from classify_filter import _NOT_PF_NAME  # noqa: E402  （与 classify_filter._NOT_PF_NAME 同一对象，禁止本地副本）


# is_private_fund：见 classify_filter（已在文件头部统一 import），本文件不再重复定义。


# ----------------------------------------------------------------- 被处罚机构/人（按实际情况填）
def subject_of(mgr, persons, measures_text):
    """被处罚机构/人：机构与人员可能只有其一存在，按决定书实际处罚对象填写。

    分隔符约定（2026-10-27 定稿）：
      * 机构与人员之间用「；」——分号前是机构、分号后是被处罚人员；
      * 多个人员之间用「、」。

    - 机构与关联人员同案被罚 → "机构名；张三、李四"（机构在前，人员在后）
    - 仅人员被罚但能从决定书取得所属机构 → "机构名；张三"
      （机构名作归属线索保留；必须同时保留人员名，否则同一机构的多份人员决定书
       在判重键「机构来源 + 被处罚机构/人 + 处罚时间」上会撞车，如上海隐上的张波/韩方圆）
    - 机构不可得 → "张三"
    """
    # 2026-10-29：组装前统一过一遍碎片词表（唯一入口）。
    # ① 剔 NOISE_TOKENS；② 剔「X+虚词」型重复碎片（「杜胜于」与「杜胜」并存 → 留杜胜）。
    # 放在这里而非各调用点：subject_of 是 D 列的唯一出口，漏一处就等于没修。
    persons = _SC.drop_noise_tokens(persons or [], keep={mgr} if mgr else None)
    if mgr:
        if persons:
            return mgr + "；" + "、".join(persons)
        return mgr
    return "、".join(persons) if persons else "不明确"


def filename_of(mgr, persons, measures_text):
    """文件名：能自解释——机构案用机构名；仅人员案用"姓名-所属机构"保留归属线索。

    2026-10-29 硬规则：**文件名里绝不允许出现「不明确/待补/缺位」等占位词**。
    主体确实抽不出来时退回「待核-<8位内容哈希>」，既保证唯一、可追溯，又让
    verify 的「占位命名」断言一眼能抓到（此前国厚资产两案落成 不明确.pdf / 不明确（2）.pdf）。
    """
    mgr_hit = bool(mgr) and mgr in (measures_text or "")
    if mgr_hit:
        return mgr + ".pdf"
    if persons:
        return persons[0] + ("-" + mgr if mgr else "") + ".pdf"
    if mgr:
        return mgr + ".pdf"
    import hashlib
    h = hashlib.sha256((measures_text or "").encode("utf-8")).hexdigest()[:8]
    print("  [WARN] 主体抽取失败，文件名退回 待核-%s.pdf（需人工回填）" % h)
    return "待核-%s.pdf" % h


# ----------------------------------------------------------------- 主流程
def build_csrc():
    recs = []
    for it in _load_items(CSRC_JSON):
        text = tidy(it.get("text") or "")
        _mgr0, _p0 = entity_from_text(text, it.get("title", ""))
        ok, why = decide(text, _mgr0 or "", it.get("title", ""), REGISTRY)
        if not ok:
            EXCLUDED.append({"source": "证监局", "title": it.get("title", ""),
                             "url": it.get("url", ""), "reason": "非私募（%s）" % why})
            continue
        src = extract_core.extract_with_contract(text, "csrc")
        manager, persons = _mgr0, _p0
        mgr_hit = bool(manager) and manager in src["measures"]
        meas = split_measures(src["measures"], manager, persons)
        if not meas:
            meas = ["管理人：不明确"]
        category = classify(text, "证监局")
        recs.append({
            "source_dir": "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）",
            "category": category,
            "subject": subject_of(manager, persons, src["measures"]),
            "manager": manager or "不明确",
            "persons": persons,
            "measures_text": src["measures"],
            "is_inst": mgr_hit,
            "date_decided": it.get("date_fwrq") or it.get("date_published") or "",
            "publish_date": it.get("date_published") or "待补",
            "measures": meas,
            "grounds": [g for g in src["grounds"].split("；") if g] or [src["grounds"]],
            "evidence": [e for e in src["evidence"].split("；") if e] or [src["evidence"]],
            "legal_basis": [l for l in src["legal_basis"].split("；") if l] or [src["legal_basis"]],
            "defense": src["defense"] or "无",
            "files": [{"src": _resolve_src(it),
                       "name": filename_of(manager, persons, src["measures"]),
                       "person": persons[0] if persons else ""}],
            "urls": [it.get("url", "")],
            "typical": False,
            "source": "证监局",
        })
    return recs


# OCR 会在「年」前插句点（"2026.年8月25日"）、在「日」前插全角斜杠（"7月1／日"），都要容忍
_DATE_PAT = re.compile(r"(\d{4})[.．。､]?年(\d{1,2})月(\d{1,2})\s*[／/|丨]?\s*日")


def extract_sign_date(text):
    """抽决定书落款日期：优先取【文末】最后一个可解析年月日（落款在文末）。

    坑①：中基协 OCR 常在日与「日」之间插全角斜杠（"2026年7月1／日"），严格正则匹配不到，
    会误取正文里更早的日期（如 上海隐上 误取成 2023-09-08）。
    坑②（2026-10-30，严重）：OCR 不只是插斜杠——实测中基协扫描件落款常被识别成
    「2026,年5月。26日」「2025°年11月25日」「2025年12月-24日」，本地 `_DATE_PAT`
    要求「年—月—日」紧邻数字，这类**月/日之前夹 OCR 杂字符**的落款一个都匹配不到 →
    回落 `or pub`（挂网日）→ 处罚时间被写成发布日。6 条中基协案例（券期资本/珠海有容/
    广东华迪/横琴均成/北京诺壹/前海坚石）台账处罚时间全错成挂网日，真实落款早 3~8 个月，
    verify B9 命中。真因是「同一概念两套日期正则」——sign_date.py 早已有年/月/日之间
    允许 0~4 个杂字符的容错口径，这里却另立一套。现改为**复用 sign_date.parse_dates**
    （唯一事实源），并传 max_date=今天 过滤「整改期限」这类未来日期。
    """
    try:
        import sign_date as _SD
        cand = None
        for seg in ((text or "")[-600:], text or ""):
            ds = _SD.parse_dates(seg, max_date=_SD.today_iso())
            if ds:
                cand = ds[-1]
                break
        if cand:
            return cand
    except Exception:
        pass
    # 兜底：本地宽松正则（sign_date 不可用时）
    flat = re.sub(r"\s+", "", text or "")
    ds = _DATE_PAT.findall(flat)
    if not ds:
        return ""
    dt = _DATE_PAT.findall(flat[-300:])
    y, m, d = (dt or ds)[-1]
    try:
        return "%04d-%02d-%02d" % (int(y), int(m), int(d))
    except Exception:
        return ""


# ------------------------------------------------- 写入期断言：处罚时间不许静默回落挂网日
# 2026-10-30 用户裁定（原话）："如果处理时确保没问题那其实没必要去验证复核的，
# 除了某些必要的在处理过程没有确保"。
# ⇒ 把「事后全量 OCR 复扫」改成「写入期硬断言」，避免同一份文件跑两遍算法：
#   落款抽不到时**不再** `or pub` 静默回落（那正是 B9 的成因），改按来源分流：
#     · 中基协 / 交易所：决定书按惯例必有落款 → 抽不到就是抽取失败 → 写「待核」，不许编日期；
#     · 证监局（官网内联正文）：确有 38 条文书正文只写检查/整改日期、无落款 → 允许取平台发文日，
#       但必须**留痕**（AUDIT 记 date_source=publish），且进 verify 观察项。
DATE_HOLD = "待核"
_SIGN_REQUIRED = ("中基协", "交易所")


def date_or_hold(text, pub, source, note=""):
    """返回 (处罚时间, 取数来源)。取数来源 ∈ {sign, publish, hold}。

    sign    = 正文档末落款（唯一正确口径）
    publish = 平台发文/挂网日（仅证监局无落款文书允许，且必须留痕）
    hold    = 抽不到落款（中基协/交易所），值写「待核」，交人工，绝不拿挂网日顶替
    """
    d = extract_sign_date(text)
    if d:
        return d, "sign"
    pub = (pub or "").strip() or "待补"
    if source in _SIGN_REQUIRED:
        AUDIT.append({"source": source, "note": "落款抽取失败，处罚时间置「待核」（%s）" % note,
                      "publish": pub, "reason": "date_hold"})
        return DATE_HOLD, "hold"
    AUDIT.append({"source": source, "note": "正文无落款，取平台发文日期（%s）" % note,
                  "publish": pub, "reason": "date_publish"})
    return pub, "publish"


def csrc_filename(mgr, persons):
    """证监局决定书文件名：管理人全称.pdf；管理人未知时用 姓名-不明确.pdf 区分。"""
    if mgr:
        return mgr + ".pdf"
    if persons:
        return persons[0] + "-不明确.pdf"
    return "不明确.pdf"


def ensure_unique_names(recs):
    """保证全库文件名唯一（同一管理人多份决定书、多个'不明确'会重名）。重名加 （2）（3）后缀。"""
    used = {}
    for r in recs:
        for f in r["files"]:
            base = f["name"]
            if base in used:
                used[base] += 1
                stem, ext = os.path.splitext(base)
                f["name"] = "%s（%d）%s" % (stem, used[base], ext)
            else:
                used[base] = 1


def _amac_filename(it, mgr, persons):
    if it["column"] == "scfjg":
        return (mgr or "不明确") + ".pdf"
    nm = persons[0] if persons else "不明确"
    return nm + "-" + (mgr or "不明确") + ".pdf"


def amac_text(it):
    """中基协 item 的 text 字段是 OCR 文本文件路径（相对 tmp_pilot），需读文件。"""
    p = it.get("text") or ""
    if re.search(r"\.txt$", p) or "\\" in p or "/" in p:
        fp = p if os.path.isabs(p) else os.path.join(HERE, p)
        if os.path.exists(fp):
            return open(fp, encoding="utf-8", errors="ignore").read()
        return ""
    return p


def build_amac():
    """中基协：每个决定书（机构 OR 人员）独立成行，按实际情况收集，不强行合并配对。

    用户 2026-10-26 明确：机构与对应人员的处罚决定往往不同时作出，可能先有其一。
    故技能按实际收集到的决定书逐条建行；若同批次机构+人员都在窗口内则各自成行、
    共享管理人名称，互不冒充；确有把握配对时可另行合并，但默认不强制。
    """
    recs = []
    # 中基协同一份决定书会被「受处分机构」「受处分人员」两栏【同时收录】且 URL 完全相同
    # （2026-10-27 实测：41 条里 7 条跨栏重复 → 台账 A3 判重报警）。必须按 URL 去重。
    seen_url = set()
    for it in _load_items(AMAC_JSON):
        u = it.get("url", "")
        if u in seen_url:
            continue
        seen_url.add(u)
        text = tidy(amac_text(it))
        title = it.get("title", "")
        # 中基协纪律处分栏目本身即私募自律监管范畴（不存在非私募文书），过滤仅作审计提示，
        # 不据此剔除——否则"纪律处分决定书（张三）"这类纯人名标题会被名称特征误杀。
        ok, why = decide(text, title, title=title, registry=REGISTRY)
        if not ok:
            AUDIT.append({"source": "中基协", "title": title, "url": it.get("url", ""),
                          "note": "未出现私募字样（%s），因来源为协会自律处分仍收录" % why})
        mgr, persons = entity_from_text(text, title)
        src = extract_core.extract_with_contract(text, "amac")
        meas = split_measures(src["measures"], mgr, persons)
        if not meas:
            if it["column"] == "scfry":
                nm = persons[0] if persons else "不明确"
                meas = ["高管（%s）：不明确" % nm]
            else:
                meas = ["管理人：不明确"]
        mgr_hit = (it["column"] == "scfjg") or (bool(mgr) and mgr in src["measures"])
        pub = (it.get("publish") or "")[:10] or "待补"
        # 写入期断言：落款抽不到时不再 `or pub` 静默回落（B9 成因），一律走 date_or_hold 分流。
        decided, date_src = date_or_hold(text, pub, "中基协", title)
        recs.append({
            "source_dir": "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）",
            "category": classify(text, "中基协"),
            "subject": subject_of(mgr, persons, src["measures"]),
            "manager": mgr or "不明确",
            "persons": persons,
            "measures_text": src["measures"],
            "is_inst": mgr_hit,
            "date_decided": decided,
            "date_source": date_src,
            "publish_date": pub,
            "measures": meas,
            "grounds": [g for g in src["grounds"].split("；") if g] or [src["grounds"]],
            "evidence": [e for e in src["evidence"].split("；") if e] or [src["evidence"]],
            "legal_basis": [l for l in src["legal_basis"].split("；") if l] or [src["legal_basis"]],
            "defense": src["defense"] or "无",
            "files": [{"src": _resolve_src(it),
                       "name": filename_of(mgr, persons, src["measures"]),
                       "person": persons[0] if persons else ""}],
            "urls": [it["url"]],
            "typical": False,
            "source": "中基协",
        })
    return recs


# ----------------------------------------------------------------- 按机构合并（同一机构的多份决定书 → 一行）
# 用户 2026-10-27 明确要求：机构与关联人员的处罚若同时存在，必须落在【同一行】的同一格子里
# （机构在前、人员在后），而不是分成两行。本步把「同一机构」的多份决定书（机构级 + 人员级）
# 聚合成一条台账行：被处罚机构/人 = "机构名；人员1、人员2"，处罚措施 = 机构措施在前、人员措施随后，
# 案例文件 = 机构 PDF + 各人员 PDF（命名均体现机构）。机构不可得时退化为纯人员行。
_INST_SUFFIX = re.compile(r"(公司|合伙|中心|企业)$")


def _rec_source_text(r):
    """按 url 回查源正文（合并时用于把「机构名不可得的个人案」挂回其所属机构）。"""
    u = (r.get("urls") or [""])[0]
    it = next((x for x in _load_items(CSRC_JSON) if x.get("url") == u), None)
    if it:
        return tidy(it.get("text") or "")
    it2 = next((x for x in _load_items(AMAC_JSON) if x.get("url") == u), None)
    if it2:
        return tidy(amac_text(it2))
    return ""


def merge_by_institution(recs):
    """按「同一机构」合并机构级 + 人员级文书为一行（用户 2026-10-27 硬要求）。

    坑（2026-10-27）：分组键必须带【来源】。同名机构可能同时被证监局与交易所点名，
    只按机构名分组会把两个来源的文书并成一行、source_dir 互串 → 台账一行横跨两个大类。
    """
    groups, known, solo = {}, [], []
    for r in recs:
        m = r.get("manager", "")
        if m and m != "不明确" and _INST_SUFFIX.search(m):
            key = (r.get("source", ""), m)      # 来源 + 机构名 双键
            if key not in groups:
                groups[key] = []
                known.append(key)
            groups[key].append(r)
        else:
            solo.append(r)
    # 把「机构名不可得的个人案」按其正文是否精确命中某已知机构名，挂回该机构组
    attached = set()
    for r in solo:
        t = re.sub(r"\s+", "", _rec_source_text(r))
        hit = next((k for k in known if k[1] in t and k[0] == r.get("source", "")), None)
        if hit:
            groups[hit].append(r)
            attached.add(id(r))
    out = [_merge_group(k[1], groups[k]) for k in known]
    # 真正机构不可得的个人案 → 各自成行
    for r in solo:
        if id(r) not in attached:
            out.append(r)
    return out


def _merge_group(mgr, rs):
    persons, files, urls, meas_parts = [], [], [], []
    inst_rec = None
    for r in rs:
        for p in r.get("persons") or []:
            if p not in persons:
                persons.append(p)
        for f in r.get("files") or []:
            files.append(f)
        urls.extend(r.get("urls") or [])
        meas_parts.extend(r.get("measures") or [])
        if r.get("is_inst") and inst_rec is None:
            inst_rec = r
    # 机构级文书排前（保证「机构措施在前」）
    rs_sorted = sorted(rs, key=lambda x: 0 if x.get("is_inst") else 1)
    meas_parts = []
    for r in rs_sorted:
        for m in r.get("measures") or []:
            if m not in meas_parts:
                meas_parts.append(m)
    core = inst_rec or rs[0]
    subject = mgr + ("；" + "、".join(persons) if persons else "")
    # 文件名规范化：机构级文书 → 机构名.pdf；人员级文书 → 姓名-机构名.pdf（不含职务）
    suffix = "-" + mgr + ".pdf"
    for f in files:
        nm = f.get("name", "")
        if nm == mgr + ".pdf" or nm.endswith(suffix):
            continue
        person = f.get("person", "")
        base = re.sub(r"-不明确\.pdf$", "", nm)
        base = re.sub(r"\.pdf$", "", base)
        f["name"] = (person or base) + suffix
    return {
        "source_dir": core["source_dir"],
        "category": core["category"],
        "subject": subject,
        "manager": mgr,
        "persons": persons,
        "measures_text": core.get("measures_text", ""),
        "is_inst": True,
        "date_decided": core.get("date_decided", ""),
        "date_source": core.get("date_source", ""),
        "publish_date": core.get("publish_date", "待补"),
        "measures": meas_parts,
        "grounds": core.get("grounds", []),
        "evidence": core.get("evidence", []),
        "legal_basis": core.get("legal_basis", []),
        "defense": core.get("defense", "无"),
        "files": files,
        "urls": urls,
        "typical": core.get("typical", False),
        "source": core.get("source", "中基协"),
    }


EXCH_DIR = "03_交易所监管措施（沪深交易所、北交所，私募产品交易、异常交易纪律）"


def _exch_text(it):
    """交易所文书正文：优先已落盘的 PDF（深/北交所有本地缓存），回退正文开头 200 字。

    坑（2026-10-27）：上交所只给 URL 不给本地文件，正文不可得时若拿 body_head 去跑
    is_private_fund，200 字窗口太小 → 误判；故正文 <60 字直接判「不可得」剔除并登记，
    绝不靠残缺文本硬判。
    """
    fp = (it.get("file") or it.get("pdf") or "")
    if fp and not fp.startswith("http") and os.path.exists(fp):
        try:
            import fitz
            d = fitz.open(fp)
            t = "".join(pg.get_text() for pg in d)
            d.close()
            if t.strip():
                return t
        except Exception:
            pass
    return ""


def build_exch():
    """沪深北三所监管措施/纪律处分 → 03 大类。

    只有真正落到「私募」的才收（is_private_fund 三层判定）；交易所文书主体多为上市公司、
    券商营业部、异常交易个人，命中率天然极低，属正常现象而非抓取失败。
    """
    if not EXCH_JSON or not os.path.exists(EXCH_JSON):
        return []
    d = json.load(open(EXCH_JSON, encoding="utf-8"))
    items = []
    for k in ("sse", "szse", "bse"):
        for it in (d.get(k) or []):
            if not isinstance(it, dict):
                continue
            it = dict(it)
            it["_exch"] = k.upper()
            items.append(it)
    recs, seen = [], set()
    for it in items:
        key = (it.get("url") or it.get("pdf") or it.get("file") or "",
               it.get("title") or it.get("subj") or "")
        if key in seen:
            continue
        seen.add(key)
        subj = it.get("subj") or it.get("target") or it.get("title") or ""
        url = it.get("url") or it.get("pdf") or it.get("file") or ""
        # 交易所决定书正文：先下载 PDF（带节流+退避，复用缓存）再抽全文；
        # 上交所/北交所抓取阶段只留 URL，必须在 build 阶段兜底抽取，否则无正文可判私募。
        # 修复（2026-10-27 §13.6 完备性）：正文抽不到不能丢案——无文字层走 OCR 兜底，
        # 仍不可得但属「强命中」（抓取阶段已按主体名/标题判定为私募机构）的，仍以强命中收录，
        # 避免把真私募案（如海之源/君之健/汐泰等）漏掉。
        local = _exch_local(url, os.path.join(os.path.dirname(EXCH_JSON or "."), "dl_exch")) if url else ""
        raw_text = ""
        if local and os.path.exists(local):
            try:
                from fetch_amac import pdf_text as _ocr_pdf_text
                raw_text, _how = _ocr_pdf_text(
                    local, os.path.join(os.path.dirname(EXCH_JSON or "."), "ocr_exch"))
            except Exception:
                raw_text = ""
        text = tidy(raw_text) if len(raw_text.strip()) >= 60 else tidy(_exch_text(it))
        ok, why = decide(text, subj, it.get("title") or subj, REGISTRY)
        if (not ok) and it.get("hit_kind") == "strong":
            # 交易所强命中兜底（2026-10-30 加固）：仅当主体【确为私募管理人】才强制收录，
            # 否则仍交由 decide 按非私募剔除——防止爬虫把「网络科技 / 企业管理咨询 / 上市公司重整」
            # 等非私募主体误标 strong 而绕过 decide 的剔除（曜瞿如 / 问美 / 福石 即此漏网）。
            # 「确为私募管理人」= 名册（在册+已注销）命中，或名称/标题明确含私募标识。
            _strong_ok = (
                (REGISTRY.has_data() and bool(REGISTRY.is_registered_manager(subj)))
                or re.search(r"私募|基金管理|证券投资基金", subj + (it.get("title") or ""))
            )
            if _strong_ok:
                ok, why = True, "交易所强命中（主体名即私募机构，正文不可得仍收录）"
        if not ok:
            EXCLUDED.append({"source": "交易所", "title": (it.get("title") or subj)[:80],
                             "url": it.get("url") or it.get("pdf") or "",
                             "reason": "非私募（%s）" % why})
            continue
        mgr, persons = entity_from_text(text, it.get("title") or subj)
        src = extract_core.extract_with_contract(text, "csrc")
        meas = split_measures(src["measures"], mgr, persons)
        if not meas:
            meas = ["管理人：不明确"]
        date = (it.get("date") or "")[:10]
        # 注：url / local 已在上方判定私募时计算并复用（避免二次下载）
        recs.append({
            "source_dir": EXCH_DIR,
            "category": classify(text, "交易所"),
            "subject": subject_of(mgr, persons, src["measures"]),
            "manager": mgr or "不明确",
            "persons": persons,
            "measures_text": src["measures"],
            "is_inst": bool(mgr) and mgr in src["measures"],
            "date_decided": date,
            "publish_date": date or "待补",
            "measures": meas,
            "grounds": [g for g in src["grounds"].split("；") if g] or [src["grounds"]],
            "evidence": [e for e in src["evidence"].split("；") if e] or [src["evidence"]],
            "legal_basis": [l for l in src["legal_basis"].split("；") if l] or [src["legal_basis"]],
            "defense": src["defense"] or "无",
            "files": [{"src": local or url,
                       "name": filename_of(mgr, persons, src["measures"]),
                       "person": persons[0] if persons else ""}],
            "urls": [url],
            "typical": False,
            "source": "交易所",
        })
    return recs


def _load_items(path):
    """读取抓取产物的 items；文件不存在返回空列表（单信源跑批时不该崩）。

    2026-10-27：`--no-amac` / `--no-exchange` 只抓证监局时，amac.json 根本不存在，
    旧实现直接 FileNotFoundError 中断整轮构建。
    """
    if not path or not os.path.exists(path):
        return []
    try:
        return json.load(open(path, encoding="utf-8")).get("items", [])
    except Exception as e:
        print("[WARN] 读取失败，按空处理：%s（%s）" % (path, e))
        return []


def _dump_json(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    json.dump(obj, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def main(argv=None):
    import argparse as _ap
    global CSRC_JSON, AMAC_JSON, EXCH_JSON, OUT, EXCLUDED_OUT, AUDIT_OUT, REGISTRY
    p = _ap.ArgumentParser(description="案例记录构建（筛选 → 分类 → 机构+人员合并 → 三核心列 → 命名）")
    p.add_argument("--csrc-json", default=CSRC_JSON)
    p.add_argument("--amac-json", default=AMAC_JSON)
    p.add_argument("--exch-json", default=EXCH_JSON)
    p.add_argument("--out", default=OUT)
    p.add_argument("--excluded-out", default=EXCLUDED_OUT)
    p.add_argument("--audit-out", default=AUDIT_OUT)
    p.add_argument("--no-exch", action="store_true", help="跳过交易所（03）分支")
    p.add_argument("--registry-dir", default=_AMAC_REG_DIR,
                   help="AMAC 管理人名册目录（默认技能下 amac_registry/）；缺失则退化为纯文本判定")
    p.add_argument("--no-registry", action="store_true",
                   help="关闭 P3 名册闸门（仅用 P1/P2 文本判定）")
    a = p.parse_args(argv)
    CSRC_JSON, AMAC_JSON, OUT = a.csrc_json, a.amac_json, a.out
    EXCH_JSON = "" if a.no_exch else a.exch_json
    EXCLUDED_OUT, AUDIT_OUT = a.excluded_out, a.audit_out
    # P3 名册闸门：加载 AMAC 管理人名册（在册+已注销）；目录不存在/无数据时 REGISTRY 保持 None，
    # 退化为纯文本判定（保证可移植首跑不中断；名册须随库更新重下）。
    if not a.no_registry:
        try:
            REGISTRY = AmacRegistry(os.path.abspath(a.registry_dir))
            print("[P3 名册] 已加载：", REGISTRY.summary() if REGISTRY.has_data()
                  else "目录无数据，退化为纯文本判定")
        except Exception as e:
            print("[P3 名册] 加载失败，退化为纯文本判定：%s" % e)
            REGISTRY = None
    recs = build_csrc() + build_amac() + build_exch()
    recs = merge_by_institution(recs)
    ensure_unique_names(recs)
    from collections import Counter
    print("构建完成：%d 条（已按机构合并）→ %s" % (len(recs), OUT))
    print("  证监局:", sum(1 for r in recs if r["source"] == "证监局"),
          "中基协:", sum(1 for r in recs if r["source"] == "中基协"),
          "交易所:", sum(1 for r in recs if r["source"] == "交易所"))
    print("  机构+人员同案（D 含；）:", sum(1 for r in recs if "；" in r["subject"]))
    print("  分类:", dict(Counter(r["category"] for r in recs)))
    print("  被处罚机构/人=不明确:", sum(1 for r in recs if r.get("subject") == "不明确"))
    print("  含 高管：不明确:", sum(1 for r in recs
                                  if any("高管（" in m and "不明确" in m for m in r["measures"])))
    print("  剔除（非私募）:", len(EXCLUDED), " 中基协审计提示:", len(AUDIT))
    os.makedirs(os.path.dirname(os.path.abspath(OUT)) or ".", exist_ok=True)
    json.dump(recs, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    _dump_json(EXCLUDED_OUT, EXCLUDED)
    _dump_json(AUDIT_OUT, AUDIT)
    print("  剔除明细 →", EXCLUDED_OUT, " 审计提示 →", AUDIT_OUT)
    return 0


if __name__ == "__main__":
    main()
