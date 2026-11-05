#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""交易所（上交所/深交所/北交所）监管处罚条目全量拉取与私募识别。

用法：
    python fetch_exchange.py --start 2025-09-24 --end 2026-09-29 \
        --out exch_candidates.json [--cache-dir ./.cache_exch] [--skip-bse]

产出 JSON：
    {"sse": [...], "szse": [...], "bse": [...], "stats": {...}}
每条记录含 date/title/target/measure/file/section(或 cat)/hit_kind(strong|weak)/body_head

关键实现要点（踩过的坑，勿改）：
  * 深交所列表 CATALOGID 必须带 `_snapshot` 后缀，裸 CID 会返回"请补充检索条件"提示
    （误判为检索式栏目会让一年窗口只拿到 1/5：36 条 → 404 条）
  * 深交所附件域名必须是 reportdocs.static.szse.cn，www.szse.cn 下 404（会产出 0 字节文件）
  * **深交所附件 CDN 对本 IP 有频控**：并发猛下会触发整段 403 封锁（38 分钟 0 份）。
    附件下载必须**顺序 + 每份间隔 + 遇 403 退避重试 + 续跑**，禁用线程池（见 0.4.97）
  * 深交所列表不返回被处罚主体，必须下载 PDF 抽正文后用正则抽主体名
  * 上交所附件有 .pdf 也有 .doc；表格有 tr/td 与 dl/dd 两种结构
  * 北交所 bse.cn 有 JS 反爬，必须用无头浏览器（playwright + 本机 Chrome）
  * 下载结果必须校验 %PDF- 与文本可抽取，否则计为失败并计入 stats
"""
import argparse
import json
import os
import re
import ssl
import sys
import time
import urllib.request

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE

# 2026-10-28：抓取完整性自证——把"自己都不信"的情况记下来（频控空返回、
# 平台声明数远大于实抓数），供 assert_crawl_complete.py 统一判定。
# 只记录事实，不做判断，也不改变既有重试/退避行为。
CRAWL_WARNINGS = []

# 2026-10-28：平台侧「最新一条」的日期，用于自证"窗口内 0 条 = 真无新增"。
# 交易所三所的自律措施本来就极稀疏（有私募字样的更是少数），日常增量 0 条是常态；
# 没有这个字段，断言只能把 0 条一律当 FAIL，等于每轮都被卡住。
CRAWL_LATEST = {}

# 2026-10-28：深交所逐栏目「声明条数 vs 实抓条数」对账。
# 用途：断言 R4 现在能证明"声明 193 条就抓到 193 条"，而不是只看总数非零。
# 以前没有它，深交所常年只拿到 1/5 却毫无征兆（旧 CID 也返回 200 + 少量数据）。
CRAWL_CATS = {}
# 深交所正文下载的自证账本（2026-10-28）：candidates / screened_out / downloaded /
# cached / dl_err / blocked 六个数必须一起报，否则「0 命中」到底是"无私募案例"
# 还是"被 403 挡住"就说不清 —— 这两种结论的处置完全相反。
SZSE_BODY_WHY = {}

# 【已废弃·勿复活】曾把深交所 4 个栏目判定为"检索式栏目"（不带关键词就 pagecount=0），
# 据此提前 return、只做关键词抽样兜底 ⇒ 深交所常年只拿到真实数据的 1/5。
# 2026-10-28 用浏览器抓包定位真因：**CATALOGID 少了 `_snapshot` 后缀**。
# 页面实际请求的是 `<CID>_snapshot`，用裸 CID 请求才会返回"检索式"那个提示头。
# 改用 _snapshot 后五栏全部可无条件枚举（实测 0 + 61 + 193 + 129 + 11 条）。
# 结论：深交所不存在"检索式栏目"，也不存在频控；再遇到 0 条，先查 CID 后缀。


def cache_name(url):
    """缓存文件名必须按 URL 哈希命名，不能按列表序号命名。

    2026-10-08 实测教训：原实现用 ``szse_<enumerate序号>.pdf`` 命名，
    而序号是**窗口过滤之后**的下标。窗口一旦滚动（起点前移/后移），
    同一序号对应的条目就换人了，旧缓存被张冠李戴地复用：
    idx284 标题是「上海曜瞿如…」，读出来的正文却是「宁波GQY视讯」。
    ⇒ 表现为「成功解析 402/402」但主体判定全错（假阴性 + 假阳性）。
    """
    import hashlib
    return hashlib.sha1((url or "").encode("utf-8")).hexdigest()[:20] + ".pdf"

# ---------------- 上交所 ----------------
SSE_SECTIONS = {
    "会员监管-监管措施": "regulation/members/measures",
    "会员监管-纪律处分": "regulation/members/disposition",
    "交易监管-纪律处分": "regulation/trading/disposition",
    "发行上市审核-监管措施": "regulation/listing/measures",
    "自律管理复核": "regulation/review",
    "自律监管与市场服务事项": "regulation/ma",
}
SSE_DOC_RE = re.compile(r'href="(/regulation/[^"]+?\.(?:pdf|docx?|xlsx?))\s*"', re.I)
TITLE_RE = re.compile(r'title="([^"]*)"')

# ---------------- 深交所 ----------------
# 坑：同一个 CATALOGID 下不同 TABKEY 是不同数据集，默认 tab1 可能返回 0 条。
#     实测 1800_jgxxgk_cf（纪律处分）必须 TABKEY=tab2 才有数据（145 页），
#     用 tab1 直连返回 0 → 曾误判「该栏目无数据」而漏掉整个纪律处分栏目。
#     因此每个 CATALOGID 必须显式声明要跑的 TABKEY 列表。
# ⚠ 2026-10-28 重大修正：以上 5 个 CATALOGID **全部是错的**，必须带 `_snapshot` 后缀。
# 发现过程：接口自己返回 metadata.header「请同时输入标题关键词…进行检索」、recordcount 恒 0，
# 旧实现据此判定"检索式栏目/被频控"，实际上——**页面真正请求的是 `<CID>_snapshot`**。
# 用浏览器抓包才看到（手工猜参数怎么试都是 0，因为 CID 本身就不对）。
# 实测对比（窗口 2025-10-28~2026-10-28）：
#   members_supervision(0)   → members_supervision_snapshot   (快照，pagesize=-1，一次返回)
#   ZJJGJG(0)                → ZJJGJG_snapshot                (61 条)
#   1800_jgxxgk(36)          → 1800_jgxxgk_snapshot           (193 条) ← 旧 CID 只拿到 1/5！
#   1800_jgxxgk_cf(0)        → 1800_jgxxgk_cf_snapshot + tab2 (129 条)
#   ZQ_JGCS(0)               → ZQ_JGCS_snapshot               (11 条)
# 也就是说：**深交所这条线此前几乎是全盲的**，而且"看起来像频控"的假象持续了很久。
# 字段与参数差异（必须按栏目区分，写错就恒 0）：
#   日期参数：members/ZJJGJG/ZQ_JGCS 用 txtStartDate+txtEndDate；
#             1800_jgxxgk/1800_jgxxgk_cf 用 txtDate+txtEnd（YYYY-MM-DD）。
#   TABKEY  ：1800_jgxxgk_cf 必须 tab2（tab1 恒 0），其余 tab1。
#   特殊    ：members_supervision_snapshot 带日期参数反而返回 0 → 只能不带日期查，
#             且要 PAGENO=0 + loading=first（pagesize=-1，一次返回最新若干条）。
SZSE_CATS = {
    "members_supervision_snapshot": ("会员及其他交易参与人监管",
                                     "https://www.szse.cn/disclosure/supervision/members/index.html",
                                     ["tab1"], ("txtStartDate", "txtEndDate"), "no_date"),
    "ZJJGJG_snapshot": ("中介机构监管",
                        "https://www.szse.cn/disclosure/supervision/supervise/index.html",
                        ["tab1"], ("txtStartDate", "txtEndDate"), "date"),
    "1800_jgxxgk_snapshot": ("监管措施(上市公司)",
                             "https://www.szse.cn/disclosure/supervision/measure/measure/index.html",
                             ["tab1"], ("txtDate", "txtEnd"), "date"),
    "1800_jgxxgk_cf_snapshot": ("纪律处分(上市公司)",
                                "https://www.szse.cn/disclosure/supervision/measure/measure/index.html",
                                ["tab2"], ("txtDate", "txtEnd"), "date"),
    "ZQ_JGCS_snapshot": ("债券监管措施",
                         "https://www.szse.cn/disclosure/supervision/measure/measure/index.html",
                         ["tab1"], ("txtStartDate", "txtEndDate"), "date"),
}
SZSE_STATIC = "https://reportdocs.static.szse.cn"

# ---------------- 北交所 ----------------
BSE_TARGETS = [
    ("交易监管", "https://www.bse.cn/disclosure/trade_supervise.html"),
    ("自律监管措施", "https://www.bse.cn/disclosure/supervise.html"),
    ("纪律处分", "https://www.bse.cn/disclosure/disciplinary_aciton.html"),
    ("会员及其他交易参与人监管", "https://www.bse.cn/disclosure/securities_trader_supervise.html"),
]

SUBJ_KEY = ["私募基金", "私募证券", "私募股权", "创业投资", "股权投资", "投资管理",
            "资产管理", "投资基金", "资本管理", "合伙企业"]
# 深交所正文下载前的**元数据预筛词表**（2026-10-28）：
# 深交所一年窗口 404 条记录里，能称得上"私募监管处罚"的极少（历史仅汇垠澳丰等个位数），
# 其余是会计师事务所/资产评估机构/上市公司/债券发行人。旧实现对 362 个附件全部下载正文，
# 既慢又极易触发静态域名 403 整段封锁（实测 38 分钟 0 份）。
# 现在先按「标题+对象+措施+文号」做粗筛，只下载命中者的正文：362 → 54（约 15%）。
# 词表刻意放宽（含 投资/资本/合伙/资产管理 等），宁可多下不可漏掉可能的私募主体。
# 需全量下正文时用 --szse-full-body（会在封锁期卡很久，慎用）。
SZSE_BODY_KW = ("私募", "基金", "投资", "资产管理", "合伙", "资本", "创投",
                "股权", "信托", "理财", "资产管理计划")
# 坑（2026-09-29）：句式词表必须覆盖深交所**纪律处分**栏目的写法
#   「关于对A、B、C、D给予纪律处分的决定」—— 早期词表只有 监管函/监管措施/通报批评/公开谴责/限制交易，
#   漏了「纪律处分」→ 该栏目 2894 条记录的主体**全部抽空**，主体级命中判定完全失效。
BODY_RE = re.compile(r"关于对(.{2,120}?)(?:予以|采取|给予|的)?"
                     r"(?:监管函|监管措施|自律监管措施|通报批评|公开谴责|限制交易|纪律处分|风险警示)")


def extract_subject(text):
    """从决定书正文抽取被处罚主体名（可能多个，返回顿号分隔的原串）。

    坑1：PDF 抽出的标题常跨行，例如
        「关于对广东汇垠澳丰股权投资基金管理\\n有限公司的监管函」
    正则的 `.` 默认不匹配换行，直接 search 会取不到 —— 必须先去掉空白再匹配
    （实测：不去空白时汇垠澳丰/杭州御硕两条都漏掉）。
    坑2：深交所纪律处分是**多主体**写法（「关于对A、B、C、D给予纪律处分的决定」），
    `.{2,60}?` 的非贪婪 + 短上限会截断；上限放宽到 120 并允许顿号，
    由 `split_subjects()` 拆成多个主体逐个判定。
    """
    head = re.sub(r"\s+", "", text[:1500])
    m = BODY_RE.search(head)
    return m.group(1).strip("：: ,，") if m else ""


def split_subjects(subj):
    """把「A、B、C」或「A、B及C」拆成单个主体列表；顺带剥掉「及相关当事人」等套话。"""
    if not subj:
        return []
    s = re.sub(r"(及相关当事人|及相关责任主体|及签字注册会计师.*|及注册会计师.*|"
               r"及资产评估师.*|及签字评估师.*)$", "", subj)
    parts = re.split(r"[、,，]|以及|及(?=[^、]*?(?:公司|中心|合伙|企业|基金|事务所))", s)
    return [p.strip("：: ") for p in parts if len(p.strip()) >= 3]


# 最近一次请求的 HTTP 状态码（2026-10-28）：
# 旧版 http_get 把异常一律吞掉、只返回 None，调用方无从区分
# 「网络超时」「404 文件不存在」「403 整段封锁」——而这三者的处置完全不同。
# 尤其是 403：深交所静态域名对本 IP 封锁时，每个附件都要烧满重试才失败，
# 实测 362 个附件卡死 38 分钟、一份都没下下来，且日志全程空白（无任何信号）。
LAST_STATUS = {"code": 0, "err": ""}


def http_get(url, ref=None, timeout=40, retry=3, binary=False):
    for _ in range(retry):
        try:
            h = {"User-Agent": UA}
            if ref:
                h["Referer"] = ref
            r = urllib.request.urlopen(urllib.request.Request(url, headers=h),
                                       timeout=timeout, context=SSL_CTX)
            b = r.read()
            LAST_STATUS.update(code=int(getattr(r, "status", 200) or 200), err="")
            return b if binary else b.decode("utf-8", "ignore")
        except urllib.error.HTTPError as e:
            LAST_STATUS.update(code=int(getattr(e, "code", 0) or 0), err=str(e)[:120])
            if int(getattr(e, "code", 0) or 0) in (403, 429):
                break          # 封锁/限流：重试只会更久，立刻交回调用方做熔断
        except Exception as e:
            LAST_STATUS.update(code=0, err=type(e).__name__ + " " + str(e)[:120])
            time.sleep(1.5)
    return None


def _clean(s):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", s)).strip()


# ================= 列表缓存与频控探测（2026-10-27 重构：机制化解决深交所频控，不再"躲"） =================
# 历史教训（2026-10-12/10-08）：深交所频控触发成本极低，重复重拉列表 API 会让
# 5 个 CATALOGID 全部返回 0 条（与"真无数据"无法区分）；封锁期请求返回空 data 而非异常，
# 不能因为"没报错"就当解封。
# 解决机制（四件套，与附件下载已有的缓存/退避互补）：
#   1. 列表快照缓存：列表 API 命中缓存后绝不重拉，根治"重复重拉触发频控"。
#   2. 频控探测：列表请求"成功但 data 为空"→ 判为疑似频控，退避后重试一次；不空再缓存。
#   3. 判解封标准恒为 len(data) > 0（不是"无异常"）。
#   4. 附件下载已有 URL 哈希缓存 + 并发 3 + 每次节流 + 403 指数退避（见 fetch_szse.job）。
# 因此：合规 re-run 不会触发频控（列表与附件都走缓存），无需"单独跑/避开"之类 workaround。

def _list_cache_path(cache_dir, tag):
    return os.path.join(cache_dir, "list_" + tag + ".json")


# 列表快照 TTL：12h 内复用（避开开发期反复重拉触发频控）；
# 超过则刷新——每日定时必过期重拉，保证增量抓到新发布文书（完整性优先）。
LIST_CACHE_TTL = 12 * 3600


def _load_list_cache(cache_dir, tag):
    p = _list_cache_path(cache_dir, tag)
    if os.path.exists(p):
        try:
            if time.time() - os.path.getmtime(p) <= LIST_CACHE_TTL:
                return json.load(open(p, encoding="utf-8"))
        except Exception:
            return None
    return None


def _save_list_cache(cache_dir, tag, obj):
    try:
        os.makedirs(cache_dir, exist_ok=True)
        json.dump(obj, open(_list_cache_path(cache_dir, tag), "w", encoding="utf-8"))
    except Exception:
        pass


def _api_list(cache_dir, tag, url, ref, refresh):
    """拉一次列表接口，带缓存与频控探测。返回解析后的 JSON（通常是 dict，含 data 键）。"""
    if not refresh:
        cached = _load_list_cache(cache_dir, tag)
        if cached is not None and (cached.get("data") or []):   # 非空缓存直接复用，不重拉
            return cached
    # 空缓存 / 强制刷新 / 空缓存命中 → 实拉一次（频控期空返回会被重试甄别）
    res = {}
    # 频控甄别升级（2026-10-28 实测）：深交所在被限流时会**持续**返回 200+data:[]，
    # 只重试一次的 3 秒退避根本等不到解封 → 整轮 5 个栏目全部 0 条，
    # 断言只能看到"交易所 0 条"，与"真无新增"无法区分。
    # 现在按 5/15/40 秒三级退避，仍空才认定（并记 warning，交由断言判 FAIL）。
    for i, wait in enumerate((-1, 5, 15, 40)):
        if wait >= 0:
            time.sleep(wait)
        b = http_get(url, ref=ref, timeout=40, retry=3)
        if not b:
            continue
        try:
            j = json.loads(b)
            res = j[0] if isinstance(j, list) else j
        except Exception:
            continue
        if (res.get("data") or []):
            break
        # 注意：曾在这里按 metadata.header「请同时输入标题关键词」判定"检索式栏目"并提前返回。
        # 2026-10-28 已证伪——真正原因是 CATALOGID 少了 `_snapshot` 后缀（见 SZSE_CATS 注释）。
        # 该提前返回会把"CID 写错"伪装成"接口只支持检索"，永久掩盖问题，故删除。
        # 现在只按「声明条数」判断：声明 0 条就是真没有，不再退避重试（省时间）。
        declared = int(((res.get("metadata") or {}).get("recordcount") or 0) or 0)
        if declared == 0:
            break
        # 声明有数据但返回空 = 确定的频控特征
        declared = int(((res.get("metadata") or {}).get("recordcount")
                        or (res.get("metadata") or {}).get("pagecount") or 0) or 0)
        print("  ⚠ [频控探测] %s 列表返回 0 条但请求成功（声明 %d 条，疑似限流，非真无数据），"
              "第 %d 次退避重试…" % (tag, declared, i + 1), file=sys.stderr)
        CRAWL_WARNINGS.append({"venue": "szse", "kind": "throttle_probe", "tag": tag,
                               "declared": declared})
    if (res.get("data") or []):
        _save_list_cache(cache_dir, tag, res)   # 只缓存非空结果，频控空返回不缓存（下轮再验证）
    return res


# ================= 上交所 =================
SSE_PAGE_RE = re.compile(r'href="(/regulation/[^"]+?\.shtml)\s*"', re.I)


def sse_parse(html):
    r"""解析上交所栏目列表。

    2026-10-28 实测：早期（2008~2010）文书没有 PDF 附件，正文直接是 `.shtml` 详情页，
    只认 `\.pdf|\.doc` 会漏掉这些行 —— 「交易监管-纪律处分」声明 56 实抓 47 的 8 条差额全在这里
    （末页 11 行只有 3 行带附件）。虽然这些老文书永远落不进 365 天窗口，
    但"实抓 < 声明"会让断言长期挂一条无法解释的告警，也让人分不清"漏抓"与"格式不同"。
    这里补一条 `.shtml` 后备：仍按日期过滤，窗口外的自然被剔除。
    """
    out, seen = [], set()
    for block in re.findall(r"<dd[^>]*>(.*?)</dd>", html or "", re.S):
        href = SSE_DOC_RE.search(block) or SSE_PAGE_RE.search(block)
        if not href:
            continue
        title = TITLE_RE.search(block)
        span = re.search(r"<span[^>]*>(.*?)</span>", block, re.S)
        d = re.search(r"(\d{4}-\d{2}-\d{2})", _clean(span.group(1)) if span else "")
        if href.group(1) in seen:
            continue
        seen.add(href.group(1))
        out.append({"measure": "", "title": (title.group(1) if title else _clean(block))[:200],
                    "target": "", "date": d.group(1) if d else "",
                    "pdf": "https://www.sse.com.cn" + href.group(1)})
    for r in re.findall(r"<tr[^>]*>(.*?)</tr>", html or "", re.S):
        href = SSE_DOC_RE.search(r) or SSE_PAGE_RE.search(r)
        if not href:
            continue
        title = TITLE_RE.search(r)
        cells = [_clean(t) for t in re.findall(r"<td[^>]*>(.*?)</td>", r, re.S)]
        if href.group(1) in seen:
            continue
        seen.add(href.group(1))
        d = None
        for c in reversed(cells):
            d = re.search(r"(\d{4}-\d{2}-\d{2})", c)
            if d:
                break
        out.append({"measure": cells[0] if cells else "",
                    "title": (title.group(1) if title else
                              (cells[1] if len(cells) > 1 else ""))[:200],
                    "target": cells[2] if len(cells) > 2 else "",
                    "date": d.group(1) if d else "",
                    "pdf": "https://www.sse.com.cn" + href.group(1)})
    return out


def fetch_sse(start, end, cache_dir="", refresh=False):
    rows = []
    os.makedirs(cache_dir, exist_ok=True) if cache_dir else None
    for name, sec in SSE_SECTIONS.items():
        base = "https://www.sse.com.cn/" + sec + "/"
        html = _load_list_cache(cache_dir, "sse_" + name) if cache_dir and not refresh else None
        if html is None:
            html = http_get(base, ref="https://www.sse.com.cn/")
            if html and cache_dir:
                _save_list_cache(cache_dir, "sse_" + name, {"data": html})
        # 缓存存的是 {"data": <html>}（为了兼容 _api_list 的 JSON 结构），取回时必须拆开。
        # 2026-10-28 实测：不拆的话第二次运行 re.search 直接 TypeError: expected string,
        # got 'dict' → 整个交易所阶段崩溃 → 断言读到上一轮残留 exch.json，
        # 报出 3 条与本轮无关的 FAIL（"看起来像频控，其实是代码崩"）。
        if isinstance(html, dict):
            html = html.get("data") or ""
        if not html:
            print(f"  [SSE] {name} 首页获取失败", file=sys.stderr)
            continue
        m = re.search(r"createPageHTML\('paging',\s*(\d+),\s*(\d+),'([^']*)','([^']*)',(\d+)", html)
        if m:
            pages, prefix, ext, total = int(m.group(1)), m.group(3), m.group(4), int(m.group(5))
        else:
            pages, prefix, ext, total = 1, "s_list", "shtml", 0
        got = sse_parse(html)
        for n in range(2, pages + 1):
            url = base + f"{prefix}_{n}.{ext}"
            ph = _load_list_cache(cache_dir, "sse_%s_%d" % (name, n)) if cache_dir and not refresh else None
            if ph is None:
                ph = http_get(url, ref="https://www.sse.com.cn/")
                if ph and cache_dir:
                    _save_list_cache(cache_dir, "sse_%s_%d" % (name, n), {"data": ph})
            if isinstance(ph, dict):        # 同上：缓存结构是 {"data": html}，必须拆开
                ph = ph.get("data") or ""
            if ph:
                got += sse_parse(ph)
        if total and len(got) < total * 0.5:
            print(f"  ⚠ [SSE] {name} 声明 {total} 实抓 {len(got)}（偏低，疑似被频控/拦截，"
                  "非真无数据）", file=sys.stderr)
            CRAWL_WARNINGS.append({"venue": "sse", "kind": "declared_vs_fetched",
                                   "section": name, "declared": total,
                                   "fetched": len(got), "pages": pages})
        print(f"  [SSE] {name:16s} 声明{total:4d} 实抓{len(got):4d} 页{pages}")
        for r in got:
            r["section"] = name
        rows += got
    dated = [r["date"] for r in rows if r.get("date")]
    CRAWL_LATEST["sse"] = max(dated) if dated else ""
    win = [r for r in rows if start <= r["date"] <= end]
    for r in win:
        r["hit_kind"] = "strong" if [k for k in SUBJ_KEY
                                     if k in r["title"] or k in r["target"]] else ""
    return win


# ================= 深交所 =================
def szse_api(cid, ref, page=1, tab="tab1", cache_dir="", refresh=False,
             params=(), loading_first=False, tag_suffix=""):
    """深交所列表接口。params=[(k,v)...] 用于日期范围；loading_first 用于快照型栏目。

    缓存 tag 必须带上参数，否则「带日期」与「不带日期（最新探测）」两次查询会互相串味：
    先查到带日期的结果被缓存，再查不带日期时命中缓存返回同一份，
    latest 自证就永远是窗口内的日期，等于自证了个寂寞。
    """
    u = (f"https://www.szse.cn/api/report/ShowReport/data?SHOWTYPE=JSON&CATALOGID={cid}"
         f"&TABKEY={tab}&PAGENO={page}&random=0.5")
    for k, v in params:
        u += "&%s=%s" % (k, urllib.parse.quote(str(v)))
    if loading_first:
        u += "&loading=first"
    tag = "szse_%s_%s_%d%s%s" % (cid, tab, page, tag_suffix,
                                 ("_" + "_".join("%s%s" % (k, v) for k, v in params)) if params else "")
    return _api_list(cache_dir, tag, u, ref, refresh)


def szse_norm(cid, rec):
    """归一化深交所记录。

    坑：不同 CATALOGID 的字段名完全不同，必须全兼容，否则新栏目恒为空：
        1800_jgxxgk    → fhrq / hjbt / gkxx_sjdx / gkxx_jgcs
        1800_jgxxgk_cf → xx_fwrq / xx_bt / ck / xx_cflb   （日期字段是 xx_fwrq）
        ZQ_JGCS        → fhrq / hjbt / dx / hh / lx       （对象在 dx，文号在 hh）
    """
    html = rec.get("hjbt") or rec.get("ck") or rec.get("hjnr") or rec.get("title") or ""
    m = re.search(r"encode-open='([^']+)'", html) or re.search(r'href="(http[^"]+)"', html)
    link = m.group(1) if m else ""
    if link.startswith("/"):
        link = "https://www.szse.cn" + link
    return {"cat": cid,
            "date": str(rec.get("fhrq") or rec.get("xx_fwrq") or rec.get("punishdate")
                        or rec.get("gkxx_gdrq") or "")[:10],
            "title": _clean(rec.get("xx_bt") or rec.get("hjbt") or rec.get("hjnr")
                            or rec.get("title") or "")[:220],
            "target": _clean(rec.get("dx") or re.sub(r"<[^>]+>", "/", rec.get("gkxx_sjdx") or "")),
            "measure": rec.get("xx_cflb") or rec.get("gkxx_jgcs") or rec.get("lx") or "",
            "doc_no": rec.get("hh") or "", "file": link}


def szse_fix(u):
    """www.szse.cn/UpFiles/... → reportdocs.static.szse.cn/UpFiles/...（否则 404）"""
    m = re.search(r"(/UpFiles/.+)$", u)
    return SZSE_STATIC + m.group(1) if m else u


def fetch_szse(start, end, cache_dir, refresh=False, full_body=False):
    os.makedirs(cache_dir, exist_ok=True)
    rows = []
    for cid, (name, ref, tabs, datep, mode) in SZSE_CATS.items():
        recs = []
        for tab in tabs:
            declared = 0
            if mode == "no_date":
                # 会员监管是快照型（pagesize=-1），且带日期参数会返回 0 → 只能不带日期 + PAGENO=0
                res = szse_api(cid, ref, 0, tab, cache_dir, refresh,
                               params=(), loading_first=True, tag_suffix="_nodate")
                got = list(res.get("data") or [])
                declared = int(((res.get("metadata") or {}).get("recordcount") or 0) or 0)
                print(f"  [SZSE] {name:20s} {tab} 共{len(got):5d}（快照型，不带日期，PAGENO=0）")
                recs += got
            else:
                pr = ((datep[0], start), (datep[1], end)) if datep else ()
                first = szse_api(cid, ref, 1, tab, cache_dir, refresh, params=pr)
                pages = int((first.get("metadata") or {}).get("pagecount") or 1)
                declared = int(((first.get("metadata") or {}).get("recordcount") or 0) or 0)
                got = list(first.get("data") or [])
                for p in range(2, pages + 1):
                    got += list(szse_api(cid, ref, p, tab, cache_dir, refresh, params=pr).get("data") or [])
                print(f"  [SZSE] {name:20s} {tab} 共{len(got):5d} 页{pages}"
                      + (f" 声明{declared}" if declared else ""))
                recs += got
            CRAWL_CATS[f"{cid}/{tab}"] = {"name": name, "declared": declared, "got": len(got)}
            if declared and len(got) < declared:
                # 声明数 > 实抓数 是唯一可信的"漏抓"信号（分页没翻完 / 中途被限流）。
                CRAWL_WARNINGS.append({"venue": "szse", "kind": "page_short", "tag": f"{cid}/{tab}",
                                       "declared": declared, "got": len(got)})
            # 自证「平台最新一条」：不带日期查第 1 页（列表按日期倒序），
            # 这样窗口内 0 条时才能证明"确实没有更新"，而不是"被日期参数筛没了"。
            probe = szse_api(cid, ref, 1, tab, cache_dir, refresh,
                             params=(), tag_suffix="_latest")
            for r in (probe.get("data") or [])[:1]:
                d = szse_norm(cid, r).get("date") or ""
                if d and d > (CRAWL_LATEST.get("szse") or ""):
                    CRAWL_LATEST["szse"] = d
        rows += [szse_norm(cid, r) for r in recs]
    dated = [r["date"] for r in rows if r.get("date")]
    CRAWL_LATEST["szse"] = max(dated) if dated else ""
    win_all = [r for r in rows if start <= r["date"] <= end and r["file"]]
    # 元数据预筛：只下载「可能是私募主体」的正文（见 SZSE_BODY_KW 注释）
    if full_body:
        win, screened_out = win_all, 0
    else:
        win, screened_out = [], 0
        for r in win_all:
            meta = "%s %s %s %s" % (r.get("title") or "", r.get("target") or "",
                                    r.get("measure") or "", r.get("doc_no") or "")
            if any(k in meta for k in SZSE_BODY_KW):
                win.append(r)
            else:
                screened_out += 1
    print(f"  [SZSE] 窗口内 {len(win_all)} 条 → 元数据预筛后 {len(win)} 条需下正文"
          f"（筛掉 {screened_out} 条：标题/对象/措施均不含私募相关词）")
    why = {"cached": 0, "dl_http": 0, "dl_err": 0, "screened_out": screened_out,
           "candidates": len(win_all), "blocked": 0}
    # 礼貌顺序下载（2026-10-28 修正，落实用户「在频控限制以下获取」要求）：
    # 深交所静态资源域名 reportdocs.static.szse.cn 对本 IP 有频控——并发 + 高频会触发
    # 整段 403 封锁（实测并发 3 + 0.8s 间隔卡死 38 分钟 0 份，且日志全空无任何信号）。
    # 对策（与证监局/中基协"礼貌重试 + 续跑"完全一致，深交所不再特殊对待）：
    #   ① 完全顺序（不用线程池），每两份之间 sleep DL_GAP；
    #   ② 单文件遇 403/429 不"熔断跳过"，而是短退避后重试同一文件（最多 6 次）；
    #   ③ 仅当"连续 8 次 403，疑似本 IP 进入冷却"才暂停整段，已下的走缓存，
    #      下轮自动续下（不在一个进程里干等数小时冷却，交给定时任务分次完成）；
    #   ④ 全程只写附件缓存、绝不动台账/PDF，故任何中途停止都 100% 可恢复。
    DL_GAP = float(os.environ.get("SZSE_DL_GAP", "2.5"))
    consec_403 = 0
    blocked = False
    out = []
    n = len(win)
    for i, r in enumerate(win):
        p = os.path.join(cache_dir, cache_name(szse_fix(r["file"])))
        data = None
        # 缓存命中：已下过且是合法 PDF 的直接复用（续跑 / 重跑都不重下）
        if os.path.exists(p) and os.path.getsize(p) > 1024:
            _b = open(p, "rb").read()
            if _b[:5] == b"%PDF-":
                data = _b
                why["cached"] += 1
        if data is None:
            for attempt in range(6):
                data = http_get(szse_fix(r["file"]),
                                ref=SZSE_CATS.get(r["cat"], ("", "", ""))[1],
                                timeout=30, binary=True, retry=1)
                if data and len(data) >= 1024 and data.startswith(b"%PDF"):
                    consec_403 = 0
                    break
                code = int(LAST_STATUS.get("code") or 0)
                if code in (403, 429):
                    consec_403 += 1
                    backoff = min(5 * (2 ** attempt), 120)   # 5,10,20,40,80,120 秒
                    print("  [SZSE] #%d 遇 %d（连续%d），退避 %.0fs 后重试同一文件"
                          % (i, code, consec_403, backoff), file=sys.stderr)
                    time.sleep(backoff)
                    if consec_403 >= 8:
                        blocked = True
                        break
                else:
                    consec_403 = 0
                    time.sleep(2)
            if blocked:
                why["blocked"] = n - i
                break
            if data is None:
                why["dl_err"] += 1
                out.append({"idx": i, "ok": False, **r})
                continue
            open(p, "wb").write(data)
            why["dl_http"] += 1
            time.sleep(DL_GAP)        # 礼貌节流：两份之间留间隔，避免触发频控
        # 抽正文 + 抽主体
        try:
            import fitz
            d = fitz.open(p)
            t = "".join(pg.get_text() for pg in d)
            d.close()
        except Exception:
            why["dl_err"] += 1
            out.append({"idx": i, "ok": False, **r})
            continue
        if not t.strip():
            why["dl_err"] += 1
            out.append({"idx": i, "ok": False, **r})
            continue
        subj = extract_subject(t)
        subs = split_subjects(subj)
        strong = [k for k in SUBJ_KEY if any(k in s for s in subs)]
        rec = {"idx": i, "ok": True, "subj": subj[:120], "subjects": subs[:8],
               "body_head": _clean(t[:200]),
               "hit_kind": "strong" if strong else ("weak" if
                           [k for k in SUBJ_KEY if k in t[:600]] else ""), **r}
        rec["file"] = szse_fix(r["file"])    # 必须放 **r 之后，否则被原始 www 链接覆盖
        out.append(rec)
    ok = sum(1 for r in out if r.get("ok"))
    not_scanned = len(win_all) - len(win)
    print(f"  [SZSE] 成功解析 {ok} / {len(win)}  缓存命中 {why['cached']} "
          f"新下载 {why['dl_http']} 失败 {why['dl_err']} 被封锁暂停 {why['blocked']}")
    if blocked:
        print("  ⚠ [SZSE] 静态域名对本 IP 临时封锁：本轮余 %d 个候选附件未下载，"
              "已暂停、已下部分走缓存，下轮（或定时任务）自动续下。"
              "**未取到正文 ≠ 无私募案例**——列表完整性不受影响（声明数==实抓数）。"
              % why["blocked"], file=sys.stderr)
        CRAWL_WARNINGS.append(
            {"venue": "szse", "kind": "static_403_blocked", "tag": "reportdocs.static.szse.cn",
             "note": "连续 403/429，静态域名对本 IP 临时封锁；已暂停本轮剩余附件下载，"
                     "已下的走缓存，下轮自动续下（列表完整性不受影响）"})
    elif why["dl_err"]:
        print("  ⚠ [SZSE] %d 个附件下载失败（非封锁类，需排查 URL/网络），"
              "下轮自动续下。" % why["dl_err"], file=sys.stderr)
    SZSE_BODY_WHY.update(why)
    SZSE_BODY_WHY["not_scanned"] = not_scanned
    SZSE_BODY_WHY["ok"] = ok
    return out


# ================= 北交所 =================
def fetch_bse(start, end, cache_dir, skip=False):
    if skip:
        return []
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        print("  [BSE] 未安装 playwright，跳过（北交所必须浏览器承载）", file=sys.stderr)
        return []
    # 可迁移（2026-10-27）：不要写死 Windows 安装路径。优先环境变量，再由
    # save_decision_pdf.browser_candidates() 跨平台探测；都找不到就交给 playwright
    # 自带的 chromium（executable_path=None），而不是直接崩在 launch 上。
    chrome = os.environ.get("CHROME_PATH") or ""
    if not chrome or not os.path.exists(chrome):
        try:
            import save_decision_pdf as _sdp
            chrome = next((p for p in _sdp.browser_candidates() if os.path.exists(p)), "")
        except Exception:
            chrome = ""
    if not chrome:
        # playwright 新版 headless=True 默认要 chromium_headless_shell；
        # 但 `playwright install chromium` 装的是**完整 chromium**。显式取完整二进制路径
        # 传给 executable_path，headless 也用它，无需再装 headless_shell（可移植）。
        try:
            with sync_playwright() as _p:
                chrome = _p.chromium.executable_path
        except Exception:
            chrome = ""
        if chrome:
            print("  [BSE] 未探测到本机浏览器，改用 playwright 完整 chromium", file=sys.stderr)
    os.makedirs(cache_dir, exist_ok=True)
    rows, docs = {}, []
    with sync_playwright() as p:
        try:
            br = p.chromium.launch(executable_path=(chrome or None), headless=True,
                                   args=["--no-sandbox", "--disable-gpu"])
        except Exception as e:
            print("  [BSE] 浏览器启动失败:", repr(e)[:120], file=sys.stderr)
            return []
        ctx = br.new_context(user_agent=UA)
        pg = ctx.new_page()
        for name, url in BSE_TARGETS:
            try:
                pg.goto(url, wait_until="networkidle", timeout=90000)
                pg.wait_for_timeout(3000)
                raw = pg.eval_on_selector_all("table tr", """els => els.map(tr => ({
                    cells:[...tr.querySelectorAll('td')].map(td=>td.innerText.trim().replace(/\\s+/g,' ')),
                    hrefs:[...tr.querySelectorAll('a')].map(a=>a.getAttribute('href')||'')}))""")
                keep = [r for r in raw if r["cells"] and
                        re.search(r"\d{4}-\d{2}-\d{2}", " ".join(r["cells"]))]
                print(f"  [BSE] {name}: 行 {len(keep)}")
                for r in keep:
                    d = re.search(r"(\d{4}-\d{2}-\d{2})", " ".join(r["cells"]))
                    if d and d.group(1) > (CRAWL_LATEST.get("bse") or ""):
                        CRAWL_LATEST["bse"] = d.group(1)
                    if not (d and start <= d.group(1) <= end):
                        continue
                    for h in r["hrefs"]:
                        if h and re.search(r"\.(pdf|docx?|xlsx?)", h, re.I):
                            docs.append({"cat": name, "date": d.group(1),
                                         "title": r["cells"][0][:120],
                                         "url": h if h.startswith("http") else
                                         ("https://www.bse.cn" + h if h.startswith("/")
                                          else "https://www.bse.cn/" + h)})
            except Exception as e:
                print("  [BSE]", name, "ERR", repr(e)[:100], file=sys.stderr)
        out, why = [], {"cached": 0, "dl_http": 0, "dl_err": 0, "txt_err": 0, "blank": 0}
        for i, d in enumerate(docs):
            fn = os.path.join(cache_dir, cache_name(d["url"]))
            data = None
            if os.path.exists(fn) and os.path.getsize(fn) > 1024:
                data = open(fn, "rb").read()      # 复用缓存，避免重复触发 WAF
                why["cached"] += 1
            if data is None:
                # 坑：bse.cn 有 WAF（403 页面来自 blocksrc.haplat.net），短时间多请求会被
                #     整段 IP 限流 —— 表现为「全部附件 403 / 2957 字节 HTML」。
                #     对策：逐个请求 + 随机间隔 + 403 退避重试，失败必须显式报警而不是静默 continue。
                for attempt in range(3):
                    try:
                        r = ctx.request.get(d["url"], timeout=90000)
                        b = r.body()
                        if r.status == 200 and len(b) >= 500:
                            data = b
                            open(fn, "wb").write(b)
                            break
                        why["dl_http"] += 0 if attempt < 2 else 1
                        if attempt == 2:
                            print(f"    [BSE-DL-FAIL] status={r.status} size={len(b)} {d['url'][:90]}")
                        time.sleep(3 + attempt * 5)
                    except Exception as e:
                        if attempt == 2:
                            why["dl_err"] += 1
                            print(f"    [BSE-DL-ERR] {repr(e)[:90]} {d['url'][:90]}")
                        time.sleep(3 + attempt * 5)
                time.sleep(1.5)   # 节流，降低被 WAF 判定为机器人的概率
            if data is None:
                continue
            try:
                import fitz
                doc = fitz.open(fn)
                t = "".join(pg2.get_text() for pg2 in doc)
                doc.close()
            except Exception as e:
                why["txt_err"] += 1
                print(f"    [BSE-TXT-ERR] {repr(e)[:90]} {fn}")
                continue
            if not t.strip():
                why["blank"] += 1
                print(f"    [BSE-BLANK] 无可抽取文字（扫描件） {d['url'][:90]}")
                continue
            subj = extract_subject(t)
            subs = split_subjects(subj)
            strong = [k for k in SUBJ_KEY if any(k in s for s in subs)]
            out.append({**d, "file": d["url"], "subj": subj[:120], "subjects": subs[:8],
                        "body_head": _clean(t[:200]),
                        "hit_kind": "strong" if strong else (
                            "weak" if [k for k in SUBJ_KEY if k in t[:600]] else "")})
        print(f"  [BSE] 窗口内附件 {len(docs)}，下载并解析成功 {len(out)}  "
              f"失败明细 {why}")
        if docs and len(out) < len(docs) * 0.8:
            print("  ⚠ [BSE] 解析率过低：若 dl_http 很高，极可能是 bse.cn 的 WAF 已对本机 IP 限流"
                  "（403 + 2957 字节拦截页）。请放慢速度、复用 --cache-dir 缓存，"
                  "或间隔一段时间后再跑；**不得**把它当成「北交所无私募案例」的结论。",
                  file=sys.stderr)
        br.close()
        return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True, help="窗口起点 YYYY-MM-DD")
    ap.add_argument("--end", required=True, help="窗口终点 YYYY-MM-DD")
    ap.add_argument("--out", required=True, help="输出 JSON 路径")
    ap.add_argument("--cache-dir", default="./.cache_exch")
    ap.add_argument("--skip-bse", action="store_true")
    ap.add_argument("--refresh-lists", action="store_true",
                    help="强制重拉列表（默认复用 --cache-dir 里的列表快照，避免重拉触发频控）")
    ap.add_argument("--szse-full-body", action="store_true",
                    help="深交所下正文时不做元数据预筛（默认只下命中私募相关词的，"
                         "全量下极易触发静态域名 403 封锁，慎用）")
    a = ap.parse_args()

    print("== 上交所 ==")
    sse = fetch_sse(a.start, a.end, os.path.join(a.cache_dir, "sse"), a.refresh_lists)
    print("== 深交所 ==")
    szse = fetch_szse(a.start, a.end, os.path.join(a.cache_dir, "szse"),
                      a.refresh_lists, a.szse_full_body)
    print("== 北交所 ==")
    bse = fetch_bse(a.start, a.end, os.path.join(a.cache_dir, "bse"), a.skip_bse)

    res = {"sse": sse, "szse": szse, "bse": bse,
           "stats": {"window": [a.start, a.end],
                     "sse_total": len(sse), "sse_strong": sum(1 for r in sse if r.get("hit_kind") == "strong"),
                     "szse_total": len(szse), "szse_ok": sum(1 for r in szse if r.get("ok")),
                     "szse_strong": sum(1 for r in szse if r.get("hit_kind") == "strong"),
                     "szse_body": dict(SZSE_BODY_WHY),
                     "bse_total": len(bse), "bse_strong": sum(1 for r in bse if r.get("hit_kind") == "strong"),
                     "latest": dict(CRAWL_LATEST),
                     "latest_seen": max([v for v in CRAWL_LATEST.values() if v], default=""),
                     "szse_cats": dict(CRAWL_CATS),
                     "warnings": CRAWL_WARNINGS}}
    json.dump(res, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\n== 汇总 ==", json.dumps(res["stats"], ensure_ascii=False))
    print("== 私募强命中候选 ==")
    for r in sse:
        if r.get("hit_kind") == "strong":
            print("  SSE ", r["date"], r.get("title", "")[:60])
    for r in szse:
        if r.get("hit_kind") == "strong":
            print("  SZSE", r["date"], r.get("subj", "")[:60])
    for r in bse:
        if r.get("hit_kind") == "strong":
            print("  BSE ", r["date"], r.get("subj") or r.get("title", "")[:60])


if __name__ == "__main__":
    main()
