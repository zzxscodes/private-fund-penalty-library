# -*- coding: utf-8 -*-
"""私募监管处罚案例库 · 完整性与分类准确性校验（v2 格式）。

用法：
    python verify_library.py --root "<案例库根目录>" [--start 2025-10-26] [--end 2026-10-26]
                             [--window-by decided|publish] [--online]
                             [--ocr-scan sample:20] [--ocr-cache DIR] [--json report.json]

v2（2026-10-26）随台账格式 v2 重写：
  * 一行 = 一个案例（管理人 + 高管决定书合并），A1 改为「台账 K/L 列 ↔ 磁盘」双向核对；
  * 文号列已删除：A4 文号连号、B3 文号回现随之取消，判重改「局+管理人+处罚时间」+文件+链接；
  * 新增 B10 三核心列（事由/证据/依据）完整性检查；
  * B8/B9 落款 vs 挂网发布日期（台账 C 列），并新增 --window-by publish 支持按发布时间筛窗口；
  * D 段法规引用口径改为「决定书正文 ∪ 核心处罚依据列」。

四大类校验：
  A 完整性   —— 案例行↔磁盘文件双向、多值列对齐、窗口、重复、平台对账、三信源覆盖
  B 内容正确 —— PDF 可打开/有文字层、落款回现、管理人名称回现、文件命名规范、
               非占位、非空壳、日期逻辑、三核心列非空
  C 分类准确 —— 类型枚举合法、目录==台账、规则重分类比对、其他类占比
  D 法规依据 —— 06 目录↔清单一致、案例引用法规是否缺失、零引用法规

退出码 0 = 全部通过；1 = 存在必须处理的问题。
"""
import argparse, json, os, re, sys, urllib.request, urllib.parse, ssl
from collections import Counter, defaultdict
from datetime import date, datetime
import ledger_style as LS  # noqa: E402  运维目录唯一真源（work_dir）

try:
    import fitz  # pymupdf
except ImportError:
    print("缺少依赖：pip install pymupdf openpyxl")
    sys.exit(2)
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
CHANNELS = json.load(open(os.path.join(HERE, "..", "references", "bureau_channels.json"),
                          encoding="utf-8"))
BN = LS.BUREAU_CN            # 辖区中文名：真源在 ledger_style.BUREAU_CN（网页台账也用这份）

# 分类目录 / 关键词 / 加权 / 置信度阈值：统一从 classify_filter 导入（与 build_pipeline 同口径）
from classify_filter import (AMAC_CATS, CSRC_CATS, EXCH_CATS, CATS, STRONG_KW,
                            recat, rule_classify, classify)  # noqa: E402

UNKNOWN = ("不明确", "待补", "无")
# 2026-10-29 新增：文件名/主体里绝不允许出现的占位词。
# 教训：国厚资产两案落成「不明确.pdf」，珠海通沛复核书落成「不明确-珠海通沛…pdf」，
# 景行天成/中诚新川把正文里的「合规风控负责人缺位」当成高管名 → 「缺位-XXX.pdf」。
BAD_TOKEN = ("不明确", "待补", "缺位", "待核", "未命名", "无标题", "未知")


def _split_md_row(line):
    """拆一行 Markdown 表格，还原被转义的 `\\|`。"""
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip().replace("\\|", "|") for c in line.split("|")]


def _load_list_file(root, stem):
    """读取库根目录的清单文件：优先 `<stem>.md`（Markdown 表格），回退 `<stem>.json`。"""
    md_p = os.path.join(root, stem + ".md")
    js_p = os.path.join(root, stem + ".json")
    if os.path.exists(md_p):
        try:
            rows, header = [], None
            for ln in open(md_p, encoding="utf-8"):
                ln = ln.rstrip("\n")
                if not ln.strip().startswith("|"):
                    continue
                cells = _split_md_row(ln)
                if not cells or all(set(c) <= set("-: ") for c in cells if c):
                    continue                      # 分隔行 |---|---|
                if header is None:
                    header = [c.strip("` `#").strip() for c in cells]
                    continue
                if len(cells) < len(header):
                    continue
                rows.append({header[i]: cells[i] for i in range(len(header)) if header[i]})
            return rows, os.path.basename(md_p)
        except Exception as e:
            print(f"!! 清单 {stem}.md 解析失败（回退 json）：{e}")
    if os.path.exists(js_p):
        try:
            return json.load(open(js_p, encoding="utf-8")), os.path.basename(js_p)
        except Exception as e:
            print(f"!! 清单 {stem}.json 解析失败（忽略）：{e}")
    return [], None


def split_multi(v):
    """K/L/M 多值列按「；」拆回数组（兼容半角分号、换行）。"""
    v = str(v or "")
    for sep in (";", "\n"):
        v = v.replace(sep, "；")
    return [x.strip() for x in v.split("；") if x.strip()]


# CATS / STRONG_KW 见 classify_filter（统一来源，已在文件头部 import）。
PLACEHOLDER = re.compile(r'(占位|placeholder|待补|README|说明文件|\.txt$|\.url$|\.lnk$)', re.I)

# CATS / STRONG_KW 见 classify_filter（统一来源，已顶部 import）。
# 决定书正文常用简称 → 法规全称（避免把简称误报成「清单缺失」）
ALIAS = {
    "暂行办法": "私募投资基金监督管理暂行办法", "私募办法": "私募投资基金监督管理暂行办法",
    "私募管理办法": "私募投资基金监督管理暂行办法", "私募监管办法": "私募投资基金监督管理暂行办法",
    "私募基金监管办法": "私募投资基金监督管理暂行办法",
    "私募暂行办法": "私募投资基金监督管理暂行办法", "私募条例": "私募投资基金监督管理条例",
    "纪律处分办法": "中国证券投资基金业协会自律管理和纪律处分措施实施办法",
    "若干规定": "关于加强私募投资基金监管的若干规定",
    "加强私募监管规定": "关于加强私募投资基金监管的若干规定",
    # 同一部规定的常见异写（2026-10-27 一年库重建实测：这两种写法各被引用 3~4 次，
    # 不归一会被误报成「06 缺失」，进而诱导重复下载）
    "关于加强私募投资基金监督管理的若干规定": "关于加强私募投资基金监管的若干规定",
    "关于加强私募基金监管的若干规定": "关于加强私募投资基金监管的若干规定",
    "收购办法": "上市公司收购管理办法",
    "信披办法": "私募投资基金信息披露管理办法", "信息披露办法": "私募投资基金信息披露管理办法",
    "适当性办法": "证券期货投资者适当性管理办法", "适当性管理办法": "证券期货投资者适当性管理办法",
    "投资者适当性办法": "证券期货投资者适当性管理办法",
    "募集办法": "私募投资基金募集行为管理办法",
    "登记备案办法": "私募投资基金登记备案办法",
    "内控指引": "私募投资基金管理人内部控制指引", "内部控制指引": "私募投资基金管理人内部控制指引",
    "自律管理和纪律处分措施实施办法": "中国证券投资基金业协会自律管理和纪律处分措施实施办法",
    "自律管理办法": "中国证券投资基金业协会自律管理和纪律处分措施实施办法",
    "廉洁从业规定": "证券期货经营机构及其工作人员廉洁从业规定",
    "私募基金登备办法": "私募投资基金登记备案办法",
    # 2026-10-27 一年库重建实测：被引用 5 次但为已废止/被《私募投资基金登记备案办法》取代的旧版，
    # 归一避免误报「06 缺失」并防止重复下载旧版
    "私募投资基金管理人登记和基金备案办法": "私募投资基金登记备案办法",
    "基金法": "中华人民共和国证券投资基金法", "证券投资基金法": "中华人民共和国证券投资基金法",
    "私募资管办法": "证券期货经营机构私募资产管理业务管理办法",
    "资管运作规定": "证券期货经营机构私募资产管理计划运作管理规定",
    "失联处理指引": "私募基金管理人失联处理指引",
    "登记指引第1号": "私募基金管理人登记指引第1号", "登记指引第2号": "私募基金管理人登记指引第2号",
    "登记指引第3号": "私募基金管理人登记指引第3号",
}
NOISE_LAW = ("管理办法", "暂行规定", "实施细则", "通知", "公告", "意见", "指引", "规定", "条例", "法", "实施办法")
DATE_RE = re.compile(r'(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日')

# 「核心列」敷衍词：出现即视为没按「通读全文、列全」的要求填
PERFUNCTORY = ("详见", "同上", "略", "见决定书", "见原文", "同前")
# 占位值（2026-11-02 补）：上一版只用 PERFUNCTORY + `len(v) < 12` 判敷衍，
# 于是「高管（吴加维）：不明确」这种 13 字符的占位值完美穿透（实测 1 例）。
# 正确做法是**先剥掉「管理人：/高管（X）：」标签**，再看剩余核心文本是不是占位词。
PLACEHOLDER_VAL = frozenset({"不明确", "不详", "未知", "未明确", "待补", "待核", "待定",
                             "无", "暂无", "空", "—", "-", "--", "N/A", "n/a"})
_SUBLABEL = re.compile(r"^(?:管理人|机构|人员|高管)（?[^）]{0,40}）?[：:]\s*")


def flat(t):
    return re.sub(r'\s+', '', t or '')


# recat / rule_classify 见 classify_filter（统一来源，已在文件头部 import）。

# 红通道 OCR 的唯一实现在 ocrlib（2026-10-30 收敛）：dpi / 最大像素 / 缓存键只此一处，
# 并发预填脚本 ocr_prefill.py 与本校验共用，改一处即全体生效（0.5.6）。
import ocrlib

OCR_MAX_PX = ocrlib.OCR_MAX_PX


def ocr_pdf(path, dpi=ocrlib.OCR_DPI, red=True):
    """扫描件内容复核用的 OCR（薄封装，保留旧签名）。实现见 ocrlib.ocr_pdf。"""
    return ocrlib.ocr_pdf(path, dpi=dpi, red=red)


def pick_latest_date(txt, today):
    cands = set()
    for y, m, d in DATE_RE.findall(txt):
        try:
            dt = date(int(y), int(m), int(d))
        except Exception:
            continue
        if date(2015, 1, 1) <= dt <= today:
            cands.add(dt)
    return max(cands).isoformat() if cands else ""


# ---------------------------------------------- OCR 容错的落款日期回现
# 扫描件经 OCR 后落款几乎必然被拆坏，直接比对会全假报警。四种实测噪声形态：
#   A 组件间插标点  B 末尾「日」丢失  C 年份被排到末尾  D 全角/中文数字
_CN_UNIT = {'〇': 0, '零': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9}
_FW_DIG = {chr(0xFF10 + i): str(i) for i in range(10)}          # 全角０-９
_NOISE = re.compile(r"[,，。、;；:：·'\"“”‘’()（）\[\]【】{}<>《》|/\\／_\-\.·\s\$￥#@!?？！～]")
_CN_Y = r"[〇零一二三四五六七八九]{4}"
_CN_MD = r"[〇零一二三四五六七八九十百]{1,3}"


def cn_int(s):
    """中文数字/阿拉伯串 → int。二〇二六→2026；十→10；二十六→26。"""
    s = (s or "").strip()
    if not s:
        return None
    s = "".join(_FW_DIG.get(c, c) for c in s)
    if s.isdigit():
        return int(s)
    if "十" in s:
        a, _, b = s.partition("十")
        hi = 1 if a == "" else (cn_int(a) or 0)
        lo = cn_int(b) if b else 0
        return hi * 10 + lo if hi is not None and lo is not None else None
    v, ok = 0, False
    for ch in s:
        if ch in _CN_UNIT:
            v, ok = v * 10 + _CN_UNIT[ch], True
        else:
            return None
    return v if ok else None


def ocr_date_candidates(txt, tail=1500, today=None):
    """OCR 容错抽取落款日期。返回 (full:set['YYYY-MM-DD'], md:set['MM-DD'])。"""
    t = _NOISE.sub("", flat(txt))
    t = t[-tail:] if tail else t
    today = today or date.today()
    full, md = set(), set()

    def add(y, m, d, cn=False):
        yi, mi, di = (cn_int(y), cn_int(m), cn_int(d)) if cn else \
                     (int(y), int(m), int(d))
        if None in (yi, mi, di):
            return
        try:
            dt = date(yi, mi, di)
        except Exception:
            return
        if date(2015, 1, 1) <= dt <= today:      # 过滤 OCR 编造的未来日期
            full.add(dt.isoformat())

    for y, m, d in re.findall(rf"(20\d{{2}})年(\d{{1,2}})月(\d{{1,2}})日?", t):
        add(y, m, d)
    for m, d, y in re.findall(rf"年(\d{{1,2}})月(\d{{1,2}})日(20\d{{2}})", t):
        add(y, m, d)
    for y, m, d in re.findall(rf"({_CN_Y})年({_CN_MD})月({_CN_MD})日?", t):
        add(y, m, d, cn=True)
    if not full:
        for m, d in re.findall(r"(\d{1,2})月(\d{1,2})日", t):
            mi, di = int(m), int(d)
            if 1 <= mi <= 12 and 1 <= di <= 31:
                md.add(f"{mi:02d}-{di:02d}")
    return full, md


# ---------------------------------------------------------------- 读取台账（v2）
def load_ledger(root):
    """读取 v2 台账。

    返回 (wb, cases, files)：
      cases —— 一行一案例：filenames/paths/urls 为数组（一一对应）；
      files —— 展开成逐文件行（A1 双向核对、B 段逐文件校验用）。
    """
    wb = load_workbook(os.path.join(root, "04_综合索引台账【总表】.xlsx"))
    ws = wb["综合索引台账"]
    cases, files = [], []
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=15, values_only=True), start=2):
        if not any(v is not None and str(v).strip() for v in r[1:]):
            continue                       # 尾部空行（幽灵行占位）
        dv = r[1]
        decided = dv.strftime("%Y-%m-%d") if hasattr(dv, "strftime") else str(dv or "").strip()
        pv = r[2]
        publish = pv.strftime("%Y-%m-%d") if hasattr(pv, "strftime") else str(pv or "").strip()
        fns = split_multi(r[11])
        pts = split_multi(r[12])
        us = split_multi(r[13])
        # L 列（存放路径）**同一案例的文书必然同目录** → 台账只写一条。
        # 内部仍展开成逐文件数组，方便 B 段/A1 逐文件核对（2026-11-02 用户口径）。
        paths_raw = list(pts)
        if fns and len(pts) == 1 and len(fns) > 1:
            pts = pts * len(fns)
        c = {"row": i, "seq": r[0], "decided": decided, "publish": publish,
             "subject": str(r[3] or "").strip(), "manager": str(r[3] or "").strip(), "measures": str(r[4] or "").strip(),
             "grounds": str(r[5] or "").strip(), "evidence": str(r[6] or "").strip(),
             "legal_basis": str(r[7] or "").strip(), "defense": str(r[8] or "").strip(),
             "category": str(r[9] or "").strip(),
             "source": str(r[10] or "").strip(),
             "filenames": fns, "paths": pts, "paths_raw": paths_raw, "urls": us,
             "typical": str(r[14] or "").strip()}
        cases.append(c)
        for j, fn in enumerate(fns):
            files.append({"row": i, "case": c, "idx": j,
                          "filename": fn,
                          "path": pts[j] if j < len(pts) else "",
                          "url": us[j] if j < len(us) else "",
                          "source": c["source"], "decided": decided,
                          "manager": c["manager"]})
    return wb, cases, files


def scan_disk(root):
    disk = []
    for d1 in sorted(os.listdir(root)):
        p = os.path.join(root, d1)
        if not os.path.isdir(p) or d1[:2] not in ("01", "02", "03"):
            continue
        for d2 in sorted(os.listdir(p)):
            q = os.path.join(p, d2)
            if os.path.isdir(q):
                for f in os.listdir(q):
                    if f.lower().endswith(".pdf"):
                        disk.append({"filename": f, "path": d1 + "/" + d2,
                                     "abs": os.path.join(q, f)})
    return disk


# ---------------------------------------------------------------- 在线对账
def http_json(url, ref="https://www.csrc.gov.cn/", retry=2):
    ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
    h = {"User-Agent": "Mozilla/5.0 Chrome/120", "Referer": ref}
    for i in range(retry):
        try:
            return json.loads(urllib.request.urlopen(
                urllib.request.Request(url, headers=h), timeout=30, context=ctx).read()
                .decode("utf-8", "ignore"))
        except Exception:
            if i == retry - 1:
                return None


STRONG = ['私募', '基金管理', '投资基金', '股权投资', '创业投资', '股权投资基金',
          '创业投资基金', '私募基金', '基金份额', '基金财产', '基金管理人']
NOT_PE = ['上市公司', '期货有限公司', '证券有限责任公司', '证券营业部', '会计师事务所',
          '商业银行', '保险公司', '基金销售', '证券投资咨询', '资产评估', '律师事务所',
          '营业部', '支行', '证券公司从业人员', '期货从业人员', '证券基金经营机构']
PE_SAVE = ['私募基金', '私募投资基金', '私募办法', '私募条例', '基金管理人', '基金管理业务', '私募']

INST = ['私募', '基金管理', '投资基金', '股权投资', '创业投资', '资产管理', '投资管理',
        '投资合伙企业', '创业投资合伙企业', '股权投资合伙企业']


def is_private(title, content):
    """按决定书标题+正文判定私募相关性。
    返回 2=强命中（必须入库）、1=弱命中（待核提示）、0=非私募。"""
    t = title or ''
    c = re.sub(r'\s+', '', content or '')
    if any(k in t for k in NOT_PE):
        return 0
    inst = any(k in t for k in INST)
    if not inst:
        return 1 if ('私募基金' in c or '私募投资基金' in c) else 0
    if '股份有限公司' in t and not any(k in c for k in PE_SAVE):
        return 0
    return 2


def platform_items(bureau, info, start, end, kinds=("measures", "penalty")):
    """取某局窗口内全部条目（外部基准），返回 (条目列表, 空正文条目数)。"""
    items, empty = [], 0
    for kind in kinds:
        ch = (info.get(kind) or {}).get("id")
        if not ch:
            continue
        page = 1
        while page <= 200:
            u = (f"https://www.csrc.gov.cn/searchList/{ch}?_isAgg=true&_isJson=true"
                 f"&_pageSize=20&_template=index&_rangeTimeGte={start}&_rangeTimeLte={end}"
                 f"&_channelName=&page={page}")
            j = http_json(u, ref=f"https://www.csrc.gov.cn/{bureau}/")
            if not j or not j.get("data"):
                break
            res = j["data"].get("results") or []
            if not res:
                break
            for blk in res:
                meta = {}
                for g in blk.get("domainMetaList", []):
                    for it in g.get("resultList", []):
                        meta[it.get("key")] = it.get("value", "")
                dt = (meta.get("fwrq") or blk.get("publishedTimeStr") or "")[:10]
                items.append({"title": blk.get("title", ""), "url": blk.get("url", ""),
                              "content": blk.get("content") or "", "date": dt,
                              "wh": meta.get("wh", "")})
            if len(res) < 20:
                break
            page += 1
    for it in items:
        if not (it["content"] or "").strip():
            empty += 1
    return items, empty


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--start", default=None)
    ap.add_argument("--end", default=None)
    ap.add_argument("--days", type=int, default=365,
                    help="库窗口天数（含当日，与 run_update 同口径），默认 365")
    ap.add_argument("--window-by", choices=("decided", "publish"), default="decided",
                    help="窗口筛选口径：decided=处罚时间（落款，默认）；publish=挂网发布日期")
    ap.add_argument("--online", action="store_true", help="联网做平台总数对账")
    # OCR 复核策略（2026-10-30 依用户裁定下调）：扫描件的处罚时间正确性**已在写入期断言**
    # （build_pipeline.date_or_hold：落款抽不到 → 中基协/交易所写「待核」，绝不回落挂网日），
    # 事后每轮再对所有扫描件重跑一遍 OCR，等于同样输入跑第二遍算法，边际收益极低。
    # ⇒ 日常只抽 20 件做"校验器冒烟"（防的不是数据错，是校验器自己判断失效，如 BODY_MARK 撞词）；
    #    `all` 仅在两种场合跑：① 大改造后的术后复查；② 每季度一次全量体检。
    ap.add_argument("--ocr-scan", default="sample:20",
                    help="无文字层（扫描件）PDF 的 OCR 复核策略：none / sample:N / all"
                         "（默认 sample:20；all 用于术后复查或季度体检）")
    ap.add_argument("--ocr-cache", default=None,
                    help="OCR 文本缓存目录（按文件 sha1 命名）")
    ap.add_argument("--json", default=None)
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    # OCR 文本缓存默认落在运维主目录（库父目录/_案例库维护留痕/<库名>/_work/ocr_cache），
    # 不写进交付库；传 --ocr-cache 可覆盖。
    if not a.ocr_cache:
        a.ocr_cache = os.path.join(LS.work_dir(root), "ocr_cache")
    _m = re.match(r"(none|all|sample):?(\d*)", (a.ocr_scan or "sample:20").strip())
    if not _m or _m.group(1) == "none":
        ocr_budget = None
    elif _m.group(1) == "all":
        ocr_budget = 10 ** 9
    else:
        ocr_budget = int(_m.group(2) or 40)
    ocr_used = 0
    ocr_cached_hits = 0
    ocr_skipped = 0
    today = date.today()
    end = a.end or today.isoformat()
    # 窗口口径统一走 winlib（唯一事实源）：旧版用 `date(year-1, ...)`（365/366 天不定），
    # 与 run_update 的 `end-(days-1)` 差 1 天 → verify 判合格的行会被 retire 移出。
    import winlib
    start = a.start or winlib.window(end, a.days)[0]

    wb, cases, files = load_ledger(root)
    OCR_CACHE = a.ocr_cache
    if OCR_CACHE:
        os.makedirs(OCR_CACHE, exist_ok=True)

    def _cache_path(fp):
        return ocrlib.cache_path(fp, OCR_CACHE)

    def ocr_cached(fp):
        if not OCR_CACHE:
            return None
        p = _cache_path(fp)
        if os.path.exists(p):
            try:
                return open(p, encoding="utf-8").read()
            except Exception:
                return None
        return None

    def ocr_save_cache(fp, t):
        if not OCR_CACHE:
            return
        try:
            open(_cache_path(fp), "w", encoding="utf-8").write(t or "")
        except Exception:
            pass

    disk = scan_disk(root)

    # 人工复核豁免清单：B4/B5 命中但已人眼裁图定案的条目不再计入结论。
    MANUAL_OK = {}
    _raw, _mf = _load_list_file(root, "校验人工复核豁免清单")
    if _raw:
        try:
            _bad = 0
            for it in _raw:
                if (it.get("filename") and it.get("check")
                        and it.get("evidence") and it.get("verified_on")):
                    MANUAL_OK[(it["filename"], it["check"])] = it
                elif it.get("filename"):
                    _bad += 1
            if _bad:
                print(f"!! 豁免清单 {_bad} 条缺少 evidence/verified_on，已丢弃")
        except Exception as e:
            print(f"!! 豁免清单解析失败（忽略）：{e}")
    rep = {"window": [start, end], "window_by": a.window_by,
           "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "A": {}, "B": {}, "C": {}, "D": {}}
    problems = []

    print("=" * 72)
    print(f"私募监管处罚案例库 · 校验报告（v2）   窗口 {start} ~ {end}"
          f"（口径：{'处罚时间' if a.window_by == 'decided' else '挂网发布日期'}）")
    print("=" * 72)

    # ---------------- A 完整性 ----------------
    print("\n【A 完整性】")
    # A1 台账 K/L 列 ↔ 磁盘 双向核对（一行=一个案例，K 列显式登记全部决定书）
    led_keys = {(f["filename"], f["path"]) for f in files if f["path"]}
    led_nopath = [f for f in files if not f["path"]]
    disk_keys = {(d["filename"], d["path"]) for d in disk}
    orphan = [d for d in disk if (d["filename"], d["path"]) not in led_keys]
    missing = [f for f in files if f["path"] and (f["filename"], f["path"]) not in disk_keys]
    print(f"  A1 案例 {len(cases)} 行（决定书文件 {len(files)} 份）/ 磁盘 {len(disk)} 个 PDF"
          f" ｜ 孤儿 {len(orphan)} ｜ 缺失 {len(missing)} ｜ 未登记路径 {len(led_nopath)}")
    rep["A"]["cases"] = len(cases); rep["A"]["files"] = len(files)
    rep["A"]["disk_pdfs"] = len(disk)
    rep["A"]["orphan"] = orphan[:50]; rep["A"]["missing"] = [
        {"row": f["row"], "filename": f["filename"], "path": f["path"]} for f in missing[:50]]
    if orphan or missing or led_nopath:
        problems.append("A1 台账与磁盘未对齐")
        for o in orphan[:10]:
            print(f"      孤儿（磁盘有、台账未登记）：{o['path'][:34]}/{o['filename'][:44]}")
        for m in missing[:10]:
            print(f"      缺失（台账有、磁盘无）：{m['path'][:30]}/{m['filename'][:44]}（台账第{m['row']}行）")
        for f in led_nopath[:10]:
            print(f"      未登记路径：第{f['row']}行 {f['filename'][:44]}")

    # A1b 多值列对齐（2026-11-02 口径修正）：
    #   K 案例文件 —— 必须与文书数一致（N 条）；
    #   L 存放路径 —— 同一案例同目录，**只写 1 条**（兼容旧库的 N 条重复写法）；
    #   M 原文链接 —— 每个文书一条，允许 1 条覆盖全案（同一页面打包下载）。
    misalign = [c for c in cases
                if not c["filenames"]
                or len(c.get("paths_raw") or []) not in (1, len(c["filenames"]))
                or len(c["urls"]) not in (1, len(c["filenames"]))]
    print(f"  A1b 案例文件/路径/链接 多值列不对齐：{len(misalign)}")
    rep["A"]["misaligned"] = [{"row": c["row"], "fn": len(c["filenames"]),
                               "pt": len(c.get("paths_raw") or []), "url": len(c["urls"])}
                              for c in misalign[:40]]
    if misalign:
        problems.append("A1b 案例文件/路径/链接数目不对齐（K=N、L=1（同目录）、M=N 或 1）")
        for c in misalign[:8]:
            print(f"      第{c['row']}行 文件{len(c['filenames'])} / 路径{len(c.get('paths_raw') or [])}"
                  f" / 链接{len(c['urls'])}")
    # A1c 存放路径重复写（同一案例把同一个目录写了 N 遍）——结构冗余，台账可读性差
    dup_paths = [c for c in cases
                 if len(c.get("paths_raw") or []) > 1
                 and len(set(c["paths_raw"])) == 1]
    print(f"  A1c 存放路径按文件数重复填写：{len(dup_paths)}")
    rep["A"]["dup_paths"] = [{"row": c["row"], "n": len(c["paths_raw"])} for c in dup_paths[:40]]
    for c in dup_paths[:6]:
        print(f"      行{c['row']} 同一路径写了 {len(c['paths_raw'])} 遍")
    if dup_paths:
        problems.append(f"A1c 有 {len(dup_paths)} 行把同一存放路径重复写了 N 遍（应只写 1 条）")

    # A2 窗口检查（默认按处罚时间；--window-by publish 时按挂网发布日期）
    def _wkey(c):
        return c["decided"] if a.window_by == "decided" else c["publish"]

    def _ok_date(s):
        return bool(re.match(r"^\d{4}-\d{2}-\d{2}$", s or ""))

    outw, soften = [], []
    for c in cases:
        k = _wkey(c)
        if not k or k == "待补":
            continue
        if start <= k <= end:
            continue
        # 2026-10-29 修：区分「双超窗」与「仅落款超窗、挂网仍在窗内」。
        # 保留窗口（按落款）比抓取窗口（按挂网）更窄，证监局文书又普遍滞后刊发 9~23 天，
        # 于是本窗口刚收的案例会因落款早 1~3 天被判超窗出库，下一轮又被抓回来 —— 震荡。
        # ⇒ 只有落款、挂网双双出窗才判 problems；仅落款出窗的记观察项并保留在库。
        p = c["publish"]
        if (a.window_by == "decided" and _ok_date(p) and start <= p <= end and k < start):
            soften.append(c)
        else:
            outw.append(c)

    skipped_win = [c for c in cases
                   if a.window_by == "publish"
                   and (not c["publish"] or c["publish"] == "待补")]
    # 2026-10-28 修：原写成 `_wkey(cases[0]) and ''`（一段恒为空串的死代码），
    # 空库时 cases 为空 → IndexError，A2 之后的 B/C/D 全部没跑，报告残缺。
    # 而"空库"恰恰是他人首次使用技能的必经状态，也是定时任务首跑的状态。
    print(f"  A2 窗口外（按{a.window_by}）：{len(outw)}"
          + (f" ｜ 仅落款出窗、挂网在窗内（观察项，保留）{len(soften)}" if soften else "")
          + (f" ｜ 发布日期待补无法核窗口 {len(skipped_win)} 条" if skipped_win else ""))
    rep["A"]["out_of_window"] = [{"row": c["row"], "date": _wkey(c),
                                  "manager": c["manager"][:30]} for c in outw[:50]]
    rep["A"]["out_of_window_soft"] = [{"row": c["row"], "decided": c["decided"],
                                       "publish": c["publish"],
                                       "manager": c["manager"][:30]} for c in soften[:50]]
    if outw:
        problems.append("A2 存在窗口外记录")
        for c in outw[:10]:
            print(f"      {_wkey(c)} {c['manager'][:30]}")

    # A3 判重（v2 无文号）：案例键 = (来源, 管理人名称, 处罚时间)；文件重复也算
    # 2026-10-29 修正：文件身份 = (存放路径, 文件名)，与「定位文件=目录+basename」
    # 的存放约定一致。同一家企业常同时被中基协自律处分 + 属地证监局行政处罚，
    # 两份文书都按「机构名.pdf」命名、落在不同来源目录下 —— 这是**两个真实文件**，
    # 不是重复。旧版只按 basename 判重，把这类正常情形全报成 A3 problems（实测 17 例误报）。
    dup_case = [f"{s} {m} [{d}]" for (s, m, d), v in
                Counter((c["source"], c["manager"], c["decided"])
                        for c in cases if c["manager"] and c["manager"] not in UNKNOWN).items()
                if v > 1]
    seen_fn = defaultdict(list)
    for f in files:
        seen_fn[(f["path"], f["filename"])].append(f["row"])
    dup_fn = {k: v for k, v in seen_fn.items() if len(v) > 1}
    # 同名不同目录：仅作观察项（跨来源同名文书），不进 problems
    _by_name = defaultdict(set)
    for f in files:
        _by_name[f["filename"]].add(f["path"])
    cross_dir = {k: v for k, v in _by_name.items() if len(v) > 1}
    print(f"  A3 重复：同案例键（局+管理人+处罚时间）{len(dup_case)} 组 ｜"
          f" 同目录跨行重复文件 {len(dup_fn)} 个 ｜"
          f" 同名不同目录（跨来源，正常）{len(cross_dir)} 个")
    rep["A"]["dup_case"] = dup_case[:30]
    rep["A"]["dup_filename"] = {"%s/%s" % k: v for k, v in list(dup_fn.items())[:30]}
    rep["A"]["same_name_cross_dir"] = len(cross_dir)
    if dup_case:
        problems.append("A3 同案例键重复入库")
        for k in dup_case[:10]:
            print(f"      案例重复：{k}")
    if dup_fn:
        problems.append("A3 同一文件登记在多行（一份决定书只能属于一个案例）")
        for k, v in list(dup_fn.items())[:8]:
            print(f"      文件重复：{k[1][:44]} @ {k[0][-18:]} → 行 {v}")

    # A3b 同案拆行（2026-11-02 新增）：判重键按 v2 是 **(局, 管理人, 处罚时间)**，
    # 语义是「一家机构同一天的一个处罚事件 = 一行，机构与高管合并」。但历史落地把键写成了
    # 「D 列整串（机构+人员）」，于是同一案的机构决定书与某高管决定书（人员名单不同）
    # 各占一行。实测 3 组：券期资本（告知书〔2026〕75号）、杭州合信（〔2025〕433号）、
    # 浙江优策（〔2025〕356号），三对的**事先告知书文号完全相同**，逐份 OCR 已证实同案。
    # 主体名一被「曾登记为/实际负责」污染，同机构就分成了两个"不同机构"，这类拆行更隐蔽。
    # 故这里按「剥角色前缀后的机构名 + 落款日 + 来源」判，比 A3 的整串键更贴近 v2 口径。
    from subject_clean import strip_role_prefix as _srp
    _grp = defaultdict(list)
    for c in cases:
        if not c["manager"] or c["manager"] in UNKNOWN:
            continue
        _org = _srp(re.split(r"[；;]", c["manager"])[0].strip())
        _grp[(c["source"], _org, c["decided"])].append(c["row"])
    dup_case2 = [(k, v) for k, v in _grp.items() if len(v) > 1]
    print(f"  A3b 同案拆行（同来源+同机构+同落款日 多行）：{len(dup_case2)} 组")
    rep["A"]["dup_case_org"] = [{"key": list(k), "rows": v} for k, v in dup_case2[:30]]
    if dup_case2:
        problems.append("A3b 同一案例被拆成多行（v2 口径：机构同日落款应合并为一行）")
        for k, v in dup_case2[:10]:
            print(f"      拆行：{k[0]} {k[2]} {k[1][:30]} → 行 {v}")

    src_cnt = Counter(c["source"] for c in cases)
    amac = sum(v for k, v in src_cnt.items() if k == "中基协")
    exch = sum(v for k, v in src_cnt.items() if any(x in k for x in ("交易所",)))
    csrc = len(cases) - amac - exch
    print(f"  A5 三信源覆盖：中基协 {amac} ｜ 证监局 {csrc} ｜ 交易所 {exch}")
    rep["A"]["sources"] = {"中基协": amac, "证监局": csrc, "交易所": exch}
    if amac == 0 or csrc == 0:
        problems.append("A5 信源覆盖不全")

    if a.online:
        # 外部基准：逐局拉窗口内全部条目 → 判私募相关 → 与库内比对。
        # v2 无文号列：比对通道 = ①原文链接（库内 M 列）②管理人名称（库内 D 列）。
        lib_u = set()
        lib_mg_full = defaultdict(set)     # 局 -> {管理人全称（去空白）}
        lib_mg_t12 = defaultdict(set)      # 局 -> {管理人名去括号后前 12 字}
        for c in cases:
            if c["source"] == "中基协" or "交易所" in c["source"]:
                continue
            b = c["source"].replace("证监局", "")
            for u in c["urls"]:
                lib_u.add(str(u).replace("https:", "").lstrip("/"))
            mg = re.sub(r"\s+", "", c["manager"])
            if mg and mg not in UNKNOWN:
                lib_mg_full[b].add(mg)
                t12 = re.sub(r"[（(].*", "", mg).strip()[:12]
                if len(t12) >= 6:
                    lib_mg_t12[b].add(t12)
        # 不收录清单豁免：URL / 文号（清单仍保留文号列）/ 主体名三通道
        excl_u, excl_n = set(), set()
        excl_tgf, excl_tg12 = set(), set()
        excl_items = []

        def _norm_docno(s):
            s = re.sub(r"\s+", "", str(s or ""))
            m = re.search(r"[〔\[（(【〖]\s*(20\d{2})\s*[〕\]）)】〗]\s*(\d+)\s*号", s)
            return s, (m.group(0) if m else "")

        def _norm_target(s):
            f = re.sub(r"\s+", "", str(s or ""))
            t = re.sub(r"[（(].*", "", f).strip()[:12]
            return f, t

        _excl_rows, excl_f = _load_list_file(root, "不收录清单")
        if _excl_rows:
            try:
                for x in _excl_rows:
                    if x.get("url"):
                        excl_u.add(str(x["url"]).replace("https:", "").lstrip("/"))
                    raw, short = _norm_docno(x.get("doc_no"))
                    if raw:
                        excl_n.add(raw)
                    if short:
                        excl_n.add(short)
                    tf, t12 = _norm_target(x.get("target"))
                    if tf:
                        excl_tgf.add(tf)
                        if len(t12) >= 10:
                            excl_tg12.add(t12)
                    excl_items.append(x)
            except Exception:
                pass
        gaps, empty_total, plat_total = [], 0, 0
        skipped = []
        for b, info in CHANNELS.items():
            cn = BN.get(b, b)
            items, empty = platform_items(b, info, start, end)
            empty_total += empty
            scored = [(is_private(it["title"], it["content"]), it) for it in items]
            cand = [it for lv, it in scored if lv == 2]
            weak = [it for lv, it in scored if lv == 1]
            plat_total += len(cand)
            miss, weak_miss = [], []
            for it in cand:
                if not (start <= (it["date"] or "") <= end):
                    continue      # 平台按挂网时间返回，决定日期在窗口外的不算缺口
                tm = re.search(r"关于对(.+?)(?:采取|予以|实施)", it["title"] or "")
                tgf = re.sub(r"\s+", "", tm.group(1)) if tm else ""
                tg = re.sub(r"[（(].*", "", tgf).strip()[:12] if tgf else ""
                n = re.search(r"[〔\[（(【〖]\s*(20\d{2})\s*[〕\]）)】〗]\s*(\d+)\s*号",
                              re.sub(r"\s+", "", it["wh"] or ""))
                no = n.group(0) if n else ""
                # 库内命中：URL / 管理人全称 / 管理人前 12 字
                in_lib = (it["url"] and str(it["url"]).replace("https:", "").lstrip("/") in lib_u) \
                    or (tgf and tgf in lib_mg_full[cn]) \
                    or (tg and len(tg) >= 6 and tg in lib_mg_t12[cn])
                if in_lib:
                    continue
                ukey = str(it["url"] or "").replace("https:", "").lstrip("/")
                ex_hit = (ukey in excl_u or (no and no in excl_n)
                          or (tgf and tgf in excl_tgf)
                          or (tg and len(tg) >= 10 and tg in excl_tg12))
                if ex_hit:
                    skipped.append({"bureau": cn, "date": it["date"],
                                    "title": (it["title"] or "")[:56],
                                    "reason": next((e.get("reason", "") for e in excl_items
                                                    if str(e.get("url", "")).replace("https:", "")
                                                    .lstrip("/") == ukey
                                                    or _norm_target(e.get("target"))[0] == tgf),
                                                   "已判定不收录")})
                    continue
                miss.append({"date": it["date"], "title": (it["title"] or "")[:56],
                             "url": it["url"], "no_content": not (it["content"] or "").strip()})
            for it in weak:
                if not (start <= (it["date"] or "") <= end):
                    continue
                tm = re.search(r"关于对(.+?)(?:采取|予以|实施)", it["title"] or "")
                tg = re.sub(r"[（(].*", "", tm.group(1)).strip()[:12] if tm else ""
                if tg and tg in lib_mg_t12[cn]:
                    continue
                weak_miss.append({"date": it["date"], "title": (it["title"] or "")[:56],
                                  "url": it["url"]})
            if miss or weak_miss:
                # 列表有封顶，必须另记真值（不能用 len(切片) 当计数）
                gaps.append({"bureau": cn, "platform_private": len(cand),
                             "missing": miss[:25], "weak": weak_miss[:10],
                             "missing_total": len(miss), "weak_total": len(weak_miss)})
        n_miss = sum(g["missing_total"] for g in gaps)
        n_weak = sum(g["weak_total"] for g in gaps)
        print(f"  A6 平台窗口内「私募机构」条目 {plat_total} 条（36 局全量比对）｜"
              f"强缺口 {n_miss} 条 ｜ 待核提示 {n_weak} 条")
        print(f"      其中平台正文为空（附件型决定书）共 {empty_total} 条 —— "
              f"这批必须走附件补抓，否则必然漏")
        if skipped:
            print(f"      已判定不收录 {len(skipped)} 条"
                  f"（见案例库根目录「{excl_f or '不收录清单.md'}」）：")
            for s in skipped[:8]:
                print(f"        × {s['bureau']} {s['date']} {s['title'][:44]} —— {s['reason'][:30]}")
        rep["A"]["platform_excluded"] = skipped
        rep["A"]["platform_private"] = plat_total
        rep["A"]["platform_empty_content"] = empty_total
        rep["A"]["platform_gap"] = gaps
        for g in sorted(gaps, key=lambda x: -len(x["missing"]))[:8]:
            print(f"      缺口 {g['bureau']}：强 {len(g['missing'])} 条 / 待核 "
                  f"{len(g.get('weak', []))}（平台私募候选 {g['platform_private']}）")
        for g in sorted(gaps, key=lambda x: -len(x["missing"]))[:3]:
            for m in g["missing"][:5]:
                print(f"        · {m['date']} {m['title'][:48]}"
                      f"{' [附件型]' if m['no_content'] else ''}")
        if n_miss:
            problems.append(f"A6 平台有 {n_miss} 条私募机构决定书未入库")
    else:
        print("  A6 平台对账：未开启（加 --online 逐局拉取窗口内全量条目做外部基准比对）")

    # ---------------- B 内容正确性 ----------------
    print("\n【B 内容正确性】")
    b_empty, b_nolayer, b_date, b_tgt, b_name, b_place = [], [], [], [], [], []
    date_tolerant, date_weak = [], []
    date_person = []       # v2.1 高管决定书落款与案例处罚时间差 ≤60 天（同批文书，放行）
    date_manual, tgt_manual, name_manual = [], [], []
    b_nobody = []
    b_pub, b_lag, b_zero = [], [], []
    b_core = []            # B10 三核心列缺失/敷衍
    scanned_fns = set()    # 无文字层（扫描件）文件名集合：其名称回现检查以 OCR 为准、文字层检查不适用
    # 复核决定书文体不含「经查/当事人」，缺了会把复核决定书误判成空壳（2026-09-28 实测）
    BODY_MARK = ("经查", "经查明", "当事人", "违反", "违规事实", "我局决定",
                 "申请人", "本会", "复核")
    scanned = 0
    # 行级「日期已可信」标记（2026-10-30 用户裁定）：**复核决定才是最终决定**，
    # 台账的处罚时间与处罚措施都记最终复核结果。于是同一行天然含多份落款不同的文书
    # （原处分决定 4-30 → 复核决定 8-25 维持），逐份比对必然把**前序文书**判成"落款未回现"。
    # 正确口径：只要该行**有任一份**文书的落款 = 台账处罚时间（且它就是最终决定的落款），
    # 这一行的日期即成立；其余文书是"前序文书"（告知/原处分/复核前），只作观察项。
    row_date_hit = {}
    for f in files:
        r = f
        fp = os.path.join(root, f["path"], f["filename"]) if f["path"] else ""
        if PLACEHOLDER.search(f["filename"]):
            b_place.append({"row": f["row"], "filename": f["filename"]})
            continue
        if not fp or not os.path.exists(fp):
            continue
        try:
            if open(fp, "rb").read(5) != b"%PDF-":
                b_place.append({"row": f["row"], "filename": f["filename"]}); continue
            d = fitz.open(fp)
            npages = d.page_count
            txt = "".join(p.get_text() for p in d)
            d.close()
        except Exception as e:
            b_place.append({"row": f["row"], "filename": f["filename"], "err": str(e)[:60]})
            continue
        if npages < 1 or os.path.getsize(fp) < 3000:
            b_empty.append({"row": f["row"], "filename": f["filename"]})
            continue
        body = flat(txt)
        if len(body) < 60:
            scanned += 1
            scanned_fns.add(f["filename"])
            b_nolayer.append({"row": f["row"], "filename": f["filename"],
                              "source": f["source"]})
            if ocr_budget is None or ocr_used >= ocr_budget:
                ocr_skipped += 1
                continue
            try:
                cached = ocr_cached(fp)
                if cached is not None:
                    t2 = cached
                    ocr_cached_hits += 1
                else:
                    ocr_used += 1
                    t2 = ocr_pdf(fp)
                    ocr_save_cache(fp, t2)
                body = flat(t2)
                txt = txt + "\n" + t2
            except Exception:
                continue
            if len(body) < 60:
                continue
        # 有文字层却不含任何决定书正文标记 → 空壳 PDF（正文在附件里）
        if not any(k in body for k in BODY_MARK):
            b_nobody.append({"row": f["row"], "source": f["source"], "len": len(body),
                             "filename": f["filename"][:44], "url": f["url"]})
            continue
        # B4 落款回现：台账处罚时间必须以「YYYY年M月D日」形式出现在决定书正文里
        # v2.1（2026-10-26 试点实测）：同一案例内**高管决定书落款日可与机构决定书不同**
        # （实测：上海隐上机构决定书 7-17、其三份高管决定书 7-1，处分文号非按签发日排序）。
        # 规则：机构决定书（文件名==管理人全称.pdf）严格比对案例处罚时间；
        #       高管决定书放宽为「正文最晚日期与案例处罚时间相差 ≤60 天」即放行（单列观察项）。
        mg_c = re.sub(r"\s+", "", r["manager"])
        mg_core = re.split(r"[（(]", mg_c)[0].strip() if mg_c and mg_c not in UNKNOWN else ""
        is_inst = bool(mg_core) and f["filename"] == mg_core + ".pdf"
        if r["decided"] and len(r["decided"]) == 10:
            y, m, d2 = r["decided"].split("-")
            pats = [f"{y}年{int(m)}月{int(d2)}日", f"{y}年{m}月{d2}日",
                    f"{y}年{int(m)}月{d2}日", f"{y}年{m}月{int(d2)}日",
                    f"{y}年{int(m):02d}月{int(d2):02d}日"]
            if not any(p in body for p in pats):
                tail_src = txt if len(txt) > len(body) else (txt + "\n" + body)
                full, md = ocr_date_candidates(tail_src)
                if r["decided"] in full:
                    date_tolerant.append({"row": f["row"], "ledger": r["decided"],
                                          "filename": f["filename"][:40]})
                    row_date_hit[f["row"]] = True
                elif f"{int(m):02d}-{int(d2):02d}" in md:
                    date_weak.append({"row": f["row"], "ledger": r["decided"],
                                      "filename": f["filename"][:40]})
                    row_date_hit[f["row"]] = True
                elif not is_inst:
                    # 高管决定书：正文最晚日期与案例处罚时间 ≤60 天即视为同批文书
                    td = pick_latest_date(txt, today)
                    try:
                        diff = abs((date.fromisoformat(td) - date.fromisoformat(r["decided"])).days) if td else 999
                    except Exception:
                        diff = 999
                    if td and diff <= 60:
                        date_person.append({"row": f["row"], "ledger": r["decided"],
                                            "file_date": td, "diff_days": diff,
                                            "filename": f["filename"][:40]})
                        row_date_hit[f["row"]] = True
                    else:
                        ex = MANUAL_OK.get((f["filename"], "B4"))
                        if ex:
                            date_manual.append({"row": f["row"], "ledger": r["decided"],
                                                "filename": f["filename"][:40],
                                                "note": str(ex.get("evidence"))[:80],
                                                "on": ex.get("verified_on")})
                        else:
                            ld = pick_latest_date(txt, today)
                            b_date.append({"row": f["row"], "ledger": r["decided"],
                                           "pdf_latest_date": ld, "filename": f["filename"][:40],
                                           "source": f["source"],
                                           "ocr_full": sorted(full)[-3:], "ocr_md": sorted(md)[-3:]})
                else:
                    ex = MANUAL_OK.get((f["filename"], "B4"))
                    if ex:
                        date_manual.append({"row": f["row"], "ledger": r["decided"],
                                            "filename": f["filename"][:40],
                                            "note": str(ex.get("evidence"))[:80],
                                            "on": ex.get("verified_on")})
                    else:
                        ld = pick_latest_date(txt, today)
                        b_date.append({"row": f["row"], "ledger": r["decided"],
                                       "pdf_latest_date": ld, "filename": f["filename"][:40],
                                       "source": f["source"],
                                       "ocr_full": sorted(full)[-3:], "ocr_md": sorted(md)[-3:]})
            else:
                row_date_hit[f["row"]] = True      # 正文原样含台账落款 → 本行日期成立
        # B5 名称回现（v2 主体列）：文件名（不带.pdf）首段（机构名或人员名）应在本 PDF 正文回现。
        # 人员文件形如「姓名-机构名.pdf」，取「姓名」段核对；机构文件（机构名.pdf）取机构段核对。
        fn = f["filename"].replace(".pdf", "")
        # 2026-10-27 修复：文件名里的「（2）/（3）…」是同名主体的去重序号后缀，
        # 决定书正文绝不会出现该后缀 → 比对前先剥掉，否则会把
        # 「深圳市券期资本管理有限公司（2）」误判成「管理人名称未在正文回现」
        # （实测 22 条里 18 条属此类良性假报警，非数据问题）。
        _strip_dup = lambda s: re.sub(r"[（(]\s*\d+\s*[）)]$", "", s).strip()
        # 2026-11-02 修复：文件名尾部的「（2026-09-30）」「（2）」「（原决定）」「（复核决定）」
        # 等语义后缀要**先整体剥掉再按 - 切段**。否则「厦门市鑫鼎盛控股有限公司（2026-09-30）」
        # 会被按日期里的连字符切成 [「…（2026」「09」「30）」]，取末段得「30）」→
        # 报出「管理人名称未在正文回现」这种纯自造的假报警（实测 1 条）。
        # 只剥「含数字 / 原决定 / 复核决定」的尾巴，**不动**「（有限合伙）」这类主体名组成部分。
        _strip_tail = lambda s: re.sub(
            r"[（(](?:[^（）()]*\d[^（）()]*|原决定|复核决定)[）)]$", "", s).strip()
        fn_nt = _strip_tail(fn)
        fn_core = _strip_dup(fn_nt)
        segs = [_strip_dup(s) for s in fn_nt.split("-") if s.strip()]
        # 机构文件：「机构名.pdf」单段；高管责任文件：「职务姓名-机构名.pdf」多段。
        # 高管文件只要求文件名任一主体片段（机构名必然在正文回现）出现即可，
        # 避免「职务+姓名」整串未逐字出现造成的良性假报警（2026-10-27 实测 34 条
        # 中约 26 条为此类，属文件名规范内的差异，非数据问题）。
        cands = segs if len(segs) > 1 else [fn_core]
        # 2026-10-28 修复：比对前归一化，消除三类良性假报警（实测 4 条假报警全部命中）：
        #  ① 全角（）vs 半角 ()：正文「元德坤私募基金管理(南京)有限公司」用半角，
        #     台账/文件名用全角 → 原样 in 比对必然 False；
        #  ② 不间断空格 \xa0 / 全角空格：正文「合肥水木信保…中心\xa0(有限合伙)」；
        #  ③ 扫描件 body 已是 OCR 文本，OCR 常把括号/空格识别成异体。
        # 归一化后再比，既不放过真实遗漏，也不误报格式差异。
        def _norm(s):
            s = str(s).replace("\u3000", "").replace("\xa0", "")
            s = re.sub(r"\s+", "", s)
            return s.replace("（", "(").replace("）", ")")
        if any(_norm(c) in _norm(body) for c in cands):
            continue
        # 无文字层扫描件：文件名由主体派生（文件名==主体即天然保证归属），
        # 名称回现应以 OCR 文本为准，纯文字层检查在此不适用 → 不报问题（避免冗余误报）。
        if f["filename"] in scanned_fns:
            continue
        head = segs[-1] if len(segs) > 1 else segs[0]
        if head and len(head) >= 2:
            ex = MANUAL_OK.get((f["filename"], "B5"))
            if ex:
                tgt_manual.append({"row": f["row"], "target": head[:30],
                                   "note": str(ex.get("evidence"))[:80]})
            else:
                b_tgt.append({"row": f["row"], "target": head[:30],
                              "filename": f["filename"][:40]})
        # B6 文件命名规范（v2）：文件名必须含本行【机构】名称核心。
        # D 列格式「机构；人员」——比对只取分号前的机构部分；
        # 人员文件 = 姓名-机构名.pdf 同样含机构，属规范内差异，不计入问题。
        mg = re.sub(r"\s+", "", r["manager"])
        if mg and mg not in UNKNOWN:
            mg_inst = re.split(r"[；;]", mg)[0]
            mg_core = re.split(r"[（(]", mg_inst)[0].strip()
            if mg_core and mg_core not in f["filename"]:
                ex6 = MANUAL_OK.get((f["filename"], "B6"))
                if ex6:
                    name_manual.append({"row": f["row"], "filename": f["filename"][:44],
                                       "note": str(ex6.get("evidence"))[:80]})
                else:
                    b_name.append({"row": f["row"], "filename": f["filename"][:44],
                                   "manager": mg_core[:24]})
    # B8/B9 日期逻辑：处罚时间（落款）≤ 挂网发布日期（台账 C 列）
    for c in cases:
        if not c["decided"] or len(c["decided"]) != 10:
            continue
        pub = c["publish"]
        if not pub or pub == "待补":
            continue
        try:
            lag = (date.fromisoformat(pub) - date.fromisoformat(c["decided"])).days
        except Exception:
            continue
        if lag < 0:
            b_pub.append({"row": c["row"], "decided": c["decided"], "publish": pub,
                          "manager": c["manager"][:20]})
        elif lag > 60:
            b_lag.append({"row": c["row"], "decided": c["decided"], "publish": pub,
                          "days": lag, "manager": c["manager"][:20]})
        elif lag == 0 and c["source"] == "中基协":
            # 中基协落款→挂网滞后 5~149 天，同日即疑似把发布日复制成了落款
            b_zero.append({"row": c["row"], "decided": c["decided"], "publish": pub,
                           "manager": c["manager"][:20]})
    # B10 三核心列完整性：F/G/H/E 禁止留空或敷衍
    for c in cases:
        for col, key in (("核心处罚事由", "grounds"), ("核心证据材料", "evidence"),
                         ("核心处罚依据", "legal_basis"), ("处罚措施", "measures")):
            v = (c[key] or "").strip()
            if not v:
                b_core.append({"row": c["row"], "col": col, "issue": "留空",
                               "manager": c["manager"][:24]})
                continue
            # 剥掉「管理人：/高管（X）：」子标签后再判占位（2026-11-02 修）
            core = _SUBLABEL.sub("", v).strip()
            core = re.sub(r"^[；;，,\s]+", "", core)
            if core in PLACEHOLDER_VAL:
                b_core.append({"row": c["row"], "col": col, "issue": "占位（" + core + "）",
                               "manager": c["manager"][:24]})
            elif len(core) < 12 and any(w in core for w in PERFUNCTORY):
                b_core.append({"row": c["row"], "col": col, "issue": "敷衍（" + core[:10] + "）",
                               "manager": c["manager"][:24]})
    n_pub_pending = sum(1 for c in cases if not c["publish"] or c["publish"] == "待补")
    print(f"  B1 非PDF/坏文件/占位：{len(b_place)}")
    print(f"  B2 页数或体积异常：{len(b_empty)} ｜ 无文字层（扫描件）：{len(b_nolayer)}")
    print(f"  B4 处罚时间与正文落款不一致：{len(b_date)}"
          f"（OCR 容错命中 {len(date_tolerant)}、仅月日回现 {len(date_weak)}、"
          f"高管决定书落款差异≤60天 {len(date_person)}、人工豁免 {len(date_manual)}，均判为通过）")
    print(f"  B5 管理人名称未在正文回现：{len(b_tgt)}"
          f"{'（人工豁免 ' + str(len(tgt_manual)) + '）' if tgt_manual else ''}")
    print(f"  B6 文件名未含本案例管理人名称：{len(b_name)}"
          f"{'（人工豁免 ' + str(len(name_manual)) + '）' if name_manual else ''}")
    print(f"  B7 正文缺失（只有导航框架的空壳 PDF）：{len(b_nobody)}"
          f"{' —— 需走附件补抓' if b_nobody else ''}")
    print(f"  B8 处罚时间晚于挂网发布日期（逻辑错误）：{len(b_pub)}"
          f" ｜ 滞后 >60 天（刊发滞后，仅提示）：{len(b_lag)}")
    print(f"  B9 中基协处罚时间 == 挂网日（复制粘贴预警）：{len(b_zero)}")
    print(f"  B10 三核心列留空/敷衍：{len(b_core)}"
          f" ｜ 发布日期待补：{n_pub_pending}（观察项，不进 problems）")
    rep["B"] = {"bad_file": b_place[:40], "empty": b_empty[:40],
                "no_text_layer": len(b_nolayer), "no_body": b_nobody[:40],
                "date_mismatch": b_date[:40], "date_tolerant": len(date_tolerant),
                "date_weak": len(date_weak), "date_person_ok": date_person[:20],
                "date_manual_ok": date_manual, "target_manual_ok": tgt_manual,
                "manager_not_in_pdf": b_tgt[:40], "filename_mismatch": b_name[:40],
                "date_after_publish": b_pub[:40], "publish_lag_60d": b_lag[:40],
                "date_equals_publish": b_zero[:200],
                "core_missing": b_core[:40], "publish_pending": n_pub_pending}
    if b_pub:
        for x in b_pub[:8]:
            print(f"      日期晚于发布：行{x['row']} 处罚{x['decided']} > 挂网{x['publish']}"
                  f" | {x['manager']}")
        problems.append(f"B8 处罚时间晚于挂网发布日期 {len(b_pub)} 条（逻辑不可能，需复核）")
    if len(b_zero) >= 5:
        for x in b_zero[:8]:
            print(f"      日期==挂网日：行{x['row']} 处罚{x['decided']} = 挂网{x['publish']}"
                  f" | {x['manager']}")
        problems.append(f"B9 中基协处罚时间 == 挂网日 {len(b_zero)} 条"
                        f"（疑似把发布日复制成了落款，需 OCR 复核落款）")
    if b_nobody:
        for x in b_nobody[:8]:
            print(f"      空壳：行{x['row']} [{x['source']}] {x['filename'][:44]}")
        problems.append(f"B7 正文缺失 {len(b_nobody)} 条（需附件补抓）")
    # 2026-10-27 口径修正：证监局网页版决定书**正文通常不印落款日期**（只给文号），
    # 台账处罚时间取平台发文日期字段，PDF 里自然找不到 → 属已知局限（观察项），
    # 只有"正文里另有落款且与台账不一致"才是真错误。
    b_date_nosign = [x for x in b_date if not x.get("pdf_latest_date")]
    # 证监局网页版决定书普遍不印落款（正文里的日期多为检查时间/整改期限），
    # 其处罚时间取平台发文日期字段，与正文任何日期比对都可能"不一致"→ 一律观察项。
    b_date_real = [x for x in b_date if x.get("pdf_latest_date")
                   and x.get("source") != "证监局"]
    b_date_web = [x for x in b_date if x.get("pdf_latest_date")
                  and x.get("source") == "证监局"]
    if b_date_web:
        print(f"      （证监局网页版无落款 {len(b_date_web)} 条：正文日期为检查/整改等非落款日期，"
              f"台账取平台发文日期 → 观察项，不计入结论）")
    if b_date_nosign:
        print(f"      （正文/PDF 未载明落款日期 {len(b_date_nosign)} 条：证监局网页版决定书普遍不印落款，"
              f"台账取平台发文日期字段 → 观察项，不计入结论）")
    # 2026-10-29 修正：1 天以内的落款近差，本质是 OCR/抽取把 PDF 里某个相邻日期当成了最晚日期
    # （真实落款日通常就在 PDF 内、只是没被 pats 命中），属已知抽取局限 → 观察项，不进 problems。
    # 否则正确台账日期会被错误判定成"需处理"，与已验收的干净态矛盾（曜瞿如/福石即此类，差 1 天）。
    b_date_real_problem = []
    b_date_superseded = []      # 前序文书（原处分/告知/复核前）落款早于最终决定 → 观察项
    for x in b_date_real:
        if x.get("row") in row_date_hit:
            # 本行已有文书落款 = 台账处罚时间（且台账记的是**最终**决定）→ 这份是前序文书，正常
            b_date_superseded.append(x)
            continue
        ld = x.get("pdf_latest_date")
        try:
            _diff = abs((date.fromisoformat(ld) - date.fromisoformat(x["ledger"])).days)
        except Exception:
            _diff = 999
        if _diff <= 1:
            print(f"      （落款近差 {_diff} 天，OCR 抽取局限，观察项：行{x.get('row')} "
                  f"{x.get('filename')} 台账{x['ledger']}/PDF{ld}）")
        else:
            b_date_real_problem.append(x)
    b_date_real = b_date_real_problem
    if b_date_superseded:
        print(f"      （前序文书 {len(b_date_superseded)} 份：落款早于最终决定、同行的最终决定落款已回现 "
              f"→ 观察项，不计入结论；如行{b_date_superseded[0].get('row')} "
              f"{b_date_superseded[0].get('filename')} "
              f"PDF{b_date_superseded[0].get('pdf_latest_date')}/台账{b_date_superseded[0]['ledger']}）")
    for v, label in [(b_date_real, "落款未回现"), (b_tgt, "管理人名称未回现"),
                     (b_name, "文件命名不符"), (b_empty, "异常文件"),
                     (b_core, "核心列缺失/敷衍")]:
        for x in v[:6]:
            print(f"      {label}：行{x.get('row')} {str(x)[:88]}")
        if v:
            problems.append(f"B {label} {len(v)} 条")
    if scanned:
        checked = ocr_used + ocr_cached_hits
        cov = 100 * checked // max(1, scanned)
        print(f"      （无文字层 {scanned} 件，已 OCR 复核 {checked} 件 = {cov}%"
              f"｜本次实跑 {ocr_used}、缓存命中 {ocr_cached_hits}"
              f"{'、超出预算跳过 ' + str(ocr_skipped) if ocr_skipped else ''}）")
        rep["B"]["scanned"] = scanned
        rep["B"]["scanned_ocr_checked"] = checked
        rep["B"]["scanned_ocr_coverage"] = round(checked / max(1, scanned), 3)
        # OCR 覆盖率是否算"问题"取决于运行模式：
        #   · --ocr-scan all：意图 100% 复核，<99% 才是真问题；
        #   · 默认 sample:<n> / none：只抽少量样本，覆盖率低是预算设计内 → 观察项，不进 problems。
        _ocr_mode = (a.ocr_scan or "sample:20").split(":")[0]
        if _ocr_mode == "all":
            # 意图 100% 复核：覆盖率 <99% 才是真问题（少量扫描件 OCR 仍拆坏属容差）
            if scanned and cov < 99:
                problems.append(f"B 扫描件 OCR 覆盖率 {cov}%（{scanned - checked} 件未复核）")
        elif scanned:
            # none / sample：覆盖率低是预算设计内 → 观察项，绝不进 problems
            if checked == 0:
                print(f"      （OCR 模式 {_ocr_mode}，未做 OCR 复核为预期内，非问题；"
                      f"需 100% 复核请加 --ocr-scan all）")
            else:
                print(f"      （OCR 抽样模式，覆盖率 {cov}% 为预算内，非问题；"
                      f"需 100% 复核请加 --ocr-scan all）")

    # ---------------- C 分类准确性 ----------------
    print("\n【C 分类准确性】")
    legal = set(AMAC_CATS) | set(CSRC_CATS) | set(EXCH_CATS)
    bad_cat = [c for c in cases if c["category"] not in legal]
    wrong_dir = []
    for c in cases:
        for pt in c["paths"]:
            expect = pt.rsplit("/", 1)[-1]
            if c["category"] and expect and c["category"] != expect:
                wrong_dir.append({"row": c["row"], "category": c["category"], "dir": expect})
    # C1b 来源↔分类枚举匹配：交易所(03)只能用 EXCH_CATS，用六分法名即为路由错误。
    # 教训（2026-10-29）：15 个交易所案例全被套上 02_登记备案/03_投资运作/04_内控治理，
    # 顶层 03 下因此长出 3 个不存在的子目录，而正确的 01/02/03/06 四个目录 0 文件。
    src_bad = [c for c in cases
               if c["source"] == "交易所"
               and re.sub(r"（次要.*", "", c["category"]) not in set(EXCH_CATS)]
    print(f"  C1b 交易所案例误用通用六分法分类：{len(src_bad)}")
    rep["C"]["exch_wrong_taxonomy"] = [{"row": c["row"], "category": c["category"]}
                                       for c in src_bad[:40]]
    for c in src_bad[:6]:
        print(f"      行{c['row']} 交易所[{c['category'][:24]}]")
    if src_bad:
        problems.append(f"C1b 有 {len(src_bad)} 个交易所案例误用六分法分类（应用 01/02/03/04 四分法）")

    # C1c 03 顶层目录白名单（2026-11-02 用户实拍）：03 只允许 01/02/03/04 四个目录。
    # 实测事故：03 下真出现过 06_其他行政违规 目录（「06」是六分法序号、「行政」是行政监管口径），
    # 里面那份决定书本是「拉抬股价、异常交易」→ 必须归 01。只看台账分类（C1b）拦不住
    # 「目录建错了但台账没写错」/「历史残留目录」这两种情形，故按磁盘目录再判一次。
    # 注意：兜底类别**必须存在**（04_其他自律监管），只是不得借用六分法的 06。
    # 2026-11-02 首轮曾把 03 收敛成三分法、兜底塞进 03_信息报送，导致「其他类别消失」被用户当场质疑。
    exch_top = os.path.join(root, "03_交易所监管措施（沪深交易所、北交所，私募产品交易、异常交易纪律）")
    illegal_dirs, missing_dirs = [], []
    if os.path.isdir(exch_top):
        present = {d for d in os.listdir(exch_top) if os.path.isdir(os.path.join(exch_top, d))}
        for d in sorted(present):
            if d not in set(EXCH_CATS):
                n = len([f for f in os.listdir(os.path.join(exch_top, d)) if f.lower().endswith(".pdf")])
                illegal_dirs.append({"dir": d, "pdf": n})
        missing_dirs = [d for d in EXCH_CATS if d not in present]
    print(f"  C1c 03 顶层非四分法目录：{len(illegal_dirs)} ｜ 缺失目录：{len(missing_dirs)}")
    rep["C"]["exch_illegal_dirs"] = illegal_dirs
    rep["C"]["exch_missing_dirs"] = missing_dirs
    for x in illegal_dirs[:6]:
        print(f"      ✗ 非法目录：{x['dir'][:28]}（{x['pdf']} 份 PDF）")
    for d in missing_dirs:
        print(f"      ✗ 缺失目录：{d}")
    if illegal_dirs:
        problems.append("C1c 03 顶层出现四分法之外的目录（%s），交易所只有 01/02/03/04"
                        % "、".join(x["dir"][:16] for x in illegal_dirs[:3]))
    if missing_dirs:
        # 兜底类别是业务实体，只能改名不能删（2026-11-02 用户质疑「其他类别没有了」）
        problems.append("C1c 03 顶层缺失应分类别（%s）——兜底类别必须保留"
                        % "、".join(d[:10] for d in missing_dirs))

    print(f"  C1 违规类型不在枚举内：{len(bad_cat)}")
    print(f"  C2 存放目录与台账违规类型不一致：{len(wrong_dir)}")
    rep["C"]["illegal_category"] = [{"row": c["row"], "category": c["category"]}
                                    for c in bad_cat[:40]]
    rep["C"]["dir_mismatch"] = wrong_dir[:40]
    if bad_cat or wrong_dir:
        problems.append("C 分类值或目录不一致")
        for c in bad_cat[:6]:
            print(f"      非法分类：行{c['row']} {c['category'][:30]}")
        for x in wrong_dir[:6]:
            print(f"      目录不符：行{x['row']} 台账[{x['category'][:20]}] 目录[{x['dir'][:20]}]")

    # 规则重分类比对（只在规则「自信」时才比对）
    mismatch, judged = [], 0
    for f in files:
        fp = os.path.join(root, f["path"], f["filename"]) if f["path"] else ""
        if not fp or not os.path.exists(fp) or f["source"] == "中基协":
            continue
        try:
            d = fitz.open(fp); t = flat("".join(p.get_text() for p in d)); d.close()
        except Exception:
            continue
        if len(t) < 60:
            continue
        # 2026-10-29 修正：必须走 classify(text, source) 才能让交易所落到四分法；
        # 旧写法 recat(rule_classify(...)) 对交易所套用六分法，C3 与台账永远对不上。
        if f["source"] == "交易所":
            pred = classify(t, "交易所")
        else:
            pred = recat(rule_classify(t, need_confident=True), f["source"])
        if not pred:
            continue
        judged += 1
        cur = re.sub(r"（次要.*", "", f["case"]["category"])
        if cur and pred[:2] != cur[:2]:
            mismatch.append({"row": f["row"], "ledger": cur, "rule": pred,
                             "manager": f["manager"][:26]})
    actionable = [x for x in mismatch if x["ledger"].startswith("06_")]
    rate = len(mismatch) / max(1, judged)
    print(f"  C3 规则重分类与台账不一致：{len(mismatch)}/{judged} = {rate:.0%}"
          f"（多违规并存取主分类，属常态；>25% 才判为问题）")
    print(f"      其中台账为「06_其他」但规则可明确归类：{len(actionable)} 条（建议复核细分）")
    rep["C"]["reclassify_mismatch"] = mismatch[:60]
    rep["C"]["reclassify_rate"] = round(rate, 3)
    rep["C"]["reclassify_actionable"] = actionable[:40]
    for x in actionable[:8]:
        print(f"      行{x['row']} 台账[{x['ledger'][:18]}] → 规则[{x['rule'][:18]}] "
              f"{x['manager'][:22]}")
    if rate > 0.25:
        problems.append(f"C3 规则重分类差异率 {rate:.0%} 偏高，分类器与规则已脱节")
    if len(actionable) > 20:
        problems.append(f"C3 有 {len(actionable)} 条「06_其他」可被规则明确归类，需细分")
    other = sum(1 for c in cases if c["category"].startswith("06_"))
    print(f"  C4 「06_其他」占比：{other}/{len(cases)} = {100*other//max(1,len(cases))}%"
          f"（>15% 需人工细分）")
    rep["C"]["other_ratio"] = round(other / max(1, len(cases)), 3)
    if len(cases) and other / len(cases) > 0.15:
        problems.append("C4 其他类占比偏高，需细分")

    # ---------------- C5/C6/C7 库内容结构（2026-10-29 新增，回应「只对数量不对内容」） ---
    # C5 分类目录空置：目录存在但 0 份 PDF → 要么分类器把案例路由错了，要么目录该删。
    # 区分两种空目录：
    #   ① 路由错误型——台账里有该分类的案例，却没落到这个目录 → 必须修（真问题）；
    #   ② 本期确无案例——台账里该分类就是 0 条 → 记观察项，不算问题。
    SRC_OF = {"01": "中基协", "02": "证监局", "03": "交易所"}
    cnt = Counter((c["source"], re.sub(r"（次要.*", "", c["category"])) for c in cases)
    empty_dirs, empty_by_design = [], []
    for d1 in sorted(os.listdir(root)):
        p1 = os.path.join(root, d1)
        if not os.path.isdir(p1) or d1[:2] not in ("01", "02", "03"):
            continue
        for d2 in sorted(os.listdir(p1)):
            p2 = os.path.join(p1, d2)
            if not os.path.isdir(p2):
                continue
            n = len([f for f in os.listdir(p2) if f.lower().endswith(".pdf")])
            if n:
                continue
            key = (SRC_OF.get(d1[:2], ""), d2)
            (empty_dirs if cnt.get(key, 0) > 0 else empty_by_design).append(
                {"dir": f"{d1}/{d2}", "ledger_rows": cnt.get(key, 0)})
    print(f"  C5 分类目录空置：路由错误 {len(empty_dirs)} ｜ 本期确无案例 {len(empty_by_design)}")
    rep["C"]["empty_dirs"] = empty_dirs
    rep["C"]["empty_by_design"] = [x["dir"] for x in empty_by_design]
    for x in empty_dirs:
        print(f"      ✗ 路由错误：{x['dir'].split('/')[-1][:34]}（台账有 {x['ledger_rows']} 案却 0 文件）")
    for x in empty_by_design:
        print(f"      · 本期无案例：{x['dir'].split('/')[-1][:34]}")
    if empty_dirs:
        problems.append(f"C5 有 {len(empty_dirs)} 个分类目录为空但台账有案例"
                        f"（{('、'.join(x['dir'].split('/')[-1] for x in empty_dirs))[:60]}）")

    # C6 占位/碎片命名：文件名 / 台账主体 里不得出现「不明确/缺位/待核」等占位词，
    # 也不得出现「至今/职位/签章/挂名/离职…」这类正文碎片（2026-10-29 实测 4 个
    # 「职位-上海通赛…pdf」「签章-广东横琴…pdf」式文件名 + 9 行 D 列污染）。
    # 词表统一从 subject_clean 取，禁止在这里再维护一份。
    import subject_clean as _SC
    bad_names = []
    for f in files:
        fn = f["filename"] or ""
        if any(k in fn for k in BAD_TOKEN):
            bad_names.append({"row": f["row"], "filename": fn, "why": "文件名占位"})
            continue
        m = re.match(r"^([^\-—]+)[\-—].+\.pdf$", fn)
        if m and _SC.is_noise_token(m.group(1)):
            bad_names.append({"row": f["row"], "filename": fn,
                              "why": "文件名以碎片词开头（%s）" % m.group(1)})
    for c in cases:
        if any(k in (c["subject"] or "") for k in ("不明确", "待补")) \
                or re.search(r"；.*缺位", c["subject"] or ""):
            bad_names.append({"row": c["row"], "filename": c["subject"][:40], "why": "主体占位"})
    # 文件名不得带「（2）（3）…」下载去重序号（2026-11-02 用户实拍）：
    # 同一主体有多份决定书时下载器会加序号，落到台账里既看不出是哪一份，
    # 又让人误以为是「其他分类」的副本。规范：改用落款日「（YYYY-MM-DD）」
    # （rename_dup_suffix.py 一次性修复）。
    dup_seq = []
    for f in files:
        fn = f["filename"] or ""
        if re.search(r"（\d+）(?=\.pdf$)", fn, re.I):
            dup_seq.append({"row": f["row"], "filename": fn})
    print(f"  C6b 文件名带（N）下载序号：{len(dup_seq)}")
    for x in dup_seq[:6]:
        print(f"      ✗ 行{x['row']} {x['filename'][:44]}")
    if dup_seq:
        problems.append(f"C6b 有 {len(dup_seq)} 个文件名带（N）序号后缀，应改为落款日")
    # C6c 复核案文件名必须标明「（原决定）/（复核决定）」（2026-11-02 用户裁定：
    # 日期后缀不表意，双份文书必须一眼分清哪份是复核）。
    # 2026-11-02 增补：原实现只认「申辩列含【复核决定」这一个闸门，而复核决定书是扫描件、
    # 台账最初又写着「无」→ 闸门自锁，上海通赛/汕头宜华/上海世联行三案都成了漏网。
    # 现改为**双向**：① 申辩有【复核决定 → 文件名必须有（复核决定）；② 文件名有（复核决定）
    # 而申辩没有【复核决定 → 台账漏记复核结论（同属内容缺失，必须补）。
    unlabeled, orphan_lbl = [], []
    for c in cases:
        names = c.get("filenames") or []
        has_mark = "【复核决定" in str(c.get("defense") or "")
        has_name = any("（复核决定）" in (n or "") for n in names)
        if has_mark and names and not has_name:
            unlabeled.append({"row": c["row"], "files": names[:3]})
        elif has_name and not has_mark:
            orphan_lbl.append({"row": c["row"], "files": [n for n in names if "（复核决定）" in (n or "")][:2]})
    print(f"  C6c 复核案未标注（复核决定）：{len(unlabeled)} ｜ 文件名标了复核但台账无复核结论：{len(orphan_lbl)}")
    for x in unlabeled[:6]:
        print(f"      ✗ 行{x['row']} {(x['files'][0] if x['files'] else '')[:44]}")
    for x in orphan_lbl[:6]:
        print(f"      ✗ 行{x['row']} 台账申辩列缺【复核决定】：{(x['files'][0] if x['files'] else '')[:40]}")
    if unlabeled:
        problems.append(f"C6c 有 {len(unlabeled)} 个复核案的文件名未标明（复核决定）")
    if orphan_lbl:
        problems.append(f"C6c 有 {len(orphan_lbl)} 个案子文件名标了（复核决定）但台账申辩列没有【复核决定】内容")
    print(f"  C6 占位/碎片命名（不明确/缺位/待核/至今/职位/签章…）：{len(bad_names)}")
    rep["C"]["placeholder_names"] = bad_names[:40]
    for x in bad_names[:8]:
        print(f"      行{x['row']} {x['why']}：{x['filename'][:52]}")
    if bad_names:
        problems.append(f"C6 有 {len(bad_names)} 处占位/碎片命名需回填主体或重命名")

    # C8 台账主体列碎片：D 列「机构；人名1、人名2」的人名段不得出现正文碎片
    # （旧版 is_valid_subject 把「2~8 汉字」一律当人名，实测混入 9 行 / 13 个碎片）。
    dirty_subj = []
    for c in cases:
        segs = [s.strip() for s in re.split(r"[；;]", c["subject"] or "") if s.strip()]
        for s in segs[1:]:
            for nm in [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]:
                if _SC.is_noise_token(nm):
                    dirty_subj.append({"row": c["row"], "token": nm, "subject": c["subject"][:40]})
    print(f"  C8 台账主体含碎片词（至今/职位/签章/离职…）：{len(dirty_subj)}")
    rep["C"]["subject_noise"] = dirty_subj[:40]
    for x in dirty_subj[:8]:
        print(f"      行{x['row']} 碎片「{x['token']}」：{x['subject']}")
    if dirty_subj:
        problems.append(f"C8 有 {len(dirty_subj)} 处主体碎片词需清洗（content_repair --only purge）")

    # C11 机构名前角色短语（2026-11-02 新增）：「曾登记为X公司」「实际负责X公司」。
    # 危害不止难看：机构名一被污染，同一家中基协/证监局的多份文书就不再归为同一主体，
    # 「同机构多案加日期」「同案拆行判重」全部失效（浙江优策 4 行曾被拆成 3 个"不同机构"）。
    role_bad = []
    for c in cases:
        _first = re.split(r"[；;]", c["subject"] or "")[0].strip()
        _fix = _SC.strip_role_prefix(_first)
        if _fix != _first:
            role_bad.append({"row": c["row"], "old": _first[:40], "new": _fix[:40]})
    print(f"  C11 机构名带角色前缀（曾登记为/实际负责…）：{len(role_bad)}")
    rep["C"]["role_prefix"] = role_bad[:40]
    for x in role_bad[:8]:
        print(f"      行{x['row']} {x['old']} → {x['new']}")
    if role_bad:
        problems.append(f"C11 有 {len(role_bad)} 行机构名带角色前缀需剥离（normalize_subjects.py --apply）")

    # C7 跨局合并：同一行若挂着不同证监局（URL 首段局名不同）的文书，说明多案被并成一行。
    bureau_re = re.compile(r"csrc\.gov\.cn/([a-z]+)/")
    merged = []
    for c in cases:
        bs = [m.group(1) for u in c["urls"] for m in [bureau_re.search(u or "")] if m]
        if len(set(bs)) > 1:
            merged.append({"row": c["row"], "bureaus": sorted(set(bs)),
                           "subject": c["subject"][:40]})
    print(f"  C7 跨局合并行（一行含多个证监局文书）：{len(merged)}")
    rep["C"]["cross_bureau_merged"] = merged[:20]
    for x in merged[:6]:
        print(f"      行{x['row']} {x['bureaus']} {x['subject'][:30]}")
    if merged:
        problems.append(f"C7 有 {len(merged)} 行把不同证监局的多个案例并成了一行，需拆分")

    # C9 典型案例副本对齐：标「是否典型案例=是」的行，其案例文件必须在 05_重点典型案例
    # 下（递归任意子类目录）有同名副本；否则典型目录与台账脱节。根因：content_repair 改主库
    # 文件名时未联动 05 副本（实测 河南智慧和信 一例旧名残留），现已手动补齐并固化此断言。
    typ05_dir = os.path.join(root, "05_重点典型案例")
    typ05_files = set()
    if os.path.isdir(typ05_dir):
        for _r, _d, fs in os.walk(typ05_dir):
            for f in fs:
                if f.lower().endswith(".pdf"):
                    typ05_files.add(f)
    typ_missing = []
    for c in cases:
        if c["typical"] != "是":
            continue
        for fn in c["filenames"]:
            fn = (fn or "").strip()
            if not fn:
                continue
            if fn not in typ05_files:
                typ_missing.append({"row": c["row"], "filename": fn})
    typ_all = {fn.strip() for c in cases if c["typical"] == "是"
               for fn in c["filenames"] if (fn or "").strip()}
    typ_extra = sorted(f for f in typ05_files if f not in typ_all)
    print(f"  C9 典型行→05副本缺失：{len(typ_missing)} ｜ 05 多余副本（观察）：{len(typ_extra)}")
    rep["C"]["typical_missing_05"] = typ_missing[:20]
    rep["C"]["typical_extra_05"] = typ_extra[:20]
    for x in typ_missing[:8]:
        print(f"      行{x['row']} 05缺副本：{x['filename'][:52]}")
    if typ_missing:
        problems.append(f"C9 有 {len(typ_missing)} 个标典型的案例在 05 目录缺副本")

    # C10 库根条目白名单（2026-11-02 新增，用户长期原则「库根只放交付物」）
    # 实测事故：维护脚本习惯 `ledger.replace(".xlsx","_bak_x.xlsx")`，把备份直接写在库根，
    # 一轮下来多出 8 个 `04_综合索引台账【总表】_bak_*.xlsx`，用户第一眼就看到「怎么多了一堆表」。
    # 白名单：01~06 六个目录 + 台账 xlsx / html（同一份台账的两种形态）
    #         + 3 个 .md（不收录清单/更新记录/校验人工复核豁免清单）。
    # 备份一律放库外：<库父目录>/_案例库维护留痕/台账备份/（见 scripts/ledger_bak.py）。
    _allow_files = {"不收录清单.md", "更新记录.md", "校验人工复核豁免清单.md"}
    stray = []
    legacy_macro = []                     # 已停用的 .xlsm，只提示删除，不再校验内容
    for name in sorted(os.listdir(root)):
        p = os.path.join(root, name)
        if os.path.isdir(p):
            if not re.match(r"^0[1-6]_", name):
                stray.append(name + "/")
        elif name.startswith("~$") or name.startswith("."):
            continue                      # Office 临时锁文件不计数
        elif name in _allow_files:
            continue
        elif re.match(r"^04_.*\.xlsx$", name):
            continue                      # 台账正本（只允许 1 个）
        elif re.match(r"^04_.*\.xlsm$", name):
            # 2026-11-04 用户决定：宏版停用、不再输出。库根若还留着旧文件就报出来，
            # 但不再做「与正本四部件逐字节比对」——那个检查是给还在产出的文件用的。
            legacy_macro.append(name)
            stray.append(name)
            continue
        elif re.match(r"^04_.*\.html$", name):
            continue                      # 网页版检索台（与 xlsx 同名同放库根，用户要求）
        else:
            stray.append(name)

    if legacy_macro:
        print(f"      （库根还有已停用的宏版：{'、'.join(legacy_macro)}——可直接删除）")

    print(f"  C10 库根非白名单条目：{len(stray)}")
    rep["C"]["root_stray"] = stray
    for x in stray[:10]:
        print(f"      ✗ 多余：{x[:60]}")
    if stray:
        problems.append("C10 库根出现 %d 个非白名单条目（备份/临时文件必须放库外）：%s"
                        % (len(stray), "、".join(stray[:5])))

    # C12 逐格内容形态（2026-11-02 新增）
    # 用户实拍：F 列写成「91310109312278503E ）在管理私募基金财产过程中…」——主体名和左括号
    # 被旧抽取器 `(?:经查)[^：:]*[:：]` 一起吃掉了（它吃到了「统一社会信用代码：」的冒号）。
    # 这类**格子级污染**列级规则查不出来：列非空、枚举合法、长度正常，必须逐格过形态。
    # 复用 audit_cells.check_cell（唯一判据），只统计「非观察项」——
    #   带 `'` 的码（如 X21'单条整段）是原文本身如此，不算问题。
    from audit_cells import check_cell as _cellchk
    _COLKEY = ["seq", "decided", "publish", "subject", "measures", "grounds",
               "evidence", "legal_basis", "defense", "category", "source",
               "filenames", "paths", "urls", "typical"]
    cell_bad = {}
    for c in cases:
        for ci, key in enumerate(_COLKEY):
            v = c.get(key)
            s = "；".join(v) if isinstance(v, list) else str(v or "")
            for code in _cellchk(ci, s, c["row"]):
                if "'" in code:
                    continue
                e = cell_bad.setdefault(code, [0, c["row"]])
                e[0] += 1
    n_cell = sum(v[0] for v in cell_bad.values())
    print(f"  C12 逐格形态污染（空格/残片/越界/半角标点/叠字）：{n_cell}")
    rep["C"]["cell_pollution"] = n_cell
    for code, (n, row) in sorted(cell_bad.items(), key=lambda x: -x[1][0])[:6]:
        print(f"      ✗ {code}：{n} 格（首个行 {row}）")
    if cell_bad:
        problems.append("C12 台账逐格形态污染 %d 处（首个：%s 行%s）——跑 "
                        "fix_cells/fix_grounds 清理" % (n_cell, list(cell_bad)[0],
                                                     cell_bad[list(cell_bad)[0]][1]))

    # ---------------- D 法规依据 ----------------
    print("\n【D 配套法规依据】")
    law_dir = [d for d in os.listdir(root) if d.startswith("06_")]
    law_files, law_ext, law_disk = 0, Counter(), set()
    if law_dir:
        for _r, _d, fs in os.walk(os.path.join(root, law_dir[0])):
            for f in fs:
                if f.startswith("~$") or f.startswith("."):
                    continue
                law_files += 1
                law_ext[os.path.splitext(f)[1].lower() or "(无扩展名)"] += 1
                law_disk.add(f)
    try:
        lw = wb["配套法规依据清单"]
        law_rows = lw.max_row - 1
        # 按表头名取列号（2026-10-27 修复：7 列 schema 下文件名在第 5 列、引用数在第 4 列，
        # 旧代码写死 11/10 导致 listed 为空、把 9 部法规全判成「磁盘有清单无」的假 D1）。
        _lh = [str(lw.cell(row=1, column=c).value or "").strip() for c in range(1, lw.max_column + 1)]
        _col = lambda n: (_lh.index(n) + 1) if n in _lh else None
        _nc, _fc, _rc = _col("法规名称"), _col("文件名"), _col("库内案例引用数")
        names, fnames, refs = [], [], {}
        for i in range(2, lw.max_row + 1):
            nm = str(lw.cell(row=i, column=_nc).value or "").strip() if _nc else ""
            fn = str(lw.cell(row=i, column=_fc).value or "").strip() if _fc else ""
            rf = lw.cell(row=i, column=_rc).value if _rc else 0
            names.append(nm)
            fnames.append(fn)
            refs[nm] = rf or 0
    except Exception:
        law_rows, names, fnames, refs = 0, [], [], {}
    listed = set(f for f in fnames if f)
    law_missing = sorted(listed - law_disk)
    law_extra = sorted(law_disk - listed)
    ext_str = "、".join(f"{k}{v}" for k, v in law_ext.most_common())
    print(f"  D1 06 目录 {law_files} 份（{ext_str}）｜ 台账法规清单 {law_rows} 行")
    print(f"      清单有磁盘无：{len(law_missing)} ｜ 磁盘有清单无：{len(law_extra)}")
    for x in law_missing[:8]:
        print(f"      缺口：{x[:60]}")
    for x in law_extra[:8]:
        print(f"      未登记：{x[:60]}")
    rep["D"]["law_files"] = law_files; rep["D"]["law_rows"] = law_rows
    rep["D"]["law_missing"] = law_missing[:40]; rep["D"]["law_extra"] = law_extra[:40]
    if law_files != law_rows or law_missing or law_extra:
        problems.append("D1 法规数目与清单不一致")
    zero = [n for n, v in refs.items() if n and (v or 0) == 0]
    print(f"  D2 库内引用数为 0 的法规：{len(zero)}（私募主干规则按约定保留，仅提示）")
    rep["D"]["zero_ref"] = zero[:40]
    for n in zero[:8]:
        print(f"      零引用：{n[:44]}")
    # 引用口径（v2）：决定书正文 ∪ 台账核心处罚依据列，同一案例出现多次只算 1
    cited = Counter()
    for c in cases:
        doc_names = set()
        hb = c["legal_basis"]
        for m in re.finditer(r"《([^《》]{4,40})》", hb):
            doc_names.add(m.group(1))
        for f in [g for g in files if g["row"] == c["row"]]:
            fp = os.path.join(root, f["path"], f["filename"]) if f["path"] else ""
            if not fp or not os.path.exists(fp):
                continue
            try:
                d = fitz.open(fp); t = flat("".join(p.get_text() for p in d)); d.close()
            except Exception:
                continue
            for m in re.finditer(r"《([^《》]{4,40})》", t):
                n = m.group(1)
                if n[0].isdigit() or n[0] in "第各" or any(k in n for k in
                        ("合同", "协议", "报告", "说明", "告知", "方案", "计划", "章程", "决议")):
                    continue
                doc_names.add(n)
        for n in doc_names:
            cited[n] += 1

    def norm(s):
        return re.sub(r"[（(].*", "", re.sub(r"[\s《》]", "", s or "")).strip()

    have_norm = {norm(n) for n in names}
    have_all = have_norm | {norm(v) for v in ALIAS.values()}
    lack = []
    for n, c in cited.most_common():
        if c < 3:
            continue
        key = norm(n)
        if key in have_norm or key in have_all:
            continue
        canon = ALIAS.get(key) or ALIAS.get(norm(re.sub(r"第\d+号", "", n)))
        if canon and norm(canon) in have_all:
            continue
        if len(key) >= 5 and any(key in h or h in key for h in have_all):
            continue
        if key in NOISE_LAW:
            continue
        lack.append((n, c))
    print(f"  D3 正文/核心处罚依据点名引用 ≥3 次但清单缺失：{len(lack)}")
    rep["D"]["missing_laws"] = [{"name": n, "refs": c} for n, c in lack[:30]]
    for n, c in lack[:10]:
        print(f"      缺：{n[:40]}（被引用 {c} 次）")
    if lack:
        problems.append("D3 有法规被引用但未入库")

    # ---------------- 结论 ----------------
    print("\n" + "=" * 72)
    if problems:
        print("结论：存在问题，需处理 ——")
        for p in problems:
            print("   · " + p)
    else:
        print("结论：全部校验通过")
    print("=" * 72)
    rep["problems"] = problems
    if a.json:
        # 2026-10-27 加固：用 with 块确保 flush/close；default=str 兜底任何不可序列化对象，
        # 避免 rep 含嵌套对象时 json.dump 抛错导致报告整份丢失（曾经 exit 1 但 0 字节落盘）。
        try:
            with open(a.json, "w", encoding="utf-8") as _jf:
                json.dump(rep, _jf, ensure_ascii=False, indent=1, default=str)
            print("报告已写入", a.json)
        except Exception as _e:
            print("!! 报告写出失败：", _e)
            try:
                with open(a.json + ".fallback", "w", encoding="utf-8") as _jf:
                    json.dump(rep, _jf, ensure_ascii=False, indent=1, default=str)
                print("!! 已写出降级副本：", a.json + ".fallback")
            except Exception:
                pass
            raise
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
