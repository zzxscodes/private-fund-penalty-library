# -*- coding: utf-8 -*-
"""库内容缺陷修复（内容层，不是数量层）。

为什么需要这个脚本
------------------
2026-10-29 用户实测指出：此前验收只做「数量对账」（567 行 ↔ 871 份 PDF、孤儿 0/缺失 0），
没有深入**库内容**，于是三类硬伤全部漏过：

1. 03 交易所顶层下 4 个正确分类目录（01_异常交易 / 02_产品交易风控 / 03_信息报送 /
   06_其他行政违规）**全空**，15 个案例被套上通用六分法目录名塞进 02_登记备案 /
   03_投资运作 / 04_内控治理 —— 根因是 classify() 对 source='交易所' 只 recat 了
   「06_」前缀，其余原样返回六分法名。
2. 台账/PDF 出现「不明确.pdf」「不明确（2）.pdf」「不明确-珠海通沛….pdf」
   「缺位-景行天成….pdf」「投资管理有限公司.pdf」这类占位/残缺命名 —— 根因是主体抽取
   对深交所债券监管函（「关于对X的监管函」+「你公司作为…发行人」）完全失效、把正文
   「合规风控负责人缺位」里的「缺位」当成高管姓名、以及把地名（深圳前海）当人名后
   把机构名截成「凯创业投资有限公司」。
3. 05_重点典型案例/募集宣传自媒体违规案例 目录 0 文件（悬置空目录）——
   典型案例遴选从未选中「自媒体/互联网公开宣传」类。
4. 一行多局：把不同证监局、不同处罚对象的多个案例并成一行（行 69/82/294），
   主体列退化成「投资管理有限公司；吴明、王军华、符冠华、林建伟、之一」。

本脚本按「先审计（dry-run）→ 后修复（--apply）→ 再复验」的顺序做**可复现**的修复，
不手改 Excel。

用法
----
    python content_repair.py --root "<库根>"                     # dry-run，只打印计划
    python content_repair.py --root "<库根>" --apply             # 全部修复
    python content_repair.py --root "<库根>" --only reclass --apply   # 只做某一阶段

阶段顺序固定：reclass → refill → split → typical。修复后**必须**再跑：
    python ledger.py --root "<库根>" sort        # 挂网日期倒序 + 重排序号
    python verify_library.py --root "<库根>"     # C1b/C5/C6/C7 应全部归零
"""
import argparse
import os
import re
import shutil
import sys
from pathlib import Path
from datetime import datetime

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import fitz  # noqa: E402

import build_pipeline as BP          # noqa: E402  主体/措施/文件名
import classify_filter as CF         # noqa: E402  分类（交易所四分法）
import extract_core as EC            # noqa: E402  四核心列
import ledger_style as LS            # noqa: E402  台账列定义/追加/重排
import subject_clean as SC           # noqa: E402  碎片词表（NOISE_TOKENS 唯一事实源）
from openpyxl import load_workbook    # noqa: E402

EXCH_TOP = "03_交易所监管措施（沪深交易所、北交所，私募产品交易、异常交易纪律）"
PLACEHOLDER = ("不明确", "待补", "缺位", "待核", "未命名", "无标题")
BUREAU_RE = re.compile(r"csrc\.gov\.cn/([a-z]+)/")
INST_SUF = re.compile(r"(有限公司|有限责任公司|股份公司|股份有限公司|合伙企业|有限合伙|管理中心|事务所)$")
# 只剩行业通用词的残缺机构名（主体抽取被截断后的产物）
BARE_ORG = {"投资管理有限公司", "资产管理有限公司", "股权投资有限公司", "创业投资有限公司",
            "基金管理有限公司", "资本管理有限公司", "财富管理有限公司", "管理有限公司",
            "投资有限公司", "有限公司", "合伙", "企业"}
# 被误当成人名的地名/片区名
PLACES = ("深圳", "北京", "上海", "广州", "杭州", "宁波", "成都", "南京", "武汉", "西安",
          "天津", "重庆", "苏州", "厦门", "青岛", "济南", "郑州", "长沙", "合肥", "福州",
          "大连", "沈阳", "哈尔滨", "昆明", "南昌", "贵阳", "南宁", "太原", "石家庄",
          "兰州", "乌鲁木齐", "呼和浩特", "银川", "西宁", "拉萨", "海口", "长春", "珠海",
          "东莞", "佛山", "温州", "绍兴", "嘉兴", "无锡", "常州", "南通", "泉州", "前海",
          "横琴", "自贸区",
          # 省级行政区（2026-10-29 补：行571「河南智慧和信基金管理有限公司」被旧的
          # 「和」切分劈成「河南智慧」+「信基金管理有限公司」，前者又因省名不在表里
          # 逃过了检测，主体一直残缺到验收都没人发现）
          "河北", "山西", "辽宁", "吉林", "黑龙江", "江苏", "浙江", "安徽", "福建", "江西",
          "山东", "河南", "湖北", "湖南", "广东", "广西", "海南", "四川", "贵州", "云南",
          "陕西", "甘肃", "新疆", "西藏", "内蒙古", "宁夏", "青海")

DRY = True

# 交易所「监管函」类文书：不写「本所决定…采取」，也不列证据，自动抽取只能拿到
# 「对你公司予以书面警示」。以下两条按原文逐字核定（2026-10-29 人工读 PDF 后回填），
# 键为原文链接（唯一且稳定）。新增同类时照此追加，不要改自动抽取逻辑去硬凑。
MANUAL_FIX = {
    "https://reportdocs.static.szse.cn/UpFiles/zqjghj/sup_jghj_00019EB5D3DB813FEEF451B68BEE623F.pdf": {
        "measures": "管理人：对国厚资产管理股份有限公司予以书面警示（出具监管函，债券业务中心监管函〔2026〕第11号）",
        "evidence": "公司债券“H9国厚01”（112923）2025年年度报告披露记录：未在规定时间内披露，迟至2026-05-29披露",
        "legal_basis": "《深圳证券交易所公司债券上市规则（2023年修订）》第1.5条、第3.2.2条",
    },
    "https://reportdocs.static.szse.cn/UpFiles/zqjghj/sup_jghj_00019D241B723F3FE6BF545763EE143F.pdf": {
        "measures": "管理人：对国厚资产管理股份有限公司予以书面警示；高管（李厚文）：予以书面警示（出具监管函，债券业务中心监管函〔2026〕第4号）",
        "evidence": "中国证监会安徽监管局《行政监管措施决定书》（〔2025〕72号）；2023年、2024年年度报告；董监高变动信息披露记录",
        "legal_basis": "《深圳证券交易所公司债券上市规则（2022年修订）》第1.5条、第3.1.1条、第3.2.4条；"
                       "《深圳证券交易所公司债券上市规则（2023年修订）》第1.5条、第3.1.1条、第3.2.2条",
    },
}


# ---------------------------------------------------------------- 基础工具
def say(msg):
    print(("  [计划] " if DRY else "  [已改] ") + msg)


def txt_of(path):
    try:
        d = fitz.open(path)
        t = "".join(p.get_text() for p in d)
        d.close()
        return t
    except Exception:
        return ""


_OCR = None


def ocr_head(path, max_pages=2):
    """扫描件（无文字层）取前几页 OCR 文本——用于识别文书类型 / 兜底抽主体。"""
    global _OCR
    try:
        import numpy as np
        from rapidocr_onnxruntime import RapidOCR
        if _OCR is None:
            _OCR = RapidOCR()
        d = fitz.open(path)
        out = []
        for i in range(min(max_pages, d.page_count)):
            pix = d[i].get_pixmap(dpi=150)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n)
            res, _ = _OCR(img)
            out += [r[1] for r in (res or [])]
        d.close()
        return "\n".join(out)
    except Exception as e:
        print("  [WARN] OCR 失败：%s (%s)" % (os.path.basename(path), e))
        return ""


def doc_text(path, ocr_pages=2):
    """文字层优先，不足 60 字走 OCR（中基协附件多为扫描件）。"""
    t = txt_of(path) if path else ""
    if len(t) < 60 and path is not None:
        t = ocr_head(path, ocr_pages)
    return t


def split_multi(v):
    if v is None:
        return []
    return [p.strip() for p in re.split(r"[；;]", str(v)) if p.strip()]


def sign_date(text, fallback=""):
    tail = re.sub(r"\s+", "", text or "")[-800:]
    ds = re.findall(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", tail)
    if not ds:
        return fallback
    y, m, d = ds[-1]
    return "%s-%02d-%02d" % (y, int(m), int(d))


def find_pdf(root, d, name):
    """按「存放路径 + 文件名」定位；路径已变（前序阶段移动过 / DRY 未落地）时按顶层目录兜底搜。"""
    p = Path(root) / d / name
    if p.is_file():
        return p
    top = Path(root) / d.split("/")[0]
    for dp, _, fs in os.walk(top):
        if name in fs:
            return Path(dp) / name
    return None


def place_person(x):
    """「人名」位是否是地名碎片。

    只看**开头**：机构名被劈开后残留的碎片总以地名开头（河南智慧 / 深圳前海），
    而真实姓名里地名多在结尾或中间（王广西、李海南）——按「包含」判会把
    「王广西」误删（行239 实测），所以必须用 startswith。
    """
    return bool(x) and len(x) <= 6 and any(x.startswith(pl) for pl in PLACES)


def parse_name(fn):
    """'张三-XX有限公司.pdf' → ('张三','XX有限公司')；'XX有限公司.pdf' → ('','XX有限公司')。"""
    stem = os.path.splitext(os.path.basename(fn))[0]
    stem = re.sub(r"（\d+）$", "", stem).strip()
    if "-" in stem:
        a, b = stem.split("-", 1)
        return a.strip(), b.strip()
    return "", stem


def bad_name(fn):
    """文件名是否「不可自解释」：占位词 / 残缺机构名 / 地名被当人名。"""
    if any(k in fn for k in PLACEHOLDER):
        return True
    person, org = parse_name(fn)
    if place_person(person):
        return True                       # 「深圳前海-凯创业投资有限公司」
    if org in BARE_ORG:
        return True                       # 「投资管理有限公司.pdf」
    core = re.sub(r"[（）()]", "", org)
    if INST_SUF.search(org) and len(core) <= 7:
        return True                       # 「有限公司.pdf」「管理有限公司.pdf」
    return False


def move_pdf(root, srcdir, dstdir, name):
    """移动 + basename/isfile 双断言（硬约定：路径列只存目录，靠 basename 定位）。"""
    src = Path(root) / srcdir / os.path.basename(name)
    dst_dir = Path(root) / dstdir
    assert src.is_file(), "源文件不存在：%s" % src
    if DRY:
        say("移动 %s：%s → %s" % (os.path.basename(name), srcdir.split("/")[-1][:16],
                                dstdir.split("/")[-1][:16]))
        return os.path.basename(name)
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / os.path.basename(name)
    if dst.exists():
        stem, ext = os.path.splitext(os.path.basename(name))
        i = 2
        while dst.exists():
            dst = dst_dir / ("%s（%d）%s" % (stem, i, ext))
            i += 1
    shutil.move(str(src), str(dst))
    assert dst.is_file(), "移动后目标不存在：%s" % dst
    return dst.name


def rename_pdf(root, d, old, new):
    assert os.path.basename(old) == old and os.path.basename(new) == new, (old, new)
    src = Path(root) / d / old
    if old == new:
        return old
    assert src.is_file(), "源文件不存在：%s" % src
    dst = Path(root) / d / new
    if dst.exists():
        stem, ext = os.path.splitext(new)
        i = 2
        while (Path(root) / d / ("%s（%d）%s" % (stem, i, ext))).exists():
            i += 1
        new = "%s（%d）%s" % (stem, i, ext)
        dst = Path(root) / d / new
    if DRY:
        say("重命名 %s → %s" % (old[:44], new[:44]))
        return new
    os.rename(str(src), str(dst))
    assert dst.is_file()
    return new


def new_name_for(path, mgr, persons):
    """按文书自身内容给一个能自解释的文件名。"""
    t = doc_text(path)
    if "复核决定书" in t[:800] or "复议决定书" in t[:800]:
        # 复核/复议决定书不能按「当事人」命名，否则会被误当成新的处罚案例（珠海通沛实测）
        return ("复核决定书-%s.pdf" % mgr) if mgr else ""
    ps = [p for p in (persons or []) if p and p not in PLACEHOLDER]
    if ps and mgr:
        return "%s-%s.pdf" % (ps[0], mgr)
    if ps:
        return "%s.pdf" % ps[0]          # 机构不可得时退回纯人名，绝不产出「吴明-.pdf」
    return ("%s.pdf" % mgr) if mgr else ""


# ---------------------------------------------------------------- 台账读写
def read_rows(root):
    wb = load_workbook(os.path.join(root, LS.LEDGER_NAME))
    ws = wb[LS.SHEET_CASE]
    head = [c.value for c in ws[1]]
    assert head[:len(LS.CASE_HEADERS)] == LS.CASE_HEADERS, "台账表头与冻结规格不一致：%s" % head
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=len(LS.CASE_HEADERS),
                                       values_only=True), start=2):
        if not any(v is not None and str(v).strip() for v in r[1:]):
            continue
        rows.append({"row": i, "v": list(r)})
    return wb, ws, rows


def get(r, key):
    return r["v"][LS.COL[key]]


def put(r, key, val):
    r["v"][LS.COL[key]] = val


# ---------------------------------------------------------------- ① 交易所重分类
def phase_reclass(root, rows):
    print("\n=== ① 交易所案例按四分法重分类 ===")
    n = 0
    for r in rows:
        if str(get(r, "source") or "").strip() != "交易所":
            continue
        files, dirs = split_multi(get(r, "filenames")), split_multi(get(r, "paths"))
        if not files or not dirs:
            continue
        fp = find_pdf(root, dirs[0], files[0])
        t = doc_text(fp) if fp else ""
        if len(t) < 60:
            t = str(get(r, "grounds") or "")
        new_cat = CF.classify(t, "交易所")
        old_cat = re.sub(r"（次要.*", "", str(get(r, "category") or ""))
        if new_cat == old_cat:
            continue
        n += 1
        say("行%d %s：[%s] → [%s]" % (r["row"], str(get(r, "subject"))[:20],
                                     old_cat[:16], new_cat))
        newdir = EXCH_TOP + "/" + new_cat
        for i, f in enumerate(files):
            d = dirs[i] if i < len(dirs) else dirs[0]
            if d != newdir:
                move_pdf(root, d, newdir, f)
        put(r, "category", new_cat)
        put(r, "paths", newdir)   # L 列只写 1 条（同一案例文书必同目录，2026-11-02 用户口径）
    print("  重分类 %d 条" % n)

    # 清掉 03 顶层下不在 EXCH_CATS 内的空目录（旧六分法误建）+ 补齐合法目录
    base = Path(root) / EXCH_TOP
    if base.is_dir():
        legal = set(CF.EXCH_CATS)
        for d in sorted(base.iterdir()):
            if d.is_dir() and d.name not in legal:
                left = [f for f in d.iterdir() if f.is_file()]
                if left:
                    print("  [WARN] 非法目录仍有 %d 个文件，未删除：%s" % (len(left), d.name))
                else:
                    say("删除误建空目录 %s" % d.name)
                    if not DRY:
                        d.rmdir()
        if not DRY:
            for c in CF.EXCH_CATS:
                (base / c).mkdir(parents=True, exist_ok=True)
    return n


# ---------------------------------------------------------------- ② 占位命名回填
def entity(t):
    """从 PDF 正文抽主体：把正文开头当标题传入。

    PDF 没有「列表标题」这种外部元数据，但文书首页第一行就是标题
    （「关于对吴明采取出具警示函措施的决定」）。不传标题会丢掉 title_entities
    这条最可靠的路径，导致人名/机构名全空、回退到行级主体（行82 实测）。
    """
    return BP.entity_from_text(t or "", (t or "")[:120])


def _derive(root, d, files):
    mgr, persons = "", []
    for f in files:
        p = find_pdf(root, d, f)
        if p is None:
            continue
        m, ps = entity(doc_text(p))
        if m and not mgr:
            mgr = m
        for x in ps:
            if x not in persons:
                persons.append(x)
    return mgr, persons


def phase_refill(root, rows):
    print("\n=== ② 占位主体回填与重命名 ===")
    n = 0
    for r in rows:
        files = split_multi(get(r, "filenames"))
        dirs = split_multi(get(r, "paths"))
        subj = str(get(r, "subject") or "")
        if not files:
            continue
        urls0 = split_multi(get(r, "urls"))
        if len({_bureau(u) for u in urls0}) > 1:
            print("  [跳过] 行%d 是跨局合并行，交由 split 阶段拆分后再命名" % r["row"])
            continue
        dirty = any(bad_name(f) for f in files)
        dirty = dirty or any(k in subj for k in ("不明确", "待补")) or bool(
            re.search(r"[；、]缺位", subj))
        if "；" in subj:
            dirty = dirty or any(place_person(x) for x in subj.split("；", 1)[1].split("、"))
        if not dirty:
            continue
        d0 = dirs[0] if dirs else ""
        mgr, persons = _derive(root, d0, files)
        n += 1

        # --- 主体列：只在被污染时重算；否则保留原件（OCR 会把「贲金锋」认成「责金锋」）---
        polluted = (any(k in subj for k in ("不明确", "待补"))
                    or bool(re.search(r"[；、]缺位", subj)))
        old_persons = []
        if "；" in subj:
            old_persons = [x.strip() for x in subj.split("；", 1)[1].split("、") if x.strip()]
        if mgr and (polluted or any(place_person(x) for x in old_persons)
                    or old_persons == ["之一"]):
            merged = []
            for x in list(persons) + old_persons:
                if (x and x not in merged and x not in ("缺位", "不明确", "待补", "之一")
                        and not place_person(x)):
                    merged.append(x)
            newsubj = BP.subject_of(mgr, merged, "")
            say("行%d 主体：%r → %r" % (r["row"], subj[:26], newsubj[:34]))
            put(r, "subject", newsubj)

        # --- 文件名：逐个按该文书自身内容重建 ---
        for i, f in enumerate(files):
            if not bad_name(f):
                continue
            d = dirs[i] if i < len(dirs) else d0
            p = find_pdf(root, d, f)
            if p is None:
                continue
            d_real = str(p.parent.relative_to(root)).replace("\\", "/")
            m1, p1 = entity(doc_text(p))
            nn = new_name_for(p, m1 or mgr, p1)
            if not nn or nn.startswith(".pdf"):
                print("  [WARN] 行%d 无法确定新文件名，跳过：%s" % (r["row"], f))
                continue
            files[i] = rename_pdf(root, d_real, f, nn)
            dirs[i if i < len(dirs) else 0] = d_real
        put(r, "filenames", "；".join(files))

        # --- 处罚措施 / 证据：仍是占位时用正文重抽（「你公司」→ 实名）---
        fp0 = find_pdf(root, d0, files[0]) if files else None
        t0 = doc_text(fp0) if fp0 else ""
        if "不明确" in str(get(r, "measures") or "") and len(t0) > 60:
            ms = EC.extract_with_contract(t0, "csrc")["measures"]
            ms = ms.replace("你公司", mgr or "").replace("你营业部", mgr or "")
            meas = BP.split_measures(ms, mgr, [p for p in persons if p not in PLACEHOLDER])
            if meas:
                say("行%d 处罚措施回填：%s" % (r["row"], "；".join(meas)[:70]))
                put(r, "measures", "；".join(meas))
        if (not str(get(r, "evidence") or "").strip()
                or "未列举" in str(get(r, "evidence") or "")) and len(t0) > 60:
            ev = EC.extract_with_contract(t0, "csrc")["evidence"]
            if ev:
                put(r, "evidence", ev)

        # --- 人工核定覆盖（仅交易所监管函 2 条，见 MANUAL_FIX 注释）---
        for u in split_multi(get(r, "urls")):
            fix = MANUAL_FIX.get(u.strip())
            if not fix:
                continue
            for k, v2 in fix.items():
                put(r, k, v2)
            say("行%d 按人工核定覆盖 %s" % (r["row"], "、".join(fix)))
    print("  回填 %d 条" % n)
    return n


# ---------------------------------------------------------------- ③ 拆分跨局合并行
def _bureau(url):
    m = BUREAU_RE.search(url or "")
    if m:
        return m.group(1)
    for host, tag in (("szse.cn", "szse"), ("sse.com.cn", "sse"), ("bse.cn", "bse")):
        if host in (url or ""):
            return tag
    return "?"


def phase_split(root, rows):
    print("\n=== ③ 拆分「一行多局」的错误合并 ===")
    new_rows, n_split = [], 0
    for r in rows:
        files, dirs, urls = (split_multi(get(r, "filenames")),
                             split_multi(get(r, "paths")), split_multi(get(r, "urls")))
        if len({_bureau(u) for u in urls}) < 2 or len(files) < 2:
            continue
        n_split += 1
        groups = {}
        for i, f in enumerate(files):
            d = dirs[i] if i < len(dirs) else dirs[0]
            p = find_pdf(root, d, f)
            t = doc_text(p) if p else ""
            m, ps = entity(t)
            dec = sign_date(t, str(get(r, "decided") or ""))
            # 归组键 = (局, 落款日期)：同一局同日多份文书是同一案例的机构+高管分解，
            # 合并一行（与「判重=(局,管理人,时间)」口径一致）；不同局必然不同案例。
            key = (_bureau(urls[i] if i < len(urls) else ""), dec)
            g = groups.setdefault(key, {"files": [], "dirs": [], "urls": [], "mgr": "",
                                        "_best": 0, "persons": [], "text": "", "per": {}})
            g["files"].append(f)
            g["dirs"].append(d)
            g["urls"].append(urls[i] if i < len(urls) else "")
            g["per"][f] = (m, ps, p)
            for x in ps:
                if x not in g["persons"]:
                    g["persons"].append(x)
            if len(t) > len(g["text"]):
                g["text"] = t
            if m and len(t) >= g["_best"]:
                g["mgr"], g["_best"] = m, len(t)
        if len(groups) < 2:
            print("  [提示] 行%d 虽跨局但按(局,落款)归为 1 组，跳过" % r["row"])
            continue
        say("行%d %s：拆成 %d 行（局=%s）" % (
            r["row"], str(get(r, "subject"))[:20], len(groups),
            ",".join(sorted({k[0] for k in groups}))))
        for idx, (key, g) in enumerate(groups.items()):
            v = list(r["v"])
            ni = LS.COL
            v[ni["decided"]] = key[1]
            v[ni["subject"]] = BP.subject_of(
                g["mgr"], [p for p in g["persons"] if not place_person(p)], "")
            src = EC.extract_with_contract(g["text"], "csrc") if len(g["text"]) > 60 else {}
            meas = BP.split_measures(
                (src.get("measures") or "").replace("你公司", g["mgr"] or ""),
                g["mgr"], g["persons"]) if src else []
            v[ni["measures"]] = "；".join(meas) or str(get(r, "measures") or "")
            v[ni["grounds"]] = (src.get("grounds") or str(get(r, "grounds") or ""))[:2000]
            v[ni["evidence"]] = src.get("evidence") or str(get(r, "evidence") or "")
            v[ni["legal_basis"]] = src.get("legal_basis") or str(get(r, "legal_basis") or "")
            v[ni["defense"]] = src.get("defense") or str(get(r, "defense") or "无")
            # 正文 ≥200 字才重新分类（单页警示函关键词密度不足，reclass 会随机漂移）
            cat = (CF.classify(g["text"], "证监局") if len(g["text"]) >= 200
                   else re.sub(r"（次要.*", "", str(get(r, "category") or "")))
            v[ni["category"]] = cat
            v[ni["source"]] = get(r, "source")
            # 文件名按本组文书自身内容重建（原行文件名多已残缺：投资管理有限公司.pdf）
            names = []
            for f in g["files"]:
                m1, p1, pth = g["per"][f]
                nn = new_name_for(pth, m1 or g["mgr"], p1) if bad_name(f) else f
                names.append(nn or f)
            # 先把文件搬到本组分类目录（用旧名搬），再在原地改成新名——顺序不能颠倒，
            # 否则「按新名去搬旧文件」必然 assert 失败（2026-10-29 实测）。
            src_top = g["dirs"][0].split("/")[0]
            want = src_top + "/" + cat
            tgt = list(g["dirs"])
            for i2, f2 in enumerate(g["files"]):
                cur, cur_dir = f2, g["dirs"][i2]
                if want != cur_dir:
                    cur = move_pdf(root, cur_dir, want, cur)
                    tgt[i2] = want
                if names[i2] and names[i2] != cur:
                    if DRY and want != cur_dir:
                        # dry-run 下 move_pdf 不会真的搬文件，改名也就无从落地：
                        # 直接把「移动+改名」合并记入计划，避免 assert 误报。
                        say("并改名为 %s" % names[i2][:44])
                        cur = names[i2]
                    else:
                        cur = rename_pdf(root, want if want != cur_dir else cur_dir,
                                         cur, names[i2])
                names[i2] = cur
            v[ni["filenames"]] = "；".join(names)
            v[ni["paths"]] = "；".join(tgt)
            v[ni["urls"]] = "；".join(g["urls"])
            v[ni["typical"]] = "否"
            if idx == 0:
                r["v"] = v
            else:
                new_rows.append({"row": None, "v": v})
    print("  拆分 %d 行 → 新增 %d 行" % (n_split, len(new_rows)))
    return new_rows


# ---------------------------------------------------------------- ⑤ 主体名与正文一致性
def phase_check_name(root, rows):
    """主体机构名必须在决定书正文里原样出现；不出现 → 用正文全称替换。

    典型症状：台账写「海南海德资本股份有限公司」，正文是「海南海德资本管理股份
    有限公司」（少「管理」二字）——数量校验永远发现不了，但人一眼就觉得不对。
    只处理**有文字层**的 PDF（扫描件跳过，避免 235 次 OCR）。
    """
    print("\n=== ⑤ 主体名与正文一致性 ===")
    n = 0
    for r in rows:
        subj = str(get(r, "subject") or "")
        if not subj or any(k in subj for k in ("不明确", "待补")):
            continue
        inst = subj.split("；")[0].strip()
        if not inst:
            continue
        files, dirs = split_multi(get(r, "filenames")), split_multi(get(r, "paths"))
        if not files:
            continue
        p = find_pdf(root, dirs[0] if dirs else "", files[0])
        t = txt_of(p) if p else ""
        if len(t) < 60:
            continue
        flat = re.sub(r"\s+", "", t)
        if inst in flat:
            continue
        m, _ = entity(t)
        # 正文里的半角括号与台账的全角括号只是排版差异（元德坤/中投汇恒实测），
        # 统一成全角再比，否则会被误判成「主体不一致」而去改原本正确的名字。
        m = (m or "").replace("(", "（").replace(")", "）")
        if not (m and m != inst and m in flat.replace("(", "（").replace(")", "）")):
            continue
        # 允许「漏字型」差异（海南海德资本[管理]股份有限公司）：单靠子串包含关系
        # 抓不到，用相似度兜底；0.8 以上才认，避免把不相干的机构名替换进来。
        import difflib
        if difflib.SequenceMatcher(None, inst, m).ratio() < 0.8:
            continue
        n += 1
        say("行%d 主体补全：%s → %s" % (r["row"], inst[:26], m[:30]))
        newsubj = m + subj[len(inst):]
        put(r, "subject", newsubj)
        for i, f in enumerate(files):
            if not bad_name(f) and inst not in f:
                continue
            d = dirs[i] if i < len(dirs) else (dirs[0] if dirs else "")
            d = str((find_pdf(root, d, f) or Path(root) / d / f).parent
                    .relative_to(root)).replace("\\", "/")
            files[i] = rename_pdf(root, d, f, f.replace(inst, m))
        put(r, "filenames", "；".join(files))
    print("  补全 %d 条" % n)
    return n


# ---------------------------------------------------------------- ④b 碎片词清洗
def phase_purge_noise(root, rows):
    """清洗「被当成人名的正文碎片」：至今/职位/签章/离职/以前/之后/挂名/之一/的请求…

    2026-10-29 库内实测 9 行（用户实拍第 2 类缺陷的同类问题）：
      D 列：`浙江优策…；李晓晴、李薇薇、离职、的请求、焦健、黄巍`
      文件名：`职位-上海通赛投资有限公司.pdf`、`签章-广东横琴均成…pdf`、`至今-前海闽荣…pdf`
      E 列：`高管（至今）：对前海闽荣进行警告`（其实是管理人措施，被挂在碎片名下）
    根因：`subject_clean.is_valid_subject` 把「2~8 个汉字」一律认作自然人。
    本阶段做三件事：① D 列剔除碎片；② E 列把「高管（碎片）」按措施指向改回「管理人：」或丢弃；
    ③ 文件名去掉「碎片-」前缀（不撞名才改）。
    """
    print("\n=== ④b 主体/文件名碎片词清洗 ===")
    n = 0
    for r in rows:
        subj = str(get(r, "subject") or "").strip()
        if not subj:
            continue
        segs = [s.strip() for s in re.split(r"[；;]", subj) if s.strip()]
        inst = segs[0] if segs else ""
        persons_raw = []
        for s in segs[1:]:
            persons_raw += [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]
        dirty = [x for x in persons_raw if SC.is_noise_token(x)]
        # 「X+虚词」型重复碎片（杜胜于 vs 杜胜）不算 dirty 判定依据，但 drop 会清掉
        persons = SC.drop_noise_tokens(persons_raw, keep={inst})
        file_fix = [f for f in split_multi(get(r, "filenames"))
                    if re.match(r"^([^\-—]+)[\-—]", f) and SC.is_noise_token(re.match(r"^([^\-—]+)[\-—]", f).group(1))]
        if not dirty and len(persons) == len(persons_raw) and not file_fix:
            continue
        n += 1
        say("行%d 碎片 %s" % (r["row"], "、".join(dirty) or "（虚词重复）"))

        # ① D 列重组
        parts = []
        if inst and not SC.is_noise_token(inst):
            parts.append(inst)
        if persons:
            parts.append("、".join(persons))
        if parts:
            put(r, "subject", "；".join(parts))
        say("      D：%s → %s" % (subj[:44], "；".join(parts)[:44]))

        # ② E 列：高管（碎片）→ 管理人（措施指向机构时）／丢弃（指向个人时）
        meas = str(get(r, "measures") or "")
        if meas and dirty:
            out, changed = [], False
            for seg in [s.strip() for s in re.split(r"[；;]", meas) if s.strip()]:
                m = re.match(r"^高管[（(](.+?)[）)][：:](.*)$", seg)
                if m and SC.is_noise_token(m.group(1)):
                    body = m.group(2).strip()
                    changed = True
                    if not any(p in body for p in persons):
                        out.append("管理人：" + body)
                        say("      E：「高管（%s）」→「管理人」（措施指向机构）" % m.group(1))
                    else:
                        say("      E：丢弃「高管（%s）」（措施指向个人，主体不可考）" % m.group(1))
                    continue
                out.append(seg)
            if changed:
                put(r, "measures", "；".join(out))

        # ③ 文件名去「碎片-」前缀（同目录撞名则跳过，交给人工）
        files, dirs = split_multi(get(r, "filenames")), split_multi(get(r, "paths"))
        for i, f in enumerate(files):
            m = re.match(r"^([^\-—]+)[\-—](.+\.pdf)$", f)
            if not (m and SC.is_noise_token(m.group(1))):
                continue
            d = dirs[i] if i < len(dirs) else (dirs[0] if dirs else "")
            new = m.group(2)
            if find_pdf(root, d, new) is not None:
                say("      跳过去前缀（%s 已存在同名）：%s" % (new[:28], f[:34]))
                continue
            say("      改名 %s → %s" % (f[:34], new[:34]))
            files[i] = rename_pdf(root, d, f, new)
        put(r, "filenames", "；".join(files))
    print("  清洗 %d 行" % n)
    return n


# ---------------------------------------------------------------- ⑥ 序号后缀清理
def phase_tidy_suffix(root, rows):
    """把「X（2）.pdf」这类孤立序号名改回「X.pdf」。

    成因：早期批量重命名给同名文件加（2）后，基名文件被移去别的目录/被改名，
    目录里只剩带序号那份 → 用户看到「（2）」却找不到「（1）」，属命名噪声。
    只在同目录内**不存在**基名时改，改完同步台账「案例文件」列。
    """
    print("\n=== ⑤ 孤立序号名清理 ===")
    mapping = {}          # (dir, old) -> new
    for d1 in sorted(os.listdir(root)):
        if not (os.path.isdir(os.path.join(root, d1)) and d1[:2] in ("01", "02", "03")):
            continue
        for dp, _, fs in os.walk(os.path.join(root, d1)):
            rel = os.path.relpath(dp, root).replace(os.sep, "/")
            names = set(fs)
            for f in sorted(fs):
                m = re.match(r"^(.*)（\d+）\.pdf$", f)
                if m and (m.group(1) + ".pdf") not in names:
                    mapping[(rel, f)] = m.group(1) + ".pdf"
    if not mapping:
        print("  无需清理")
        return 0
    for (d, old), new in mapping.items():
        say("去序号 %s：%s → %s" % (d.split("/")[-1][:14], old[:34], new[:34]))
        if not DRY:
            os.rename(os.path.join(root, d, old), os.path.join(root, d, new))
    # 同步台账
    for r in rows:
        files = split_multi(get(r, "filenames"))
        dirs = split_multi(get(r, "paths"))
        changed = False
        for i, f in enumerate(files):
            key = (dirs[i] if i < len(dirs) else (dirs[0] if dirs else ""), f)
            if key in mapping:
                files[i] = mapping[key]
                changed = True
        if changed:
            put(r, "filenames", "；".join(files))
    print("  清理 %d 个孤立序号名" % len(mapping))
    return len(mapping)


# ---------------------------------------------------------------- ④ 典型案例填充
TYPICAL_RULES = {
    "募集宣传自媒体违规案例": (
        r"(公众号|微信|朋友圈|互联网|网站|直播|短视频|抖音|微博|自媒体).{0,40}(宣传|推介|募集)"
        r"|(宣传|推介|募集).{0,40}(公众号|微信|朋友圈|互联网|网站|直播|短视频|抖音|微博|自媒体)",
        "通过公众号/微信/朋友圈/互联网/直播/短视频等自媒体公开宣传推介私募产品"),
}


def phase_typical(root, rows):
    print("\n=== ④ 05 重点典型案例目录填充 ===")
    added = 0
    for sub, (pat, desc) in TYPICAL_RULES.items():
        d = Path(root) / "05_重点典型案例" / sub
        have = {f.name for f in d.iterdir()} if d.is_dir() else set()
        hits = []
        for r in rows:
            if str(get(r, "typical") or "").strip() == "是":
                continue
            blob = re.sub(r"\s+", "", str(get(r, "grounds") or "") + str(get(r, "measures") or ""))
            if not re.search(pat, blob):
                continue
            files, dirs = split_multi(get(r, "filenames")), split_multi(get(r, "paths"))
            if files:
                hits.append((r, files, dirs))
        print("  %s：命中 %d 条（判定标准：%s）" % (sub, len(hits), desc))
        for r, files, dirs in hits:
            for i, f in enumerate(files):
                p = find_pdf(root, dirs[i] if i < len(dirs) else dirs[0], f)
                if p is None or f in have:
                    continue
                if DRY:
                    say("复制 %s → 05/%s/" % (f[:36], sub))
                else:
                    d.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(str(p), str(d / f))
                have.add(f)
            put(r, "typical", "是")
            added += 1
            print("      行%d %s" % (r["row"], str(get(r, "subject"))[:34]))
    print("  标记典型案例 %d 条" % added)
    return added


# ---------------------------------------------------------------- 落盘
def save(root, rows, new_rows):
    if DRY:
        return
    wb = load_workbook(os.path.join(root, LS.LEDGER_NAME))
    ws = wb[LS.SHEET_CASE]
    for r in rows:                      # 原地写回值（不碰任何样式）
        for c, v in enumerate(r["v"], start=1):
            ws.cell(row=r["row"], column=c).value = v
    for r in new_rows:                  # 追加新行（沿用上一行样式）
        LS.append_row(ws, r["v"], LS.CASE_WRAP_COLS)
    LS.renumber(ws)
    wb.save(os.path.join(root, LS.LEDGER_NAME))
    print("  已保存台账（原 %d 行 + 新增 %d 行）" % (len(rows), len(new_rows)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--only", default="reclass,refill,split,purge,typical,names,tidy")
    ap.add_argument("--apply", action="store_true", help="真正落盘（默认 dry-run）")
    a = ap.parse_args()
    global DRY
    DRY = not a.apply
    steps = [s.strip() for s in a.only.split(",") if s.strip()]

    root = a.root
    assert os.path.isdir(root), root
    print("库根：%s" % root)
    print("模式：%s" % ("DRY-RUN（只打印计划，不改文件）" if DRY else "APPLY（真正修改）"))

    if not DRY:
        bak = os.path.join(LS.work_dir(root), LS.LEDGER_NAME + ".bak_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
        os.makedirs(LS.work_dir(root), exist_ok=True)
        if not os.path.exists(os.path.join(root, LS.LEDGER_NAME)):
            sys.exit("台账不存在：%s" % os.path.join(root, LS.LEDGER_NAME))
        shutil.copy2(os.path.join(root, LS.LEDGER_NAME), bak)
        print("  已备份台账 → %s" % bak)

    _, _, rows = read_rows(root)
    print("台账数据行：%d" % len(rows))
    new_rows = []
    if "reclass" in steps:
        phase_reclass(root, rows)
    if "refill" in steps:
        phase_refill(root, rows)
    if "split" in steps:
        new_rows += phase_split(root, rows)
    if "purge" in steps:
        phase_purge_noise(root, rows)
    if "typical" in steps:
        phase_typical(root, rows)
    if "names" in steps:
        phase_check_name(root, rows)
    if "tidy" in steps:
        phase_tidy_suffix(root, rows)
    save(root, rows, new_rows)
    print("\n完成。修复后请依次执行：")
    print("  python ledger.py --root \"%s\" sort" % root)
    print("  python verify_library.py --root \"%s\"" % root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
