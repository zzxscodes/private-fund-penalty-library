# -*- coding: utf-8 -*-
"""把「一行捆了多个案例」的台账行拆成独立案例行（2026-10-30）。

用户口径（本脚本的验收标准）：
  ① 一行 = 一个案例；判定键 = (证监局/来源局, 机构, 处罚时间)。跨局、跨批次一律拆开；
  ② **人员函必须归入正确的那一个机构函案例**——不得把 A 案的处罚措施错记到 B 案行上。
     归属规则：同局优先，落款最接近者优先（人员函一般与最近一批机构函同案）；
     平局时优先「机构函落款 ≤ 人员函落款」的那个（先罚机构、后出人员函是常态）。
  ③ 同一批次内的机构函 + 人员函仍合成一行（机构在前、人员在后的 v2 格式不变）。

判定「需要拆」：把行内每份文件按 (局, 落款) 分组，分组数 > 1 即视为捆包行。
落款取法：文字层最后一个可解析日期（sign_date.parse_dates，唯一事实源）；
          无文字层（扫描件）走 OCR（sign_date.find_sign）；都失败退回台账处罚时间。

用法：
    python split_rows.py --root "<库根>"                  # 体检 + 打印拆分方案
    python split_rows.py --root "<库根>" --apply          # 落盘（先备份台账 + 备份文件）
    [--rows 23,44,54]     只处理指定台账序号
    [--max-gap 30]        行内落款跨度容忍阈值（天），默认 >30 天才判捆包

退出码：0 = 无待拆行或全部成功；1 = 有行拆分失败。
"""
import argparse, json, os, re, shutil, sys
from datetime import date, datetime
import ledger_style as LS  # noqa: E402  运维目录唯一真源（maint_dir/work_dir/...）

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import fitz  # noqa: E402
import sign_date as SD  # noqa: E402
import extract_core  # noqa: E402
from classify_filter import classify  # noqa: E402
import build_pipeline as BP  # noqa: E402

LEDGER = "04_综合索引台账【总表】.xlsx"
SHEET = "综合索引台账"
CSRC_DIR = "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）"


def split_multi(v):
    if v is None:
        return []
    return [x.strip() for x in re.split(r"[；;]", str(v).strip()) if x.strip()]


def bureau_of(url):
    m = re.match(r"https?://www\.csrc\.gov\.cn/([a-z]+)/", url or "")
    if m:
        return m.group(1)
    u = url or ""
    if "szse" in u:
        return "szse"
    if "sse.com.cn" in u:
        return "sse"
    if "bse.cn" in u:
        return "bse"
    if "amac.org.cn" in u:
        return "amac"
    return ""


# ---------------------------------------------------------------- 正文 / 落款
_OCR = None


def _ocr():
    global _OCR
    if _OCR is None:
        import numpy as np
        from rapidocr_onnxruntime import RapidOCR
        _OCR = (RapidOCR(), np)
    return _OCR


def ocr_pdf(path, max_pages=8):
    """整份 OCR（限页）：扫描件抽事由/措施必须靠它，只 OCR 末页拿不到正文。"""
    ocr, np = _ocr()
    d = fitz.open(path)
    out = []
    try:
        for i, pg in enumerate(d):
            if i >= max_pages:
                break
            zoom = 200 / 72.0
            if pg.rect.width * zoom > 2000:
                zoom *= 2000 / (pg.rect.width * zoom)
            pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if img.ndim == 3 and img.shape[2] >= 3:
                R = np.ascontiguousarray(img[:, :, 0])
                img = np.stack([R, R, R], axis=2)
            r, _ = ocr(img)
            if r:
                out.append(" ".join(t[1] for t in r))
    finally:
        d.close()
    return "\n".join(out)


def file_text(path):
    """取正文（文字层优先；扫描件才整份 OCR，限 8 页）。只在确需拆行时调用。"""
    d = fitz.open(path)
    try:
        t = "".join(p.get_text() for p in d)
        npg = d.page_count
    finally:
        d.close()
    if len(re.sub(r"\s", "", t)) < 60 and npg > 0:
        t = ocr_pdf(path)
    return t


def light_text(path):
    """只取文字层（不 OCR）——用于廉价的"要不要拆"初判。"""
    try:
        d = fitz.open(path)
        t = "".join(p.get_text() for p in d)
        d.close()
        return t
    except Exception:
        return ""


def file_date(path, text, fallback):
    ds = SD.parse_dates(text, max_date=SD.today_iso())
    if ds:
        return ds[-1], "text"
    try:
        dt, _plan, _raw = SD.find_sign(path)
        if dt:
            return dt, "ocr"
    except Exception:
        pass
    return fallback, "fallback"


def fast_date(path, cached_text, fallback):
    """阶段1 的廉价落款判定（禁止无谓 OCR）。

    优先级：抓取期缓存的正文（csrc.json/amac.json 的 text，命中率 778/790）→ 文件文字层
    → 最后才 OCR。
    关键：**只要拿到了 ≥60 字的正文，抽不到落款就立即用台账日期兜底，不再 OCR**——
    因为「正文里没有落款」是网页版决定书的常态（证监局网页版普遍只给文号），
    OCR 也变不出落款，白白烧掉几十秒（实测序号15/16/34 单行耗 40~122 秒全耗在这）。
    只有正文完全拿不到（真扫描件且无缓存）才走 OCR 定位落款。
    """
    for src in (cached_text, None):
        if src is None:
            src = light_text(path)
        if src and len(re.sub(r"\s", "", src)) >= 60:
            ds = SD.parse_dates(src, max_date=SD.today_iso())
            return (ds[-1], "text") if ds else (fallback, "no-sign")
    dt, _p, _r = SD.find_sign(path)
    return (dt, "ocr") if dt else (fallback, "fallback")


def is_person_file(fn):
    core = re.sub(r"\.pdf$", "", fn, flags=re.I)
    if "-" not in core:
        return False, ""
    head = core.split("-")[0].strip()
    return (2 <= len(head) <= 6), head


def days(a, b):
    try:
        return abs((date.fromisoformat(a) - date.fromisoformat(b)).days)
    except Exception:
        return 9999


# ---------------------------------------------------------------- 拆行
def analyze_row(root, row, max_gap, textmap=None):
    textmap = textmap or {}
    files, paths, urls = row["files"], row["paths"], row["urls"]
    if len(files) < 2:          # 单文件行无从拆起，直接跳过（不做任何 IO）
        return None
    items = []
    for i, fn in enumerate(files):
        d = paths[i] if i < len(paths) else (paths[0] if paths else "")
        u = urls[i] if i < len(urls) else (urls[0] if urls else "")
        fp = os.path.join(root, d, fn)
        if not os.path.exists(fp):
            continue
        dt, how = fast_date(fp, textmap.get(u, ""), row["date"])
        per, pname = is_person_file(fn)
        items.append({"fn": fn, "path": d, "url": u, "bureau": bureau_of(u),
                      "date": dt, "date_how": how, "is_person": per,
                      "pname": pname if per else "", "text": "", "fp": fp})
    if len(items) < 2:
        return None
    ds = sorted(set(x["date"] for x in items if x["date"]))
    span = days(ds[0], ds[-1]) if len(ds) > 1 else 0
    groups = {}
    inst = [x for x in items if not x["is_person"]]
    pers = [x for x in items if x["is_person"]]
    for x in inst:
        groups.setdefault((x["bureau"], x["date"]), []).append(x)
    for p in pers:
        cand = [k for k in groups if k[0] == p["bureau"]] or list(groups)
        if not cand:
            groups.setdefault((p["bureau"], p["date"]), []).append(p)
            continue
        best = min(cand, key=lambda k: (days(k[1], p["date"]), k[1] > p["date"]))
        # 2026-10-30 关键修正（用户明确提醒的错配风险）：
        #   同局没有机构函时，原来的 `or list(groups)` 会把人员函硬塞给**别局**的机构函，
        #   于是 A 案的处罚措施被记到 B 案行上（实测江苏百瑞赢：安徽局的汪金玉/李想被
        #   塞进陕西局 2026-06-23 那行）。正确做法：人员函与最近的机构函落款相差超过
        #   max_gap，就说明不是同一批——该人员函自成一案（这类函通常同时点名机构与人员，
        #   机构名从正文抽取，不会丢主体）。
        if days(best[1], p["date"]) > max_gap:
            groups.setdefault((p["bureau"], p["date"]), []).append(p)
        else:
            groups[best].append(p)
    if len(groups) < 2 and span <= max_gap:
        return None
    # 确需拆行 → 这时才补正文（扫描件才会 OCR）
    for x in items:
        # 必须走 build_pipeline.tidy()：去 HTML 实体 + 合并汉字间空格（"二 是"→"二是"）。
        # 少了这一步，extract_core 会把整段「经查…」当成处罚措施，事由/依据也抽不干净
        # （2026-10-30 实测：未 tidy 时措施列变成「管理人：<机构名>：经查…」）。
        x["text"] = BP.tidy(textmap.get(x["url"]) or file_text(x["fp"]))
    return {"groups": groups, "items": items, "span": span}


_CLAUSE_PROC = re.compile(r"^(如对|如果对|复议与诉讼期间|当事人如|你公司如)")


def split_clauses(t):
    return [c.strip(" ；;。") for c in re.split(r"[；;。]", t or "") if c.strip(" ；;。")]


def clean_html_entities(s):
    """清掉网页正文带来的 HTML 实体（&ensp; 等）—— 2026-10-30 实测拆行后措施列出现
    「…整改报告。&ensp;&ensp;&ensp;」这类垃圾，是内联 HTML 抽取的残留。"""
    if not s:
        return s
    s = re.sub(r"&(ensp|nbsp|emsp|thinsp|amp|lt|gt|quot|apos|middot|hellip|ldquo|rdquo);",
               lambda m: {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'",
                          "middot": "·", "hellip": "…", "ldquo": "“", "rdquo": "”"
                          }.get(m.group(1), " "), s, flags=re.I)
    s = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), s)
    return re.sub(r"[ \t\u3000]{2,}", " ", s).strip()


_HTML_FOOT = re.compile(r"(链接[:：]\s*中国政府网|主办单位[:：]|网站识别码|京ICP备|京公网安备|"
                        r"【打印】|行业相关网站|其他链接).*$", re.S)
_MEAS_TAIL = re.compile(r"[。；]?\s*(如对(本|上述|以上)?(监督)?管理措施不服|"
                        r"如果对(本|上述|以上)?(监督)?管理措施不服|复议与诉讼期间).*$", re.S)


def doc_measures(x):
    """单份文书抽「处罚措施」（**逐份抽，绝不跨文书合并**）。

    2026-10-30 定稿：tidy() 之后 extract_core 对各类文书的措施抽取质量都很好
    （机构函「对你公司采取责令改正…」、人员函「对你采取出具警示函…」、
      交易所函「对你公司予以书面警示」、处罚决定书「对××责令改正，给予警告，并处…」）。
    必须逐份抽而不是把整行原文切条回查——原台账的措施列是多案混合且已被改写，
    回查匹配率低，兜底会把 A 案的条文塞进 B 案行（实测海南港澳 5 案互相串味）。
    这里只做两件清洗：截掉网页页脚、截掉「如不服可复议/诉讼」的程序告知。
    """
    s = extract_core.extract_with_contract(x["text"], "csrc").get("measures") or ""
    s = clean_html_entities(BP.tidy(s))
    s = _HTML_FOOT.sub("", s)
    s = _MEAS_TAIL.sub("", s)
    return s.strip(" ；。;,")


def measures_for_group(inst, pers):
    """组内措施：机构函→管理人；人员函→该人员；人员函若同时点名机构则记管理人。"""
    mgr_parts, per_parts = [], {p["pname"]: [] for p in pers}
    for f in inst:
        m = doc_measures(f)
        if m and m not in mgr_parts:
            mgr_parts.append(m)
    for p in pers:
        m = doc_measures(p)
        if not m:
            continue
        if re.search(r"对你(公司|单位|分公司|企业|所)", m):
            if m not in mgr_parts:
                mgr_parts.append(m)
        elif p["pname"] and m not in per_parts[p["pname"]]:
            per_parts[p["pname"]].append(m)
    out = []
    if mgr_parts:
        out.append("管理人：" + "；".join(mgr_parts))
    for p in pers:
        if per_parts.get(p["pname"]):
            out.append("高管（%s）：%s" % (p["pname"], "；".join(per_parts[p["pname"]])))
    return out


def _norm_key(s, n=24):
    s = re.sub(r"&[a-zA-Z]+;", "", s or "")
    s = re.sub(r"[\s\u3000；;。，,、：:（）()“”\"'【】\[\]…·]", "", s)
    return s[:n]


def assign_measures(row_measures, gf, mgr, pnames):
    """从**原台账的措施列**按文书归属切分（不重抽）。

    为什么不能重抽：extract_core 对「行政处罚决定书」「交易所监管函」这类非标准
    「我局决定…」结构文书常常抓偏（实测把整段「当事人：…／经查…」当成措施，或把
    「上述行为违反了…」当成措施）。而原台账的措施列是经过多轮校订的，质量更高。
    做法：把原措施按 `；。` 切成条款，逐条回查它出现在**哪一份文书正文**里——
    出现在哪份就归哪份的当事人：机构函→管理人、人员函→该人员。这样既保原文质量，
    又能把 A 案的措施准确落到 A 案行上（用户 2026-10-30 明确提醒的坑）。
    对回查不到原文的条款，再按「对你公司/你单位」→管理人、「对你」＋单人→该人员 兜底。
    """
    clauses = split_clauses(row_measures or "")
    per = {p: [] for p in pnames}
    mgr_cl, un = [], []
    for c in clauses:
        if _CLAUSE_PROC.match(c):
            continue
        k = _norm_key(c)
        owners = [f for f in gf if k and k in _norm_key(f["text"], 10 ** 6)]
        hp = next((p for p in pnames if p and p in c), None)
        if hp and any(o.get("pname") == hp for o in owners):
            per[hp].append(c)
        elif owners:
            o = owners[0]
            if o["is_person"] and o.get("pname"):
                per[o["pname"]].append(c)
            else:
                mgr_cl.append(c)
        elif hp:
            per[hp].append(c)
        elif re.search(r"对你(公司|单位|分公司|企业|所)", c):
            mgr_cl.append(c)
        elif "对你" in c and len(pnames) == 1:
            per[pnames[0]].append(c)
        elif "对你" in c:
            mgr_cl.append(c)
        else:
            un.append(c)
    if not mgr_cl and not any(per[p] for p in pnames) and un:
        mgr_cl = un          # 组内一条都没匹配上：至少别把措施留空
    out = []
    if mgr_cl:
        out.append("管理人：" + "；".join(mgr_cl))
    for p in pnames:
        if per[p]:
            out.append("高管（%s）：%s" % (p, "；".join(per[p])))
    return out


def build_group_rows(root, row, groups, pubmap):
    """把分组结果转成台账行（15 列）。"""
    out = []
    for (bureau, dt), xs in sorted(groups.items(), key=lambda kv: kv[0][1]):
        inst = [x for x in xs if not x["is_person"]]
        pers = [x for x in xs if x["is_person"]]
        base = (inst or pers)[0]
        mgr, persons = BP.entity_from_text(base["text"], "")
        if not mgr:
            # 退化：用文件名里的机构段，并剥掉同名去重序号（否则会出现
            # 「海南港澳资讯产业股份有限公司（4）」这种带文件序号的主体名）
            mgr = re.sub(r"\.pdf$", "", (inst or pers)[-1]["fn"], flags=re.I)
            if "-" in mgr:
                mgr = mgr.split("-", 1)[1]
            mgr = re.sub(r"[（(]\s*\d+\s*[）)]$", "", mgr).strip()
        pnames = []
        for p in pers:
            nm = p["pname"] or ""
            if nm and nm not in pnames:
                pnames.append(nm)
        if pers:
            for p in BP.entity_from_text(" ".join(x["text"][:1500] for x in pers), "")[1]:
                if p not in pnames:
                    pnames.append(p)
        meas = measures_for_group(inst, pers)
        if not meas:
            # 兜底：所有文书都抽不出措施时，才回到「按原文切条 + 当事人归属」
            meas = assign_measures(row["measures"], inst + pers, mgr, pnames)
        if not meas:
            meas = ["管理人：不明确"]
        src = extract_core.extract_with_contract((inst or pers)[0]["text"], "csrc")
        text_all = (inst or pers)[0]["text"]
        cat = classify(text_all, row["source"]) if row["source"] == "证监局" else row["category"]
        fnames = [x["fn"] for x in (inst + pers)]
        fpaths = [x["path"] for x in (inst + pers)]
        furls = [x["url"] for x in (inst + pers)]
        # 分类变化 → 文件必须跟着搬到对应目录，否则台账「违规类型」与「存放路径」打架（verify C2）
        if cat and cat != row["category"] and row["source"] == "证监局":
            newdir = CSRC_DIR + "/" + cat   # 台账路径统一用 /（与全库存量一致，勿用 os.path.join 生成 \\）
            for i in range(len(fpaths)):
                src_fp = os.path.join(root, fpaths[i], fnames[i])
                dst_fp = os.path.join(root, newdir, fnames[i])
                if os.path.abspath(src_fp) == os.path.abspath(dst_fp):
                    fpaths[i] = newdir
                    continue
                if os.path.exists(src_fp):
                    os.makedirs(os.path.join(root, newdir), exist_ok=True)
                    shutil.move(src_fp, dst_fp)
                elif not os.path.exists(dst_fp):
                    # 两边都没有 → 保持原路径，交给 verify A1 报缺失（不静默改路径）
                    continue
                fpaths[i] = newdir
        pub = pubmap.get(furls[0]) or row["publish"]
        out.append({
            "date": dt, "publish": pub,
            "subject": BP.subject_of(mgr, pnames, src["measures"]) or row["subject"],
            "measures": clean_html_entities("；".join(meas)),
            "grounds": clean_html_entities(src["grounds"] or row["grounds"]),
            "evidence": clean_html_entities(src["evidence"] or row["evidence"]),
            "legal_basis": clean_html_entities(src["legal_basis"] or row["legal_basis"]),
            "defense": src["defense"] or row["defense"],
            "category": cat,
            "source": row["source"],
            "files": fnames, "paths": fpaths, "urls": furls,
            "typical": row["typical"],
            "bureau": bureau,
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--rows", default="")
    ap.add_argument("--max-gap", type=int, default=30)
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    from openpyxl import load_workbook
    lp = os.path.join(root, LEDGER)
    wb = load_workbook(lp)
    ws = wb[SHEET]
    head = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(head)}
    col = lambda n: idx[n]

    # 挂网日 + 正文映射（按 URL 精确取，避免拆出的行沿用原行的挂网日；正文复用抓取期抽取结果）
    pubmap, textmap = {}, {}
    for fn in ("csrc.json", "amac.json", "exch.json"):
        p = os.path.join(LS.work_dir(root), fn)
        if not os.path.exists(p):
            continue
        try:
            j = json.load(open(p, encoding="utf-8"))
        except Exception:
            continue
        for it in j.get("items", []):
            u = it.get("url") or ""
            d = (it.get("date_published") or it.get("publish") or "")[:10]
            if u and d and d != "待补":
                pubmap[u] = d
            if u and it.get("text"):
                textmap[u] = it["text"]

    only = set(int(x) for x in a.rows.split(",") if x.strip().isdigit())
    plans = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, col("被处罚机构/人")).value:
            continue
        seq = ws.cell(r, col("序号")).value
        if only and int(seq or 0) not in only:
            continue
        row = {"excel_row": r, "seq": seq,
               "date": str(ws.cell(r, col("处罚时间")).value or "")[:10],
               "publish": str(ws.cell(r, col("挂网发布日期")).value or "")[:10],
               "subject": str(ws.cell(r, col("被处罚机构/人")).value or ""),
               "measures": str(ws.cell(r, col("处罚措施（机构+人员）")).value or ""),
               "grounds": str(ws.cell(r, col("核心处罚事由")).value or ""),
               "evidence": str(ws.cell(r, col("核心证据材料")).value or ""),
               "legal_basis": str(ws.cell(r, col("核心处罚依据")).value or ""),
               "defense": str(ws.cell(r, col("申辩理由")).value or ""),
               "category": str(ws.cell(r, col("违规类型")).value or ""),
               "source": str(ws.cell(r, col("机构来源")).value or ""),
               "files": split_multi(ws.cell(r, col("案例文件")).value),
               "paths": split_multi(ws.cell(r, col("存放路径")).value),
               "urls": split_multi(ws.cell(r, col("原文链接")).value),
               "typical": ws.cell(r, col("是否典型案例")).value}
        import time as _t
        _t0 = _t.time()
        an = analyze_row(root, row, a.max_gap, textmap)
        _dt = _t.time() - _t0
        if _dt > 3.0:
            print("  [slow %.1fs] 序号%s %s" % (_dt, seq, row["subject"][:34]), flush=True)
        if not an:
            continue
        plans.append((row, an))
    print("待拆行 %d 行" % len(plans))
    if not plans:
        return 0
    total_new = 0
    for row, an in plans:
        print("\n=== 序号%s %s｜原处罚 %s｜落款跨度 %d 天"
              % (row["seq"], row["subject"][:44], row["date"], an["span"]))
        for (b, dt), xs in sorted(an["groups"].items(), key=lambda kv: kv[0][1]):
            kind = "机构" if any(not x["is_person"] for x in xs) else "仅人员"
            print("    · 局=%-9s 落款=%s (%s) 文件%d份: %s"
                  % (b, dt, kind, len(xs), "、".join(x["fn"][:22] for x in xs)))
        total_new += max(0, len(an["groups"]) - 1)
    print("\n合计将新增 %d 行（%d → %d）" % (total_new, 500, 500 + total_new))

    if not a.apply:
        print("\n（体检模式，未改动。加 --apply 落盘）")
        return 0

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bakdir = os.path.join(os.path.dirname(root), "_split_bak_" + stamp)
    os.makedirs(bakdir, exist_ok=True)
    shutil.copy2(lp, os.path.join(bakdir, LEDGER))

    # 从后往前处理，避免插行导致行号错位
    report = []
    for row, an in sorted(plans, key=lambda x: -x[0]["excel_row"]):
        newrows = build_group_rows(root, row, an["groups"], pubmap)
        if len(newrows) < 2:
            continue
        r = row["excel_row"]
        # 第一组覆盖原行
        first = newrows[0]
        vals = {"处罚时间": first["date"], "挂网发布日期": first["publish"],
                "被处罚机构/人": first["subject"], "处罚措施（机构+人员）": first["measures"],
                "核心处罚事由": first["grounds"], "核心证据材料": first["evidence"],
                "核心处罚依据": first["legal_basis"], "申辩理由": first["defense"],
                "违规类型": first["category"],
                "案例文件": "；".join(first["files"]), "存放路径": "；".join(first["paths"]),
                "原文链接": "；".join(first["urls"])}
        for k, v in vals.items():
            ws.cell(r, col(k)).value = v
        # 其余组插入到本行下方
        off = 0
        for extra in newrows[1:]:
            off += 1
            ws.insert_rows(r + off)
            rowvals = {"序号": None, "处罚时间": extra["date"],
                       "挂网发布日期": extra["publish"],
                       "被处罚机构/人": extra["subject"],
                       "处罚措施（机构+人员）": extra["measures"],
                       "核心处罚事由": extra["grounds"], "核心证据材料": extra["evidence"],
                       "核心处罚依据": extra["legal_basis"], "申辩理由": extra["defense"],
                       "违规类型": extra["category"], "机构来源": extra["source"],
                       "案例文件": "；".join(extra["files"]),
                       "存放路径": "；".join(extra["paths"]),
                       "原文链接": "；".join(extra["urls"]),
                       "是否典型案例": extra["typical"]}
            for k, v in rowvals.items():
                if k in idx:
                    ws.cell(r + off, col(k)).value = v
        report.append({"seq": row["seq"], "subject": row["subject"],
                       "from": 1, "to": len(newrows),
                       "detail": [{"date": x["date"], "subject": x["subject"],
                                   "bureau": x["bureau"], "files": x["files"]}
                                  for x in newrows]})
        print("  序号%s 拆为 %d 行" % (row["seq"], len(newrows)))

    # 重排序号 + 按挂网日倒序整理
    n = 0
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, col("被处罚机构/人")).value:
            continue
        n += 1
        ws.cell(r, col("序号")).value = n
    wb.save(lp)
    print("\n台账已保存：%d 行（备份 %s）" % (n, bakdir))
    rp = os.path.join(LS.work_dir(root), "split_report.json")
    json.dump(report, open(rp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("拆分报告：%s" % rp)
    return 0


if __name__ == "__main__":
    sys.exit(main())
