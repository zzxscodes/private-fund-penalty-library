# -*- coding: utf-8 -*-
"""案例库收尾三件事（此前 v3 构建全部漏做的部分）。

1. `--laws`  06_配套法规依据 的枚举 + 库内引用数统计 + 台账「配套法规依据清单」回填
              + 报告「正文点名引用 ≥3 次但 06 目录缺失」的法规（供 fetch_laws 补下载）。
2. `--typical` 05_重点典型案例 遴选与**交叉归档**：同一份决定书两处都留
              （05 子类 + 来源大类的违规实质子分类），并置台账「是否典型案例」= 是。

背景（2026-10-27 实测根因）
--------------------------
SKILL.md 第 140~142 行早已写明「05 两处留存」「案例与法规必须联动、缺一部补一部」，
但**没有任何脚本实现**：init_library 只建了 05 空目录，build_v3 只把 06 从快照复制
且未回填法规清单 sheet → 05 全空、法规清单 0 行、06 法规未真正检索下载。

用法:
    python finalize_library.py --root "<案例库根>" --laws
    python finalize_library.py --root "<案例库根>" --typical
    python finalize_library.py --root "<案例库根>" --laws --typical
"""
import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import ledger           # noqa: E402
import ledger_style as LS  # noqa: E402
from verify_library import ALIAS, NOISE_LAW  # noqa: E402  （法规简称归一，单一事实源）
import fitz             # noqa: E402

LAW_EXTS = (".docx", ".pdf")
CITE_RE = re.compile(r"《([^》]{2,60}?)》")
# 《》里不全是法规：决定书常引用「行政处罚事先告知书」「2021年年度报告」「私募基金合同」等文书。
# 直接把这些当成「缺失法规」会让操作员去下载根本不是法规的东西（2026-10-27 实测）。
DOC_NOISE = re.compile(r"(告知书|决定书|报告|合同|清单|说明|函|意见书|申请书|复核决定|"
                       r"承诺|计划|方案|报表|表格|协议|章程|制度|问卷|确认书|承诺函)")
LAW_SUFFIX = re.compile(r"(法|条例|办法|规定|指引|细则|规则|意见|通知|公告|准则|标准)$")


_PAREN_TAIL = re.compile(r"[（(](?:试行|暂行|修订|[^）)]*?修订|[^）)]*?年修订)[^）)]*[）)]\s*$")


def looks_like_law(name):
    """《》里的内容是否像一部法规（用于区分「真缺失」与「引用了文书」。

    注意：法规名常带「（试行）」「（2020年修订）」等括号后缀，直接判后缀会被
    "…实施指引（试行）" 这类**真法规**误判成非法规而漏报（2026-10-27）。
    另外：裸词如「实施细则」「管理办法」是《》截断碎片，不是真法规，按 NOISE_LAW 过滤。
    """
    n = (name or "").strip()
    if not n or DOC_NOISE.search(n):
        return False
    n = _PAREN_TAIL.sub("", n).strip()      # 先剥括号后缀再判后缀
    if n in NOISE_LAW:                      # 裸词碎片（实施细则/管理办法…）非真法规
        return False
    return bool(LAW_SUFFIX.search(n))
LAW_DIR = "06_以上各类所有案例的配套法规依据"
TYPICAL_DIR = "05_重点典型案例"
CURATION = os.path.join(HERE, "..", "references", "typical_cases.md")

# 主干法规元信息（供清单填报；未知法规留空由人工补）
METADATA = {
    "私募投资基金监督管理条例": ("国务院令第762号", "国务院", "2023-07-09", "2023-09-01", "现行有效"),
    "私募投资基金监督管理暂行办法": ("证监会令第105号", "中国证券监督管理委员会", "2014-08-21", "2014-08-21", "现行有效"),
    "关于加强私募投资基金监管的若干规定": ("证监会公告〔2020〕71号", "中国证券监督管理委员会", "2020-12-30", "2020-12-30", "现行有效"),
    "私募投资基金募集行为管理办法": ("中基协发〔2016〕7号", "中国证券投资基金业协会", "2016-04-15", "2016-07-15", "现行有效"),
    "私募投资基金信息披露管理办法": ("中基协（2016年2月4日发布）", "中国证券投资基金业协会", "2016-02-04", "2016-02-04", "现行有效"),
    "私募投资基金管理人内部控制指引": ("中基协（2016年2月1日发布）", "中国证券投资基金业协会", "2016-02-01", "2016-02-01", "现行有效"),
    "私募投资基金登记备案办法": ("中基协（2023年2月24日发布）", "中国证券投资基金业协会", "2023-02-24", "2023-05-01", "现行有效"),
    "中国证券投资基金业协会自律管理和纪律处分措施实施办法": ("中基协", "中国证券投资基金业协会", "", "", "现行有效"),
    "中华人民共和国证券投资基金法": ("主席令第9号", "全国人民代表大会常务委员会", "2012-12-28", "2013-06-01", "现行有效"),
    "证券期货投资者适当性管理办法": ("证监会令第130号", "中国证券监督管理委员会", "2016-12-12", "2017-07-01", "现行有效"),
}


def split_multi(v):
    v = str(v or "")
    for sep in (";", "\n"):
        v = v.replace(sep, "；")
    return [x.strip() for x in v.split("；") if x.strip()]


# ----------------------------------------------------------------- 06 法规
def enumerate_06(root):
    """枚举 06 目录下全部法规文件（.docx + .pdf，两者都要）。"""
    six = os.path.join(root, LAW_DIR)
    out = []
    if not os.path.isdir(six):
        return out
    for dp, _, fs in os.walk(six):
        rel = os.path.relpath(dp, six)
        for f in sorted(fs):
            if f.lower().endswith(LAW_EXTS) and not f.startswith("~$"):
                stem = os.path.splitext(f)[0]
                out.append({
                    "category": "" if rel == "." else rel,
                    "name": stem,
                    "form": "PDF" if f.lower().endswith(".pdf") else "Word",
                    "filename": f,
                    "path": "" if rel == "." else rel,
                    "size": os.path.getsize(os.path.join(dp, f)),
                })
    return out


def _flat(s):
    """压平所有空白/换行。

    实测：决定书 PDF 正文常把法规名从中间换行（"私募投资基金监督\\n管理暂行办法"），
    不去换行会与 06 文件名匹配不上 → 同一部法规被误报成「缺失」，
    进而诱使操作员重复下载（2026-10-27 修复）。
    """
    return re.sub(r"\s+", "", s or "")


def norm_law(name):
    """法规名归一：去《》→ 剥括号后缀 → 压平空白 → 简称转全称（ALIAS）。

    2026-10-27 续接修复：先剥「（试行）/（2023年修订）」等括号后缀再查 ALIAS，
    否则「私募投资基金管理人登记和基金备案办法（试行）」等旧版写法查不到别名、
    会误报「06 缺失」并诱导重复下载。
    """
    n = (name or "").strip().strip("《》").strip()
    n = _PAREN_TAIL.sub("", n).strip()
    n = _flat(n)
    return ALIAS.get(n, n)


def match_law(cited, law_names):
    """把决定书点名的法规名匹配到 06 目录里的法规文件（空白不敏感）。

    先压平后精确相等，再尝试「互相包含」（容忍书名号/简称/换行差异）。
    """
    c = norm_law(cited)
    if not c:
        return None
    flat_map = {_flat(ln): ln for ln in law_names}
    if c in flat_map:
        return flat_map[c]
    for f, ln in flat_map.items():
        if f and (c in f or f in c):
            return ln
    return None


def case_texts(root):
    """产出每个案例的可检索文本：(行号, 核心处罚依据, PDF正文)。"""
    wb, ws = ledger_sort_free_open(root)
    C = LS.COL
    for r in range(2, ws.max_row + 1):
        basis = str(ws.cell(row=r, column=C["legal_basis"] + 1).value or "")
        paths = split_multi(ws.cell(row=r, column=C["paths"] + 1).value)
        files = split_multi(ws.cell(row=r, column=C["filenames"] + 1).value)
        body = ""
        for i, fn in enumerate(files):
            d = paths[i] if i < len(paths) else (paths[0] if paths else "")
            p = os.path.join(root, d, fn)
            if os.path.isfile(p):
                try:
                    d_ = fitz.open(p)
                    body += "\n".join(pg.get_text() for pg in d_)
                    d_.close()
                except Exception:
                    pass
        yield r, basis, body


def ledger_sort_free_open(root):
    """打开台账（不做任何排序/改动），返回 (wb, ws)。"""
    from openpyxl import load_workbook
    p = os.path.join(root, LS.LEDGER_NAME)
    wb = load_workbook(p)
    ws = wb[LS.SHEET_CASE] if LS.SHEET_CASE in wb.sheetnames else wb.active
    return wb, ws


def cmd_laws(root, out_json=None):
    laws = enumerate_06(root)
    law_names = [l["name"] for l in laws]
    # 引用数：每个案例对同一部法规只计 1 次（正文 ∪ 核心处罚依据列）
    refs = {n: 0 for n in law_names}
    cited_missing = {}
    for _r, basis, body in case_texts(root):
        blob = (basis or "") + "\n" + (body or "")
        hit = set()
        for m in CITE_RE.finditer(blob):
            ln = match_law(m.group(1), law_names)
            if ln:
                hit.add(ln)
            else:
                nm = norm_law(m.group(1))
                # 只把「像法规」且引用达阈值的记为缺失，避免把告知书/年报/合同当法规
                if nm and len(nm) >= 4 and looks_like_law(nm):
                    cited_missing[nm] = cited_missing.get(nm, 0) + 1
        for n in hit:
            refs[n] += 1
    # 组装清单（LAW_FIELDS 顺序）
    # 2026-10-27 用户定稿：不再产出 文号/发布机关/发布日期/施行日期/效力状态/留存形式
    # —— 这 6 项无法保证每部法规都取到，留着只会是半空列。
    rows = []
    for l in laws:
        rows.append({
            "category": l["category"], "name": l["name"],
            "refs": refs.get(l["name"], 0),
            "filename": l["filename"], "path": l["path"], "size": l["size"],
        })
    out_json = out_json or os.path.join(root, "_laws_generated.json")
    json.dump(rows, open(out_json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    # 回填台账法规清单 sheet
    ledger.cmd_law_sheet(Path(root), out_json)
    print("06 法规文件: %d 部（%s）" % (
        len(laws), "、".join(sorted({l['category'] for l in laws})) or "无分类"))
    miss = {k: v for k, v in cited_missing.items() if v >= 3}
    if miss:
        print("\n!! 点名引用 ≥3 次但 06 目录缺失（须 fetch_laws 补下载/补生成）：")
        for k, v in sorted(miss.items(), key=lambda x: -x[1]):
            print("   %-46s 引用 %d 次" % (k, v))
    else:
        print("点名引用 ≥3 次的法规均已留存（D3 应为 0）")
    return 0


# ----------------------------------------------------------------- 05 重点典型案例
def load_curation():
    """读 references/typical_cases.md 的遴选规则表。

    返回 [(子类, [关键词...], 限定违规类型或"")]。
    表结构：子类 | 匹配关键词 | 限定违规类型 | 说明
    """
    rules = []
    if not os.path.exists(CURATION):
        print("!! 未找到遴选规则文件:", CURATION)
        return rules
    header = None
    for ln in open(CURATION, encoding="utf-8"):
        ln = ln.strip()
        if not ln.startswith("|"):
            continue
        cells = [c.strip() for c in ln.strip("|").split("|")]
        if all(set(c) <= set("-: ") for c in cells if c):
            continue
        if header is None:
            header = cells
            continue
        if len(cells) < 2 or not cells[0]:
            continue
        sub = cells[0]
        kws = [k.strip() for k in re.split(r"[、,，]", cells[1]) if k.strip()]
        # 第 3 列是「限定违规类型」；"（不限）"表示不约束
        limit = cells[2] if len(cells) > 2 else ""
        if limit in ("（不限）", "(不限)", "不限", "—", "-"):
            limit = ""
        if sub and kws:
            rules.append((sub, kws, limit))
    return rules


def cmd_typical(root):
    """05 交叉归档（幂等：先把全表重置为「否」，再按规则标「是」）。

    匹配范围为台账 D 列主体 + F 列核心处罚事由，**不读 PDF 全文**——
    实测（2026-10-27）：全文匹配会把「决定书援引了含『关联交易』字样的法条」
    的普通案件全部误判成典型（误标 125/499 = 25%）。
    """
    rules = load_curation()
    if not rules:
        print("无遴选规则，跳过")
        return 1
    wb, ws = ledger_sort_free_open(root)
    C = LS.COL
    five = os.path.join(root, TYPICAL_DIR)
    n_file = n_row = 0
    total = 0
    for r in range(2, ws.max_row + 1):
        subj = str(ws.cell(row=r, column=C["subject"] + 1).value or "")
        cat = str(ws.cell(row=r, column=C["category"] + 1).value or "")
        grounds = str(ws.cell(row=r, column=C["grounds"] + 1).value or "")
        paths = split_multi(ws.cell(row=r, column=C["paths"] + 1).value)
        files = split_multi(ws.cell(row=r, column=C["filenames"] + 1).value)
        if not files:
            continue
        total += 1
        # 幂等：先重置
        ws.cell(row=r, column=C["typical"] + 1).value = "否"
        blob = subj + " " + grounds
        for sub, kws, limit in rules:
            if limit and limit != cat:
                continue                      # 限定分类不满足
            if not any(k in blob for k in kws):
                continue
            dst = os.path.join(five, sub)
            os.makedirs(dst, exist_ok=True)
            for i, fn in enumerate(files):
                d = paths[i] if i < len(paths) else (paths[0] if paths else "")
                src = os.path.join(root, d, fn)
                if os.path.isfile(src):
                    target = os.path.join(dst, fn)
                    if not os.path.exists(target):
                        shutil.copy2(src, target)
                        n_file += 1
            ws.cell(row=r, column=C["typical"] + 1).value = "是"
            n_row += 1
            print("  典型[%s] 行%d %s" % (sub, r, subj[:30]))
            break
    wb.save(os.path.join(root, LS.LEDGER_NAME))
    pct = (n_row * 100 // total) if total else 0
    print("05 交叉归档完成：典型案例 %d / %d 行（%d%%），复制文件 %d 份"
          % (n_row, total, pct, n_file))
    if pct > 10:
        print("!! 典型案例占比 %d%% 偏高——关键词可能过宽，应收紧到主体名或加限定分类" % pct)
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--laws", action="store_true")
    ap.add_argument("--typical", action="store_true")
    ap.add_argument("--laws-out", default="")
    a = ap.parse_args()
    if not (a.laws or a.typical):
        ap.error("至少指定 --laws 或 --typical")
    if a.laws:
        cmd_laws(a.root, a.laws_out or None)
    if a.typical:
        cmd_typical(a.root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
