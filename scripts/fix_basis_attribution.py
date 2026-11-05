# -*- coding: utf-8 -*-
"""H 列「法条归错法规」定点修复（依据＝决定书原文，不靠推断）。

问题形态（2026-11-03 实测）：
  原文：违反了《关于加强私募投资基金监管的若干规定》第四条、《暂行办法》第二十三条第九项的规定
  台账：《关于加强私募投资基金监管的若干规定》第四条、第二十三条第九项
  —— 抽取 H 时把**第二部法规的名字吞掉了**，条号顺延挂在前一部法规名下。
  症状不只是「条号超条数」（《若干规定》只有 14 条却出现第二十三条），
  更危险的是**静默错**：若错挂到的法规恰好也有同号条文，任何形态检查都查不出来。

判据（三条同时满足才改，缺一即只报告）：
  1. 台账里的 (法规, 条号) 在本文书全文中**找不到**该法规带该条号的引用；
  2. 本文书里**确实有另一部法规**带着同一个条号出现；
  3. 这样的「另一部法规」**唯一**（多于一部即为歧义，不动）。

工具只做「归位」：条号、款、项一字不改，不删条目、不新增条目。
用法：
    python fix_basis_attribution.py --root <库根>                 # 干跑，打印全部差异
    python fix_basis_attribution.py --root <库根> --apply         # 落盘（自动备份 + 逐格校验）
    python fix_basis_attribution.py --root <库根> --rows 13,121   # 只查指定行
"""

import argparse
import csv
import os
import re
import shutil
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

LEDGER = "04_综合索引台账【总表】.xlsx"
CN = "一二三四五六七八九十百零〇"
# 》 与条号之间常夹文号括注（如「(证监会公告〔2020〕71号)」），必须允许跳过
# 否则《若干规定》第四条整条抓不到，候选集为空 → 本该归位的行被漏判（行13 实测）
# 另外还要容忍空格：PDF 两端对齐会把「》 第四条」排出一个空格，
# 不容忍就会误判成「文书里没有该条」→ 把对的条目挪走（行92 实测假阳性）
_BETWEEN = r"[\s\u3000]*(?:[（(][^）)]{0,40}[）)])?[\s\u3000]*"
# 条号串之间的分隔也要容忍空格：原文「第四十四条、 第四十五条」顿号后带一个空格，
# 不容忍就只抓到第四十四条、把第四十五条漏成「文书未见」（行229 实测）
_SEP = r"[\s\u3000]*[、，和及]?[\s\u3000]*"
LAW = re.compile(
    r"《([^》]{2,60})》%s((?:第[%s]+条(?:第[%s]+[款项])?%s)*)"
    % (_BETWEEN, CN, CN, _SEP)
)


def normalize(text):
    """把 PDF 抽出的文本归一成「一行」，否则条号会被硬换行切断。

    实测（行13 江苏中泓国信）：`《暂行办法》第二十\\n三条第九项` —— 不归一，
    「第二十三条」永远抓不到，归位候选为空。

    还要去掉**汉字之间的空格**：PDF 两端对齐会在字间补空格，实测出现
    「违 反了」「《私募投资基金 监督管理条例》」「第 五 十 九 条」，
    不去掉则「违反」关键词、法规名、条号全部匹配不到（行197 实测）。
    """
    t = str(text or "")
    t = t.replace("\u3000", "").replace("\xa0", "")
    t = re.sub(r"[ \t]*[\r\n]+[ \t]*", "", t)
    for _ in range(3):          # 空格可能连着几个，反复收几次才干净
        t2 = re.sub(r"(?<=[\u4e00-\u9fff》”）])\s+(?=[\u4e00-\u9fff《“（])", "", t)
        if t2 == t:
            break
        t = t2
    return t


_LAWISH = ("法", "条例", "规定", "办法", "规则", "指引", "细则", "准则", "通知", "决定", "意见")


_LAW_TAIL = ("办法", "规定", "规则", "法", "条例", "指引", "细则", "通知", "决定", "意见", "准则")


def _lawish(raw):
    """像不像一部法规的名字（用于 06 未收录、只能照抄原文的场合）。

    光看「含『规定』『办法』」不够：正则跨句误吃时会得到
    `上海证券交易所首次公开发行证券发行与三条、第六条、第七条等规定。鉴于…根据《实施细则`
    这种长串（含「等规定」），把它当新法规就会把对的挪给垃圾名（行73 实测）。
    所以加**结尾校验**：必须以法规后缀收尾，或以「第N号」「（试行）」收尾。
    """
    r = re.sub(r"\s+", "", str(raw or ""))
    if not (4 <= len(r) <= 40):
        return False
    if "。" in r or "；" in r or "，" in r:
        return False
    if r.endswith(_LAW_TAIL):
        return True
    return bool(re.search(r"第[0-9一二三四五六七八九十]+号$", r)
                or re.search(r"[（(](试行|修订|修正|暂行)[）)]$", r))


def _ocr_typo(idx, raw):
    """这个「未收录名」是不是已收录法规的 OCR 错名（只差一两个字）。

    实测行265：06 未收录名《若于规定》与已收录《若干规定》只差 1 个字，
    若把它当新法规，就会把台账里正确的《若干规定》第十二条改成《若于规定》第十二条
    —— 越修越错。名字近似即判为 OCR 错名，一律不当归位目标。
    """
    import difflib

    r = re.sub(r"\s+", "", str(raw or ""))
    for t in idx.titles:
        if t == r:
            continue
        if difflib.SequenceMatcher(None, r, t).ratio() >= 0.85:
            return t
    return ""


def doc_pairs(idx, text):
    """文书里出现过的 (法规名, 条号) 集合。

    06 未收录的法规**保留原文名**：归位依据是文书原文的位置，不是 06 库；
    名字照抄原文才不丢信息，06 缺失只影响「有没有正文可看」（批注会如实标注）。
    """
    pairs, seen = set(), []
    for m in LAW.finditer(normalize(text)):
        law, _ = idx.resolve(m.group(1), seen)
        if law and law not in seen:
            seen.append(law)
        if not law:
            law = re.sub(r"\s+", "", m.group(1))
            if not _lawish(law) or _ocr_typo(idx, law):
                continue        # 不像法规名 / 是已收录法规的错字版 → 不当候选
        for a in re.findall(r"第([%s]+)条" % CN, m.group(2)):
            pairs.add((law, "第%s条" % a))
    return pairs


def doc_near(idx, text):
    """每个「第X条」往回看最近的《…》落在哪部法 —— 用于消解同号歧义。

    与 doc_pairs 的正向串匹配互为独立证据：两者对得上才敢动。
    实测行190：原文「《内控指引》第十二条、《私募投资基金登记备案办法》第二十一条」，
    问「第二十一条属于谁」时有两部法都有第二十一条，只有位置能定案。
    """
    t = normalize(text)
    refs, seen = [], []
    for m in re.finditer(r"《([^》]{1,60})》", t):
        law, _ = idx.resolve(m.group(1), seen)
        if law and law not in seen:
            seen.append(law)
        refs.append((m.start(), law))
    out = set()
    for m in re.finditer(r"第([%s]+)条" % CN, t):
        prev = [x for x in refs if x[0] < m.start()]
        if prev and prev[-1][1]:
            out.add((prev[-1][1], "第%s条" % m.group(1)))
    return out


def doc_violation(idx, text):
    """哪些 (法规, 条号) 出现在「…违反了…」这类定性句里。

    同号多法时这是最后一道区分：处罚依据写在「违反了」句中，
    申辩/说明里引用的条文不在此列（行197 实测：《实施办法》第二十二条出现在申辩段）。
    """
    t = normalize(text)
    refs, seen = [], []
    for m in re.finditer(r"《([^》]{1,60})》", t):
        law, _ = idx.resolve(m.group(1), seen)
        if law and law not in seen:
            seen.append(law)
        refs.append((m.start(), law))
    out = set()
    for m in re.finditer(r"第([%s]+)条" % CN, t):
        prev = [x for x in refs if x[0] < m.start()]
        if not prev or not prev[-1][1]:
            continue
        if "违反" in t[max(0, prev[-1][0] - 60):m.end()]:
            out.add((prev[-1][1], "第%s条" % m.group(1)))
    return out


def read_doc(path, cache_dir=None, ocr=True):
    """读文书文本：先取 PDF 文字层，缺文字层才 OCR（走 _ocr_cache，二次运行秒回）。

    不用 lawtext.read_text 是因为它不带 OCR 缓存，508 份逐轮重 OCR 要 10 分钟以上；
    Z13/audit_files 用的是同一个 `_ocr_cache`，共用一份即可。
    """
    import fitz

    try:
        d = fitz.open(path)
        try:
            t = "".join(p.get_text() for p in d)
        finally:
            d.close()
    except Exception:
        return ""
    if len(re.sub(r"\s+", "", t)) >= 200:
        return t
    if not ocr:
        return t
    try:
        from ocrlib import ocr_text

        return ocr_text(path, cache_dir=cache_dir) or t
    except Exception:
        return t


def alias_map(idx, h):
    """本格里「法规全称 → 原文写法」，同义取最短（优先用已有的简称）。

    归位后若写成《私募投资基金监督管理暂行办法》全称，会与同格其他段落的
    《暂行办法》并存，读者以为是两部法 —— 所以沿用本格既有写法。
    """
    m = {}
    for it in idx.parse_h(h):          # 必须整格解析：简称要靠前文全称才认得出
        law, raw = it.get("law"), it.get("raw")
        if not law or not raw:
            continue
        cur = m.get(law)
        if cur is None or len(raw) < len(cur):
            m[law] = raw
    return m


def clause(a):
    s = a["no"]
    if a["kuan"]:
        s += "第%s款" % a["kuan"]
    if a["xiang"]:
        s += "第%s项" % a["xiang"]
    return s


def _seg_items(idx, h):
    """按「；」切段，并给出每段的解析结果——简称解析用**整格上下文**。

    直接对每段 parse_h 会丢上下文：《条例》《管理办法》这类简称在段内认不出，
    law 变成 None → rebuild 只能原样写回、归位失效（2026-11-03 实测 12 行自检失败）。
    """
    segs = [s.strip() for s in str(h).split("；") if s.strip()]
    items = idx.parse_h(h)                      # 整格解析 → 简称可回溯前文全称
    n_ref = len(re.findall(r"《[^》]{1,60}》", str(h)))
    if len(items) != n_ref:                     # 对不上就放弃重写，绝不猜
        return None
    out, p = [], 0
    for s in segs:
        k = len(re.findall(r"《[^》]{1,60}》", s))
        out.append((s, items[p:p + k]))
        p += k
    return out if p == len(items) else None


def _assigned(idx, text, law, no):
    """这篇 H 文本里，「law 名下的某一段」是否包含条号 no。

    必须先把 `《…》` 解析成规范名再比：重写后用的是**本格已有的简称**
    （如《私募基金监管办法》），拿全称去正则找永远找不到，
    自证就会把**正确的**改动判成失败（行408 实测：正确归位被拦下）。
    """
    seen = []
    for m in re.finditer(r"《([^》]{1,60})》([^；]*)", str(text)):
        L, _ = idx.resolve(m.group(1), seen)
        if L and L not in seen:
            seen.append(L)
        # 06 未收录的法规 resolve 给 None，此时按原文名直比（未收录目标就是照抄原文名的）
        got = L or re.sub(r"\s+", "", m.group(1))
        if got == law and no in m.group(2):
            return True
    return False


def rebuild(idx, h, moves, alias=None):
    """只拆「受影响的那一段」，其余段落一字不动。

    段内按目标法规把条款**连续分组**；源法规保留原写法（如《若干规定》），
    被归位的条款用本格已有的写法另起一段（本格没有该法规时才退回规范全称）。
    """
    alias = alias or {}
    pairs = _seg_items(idx, h)
    if pairs is None:
        return str(h)                # 段与条目对不齐 → 一个字都不动
    out = []                         # [(法规名 or "", [条款...])]
    for seg, its in pairs:
        if len(its) != 1 or not its[0]["law"] or not its[0]["arts"]:
            out.append(("", [seg]))   # 解析不到（未收录/无条款）→ 原样保留
            continue
        it = its[0]
        law, raw = it["law"], it["raw"]
        groups = []                  # [(目标法规, [条款...])]
        for a in it["arts"]:
            tgt = moves.get((law, a["no"]), law)
            if groups and groups[-1][0] == tgt:
                groups[-1][1].append(clause(a))
            else:
                groups.append((tgt, [clause(a)]))
        for tgt, cls in groups:
            name = raw if tgt == law else alias.get(tgt, tgt)
            if not name.startswith("《"):
                name = "《%s》" % name
            out.append((name, cls))
    return _merge(out)


def _merge(segs):
    """相邻同名法规的段合并——**只在条号严格递增时**合并。

    无脑合并会把「《条例》第三条第三款、第二十八条第一款」和「《条例》第三条第四款」
    拼成「第三条第三款、第二十八条第一款、第三条第四款」，条号倒退，
    正是 §0.5.16 ⑭ 说的 X28 形态（2026-11-03 实测行5）。
    """
    from lawtext import art_sort_key

    def head(cls):
        m = re.match(r"第([一二三四五六七八九十百零〇]+)条", cls)
        k = art_sort_key("第%s条" % m.group(1)) if m else None
        return k[0] if k else None

    def ascending(parts, nxt):
        a, b = head(parts[-1]), head(nxt)
        return a is not None and b is not None and b > a

    out = []
    for name, parts in segs:
        if (name and out and out[-1][0] == name and parts
                and ascending(out[-1][1], parts[0])):
            out[-1][1].extend(parts)
        else:
            out.append([name, list(parts)])
    return "；".join((n + "、".join(v)) if n else "".join(v) for n, v in out)


def scan(root, ledger=None, rows=None):
    """全库扫描 H 列的法规归属问题（**唯一实现**：修复工具与体检器 Z16 共用）。

    返回 dict(plans, report, idx, ws, ledger, work)：
      plans  —— [(行, 旧H, 新H, moves)]，已通过四条自证；
      report —— [(行, 类型, 说明, 文书, 备注)]，类型含 归位/歧义跳过/文书未见/自检失败/跳过。
    """
    import openpyxl

    import lawtext
    from check_fields_source import pick_inst_file

    root = os.path.abspath(root)
    ledger = ledger or os.path.join(root, LEDGER)
    reg = os.path.join(root, "06_以上各类所有案例的配套法规依据")
    work = os.path.join(os.path.dirname(root), "_案例库维护留痕")
    cache = os.path.join(work, "_law_index", "index.json")
    ocr_cache = os.path.join(work, "_ocr_cache")   # 与 Z13/audit_files 共用
    idx, _ = lawtext.load_or_build(reg, cache)

    ws = openpyxl.load_workbook(ledger).worksheets[0]
    only = set(rows) if rows else None

    plans, report = [], []
    for r in range(2, ws.max_row + 1):
        if only and r not in only:
            continue
        h = str(ws.cell(r, 8).value or "").strip()
        if not h:
            continue
        items = idx.parse_h(h)
        pairs = {(it["law"], a["no"]) for it in items if it["law"] for a in it["arts"]}
        if not pairs:
            continue
        d = str(ws.cell(r, 13).value or "").split("；")[0].strip()
        files = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        inst = pick_inst_file(files)
        note = ""
        if not inst:
            # 还没等到机构文书、库里只有这一份（多为人员决定书）时，
            # 用它把 H 的归属核一遍——H 本来就是从它抄来的（行131 实测）
            if len(files) == 1:
                inst, note = files[0], "仅一份文书"
            else:
                continue
        fp = os.path.join(root, d, inst)
        if not os.path.isfile(fp):
            continue
        txt = read_doc(fp, ocr_cache) or ""
        dpairs = doc_pairs(idx, txt)
        if not dpairs:
            report.append((r, "跳过", "文书无可解析引用", "", ""))
            continue
        dnears = doc_near(idx, txt)
        dviol = doc_violation(idx, txt)

        moves, ambiguous, unknown = {}, [], []
        for law, no in sorted(pairs):
            if (law, no) in dpairs:
                continue
            cands = sorted({L for (L, A) in dpairs if A == no and L != law})
            if len(cands) > 1:
                # 同号多法：先看原文「紧挨着这个条号」的那部，再看是否落在定性句里
                near = sorted({L for L in cands if (L, no) in dnears})
                if len(near) > 1:
                    near = sorted(L for L in near if (L, no) in dviol)
                cands = near if len(near) == 1 else cands
            if len(cands) == 1:
                moves[(law, no)] = cands[0]
            elif len(cands) > 1:
                ambiguous.append((law, no, cands))
            else:
                unknown.append((law, no))
        if moves:
            new_h = rebuild(idx, h, moves, alias_map(idx, h))
            # 自证四条：①条号总数不变 ②归位后挂到新法规名下 ③原错挂消失 ④不产生重复条目
            # （③④ 用文本判而不只靠 parse_h：目标法规 06 未收录时 parse_h 给的是 None）
            n_old = len(re.findall(r"第[%s]+条" % CN, h))
            n_new = len(re.findall(r"第[%s]+条" % CN, new_h))
            bad = [(L, A, L2) for (L, A), L2 in moves.items()
                   if _assigned(idx, new_h, L, A) or not _assigned(idx, new_h, L2, A)]
            # 行36 实测：原文那处是**说理段**，台账里该条已经存在，
            # 挪过去会变成两条一模一样 → 只增重复、不增信息，一律不动
            def _cnt(txt):
                c = {}
                for it in idx.parse_h(txt):
                    for a in it["arts"]:
                        k = (it["law"] or it["raw"], a["no"])
                        c[k] = c.get(k, 0) + 1
                return c
            c_old, c_new = _cnt(h), _cnt(new_h)
            # 只抓「本格内同一法规同一条号出现两次以上」——正常归位是 0→1，不算重复
            dup = [k for k, v in c_new.items() if v > 1 and v > c_old.get(k, 0)]
            if n_old != n_new or bad or dup:
                report.append((r, "自检失败", "条号 %d→%d；异常 %s；重复 %s；未纳入" %
                               (n_old, n_new, bad, dup), inst, note))
                continue
            plans.append((r, h, new_h, moves))
            for (L, A), L2 in sorted(moves.items()):
                report.append((r, "归位", "%s %s → %s" % (L, A, L2), inst, note))
        for law, no, cands in ambiguous:
            report.append((r, "歧义跳过", "%s %s 可归 %s" % (law, no, "／".join(cands)),
                           inst, note))
        for law, no in unknown:
            report.append((r, "文书未见", "%s %s" % (law, no), inst, note))

    return {"plans": plans, "report": report, "idx": idx,
            "ws": ws, "ledger": ledger, "work": work}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--rows", default="", help="只处理这些行，如 13,121")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--csv", default="")
    args = ap.parse_args()

    import openpyxl          # main 里落盘后还要重开台账复核，别只在 scan 内导入

    rows = [int(x) for x in args.rows.split(",") if x.strip()] if args.rows else None
    sc = scan(args.root, args.ledger or None, rows)
    plans, report, ws, ledger, work = (sc["plans"], sc["report"], sc["ws"],
                                       sc["ledger"], sc["work"])

    print("待改 %d 行｜差异明细 %d 条" % (len(plans), len(report)))
    for r, old, new, moves in plans:
        print("\n==== 行 %d" % r)
        print("  旧：", old)
        print("  新：", new)
    if not plans:
        print("（无需要归位的行）")

    if args.csv:
        os.makedirs(os.path.dirname(args.csv) or ".", exist_ok=True)
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["台账行", "类型", "说明", "文书", "备注"])
            w.writerows(report)
        print("\n明细：%s" % args.csv)

    if not args.apply or not plans:
        print("\n（干跑结束，未改动台账）")
        return 0

    bak = os.path.join(work, "台账备份_%s" % time.strftime("%Y%m%d"),
                       "%s_bak_%s.xlsx" % (os.path.splitext(LEDGER)[0], time.strftime("%H%M%S")))
    os.makedirs(os.path.dirname(bak), exist_ok=True)
    shutil.copy2(ledger, bak)
    print("\n已备份：%s" % bak)

    for r, old, new, moves in plans:
        assert len(new) > 10 and "《" in new, "行%d 重写结果异常：%s" % (r, new)
        ws.cell(r, 8).value = new
    wb = ws.parent
    wb.save(ledger)

    # 复核：条号一个不少
    ws2 = openpyxl.load_workbook(ledger).worksheets[0]
    bad = 0
    for r, old, new, moves in plans:
        n_old = len(re.findall(r"第[%s]+条" % CN, old))
        n_new = len(re.findall(r"第[%s]+条" % CN, new))
        if n_old != n_new:
            bad += 1
            print("  ✗ 行%d 条号数变化 %d → %d" % (r, n_old, n_new))
    print("落盘完成：%d 行｜条号数量校验 %s" % (len(plans), "通过" if not bad else "失败 %d 行" % bad))
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
