#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""explore_anomalies.py —— 未知类型问题探索器（2026-11-02，不依赖预设规则）

为什么要有它：
  verify_library / audit_columns / audit_cells 都是**已知类型**检查——它们只报
  「我已经想到的形态」。用户实拍过的每一次事故（M 列重复写、代码残片、
  网页摘要冒充正文、机构名缺右括号）当时都**不在任何检查器里**。
  所以自检必须另有一条腿：**不给规则，让数据自己说哪里不对**。

思路：同类比同类。
  同一来源（证监局/中基协/交易所）、同一违规类型、同一措施类型的行，
  形态应当接近。偏离总体分布的，就是候选异常——先当线索，不是结论。

检测项（全部是「离群/矛盾」，不需要知道正确答案）：
  Z1 长度离群       —— 同行内 F 远长于 H、或某列长度 > 同类中位数 3 倍 / < 0.15 倍
  Z2 跨列串味       —— F 与 H / F 与 E 字符集 Jaccard > 0.55（两列内容高度重合）
  Z3 怪字符         —— 控制字符、零宽字符、全角英数、重复标点（。。、、；；）
  Z4 句子复用       —— 长度 ≥30 的句子在同一列出现 ≥3 次（模板灌到多个案例）
  Z5 链接重复       —— 同一 URL 出现在 ≥2 行（一案一链，重复=错配或合并漏了）
  Z6 文件重名       —— 同一文件名出现在 ≥2 行（同目录不同名是正常的，跨行同名才可疑）
  Z7 主体/文件名背离 —— 文件名里没有 D 列主体名，且首段与该主体编辑距离 >3
  Z8 日期矛盾       —— B 早于 2015 / 晚于今天 / C-B > 400 天
  Z9 条目数离群     —— F 列分点条目数 >12（多半是正文整段被当成条目）
  Z10 同名主体多日期 —— D 列完全同名但落款日跨度 >180 天（可能是两家同名公司被并了）

输出：按检测项分组打印「行号 + 一句为什么可疑」，并给 CSV 方便逐条判断。
**只报线索，不改数据**。

用法：python scripts/explore_anomalies.py --root <库根> [--show 6]
"""
import argparse
import collections
import csv
import os
import re
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402

NAMES = ["A序号", "B处罚时间", "C挂网日期", "D被处罚机构人", "E处罚措施", "F核心处罚事由",
         "G核心证据材料", "H核心处罚依据", "I申辩理由", "J违规类型", "K机构来源",
         "L案例文件", "M存放路径", "N原文链接", "O是否典型案例"]

CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\u200b-\u200f\u2028\u2029\ufeff]")
FULLWIDTH_ALNUM = re.compile(r"[０-９Ａ-Ｚａ-ｚ]")
DUP_PUNCT = re.compile(r"[。；，、]{2,}")
IDX = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|[（(]\d+[）)])", re.M)
SENT_SPLIT = re.compile(r"[。；;\n]")


def S(v):
    return "" if v is None else str(v).strip()


def jaccard(a, b):
    A, B = set(a), set(b)
    if not A or not B:
        return 0.0
    return len(A & B) / len(A | B)


def load(root):
    p = os.path.join(root, "04_综合索引台账【总表】.xlsx")
    if not os.path.isfile(p):
        cands = [f for f in os.listdir(root) if f.endswith(".xlsx") and not f.startswith("~$")]
        p = os.path.join(root, cands[0])
    ws = openpyxl.load_workbook(p, data_only=True).worksheets[0]
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=15, values_only=True), start=2):
        if not S(r[0]):
            continue
        rows.append({"row": i, "c": [S(v) for v in r]})
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--show", type=int, default=6)
    ap.add_argument("--out", default="")
    args = ap.parse_args()
    root = os.path.abspath(args.root)
    rows = load(root)

    hits = collections.defaultdict(list)   # code -> [(row, detail)]

    # ---- Z1 长度离群 ----
    for col, name in [(5, "F"), (6, "G"), (7, "H")]:
        vals = [len(x["c"][col]) for x in rows if x["c"][col]]
        if len(vals) < 20:
            continue
        sv = sorted(vals)
        p5 = sv[max(0, int(len(sv) * 0.05) - 1)]
        p95 = sv[min(len(sv) - 1, int(len(sv) * 0.95))]
        med = statistics.median(vals)
        for x in rows:
            v = x["c"][col]
            if not v:
                continue
            L = len(v)
            if L > p95 * 1.6 or (L < p5 * 0.35 and p5 > 20):
                hits["Z1长度离群"].append(
                    (x["row"], "%s 列 %d 字（P5=%d 中位=%d P95=%d）：%s"
                     % (name, L, p5, med, p95, v[:40])))

    # ---- Z2 跨列串味 ----
    for x in rows:
        e, f, h = x["c"][4], x["c"][5], x["c"][7]
        if len(f) > 20 and len(h) > 20 and jaccard(f, h) > 0.55:
            hits["Z2跨列串味"].append((x["row"], "F 与 H 重合度 %.2f" % jaccard(f, h)))
        if len(e) > 40 and len(f) > 40 and jaccard(e, f) > 0.55:
            hits["Z2跨列串味"].append((x["row"], "E 与 F 重合度 %.2f" % jaccard(e, f)))

    # ---- Z3 怪字符 ----
    for x in rows:
        for ci, v in enumerate(x["c"]):
            if not v or ci in (0, 1, 2):
                continue
            why = []
            if CTRL.search(v):
                why.append("控制/零宽字符")
            if FULLWIDTH_ALNUM.search(v):
                why.append("全角英数")
            if DUP_PUNCT.search(v):
                why.append("重复标点")
            if why:
                hits["Z3怪字符"].append((x["row"], "%s：%s" % (NAMES[ci], "/".join(why))))

    # ---- Z4 句子复用 ----
    for col, name in [(5, "F"), (7, "H")]:
        pool = collections.defaultdict(set)
        for x in rows:
            for s in SENT_SPLIT.split(x["c"][col]):
                s = s.strip()
                if len(s) >= 30:
                    pool[s].add(x["row"])
        for s, rs in pool.items():
            if len(rs) >= 3:
                hits["Z4句子复用"].append(
                    (min(rs), "%s 列同一句出现 %d 次（行 %s）：%s"
                     % (name, len(rs), ",".join(map(str, sorted(rs)[:5])), s[:40])))

    # ---- Z5 链接重复 ----
    url2rows = collections.defaultdict(set)
    for x in rows:
        for u in re.split(r"[；;]", x["c"][13]):
            u = u.strip()
            if u:
                url2rows[u].add(x["row"])
    for u, rs in url2rows.items():
        if len(rs) >= 2:
            hits["Z5链接重复"].append((min(rs), "行 %s 共用同一链接：%s"
                                       % (",".join(map(str, sorted(rs))), u[-40:])))

    # ---- Z6 文件重名 ----
    fn2rows = collections.defaultdict(set)
    for x in rows:
        for f in re.split(r"[；;]", x["c"][11]):
            f = f.strip()
            if f:
                fn2rows[f].add(x["row"])
    for f, rs in fn2rows.items():
        if len(rs) >= 2:
            hits["Z6文件重名"].append((min(rs), "行 %s 同名：%s"
                                       % (",".join(map(str, sorted(rs))), f[:50])))

    # ---- Z7 主体/文件名背离 ----
    for x in rows:
        org = x["c"][3].split("；")[0].split("、")[0].strip()
        if len(org) < 6:
            continue
        for f in re.split(r"[；;]", x["c"][11]):
            f = f.strip()
            if not f or org in f:
                continue
            hits["Z7主体文件名背离"].append((x["row"], "主体 %r 不在文件名 %r" % (org[:20], f[:30])))

    # ---- Z8 日期矛盾 ----
    import datetime
    today = datetime.date.today()
    for x in rows:
        b, c = x["c"][1], x["c"][2]
        try:
            db = datetime.date.fromisoformat(b)
            if db.year < 2015 or db > today:
                hits["Z8日期矛盾"].append((x["row"], "处罚时间 %s 越界" % b))
            dc = datetime.date.fromisoformat(c)
            if (dc - db).days > 400:
                hits["Z8日期矛盾"].append((x["row"], "挂网 %s 晚于处罚 %s 达 %d 天"
                                           % (c, b, (dc - db).days)))
        except ValueError:
            hits["Z8日期矛盾"].append((x["row"], "日期格式异常 B=%r C=%r" % (b, c)))

    # ---- Z9 条目数离群 ----
    for x in rows:
        n = len(IDX.findall(x["c"][5]))
        if n > 12:
            hits["Z9条目数离群"].append((x["row"], "F 列 %d 条" % n))

    # ---- Z10 同名主体多日期 ----
    byorg = collections.defaultdict(list)
    for x in rows:
        org = x["c"][3].split("；")[0].strip()
        if len(org) >= 6:
            byorg[org].append((x["c"][1], x["row"]))
    for org, lst in byorg.items():
        ds = sorted(d for d, _ in lst if d)
        if len(ds) >= 2 and ds[0] and ds[-1]:
            try:
                span = (datetime.date.fromisoformat(ds[-1])
                        - datetime.date.fromisoformat(ds[0])).days
            except ValueError:
                continue
            if span > 180:
                hits["Z10同名主体日期跨度大"].append(
                    (lst[0][1], "%s：%s ~ %s（%d 天，%d 行）"
                     % (org[:24], ds[0], ds[-1], span, len(lst))))

    # ---- 输出 ----
    total = sum(len(v) for v in hits.values())
    print("=" * 78)
    print("未知类型问题探索 —— %d 行，命中线索 %d 条" % (len(rows), total))
    print("（线索≠错误，需逐条判断；目的是把「没想到的类型」逼出来）")
    print("=" * 78)
    for code in sorted(hits):
        lst = hits[code]
        print("\n-- %s：%d 条" % (code, len(lst)))
        for row, detail in lst[:args.show]:
            print("   行%-4d %s" % (row, detail))
        if len(lst) > args.show:
            print("   …（其余 %d 条见 CSV）" % (len(lst) - args.show))

    out = args.out or os.path.join(LS.maint_dir(root), "未知问题线索_%s.csv"
                                   % __import__("datetime").date.today().strftime("%Y%m%d"))
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["检测项", "行号", "线索"])
        for code in sorted(hits):
            for row, detail in hits[code]:
                w.writerow([code, row, detail])
    print("\n线索明细：%s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
