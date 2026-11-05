#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""check_fields_source.py —— 台账列 ↔ 机构文书 回现校验（2026-11-02）

用户口径（2026-11-02 定）：
  台账除「处罚措施（机构+人员）」需要合并机构文书与人员文书外，其余各列**只从机构文书取**；
  且一个案例的机构文书**只有一份**。

  这条口径一旦被破坏，症状就是「同一段话出现两遍 / 依据列同一部法写两遍」——
  台账形态完全合法（非空、有编号、枚举合法），C1~C12 全部查不出来。
  唯一可靠的查法：把台账文字拿回机构文书里**逐句回现**。

判定：
  · 机构文书 = L 列里**不含「-」**的那份（人员文书命名是「姓名-机构名.pdf」）；
    有多份时优先不含「复核」的（复核决定只是维持原决定，事实与原文一致）。
  · 每列取若干「指纹」（去掉标点后的连续中文片段，长度 ≥12），
    指纹在机构文书中命中率 < 50% → 该列记为 mismatch（内容不来自机构文书）。
  · 机构文书无文字层且未 OCR → 记 no_text，不判错（避免把「读不到」当「抄错了」）。

用法：
  python scripts/check_fields_source.py --root <库根>                 # 只查有文字层的
  python scripts/check_fields_source.py --root <库根> --ocr           # 扫描件也 OCR（慢）
  python scripts/check_fields_source.py --root <库根> --rows 36,344   # 只看指定行
  python scripts/check_fields_source.py --root <库根> --csv out.csv
"""
import argparse
import csv
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

# 默认只查 F（事由）。H 列写法是「《法规简称》条款号」，原文是「违反《法规全称》第X条」——
# 简称与全称 6-gram 对不上，属**方法不适用**而非内容错（实测 11 条线索全是这类）。
# 需要时用 --all-cols 连 G/H/I 一起查，但结论要人工甄别。
COLS_F = [(6, "F核心处罚事由")]
COLS_ALL = [(6, "F核心处罚事由"), (7, "G核心证据材料"), (8, "H核心处罚依据"), (9, "I申辩理由")]
FP_MIN = 12          # 指纹最短长度
MATCH_MIN = 0.60     # 每条最长公共子串占本条长度的比例下限


def fingerprints(text, per_col=6):
    """从一格里取指纹：去标点后的连续中文片段，优先取长的、分布均匀的。"""
    t = re.sub(r"[（(][^（）()]{0,40}[）)]", "", str(text))
    chunks = re.findall(r"[\u4e00-\u9fa5A-Za-z0-9]{12,}", t)
    if not chunks:
        return []
    chunks = sorted(chunks, key=len, reverse=True)
    out, seen = [], set()
    for c in chunks:
        if c in seen:
            continue
        seen.add(c)
        out.append(c)
        if len(out) >= per_col:
            break
    return out


_PUNC = re.compile(r"[\s\u3000，。；：、,.;:（）()《》“”\"'‘’…—–\-|｜]")


def best_cover(seg, txt):
    """seg 的文字有多少能在 txt 里找到 —— 用 6-gram 命中率。

    为什么不用「最长公共子串」：一长句中间只要有**一处**标点/空格差异，LCS 就断在
    那里，覆盖率会虚低（行188 是一整段 100+ 字，LCS 只有 19%，实际只差几个逗号）。
    6-gram 逐窗命中对局部差异免疫，判定的是「这些字是不是从这份文书来的」，
    而不是「有没有一字不差地连续出现」。
    """
    s = _PUNC.sub("", str(seg))
    t = _PUNC.sub("", txt)
    if len(s) < 12:
        return 1.0
    k = 6
    grams = [s[i:i + k] for i in range(len(s) - k + 1)]
    if not grams:
        return 1.0
    hit = sum(1 for g in grams if g in t)
    return hit / len(grams)


def col_cover(value, txt):
    """一列的整体覆盖率 = 各条（按换行/分号切）覆盖率的均值。"""
    segs = [s for s in re.split(r"[\n；;]", str(value)) if len(s.strip()) >= 8]
    if not segs:
        return 1.0
    return sum(best_cover(s, txt) for s in segs) / len(segs)


def norm(s):
    return re.sub(r"[\s\u3000]", "", str(s))


def is_person_file(n):
    """人员文书命名 = 「姓名-机构名.pdf」。不能简单用「含 -」判断 ——
    机构文书本身可能带日期「浙江优策投资管理有限公司（2025-11-25）.pdf」（实测踩坑）。"""
    base = re.sub(r"\.pdf$", "", n, flags=re.I)
    if "-" not in base:
        return False
    head = base.split("-")[0]
    return len(head) <= 6 and not re.search(r"(公司|合伙|中心|企业|基金)", head)


def pick_inst_file(names):
    """从 L 列文件名里挑机构文书：非人员文书，且优先不含「复核」（复核只是维持原决定）。"""
    inst = [n for n in names if not is_person_file(n)]
    if not inst:
        return ""
    plain = [n for n in inst if "复核" not in n]
    return (plain or inst)[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--ocr", action="store_true")
    ap.add_argument("--rows", default="")
    ap.add_argument("--csv", default="")
    ap.add_argument("--all-cols", action="store_true", help="连 G/H/I 一起查（H 列简称问题会误报）")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]
    only = {int(x) for x in args.rows.split(",") if x.strip()} if args.rows else None
    COLS = COLS_ALL if args.all_cols else COLS_F

    from probe_grounds import pdf_text_join, text_layer_ok  # noqa: E402

    ok = no_text = person_only = 0
    bad = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        if only and r not in only:
            continue
        names = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        inst = pick_inst_file(names)
        d = str(ws.cell(r, 13).value or "").split("；")[0].strip()
        if not inst:
            # 人员文书先挂网 / 该案只罚了个人 → 按人员文书填**是合规的临时态**，
            # 但「临时态」不等于「不用查」：这 34 行以前直接 continue，等于没人对过文件
            # （用户 2026-11-02：每一格都要对过实际文件）。
            # 比对源＝本行**任一**人员文书里覆盖最高的那一份（不能只挑最大文件：
            # 行418 两人的事实各自不同，挑错人就把 100% 算成 51%，行418 实测踩过）。
            person_only += 1
            fval = str(ws.cell(r, 6).value or "")
            best = None
            for n in names:
                p0 = os.path.join(root, d, n)
                if not os.path.isfile(p0):
                    continue
                t0 = norm(pdf_text_join(p0))
                if len(t0) < 300:
                    if args.ocr:
                        try:
                            import ocrlib
                            t0 = norm(ocrlib.ocr_text(p0))
                        except Exception:
                            t0 = ""
                    if len(t0) < 300:
                        continue
                cv = col_cover(fval, t0)
                if best is None or cv > best[0]:
                    best = (cv, n)
            if best is None:
                no_text += 1
                continue
            if best[0] < MATCH_MIN:
                bad.append((r, "F核心处罚事由",
                            "指纹命中 覆盖 %.0f%%（本行人员文书最佳）" % (best[0] * 100), fval[:60]))
            else:
                ok += 1
            continue
        cols_here = COLS
        fp = os.path.join(root, d, inst)
        if not os.path.isfile(fp):
            bad.append((r, "*", "机构文书不存在：%s" % inst, ""))
            continue

        txt = norm(pdf_text_join(fp))
        if len(txt) < 300:
            if args.ocr:
                try:
                    import ocrlib
                    txt = norm(ocrlib.ocr_text(fp))
                except Exception as ex:
                    bad.append((r, "*", "OCR 失败：%s" % ex, inst))
                    continue
            if len(txt) < 300:
                no_text += 1
                continue

        row_bad = []
        for c, cn in cols_here:
            v = ws.cell(r, c).value
            if not isinstance(v, str) or len(v.strip()) < 12:
                continue
            # 交易所 H 列写的是「本所规则名」，交易所文书正文里不会出现《深交所股票上市规则》
            # 这种外部法规全称 → 不适用回现校验，跳过（改由 X31 判条款号齐全）。
            if c == 8 and "交易所" in str(ws.cell(r, 11).value or ""):
                continue
            cover = col_cover(v, txt)
            if cover < MATCH_MIN:
                row_bad.append((cn, "覆盖 %.0f%%" % (cover * 100), v[:60]))
        if row_bad:
            for cn, ratio, sample in row_bad:
                bad.append((r, cn, "指纹命中 %s" % ratio, sample))
        else:
            ok += 1

    print("\n对照完成：命中 %d 行（其中仅人员文书 %d 行按人员文书核 F，待机构文书置换）"
          "｜无文字层跳过 %d 行｜线索 %d 条" % (ok, person_only, no_text, len(bad)))
    for r, cn, why, sample in bad[:40]:
        print("  行%-4d %-12s %s  %s" % (r, cn, why, sample.replace("\n", " ")[:50]))
    if len(bad) > 40:
        print("  …其余 %d 条见 CSV" % (len(bad) - 40))

    if args.csv and bad:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["行号", "列", "说明", "台账取样"])
            w.writerows(bad)
        print("\nCSV：%s" % args.csv)
    return 0


if __name__ == "__main__":
    sys.exit(main())
