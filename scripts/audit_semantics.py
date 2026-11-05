#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""audit_semantics.py —— 台账**逐格内容**语义体检（2026-11-02）

audit_cells 查的是「形态」（空格/标点/残片/越界），本工具查的是「这一格像不像
它该有的内容」。两者互补：形态全过、语义全错的格子，历史上出现过好几次
（F 列装的是 H 列的法条、I 列装的是措施、G 列装的是事由缩写）。

S 系列判据（逐列）：
  S1  序号不连续/重复/断号
  S2  B/C 非法日期，或 B 晚于 C
  S3  D 列机构名不以机构后缀收尾；D 列用「、」把两个机构分开（应只一个机构）
  S4  E 列没有「管理人：/高管（…）：」标签，或整列没有措施动词
  S5  F 列没有任何违规词（未/违/不/擅自/超/挪用/承诺/混同…）→ 可能不是事由
  S6  G 列没有证据类词（证据/材料/笔录/报告/说明/记录/底稿/资料）
  S7  H 列没有《》或没有条款号
  S8  I 列没有申辩类词（申辩/陈述/异议/复核/听证/未提出）
  S9  J/K/O 枚举越界
  S10 L 列文件名不含 D 列机构名
  S11 M 列分类末级与 J 列不匹配
  S12 N 列链接数与 L 列文件数不等（且 L 非空）
  S13 格内出现占位/脏值（None/nan/待补/待核/不明确/暂无/无/／）
  S14 E 列措施动词缺「人」（只有机构措施没有高管措施，但 D 列有人）

用法：
  python scripts/audit_semantics.py --root <库根> [--csv out.csv] [--col F]
"""
import argparse
import csv
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

# 词表要按「实际文书用词」收：证监局/协会描述违规不一定用「未/违/不」，
# 「代持」「低于」「偏离」「诱导」「保本」「出借」都是违规描述（第一版漏了，误报 9 条）
VIOL_W = re.compile(r"未|违|不|擅自|超|挪用|侵占|承诺|混同|让渡|委托|从事|缺失|缺位|"
                    r"不当|虚假|误导|隐瞒|利用|将|向|通过|使用|减|少|迟|漏|虚|"
                    r"代持|低于|高于|偏离|诱导|保本|出借|代销|优先|拆分|滚动|垫付|"
                    r"混同|兼任|挂名|借用|违规|接受|安排|"
                    # 2026-11-02 补：「存在…的问题」「存在…的情形」是证监局警示函的标准句式，
                    # 整句可以一个违规动词都没有（行138 实测），不补会误报
                    r"问题|情形|缺陷|瑕疵")
EVID_W = re.compile(r"证据|材料|笔录|报告|说明|记录|底稿|资料|合同|流水|截图|清单|台账|决议")
DEF_W = re.compile(r"申辩|陈述|异议|复核|听证|未提出|辩称|答辩")
MEAS_W = re.compile(r"警示函|责令改正|罚款|警告|公开谴责|暂停|撤销|注销|加入黑名单|"
                    r"监管措施|纪律处分|通报批评|书面警示|监管警示|责令|没收|市场禁入|"
                    r"取消|限制|约谈|出具")
# 「…合伙企业（有限合伙）」「…中心（有限合伙）」也是合法机构名（第一版不收「）」误报 28 条）
INST_TAIL = re.compile(r"(公司|合伙|中心|企业|厂|社|所|院|银行|集团|基金|协会|商会|工作室|）|\))）?$")
ENUM_J = None
ENUM_K = {"中基协", "证监局", "交易所"}
ENUM_O = {"是", "否"}
BAD_VAL = re.compile(r"^(?:None|nan|NaN|NULL|null|待补|待核|不明确|暂无|不详|未知|/|／|—|-)$")
# 注意：「无」不在脏值里 —— I 列「无」表示「未提出申辩」，是合法值（第一版把它算脏值，误报 497 条）
DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--csv", default="")
    ap.add_argument("--col", default="", help="只看某列（A~O）")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]
    hdr = {c: str(ws.cell(1, c).value or "") for c in range(1, 16)}

    # 违规类型枚举：从 01/02/03 目录名反推（唯一事实源是磁盘结构）
    cats = set()
    for d in os.listdir(root):
        if d[:2] not in ("01", "02", "03"):
            continue
        for sub in os.listdir(os.path.join(root, d)):
            if os.path.isdir(os.path.join(root, d, sub)):
                cats.add(sub)          # J 列存的就是目录全名（含「02_」前缀），别去前缀

    out = []
    seq = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        g = lambda c: str(ws.cell(r, c).value or "").strip()
        a, b, c_, d, e, f, gg, h, i = (g(1), g(2), g(3), g(4), g(5), g(6), g(7), g(8), g(9))
        j, k, l, m, n, o = g(10), g(11), g(12), g(13), g(14), g(15)
        seq.append(a)
        add = lambda col, why, val: out.append((r, col, why, val[:60]))

        if not DATE.match(b) or not DATE.match(c_):
            add("B/C", "S2非法日期", b + " / " + c_)
        elif b > c_:
            add("B/C", "S2处罚晚于挂网", b + " > " + c_)

        mgr = d.split("；")[0].strip()
        # 2026-11-02：**个人独立警示函**（文书只针对个人、正文未载机构名）的 D 列就是纯人名，
        # 这是合法的（行126 魏锋 / 行146 张守军 实测）。判据：D 只有一个主体 + 是短人名 +
        # E 列以「高管（该人）」开头 → 不报。不放宽的话，每次跑都固定误报 2 条。
        person_only = (
            "；" not in d and len(mgr) <= 4 and "《" not in mgr
            and not re.search(r"(公司|合伙|中心|企业|基金|集团|银行)", mgr)
            and e.startswith("高管（%s）" % mgr)
        )
        if mgr and not INST_TAIL.search(mgr) and not person_only:
            add("D", "S3机构名无机构后缀", mgr)
        if "、" in d.split("；")[0]:
            add("D", "S3机构名内出现顿号", d.split("；")[0])

        if not re.search(r"管理人：|高管（", e):
            add("E", "S4无机构/人员标签", e)
        if not MEAS_W.search(e):
            add("E", "S4无措施动词", e)
        # 2026-11-02：交易所/个别证监局把机构和个人写在**同一句措施**里
        # （「对你公司、李厚文予以书面警示」，行206 实测）。此时 E 列没有「高管（姓名）」
        # 段落，但人已经被点名处罚 → 不报。判据：D 里的人名在 E 列出现过即算已覆盖。
        names = [x for x in re.split(r"[；、]", d) if x and x.strip() != mgr]
        unmentioned = [n for n in names if n.strip() and n.strip() not in e]
        if names and unmentioned and "高管（" not in e:
            add("E", "S14有人无高管措施", d + " || " + e)

        if len(f) >= 10 and not VIOL_W.search(f):
            add("F", "S5无违规词", f)
        if len(gg) >= 4 and not EVID_W.search(gg):
            add("G", "S6无证据类词", gg)
        if len(h) >= 4 and "《" not in h:
            add("H", "S7无书名号", h)
        if len(h) >= 4 and not re.search(r"第[一二三四五六七八九十百零\d]+", h):
            add("H", "S7无条款号", h)
        # I 列存的是**申辩内容本身**，绝大多数不含「申辩」二字（引导句在抽取时已剥掉）
        # → 不做「必须含申辩词」的断言，只判「极短且不是『无』」这种明显缺内容的情形。
        if i and i not in ("无", "未提出") and len(i) < 8:
            add("I", "S8申辩过短", i)

        if j and j not in cats:
            add("J", "S9违规类型不在目录枚举", j)
        if k and k not in ENUM_K:
            add("K", "S9来源越界", k)
        if o and o not in ENUM_O:
            add("O", "S9典型列越界", o)

        if l:
            for nm in [x for x in re.split(r"[；;]", l) if x.strip()]:
                if mgr and mgr not in nm:
                    add("L", "S10文件名不含机构名", nm)
                    break
        if m and j:
            if not m.split("/")[-1].strip().endswith(j.strip()):
                add("M", "S11路径末级≠违规类型", m.split("/")[-1] + " vs " + j)

        n_file = len([x for x in re.split(r"[；;]", l) if x.strip()])
        n_url = len([x for x in re.split(r"[；;]", n) if x.strip()])
        if n_file and n_url and n_file != n_url:
            add("N", "S12链接数≠文件数", "%d vs %d" % (n_url, n_file))

        for col, name, v in ((2, "B", b), (3, "C", c_), (4, "D", d), (5, "E", e), (6, "F", f),
                             (7, "G", gg), (8, "H", h), (9, "I", i), (10, "J", j),
                             (11, "K", k), (13, "M", m), (15, "O", o)):
            if BAD_VAL.match(v):
                add(name, "S13占位/脏值", v)

    # S1 序号
    try:
        want = list(range(1, len(seq) + 1))
        got = [int(x) for x in seq if x.isdigit()]
        if got != want:
            bad = [k for k, v in enumerate(got, 1) if k != v][:5]
            out.append((0, "A", "S1序号不连续", "第 %s 个起错位" % bad))
    except Exception:
        out.append((0, "A", "S1序号非数字", ""))

    from collections import Counter
    cc = Counter(x[1] for x in out)
    print("\n逐格语义体检：%d 行 / %d 格" % (len(seq), len(seq) * 15))
    if not out:
        print("  未发现语义异常")
    for k2 in sorted(cc):
        print("  列 %-3s %d 条" % (k2, cc[k2]))
    print()
    want_col = args.col.strip().upper()
    shown = 0
    for r, col, why, val in out:
        if want_col and col != want_col:
            continue
        print("  行%-4s %-4s %-18s %s" % (r or "-", col, why, val))
        shown += 1
        if shown >= 45:
            print("  …")
            break

    if args.csv and out:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["行号", "列", "类型", "内容"])
            w.writerows(out)
        print("\nCSV：%s" % args.csv)
    return 0


if __name__ == "__main__":
    sys.exit(main())
