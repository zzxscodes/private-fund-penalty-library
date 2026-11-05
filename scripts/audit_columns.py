#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""audit_columns.py —— 台账「逐行逐列」检查器（2026-11-02）

与 verify_library.py 的分工：
  verify_library  面向「交付验收」——PDF↔台账双向、OCR 落款回现、平台对账、法规引用，
                  重、慢、需要 PDF 与网络；
  audit_columns   面向「人工检验前的自查」——**只看台账 15 列本身**，不打开任何 PDF，
                  秒级跑完，逐列给出可读的问题清单。两者互补，不重复。

15 列逐项断言（列名与下标全部取自 ledger_style.COL，禁止各写一份）：

  A 序号       连续 1..N、无重复、是整数
  B 处罚时间    YYYY-MM-DD、真实日期、在窗口内（给了 --start/--end 时）
  C 挂网发布日期 YYYY-MM-DD、真实日期、不早于 B
  D 被处罚机构/人 非空、机构段合法（有组织形式后缀、无「曾登记为/实际负责」角色前缀）、
                人员段 2~4 字（复姓/少数民族名放宽）、无重复、无碎片词
  E 处罚措施    非空、剥掉「管理人：/高管（X）：」标签后不是占位词（不明确/待补…）、
                每个当事人要么有「高管（X）：」段要么姓名出现在措施正文里
  F 核心处罚事由 非空、≥15 字
  G 核心证据材料 非空
  H 核心处罚依据 非空、不含「正文未载明」
  I 申辩理由    非空
  J 违规类型    非空、在该来源的枚举内（交易所=四分法）
  K 机构来源    非空、∈{中基协, 证监局, 交易所}
  L 案例文件    非空、全部 .pdf、条数 ≥1
  M 存放路径    恰好 1 条（2026-11-02 口径：一个案例的文书必同目录）、目录在磁盘上存在、
                与 J 的分类目录一致
  N 原文链接    非空、http(s)、条数 ∈ {1, 文件数}
  O 是否典型案例 值 ∈ {是, 否}；为「是」时 05 目录下必须有同名副本

退出码：0 = 无问题（观察项不计）；1 = 有问题。

用法：
  python scripts/audit_columns.py --root <库根> [--start 2025-10-31 --end 2026-10-30] [--out r.txt]
"""
import argparse
import datetime as dt
import os
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402
from classify_filter import AMAC_CATS, CSRC_CATS, EXCH_CATS  # noqa: E402
from subject_clean import is_noise_token, strip_role_prefix  # noqa: E402

C = LS.COL
ORG_TAIL = ("公司", "合伙企业", "企业", "中心", "基金", "协会", "银行", "证券", "信托",
            "集团", "事务所", "研究院", "大学", "学校", "工作室", "公社", "医院", "酒店",
            "厂", "店", "社", "部", "会", "院", "所", "行")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
PLACEHOLDER = frozenset({"不明确", "不详", "未知", "未明确", "待补", "待核", "待定",
                         "无", "暂无", "空", "—", "-", "N/A"})
LABEL_RE = re.compile(r"^(?:管理人|机构|人员|高管)（?[^）]{0,40}）?[：:]\s*")
# 原文页码标记：— 2 — / – 3 – / 第 2 页
# 只认长破折号（— –），**不认半角连字符**——否则「2026-08-15」里的 "-08-" 会被误判（实测 9 条假报）
PAGE_MARK = re.compile(r"[—–]\s*\d{1,3}\s*[—–]|第\s*\d{1,3}\s*页")


def split_multi(v):
    return [p.strip() for p in re.split(r"[；;]", str(v or "")) if p.strip()]


def parse_date(s):
    s = (s or "").strip()[:10]
    if not DATE_RE.match(s):
        return None
    try:
        return dt.date.fromisoformat(s)
    except ValueError:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--start", default="")
    ap.add_argument("--end", default="")
    ap.add_argument("--out", default="")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    ledger = next(os.path.join(root, f) for f in os.listdir(root)
                  if f.startswith("04_") and f.endswith(".xlsx") and not f.startswith("~$"))
    ws = openpyxl.load_workbook(ledger).active
    win = (parse_date(a.start), parse_date(a.end)) if a.start else (None, None)

    # 05 副本索引（O 列用）
    copies = set()
    for d in os.listdir(root):
        if d.startswith("05_"):
            for dp, _dn, fns in os.walk(os.path.join(root, d)):
                copies.update(fns)

    P = defaultdict(list)   # 问题
    OBS = defaultdict(list)  # 观察项
    rows = 0

    for r in range(2, ws.max_row + 1):
        g = lambda k: str(ws.cell(r, C[k] + 1).value or "").strip()
        subj, meas = g("subject"), g("measures")
        if not subj:
            # 空行：若整行全空则跳过，否则报
            if any(str(ws.cell(r, c + 1).value or "").strip() for c in range(15)):
                P["A 序号 空行残留"].append((r, "有内容但 D 列为空"))
            continue
        rows += 1

        # ---------- A 序号 ----------
        seq = g("seq")
        if not seq.isdigit() or int(seq) != rows:
            P["A 序号 不连续/非数字"].append((r, "序号=%r 期望 %d" % (seq, rows)))

        # ---------- B/C 日期 ----------
        b, c_ = parse_date(g("decided")), parse_date(g("publish"))
        if not b:
            P["B 处罚时间 格式/无效"].append((r, g("decided")))
        elif win[0] and win[1] and not (win[0] <= b <= win[1]):
            # 与 verify 的 A2 口径一致（2026-10-29 裁定）：台账处罚时间取的是**决定书落款**，
            # 挂网可能晚到窗口内（刊发普遍滞后 9~23 天）。只要挂网日在窗口内就是当窗口的案子，
            # 「仅落款出窗」属观察项，绝不进 problems —— 否则边界 3 行会被反复误报为需处理。
            if c_ and win[0] <= c_ <= win[1]:
                OBS["B 仅落款出窗、挂网在窗内（边界案，保留）"].append((r, "%s / 挂网 %s" % (b, c_)))
            else:
                P["B 处罚时间 窗口外"].append((r, g("decided")))
        if not c_:
            P["C 挂网日期 格式/无效"].append((r, g("publish")))
        elif b and c_ < b:
            P["C 挂网日期 早于处罚时间"].append((r, "%s < %s" % (c_, b)))
        # 原文页码标记（「— 2 —」）混进正文列：OCR 把页脚当正文了，读起来像乱码
        for lbl, key in (("E 措施", "measures"), ("F 事由", "grounds"), ("H 依据", "legal_basis")):
            if PAGE_MARK.search(g(key)):
                OBS["%s 含原文页码标记（— N —）" % lbl].append(
                    (r, PAGE_MARK.search(g(key)).group(0)))

        # ---------- D 主体 ----------
        segs = split_multi(subj)
        org = segs[0]
        if strip_role_prefix(org) != org:
            P["D 主体 机构名带角色前缀"].append((r, org))
        if not any(t in org for t in ORG_TAIL) and len(org) > 6:
            OBS["D 主体 机构段无组织形式后缀（个人案常见）"].append((r, org))
        persons = []
        for s in segs[1:]:
            persons += [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]
        if len(set(persons)) != len(persons):
            P["D 主体 人员重复"].append((r, "、".join(persons)))
        for p in persons:
            if is_noise_token(p):
                P["D 主体 人员含碎片词"].append((r, p))
            elif not re.fullmatch(r"[\u4e00-\u9fa5·]{2,8}", p):
                P["D 主体 人员名不合法"].append((r, p))

        # ---------- E 措施 ----------
        if not meas:
            P["E 措施 为空"].append((r, org[:22]))
        else:
            core = re.sub(r"^[；;，,\s]+", "", LABEL_RE.sub("", meas)).strip()
            if core in PLACEHOLDER:
                P["E 措施 占位值"].append((r, "%s ｜%s" % (org[:18], core)))
            bullets = set(re.findall(r"高管（([^）]+)）", meas))
            uncovered = [p for p in persons if p not in bullets]
            if uncovered:
                invisible = [p for p in uncovered if p not in meas]
                if invisible:
                    OBS["E 措施 当事人无专属段且姓名未现于措施（联名决定书常见）"].append(
                        (r, "、".join(invisible)))

        # ---------- F/G/H/I 内容列 ----------
        if len(g("grounds")) < 15:
            P["F 事由 过短/为空"].append((r, len(g("grounds"))))
        # F 末尾残留「根据《…》第…条」＝处罚依据句，属 H 列，不是事由。
        # 2026-11-02 实测 27 行（决定段尾巴被整句抄进 F）。句中出现的（行247「据此，根据…构成…」）
        # 是事实叙述，不判；只判「最后一句就是纯引用」这一形态。
        _tail = re.split(r"[。；\n]", str(g("grounds") or "").rstrip("。；; "))[-1]
        if re.match(r"^[，,、]?\s*根据《", _tail.strip()) and re.search(
                r"第[一二三四五六七八九十百零〇\d]+条", _tail) and not re.search(
                r"约定|合同|据此|构成|不符合|违反|存在", _tail):
            P["F 事由 尾部残留依据句"].append((r, _tail.strip()[:40]))
        if not g("evidence"):
            P["G 证据 为空"].append((r, org[:22]))
        legal = g("legal_basis")
        if not legal:
            P["H 依据 为空"].append((r, org[:22]))
        elif "未载明" in legal or legal in PLACEHOLDER:
            P["H 依据 正文未载明"].append((r, legal[:24]))
        if not g("defense"):
            P["I 申辩 为空"].append((r, org[:22]))

        # ---------- J/K 分类与来源 ----------
        cat, src = g("category"), g("source")
        if "交易所" in src or src == "交易所":
            valid = EXCH_CATS
        elif src == "中基协":
            valid = AMAC_CATS
        elif "证监局" in src:
            valid = CSRC_CATS
        else:
            valid = []
            P["K 来源 不在枚举"].append((r, src or "(空)"))
        if not cat:
            P["J 违规类型 为空"].append((r, org[:22]))
        elif valid and cat not in valid:
            P["J 违规类型 不在该来源枚举"].append((r, "%s ｜%s" % (src, cat)))

        # ---------- L 文件 ----------
        files = split_multi(ws.cell(r, C["filenames"] + 1).value)
        if not files:
            P["L 案例文件 为空"].append((r, org[:22]))
        bad_ext = [f for f in files if not f.lower().endswith(".pdf")]
        if bad_ext:
            P["L 案例文件 非PDF"].append((r, "、".join(bad_ext)[:40]))

        # ---------- M 路径 ----------
        paths = split_multi(ws.cell(r, C["paths"] + 1).value)
        if len(paths) != 1:
            P["M 存放路径 条数≠1"].append((r, "%d 条" % len(paths)))
        else:
            p0 = paths[0]
            full = p0 if os.path.isabs(p0) else os.path.join(root, p0)
            if not os.path.isdir(full):
                P["M 存放路径 目录不存在"].append((r, p0[-40:]))
            elif cat and not p0.replace("\\", "/").endswith(cat):
                P["M 存放路径 与违规类型目录不一致"].append((r, "%s ≠ %s" % (p0[-24:], cat)))

        # ---------- N 链接 ----------
        urls = split_multi(ws.cell(r, C["urls"] + 1).value)
        if not urls:
            P["N 原文链接 为空"].append((r, org[:22]))
        else:
            bad = [u for u in urls if not re.match(r"^https?://", u)]
            if bad:
                P["N 原文链接 非http"].append((r, bad[0][:40]))
            if len(urls) not in (1, len(files)):
                P["N 原文链接 条数与文件数不匹配"].append((r, "链接%d 文件%d" % (len(urls), len(files))))

        # ---------- O 典型 ----------
        typ = g("typical")
        if typ not in ("是", "否"):
            P["O 典型案例 值非法"].append((r, typ or "(空)"))
        elif typ == "是":
            if not any(f in copies for f in files):
                P["O 典型案例 05 无副本"].append((r, org[:22]))

    # ---------- 输出 ----------
    out = ["=" * 74,
           "台账逐行逐列检查   共 %d 行   问题 %d 类 / 观察 %d 类" % (rows, len(P), len(OBS)),
           "=" * 74]
    for k in sorted(P):
        out.append("\n【%s】%d 条" % (k, len(P[k])))
        for x in P[k][:10]:
            out.append("   行%-4d %s" % (x[0], x[1]))
        if len(P[k]) > 10:
            out.append("   …（其余 %d 条）" % (len(P[k]) - 10))
    if OBS:
        out.append("\n" + "-" * 74 + "\n观察项（不进结论，供人工判断）")
        for k in sorted(OBS):
            out.append("\n〔%s〕%d 条" % (k, len(OBS[k])))
            for x in OBS[k][:8]:
                out.append("   行%-4d %s" % (x[0], x[1]))
            if len(OBS[k]) > 8:
                out.append("   …（其余 %d 条）" % (len(OBS[k]) - 8))
    out.append("\n" + "=" * 74)
    out.append("结论：%s" % ("全部列通过" if not P else "存在 %d 类问题需处理" % len(P)))
    txt = "\n".join(out)
    print(txt)
    if a.out:
        open(a.out, "w", encoding="utf-8").write(txt)
        print("\n报告已写：%s" % a.out)
    sys.exit(1 if P else 0)


if __name__ == "__main__":
    main()
