# -*- coding: utf-8 -*-
"""Z14 事由条目漏抽探测（F 列「只抄了一半」）。

为什么需要这条：Z11（回现覆盖率）只能证明「这段字确实来自这份文书」，
**证明不了「这段字把文书说全了」**。行330 就是典型：
    F = 一、广州市原象私募基金管理有限公司为其他企业进行项目推荐以及提供咨询服务。
    原文 = 一是你公司为其他企业进行项目推荐以及提供咨询服务,违反了《若干规定》第四条…
          二是你公司未及时更新从业人员信息，违反了《私募办法》第二十五条第一款的规定。
F 每个字都在原文里（回现率 100%），但**第二条整个丢了**，而且第一条在「，违反了」处被切断。
H 列却抽到了「第二十五条第一款」—— 说明抽取器看得到第二条，只是没写进 F 列。

判据（只在「F 文本能在原文里定位到」时生效，否则交给 Z11）：
    ① 截断：F 末字不是「。」（说明半句被削）
    ② 漏条：F 在原文中的**结束位置之后**、下一个「根据《/我局决定」之前，
       还出现「二是 / 二、 / （二） / 2.」这类**同级条目起始标记**

    「此外，」是唯一需要二次判定的标记（行301 实测）：它既可能引出第四条违规，
    也可能只是一句「此外，我局还关注到…」的提醒。判据是**后面有没有违反定性**，
    没有就进观察项（attention），不进 problems —— 见 §14.6「判据太松也要修」。

只报候选，裁定仍走 §14.4（线索≠错误）。

用法：
    python check_grounds_complete.py --root <库根>
    python check_grounds_complete.py --root <库根> --rows 330
    python check_grounds_complete.py --root <库根> --csv out.csv
"""
import argparse
import csv
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl

from check_fields_source import is_person_file, pick_inst_file, norm

LEDGER = "04_综合索引台账【总表】.xlsx"

# 同級條目起始（能出现在「F 结束之后」的那种）
NEXT_ITEM = re.compile(r"(?:^|[，,；;。])\s*(二是|二、|（二）|\(二\)|2[.、．]|此外[，,])")
# 事实段结束 / 决定段开始的标志：到这里就不算「还漏条目」
STOP = re.compile(r"根据《|依据《|我局决定|本局决定|决定对|综上|现决定|以上事实|上述事实")

# 「此外」是歧义标记：既可能引出第四条违规（该抄），也可能只是一句提醒
# （「此外，我局还关注到…」），后者没有违反定性、不是违规条目（行301 实测）。
# 因此「此外」必须再看后面有没有违反定性，没有就归观察项、不进 problems。
VIOL_TAIL = re.compile(r"违反了|违反《|不符合|未恪守|未依照|构成违规|违规行为|违法违规")


def read_text(path):
    import fitz

    d = fitz.open(path)
    try:
        return re.sub(r"\s+", "", "".join(p.get_text() for p in d))
    finally:
        d.close()


_STRIP = re.compile(r"[\s\u3000，。；、,.;:：（）()《》〈〉“”\"'‘’]")


def _clean_with_map(s):
    """去掉标点并保留「清洗后下标 → 原文下标」的映射。

    为什么不能直接在原文里找 F 的尾部：台账把原文的「,违反了…」截成「。」收尾，
    标点一变窗口就锚不上（行330 第一版就是这么漏掉的）。所以两边都去标点再锚，
    锚到之后用映射换回原文位置，才能看清「F 结束之后原文还剩什么」。
    """
    buf, pos = [], []
    for i, ch in enumerate(s):
        if not _STRIP.match(ch):
            buf.append(ch)
            pos.append(i)
    return "".join(buf), pos


def locate_end(f_norm, doc_norm):
    """在文书里定位 F 的「结束位置」（原文下标）。

    不能直接 doc.find(F)：台账做过归一化（「你公司」→机构全称、OCR 折行拼接、
    「一是」→「一、」），F **不是逐字原文**，find 必然失败（行330/138 实测 2/2 失败）。
    改用「从尾部往前滑窗」：取 F 最后 14/12/10/8 字，找能在原文命中的最长窗口，
    以该窗口的右端点当作 F 的结束位置。尾部是逐字抄的，所以锚得住。
    """
    fc, _ = _clean_with_map(f_norm)
    dc, dmap = _clean_with_map(doc_norm)
    n = len(fc)
    for w in (14, 12, 10, 8):
        if n < w:
            continue
        p = dc.rfind(fc[-w:])
        if p >= 0:
            return dmap[p + w - 1] + 1
    return -1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--rows", default="")
    ap.add_argument("--csv", default="")
    ap.add_argument("--max-win", type=int, default=400, help="F 结束位置往后看多少字")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = a.ledger or os.path.join(root, LEDGER)
    only = {int(x) for x in re.findall(r"\d+", a.rows)} if a.rows else None

    wb = openpyxl.load_workbook(ledger, data_only=True)
    ws = wb.active

    trunc, missed, unlocated, notext = [], [], 0, 0
    attention = []   # 「此外，我局还关注到…」这类提醒句（无违反定性）→ 观察项，非问题

    for r in range(2, ws.max_row + 1):
        if only and r not in only:
            continue
        f = str(ws.cell(r, 6).value or "").strip()
        rel = str(ws.cell(r, 13).value or "").strip()
        files = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        if not f or not files:
            continue
        inst = pick_inst_file(files)
        prim = inst or (files[0] if len(files) == 1 else "")
        order = ([prim] if prim else []) + [x for x in files if x != prim]

        txt, how = "", ""
        for fn in order:
            p = os.path.join(root, rel, fn)
            if not os.path.isfile(p):
                continue
            txt, how = read_text(p), "text"
            if len(txt) >= 200:
                prim = fn
                break
        if len(txt) < 200:
            notext += 1
            continue

        end = locate_end(norm(f), norm(txt))
        if end < 0:
            unlocated += 1  # 交给 Z11 / X29 处理（页脚串味、人员视角等）
            continue

        # ① 截断：末字不是句末标点
        if not f.endswith(("。", "；", ";")) and len(f) > 30:
            trunc.append((r, prim, f[-40:]))

        # ② 漏条：往后看，遇到决定段标志就停
        win = txt[end:end + a.max_win]
        m_stop = STOP.search(win)
        if m_stop:
            win = win[:m_stop.start()]
        m = NEXT_ITEM.search(win)
        if m:
            mk = m.group(1)
            if mk.startswith("此外"):
                seg = win[m.start():]
                ms = STOP.search(seg)
                if ms:
                    seg = seg[:ms.start()]
                if not VIOL_TAIL.search(seg):
                    attention.append((r, prim, seg[:90]))  # 关注事项，非违规条目
                    m = None
            if m:
                missed.append((r, prim, mk, f[-40:], txt[end:end + 120]))

    wb.close()

    print("=" * 78)
    print("Z14 事由条目完整性探测（F 列「只抄了一半」）")
    print("=" * 78)
    print("疑似漏条 : %d" % len(missed))
    print("疑似截断 : %d" % len(trunc))
    print("F 无法在原文定位（归 Z11）: %d" % unlocated)
    print("无文字层 : %d" % notext)
    print("观察项（「此外…还关注到」非违规条目）: %d" % len(attention))

    for r, fn, mk, tail, ctx in missed:
        print(" 行%-4d 后续出现「%s」 %s\n        F 尾部：%s\n        原文续：%s"
              % (r, mk, fn[:34], tail, ctx[:110]))
    for r, fn, tail in trunc:
        print(" 行%-4d 截断 %s\n        F 尾部：%s" % (r, fn[:34], tail))
    for r, fn, seg in attention:
        print(" [观察] 行%-4d %s\n        「此外」后续无违反定性：%s" % (r, fn[:30], seg[:80]))

    if a.csv:
        with open(a.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["台账行", "类型", "文书", "明细"])
            for r, fn, mk, tail, ctx in missed:
                w.writerow([r, "漏条:" + mk, fn, ctx[:200]])
            for r, fn, tail in trunc:
                w.writerow([r, "截断", fn, tail])
            for r, fn, seg in attention:
                w.writerow([r, "观察:此外关注事项", fn, seg[:200]])
        print("CSV 已写：%s" % a.csv)

    return 0 if not (missed or trunc) else 1


if __name__ == "__main__":
    sys.exit(main())
