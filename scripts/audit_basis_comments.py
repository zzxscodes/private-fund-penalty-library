# -*- coding: utf-8 -*-
"""Z15：H 列批注体检（台账↔法规正文的一致性）。

为什么需要它：
  - 批注不是「贴上去就永远对」——06 目录换了版本、法条正文变了、有人手工编辑过批注，
    台账看起来照旧，鼠标一悬停就是旧文。
  - 判据不能是「有没有批注」这么弱，而是**用同一套生成逻辑复算一遍，逐字比对**。

检查项：
  Z15-a 有 H 值但没有批注          → 问题
  Z15-b 批注与当前法规库复算结果不一致 → 问题（陈旧/手改）
  Z15-c 批注里出现「未收录/未找到/抽取不可靠」 → 观察（需人工核原文，不算错）
  Z15-d 批注存在但 H 已清空        → 问题（残留批注会误导）
  Z15-e 批注「逐字竖排」（一行一字） → 问题（解析没归一段内硬换行；重算同源，只能靠它抓）
  Z15-f 批注框缺锚点/残留 px 单位   → 问题（Excel 会按默认小框显示，尺寸等于没设）

退出码：0 无问题 / 1 有问题 / 2 环境不满足（缺法规库）
"""

import argparse
import csv
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

LEDGER = "04_综合索引台账【总表】.xlsx"
REG = "06_以上各类所有案例的配套法规依据"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--csv", default="", help="问题明细导出")
    args = ap.parse_args()

    import openpyxl

    import basis_comment
    import lawtext

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, LEDGER)
    reg = os.path.join(root, REG)
    cache = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_law_index", "index.json")
    if not os.path.isdir(reg):
        print("✗ 找不到法规目录：%s" % reg)
        return 2

    idx, hit = lawtext.load_or_build(reg, cache, ocr=True)
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]

    out, obs = [], []
    n_h, n_cm = 0, 0
    for r in range(2, ws.max_row + 1):
        h = str(ws.cell(r, 8).value or "").strip()
        cm = ws.cell(r, 8).comment
        text = (cm.text or "") if cm else ""
        if h:
            n_h += 1
        if text:
            n_cm += 1
            # Z15-e 逐字竖排：条文本应整段折行，若出现大量「一行只有一个字」，
            # 说明解析把段内硬换行当成了段落（2026-11-03：某 docx 有 6950 处，
            # 批注显示成竖排单字，且行数暴涨把框顶到上限、正文被裁）。
            # 这一条**不能靠 Z15-b 兜底**：重算用的是同一个解析器，坏得一样时会「自洽通过」。
            vn, vl = sum(1 for ln in text.split("\n") if len(ln.strip()) == 1), len(text.split("\n"))
            if vn >= 5:
                out.append((r, "Z15-e", "批注逐字竖排：%d/%d 行是单字（解析未归一段内硬换行）" % (vn, vl)))
                continue
        if h and not text:
            out.append((r, "Z15-a", "H 有值但无批注"))
            continue
        if not h and text:
            out.append((r, "Z15-d", "H 为空但残留批注"))
            continue
        if not h:
            continue
        expect, _ = basis_comment.build(idx, h)
        if expect != text:
            # 给出第一处差异，便于定位
            i = next((k for k in range(min(len(expect), len(text))) if expect[k] != text[k]),
                     min(len(expect), len(text)))
            out.append((r, "Z15-b", "批注与法规库复算不一致（第 %d 字起：批注「%s」／应为「%s」）"
                        % (i, text[i:i + 24].replace("\n", " "),
                           expect[i:i + 24].replace("\n", " "))))
            continue
        for kw in ("未收录", "未找到", "抽取不可靠"):
            if kw in text:
                obs.append((r, "Z15-c", "批注含「%s」，需人工核原文" % kw))
                break

    # ---- Z15-f 批注框结构：尺寸要 Excel 认，必须同时有「pt 单位」和「锚点」。
    # 少任何一个，Excel 都按默认 144×79 小框显示——批注写得再全也看不全（2026-11-03 二次反馈）。
    vml_note = ""
    try:
        import re as _re
        import zipfile
        with zipfile.ZipFile(ledger) as z:
            vs = [n for n in z.namelist() if n.lower().endswith(".vml")]
            tot = anc = px = cov = 0
            for n in vs:
                v = z.read(n).decode("utf-8")
                tot += len(_re.findall(r"<(?:\w+:)?shape\b", v))
                anc += len(_re.findall(r"<(?:\w+:)?Anchor>", v))
                px += len(_re.findall(r"(?<![\w-])(?:width|height):[\d.]+px", v))
                # 框不许压 H 列：压住了鼠标就点不回原格，收不起批注（2026-11-04 用户反馈）
                for _a in _re.findall(r"<(?:\w+:)?Anchor>([^<]+)</(?:\w+:)?Anchor>", v):
                    _p = [int(x) for x in _re.findall(r"-?\d+", _a)]
                    if len(_p) >= 5 and _p[4] >= 7:
                        cov += 1
        if tot == 0:
            out.append((0, "Z15-f", "找不到批注图形（VML 缺失/为空）"))
        elif anc < tot or px or cov:
            out.append((0, "Z15-f", "批注框结构不完整：shape %d｜锚点 %d｜残留 px %d｜压住 H 列 %d —— 缺锚点/残留 px 会让 Excel 按默认小框显示；压住 H 列则点不回原格" % (tot, anc, px, cov)))
        else:
            vml_note = "｜框结构 OK：%d 个 shape 均有锚点、无 px 残留、均不压 H 列" % tot
            # 顺带报一下「有多少行靠缩字号才塞进一屏」——这是它们能一屏看全的唯一原因
            with zipfile.ZipFile(ledger) as z2:
                szs = []
                for n in z2.namelist():
                    if n.startswith("xl/comments") and n.endswith(".xml"):
                        szs += _re.findall(r'<sz val="([\d.]+)"/>', z2.read(n).decode("utf-8"))
            if szs:
                from collections import Counter
                c = Counter(szs)
                vml_note += "｜缩字号 %d 行（%s）" % (len(szs), "，".join(
                    "%spt×%d" % (k, v) for k, v in sorted(c.items())))
    except Exception as e:        # noqa: BLE001
        out.append((0, "Z15-f", "无法读取 VML：%s" % e))

    print("Z15 批注体检：H 非空 %d 行｜批注 %d 个｜%s%s" % (n_h, n_cm, "法规库缓存命中" if hit else "法规库已重建", vml_note))
    print("  问题 %d ｜观察 %d" % (len(out), len(obs)))
    for x in out[:20]:
        where = "全局  " if x[0] == 0 else "行%-4d" % x[0]
        print("   ✗ %s %s %s" % (where, x[1], x[2]))
    if len(out) > 20:
        print("   …… 其余 %d 条见 CSV" % (len(out) - 20))
    if obs:
        print("   观察（不算错）：%d 行批注提示需核原文，前 5：%s"
              % (len(obs), [x[0] for x in obs[:5]]))

    if args.csv and (out or obs):
        os.makedirs(os.path.dirname(args.csv) or ".", exist_ok=True)
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["台账行", "判据", "说明"])
            w.writerows(out + obs)
        print("   明细：%s" % args.csv)
    return 1 if out else 0


if __name__ == "__main__":
    sys.exit(main())
