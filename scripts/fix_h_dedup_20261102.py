#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_h_dedup_20261102.py —— H 列（核心处罚依据）同法规条款去重（2026-11-02）

背景（用户 2026-11-02 口径）：
  台账除「处罚措施（机构+人员）」需要合并机构与人员文书外，其余各列**只从机构文书取**，
  且一个案例的机构文书只有一份。H 列出现「同一部法、同一批条款写两遍」，就是这项口径
  被破坏留下的痕迹 —— 人员文书的依据被追加到了机构依据后面：
      行39 《私募基金监管办法》第四条第一款、第二十四条；《私募基金监管办法》第四条第一款
      行36 整组 6 段重复 2 遍（机构原决定 + 复核决定各一遍）

做法（只去重、不改写）：
  1. 按「；」切段，每段解析成 (法规名, [条款…])
  2. 同法规名按首次出现顺序聚合，条款合并去重（保持原有先后）
  3. 无法解析的段（无《》或纯说明）原样保留在最后，同样去重
  4. 断言：法规名集合不变、条款集合不变（只允许减少重复项）

用法：
  python scripts/fix_h_dedup_20261102.py --root <库根>            # 干跑
  python scripts/fix_h_dedup_20261102.py --root <库根> --apply    # 落盘
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

from ledger_bak import backup_ledger  # noqa: E402

LAW = re.compile(r"《([^》]+)》(.*)$", re.S)
ART = re.compile(r"第[一二三四五六七八九十百零\d]+条(?:第[一二三四五六七八九十\d]+[款项])?"
                 r"(?:第[一二三四五六七八九十\d]+项)?")


def parse(seg):
    """'《X》第A条、第B条' → ('《X》', ['第A条','第B条'])；解析失败返回 (None, [seg])"""
    m = LAW.match(seg.strip())
    if not m:
        return None, [seg.strip()]
    name, tail = m.group(1), m.group(2)
    arts = ART.findall(tail)
    if not arts:
        # 「《X》」后面没有条款号（如「《中华人民共和国证券投资基金法》」）→ 整段当无条款
        return "《%s》" % name, []
    return "《%s》" % name, arts


def dedup(text):
    """只删「冗余段」：某段的条款已被前面同法规段完全覆盖，或整段与前面完全相同。

    不做跨段合并 —— 同一部法分几段列（对应几个违规点）是原文写法，动了就改变了原样。
    """
    segs = [s.strip() for s in re.split(r"[；;]", text) if s.strip()]
    out, seen_arts, seen_seg = [], {}, set()
    for s in segs:
        if s in seen_seg:                      # 整段重复（多份文书拼接的典型痕迹）
            continue
        name, arts = parse(s)
        if name is not None:
            got = seen_arts.setdefault(name, [])
            if arts and all(a in got for a in arts):
                continue                       # 条款已被前面同法规段覆盖
            for a in arts:
                if a not in got:
                    got.append(a)
        seen_seg.add(s)
        out.append(s)
    return "；".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--show", type=int, default=12)
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]

    n = 0
    shown = 0
    for r in range(2, ws.max_row + 1):
        old = ws.cell(r, 8).value
        if not isinstance(old, str) or len(old.strip()) < 6:
            continue
        new = dedup(old)
        if new == old:
            continue
        # 安全断言：去重前后「法规名+条款」的多重集合必须一致（只允许去掉重复份）
        old_set = sorted(set(re.findall(r"《[^》]+》[^；;]*", old)))
        arts_o = sorted(ART.findall(old))
        arts_n = sorted(ART.findall(new))
        if len(arts_n) > len(arts_o):
            print("行%-4d 断言失败（条款变多），跳过" % r)
            continue
        n += 1
        if shown < args.show:
            print("\n行%-4d %s" % (r, str(ws.cell(r, 4).value or "")[:30]))
            print("  旧 %s" % old[:150])
            print("  新 %s" % new[:150])
            shown += 1
        if args.apply:
            ws.cell(r, 8).value = new

    print("\n去重 %d 行" % n)
    if args.apply and n:
        backup_ledger(root, ledger, "h_dedup_20261102")
        wb.save(ledger)
        print("已保存（备份在 _案例库维护留痕/台账备份/）")
    elif not args.apply:
        print("[干跑] 未落盘；加 --apply 执行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
