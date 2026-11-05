# -*- coding: utf-8 -*-
"""Z16 —— H 列「法条挂错法规」体检（只读，不改台账）。

为什么单列一个检查器：`fix_basis_attribution.py` 是**修复**工具，平时没人跑；
而 H 列归错会**静默**存在（错挂到的法规恰好也有同号条文时，任何形态检查都看不出来）。
所以必须有只读版本进 `run_selfcheck`，让每次自检都替你把 508 格 H 的法规归属对一遍文书。

判定沿用 `fix_basis_attribution.scan()`（**同一实现**，不另写一套）：
  ①台账 (法规, 条号) 在本行机构文书里找不到该法规带该条号的引用；
  ②文书里确有另一部法规带同一条号；③这样的法规唯一 → 才算「该改而未改」（问题）。

输出约定（run_selfcheck 靠这两行取数）：
    问题 N 条｜观察 M 条
其中：
  问题 —— 能定位到唯一正确法规、但台账还是错的（跑 fix_basis_attribution --apply 即可修）；
  观察 —— 歧义（多部法规都有同号，需人工判）／文书未见（可能是文书未到、或文书本身没写）。

用法：
    python check_basis_attribution.py --root <库根>
    python check_basis_attribution.py --root <库根> --rows 13,131 --detail
"""

import argparse
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--rows", default="")
    ap.add_argument("--detail", action="store_true", help="逐条打印观察项")
    a = ap.parse_args()

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import fix_basis_attribution as F

    rows = [int(x) for x in a.rows.split(",") if x.strip()] if a.rows else None
    sc = F.scan(a.root, a.ledger or None, rows)
    plans, report = sc["plans"], sc["report"]

    kinds = {}
    for item in report:
        kinds.setdefault(item[1], []).append(item)

    fixable = len(plans)
    amb = kinds.get("歧义跳过", [])
    unseen = kinds.get("文书未见", [])
    skip = kinds.get("跳过", [])
    # 「自检失败」里有两类，必须分开：
    #   重复类 —— 归位后会变成两条一模一样的，工具**故意**不动（如原文是说理段），属观察；
    #   其余（条号数变化 / 归属异常）—— 说明该改却没改成，是真问题。
    selfbad = [x for x in kinds.get("自检失败", []) if "重复" not in x[2]]
    dupbad = [x for x in kinds.get("自检失败", []) if "重复" in x[2]]
    obs = len(amb) + len(unseen) + len(skip) + len(dupbad)

    print("Z16 H 列法条归属体检：扫描 %d 行" % (len(plans) + len({x[0] for x in report})))
    print("问题 %d 条｜观察 %d 条" % (fixable + len(selfbad), obs))
    if fixable:
        print("\n【可自动归位（跑 fix_basis_attribution.py --apply）】%d 行" % fixable)
        for r, old, new, moves in plans[:15]:
            desc = "；".join("%s %s→%s" % (L, A, L2)
                            for (L, A), L2 in sorted(moves.items()))
            print("  行%-4d %s" % (r, desc))
        if fixable > 15:
            print("  …（其余 %d 行）" % (fixable - 15))
    if selfbad:
        print("\n【自检未通过（改动会引入重复/丢条，需人工看）】%d 条" % len(selfbad))
        for x in selfbad[:10]:
            print("  行%-4d %s" % (x[0], x[2][:100]))
    if dupbad:
        print("\n【归位会造成重复条目，工具故意不动（观察）】%d 条" % len(dupbad))
        for x in dupbad[:10]:
            print("  行%-4d %s" % (x[0], x[2][:100]))
    if amb and a.detail:
        print("\n【歧义（多部法规都有同号，需人工判）】%d 条" % len(amb))
        for x in amb[:20]:
            print("  行%-4d %s" % (x[0], x[2][:100]))
    if unseen and a.detail:
        print("\n【文书未见（可能文书未到，或文书本身没写这条）】%d 条" % len(unseen))
        for x in unseen[:30]:
            print("  行%-4d %s" % (x[0], x[2][:100]))

    return 1 if (fixable or selfbad) else 0


if __name__ == "__main__":
    sys.exit(main())
