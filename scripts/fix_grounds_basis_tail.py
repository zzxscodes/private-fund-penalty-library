# -*- coding: utf-8 -*-
"""删掉 F 列（核心处罚事由）末尾残留的**处罚依据句**。

症状（2026-11-02 实测 27 行）：
    F = 「…违反《私募办法》第二十四条的规定。**根据《私募办法》第三十三条的规定。**」
    最后那句是「根据哪一条作出处罚」，属于 H（核心处罚依据），不是「事由」。

    X29/GROUNDS_CUT 只截「我局决定」「以上事实有」这类，管不到「根据《…》第…条」，
    所以这些行一直带着依据尾巴。

判据（四道，全过才删）：
    ① 最后一句以「根据《」开头（前面允许有逗号/顿号）；
    ② 这句里出现「第…条」（纯引用，不是叙述）；
    ③ 这句里**没有**「约定/合同/据此/构成/不符合/违反」——
       行40「根据《托管协议》的约定…支付给管理人」、行247「据此，根据《…》第十五条规定，
       …构成骗取管理人登记的行为」是事实句，不删；
    ④ 删完剩余 ≥15 字且以句末标点收尾。

用法：
    python fix_grounds_basis_tail.py --root <库根>            # 干跑
    python fix_grounds_basis_tail.py --root <库根> --apply
"""
import argparse
import os
import re
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl

LEDGER = "04_综合索引台账【总表】.xlsx"
START = re.compile(r"^[，,、]?\s*根据《")
CLAUSE = re.compile(r"第[一二三四五六七八九十百零〇\d]+条")
KEEP = re.compile(r"约定|合同|据此|构成|不符合|违反|存在")


def split_last(f):
    """返回 (前缀, 最后一句)。按 。/；/换行 切，保留分隔符在前缀里。"""
    m = list(re.finditer(r"[。；;\n]", f))
    if not m:
        return "", f
    p = m[-1].end()
    return f[:p], f[p:]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LEDGER)
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    plan, skipped = [], []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        f = str(ws.cell(r, 6).value or "")
        if "根据《" not in f:
            continue
        # 先去掉末尾标点再切，否则最后一段永远是空串
        core = f.rstrip().rstrip("。；; ")
        head, last = split_last(core)
        if not last.strip():
            continue
        if not START.match(last.strip()):
            if "根据《" in last:
                skipped.append((r, "根据句不是最后一句（句中）", last.strip()[:60]))
            continue
        if not CLAUSE.search(last):
            skipped.append((r, "无条款号，不像依据句", last.strip()[:60]))
            continue
        if KEEP.search(last):
            skipped.append((r, "含叙述性词，保留", last.strip()[:60]))
            continue
        rest = head.rstrip("，,、 ")
        if len(re.sub(r"\s", "", rest)) < 15:
            skipped.append((r, "删完过短", last.strip()[:60]))
            continue
        plan.append((r, f, rest))

    print("可删依据尾句 %d 行｜跳过 %d 行" % (len(plan), len(skipped)))
    for r, _old, rest in plan:
        print("  行%-4d 保留结尾：…%s" % (r, rest[-56:]))
    for r, why, seg in skipped:
        print("  跳过 行%-4d %-24s %s" % (r, why, seg))

    if not a.apply:
        print("\n（干跑结束，加 --apply 落盘）")
        wb.close()
        return 0
    if not plan:
        wb.close()
        return 0

    bak_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "台账备份")
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, "04_综合索引台账【总表】_bak_basistail_%s.xlsx"
                       % datetime.now().strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ledger, bak)
    print("\n备份：%s" % bak)

    for r, _old, rest in plan:
        ws.cell(r, 6).value = rest
    wb.save(ledger)
    wb.close()

    wb2 = openpyxl.load_workbook(ledger, data_only=True)
    ws2 = wb2.active
    left = [r for r, _o, _n in plan if re.search(r"根据《[^》]*》[^。]{0,60}第[一二三四五六七八九十百零〇\d]+条[^。]*。?$",
                                                 str(ws2.cell(r, 6).value or ""))]
    wb2.close()
    print("落盘完成｜仍带依据尾句：%s" % (left or "无"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
