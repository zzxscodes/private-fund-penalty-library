# -*- coding: utf-8 -*-
"""行86（上海忠石 / 马彩茸）F、H 列来源归位 —— Z11 抓到的唯一「跨文书串源」。

问题（Z11 实测 覆盖 0%）：
    台账 F = 「被停职后超越授权操作基金产品证券账户，损害基金财产和投资者利益。」
    这句在**机构文书（责令改正）里根本没有**，它出自**高管马彩茸的个人警示函**。
    机构文书的事由是另一句：「内部控制不健全，未履行谨慎勤勉义务」。
    H 列同理：第二十三条第六项是个人决定书引的条款，机构文书引的是第四条第一款。

为什么必须改（不是「不好看」，是口径被破）：
    §15.1 F/G/H/I **只认机构文书**。这一行把个人文书的内容写进了机构视角的列，
    后果是这一行的「处罚事由」与「处罚措施（责令改正）」对不上——
    责令改正的原因写成了个人的越权交易。Z11 之所以只报这一行，
    是因为全库 402 行里只有这一行越界（其余 401 行 F 都能在机构文书回现），属离群，不是口径问题。

代价（必须记下来，别当没发生）：
    「被停职后超越授权操作基金产品证券账户」这句细节会从台账 F 列消失，只留在 PDF 里。
    这是「F 只认机构文书」这条口径的必然结果——机构文书本身只写内控不健全。
    若以后要保留个人层面的具体行为，应新增列（如「人员违规行为」），而不是往 F 里混。

用法：
    python fix_src_row86_20261102.py --root <库根>            # 干跑
    python fix_src_row86_20261102.py --root <库根> --apply    # 落盘（自动备份台账）
"""
import argparse
import os
import re
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import fitz
import openpyxl

LEDGER = "04_综合索引台账【总表】.xlsx"
ROW = 86

NEW_F = "内部控制不健全，未履行谨慎勤勉义务。"
NEW_H = "《私募投资基金监督管理暂行办法》第四条第一款；《私募投资基金监督管理暂行办法》第三十三条"

# 归位后 F 必须能在机构文书里回现；H 的每个条款号也必须能在机构文书里找到
NEEDLES = ["内部控制不健全", "未履行谨慎勤勉义务", "第四条第一款", "第三十三条"]
# 旧 F 应当能在**个人**文书里回现（证明它确实串了源，而不是我们瞎猜）
OLD_F_PROOF = ["被停职后超越授权操作基金产品证券账户", "第二十三条第六项"]


def read_pdf(path):
    d = fitz.open(path)
    try:
        return re.sub(r"\s+", "", "".join(p.get_text() for p in d))
    finally:
        d.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LEDGER)
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    rel = str(ws.cell(ROW, 13).value or "").strip()
    files = [x.strip() for x in re.split(r"[；;]", str(ws.cell(ROW, 12).value or "")) if x.strip()]
    inst = [f for f in files if "上海忠石" in f and not f.startswith("马彩茸")]
    per = [f for f in files if f.startswith("马彩茸")]
    if not inst or not per:
        print("✗ 行%d 文件清单不符预期：%s" % (ROW, files))
        return 1
    inst_p = os.path.join(root, rel, inst[0])
    per_p = os.path.join(root, rel, per[0])
    itxt, ptxt = read_pdf(inst_p), read_pdf(per_p)

    old_f = str(ws.cell(ROW, 6).value or "").strip()
    old_h = str(ws.cell(ROW, 8).value or "").strip()
    print("机构文书：%s（%d 字）" % (inst[0], len(itxt)))
    print("人员文书：%s（%d 字）" % (per[0], len(ptxt)))
    print("旧 F：%s" % old_f)
    print("新 F：%s" % NEW_F)
    print("旧 H：%s" % old_h)
    print("新 H：%s" % NEW_H)

    # 护栏 ①：新值必须在机构文书里回现
    miss = [n for n in NEEDLES if n not in itxt]
    if miss:
        print("✗ 新值未在机构文书回现：%s" % miss)
        return 1
    # 护栏 ②：旧值必须在个人文书里回现（证明「串源」判断成立）
    miss2 = [n for n in OLD_F_PROOF if n not in ptxt]
    if miss2:
        print("✗ 旧值未在人员文书回现，串源判断不成立，中止：%s" % miss2)
        return 1
    # 护栏 ③：旧值不该在机构文书里（若在，说明机构文书也有这段，不必改）
    if "被停职后超越授权" in itxt:
        print("✗ 机构文书也含该事实，本行不是串源，中止")
        return 1

    if old_f == NEW_F and old_h == NEW_H:
        print("→ 已是目标值，跳过（幂等）")
        wb.close()
        return 0

    if not a.apply:
        print("\n（干跑结束，加 --apply 落盘）")
        wb.close()
        return 0

    bak_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "台账备份")
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, "04_综合索引台账【总表】_bak_row86src_%s.xlsx"
                       % datetime.now().strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ledger, bak)
    print("备份：%s" % bak)

    ws.cell(ROW, 6).value = NEW_F
    ws.cell(ROW, 8).value = NEW_H
    wb.save(ledger)
    wb.close()

    # 落盘后复验（不信内存，重开文件读）
    wb2 = openpyxl.load_workbook(ledger, data_only=True)
    ws2 = wb2.active
    ok = (str(ws2.cell(ROW, 6).value) == NEW_F and str(ws2.cell(ROW, 8).value) == NEW_H)
    wb2.close()
    print("落盘复验：%s" % ("OK" if ok else "✗ 未写入"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
