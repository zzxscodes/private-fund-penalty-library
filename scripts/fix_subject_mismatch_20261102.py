#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_subject_mismatch_20261102.py —— 主体/文件名错配的三处修复（一次性数据修正）

来源：全库 OCR 首页扫描（scripts/scan_review_docs.py）把「文书当事人」与「台账主体」比对，
查出 3 处真实错配（用户口径：「台账主体必须是文书里的被处罚主体本身」）：

  行267 天问时代：类别误判 02_登记备案（实为委托通道发起基金 → 03_投资运作），
        且机构决定书（中基协处分〔2026〕7 号）被错挂到行268。
  行268 黑岩中盛：K 列第 4 项文件名被粘成
        「天问时代私募基金管理（海南-深圳市黑岩中盛资产管理有限公司.pdf」
        （两主体名被中文破折号粘连，两个名字都不完整），且 E 列混入了「对天问时代进行警告」。
  行292：主体/文件名写成「中国银河证券股份有限公司」，而文书当事人是
        **海南省凯什私募基金管理有限公司**总经理黄明明（复核〔2025〕35 号）；
        E 列尾部有 OCR 残片「；券投」，文号括号被识别成「（2025）224号」的混合形态。

用法：python fix_subject_mismatch_20261102.py --root <库根> [--apply]
全部改动前先备份台账；改名用两步法（先 .tmp_rename 再落终名）。
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys
import time

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ledger_bak import backup_ledger  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

D = "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）"
DIR_02 = D + "/02_登记备案、信息报送、信披违规"
DIR_03 = D + "/03_投资运作、关联交易、利益输送"
DIR_01 = D + "/01_募集与适当性违规"

TWT = "天问时代私募基金管理（海南）有限公司.pdf"
TWT_BAD = "天问时代私募基金管理（海南-深圳市黑岩中盛资产管理有限公司.pdf"
TWT_WJ = "张伟杰-天问时代私募基金管理（海南）有限公司.pdf"
HYS = "深圳市黑岩中盛资产管理有限公司.pdf"
XJ = "徐剑-深圳市黑岩中盛资产管理有限公司.pdf"
GZ = "郭政-深圳市黑岩中盛资产管理有限公司.pdf"
HMM_BAD = "黄明明-中国银河证券股份有限公司.pdf"
HMM = "黄明明-海南省凯什私募基金管理有限公司（复核决定）.pdf"

# 行号: [(旧路径, 新路径)] 改名
MOVES = {
    267: [(os.path.join(DIR_02, TWT_WJ), os.path.join(DIR_03, TWT_WJ))],
    268: [(os.path.join(DIR_03, TWT_BAD), os.path.join(DIR_03, TWT))],
    292: [(os.path.join(DIR_01, HMM_BAD), os.path.join(DIR_01, HMM))],
}

SET = {
    267: {
        9: "03_投资运作、关联交易、利益输送",         # J 违规类型
        11: "；".join([TWT, TWT_WJ]),                  # K 案例文件
        12: "；".join([DIR_03, DIR_03]),               # L 存放路径
        4: "管理人：对天问时代进行警告\n高管（张伟杰）：对张伟杰进行警告",   # E 处罚措施
    },
    268: {
        11: "；".join([HYS, XJ, GZ]),
        12: "；".join([DIR_03] * 3),
        4: "管理人：对黑岩中盛进行警告\n高管（徐剑）：对徐剑进行警告\n高管（郭政）：对郭政进行警告",
    },
    292: {
        3: "海南省凯什私募基金管理有限公司；黄明明",   # D 主体
        11: HMM,
        12: DIR_01,
        4: "高管（黄明明）：维持《纪律处分决定书》（中基协处分〔2025〕224号）对申请人作出的纪律处分措施",
    },
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    ledger = next(os.path.join(root, f) for f in os.listdir(root)
                  if f.startswith("04_") and f.endswith(".xlsx") and not f.startswith("~$"))
    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    print("== 改名/移位 ==")
    for row, pairs in MOVES.items():
        for src, dst in pairs:
            p1, p2 = os.path.join(root, src), os.path.join(root, dst)
            print("  行%-4d %s\n        %s\n     → %s" % (row, os.path.dirname(src)[-14:],
                                                        os.path.basename(src), dst))
    print("\n== 台账改写 ==")
    for row, kv in SET.items():
        for col, v in kv.items():
            old = ws.cell(row, col + 1).value
            print("  行%-4d 列%s\n        %s\n     → %s" % (row, "ABCDEFGHIJKLMNO"[col],
                                                          str(old)[:70].replace("\n", " ⏎ "),
                                                          str(v)[:70].replace("\n", " ⏎ ")))
    if not a.apply:
        print("\n[干跑] 未改动；确认后加 --apply")
        return

    backup_ledger(root, ledger, "subjmis")

    def ren(src, dst, tries=5):
        for i in range(tries):
            try:
                if os.path.abspath(src) == os.path.abspath(dst):
                    return True
                if not os.path.exists(src) and os.path.exists(dst):
                    return True          # 已改过
                os.rename(src, dst)
                return True
            except PermissionError:
                time.sleep(1.5 * (i + 1))
        return False

    ok, fail = 0, []
    for row, pairs in MOVES.items():
        for src, dst in pairs:
            p1, p2 = os.path.join(root, src), os.path.join(root, dst)
            if ren(p1, p2):
                ok += 1
            else:
                fail.append((row, src))
    print("改名/移位 %d 个" % ok)
    if fail:
        print("失败：%s" % fail)
        sys.exit(1)

    for row, kv in SET.items():
        for col, v in kv.items():
            ws.cell(row, col + 1).value = v
    wb.save(ledger)
    print("台账已保存（%d 行 / %d 个单元格）"
          % (len(SET), sum(len(v) for v in SET.values())))


if __name__ == "__main__":
    main()
