#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_subject_20261102b.py —— 两处 D 列/措施列实证修复（2026-11-02）

两处都是「按文书原文逐字核对」后才动的，不是猜：

① 行33（序号32）杭州复投：D 列第三人是「新享」。
   措施列只有徐勇升、孙勇两段 + 管理人段；事由/申辩/证据全无「新享」；
   平台上该案只有 3 份文书（3 条链接 ↔ 3 个文件），没有第 4 个人。
   ⇒ 提取碎片，删。

② 行366（序号365）马鞍山哲方：D 列「乔徽、李皓」，漏了李磊。
   决定书首页原文：名称「关于对乔徽、李皓、李磊采取出具警示函措施的决定」，
   抬头「乔徽、李皓、李磊：」，正文「乔徽、李皓未谨慎勤勉履行相关职责……
   我局决定对其采取出具警示函……李磊未谨慎勤勉履行相关职责……我局决定
   对其采取出具警示函……」。
   ⇒ D 补李磊；措施按原文拆成三段（乔徽、李皓 共用一句；李磊 单独一句），
     不再出现「高管（乔徽）」段里夹着李磊措施的情况。

用法：python scripts/fix_subject_20261102b.py --root <库根> [--apply]
脚本自带前置断言：目标单元格必须与预期原文一致，否则整行跳过（防止行号漂移后改错行）。
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

C_SUBJECT, C_MEAS = 4, 5

EXPECT = [
    {
        "tag": "杭州复投·删碎片",
        "subj_head": "杭州复投私募基金管理合伙企业（有限合伙）",
        "subj_tail": "徐勇升、孙勇、新享",
        "new_subj": "杭州复投私募基金管理合伙企业（有限合伙）；徐勇升、孙勇",
        "meas_must": "高管（孙勇）",
        "new_meas": None,          # 措施不改
    },
    {
        "tag": "马鞍山哲方·补李磊并拆段",
        "subj_head": "马鞍山哲方智能机器人投资管理有限公司",
        "subj_tail": "乔徽、李皓",
        "new_subj": "马鞍山哲方智能机器人投资管理有限公司；乔徽、李皓、李磊",
        "meas_must": "李磊未谨慎勤勉履行相关职责",
        "new_meas": (
            "高管（乔徽、李皓）：乔徽、李皓未谨慎勤勉履行相关职责，违反了《私募投资基金监督管理条例》"
            "第三条第四款、《私募投资基金监督管理暂行办法》第四条第二款的规定，根据《私募投资基金监督"
            "管理条例》第四十条第二款、《私募投资基金监督管理暂行办法》第三十三条的规定，我局决定对其"
            "采取出具警示函的行政监管措施，并记入证券期货市场诚信档案；\n"
            "高管（李磊）：李磊未谨慎勤勉履行相关职责，违反了《私募投资基金监督管理暂行办法》第四条"
            "第二款的规定，根据《私募投资基金监督管理暂行办法》第三十三条的规定，我局决定对其采取出具"
            "警示函的行政监管措施，并记入证券期货市场诚信档案；\n"
            "管理人：对你公司采取责令改正的行政监管措施，并记入证券期货市场诚信档案。你公司应在收到本"
            "决定书之日起30日内向我局提交书面整改报告"
        ),
    },
]


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

    todo = []
    for spec in EXPECT:
        hit = None
        for r in range(2, ws.max_row + 1):
            subj = str(ws.cell(r, C_SUBJECT).value or "").strip()
            meas = str(ws.cell(r, C_MEAS).value or "")
            if subj.startswith(spec["subj_head"]) and subj.endswith(spec["subj_tail"]) \
                    and spec["meas_must"] in meas:
                hit = r
                break
        if hit is None:
            print("  ✗ %s：未找到匹配行（可能已修过）" % spec["tag"])
            continue
        todo.append((hit, spec))
        print("  ✓ %s → 行%d 序号%s" % (spec["tag"], hit, ws.cell(hit, 1).value))
        print("      D: %s" % ws.cell(hit, C_SUBJECT).value)
        if spec["new_meas"]:
            print("      E: %s…" % spec["new_meas"][:40].replace("\n", " "))

    if not a.apply:
        print("\n（dry-run；--apply 落盘）")
        return
    if not todo:
        print("\n无需改动。")
        return
    from ledger_bak import backup_ledger
    print("\n台账已备份：%s" % backup_ledger(root, ledger, "fix_subject_b"))
    for r, spec in todo:
        ws.cell(r, C_SUBJECT).value = spec["new_subj"]
        if spec["new_meas"]:
            ws.cell(r, C_MEAS).value = spec["new_meas"]
        print("  已改 行%-4d %s" % (r, spec["tag"]))
    wb.save(ledger)
    print("\n完成 %d 处" % len(todo))


if __name__ == "__main__":
    main()
