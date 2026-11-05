# -*- coding: utf-8 -*-
"""2026-11-02 G 列（核心证据材料）漏抽定点补齐。

背景：G 列 451/508 为「未列举证据材料」，抽查发现 3 行原文明确写了证据句却被
写成「未列举」。本脚本只动这 3 行，且带「证据项必须在原文回现」的硬断言，
防止把别的案子的证据串到本行。

用法：
    python fix_evidence_20261102.py --root <库根>            # 预览
    python fix_evidence_20261102.py --root <库根> --apply    # 落盘（自动备份台账）
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

# (库内相对目录, 文件名, 证据项清单, 原文必须回现的关键词)
FIXES = {
    106: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/05_基金财产挪用、混同、保本保收益",
        "四川商投资本股权投资基金管理有限公司.pdf",
        "合伙协议及补充协议、咨询顾问协议、银行账户资金流水、询问笔录等证据",
        ["合伙协议及补充协议", "咨询顾问协议", "银行账户资金流水", "询问笔录"],
    ),
    213: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/02_登记备案、信息报送、信披违规",
        "深圳前海玖瀛资产管理有限公司（2026-05-29）.pdf",
        "交易流水，银行流水，相关人员询问笔录，登记备案信息，玖瀛资产及腾创投资提供、报送信息等证据",
        ["交易流水", "银行流水", "相关人员询问笔录", "登记备案信息"],
    ),
    402: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/01_募集与适当性违规",
        "深圳前海中证长凯投资管理有限公司.pdf",
        "相关协议、银行流水、询问笔录、事实确认书等证据",
        ["相关协议", "银行流水", "询问笔录", "事实确认书"],
    ),
    # 行251 属「只有人员文书」型（市场禁入决定书只针对个人，无机构文书），
    # 按「先到先填」取人员文书的证据段，合规。
    251: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/02_登记备案、信息报送、信披违规",
        "李某龙-上海雷根资产管理有限公司.pdf",
        "基金公告、基金合同、中国证券投资基金业协会提供的资料、托管人提供的资料、相关人员询问笔录等证据",
        ["基金公告", "基金合同", "托管人提供的资料", "相关人员询问笔录"],
    ),
    # 408/449 是同一批事实的中基协处分（机构件与个人件），两份文书的证据段相同。
    # 这两份都是扫描件，靠 OCR 才看得到证据句。
    408: (
        "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）/05_基金财产挪用、混同、保本保收益",
        "尉亲-浙江优策投资管理有限公司.pdf",
        "登记备案信息、基金产品合同、基金产品提供、报送信息、银行流水、工商资料、现场走访情况、投资者适当性材料、相关人员询问笔录等证据",
        ["登记备案信息", "基金产品合同", "银行流水", "工商资料", "相关人员询问笔录"],
    ),
    449: (
        "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）/02_登记备案、信息报送、信披违规",
        "浙江优策投资管理有限公司（2025-11-25）.pdf",
        "登记备案信息、基金产品合同、基金产品提供、报送信息、银行流水、工商资料、现场走访情况、投资者适当性材料、相关人员询问笔录等证据",
        ["登记备案信息", "基金产品合同", "银行流水", "工商资料", "相关人员询问笔录"],
    ),
}


def read_pdf_text(path, cache_dir=None):
    """优先文字层；扫描件（<200 字）走 OCR。

    中基协的纪律处分决定书大量是扫描件，证据句只在 OCR 后可见
    （行408/449 实测：文字层为空，OCR 后才有「…等证据证明」）。
    OCR 缓存一律派生到库**外**的维护目录，避免污染库根白名单。
    """
    import fitz

    doc = fitz.open(path)
    try:
        t = re.sub(r"\s+", "", "".join(p.get_text() for p in doc))
    finally:
        doc.close()
    if len(t) >= 200:
        return t, "text"
    try:
        import ocrlib

        return re.sub(r"\s+", "", ocrlib.ocr_text(path, cache_dir=cache_dir)), "ocr"
    except Exception as e:
        return t, "ocr_err:%s" % e


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--apply", action="store_true", help="真正写盘（默认只预览）")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LEDGER)
    # OCR 缓存放库外，避免污染库根目录白名单
    cache_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")
    if not os.path.isfile(ledger):
        sys.exit("台账不存在：%s" % ledger)

    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    # 先全部校验，任何一条断言不过就整体不写（避免半改状态）
    plan, bad = [], []
    for row, (rel, fname, evid, needles) in sorted(FIXES.items()):
        p = os.path.join(root, rel, fname)
        if not os.path.isfile(p):
            bad.append("行%d 文件缺失：%s" % (row, p))
            continue
        txt, _how = read_pdf_text(p, cache_dir)
        miss = [n for n in needles if n not in txt]
        if miss:
            bad.append("行%d 证据项未在原文回现：%s" % (row, "／".join(miss)))
            continue
        old = str(ws.cell(row, 7).value or "").strip()
        if old == evid:
            continue  # 已是目标值：幂等跳过，重跑不报错
        if "未列举" not in old:
            bad.append("行%d G 列已非占位且不等于目标值（当前=%r），请人工确认后移除该 FIXES 项" % (row, old[:40]))
            continue
        plan.append((row, old, evid))

    if bad:
        print("✗ 校验未通过，未写盘：")
        for b in bad:
            print("   -", b)
        sys.exit(1)

    print("待修 %d 行：" % len(plan))
    for row, old, new in plan:
        print("  行%-4d %s  →  %s" % (row, old, new))

    if not a.apply:
        print("\n（预览模式，加 --apply 落盘）")
        return

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ledger.replace(".xlsx", "_bak_evidence_%s.xlsx" % stamp)
    shutil.copy2(ledger, bak)

    for row, _old, new in plan:
        ws.cell(row, 7).value = new
    wb.save(ledger)
    wb.close()

    print("\n✓ 已写入 %d 行" % len(plan))
    print("  台账备份：%s" % bak)


if __name__ == "__main__":
    main()
