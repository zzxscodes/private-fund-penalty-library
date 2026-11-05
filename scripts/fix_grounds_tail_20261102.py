# -*- coding: utf-8 -*-
"""F 列事由定点补全（2026-11-02，Z14 收官）。

批量脚本 `fix_grounds_multi_20261102.py` 有 4 条硬护栏，护栏挡下的行不是「判错了」，
而是「自动改不安全」。这 3 行就是这样，只能逐条看原文、手写目标值：

  行182 扫描/分页把「特征」劈成「特2征」，自动脚本不敢删数字 → 手写，并把「2」去掉；
  行211 第二条含「你公司」二次称呼，回现比对过不去 → 手写，统一成机构全称；
  行251 原文是「一、<违规名称><事实描述>」粘连体 + 第二条带「一是/二是」子项，
        拆分会撞序号（一二二三）→ 手写，保留子项在第二条内。

每行都断言「目标值的每一句都能在原文回现」，防止手写时把别的案子串进来。
行301 不改：其「此外，我局还关注到…」是**关注事项**而非违规条目（无违反定性），
已登记进 `校验人工复核豁免清单.md`。

用法：
    python fix_grounds_tail_20261102.py --root <库根> [--apply]
"""
import argparse
import os
import re
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl

from check_fields_source import best_cover
from extract_core import _grounds_csrc  # noqa: F401  (保留依赖关系，便于日后复用)

LEDGER = "04_综合索引台账【总表】.xlsx"
COVER_MIN = 0.80

FIXES = {
    182: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/01_募集与适当性违规",
        "佛山市华潮资产管理有限公司.pdf",
        "一、未谨慎勤勉履行管理人职责，未通过投资者基本信息表、资产证明材料和投资经历证明材料等有效方式了解投资者的背景情况和财务状况。\n"
        "二、未穿透核查作为基金产品投资者的合伙企业的合伙人是否为合格投资者。\n"
        "三、未基于基金产品投资标的流动性较差的实际风险特征和程度准确认定基金产品风险等级。",
    ),
    211: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/04_内控治理、人员资质、风控缺位",
        "国信投（海南）私募基金管理有限公司.pdf",
        "一、未及时填报并定期更新有关信息。\n"
        "二、国信投（海南）私募基金管理有限公司未与合规风控负责人建立劳动关系，合规风控负责人未在公司专职履职，说明公司内部管理薄弱，未履行谨慎勤勉义务。\n"
        "三、从事与私募基金管理无关的业务。",
    ),
    251: (
        "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）/02_登记备案、信息报送、信披违规",
        "李某龙-上海雷根资产管理有限公司.pdf",
        "一、向投资者承诺投资本金不受损失和最低收益。雷根资管与雷根添宝全天候一号私募证券投资基金（以下简称添宝全天候一号）、"
        "雷根元亨私募证券投资基金、雷根优享全天候五号私募证券投资基金（以下简称优享全天候五号）等的投资者签订私募投资基金份额转让协议，"
        "承诺投资本金不受损失和最低收益。\n"
        "二、从事损害基金财产和投资者利益的投资活动。一是在雷根有道基金主要资产为已终止挂牌的原新三板某挂牌公司股票，"
        "且该公司经营异常的情况下，雷根资管仍按股票终止挂牌前价格核算基金净值，并以添宝全天候一号按相关净值申购持有雷根有道基金份额，"
        "损害了添宝全天候一号投资者的利益。二是在雷根领先一号私募证券投资基金（以下简称领先一号）主要资产为已终止挂牌的原新三板某挂牌公司股票，"
        "且该公司经营状况恶化的情况下，雷根资管仍按股票终止挂牌前价格核算基金净值，并以雷根添宝全天候五号二期私募证券投资基金"
        "（以下简称添宝全天候五号二期）按相关净值申购持有领先一号份额，损害了添宝全天候五号二期投资者的利益。\n"
        "三、未按规定履行信息披露义务。根据《雷根优享全天候五号私募证券投资基金私募基金合同》，雷根资管应当在优享全天候五号的年度报告中"
        "披露基金投资运作情况、在季度报告中披露投资组合情况。优享全天候五号的《2021年年度报告》《2022年第二季度报告》披露的"
        "“期末基金资产组合情况”均显示无私募基金产品投资，与实际情况不符，导致优享全天候五号《2021年年度报告》"
        "《2022年第二季度报告》未如实披露相关投资情况。",
    ),
    245: (
        "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）/05_基金财产挪用、混同、保本保收益",
        "国彤万和私募基金管理有限公司.pdf",
        "一、与关联机构混同管理。"
        "其一，珠海万和锦鸿科技有限公司（以下简称万和锦鸿）为国彤万和全资子公司。2021年6月，珠海锦创企业管理中心（有限合伙）（以下简称珠海锦创）成立，"
        "普通合伙人、执行事务合伙人为万和锦鸿，另有17名有限合伙人（其中16人未经工商确权）。2021年8月，珠海锦创投向珠海锦鸿欣融房地产咨询服务合伙企业（有限合伙），"
        "投资金额为2.62亿元。珠海锦创投资过程中，万和锦鸿与国彤万和在高级管理人员、财务报销等方面高度混同。其二，国彤万和与其全资子公司"
        "珠海万和锦华资产管理有限公司（以下简称万和锦华，登记为私募基金管理人）存在财务混同的情况，未有效执行内部控制。国彤万和银行账户流水显示，"
        "2021年至2025年期间，国彤万和为其子公司万和锦华的23名员工频繁报销费用。\n"
        "二、使用关联方账户代私募基金收付基金财产。其一，8850万元资金通过底层投资标的张家口锦瑞欣融房地产开发有限公司（以下简称锦瑞欣融房地产）"
        "转至国彤万和关联方珠海景熠投资中心（有限合伙）（以下简称珠海景熠）。2021年8月，珠海景熠再将上述资金转回至“万和锦兴”。其二，"
        "国彤万和将在管产品“国投万和瑞欣3号契约型私募股权投资基金”（以下简称“瑞欣3号”）的11970万元资金通过底层投资标的锦瑞欣融房地产"
        "转至国彤万和关联方珠海景熠。此后珠海景熠再将上述资金转回至“瑞欣3号”。",
    ),
}


def read_any_text(path, cache_dir=None):
    import fitz

    d = fitz.open(path)
    try:
        t = re.sub(r"\s+", "", "".join(p.get_text() for p in d))
    finally:
        d.close()
    if len(t) >= 200:
        return t
    try:
        import ocrlib

        return re.sub(r"\s+", "", ocrlib.ocr_text(path, cache_dir=cache_dir))
    except Exception:
        return t


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LEDGER)
    cache_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")

    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    plan, bad = [], []
    for row, (rel, fn, new) in sorted(FIXES.items()):
        p = os.path.join(root, rel, fn)
        if not os.path.isfile(p):
            bad.append("行%d 文件缺失" % row)
            continue
        txt = read_any_text(p, cache_dir)
        # 回现比对要先抹平「你公司 → 机构全称」的称呼置换（台账已归一、原文没有），
        # 否则带主语的那一条永远过不去（行211 实测）。
        subj = str(ws.cell(row, 4).value or "").split("；")[0].strip()
        txt_cmp = txt.replace("你公司", subj).replace("你单位", subj) if subj else txt
        low = []
        for line in new.split("\n"):
            body = re.sub(r"^\s*[一二三四五六七八九十]+、", "", line)
            if best_cover(body, txt_cmp) < COVER_MIN:
                low.append(body[:26])
        if low:
            bad.append("行%d 回现不足：%s" % (row, "／".join(low)))
            continue
        old = str(ws.cell(row, 6).value or "").strip()
        if old == new.strip():
            continue
        plan.append((row, old, new))

    if bad:
        print("✗ 校验未通过，未写盘：")
        for b in bad:
            print("   -", b)
        return 1

    print("待修 %d 行" % len(plan))
    for row, old, new in plan:
        print("\n行%d\n  旧(%d 字)：%s\n  新(%d 字)：%s"
              % (row, len(old), old[:120].replace("\n", " / "), len(new), new[:200].replace("\n", " / ")))
    if not a.apply:
        print("\n（预览模式，加 --apply 落盘）")
        return 0

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ledger.replace(".xlsx", "_bak_grounds_tail_%s.xlsx" % stamp)
    shutil.copy2(ledger, bak)
    for row, _o, new in plan:
        ws.cell(row, 6).value = new
    wb.save(ledger)
    wb.close()
    print("\n✓ 已写入 %d 行｜备份：%s" % (len(plan), bak))
    return 0


if __name__ == "__main__":
    sys.exit(main())
