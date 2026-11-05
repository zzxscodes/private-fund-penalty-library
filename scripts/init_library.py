# -*- coding: utf-8 -*-
"""初始化私募监管处罚案例库目录结构、空白台账与更新记录。

用法:
    python init_library.py --root "<工作空间路径>"

在 <工作空间路径> 下创建:
    私募监管处罚案例库（中基协+证监局+交易所）/
        01_中基协自律处分.../01~06 分类目录
        02_各地证监局行政监管.../01~06 分类目录
        03_交易所监管措施.../01~03+06 分类目录
        04_综合索引台账【总表】.xlsx
        05_重点典型案例/4 个子目录
        06_以上各类所有案例的配套法规依据/01~03 子目录
        更新记录.md
注意：「原始决定书PDF」不是目录，只是说明分类目录中留存的是原始决定书 PDF。
已存在的目录与文件不会被覆盖。
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ledger_style as LS  # noqa: E402

ROOT_NAME = "私募监管处罚案例库（中基协+证监局+交易所）"
LEDGER_NAME = LS.LEDGER_NAME

CAT_AMAC = "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）"
CAT_CSRC = "02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）"
CAT_EXCH = "03_交易所监管措施（沪深交易所、北交所，私募产品交易、异常交易纪律）"

SUB_COMMON = [
    "01_募集与适当性违规",
    "02_登记备案、信息报送、信披违规",
    "03_投资运作、关联交易、利益输送",
    "04_内控治理、人员资质、风控缺位",
    "05_基金财产挪用、混同、保本保收益",
]

STRUCTURE = {
    CAT_AMAC: SUB_COMMON + ["06_其他自律违规"],
    CAT_CSRC: SUB_COMMON + ["06_其他行政违规"],
    CAT_EXCH: [
        "01_异常交易、对倒、价格操纵",
        "02_产品交易风控、持仓限制违规",
        "03_信息报送、交易行为自律监管",
        # 2026-11-02 用户二次裁定：03 的「其他」保留，但必须用 03 自己的序号与口径
        # ——`04_其他自律监管`，**不是**六分法的 `06_其他行政违规`（06 是证监局序号、
        # 「行政」是行政监管口径；实测曾出现 03/06 目录，看着像串了分类体系）。
        "04_其他自律监管",
    ],
    "05_重点典型案例": [
        "玖瀛资产+腾创投资案",
        "关联交易典型案例",
        "风控负责人履职违规案例",
        "募集宣传自媒体违规案例",
    ],
    "06_以上各类所有案例的配套法规依据": [
        "01_法律与行政法规",
        "02_证监会部门规章与规范性文件",
        "03_中基协自律规则",
        # 2026-09-29 补：此前缺交易所规则类，导致深交所案例引用的
        # 《创业板股票上市规则》等上市/交易规则无处安放（御硕、汇垠澳丰、远致瑞信三条均引用）
        "04_交易所业务规则（沪深北交易所）",
    ],
}


def create_dirs(root: Path) -> int:
    count = 0
    for top, subs in STRUCTURE.items():
        top_dir = root / top
        if not top_dir.exists():
            top_dir.mkdir(parents=True)
            count += 1
        for sub in subs:
            d = top_dir / sub
            if not d.exists():
                d.mkdir(parents=True)
                count += 1
    return count


def create_ledger(root: Path) -> bool:
    """按 ledger_style.py 的冻结规格新建台账；已存在则原样保留（不重建、不改排版）。"""
    ledger = root / LEDGER_NAME
    if ledger.exists():
        return False
    try:
        from openpyxl import Workbook  # noqa: F401
    except ImportError:
        print("ERROR: 缺少 openpyxl，请在托管 Python 隔离 venv 中安装: pip install openpyxl", file=sys.stderr)
        sys.exit(2)
    wb = LS.new_workbook()
    LS.ensure_law_sheet(wb)
    wb.save(ledger)
    return True


def create_log(root: Path) -> bool:
    log = root / "更新记录.md"
    if log.exists():
        return False
    log.write_text(
        "# 私募监管处罚案例库 更新记录\n\n"
        "每次执行归集后，在此追加：执行日期、各信源核查情况、新增条目清单。\n",
        encoding="utf-8",
    )
    return True


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="工作空间路径（案例库根文件夹将创建在其下）")
    args = ap.parse_args()
    root = Path(args.root) / ROOT_NAME
    root.mkdir(parents=True, exist_ok=True)
    n_dirs = create_dirs(root)
    ledger_ok = create_ledger(root)
    log_ok = create_log(root)
    print(f"案例库根目录: {root}")
    print(f"新建目录: {n_dirs} 个")
    print(f"台账: {'已创建 ' + LEDGER_NAME if ledger_ok else '已存在，跳过'}")
    print(f"更新记录: {'已创建 更新记录.md' if log_ok else '已存在，跳过'}")


if __name__ == "__main__":
    main()
