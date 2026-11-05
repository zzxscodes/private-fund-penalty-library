# -*- coding: utf-8 -*-
"""ledger_style 内容层格式规则单测（2026-11-02 用户裁定口径）。

守护四件事，任何一处回退都必须 FAIL：
  1) 处罚措施：每主体一行、多高管相邻、组间按原文先后；
  2) 复核决定：空一行另起；
  3) 核心事由：一点一换行、序号统一「一、」；
  4) 反例：23.74 / 第一条 / 第3.3节 / 第（八）项 绝不能被当序号切开；
     枚举收尾「A、B等相关事项。」绝不能被当残头删掉。

运行：python utest_ledger_style.py
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import ledger_style as LS  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

_n_pass = _n_fail = 0


def ck(name, cond, got=""):
    global _n_pass, _n_fail
    if cond:
        _n_pass += 1
        print(f"  [PASS] {name}")
    else:
        _n_fail += 1
        print(f"  [FAIL] {name}\n         got: {got!r}")


def eq(name, got, exp):
    ck(name, got == exp, got)


# ---------------------------------------------------------------- 1 处罚措施
print("=== 处罚措施：按主体分行 / 多高管相邻 / 组间按原文先后 ===")
m1 = ("高管（李宝勇）：暂停受理私募基金产品备案六个月；"
      "管理人：公开谴责，暂停受理私募基金产品备案十二个月；"
      "高管（赵杨博）：公开谴责")
eq("用户样例：高管组相邻、管理人后置",
   LS.normalize_measures_text(m1),
   "高管（李宝勇）：暂停受理私募基金产品备案六个月；\n"
   "高管（赵杨博）：公开谴责\n"
   "管理人：公开谴责，暂停受理私募基金产品备案十二个月；")

m2 = "管理人：警告；高管（甲）：警告；高管（乙）：罚款"
eq("管理人原文在前则保持在前",
   LS.normalize_measures_text(m2),
   "管理人：警告；\n高管（甲）：警告；\n高管（乙）：罚款")

eq("列表入口与文本入口同口径",
   LS.format_measures_list(["管理人：警告", "高管（甲）：警告"]),
   "管理人：警告\n高管（甲）：警告")

eq("幂等：重复规范化不变",
   LS.normalize_measures_text(LS.normalize_measures_text(m1)),
   LS.normalize_measures_text(m1))

# ---------------------------------------------------------------- 2 复核另起
print("=== 复核决定：空一行另起 ===")
m3 = "管理人：暂停受理；高管（张雅薇）：警告 【复核决定·2026-07-02】维持原决定"
eq("复核标注前插入空行",
   LS.normalize_measures_text(m3).split("\n")[-2:],
   ["", "【复核决定·2026-07-02】维持原决定"])
ck("复核标注只出现一次空行",
   "\n\n\n" not in LS.normalize_measures_text(m3),
   LS.normalize_measures_text(m3))
eq("复核幂等",
   LS.normalize_measures_text(LS.normalize_measures_text(m3)),
   LS.normalize_measures_text(m3))

# ---------------------------------------------------------------- 3 事由列表化
print("=== 核心事由：一点一行 + 序号统一「一、」 ===")
g1 = ("一、 部分私募基金未托管且未在基金合同中明确约定保障私募基金财产安全的制度措施 。 "
      "二、 未在中国证券投资基金业协会系统及时填报并定期更新私募基金管理人子公司及员工人数的有关信息 。 "
      "三、与关联私募基金管理人混同办公。")
eq("用户样例①：句号分隔的点各占一行",
   LS.normalize_grounds_items(g1),
   "一、 部分私募基金未托管且未在基金合同中明确约定保障私募基金财产安全的制度措施 。\n"
   "二、 未在中国证券投资基金业协会系统及时填报并定期更新私募基金管理人子公司及员工人数的有关信息 。\n"
   "三、与关联私募基金管理人混同办公。")

g2 = ("一是专职人员少于 5 人，未按规定及时填报；二是合规风控负责人兼职；"
      "三是挪用基金财产；四是未按规定开展投资者适当性管理工作。")
eq("用户样例②：一是→一、且每条一行",
   LS.normalize_grounds_items(g2),
   "一、专职人员少于 5 人，未按规定及时填报；\n"
   "二、合规风控负责人兼职；\n"
   "三、挪用基金财产；\n"
   "四、未按规定开展投资者适当性管理工作。")

eq("（一）式归一", LS.normalize_grounds_items("（一）未披露； （二）未勤勉。"),
   "一、未披露；\n二、未勤勉。")
eq("1. 式归一", LS.normalize_grounds_items("1. 无办公场所; 2. 未备案;"),
   "一、 无办公场所;\n二、 未备案;")
eq("逗号后的序号也切（实测 6 例）",
   LS.normalize_grounds_items("存在以下问题，一是披露不及时，二是披露不准确。"),
   "存在以下问题\n一、披露不及时\n二、披露不准确。")
eq("冒号后的序号也切",
   LS.normalize_grounds_items("存在以下情形: 一是管理未备案基金"),
   "存在以下情形:\n一、管理未备案基金")
eq("序号后重复逗号被吃掉（「一是，2016年」→「一、2016年」）",
   LS.normalize_grounds_items("一、，法泷投资在协会登记。\n"), "一、法泷投资在协会登记。")
eq("无标点粘连：枚举上下文中「三是」前断行",
   LS.normalize_grounds_items("二、报送信息存在虚假记载三是兼营其他业务。").split("\n"),
   ["二、报送信息存在虚假记载", "三、兼营其他业务。"])
eq("列表化幂等", LS.normalize_grounds_items(LS.normalize_grounds_items(g2)),
   LS.normalize_grounds_items(g2))

# ---------------------------------------------------------------- 4 反例
print("=== 反例：正文数字/条款不得当序号切开 ===")
for bad in ("23.74 元的份额对应收益",
            "不符合《私募投资基金监督管理条例》第一条的规定",
            "详见第3.3节说明",
            "违反第九条第一款第（八）项"):
    eq(f"不切：{bad[:14]}…", LS.normalize_grounds_items(bad), bad)
eq("单点无序号不动", LS.normalize_grounds_items("未按规定开展投资者适当性管理工作。"),
   "未按规定开展投资者适当性管理工作。")

print("=== 反例：残头删除边界 ===")
eq("删残头 A 类", LS.clean_grounds_tail("公司未履行谨慎勤勉义务。上述行为"),
   "公司未履行谨慎勤勉义务。")
eq("删残头 B 类（短主语连带）", LS.clean_grounds_tail("公司未履行义务。武汉金瑞麟上述行为"),
   "公司未履行义务。")
eq("保留枚举收尾「等X。」", LS.clean_grounds_tail("未落实A、B等相关事项。"),
   "未落实A、B等相关事项。")
eq("谓语收尾不删", LS.clean_grounds_tail("公司存在上述行为"), "公司存在上述行为")

print("=== 页码噪声（PDF 分页切词）===")
eq("页码夹在词中复原", LS.strip_page_noise("列为重点监控账 -2- 户、暂停交易"),
   "列为重点监控账户、暂停交易")
eq("页码夹在条款中复原", LS.strip_page_noise("第十三条第一 - 2 - 款的规定"),
   "第十三条第一款的规定")
eq("不误伤日期", LS.strip_page_noise("2026-2-24 作出决定"), "2026-2-24 作出决定")
eq("页码清洗幂等", LS.strip_page_noise(LS.strip_page_noise("账 -2- 户")), "账户")

print("-" * 60)
print(f"通过 {_n_pass} / 失败 {_n_fail}")
raise SystemExit(0 if _n_fail == 0 else 1)
