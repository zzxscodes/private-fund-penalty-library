#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""subject_clean 的单测：碎片词形态规则（2026-11-02 新增）

背景：C8 曾只做精确匹配，于是「离职后」「的安排」这类「词根+虚词」形态穿透，
汕头宜华一行主体被写成「肖伟昊、袁榕、离职后、的安排、林浩」。
本测锁定：① 形态碎片必须识别；② 真实姓氏不得误伤。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import subject_clean as SC  # noqa: E402

ok = fail = 0


def chk(cond, msg, detail=""):
    global ok, fail
    if cond:
        ok += 1
        print("  [PASS] %s" % msg)
    else:
        fail += 1
        print("  [FAIL] %s %s" % (msg, detail))


# ① 形态碎片：虚词开头
for t in ("离职后", "的安排", "的请求", "在任职", "为他人", "与其", "并将", "以上", "其中"):
    chk(SC.is_noise_token(t), "虚词/精确碎片：%s" % t, t)

# ② 形态碎片：噪声词根 + 1~2 字
for t in ("任职期间", "离职以后", "挂名期间", "签章页"):
    chk(SC.is_noise_token(t), "词根+后缀碎片：%s" % t, t)

# ③ 反例：真实姓名（含常见姓氏）绝不能判碎片
for t in ("于某", "向某", "时某", "成某", "和某", "肖伟昊", "袁榕", "林浩", "王券期",
          "胡逸之", "贲金锋", "欧阳修", "阿依古丽·买买提"):
    chk(not SC.is_noise_token(t), "真实姓名不误伤：%s" % t, t)

# ④ 机构名不得判碎片
for t in ("深圳市黑岩中盛资产管理有限公司", "天问时代私募基金管理（海南）有限公司"):
    chk(not SC.is_noise_token(t), "机构名不误伤：%s" % t, t)

print("-" * 60)
print("通过 %d / 失败 %d" % (ok, fail))
sys.exit(1 if fail else 0)
