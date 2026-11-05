# -*- coding: utf-8 -*-
"""winlib 单元测试 + 「窗口口径唯一事实源」结构断言。

覆盖：正常 / 边界（含当日、恰在端点）/ 反例（旧的双实现公式必须绝迹）。
2026-10-29 事故：retire 与 verify 各算一套窗口起点，差 1 天，把落款恰在
边界的案例判成超窗移出，而校验报告同时显示它在窗内。
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import winlib

FAILS = []


def chk(cond, msg, got=""):
    if cond:
        print("  OK   " + msg)
    else:
        print("  FAIL " + msg + (f"  → 实得 {got!r}" if got != "" else ""))
        FAILS.append(msg)


print("=== 1) window() 正常用例 ===")
chk(winlib.window("2026-10-29", 365) == ("2025-10-30", "2026-10-29"),
    "365 天含当日 → 起点 2025-10-30", winlib.window("2026-10-29", 365))
chk(winlib.window("2026-10-29", 1) == ("2026-10-29", "2026-10-29"),
    "days=1 → 只有当日", winlib.window("2026-10-29", 1))

print("\n=== 2) 边界：inclusive 天数恒等于 --days（含闰年 2 月）===")
import datetime as dt
for end, days in (("2026-10-29", 365), ("2028-03-01", 365), ("2028-03-01", 366),
                  ("2026-01-01", 30), ("2027-12-31", 365)):
    s, e = winlib.window(end, days)
    n = (dt.date.fromisoformat(e) - dt.date.fromisoformat(s)).days + 1
    chk(n == days, f"{end} days={days} → 实含 {n} 天", n)

print("\n=== 3) 边界：端点判定 ===")
s, e = winlib.window("2026-10-29", 365)          # 2025-10-30 ~ 2026-10-29
chk(winlib.in_window("2025-10-30", s, e), "起点当日算在窗内")
chk(winlib.in_window("2026-10-29", s, e), "终点当日算在窗内")
chk(not winlib.in_window("2025-10-29", s, e), "起点前一天不算（事故日：曾被判在窗内）")
chk(not winlib.in_window("2026-10-30", s, e), "终点后一天不算")

print("\n=== 4) 反例：非法/空日期一律 False ===")
for bad in ("", None, "待补", "2026-1-1", "20261029", "2026-10-29 00:00:00"):
    chk(winlib.in_window(bad, s, e) is False, f"非法日期 {bad!r} → False",
        winlib.in_window(bad, s, e))

print("\n=== 5) 结构断言：窗口口径不得再出现第二套实现 ===")
here = os.path.dirname(os.path.abspath(__file__))
BAD = [
    (r"timedelta\(days=args\.days\s*-\s*1\)", "run_update 旧公式 end-(days-1)"),
    (r"date\(today\.year\s*-\s*1", "verify 旧公式 date(year-1)"),
]
for fn in ("run_update.py", "verify_library.py"):
    src = open(os.path.join(here, fn), encoding="utf-8").read()
    chk("import winlib" in src, f"{fn} 必须 import winlib")
    for pat, label in BAD:
        chk(re.search(pat, src) is None, f"{fn} 不得残留{label}")

print("\n" + ("全部通过（%d 项失败）" % len(FAILS) if not FAILS else "存在 %d 项失败" % len(FAILS)))
sys.exit(1 if FAILS else 0)
