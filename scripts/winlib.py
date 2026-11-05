# -*- coding: utf-8 -*-
"""库窗口的**唯一事实源**（2026-10-29 立）。

背景（真实事故）：窗口起点曾在两处各算一套——
  * `run_update.py`：`end - (days-1)` → days=365 时起点 2025-10-30
  * `verify_library.py`：`date(today.year-1, ...)` → 起点 2025-10-29
差 1 天，于是 verify 说「在窗内、合格」的行，下一轮 retire 判「超窗」直接移出库。
实测代价：2026-10-29 一轮 retire 把「浙江洋哲晨资产管理有限公司；陈阳」（落款 2025-10-29）
当超窗移出，而同一秒的校验报告仍认为它在窗口内 —— 两个模块互相打脸，用户看到的是
「案例莫名消失」。同一业务概念只允许有一个实现，故抽出本模块。

口径（与 SKILL 13.1 一致）：窗口 = 含当日共 `days` 个自然日，
即 `[end - (days-1), end]`。默认 365 天。
"""
import datetime as _dt
import re

_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")


def window(end=None, days=365):
    """返回 (start, end) 两个 ISO 日期字符串。

    - `end`：窗口终点，None=今天；支持 date / 'YYYY-MM-DD'。
    - `days`：含当日共多少天，默认 365。`days=1` → 只有 end 当天。
    """
    if end is None:
        e = _dt.date.today()
    elif isinstance(end, _dt.date):
        e = end
    else:
        e = _dt.date.fromisoformat(str(end)[:10])
    n = max(1, int(days))
    s = e - _dt.timedelta(days=n - 1)
    return s.isoformat(), e.isoformat()


def in_window(day, start, end):
    """某日期是否落在 [start, end]（含端点）。

    只接受 `date`/`datetime` 对象或严格的 `YYYY-MM-DD` 字符串；
    其余（空、'待补'、'2026-1-1'、'20261029'、带时分的字符串）一律 False。
    校验器必须严格：宽进会把格式错误静默当成合法日期，窗口判定就失效了。
    """
    if isinstance(day, (_dt.datetime, _dt.date)):
        d = day.isoformat()[:10]
    else:
        d = str(day or "").strip()
    if not _DATE_RE.fullmatch(d):
        return False
    return start <= d <= end
