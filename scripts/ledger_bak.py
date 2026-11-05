#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""ledger_bak.py —— 台账备份路径的唯一真源（2026-11-02）

为什么必须有这个：**库根只有 9 项白名单**（01~06 目录 + 台账 + 3 个 .md）。
历史脚本习惯 `ledger.replace(".xlsx", "_bak_xxx.xlsx")`，直接写在库根 ——
一轮维护跑下来库里多出 8 个备份 xlsx，白名单立刻破功（用户会一眼看到「怎么多了一堆表」）。

本模块统一把备份放到**库外**：`<库父目录>/_案例库维护留痕/台账备份/<台账名>_bak_<tag>_<时间>.xlsx`。
路径全部从 `--root` 派生，不硬编码任何本机路径（可迁移性自检会把硬编码路径当错误）。

用法：
    from ledger_bak import backup_ledger
    backup_ledger(root, ledger_path, "before_rename")
"""
import datetime as dt
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)


def bak_dir(root):
    """库外备份目录（不存在则建）。"""
    d = os.path.join(os.path.dirname(os.path.abspath(root)), "_案例库维护留痕", "台账备份")
    os.makedirs(d, exist_ok=True)
    return d


def backup_ledger(root, ledger_path, tag="manual"):
    """复制台账到库外备份目录，返回备份文件绝对路径。"""
    stem = os.path.splitext(os.path.basename(ledger_path))[0].replace("【", "").replace("】", "")
    name = "%s_bak_%s_%s.xlsx" % (stem, tag, dt.datetime.now().strftime("%Y%m%d_%H%M%S"))
    dst = os.path.join(bak_dir(root), name)
    shutil.copy2(ledger_path, dst)
    print("台账备份（库外）：%s" % dst)
    return dst


if __name__ == "__main__":
    print(bak_dir(sys.argv[1] if len(sys.argv) > 1 else "."))
