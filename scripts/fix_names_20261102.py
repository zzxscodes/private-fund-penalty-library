#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""fix_names_20261102.py —— 交易所分类收尾同批：01 中基协「文件名主体/复核」语义修正

一次性的数据修正（不通用，故不放进 SKILL 流程；逻辑与证据见 更新记录.md 第十一节·五）。
根因：下载器按「台账行主体 + 同名序号」命名，遇到
  ① 机构决定书与高管决定书同页下载 → 机构那份被冠上高管名（浙江贯邦 / 陕西望和）；
  ② 原决定与复核决定同名 → 只按（2）加日期，没看文书类型（上海通赛 / 汕头宜华 / 上海世联行）；
  ③ 机构仅一份决定却带了日期后缀（上海世联行）。
判定依据：逐份 OCR 首页标题（纪律处分决定书 / 纪律处分复核决定书）+ 文号。

用法：python fix_names_20261102.py --root <库根> [--apply]
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ledger_bak import backup_ledger  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

# (行号, 旧名, 新名)  —— 行号=Excel 1-based
RENAME = [
    # 行149 浙江贯邦：111号=机构，113号=胡渡洋
    (149, "胡渡洋-浙江贯邦私募基金管理有限公司.pdf", "浙江贯邦私募基金管理有限公司.pdf"),
    (149, "胡渡洋-浙江贯邦私募基金管理有限公司（2026-04-13）.pdf",
          "胡渡洋-浙江贯邦私募基金管理有限公司.pdf"),
    # 行250 陕西望和：9号=机构，10号=孙小明
    (250, "孙小明-陕西望和私募基金管理有限公司.pdf", "陕西望和私募基金管理有限公司.pdf"),
    (250, "孙小明-陕西望和私募基金管理有限公司（2026-01-13）.pdf",
          "孙小明-陕西望和私募基金管理有限公司.pdf"),
    # 行226 上海通赛·李承恩：49号原决定 + 复核12号
    (226, "李承恩-上海通赛投资有限公司.pdf", "李承恩-上海通赛投资有限公司（原决定）.pdf"),
    (226, "李承恩-上海通赛投资有限公司（2026-05-26）.pdf",
          "李承恩-上海通赛投资有限公司（复核决定）.pdf"),
    # 行232 汕头宜华·肖伟昊：无后缀那份才是复核10号，44号是原决定
    (232, "肖伟昊-汕头宜华投资有限公司.pdf", "肖伟昊-汕头宜华投资有限公司（复核决定）.pdf"),
    (232, "肖伟昊-汕头宜华投资有限公司（2026-02-11）.pdf",
          "肖伟昊-汕头宜华投资有限公司（原决定）.pdf"),
    # 行207 上海世联行：50号机构（唯一，去日期）+ 51号原决定 + 复核19号
    (207, "上海世联行股权投资管理有限公司（2026-03-02）.pdf", "上海世联行股权投资管理有限公司.pdf"),
    (207, "袁鸿昌-上海世联行股权投资管理有限公司.pdf",
          "袁鸿昌-上海世联行股权投资管理有限公司（原决定）.pdf"),
    (207, "袁鸿昌-上海世联行股权投资管理有限公司（2026-07-02）.pdf",
          "袁鸿昌-上海世联行股权投资管理有限公司（复核决定）.pdf"),
]

# 主体列碎片清洗：行号 → 新值
SUBJ_FIX = {232: "汕头宜华投资有限公司；肖伟昊、袁榕、林浩"}


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
    assert str(ws.cell(1, 12).value) == "案例文件"

    # 目录索引（01/02/03/05 全域按 basename 找）
    where = {}
    for dp, _dn, fns in os.walk(root):
        if os.path.basename(dp)[:2] not in ("01", "02", "03", "05"):
            continue
        for f in fns:
            if f.lower().endswith(".pdf"):
                where.setdefault(f, []).append(os.path.join(dp, f))

    print("== 改名计划 ==")
    plan = []
    for row, old, new in RENAME:
        cellv = str(ws.cell(row, 12).value or "")
        if old not in cellv:
            print("  ! 行%d 台账 K 列不含旧名，跳过：%s" % (row, old))
            continue
        paths = where.get(old, [])
        if not paths:
            # 断点续跑：上次已改名、台账没跟上（或文件被占用重试成功但台账未写）
            if where.get(new):
                print("  · 行%d 磁盘已是新名，仅补台账：%s" % (row, new))
                plan.append((row, old, new, []))
            else:
                print("  ! 磁盘找不到（新旧名都不存在）：%s" % old)
            continue
        plan.append((row, old, new, paths))
        for p in paths:
            print("  行%-4d %s\n        %s\n     → %s" % (row, os.path.basename(os.path.dirname(p)), old, new))

    print("\n== 主体列清洗 ==")
    for row, val in SUBJ_FIX.items():
        print("  行%-4d %s\n     → %s" % (row, ws.cell(row, 4).value, val))

    if not a.apply:
        print("\n[干跑] 未改动；确认后加 --apply")
        return

    backup_ledger(root, ledger, "reviewname")

    # 两步改名：先全部改成临时名，再改成终名，避免 A→B 与 B→C 互相覆盖
    # 2026-11-02 实测踩坑：全库 OCR 扫描进程正打开着 PDF，os.rename 抛 WinError 32，
    # 且已改名的 8 个文件停在 .tmp_rename 中间态 → 必须「重试 + 失败不中断 + 记失败清单」，
    # 并且回滚时要能识别 .tmp_rename（见回滚说明）。
    import time

    def ren(a, b, tries=4):
        for i in range(tries):
            try:
                os.rename(a, b)
                return True
            except PermissionError:
                time.sleep(1.2 * (i + 1))
            except OSError as e:
                print("  ! OSError %s：%s" % (e, os.path.basename(a)))
                return False
        return False

    good, bad = [], []
    for item in plan:
        row, old, new, paths = item
        if not paths:                      # 仅补台账
            good.append(item)
            continue
        tpaths = [p + ".tmp_rename" for p in paths]
        if all(ren(p, t) for p, t in zip(paths, tpaths)):
            good.append(item)
        else:
            bad.append((row, old, "临时名阶段失败（文件被占用？）"))
            # 尽力回滚已改的
            for p, t in zip(paths, tpaths):
                if os.path.exists(t) and not os.path.exists(p):
                    ren(t, p)
    done = 0
    for row, old, new, paths in good:
        for p in paths:
            if ren(p + ".tmp_rename", os.path.join(os.path.dirname(p), new)):
                done += 1
    print("改名 %d 个文件" % done)
    if bad:
        print("失败 %d 项（台账对应行未改）：%s" % (len(bad), bad))

    # 台账 K 列：**单遍正则**替换。2026-11-02 踩坑：先写占位符 \x00 再替换会被
    # openpyxl 的 IllegalCharacterError 直接拒绝（cell.value 赋值即校验），
    # 结果文件已改名、台账没跟上 → A1 孤儿激增。单遍替换既避开暂存字符，也天然防「新名被再替换」。
    ok_items = [(row, old, new) for row, old, new, _p in good]
    by_row = {}
    for row, old, new in ok_items:
        by_row.setdefault(row, []).append((old, new))
    for row, pairs in by_row.items():
        cell = ws.cell(row, 12)
        v = str(cell.value or "")
        pat = re.compile("|".join(re.escape(o) for o, _n in
                                  sorted(pairs, key=lambda x: -len(x[0]))))
        mp = dict(pairs)
        cell.value = pat.sub(lambda m: mp[m.group(0)], v)
    for row, val in SUBJ_FIX.items():
        ws.cell(row, 4).value = val
    wb.save(ledger)
    print("台账已保存（K 列 %d 处、D 列 %d 处）" % (len(ok_items), len(SUBJ_FIX)))


if __name__ == "__main__":
    main()
