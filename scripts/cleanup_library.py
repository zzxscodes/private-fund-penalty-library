#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""cleanup_library.py —— 案例库「终态整洁」收尾（2026-10-28 用户要求 #1）

需求原话：
  「案例库技能执行后终态时案例库本身内部应该要是干净的，不相干的要清理，
    需要保留的才保留在库中或其他位置。」

设计原则（与全局约定一致，勿改）：
  1. **只移不删**。任何被判定为「不相干」的东西一律 *move* 到库外留痕区，
     绝不 unlink/rmtree —— 误判可 100% 回退。
  2. **白名单制**。除了白名单，库内一律不允许留存；新增白名单必须写在
     ``ROOT_KEEP_FILES`` / ``ROOT_KEEP_DIR_RE`` 里，不要在别处散落特判。
  3. **默认 dry-run**。不显式加 ``--apply`` 就只打印计划，不动盘。
     由 ``run_update.py`` 的 ``cleanup`` 阶段自动带 ``--apply`` 调用。
  4. **留痕区在库外**：``<库父目录>/_案例库维护留痕/<库名>/<YYYY-MM-DD>/``。
     绝不能是库内子目录 —— 否则下轮的清理又把它当杂物搬走（无限套娃）。

白名单（库根目录）：
  · 目录：``01_``~``06_`` 六个分类目录（含未来新增的 ``07_``…）
  · 文件：04_综合索引台账【总表】.xlsx（正本）/ .html（网页检索台）
          / 不收录清单.md / 更新记录.md / 校验人工复核豁免清单.md

分类目录内部（递归）：
  · 只放行 ``.pdf``（决定书正文）与目录（分类层级）
  · 其余一律视为杂物移出（``.json`` / ``.bak*`` / ``.tmp`` / ``__pycache__`` /
    ``Thumbs.db`` / ``.DS_Store`` / 误放的 md、log、xlsx…）
  · 注意：``05_重点典型案例`` 里放的是 PDF 副本，同样只放行 .pdf。

用法：
  python cleanup_library.py --root "<案例库根>"            # 预演
  python cleanup_library.py --root "<案例库根>" --apply    # 执行
  python cleanup_library.py --root "<案例库根>" --apply --archive "<自定义留痕区>"
"""

import os
import re
import sys
import shutil
import datetime
import argparse

ROOT_KEEP_DIR_RE = re.compile(r"^\d{2}_")           # 01_ … 06_ 分类目录
ROOT_KEEP_FILES = {
    "04_综合索引台账【总表】.xlsx",
    # 台账的网页形态（2026-11-04 用户要求同放库根、同名）。
    # .xlsm 宏版已于同日停用（Office 未激活，宏在用户端正规路径下不执行），不再产出。
    "04_综合索引台账【总表】.html",
    "不收录清单.md",
    "更新记录.md",
    "校验人工复核豁免清单.md",
}
# 分类目录内允许留存的后缀：
#   01~05 只放行 .pdf（决定书正文只有 PDF，其余一律杂物）
#   06 法规目录放行 .pdf/.doc/.docx —— 法规原文是 docx，
#     2026-10-28 首版只写 .pdf 差点把 40 部法规全当垃圾搬走（预演才发现）。
INNER_KEEP_SUFFIX = {".pdf"}
INNER_KEEP_SUFFIX_LAW = {".pdf", ".doc", ".docx"}
LAW_DIR_RE = re.compile(r"^06_")
# 无论在哪都要清掉的系统/缓存垃圾
JUNK_NAMES = {
    "__pycache__", ".DS_Store", "Thumbs.db", "thumbs.db",
    ".git", ".ipynb_checkpoints",
}


def _move(src, dst_dir, moved, dry):
    """移动单文件/单目录到留痕区；同名的加时间戳后缀，绝不覆盖、绝不删除。"""
    os.makedirs(dst_dir, exist_ok=True)
    base = os.path.basename(src.rstrip("\\/")) or "unnamed"
    dst = os.path.join(dst_dir, base)
    if os.path.exists(dst):
        stem, ext = os.path.splitext(base)
        dst = os.path.join(dst_dir, "%s_%s%s" % (
            stem, datetime.datetime.now().strftime("%H%M%S"), ext))
        if os.path.exists(dst):
            return False
    rel = os.path.relpath(src, dst_dir)
    print("    %s  %s" % ("[移出]" if not dry else "[待移]", rel))
    if not dry:
        try:
            shutil.move(src, dst)
            moved.append((src, dst))
            return True
        except Exception as e:
            print("      !! 移动失败（保留在原地）：%s" % e)
            return False
    moved.append((src, dst))
    return True


def cleanup(root, archive, dry=True):
    root = os.path.abspath(root)
    if not os.path.isdir(root):
        raise SystemExit("案例库根不存在：%s" % root)
    lib = os.path.basename(root.rstrip("\\/"))
    stamp = datetime.date.today().isoformat()
    archive = os.path.abspath(archive) if archive else os.path.join(
        os.path.dirname(root), "_案例库维护留痕", lib, stamp)
    # 留痕区绝不能在库内（否则自我套娃）
    if os.path.commonpath([os.path.abspath(archive), root]) == root:
        raise SystemExit("留痕区不能位于案例库内部：%s" % archive)

    moved, kept = [], 0
    print("案例库根: %s" % root)
    print("留痕区  : %s" % archive)
    print("模式    : %s\n" % ("预演（未动盘）" if dry else "执行"))

    print("== 根目录 ==")
    for nm in sorted(os.listdir(root)):
        p = os.path.join(root, nm)
        if nm in JUNK_NAMES:
            _move(p, archive, moved, dry)
            continue
        if os.path.isdir(p):
            if ROOT_KEEP_DIR_RE.match(nm):
                kept += 1
                print("    [保留] %s/" % nm)
            else:
                _move(p, os.path.join(archive, "root"), moved, dry)
            continue
        if nm in ROOT_KEEP_FILES:
            kept += 1
            print("    [保留] %s" % nm)
            continue
        # 台账的历史备份（04_…xlsx.bak_*）与清单备份同样移出，库内只留当前版本
        _move(p, os.path.join(archive, "root"), moved, dry)

    print("\n== 分类目录（递归，只放行 .pdf 与目录）==")
    for nm in sorted(os.listdir(root)):
        p = os.path.join(root, nm)
        if not (os.path.isdir(p) and ROOT_KEEP_DIR_RE.match(nm)):
            continue
        allow = INNER_KEEP_SUFFIX_LAW if LAW_DIR_RE.match(nm) else INNER_KEEP_SUFFIX
        for dp, dns, fns in os.walk(p):
            dns[:] = [d for d in dns if d not in JUNK_NAMES]
            inner = os.path.relpath(dp, root)
            for fn in sorted(fns):
                fp = os.path.join(dp, fn)
                if fn in JUNK_NAMES:
                    _move(fp, os.path.join(archive, inner), moved, dry)
                    continue
                ext = os.path.splitext(fn)[1].lower()
                if ext in allow:
                    kept += 1
                    continue
                _move(fp, os.path.join(archive, inner), moved, dry)
            for d in list(dns):
                if d in JUNK_NAMES:
                    _move(os.path.join(dp, d), os.path.join(archive, inner), moved, dry)

    print("\n---- 汇总 ----")
    print("保留 %d 项；%s %d 项 → %s" % (kept, "待移出" if dry else "已移出", len(moved), archive))
    return len(moved)


def main():
    ap = argparse.ArgumentParser(description="案例库终态整洁收尾（只移不删）")
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--archive", default="", help="留痕区（默认 <库父>/_案例库维护留痕/<库名>/<日期>）")
    ap.add_argument("--apply", action="store_true", help="真正执行移动（默认只预演）")
    a = ap.parse_args()
    n = cleanup(a.root, a.archive, dry=not a.apply)
    if not a.apply and n:
        print("\n（这是预演。确认无误后加 --apply 执行；所有条目均为移动，不会删除。）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
