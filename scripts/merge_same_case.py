#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""merge_same_case.py —— 同一案例被拆成多行时并回一行（2026-11-02）

## 为什么需要这个脚本

台账 v2 的判重键是 **(来源, 管理人, 处罚时间)**，语义是「一家机构同一天的一个处罚事件 = 一行，
机构与高管合并」。但落地时键被写成了「D 列整串（机构+人员）」——于是同一案的
「机构决定书」与「某高管决定书」只要人员名单不同，就被当成两个案例各占一行。

真正把它暴露出来的是**主体名污染**：中基协对个人的决定书正文写
「当事人：刘凯歌……实际负责浙江优策投资管理有限公司」，抽取时把角色短语当机构名，
于是三个案子各被劈成两行：

    中基协字〔2026〕75号  → 处分〔2026〕163号（机构+王券期+郭玉兰） / 165号（侯荐珉）
    中基协字〔2025〕433号 → 处分〔2026〕62号（机构+3人）          / 65号（张蕾）
    中基协字〔2025〕356号 → 处分〔2025〕333号（机构+4人）          / 335号（刘凯歌）

三对的**事先告知书文号完全相同**，铁证是同一个案子（已逐份 OCR 首页核对）。

## 规则

按 (来源, 机构名[剥角色前缀], 落款日) 分组，≥2 行的组判为「同一案被拆行」：
  · 保留文件数最多的一行（并列取行号小的）
  · 人员并进 D 列、措施逐条并进 E 列（已有的「高管（X）」不重复追加）
  · K 列文件、M 列链接并集
  · F/G/H/I（事由/证据/依据/申辩）以保留行为准；保留行为空/「无」时取被合并行
  · 违规类型、存放路径不一致则**不合并**，报冲突交人工（跨类型目录强行合并会破 C2）
最后按 序号 重排 1..N。

用法：
  python scripts/merge_same_case.py --root <库根>            # dry-run
  python scripts/merge_same_case.py --root <库根> --apply
"""
import argparse
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

from subject_clean import strip_role_prefix  # noqa: E402

# 列下标（0-based）
C_SEQ, C_DECIDED, C_SUBJECT, C_MEAS, C_GROUNDS, C_EVID, C_LEGAL, C_DEFENSE = 0, 1, 3, 4, 5, 6, 7, 8
C_CAT, C_SRC, C_FILES, C_PATHS, C_URLS, C_TYP = 9, 10, 11, 12, 13, 14
EMPTY = ("", "无", "未列举证据材料", "不明确", "正文未载明处罚依据")


def split_multi(v):
    return [p.strip() for p in re.split(r"[；;]", str(v or "")) if p.strip()]


def join_multi(xs):
    out, seen = [], set()
    for x in xs:
        x = str(x or "").strip()
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return "；".join(out)


def org_of(subject):
    segs = split_multi(subject)
    return strip_role_prefix(segs[0]) if segs else ""


def persons_of(subject):
    segs = split_multi(subject)
    out = []
    for s in segs[1:]:
        out += [x.strip() for x in re.split(r"[、，,]", s) if x.strip()]
    return out


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

    rows = []
    for r in range(2, ws.max_row + 1):
        subj = str(ws.cell(r, C_SUBJECT + 1).value or "").strip()
        if not subj:
            continue
        rows.append({
            "row": r,
            "org": org_of(subj),
            "decided": str(ws.cell(r, C_DECIDED + 1).value or "")[:10],
            "src": str(ws.cell(r, C_SRC + 1).value or "").strip(),
            "cat": str(ws.cell(r, C_CAT + 1).value or "").strip(),
            "paths": str(ws.cell(r, C_PATHS + 1).value or "").strip(),
            "files": split_multi(ws.cell(r, C_FILES + 1).value),
        })
    groups = defaultdict(list)
    for x in rows:
        groups[(x["src"], x["org"], x["decided"])].append(x)
    multi = {k: v for k, v in groups.items() if len(v) >= 2}

    print("== 同案拆行体检：%d 行 → 命中 %d 组 ==" % (len(rows), len(multi)))
    plans, conflicts = [], []
    for k, g in multi.items():
        g = sorted(g, key=lambda x: (-len(x["files"]), x["row"]))
        keep, drops = g[0], g[1:]
        why = []
        for d in drops:
            if d["cat"] != keep["cat"]:
                why.append("违规类型不一致（%s vs %s）" % (keep["cat"], d["cat"]))
            if d["paths"] != keep["paths"]:
                why.append("存放路径不一致")
        if why:
            conflicts.append((k, keep["row"], [d["row"] for d in drops], why))
            continue
        plans.append((k, keep, drops))
        print("\n【%s %s %s】保留 行%d（%d 份文件）" % (k[0], k[2], k[1], keep["row"], len(keep["files"])))
        for d in drops:
            print("   并入 行%-4d 序号%s 文件=%s" % (d["row"], ws.cell(d["row"], 1).value,
                                                "；".join(d["files"])[:50]))
    if conflicts:
        print("\n【冲突，不合并，交人工】%d 组" % len(conflicts))
        for k, keep, drops, why in conflicts:
            print("  %s %s %s ｜行%d ← %s ｜%s" % (k[0], k[2], k[1][:26], keep, drops, "；".join(why)))
    if not plans:
        print("\n无需合并。")
        return
    if not a.apply:
        print("\n（dry-run，未改动；加 --apply 落盘）")
        return

    from ledger_bak import backup_ledger
    print("\n台账已备份：%s" % backup_ledger(root, ledger, "merge_same_case"))
    deleted = []
    for k, keep, drops in plans:
        r = keep["row"]
        subj = str(ws.cell(r, C_SUBJECT + 1).value or "").strip()
        persons = persons_of(subj)
        meas = str(ws.cell(r, C_MEAS + 1).value or "").strip()
        files = split_multi(ws.cell(r, C_FILES + 1).value)
        urls = split_multi(ws.cell(r, C_URLS + 1).value)
        typ = str(ws.cell(r, C_TYP + 1).value or "").strip()
        for d in drops:
            dr = d["row"]
            dsub = str(ws.cell(dr, C_SUBJECT + 1).value or "").strip()
            persons += [p for p in persons_of(dsub) if p not in persons]
            # 措施：逐「高管（X）」追加，已有同名不重复
            dmeas = str(ws.cell(dr, C_MEAS + 1).value or "").strip()
            have = set(re.findall(r"高管（([^）]+)）", meas))
            for blk in [b.strip() for b in dmeas.split("\n") if b.strip()]:
                m = re.match(r"^(高管|管理人员|人员|机构|管理人)（([^）]+)）", blk)
                if m and m.group(1) == "高管" and m.group(2) in have:
                    continue
                if blk in meas:
                    continue
                meas = (meas + "\n" + blk).strip() if meas else blk
                if m:
                    have.add(m.group(2))
            files += [f for f in split_multi(ws.cell(dr, C_FILES + 1).value) if f not in files]
            urls += [u for u in split_multi(ws.cell(dr, C_URLS + 1).value) if u not in urls]
            # 保留行为空/敷衍时取被合并行
            for c in (C_GROUNDS, C_EVID, C_LEGAL, C_DEFENSE):
                cv = str(ws.cell(r, c + 1).value or "").strip()
                if cv in EMPTY or len(cv) < 4:
                    dv = str(ws.cell(dr, c + 1).value or "").strip()
                    if dv and dv not in EMPTY and len(dv) >= 4:
                        ws.cell(r, c + 1).value = dv
            if str(ws.cell(dr, C_TYP + 1).value or "").strip() == "是":
                typ = "是"
            deleted.append(dr)
        ws.cell(r, C_SUBJECT + 1).value = keep["org"] + (
            "；" + "、".join(persons) if persons else "")
        ws.cell(r, C_MEAS + 1).value = meas
        ws.cell(r, C_FILES + 1).value = "；".join(files)
        ws.cell(r, C_URLS + 1).value = "；".join(urls)
        ws.cell(r, C_TYP + 1).value = typ if typ else "否"
        print("  行%-4d 合并完成：人员 %d 人 / 文件 %d 份" % (r, len(persons), len(files)))

    for dr in sorted(set(deleted), reverse=True):
        ws.delete_rows(dr, 1)
    # 序号重排
    n = 0
    for r in range(2, ws.max_row + 1):
        if not str(ws.cell(r, C_SUBJECT + 1).value or "").strip():
            continue
        n += 1
        ws.cell(r, C_SEQ + 1).value = n
    wb.save(ledger)
    print("\n已删除 %d 行并重排序号，现共 %d 行" % (len(set(deleted)), n))


if __name__ == "__main__":
    main()
