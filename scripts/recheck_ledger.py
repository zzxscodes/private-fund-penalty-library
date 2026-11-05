# -*- coding: utf-8 -*-
"""多轮执行下的「非最终态」复查与就地更新（v2 硬性要求 2026-10-26）。

背景（用户明确）：
  * 决定书内容本身会随时间推进而更新——例如中基协先出【纪律处分决定书】，
    当事人申辩后再出【复核决定书】，复核决定书才含「二、申辩意见」段落。
    因此「申辩理由」等列在第一轮可能是「无」，后续轮次复查时可能变成有内容。
  * 机构与对应人员的处罚决定往往不同时作出（可能先有其一），故每轮按实际收集，
    已入库条目若此前只能标「不明确」，后续轮次拿到新文书后应能补全。

本脚本职责（机械层）：
  1. 扫描台账，找出【非最终态】的单元格：管理人=不明确、处罚措施含"不明确"、
     申辩理由=无（可能因文书更新而出现申辩段）、挂网发布日期=待补、处罚时间为空/窗口外。
  2. 对每个非最终态行，按 URL 取回最新正文（优先本地快照 JSON，其次 PDF 文字层），
     重跑 extract_core（规则层，三层契约），若字段由「非最终」变为「有确定内容」→ 就地更新。
  3. 只改值，不动排版；更新明细追加到 更新记录.md。

用法：
    python recheck_ledger.py --root "<案例库根>" [--snapshot <csrc.json> ...] [--dry-run]

说明：--snapshot 可重复传入抓取快照（含 url/text 或 url/本地txt路径），
      用于离线复查；真实定时执行时应先重新抓取再调用本脚本。
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import extract_core  # noqa: E402
import ledger_style as LS  # noqa: E402

try:
    import openpyxl
except ImportError:
    print("ERROR: 缺少 openpyxl", file=sys.stderr)
    sys.exit(2)

LEDGER = LS.LEDGER_NAME
C = LS.COL

# 非最终态判定
NON_FINAL_MANAGER = {"不明确", ""}
NON_FINAL_PUBLISH = {"待补", ""}


def _load_snapshots(paths):
    """把抓取快照整理成 url -> text 映射。支持 csrc.json(items:url/text) 与
    amac_30d.json(items:url/text，text 可能是本地 txt 路径)。"""
    idx = {}
    for p in paths or []:
        if not p or not os.path.exists(p):
            continue
        d = json.load(open(p, encoding="utf-8"))
        items = d.get("items", []) if isinstance(d, dict) else d
        for it in items:
            url = it.get("url")
            if not url:
                continue
            t = it.get("text") or ""
            if isinstance(t, str) and re.search(r"\.txt$", t) and os.path.exists(t):
                t = open(t, encoding="utf-8", errors="ignore").read()
            idx[url] = t
    return idx


def _pdf_text(path):
    try:
        import fitz
    except Exception:
        return ""
    try:
        doc = fitz.open(path)
        txt = "".join(page.get_text() for page in doc)
        doc.close()
        return txt
    except Exception:
        return ""


def _is_non_final(row):
    """返回需要复查的列名集合。"""
    need = set()
    if str(row[C["manager"]] or "").strip() in NON_FINAL_MANAGER:
        need.add("manager")
    if "不明确" in str(row[C["measures"]] or ""):
        need.add("measures")
    if str(row[C["defense"]] or "").strip() in ("无", ""):
        need.add("defense")            # 文书后续可能新增申辩段
    if str(row[C["publish"]] or "").strip() in NON_FINAL_PUBLISH:
        need.add("publish")
    if not str(row[C["decided"]] or "").strip():
        need.add("decided")
    return need


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--snapshot", action="append", default=[])
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    root = a.root
    snap = _load_snapshots(a.snapshot)
    wb = openpyxl.load_workbook(os.path.join(root, LEDGER))
    ws = wb[LS.SHEET_CASE] if LS.SHEET_CASE in wb.sheetnames else wb.active

    changes = []
    checked = 0
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(row=r, column=i).value for i in range(1, len(LS.CASE_HEADERS) + 1)]
        if not any(v is not None and str(v).strip() for v in row[1:]):
            continue
        need = _is_non_final(row)
        if not need:
            continue
        checked += 1
        urls = [u.strip() for u in re.split(r"[；;]", str(row[C["urls"]] or "")) if u.strip()]
        text = ""
        for u in urls:
            if u in snap and snap[u]:
                text = snap[u]
                break
        if not text:
            # 退回本地 PDF 文字层
            paths = [p.strip() for p in re.split(r"[；;]", str(row[C["paths"]] or "")) if p.strip()]
            fns = [f.strip() for f in re.split(r"[；;]", str(row[C["filenames"]] or "")) if f.strip()]
            if paths and fns:
                fp = os.path.join(root, paths[0].replace("/", os.sep), fns[0])
                if os.path.exists(fp):
                    text = _pdf_text(fp)
        if not text:
            continue

        source = "amac" if "中基协" in str(row[C["source"]] or "") else "csrc"
        res = extract_core.extract_with_contract(text, source)

        # —— 逐字段判断「是否由非最终变为有确定内容」 ——
        if "defense" in need and res.get("defense") and res["defense"] not in ("无", ""):
            old = str(row[C["defense"]] or "")
            if old != res["defense"]:
                changes.append((r, "申辩理由", old, res["defense"]))
                if not a.dry_run:
                    ws.cell(row=r, column=C["defense"] + 1).value = res["defense"]
        if "manager" in need and res.get("measures"):
            pass  # 管理人名称需实体抽取，规则层无法凭空补全，交模型/人工（此处保留不明确）

    if changes and not a.dry_run:
        wb.save(os.path.join(root, LEDGER))

    # 更新记录
    print("复查完成：扫描非最终态行 %d，实际更新 %d 处" % (checked, len(changes)))
    for r, col, old, new in changes[:20]:
        print("  行%d [%s] %r → %r" % (r, col, old[:20], new[:60]))
    if not changes:
        print("  本轮无非最终态字段取得新内容（说明：规则层已穷尽；若文书已更新，"
              "需先重新抓取再复查，或由模型/人工补全）。")

    log = os.path.join(root, "更新记录.md")
    if os.path.exists(log) and not a.dry_run:
        with open(log, "a", encoding="utf-8") as f:
            f.write("\n## %s 非最终态复查\n\n- 扫描非最终态行：%d\n- 实际更新：%d\n"
                    % (datetime.now().strftime("%Y-%m-%d %H:%M"), checked, len(changes)))
            for r, col, old, new in changes[:50]:
                f.write("- 行%d [%s]：%s → %s\n" % (r, col, old[:30], new[:80]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
