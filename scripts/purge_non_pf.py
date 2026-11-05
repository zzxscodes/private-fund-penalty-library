# -*- coding: utf-8 -*-
"""
purge_non_pf.py —— 清理库内明确非私募误收（期货 / 券商 / 银行 / 信托 / 保险）。

背景（2026-10-29）：用户发现台账混入金鹏期货、华创期货等经纪商。本库主题是
【私募基金管理人】监管处罚，证监局平台下发的期货公司、证券公司、银行、信托、保险等
文书虽被 is_private_fund 判定「非私募」并登记进「不收录清单.md」，但因早期建库残留
未被回捞，仍滞留在案例台账里。

清理动作（--apply 才落盘，自动备份台账）：
  1) 识别主体名不含「私募」且命中「期货 / 券商明确后缀 / 银行 / 信托 / 保险 / 公募」的行；
  2) 把这些行的决定书 PDF 移入「库父目录/_case_excluded_non_pf/<来源>/」（可恢复，非物理删除）；
  3) 从台账删除这些行（整表安全重写 + 重套样式 + renumber，复用 ledger_style）；
  4) 把清理的主体登记进库根「不收录清单.md」（5 列格式，与现有条目一致，去重）。

判定只认「不含私募 且 命中排除词」，因此「XX基金管理有限公司 / XX股权投资基金管理有限公司」
等真私募不会误伤（它们既不含私募二字、也不命中排除词）。

用法：
  python scripts/purge_non_pf.py --root "<库>"            # dry-run
  python scripts/purge_non_pf.py --root "<库>" --apply     # 执行
"""
import os
import re
import sys
import glob
import shutil
import argparse
import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl
import ledger_style as LS
# 排除词唯一事实源：classify_filter._NOT_PF_NAME（禁止本地副本，见 SKILL.md 0.5.2）。
from classify_filter import _NOT_PF_NAME


def is_non_pf(subj):
    """明确非私募：整行不含『私募』且命中排除词（期货/券商/银行/信托/保险/公募）。"""
    if "私募" in (subj or ""):
        return False
    return bool(_NOT_PF_NAME.search(subj or ""))


def find_pdf(root, d, name):
    cand = os.path.join(root, d, name) if d else os.path.join(root, name)
    if os.path.isfile(cand):
        return cand
    hits = glob.glob(os.path.join(root, "0[123]_*", "**", name), recursive=True)
    return hits[0] if hits else None


def load_list_file(path):
    """读库根不收录清单，返回 (entries, urls, max_no)。格式 5 列 #/bureau/target/url/reason。"""
    entries, urls, max_no = [], set(), 0
    if not os.path.isfile(path):
        return entries, urls, max_no
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line.startswith("|"):
                continue
            cells = [c.strip() for c in line.strip("|").split("|")]
            if len(cells) < 5:
                continue
            no, bureau, target, url, reason = cells[0], cells[1], cells[2], cells[3], cells[4]
            if not target or target.startswith("`target`") or target == "target":
                continue
            entries.append((no, bureau, target, url, reason))
            if url:
                urls.add(url)
            try:
                max_no = max(max_no, int(no))
            except ValueError:
                pass
    return entries, urls, max_no


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True,
                    help="案例库根目录（如：私募监管处罚案例库（中基协+证监局+交易所））")
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    root = a.root
    # 备份/排除目录与库根同级（可迁移：不再写死本机绝对路径）。
    parent = os.path.dirname(os.path.abspath(root))
    SNAP = os.path.join(LS.maint_dir(root), "_case_repair_snapshot")
    EXCLUDE_DIR = LS.case_excluded_dir(root)
    wb = openpyxl.load_workbook(os.path.join(root, LS.LEDGER_NAME))
    ws = wb.active
    hdr = [c.value for c in ws[1]]
    I = {h: i for i, h in enumerate(hdr)}

    rows = []
    for r in range(2, ws.max_row + 1):
        v = [ws.cell(r, c + 1).value for c in range(len(hdr))]
        if not any(x not in (None, "") for x in v[1:12]):
            continue
        rows.append((r, v))

    hit = []
    for r, v in rows:
        subj = str(v[I["被处罚机构/人"]] or "")
        src = str(v[I["机构来源"]] or "")
        if is_non_pf(subj):
            hit.append((r, v, src, subj))

    print("识别非私募行：%d" % len(hit))
    for r, v, src, subj in hit:
        print("  行%-4d [%-3s] %s" % (r, src[:3], subj[:42]))

    if not a.apply:
        print("\n（dry-run，未改动。加 --apply 执行清理）")
        return

    # ---- 备份台账 ----
    os.makedirs(SNAP, exist_ok=True)
    bk = os.path.join(SNAP, "ledger_pre_purge_nonpf_%s.xlsx"
                      % datetime.date.today().isoformat())
    shutil.copy2(os.path.join(root, LS.LEDGER_NAME), bk)
    print("\n台账已备份：", bk)

    # ---- 1) 移出 PDF ----
    os.makedirs(EXCLUDE_DIR, exist_ok=True)
    moved = 0
    for r, v, src, subj in hit:
        paths = [p.strip() for p in str(v[I["存放路径"]] or "").split("；") if p.strip()]
        fns = [f.strip() for f in str(v[I["案例文件"]] or "").split("；") if f.strip()]
        dest_dir = os.path.join(EXCLUDE_DIR, src or "未知")
        os.makedirs(dest_dir, exist_ok=True)
        for i, fn in enumerate(fns):
            d = paths[i] if i < len(paths) else (paths[-1] if paths else "")
            src_pdf = find_pdf(root, d, fn)
            if not src_pdf:
                print("    [WARN] 找不到 PDF：行%d %s" % (r, fn))
                continue
            dst = os.path.join(dest_dir, os.path.basename(fn))
            if os.path.abspath(src_pdf) != os.path.abspath(dst):
                shutil.move(src_pdf, dst)
                moved += 1
    print("移出 PDF：%d 份 → %s" % (moved, EXCLUDE_DIR))

    # ---- 2) 删行（整表安全重写：保留行 append_row + renumber）----
    keep = [(r, v) for (r, v) in rows if (r, v) not in [(h[0], h[1]) for h in hit]]
    # 直接按行号差集收集保留行
    hit_rows = set(h[0] for h in hit)
    keep_vals = [v for (r, v) in rows if r not in hit_rows]
    for r_i in range(ws.max_row, 1, -1):
        ws.delete_rows(r_i)
    for v in keep_vals:
        LS.append_row(ws, v, LS.CASE_WRAP_COLS)
    LS.renumber(ws)
    wb.save(os.path.join(root, LS.LEDGER_NAME))
    print("台账重写完成：保留 %d 行，删除 %d 行" % (len(keep_vals), len(hit)))

    # ---- 3) 登记不收录清单（去重追加）----
    list_path = os.path.join(root, "不收录清单.md")
    _, urls, max_no = load_list_file(list_path)
    added = 0
    with open(list_path, "a", encoding="utf-8") as f:
        f.write("\n<!-- 2026-10-29 清理非私募误收（purge_non_pf.py） -->\n")
        for r, v, src, subj in hit:
            url = str(v[I["原文链接"]] or "").strip()
            if url and url in urls:
                continue
            max_no += 1
            target = subj.replace("|", "／")
            f.write("| %d | %s | %s | %s | 非私募（期货公司/证券公司/银行，非私募基金管理人） |\n"
                    % (max_no, src, target, url))
            urls.add(url)
            added += 1
    print("不收录清单新增：%d 条（去重后）" % added)
    print("\n完成。被清理的 %d 行 PDF 已移至 %s，台账已备份，可随时恢复。" % (len(hit), EXCLUDE_DIR))


if __name__ == "__main__":
    main()
