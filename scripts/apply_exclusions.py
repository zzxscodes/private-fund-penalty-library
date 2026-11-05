# -*- coding: utf-8 -*-
"""apply_exclusions.py — 依据 screen_nonpf v3 的"非私募"命中，落地剔除。
做法（保守、可回滚）：
  * 备份台账 .xlsx 到同目录 .bak_<date>
  * 解析 v3 扫描输出的"非私募"命中行（rel 路径）
  * 将每个命中 PDF 移动到库父目录 _case_excluded_non_pf/ 下（镜像子目录结构，只移不删）
  * 在台账中按「案例文件(col12) 含该 basename」定位并删除对应行（自底向上删，避免错位）
  * 向 不收录清单.md 追加登记（# 顺延）
  * 更新 不收录清单.md 头部「条数」
用法：python apply_exclusions.py --root <库根> --scan <v3扫描文件>
"""
import argparse, os, re, shutil, datetime
import openpyxl
import ledger_style as LS  # noqa: E402  运维目录唯一真源（case_excluded_dir）

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--scan", required=True)
    args = ap.parse_args()
    ROOT = args.root
    EXCL = LS.case_excluded_dir(ROOT)
    LIST = os.path.join(ROOT, "不收录清单.md")

    # 1) 解析扫描命中
    rels = []
    with open(args.scan, encoding="utf-8") as f:
        for line in f:
            m = re.match(r"\s*(.+?)\s*<<\s*(主体名明显非私募|仅引用)", line)
            if m:
                rels.append(m.group(1).strip())
    print("[apply] v3 非私募命中: %d" % len(rels))

    # 2) 备份台账（落在运维主目录，不污染交付库）
    os.makedirs(LS.maint_dir(ROOT), exist_ok=True)
    bk = os.path.join(LS.maint_dir(ROOT), "台账备份_apply_exclusions_" + datetime.date.today().isoformat() + ".xlsx")
    shutil.copy2(LEDGER, bk)
    print("[apply] 台账备份 -> %s" % bk)

    # 3) 移动 PDF + 收集台账行
    wb = openpyxl.load_workbook(LEDGER)
    ws = wb.active
    rows_del = set()
    recs = []
    moved = 0
    for rel in rels:
        src = os.path.join(ROOT, rel)
        base = os.path.basename(rel)
        if not os.path.exists(src):
            print("  [skip] 文件不存在(可能已移): %s" % rel)
            continue
        dst = os.path.join(EXCL, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        for ri in range(2, ws.max_row + 1):
            c12 = ws.cell(row=ri, column=12).value or ""
            if base in c12 and ri not in rows_del:
                rows_del.add(ri)
                recs.append((ri, c12, ws.cell(row=ri, column=11).value,
                             ws.cell(row=ri, column=14).value))
        shutil.move(src, dst)
        moved += 1
        print("  [moved] %s" % rel)

    # 4) 自底向上删行
    for ri in sorted(rows_del, reverse=True):
        ws.delete_rows(ri, 1)
    wb.save(LEDGER)
    print("[apply] 删除台账行 %d 条: %s" % (len(rows_del), sorted(rows_del)))

    # 5) 追加不收录清单
    txt = open(LIST, encoding="utf-8").read()
    maxno = 0
    for l in txt.splitlines():
        m = re.match(r"\|\s*(\d+)\s*\|", l)
        if m:
            maxno = max(maxno, int(m.group(1)))
    newrows = []
    for (ri, c12, src, url) in recs:
        maxno += 1
        target = c12.replace("；", " / ")[:220]
        reason = "非私募（与私募基金监管无关）：screen_nonpf v3 按主体名/引用法规判定，不属私募基金管理人/私募基金处罚"
        newrows.append("| %d | %s | %s | %s | %s |" %
                       (maxno, src or "", target, url or "", reason))
    with open(LIST, "a", encoding="utf-8") as f:
        f.write("\n" + "\n".join(newrows) + "\n")
    # 6) 更新头部条数
    newtxt = re.sub(r"(条数：\*\*\d+\*\*)", "条数：**%d**" % maxno, txt, count=1)
    if newtxt != txt:
        open(LIST, "w", encoding="utf-8").write(newtxt)
    print("[apply] 不收录清单新增 %d 条，现总计 %d" % (len(recs), maxno))
    print("[apply] 完成：移动 %d 个 PDF，删除 %d 行台账。" % (moved, len(rows_del)))

if __name__ == "__main__":
    main()
