# -*- coding: utf-8 -*-
"""v2 入库机械层：把「已抽取好三核心列的案例记录」落盘 + 写台账 + 重排 + 校验。

输入 JSON（records 数组）每条 = 一个 v2 案例行（机构 + 高管合并）：
{
  "source_dir": "01_中基协自律处分（纪律处分、公开谴责、暂停备案、注销登记）",   # 顶层分类目录名
  "category":   "04_内控治理、人员资质、风控缺位",                              # 违规类型子目录名
  "manager":    "上海隐上私募基金管理有限公司",
  "date_decided": "2026-07-17",          # 处罚时间（落款/决定日期）
  "publish_date": "2026-10-16",          # 挂网发布日期（取不到标 待补）
  "measures": ["管理人：撤销管理人登记",
               "高管（现任法定代表人韩方圆）：取消基金从业资格", ...],   # 分述，与 files 对齐
  "grounds":  ["1.未建立健全内部控制……", ...],     # 核心处罚事由
  "evidence": ["互联网公开信息、……", ...],          # 核心证据材料
  "legal_basis": ["《私募投资基金监督管理暂行办法》第四条", ...],  # 核心处罚依据（到条款号）
  "files": [{"src": "<本地PDF路径>", "name": "上海隐上私募基金管理有限公司.pdf"},
            {"src": "...", "name": "现任法定代表人韩方圆-上海隐上私募基金管理有限公司.pdf"}, ...],
  "urls": ["https://...", ...],
  "typical": false
}

为什么独立成脚本：抽取（读全文、列全三核心列）是高质量人工/子代理活；落盘/写台账/重排/校验是机械活。
本脚本只做机械活，不读决定书，避免在主会话里反复改台账样式。
"""
import sys, os, json, shutil, argparse
import openpyxl
import ledger_style as LS

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

# 14 列顺序（与 ledger_style.COL 一致）
HEADERS = ["序号", "处罚时间", "挂网发布日期", "管理人名称", "处罚措施（管理人+高管）",
           "核心处罚事由", "核心证据材料", "核心处罚依据", "申辩理由", "违规类型", "机构来源",
           "案例文件", "存放路径", "原文链接", "是否典型案例"]
SEP = "；"


def _join(arr):
    return SEP.join(x for x in arr if x)


def _date_num(s):
    try:
        y, m, d = s.split("-")
        return int(y) * 10000 + int(m) * 100 + int(d)
    except Exception:
        return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--records", required=True, help="抽取好的案例记录 JSON")
    ap.add_argument("--ledger", default="04_综合索引台账【总表】.xlsx")
    a = ap.parse_args()

    recs = json.load(open(a.records, encoding="utf-8"))
    if isinstance(recs, dict):
        recs = recs.get("records", [])
    wb = openpyxl.load_workbook(os.path.join(a.root, a.ledger))
    ws = wb.active

    # 现有最大序号
    max_seq = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r and r[0] and str(r[0]).isdigit():
            max_seq = max(max_seq, int(r[0]))
    added = 0
    for rec in recs:
        sd = rec["source_dir"]; cat = rec["category"]
        mgr = rec.get("subject", rec.get("manager", ""))
        dest_dir = os.path.join(a.root, sd, cat)
        os.makedirs(dest_dir, exist_ok=True)
        names, paths, urls = [], [], []
        for f in rec.get("files", []):
            src = f["src"]; name = f["name"]
            dst = os.path.join(dest_dir, name)
            if src and os.path.exists(src) and not os.path.exists(dst):
                shutil.copy2(src, dst)
            names.append(name)
            # 存放路径列只存分类目录（不含文件名），与 v2 约定一致；
            # 校验时按 ROOT + 存放路径 + basename(案例文件) 拼回磁盘路径。
            paths.append(os.path.dirname(os.path.relpath(dst, a.root)).replace("\\", "/"))
            urls.append("")  # 占位，下面用 rec.urls 覆盖
        # 原文链接与各文件一一对应（不足则复用最后一个）
        ulist = rec.get("urls", [])
        url_cells = []
        for i in range(len(names)):
            url_cells.append(ulist[i] if i < len(ulist) else (ulist[-1] if ulist else ""))
        max_seq += 1
        measures = LS.format_measures_list(rec.get("measures", []))  # 按主体分行（其余列仍用「；」）
        grounds = LS.normalize_grounds_items(_join(rec.get("grounds", [])))  # 一点一行、序号统一「一、」
        evidence = _join(rec.get("evidence", []))
        legal = _join(rec.get("legal_basis", []))
        row = [max_seq, rec.get("date_decided", ""), rec.get("publish_date", "待补"), mgr,
               measures, grounds, evidence, legal, rec.get("defense", "无"),
               cat, rec.get("source", "中基协"),
               _join(names), LS.paths_cell(paths), _join(url_cells),
               "是" if rec.get("typical") else "否"]
        ws.append(row)
        added += 1
        print("  + [seq %d] %s | %s | 文件 %d" % (max_seq, mgr, cat, len(names)))

    # 重排：按挂网发布日期 DESC（待补最后），其次处罚时间 DESC
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    keys = []
    for r in rows:
        if r[0] is None:
            continue
        pub = r[2] or ""
        pub_key = (1, 0) if (not pub or pub == "待补") else (0, -_date_num(pub))
        keys.append((pub_key, -_date_num(r[1] or ""), str(r[9]), str(r[10]), list(r)))
    keys.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    for i, k in enumerate(keys, start=2):
        for j, v in enumerate(k[4], start=1):
            ws.cell(i, j, v if j != 1 else (i - 1))
    wb.save(os.path.join(a.root, a.ledger))
    print("\n入库完成：新增 %d 行；台账现 %d 行（已按挂网发布日倒序重排）。"
          % (added, len(keys)))
    print("下一步：运行 verify_library.py 校验全库。")


if __name__ == "__main__":
    main()
