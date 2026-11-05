# -*- coding: utf-8 -*-
"""综合索引台账（04_综合索引台账【总表】.xlsx）的查重与**增量追加**（v2 格式）。

!!! 核心原则：台账只做「内容增加」，绝不重建、绝不改排版 !!!
  本脚本所有写操作都是：打开已有工作簿 → 在末尾追加行 → 保存。
  表头文字、列顺序、列宽、配色、字体、冻结窗格一律不动；
  新增数据行自动复制上一数据行的样式。只有文件不存在时才按冻结规格新建。

v2（2026-10-26）：一行=一个案例（机构+高管决定书合并为一行）；
无文号列；新增挂网发布日期/核心处罚事由/核心证据材料/核心处罚依据。
列定义见 ledger_style.py 与 references/ledger-format.md。

用法:
    python ledger.py --root "<案例库根路径>" check --key "<管理人名或文件名或链接>"
    python ledger.py --root "<案例库根路径>" add --json "<记录JSON>"
    python ledger.py --root "<案例库根路径>" add-batch --file "<记录数组JSON文件>"
    python ledger.py --root "<案例库根路径>" latest [--source "<机构来源关键词>"]
    python ledger.py --root "<案例库根路径>" law-sheet --file "<法规数组JSON文件>"
    python ledger.py --root "<案例库根路径>" sort
    python ledger.py --root "<案例库根路径>" verify

check:     按管理人名称/案例文件/原文链接/处罚时间查重。退出码 0=未收录，1=已收录。
add:       追加一条案例记录（先按案例聚合好机构+人员决定书再提交）。
add-batch: 批量追加（内部按 案例键(source+manager+date_decided)/文件名/链接 去重）。
latest:    输出最新处罚时间（决定书落款），用于默认窗口的增量检索起点。
law-sheet: 更新「配套法规依据清单」——按「法规名称」匹配，原地刷新或末尾追加。
sort:      按「挂网发布日期倒序（最新在前，待补最后）」重排；同发布日期按处罚时间倒序。
verify:    校验台账格式是否仍是 v2 冻结规格。退出码 0=格式未被改动。
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ledger_style as LS  # noqa: E402

# 提交记录的字段（filenames/paths/urls 为数组，其余字符串）
CASE_FIELDS = [
    "date_decided", "publish_date", "subject", "measures",
    "grounds", "evidence", "legal_basis", "defense", "category", "source",
    "filenames", "paths", "urls", "typical",
]
# 2026-10-27 用户定稿：移除无法保证取到的 6 个字段（文号/发布机关/发布日期/
# 施行日期/效力状态/留存形式），只留能稳定取到的 6 个。
LAW_FIELDS = ["category", "name", "refs", "filename", "path", "size"]


def _open(root: Path):
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("ERROR: 缺少 openpyxl，请在托管 Python 隔离 venv 中安装: pip install openpyxl", file=sys.stderr)
        sys.exit(2)
    ledger = root / LS.LEDGER_NAME
    if not ledger.exists():
        print(f"台账不存在，按冻结规格新建: {ledger}", file=sys.stderr)
        wb = LS.new_workbook()
        wb.save(ledger)
    wb = load_workbook(ledger)
    if LS.SHEET_CASE in wb.sheetnames:
        ws = wb[LS.SHEET_CASE]
    else:
        ws = wb.active
        ws.title = LS.SHEET_CASE
    return wb, ws


def _cell(row, idx):
    v = row[idx] if idx < len(row) else None
    return "" if v is None else str(v).strip()


def _split_multi(v: str):
    """把 K/L/M 多值列按「；」拆回数组（兼容半角分号、换行）。"""
    for sep in ("；", ";", "\n"):
        v = v.replace(sep, "；")
    return [x.strip() for x in v.split("；") if x.strip()]


def _existing_keys(ws):
    """已收录键集合（v2 无文号）：
    ("case", source, manager, date_decided) + 每个文件名 + 每条链接。
    """
    C = LS.COL
    keys = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        source = _cell(row, C["source"])
        manager = _cell(row, C["subject"])
        decided = _cell(row, C["decided"])
        if source and manager and decided:
            keys.add(("case", source, manager, decided))
        for fn in _split_multi(_cell(row, C["filenames"])):
            keys.add(("fn", fn))
        for u in _split_multi(_cell(row, C["urls"])):
            keys.add(("url", u))
    return keys


def _keys_of(rec: dict):
    """待追加记录的去重键（与 _existing_keys 同构）。

    v2 判重不再用文号（文号列已删）。案例级主键 = 机构来源+管理人名称+处罚时间。
    """
    ks = set()
    source = str(rec.get("source") or "").strip()
    manager = str(rec.get("subject") or rec.get("manager") or "").strip()
    decided = str(rec.get("date_decided") or "").strip()
    if source and manager and decided:
        ks.add(("case", source, manager, decided))
    for fn in rec.get("filenames") or []:
        if str(fn).strip():
            ks.add(("fn", str(fn).strip()))
    for u in rec.get("urls") or []:
        if str(u).strip():
            ks.add(("url", str(u).strip()))
    return ks


def cmd_check(root: Path, key: str) -> int:
    _, ws = _open(root)
    key = key.strip()
    C = LS.COL
    for row in ws.iter_rows(min_row=2, values_only=True):
        for idx in (C["subject"], C["filenames"], C["urls"], C["decided"]):
            if idx < len(row) and row[idx] is not None and key and key in str(row[idx]):
                print(f"DUPLICATE: 第{row[0]}行 命中[{LS.CASE_HEADERS[idx]}]: {row[idx]}")
                return 1
    print("NEW: 台账中未收录")
    return 0


def _append_cases(root: Path, wb, ws, records) -> int:
    C = LS.COL
    have = _existing_keys(ws)
    added = 0
    for rec in records:
        ks = _keys_of(rec)
        if ks & have:
            print(f"SKIP(已收录): {str(rec.get('subject') or rec.get('manager', ''))[:30]} | {rec.get('date_decided', '')}")
            continue
        row = [0,
               rec.get("date_decided", ""),
               rec.get("publish_date", "待补"),
               rec.get("subject", rec.get("manager", "")),
               rec.get("measures", ""),
               rec.get("grounds", ""),
               rec.get("evidence", ""),
               rec.get("legal_basis", ""),
               rec.get("defense", ""),
               rec.get("category", ""),
               rec.get("source", ""),
               "；".join(rec.get("filenames") or []),
               LS.paths_cell(rec.get("paths") or []),
               "；".join(rec.get("urls") or []),
               rec.get("typical", "否")]
        LS.append_row(ws, row, LS.CASE_WRAP_COLS)
        have |= ks
        added += 1
        print(f"ADDED: {str(row[C['subject']])[:34]} | {row[C['decided']]} | {row[C['source']]}")
    if added:
        LS.renumber(ws)
        wb.save(root / LS.LEDGER_NAME)
    return added


def cmd_add(root: Path, record: dict) -> int:
    wb, ws = _open(root)
    n = _append_cases(root, wb, ws, [record])
    print(f"新增 {n} 条，台账现有 {ws.max_row - 1} 条")
    return 0


def cmd_add_batch(root: Path, path: str) -> int:
    records = json.load(open(path, encoding="utf-8"))
    wb, ws = _open(root)
    n = _append_cases(root, wb, ws, records)
    print(f"批量新增 {n} / 提交 {len(records)} 条，台账现有 {ws.max_row - 1} 条")
    return 0


def cmd_latest(root: Path, source: str) -> int:
    _, ws = _open(root)
    C = LS.COL
    dates = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[C["decided"]] is None:
            continue
        if source and source not in _cell(row, C["source"]):
            continue
        dates.append(_cell(row, C["decided"]))
    print(max(dates) if dates else "NONE")
    return 0


def cmd_law_sheet(root: Path, path: str) -> int:
    """更新法规清单：已存在按名称原地更新，新增追加到末尾。不改排版。"""
    laws = json.load(open(path, encoding="utf-8"))
    wb, _ = _open(root)
    ws = LS.ensure_law_sheet(wb)
    idx = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=3).value
        if v:
            idx[str(v).strip()] = r
    added = updated = 0
    for law in laws:
        name = str(law.get("name") or "").strip()
        if not name:
            continue
        if name in idx:
            r = idx[name]
            # 列号按 LAW_FIELDS 现序推导（第 1 列是序号，故字段 i 落在第 i+2 列）
            for i, field in enumerate(LAW_FIELDS):
                if field == "name":
                    continue
                val = law.get(field)
                if val not in (None, ""):
                    ws.cell(row=r, column=i + 2).value = val
            # 引用数是每次都要刷新的字段
            if law.get("refs") not in (None, ""):
                ws.cell(row=r, column=LAW_FIELDS.index("refs") + 2).value = law["refs"]
            updated += 1
        else:
            row = [0] + [law.get(f, "") for f in LAW_FIELDS]
            LS.append_row(ws, row, LS.LAW_WRAP_COLS)
            idx[name] = ws.max_row
            added += 1
    if added:
        LS.renumber(ws)
    wb.save(root / LS.LEDGER_NAME)
    print(f"法规清单：新增 {added} 条，更新 {updated} 条，共 {ws.max_row - 1} 条")
    return 0


def _date_num(d: str):
    """'YYYY-MM-DD'/'YYYY-MM' → 可比较数值；解析失败返回 0。"""
    d = str(d or "")
    ym = (d[:4] + d[5:7]).ljust(6, "0") if len(d) >= 7 else "000000"
    day = d[8:10] if len(d) >= 10 else "00"
    try:
        return int(ym) * 100 + int(day)
    except Exception:
        return 0


def _sort_key(row):
    """排序键（v2）：挂网发布日期 DESC（「待补」排最后）→ 处罚时间 DESC
    → 机构来源 ASC → 案例文件 ASC 兜底稳定。"""
    C = LS.COL
    pub = _cell(row, C["publish"])
    if not pub or pub == "待补":
        pub_key = (1, 0)          # 待补/空 排最后
    else:
        pub_key = (0, -_date_num(pub))
    return (pub_key,
            -_date_num(_cell(row, C["decided"])),
            _cell(row, C["source"]),
            _cell(row, C["filenames"]))


def cmd_sort(root: Path) -> int:
    """按「挂网发布日期倒序（待补最后）」重排案例台账。

    只搬「值」，不碰任何样式：openpyxl 的样式挂在单元格上，把排序后的值写回
    原单元格区域，表头/列宽/配色/边框/冻结窗格全部保持不动。
    """
    wb, ws = _open(root)
    if ws.max_row < 3:
        print("数据不足 2 行，无需排序")
        return 0
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]
    before = [(str(r[LS.COL["publish"]] or ""), str(r[LS.COL["decided"]] or "")) for r in rows]
    rows.sort(key=_sort_key)
    after = [(str(r[LS.COL["publish"]] or ""), str(r[LS.COL["decided"]] or "")) for r in rows]
    if before == after:
        # 顺序没变也要 renumber——清除「尾部幽灵行」
        n = LS.renumber(ws)
        wb.save(root / LS.LEDGER_NAME)
        print(f"已是目标顺序；已重排序号并清除尾部空行，实际数据行 {n}")
        return 0
    for i, vals in enumerate(rows, start=2):
        for c, v in enumerate(vals, start=1):
            ws.cell(row=i, column=c).value = v
    LS.renumber(ws)
    wb.save(root / LS.LEDGER_NAME)
    n_pending = sum(1 for r in rows if not str(r[LS.COL["publish"]] or "").strip() or str(r[LS.COL["publish"]]) == "待补")
    print(f"已按挂网发布日期倒序重排 {len(rows)} 条（最新在前，待补 {n_pending} 条排最后）")
    return 0


# 缺失列的默认值：宁可写明确的占位，也不要留空让人误以为"没提取到"
_MIGRATE_DEFAULT = {"申辩理由": "无", "是否典型案例": "否", "挂网发布日期": "待补"}

# 列名别名：旧版本用过别的叫法。**必须先改名再对齐**，否则会被判成
# 「缺列 + 多余列」→ 旧列被删、新列留空 = 整列数据静默丢失。
# （2026-10-27 实测：现役库是「管理人名称」「处罚措施（管理人+高管）」，
#    v3 冻结规格是「被处罚机构/人」「处罚措施（机构+人员）」，不改名即丢两列内容。）
_ALIAS = {
    "被处罚机构/人": ["管理人名称", "被处罚机构", "处罚对象", "机构名称", "当事人"],
    "处罚措施（机构+人员）": ["处罚措施（管理人+高管）", "处罚措施（机构+高管）",
                             "处罚措施", "处罚结果"],
    "处罚时间": ["决定日期", "处罚日期", "落款日期"],
    "挂网发布日期": ["挂网日期", "发布日期", "公开日期"],
    "违规类型": ["违规类别", "分类"],
    "机构来源": ["来源", "机构"],
    "案例文件": ["文件名", "决定书文件"],
    "存放路径": ["路径"],
    "原文链接": ["链接", "原文网址"],
    "是否典型案例": ["典型案例"],
    "核心处罚事由": ["案由摘要", "案由", "违规事实"],
}
_REV_ALIAS = {old: new for new, olds in _ALIAS.items() for old in olds}


def _canon(name):
    return _REV_ALIAS.get(name, name)


def cmd_migrate(root: Path) -> int:
    """把旧格式台账**就地升级**到冻结的 v3 规格（15 列 + 法规清单 7 列），保留全部已有行。

    为什么要有这一步（2026-10-27）：
      技能升级到 v3 后新增了「申辩理由」列；而现役库可能是 v2（14 列）甚至更早的 12 列。
      直接按 v3 列序追加会**整行错位**（违规类型写进申辩理由列…），是静默数据损坏。
      可迁移场景（拷到别人机器上遇到旧库）同样必须先 migrate 再追加。
    做法：
      1) 按目标表头顺序逐列对齐：缺的列 insert_cols 补上并填默认值；
         目标里没有的旧列（如「文号」「案由摘要」）删除并**打印警告**（v3 判重不用文号）。
      2) 表头/列宽/配色/冻结窗格按 ledger_style 重套（不动数据行内容）。
      3) 法规清单 sheet 存在但列不是 7 列时，按列名映射重建，保留能对上的值。
    """
    from openpyxl.utils import get_column_letter

    wb, ws = _open(root)
    cur = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    renamed = [(o, _canon(o)) for o in cur if _canon(o) != o]
    if renamed:
        for i, (o, n) in [(i, (o, _canon(o))) for i, o in enumerate(cur)]:
            if o != n:
                ws.cell(row=1, column=i + 1).value = n
        print("列名规范化（旧名 → v3 名）：%s"
              % "；".join("%s→%s" % (o, n) for o, n in renamed))
    cur = [_canon(x) for x in cur]
    if cur == LS.CASE_HEADERS:
        print("案例台账已是 v3 冻结规格，无需迁移")
    else:
        print("当前表头: %s" % cur)
        ins, drop = [], []
        # --- 1) 补缺失列（按目标顺序就地插入，保证列序与冻结规格一致）
        pos = 0                                   # 0-based 指向当前列
        for name in LS.CASE_HEADERS:
            if pos < len(cur) and cur[pos] == name:
                pos += 1
                continue
            if name in cur:                       # 存在但位置不对 → 靠后续对齐处理
                pos += 1
                continue
            ws.insert_cols(pos + 1)               # openpyxl 是 1-based
            ws.cell(row=1, column=pos + 1).value = name
            cur.insert(pos, name)
            ins.append(name)
            dv = _MIGRATE_DEFAULT.get(name, "")
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=pos + 1).value = dv
            pos += 1
        # --- 2) 删多余列（从右往左删，避免下标漂移）
        for i in range(len(cur) - 1, -1, -1):
            if cur[i] not in LS.CASE_HEADERS:
                drop.append(cur[i])
                ws.delete_cols(i + 1)
                cur.pop(i)
        # --- 3) 重套样式
        for i, name in enumerate(LS.CASE_HEADERS, start=1):
            c = ws.cell(row=1, column=i)
            c.value = name
            LS.header_cell(c)
            ws.column_dimensions[get_column_letter(i)].width = LS.CASE_WIDTHS[i - 1]
        ws.freeze_panes = LS.FREEZE
        n = LS.renumber(ws)
        print("已迁移到 v3：新增列 %s；删除旧列 %s；数据行 %d" % (ins or "无", drop or "无", n))
        if drop:
            print("  ⚠ 被删列的内容已丢失（如需保留请先备份台账）：%s" % "、".join(drop))

    # --- 4) 法规清单：列名映射重建（旧 13 列 → 7 列）
    if LS.SHEET_LAW in wb.sheetnames:
        lw = wb[LS.SHEET_LAW]
        lh = [str(c.value).strip() if c.value is not None else "" for c in lw[1]]
        if lh != LS.LAW_HEADERS:
            idx = {n: (lh.index(n) if n in lh else -1) for n in LS.LAW_HEADERS}
            old_rows = [list(r) for r in lw.iter_rows(min_row=2, values_only=True)]
            wb.remove(lw)
            lw = LS.ensure_law_sheet(wb)
            for r in old_rows:
                if not any(v is not None and str(v).strip() for v in r):
                    continue
                lw.append([(r[idx[n]] if idx[n] >= 0 and idx[n] < len(r) else None)
                           for n in LS.LAW_HEADERS])
            for i in range(1, len(LS.LAW_HEADERS) + 1):
                lw.cell(row=1, column=i).value = LS.LAW_HEADERS[i - 1]
                LS.header_cell(lw.cell(row=1, column=i))
                lw.column_dimensions[get_column_letter(i)].width = LS.LAW_WIDTHS[i - 1]
            print("法规清单已从 %d 列重建为 %d 列（保留 %d 行）" % (len(lh), len(LS.LAW_HEADERS), len(old_rows)))
    else:
        LS.ensure_law_sheet(wb)
        print("已补建「%s」工作表" % LS.SHEET_LAW)

    wb.save(root / LS.LEDGER_NAME)
    return 0


def cmd_verify(root: Path) -> int:
    from openpyxl.utils import get_column_letter

    wb, ws = _open(root)
    ok = True

    def chk(cond, msg):
        nonlocal ok
        print(("  OK   " if cond else "  FAIL ") + msg)
        if not cond:
            ok = False

    print(f"工作表: {wb.sheetnames}")
    chk(ws.title == LS.SHEET_CASE, f"主工作表名 = {LS.SHEET_CASE}")
    chk([c.value for c in ws[1]] == LS.CASE_HEADERS, "表头与冻结规格 v2 一致")
    chk(ws.max_column == len(LS.CASE_HEADERS), f"列数 = {len(LS.CASE_HEADERS)}")
    widths = [ws.column_dimensions[get_column_letter(i)].width for i in range(1, len(LS.CASE_HEADERS) + 1)]
    chk(widths == [float(w) for w in LS.CASE_WIDTHS], f"列宽 = {LS.CASE_WIDTHS}")
    h = ws.cell(row=1, column=1)
    chk(h.fill.fgColor.rgb.endswith(LS.HEADER_FILL) and h.font.bold,
        f"表头配色 = #{LS.HEADER_FILL} 且加粗")
    chk(ws.freeze_panes == LS.FREEZE, f"冻结窗格 = {LS.FREEZE}")
    chk(LS.SHEET_LAW in wb.sheetnames, f"存在「{LS.SHEET_LAW}」工作表")
    # 只数 B..N 有内容的行，不信 max_row-1（尾部幽灵行）
    n_case = sum(1 for r in ws.iter_rows(min_row=2, values_only=True)
                 if any(v is not None and str(v).strip() for v in r[1:]))
    print(f"案例行数: {n_case}（max_row-1={ws.max_row - 1}，差值即尾部空行）")
    if LS.SHEET_LAW in wb.sheetnames:
        lw = wb[LS.SHEET_LAW]
        n_law = sum(1 for r in lw.iter_rows(min_row=2, values_only=True)
                    if any(v is not None and str(v).strip() for v in r))
        print(f"法规行数: {n_law}")
    print("结论:", "格式未被改动，可继续增量追加" if ok else "格式已被改动，需按 ledger_style.py 修复后再追加")
    return 0 if ok else 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="案例库根路径（含台账的文件夹）")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("check")
    p.add_argument("--key", required=True)
    p = sub.add_parser("add")
    p.add_argument("--json", required=True)
    p = sub.add_parser("add-batch")
    p.add_argument("--file", required=True)
    p = sub.add_parser("latest")
    p.add_argument("--source", default="")
    p = sub.add_parser("law-sheet")
    p.add_argument("--file", required=True)
    sub.add_parser("sort")
    sub.add_parser("verify")
    sub.add_parser("migrate", help="把旧格式台账就地升级到 v3 冻结规格（保留数据行）")
    args = ap.parse_args()
    root = Path(args.root)

    if args.cmd == "check":
        return cmd_check(root, args.key)
    if args.cmd == "add":
        return cmd_add(root, json.loads(args.json))
    if args.cmd == "add-batch":
        return cmd_add_batch(root, args.file)
    if args.cmd == "latest":
        return cmd_latest(root, args.source)
    if args.cmd == "law-sheet":
        return cmd_law_sheet(root, args.file)
    if args.cmd == "sort":
        return cmd_sort(root)
    if args.cmd == "verify":
        return cmd_verify(root)
    if args.cmd == "migrate":
        return cmd_migrate(root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
