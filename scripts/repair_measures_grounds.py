# -*- coding: utf-8 -*-
"""台账内容层修复：处罚措施按主体分行 + 复核决定空行另起 + 核心事由去残头。

用户口径（2026-11-02）：
  1) 处罚措施：每个「管理人：…」「高管（姓名）：…」各占一行；多高管相邻；
     高管组/管理人组按原文先后排序。
  2) 复核决定：在处罚措施/申辩理由里空一行另起一段。
  3) 核心处罚事由：删除末尾断句残头（「。上述行为」「。 上述情形」等）。

用法：
  python repair_measures_grounds.py --root <库根>            # 干跑，只报告
  python repair_measures_grounds.py --root <库根> --apply    # 落盘（自动备份台账）

设计：幂等（可重复跑）；路径一律从库根派生，备份落 <库根父>/_案例库维护留痕/<库名>/_work/。
"""
import os
import re
import sys
import shutil
import argparse
from datetime import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import ledger_style as LS  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

# 1-based 列号（与 ledger_style.COL 对齐：measures=4→5、grounds=5→6、defense=8→9）
C_MEAS = LS.COL["measures"] + 1
C_GROUNDS = LS.COL["grounds"] + 1
C_DEF = LS.COL["defense"] + 1
C_SUBJ = LS.COL["subject"] + 1

TYPOS = [("公开遣责", "公开谴责")]  # 明显 OCR/抽取错字

# 管理人措施误取「权利告知句」（实测 15 例，深圳证监局 2026 版式：正文只有
# 「我局决定对 A、B 采取出具警示函的行政监管措施」，没有分主体的「对你公司…」句，
# 抽取器把紧随其后的「如对本行政监管措施不服…」当成了管理人措施）。
_BOILER = re.compile(r'^管理人[：:]\s*如对本?[^；\n]*?(?:不服|复议|诉讼)')
_DECIDE = re.compile(r'对(.+?)(?:及|、)([^、及]+?)采取(.+?)(?:的行政监管措施|的监督管理措施)')


def _norm_name(s):
    return re.sub(r'[\s（）()]', '', str(s or ''))


def fix_boilerplate_measure(meas_text, subject):
    """管理人措施=权利告知句时，从高管措施句反推管理人措施。返回 (new_text, org)。无改动返回 (None, None)。"""
    if not meas_text or "不服" not in str(meas_text):
        return None, None
    lines = str(meas_text).split("\n")
    hit = [i for i, ln in enumerate(lines) if _BOILER.match(ln.strip())]
    if not hit:
        return None, None
    subj = str(subject or "")
    org = subj.split("；")[0].strip()
    persons = set()
    if "；" in subj:
        persons |= {p.strip() for p in re.split(r"[；、]", subj.split("；", 1)[1]) if p.strip()}
    for ln in lines:
        if "采取" not in ln:
            continue
        g = re.search(r"高管（([^）]*)）", ln)
        if g:
            persons.add(g.group(1).strip())
        i, j = ln.find("对"), ln.find("采取")
        if i < 0 or j < 0 or j <= i:
            continue
        head = ln[i + 1:j]
        # 从尾部逐个剥掉人名，剩下的是机构全称（兼容「A、B、C采取」与「A及B采取」）
        seg = [x for x in re.split(r"[、及]", head) if x.strip()]
        while len(seg) > 1 and seg[-1].strip() in persons:
            seg.pop()
        org_full = "、".join(seg).strip()
        if org and org_full and _norm_name(org_full).startswith(_norm_name(org)):
            tail = ln[j + 2:]
            mm = re.search(r"^(.*?)的(?:行政监管措施|监督管理措施)(.*)$", tail)
            action, rest = (mm.group(1), mm.group(2)) if mm else (tail, "")
            new_line = "管理人：对%s采取%s的行政监管措施%s" % (org_full, action, rest)
            lines[hit[0]] = new_line
            return "\n".join(lines), org
    return None, org  # 有缺陷但无法安全反推 → 交人工

# 兜底诊断：删完仍以残头词结尾（严格正则未覆盖 → 需人工看）
_FRAG_WORD_END = re.compile(
    r'(?:上述|以上|以下)(?:行为|情形|事实|事项|问题|内容|违规行为|违法行为|违法事实|违规事实)[。；;]?\s*$'
)


def fix_typos(s):
    if not s:
        return s
    for a, b in TYPOS:
        s = s.replace(a, b)
    return s


def _charset(s):
    """忽略空白与分隔符后的字符多重集合（用于守恒断言：只允许重排/换行，不允许删字）。"""
    return sorted(re.sub(r"[\s；]", "", str(s or "")))


def main(argv=None):
    ap = argparse.ArgumentParser(description="台账内容层修复：处罚措施分行 / 复核另起 / 事由去残头")
    ap.add_argument("--root", required=True, help="案例库根目录（含 04_综合索引台账【总表】.xlsx）")
    ap.add_argument("--apply", action="store_true", help="落盘（默认仅干跑）")
    a = ap.parse_args(argv)

    root = os.path.abspath(a.root)
    ledger = os.path.join(root, LS.LEDGER_NAME)
    if not os.path.isfile(ledger):
        print("!! 未找到台账：", ledger)
        return 2

    wb = load_workbook(ledger, data_only=False)
    ws = wb[LS.SHEET_CASE]

    meas_chg, grounds_chg, def_chg = [], [], []
    bol_chg, bol_manual = [], []
    residual = []
    typo_n = 0

    for r in range(2, ws.max_row + 1):
        cell_m = ws.cell(r, C_MEAS)
        cell_g = ws.cell(r, C_GROUNDS)
        cell_d = ws.cell(r, C_DEF)
        old_m, old_g, old_d = cell_m.value, cell_g.value, cell_d.value

        # --- 处罚措施：先修「管理人措施误取权利告知句」，再分行排序 ---
        bol_m, bol_org = fix_boilerplate_measure(old_m, ws.cell(r, C_SUBJ).value)
        if bol_m:
            bol_chg.append((r, old_m, bol_m))
            old_m = bol_m
        elif bol_org:
            bol_manual.append((r, bol_org))

        # --- 处罚措施：分行 + 排序（先对纯重排结果做守恒断言，再单独修错字）---
        base_m = LS.strip_page_noise(old_m)  # 页码噪声属删除项，不参与守恒比对
        norm_m = LS.normalize_measures_text(base_m)
        if norm_m != base_m:
            assert _charset(norm_m) == _charset(base_m), f"行{r} 处罚措施内容不守恒，已中止"
        new_m = fix_typos(norm_m)
        if new_m != cell_m.value:
            meas_chg.append((r, cell_m.value, new_m))
            cell_m.value = new_m
            cell_m.alignment = Alignment(vertical="top", wrap_text=True)

        # --- 核心处罚事由：去页码噪声 + 去残头 + 序号列表化（一是→一、、每条一行）---
        new_g = fix_typos(LS.normalize_grounds_items(
            LS.clean_grounds_tail(LS.strip_page_noise(old_g))))
        if new_g != old_g:
            grounds_chg.append((r, old_g, new_g))
            cell_g.value = new_g
            cell_g.alignment = Alignment(vertical="top", wrap_text=True)
        if new_g and _FRAG_WORD_END.search(str(new_g)):
            residual.append((r, new_g))

        # --- 申辩理由：复核决定空行另起 ---
        if old_d and "【复核决定" in str(old_d) and "\n\n【复核决定" not in str(old_d):
            s = str(old_d)
            i = s.find("【复核决定")
            new_d = fix_typos(s[:i].rstrip() + "\n\n" + s[i:])
            def_chg.append((r, old_d, new_d))
            cell_d.value = new_d
            cell_d.alignment = Alignment(vertical="top", wrap_text=True)

        # 错字计数（只统计，不重复改）
        for old in (old_m, old_g, old_d):
            if old and any(x in str(old) for x, _ in TYPOS):
                typo_n += 1

    print("=" * 70)
    print(f"台账：{ledger}")
    print(f"管理人措施误取权利告知句（已反推修复）：{len(bol_chg)} 行")
    print(f"处罚措施需按主体分行/排序：{len(meas_chg)} 行")
    print(f"核心处罚事由去残头/列表化：{len(grounds_chg)} 行")
    print(f"申辩理由复核决定需空行：  {len(def_chg)} 行")
    print(f"含明显错字的单元格：      {typo_n} 个")
    print(f"兜底：删后仍以残头词结尾：{len(residual)} 行（应为 0；非 0 需人工看）")
    print("=" * 70)

    def show(title, items, n=4, lim=150):
        if not items:
            return
        print(f"\n--- {title} 抽样（前 {min(n, len(items))} 条）---")
        for it in items[:n]:
            print(f"\n行{it[0]}")
            print("  原:", repr(str(it[1])[:lim]))
            print("  新:", repr(str(it[2])[:lim]))

    show("管理人措施 反推修复", bol_chg, n=3, lim=120)
    show("处罚措施 分行/排序", meas_chg, n=2)
    show("核心处罚事由 去残头", grounds_chg, lim=110)
    show("申辩理由 复核另起", def_chg, n=2, lim=90)

    if bol_manual:
        print("\n!! 以下行管理人措施=权利告知句，且无法安全反推（需人工按 PDF 修）：")
        for r, org in bol_manual:
            print(f"  行{r}: {org}")

    if residual:
        print("\n!! 以下行核心事由删后仍以残头词结尾（需人工确认）：")
        for r, g in residual[:20]:
            print(f"  行{r}: …{str(g)[-50:]!r}")

    if not a.apply:
        print("\n[干跑] 未落盘。确认无误后加 --apply。")
        return 0 if not residual else 1

    wd = LS.work_dir(root)
    os.makedirs(wd, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = os.path.join(wd, f"ledger_bak_measures_grounds_{ts}.xlsx")
    shutil.copy2(ledger, bak)
    print("\n已备份台账 →", bak)

    wb.save(ledger)
    print("已落盘：", ledger)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
