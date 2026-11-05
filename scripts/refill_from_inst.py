#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""refill_from_inst.py —— 用机构文书置换台账各列内容（2026-11-02）

用户口径（2026-11-02）：
  · 台账除「处罚措施（机构+人员）」外，其余各列**只从机构文书取**；
  · 一个案例的机构文书**只有一个**；
  · 但有的案例**人员文书先挂网**，那就先按人员文书填；**后来机构文书出现了，要置换**。

  → 所以本工具只干一件事：找出「有机构文书、但列内容还是人员文书/网页摘要那套」的行，
    用机构文书重抽覆盖。没有机构文书的行一律不动（那是合规的，登记待补即可）。

判定（复用 check_fields_source）：
  每列按「条」切分，逐条在机构文书中找最长公共子串；覆盖率 < 60% 判为「不是从机构文书来的」。

保护（宁可不动，不可改错）：
  · 重抽为空 / 归一后 < 20 字 → 保留旧值
  · 重抽含「我局决定/现决定/违反了」等越界句 → 保留旧值
  · I 列（申辩）机构文书常有「未提出陈述申辩」→ 重抽为空属正常，保留旧值
  · 任何改动都打印旧新对照，可逐条复核

用法：
  python scripts/refill_from_inst.py --root <库根>              # 干跑
  python scripts/refill_from_inst.py --root <库根> --apply      # 落盘
  python scripts/refill_from_inst.py --root <库根> --rows 86,54 # 只看指定行
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import extract_core as EC  # noqa: E402
import ledger_style as LS  # noqa: E402
from check_fields_source import COLS_F, MATCH_MIN, col_cover, pick_inst_file
from audit_cells import FOOTER  # noqa: E402
from ledger_bak import backup_ledger  # noqa: E402
from probe_grounds import pdf_text_join, text_layer_ok  # noqa: E402

OVERREACH = re.compile(r"我局决定|本局决定|现决定|免予处罚|应当自收到本决定书|"
                       r"如不服|复议|诉讼")
COL2KEY = {6: "grounds", 7: "evidence", 8: "legal_basis", 9: "defense"}
PUNC = re.compile(r"[\s，。；：、,.;:（）()《》“”\"'…—–\-]")


def _sig(s):
    """去标点后的骨架：只差标点的改动不算实质改动（原文半角 vs 台账全角）。"""
    return PUNC.sub("", str(s))


def _to_full_punc(s):
    """中文语境半角标点 → 全角（与 fix_cells_20261102 同口径；数字间的逗号/冒号保持）。"""
    def rp(m):
        ch = m.group(0)
        i = m.start()
        a = s[i - 1] if i else ""
        b = s[i + 1] if i + 1 < len(s) else ""
        if a.isdigit() and b.isdigit():
            return ch
        return {",": "，", ";": "；", ":": "："}.get(ch, ch)
    return re.sub(r"[,;:]", rp, str(s))


def read_text(fp, use_ocr=True, cache_dir=None):
    t = pdf_text_join(fp) if text_layer_ok(fp) else ""
    if len(t) < 300 and use_ocr:
        try:
            import ocrlib
            t = ocrlib.ocr_text(fp, cache_dir=cache_dir)
        except Exception:
            t = ""
    return t.replace("\r", " ")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--rows", default="")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--show", type=int, default=10)
    ap.add_argument("--no-ocr", action="store_true")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]
    only = {int(x) for x in args.rows.split(",") if x.strip()} if args.rows else None
    # OCR 缓存放库外（库根只允许白名单条目，C10 会报）；路径从库根派生，保证可迁移
    cache_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")

    n_row = n_cell = n_skip = 0
    shown = 0
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        if only and r not in only:
            continue
        source = str(ws.cell(r, 11).value or "")
        if "交易所" in source:
            continue                        # 交易所规则名由映射得出，不适用回现校验
        names = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        inst = pick_inst_file(names)
        if not inst:
            continue                        # 只有人员文书 → 按人员填是对的
        d = str(ws.cell(r, 13).value or "").split("；")[0].strip()
        fp = os.path.join(root, d, inst)
        if not os.path.isfile(fp):
            continue
        txt = read_text(fp, use_ocr=not args.no_ocr, cache_dir=cache_dir)
        if len(txt) < 300:
            n_skip += 1
            continue

        low = []
        for c, cn in COLS_F:
            v = ws.cell(r, c).value
            # 只置换 F（事由）。H/G/I 不参与自动置换：
            #   实测行265 —— OCR 把《关于加强**私募**投资基金监管的若干规定》读成
            #   《关于加强 私募投资基金监管的若**于**规定》，法条名反而被写坏。
            #   H 列的重复/漏抽用 fix_h_dedup / fix_h_basis 定点处理，不做自动覆盖。
            if c != 6:
                continue
            if not isinstance(v, str) or len(v.strip()) < 12:
                continue
            # 页脚串味必须**强制重抽**：页脚文字本身就在这份文书里，覆盖率会是 100%，
            # 靠覆盖率永远抓不到（行312/314/365 就是这么漏的）。
            if FOOTER.search(v) or col_cover(v, txt) < MATCH_MIN:
                low.append((c, cn))
        if not low:
            continue

        try:
            e = EC.extract_amac(txt) if source == "中基协" else EC.extract_csrc(txt)
        except Exception as ex:
            print("行%-4d 抽取异常：%s" % (r, ex))
            n_skip += 1
            continue
        if not e:
            n_skip += 1
            continue

        subject = str(ws.cell(r, 4).value or "")
        changed = []
        for c, cn in low:
            old = ws.cell(r, c).value
            key = COL2KEY[c]
            raw = (e.get(key) or "").strip()
            new = raw
            if key == "grounds":
                new = LS.normalize_grounds(new, subject, split_parallel=True)
            elif key == "legal_basis":
                new = LS.normalize_grounds_items(new)
            new = _to_full_punc(new)
            if len(new) < 20:
                continue
            # OCR 断词：中文之间夹空格（「关于加强 私募投资基金监管」）→ 抽取文本不可信，拒绝
            if re.search(r"[\u4e00-\u9fa5][ \t\u3000]+[\u4e00-\u9fa5]", new):
                continue
            if OVERREACH.search(new[:80]):
                continue
            # 置换后必须能在机构文书里回现（否则说明抽取又跑偏了）。
            # 注意用**归一前**的 raw 回现：normalize 会把「你公司」换成主体全称，
            # 换完再拿回原文比对必然对不上（行80 山东神光第一版就是被这条误拦的）。
            if col_cover(raw, txt) < 0.8:
                continue
            if _sig(new) == _sig(old or ""):     # 只差标点 → 不是实质改动
                continue
            changed.append((c, cn, old, new))

        if not changed:
            continue
        n_row += 1
        n_cell += len(changed)
        if shown < args.show:
            print("\n行%-4d %s（机构文书 %s）" % (r, subject[:26], inst))
            for c, cn, old, new in changed:
                print("  %s\n    旧 %s\n    新 %s" % (cn, str(old)[:90].replace("\n", " ⏎ "),
                                                 new[:90].replace("\n", " ⏎ ")))
            shown += 1
        if args.apply:
            for c, cn, old, new in changed:
                ws.cell(r, c).value = new

    print("\n置换 %d 行 / %d 格｜跳过 %d 行（读不到原文）" % (n_row, n_cell, n_skip))
    if args.apply and n_row:
        backup_ledger(root, ledger, "refill_inst_20261102")
        wb.save(ledger)
        print("已保存（备份在 _案例库维护留痕/台账备份/）")
    elif not args.apply:
        print("[干跑] 未落盘；加 --apply 执行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
