# -*- coding: utf-8 -*-
"""平台对账（A6 的轻量替代）：只拉列表 → 与库内/不收录清单比对 → 可选逐条取正文复判。

为什么需要（2026-10-28 实测背景）：
  A6 `--online` 要逐局联网重拉**并抽正文**，一轮十几分钟，日常跑不起；
  但上一轮一年窗口建库确实漏了 12 条（广东局 9 条正文被静默丢弃 + 2 条交易所 + 1 条中基协），
  说明"没跑 A6"就等于"不知道完整不完整"。
  本脚本把这件事压到可日常执行：
    第 1 段（默认）只翻列表，约 1.5 分钟，按 URL / 主体名与库内 + 不收录清单比对；
    第 2 段（--body）对"列表有、库里没有"的条目逐条取正文，用 is_private_fund 复判，
    把"标题不含私募字样但正文是私募案"的盲区也覆盖（广东局 8 条就是这样找出来的）。

用法：
  python recon_platform.py --root "<案例库根>" --start 2025-10-28 --end 2026-10-28
  python recon_platform.py --root "<案例库根>" --start 2025-10-28 --end 2026-10-28 --body
  # 退出码 0=PASS（无真缺口）；1=存在真缺口（列表有、库内无、正文判为私募）
"""
import argparse
import json
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import openpyxl                     # noqa: E402
import fetch_csrc as FC             # noqa: E402
from classify_filter import is_private_fund  # noqa: E402
# P3 名册闸门：与 build 管道统一走 decide（文本判定 + AMAC 管理人名册复核）
from amac_lookup import decide, AmacRegistry, DEFAULT_DIR as _AMAC_REG_DIR  # noqa: E402

KEY = re.compile(r"私募|基金管理人|股权投资基金|创投基金|基金业协会")


def norm_u(u):
    s = str(u or "").strip()
    s = re.sub(r"^https?://", "", s)
    s = re.sub(r"^www\.", "", s)
    return s.split("?")[0].split("#")[0].rstrip("/").lower()


def load_lib(root):
    p = os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb[wb.sheetnames[0]]
    urls, names = set(), []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not any(x not in (None, "") for x in r[1:]):
            continue
        for u in str(r[13] or "").split("；"):
            if norm_u(u):
                urls.add(norm_u(u))
        nm = re.sub(r"\s+", "", str(r[3] or "").split("；")[0])
        if nm:
            names.append(nm)
    md = os.path.join(root, "不收录清单.md")
    if os.path.exists(md):
        for line in open(md, encoding="utf-8"):
            if line.startswith("|"):
                for c in [x.strip() for x in line.strip().strip("|").split("|")]:
                    if c.startswith("http"):
                        urls.add(norm_u(c))
    return urls, names


def main():
    ap = argparse.ArgumentParser(description="平台列表 ↔ 库内 完整性对账")
    ap.add_argument("--root", required=True)
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    ap.add_argument("--body", action="store_true",
                    help="对未匹配条目取正文复判（更准，慢约 3~5 分钟）")
    ap.add_argument("--work", default="", help="中间产物目录，默认 <root>_recon")
    ap.add_argument("--json", default="", help="结果 JSON 路径")
    ap.add_argument("--registry-dir", default=_AMAC_REG_DIR,
                   help="AMAC 管理人名册目录；缺失则退化为纯文本判定")
    ap.add_argument("--no-registry", action="store_true", help="关闭 P3 名册闸门")
    a = ap.parse_args()

    registry = None
    if not a.no_registry:
        try:
            registry = AmacRegistry(os.path.abspath(a.registry_dir))
        except Exception:
            registry = None

    root = os.path.abspath(a.root)
    work = os.path.abspath(a.work or (root + "_recon"))
    os.makedirs(work, exist_ok=True)
    out = a.json or os.path.join(work, "recon_%s_%s.json" % (a.start, a.end))

    have, names = load_lib(root)
    print("库内 + 不收录清单 已知链接：%d 条" % len(have))

    items, chans = [], []
    for pinyin, cn, cid_meas, cid_pen in FC.BUREAUS:
        for col, cid in (("行政监管措施", cid_meas), ("行政处罚决定", cid_pen)):
            lst, meta = FC.fetch_list(cid, a.start, a.end)
            chans.append({"bureau": cn, "col": col, "items_window": len(lst), **meta})
            for it in lst:
                it["_bureau_cn"], it["_col"] = cn, col
                items.append(it)
    print("平台窗口条目：%d（翻页异常栏目 %d 个）"
          % (len(items), sum(1 for c in chans
                             if c.get("terminated_by") not in ("start_boundary", "short_page"))))

    matched, rest = 0, []
    for it in items:
        u = norm_u(it["url"])
        if u in have:
            matched += 1
            continue
        t = re.sub(r"\s+", "", it["title"])
        if any(n and len(n) >= 6 and n in t for n in names):
            matched += 1
            continue
        rest.append(it)

    gaps, errs, nonpriv = [], [], 0
    if a.body:
        print("逐条取正文复判 %d 条（--body）…" % len(rest))
        for it in rest:
            kind, payload, ext = FC.get_content(it["manuscriptId"], it["url"])
            text = payload if kind == "inline" else ""
            if kind == "attach":
                _fn, text = FC.save_attach(payload, ext, it["url"],
                                           os.path.join(work, "dl"),
                                           os.path.join(work, "ocr"))
            if not text or len(re.sub(r"\s", "", text)) < 60:
                errs.append({"bureau": it["_bureau_cn"], "title": it["title"],
                             "url": it["url"]})
                continue
            ok, why = decide(text, it["title"], "", registry)
            if ok:
                gaps.append({"bureau": it["_bureau_cn"], "col": it["_col"],
                             "date": it.get("date_published", ""), "title": it["title"],
                             "url": it["url"], "why": why,
                             "head": re.sub(r"\s+", " ", text)[:200]})
                print("  [真缺口] %s %s %s" % (it["_bureau_cn"],
                                              it.get("date_published"), it["title"][:44]))
            else:
                nonpriv += 1
            time.sleep(0.2)
    else:
        for it in rest:
            t = re.sub(r"\s+", "", it["title"])
            if KEY.search(t):
                gaps.append({"bureau": it["_bureau_cn"], "col": it["_col"],
                             "date": it.get("date_published", ""), "title": it["title"],
                             "url": it["url"], "why": "标题含私募/基金关键词"})
            else:
                nonpriv += 1

    res = {"window": [a.start, a.end], "platform_items": len(items),
           "matched": matched, "gaps": gaps, "unreadable": errs,
           "non_private_candidates": nonpriv, "channels": chans}
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print("\n==== 对账结果 ====")
    print("平台 %d ｜ 已解释 %d ｜ 真缺口 %d ｜ 非私募 %d ｜ 正文不可读 %d"
          % (len(items), matched, len(gaps), nonpriv, len(errs)))
    for g in gaps[:20]:
        print("   - %s %s %s" % (g["bureau"], g["date"], g["title"][:52]))
    print("明细：", out)
    return 0 if not gaps else 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.exit(main())
