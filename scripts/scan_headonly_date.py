# -*- coding: utf-8 -*-
"""补充扫描：台账「决定日期」是否其实是挂网日/页头日期，而非决定书文末落款。

背景（2026-09-29 首版 / 2026-10-09 重大修正）：
B4 做**全文**比对，而证监局页面头部必然印有「发文日期 YYYY年M月D日」，台账若复制了挂网日，
B4 全文比照样命中 → 漏检。本脚本专治这个盲区。

> **2026-10-09 修正（0.4.35）：首版把「正文 <2400 字」的 PDF 直接跳过，是致命设计错误。**
> 证监局网页打印件的正文只有 700~1500 字，头尾区间必然重叠，首版据此 `continue` → **恒报 0**，
> 而实际上 163 条台账日期全是挂网日（滞后落款 1~59 天）。落款提取必须换用**锚点法**，不能靠头尾切分。

判定方法（改用人眼复核过的三段式落款提取，见 `sign_date()`）：
  1. **锚点法**：`(?:证监局|监管局|管理局|基金业协会|交易所)` + 日期 —— 覆盖附件型正式文书与紧邻落款；
  2. **打印标记法**：取「不停止执行/提起诉讼」到「【打印】」之间的最后一个日期
     —— 覆盖网页打印件里被页脚噪声（链接：中国政府网 / 联系我们|法律声明|）隔开的落款；
  3. **尾部兜底**：末 1500 字最后一个日期，且前 8 字不得含「发文日期/日期」（那是挂网日）。

台账日期 ≠ 提取落款 即列为疑似（分「落款更早」与「落款更晚（异常）」两类）。

用法：
  python scan_headonly_date.py --root <案例库根> [--json out.json] [--selftest]
只读，不修改任何文件。无文字层的扫描件自动跳过（那批走 B4 的 OCR 通道）。
"""
import os
import re
import sys
import json
import argparse

if hasattr(sys.stdout, "buffer"):
    sys.stdout = __import__("io").TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ANCHOR = re.compile(
    r"(?:证监局|监管局|管理局|基金业协会|交易所)\s*"
    r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
HEADPUB = re.compile(r"发文日期\s*(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
ANYDATE = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
PRINT_SEG = re.compile(r"(?:不停止执行|提起诉讼|诉讼期间)(.{0,400}?)【打印】", re.S)


def _fmt(a, b, c):
    return "%04d-%02d-%02d" % (int(a), int(b), int(c))


def sign_date(txt):
    """返回 (落款日期, 页头挂网日)；取不到落款时返回 ('', pub)。"""
    n = re.sub(r"\s+", "", txt or "")
    m = HEADPUB.search(n)
    pub = _fmt(*m.groups()) if m else ""
    a = [_fmt(*g) for g in ANCHOR.findall(n)]
    if a:
        return a[-1], pub
    seg = None
    for seg in PRINT_SEG.finditer(n):
        pass
    if seg is not None:
        d = [_fmt(*g) for g in ANYDATE.findall(seg.group(1))]
        if d:
            return d[-1], pub
    tail = n[-1500:]
    out = []
    for mm in ANYDATE.finditer(tail):
        pre = tail[max(0, mm.start() - 8): mm.start()]
        if "发文日期" in pre or pre.endswith("日期"):
            continue
        out.append(_fmt(*mm.groups()))
    return (out[-1] if out else ""), pub


# ---------------- 自检（正常 / 边界 / 噪声 / 反例） ----------------
def selftest():
    cases = [
        ("A-紧邻",    "…浙江证监局2026年7月30日【打印】", "2026-07-30"),
        ("A-页脚噪声", "…不停止执行。中国证券监督管理委员会山西监管局链接：中国政府网…联系我们|法律声明|2026年9月15日【打印】", "2026-09-15"),
        ("A-头部挂网日干扰", "发文日期2026年09月24日…不停止执行。…法律声明|2026年9月15日【打印】", "2026-09-15"),
        ("B-附件型",   "…广东证监局2026年7月13日抄送：证监会市场二司。广东证监局办公室2026年7月15日印发", "2026-07-13"),
        ("C-交易所",   "…深圳证券交易所2026年7月29日", "2026-07-29"),
        ("D-中基协",   "…中国证券投资基金业协会2026年4月30日", "2026-04-30"),
        ("E-多空格",   "…厦门证监局 2026 年 7 月 21 日【打印】", "2026-07-21"),
        ("反例-无落款", "全文没有任何日期", ""),
        ("反例-只有挂网日", "发文日期2026年10月09日名称关于对某某采取出具警示函措施的决定", ""),
    ]
    ok = True
    for name, s, exp in cases:
        got, _ = sign_date(s)
        good = got == exp
        ok = ok and good
        print("   [%s] %-18s exp=%-12s got=%s"
              % ("OK  " if good else "FAIL", name, exp or "(空)", got or "(空)"))
    print("自检", "全部通过" if ok else "存在失败用例")
    return ok


def load_rows(xlsx):
    """v2 列布局（2026-10-26）：B=1 处罚时间 | C=2 挂网发布日期 | D=3 管理人名称
    | K=10 案例文件（多值「；」分隔）| L=11 存放路径（与 K 一一对应）。
    K/L 展开成逐文件行，一个案例行产出 N 个待扫文件。"""
    import openpyxl
    ws = openpyxl.load_workbook(xlsx).active
    out = []
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=14, values_only=True), start=2):
        fns = [x.strip() for x in str(r[10] or "").replace(";", "；").split("；") if x.strip()]
        pts = [x.strip() for x in str(r[11] or "").replace(";", "；").split("；") if x.strip()]
        for j, fn in enumerate(fns):
            out.append({"row": i, "date": str(r[1] or ""), "source": str(r[9] or ""),
                        "target": str(r[3] or ""), "doc_no": "",
                        "filename": fn, "path": pts[j] if j < len(pts) else ""})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--json", default="")
    ap.add_argument("--selftest", action="store_true", help="只跑内置单元测试")
    a = ap.parse_args()

    if a.selftest:
        return 0 if selftest() else 1

    xlsx = os.path.join(a.root, "04_综合索引台账【总表】.xlsx")
    rows = load_rows(xlsx)
    try:
        import fitz
    except ImportError:
        import pymupdf as fitz

    cand, notext, nofile, short = [], 0, 0, 0
    for r in rows:
        if not re.match(r"\d{4}-\d{2}-\d{2}", r["date"]) or not r["filename"] or not r["path"]:
            continue
        p = os.path.join(a.root, r["path"].replace("/", os.sep), r["filename"])
        if not os.path.exists(p):
            nofile += 1
            continue
        try:
            d = fitz.open(p)
            raw = "".join(pg.get_text() for pg in d)
            d.close()
        except Exception:
            continue
        if len(raw.strip()) < 200:
            notext += 1
            continue
        sign, pub = sign_date(raw)
        if not sign:
            short += 1
            continue
        if sign == r["date"]:
            continue
        cand.append({"row": r["row"], "ledger_date": r["date"], "sign_date": sign,
                     "head_pub": pub, "source": r["source"], "target": r["target"],
                     "doc_no": r["doc_no"], "filename": r["filename"],
                     "kind": "落款更早" if sign < r["date"] else "落款更晚(异常)"})

    print(f"扫描台账 {len(rows)} 行 ｜ 缺文件 {nofile} ｜ 无文字层(跳过) {notext} ｜ 取不到落款 {short}")
    print(f"疑似「挂网日冒充落款」{len(cand)} 条（不进 problems，需人眼确认后批量改）：")
    for c in cand[:60]:
        print(f"  行{c['row']} 台账 {c['ledger_date']} → 落款 {c['sign_date']} [{c['kind']}] ｜ "
              f"{c['source']} ｜ {c['target'][:20]} ｜ {c['filename'][:40]}")
    if a.json:
        json.dump({"candidates": cand, "scanned": len(rows), "no_text_layer": notext,
                   "no_sign": short},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("报告已写入", a.json)
    return 0


if __name__ == "__main__":
    sys.exit(main())
