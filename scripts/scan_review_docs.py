#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""scan_review_docs.py —— 扫描「复核决定书」并核对文件名语义（2026-11-02）

为什么需要：`label_review_decisions.py` 的闸门是「台账申辩列含【复核决定」，
但复核决定书是**扫描件**、台账最初又没写复核内容 → 闸门自锁，复核决定永远标不出来。
实测：汕头宜华·肖伟昊、上海世联行·袁鸿昌、上海通赛·李承恩 三案都是复核案，
文件名却一个带日期、一个不带，且无（原决定）/（复核决定）标注。

做法：逐个 PDF 取首页文本（有文字层直接读，扫描件 OCR 首页），
按「复核决定书 / 复核申请 / 复核意见」判定，输出 JSON 清单：
    {"file": 路径, "kind": "复核"|"原决定"|"未知", "wenhao": 文号, "party": 当事人}

用法：python scripts/scan_review_docs.py --root <库根> [--dir 01_xxx] [--jobs 6] [--out f.json]
"""
import argparse
import json
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ocrlib  # noqa: E402

REVIEW_PAT = re.compile(r"复核决定书|复核申请|复核意见|复核（\d{4}）")
WENHAO = re.compile(r"中基协(复核|处分)[（(【\[]?(\d{4})[）)】\]]?(\d+)号")
PARTY = re.compile(r"(?:申请人|当事人)[：:]\s*([\u4e00-\u9fa5·]{2,12})")


def head_text(path):
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    d = fitz.open(path)
    try:
        t = d[0].get_text()
        if len(re.sub(r"\s", "", t)) >= 60:
            return re.sub(r"\s+", " ", t), "text"
        del t
        return re.sub(r"\s+", " ", ocrlib.ocr_text(path, pages=[0]) or ""), "ocr"
    finally:
        d.close()


def judge(t):
    if not t:
        return "未知", "", ""
    m = WENHAO.search(t.replace(" ", ""))
    wen = ("%s%s" % (m.group(1), m.group(3))) if m else ""
    p = PARTY.search(t)
    party = p.group(1) if p else ""
    if "复核决定书" in t or (REVIEW_PAT.search(t) and "复核申请" in t):
        return "复核", wen, party
    if "纪律处分决定书" in t or "处分决定书" in t:
        return "原决定", wen, party
    return "未知", wen, party


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--dir", default="01_")
    ap.add_argument("--jobs", type=int, default=6)
    ap.add_argument("--out", default="")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    sub = [d for d in os.listdir(root) if d.startswith(a.dir)]
    if not sub:
        sys.exit("找不到目录 %s*" % a.dir)
    files = []
    for dp, _dn, fns in os.walk(os.path.join(root, sub[0])):
        for f in fns:
            if f.lower().endswith(".pdf"):
                files.append(os.path.join(dp, f))
    files.sort()
    print("待扫描 %d 份" % len(files))

    def work(p):
        try:
            t, how = head_text(p)
        except Exception as e:  # noqa: BLE001
            return {"file": p, "kind": "错误", "err": str(e)}
        kind, wen, party = judge(t)
        return {"file": p, "kind": kind, "wenhao": wen, "party": party,
                "how": how, "head": t[:120]}

    out = []
    with ThreadPoolExecutor(max_workers=a.jobs) as ex:
        for i, r in enumerate(ex.map(work, files), 1):
            out.append(r)
            if i % 25 == 0 or i == len(files):
                print("  ...%d/%d" % (i, len(files)))
    rv = [x for x in out if x.get("kind") == "复核"]
    print("\n识别为复核决定书：%d 份" % len(rv))
    for x in rv:
        print("  %s ｜ %s ｜ %s" % (os.path.basename(x["file"])[:56], x.get("wenhao"), x.get("party")))
    unk = [x for x in out if x.get("kind") in ("未知", "错误")]
    if unk:
        print("\n未能判定：%d 份" % len(unk))
        for x in unk[:10]:
            print("  %s %s" % (os.path.basename(x["file"])[:50], x.get("err") or x.get("head", "")[:40]))
    if a.out:
        with open(a.out, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False, indent=1)
        print("\n清单已写：%s" % a.out)


if __name__ == "__main__":
    main()
