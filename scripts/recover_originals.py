# -*- coding: utf-8 -*-
"""recover_originals.py — 把"文本重建"的垃圾决定书 PDF 还原为官网原决定书。

背景（2026-10-29 严重问题）：
    早期 fetch_csrc 对证监局内联 HTML 正文走文本重建（make_pdf_from_text），
    产物带「原文链接：」与「&ensp;」等 HTML 实体垃圾、丢失官方版式，常被误认为
    OCR 乱码。本脚本把这类 PDF 用无头浏览器把官网原页面打印为 PDF（html2pdf），
    还原为带官方导航栏/版式的"官网原决定书"。

做法：
    - 扫描 02（证监局）与 05（典型案例）下所有 PDF；
    - 仅处理"文本重建"产物（含「原文链接：」或「&ensp;」）且能抽出官网 URL 的；
    - 抽出 URL → 调 save_decision_pdf.py html2pdf 重新打印为 PDF；
    - 校验新 PDF 干净（无「原文链接：」/「&ensp;」且体积 > 10KB）后才替换原文件；
    - 原文件先备份到库父目录 _recover_backup/（只移不删，可恢复）；
    - 抽不到 URL / 打印失败的文件：原样保留并记入状态，供人工处理；
    - 状态文件 _recover_state.json 落库根，可断点续跑（已 done 的文件跳过）。

用法：
    python recover_originals.py --root "<库根>" [--limit N] [--delay 1.5] [--no-backup]
"""
import argparse
import importlib.util
import json
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SDP = os.path.join(HERE, "save_decision_pdf.py")
PY = sys.executable

URL_RE = re.compile(r"原文链接[:：]\s*(\S+)")
RECON_MARK = ("原文链接：", "&ensp;")


def load_sdp():
    spec = importlib.util.spec_from_file_location("save_decision_pdf", SDP)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def is_reconstructed(p):
    try:
        import fitz
        d = fitz.open(p)
        t = "".join(pg.get_text() for pg in d)
        d.close()
    except Exception:
        return False
    return any(mk in t for mk in RECON_MARK)


def extract_url(p):
    try:
        import fitz
        d = fitz.open(p)
        t = "".join(pg.get_text() for pg in d)
        d.close()
    except Exception:
        return None
    m = URL_RE.search(t)
    return m.group(1) if m else None


def is_clean(p):
    try:
        import fitz
        d = fitz.open(p)
        t = "".join(pg.get_text() for pg in d)
        d.close()
    except Exception:
        return False
    if any(mk in t for mk in RECON_MARK):
        return False
    return os.path.getsize(p) > 10 * 1024


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--limit", type=int, default=0, help="仅处理前 N 个（测试用，0=全部）")
    ap.add_argument("--delay", type=float, default=1.5, help="每次打印后间隔秒")
    ap.add_argument("--no-backup", action="store_true", help="不备份原文件（默认备份）")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    parent = os.path.dirname(root)
    backup_dir = os.path.join(parent, "_recover_backup")
    state_path = os.path.join(root, "_recover_state.json")
    os.makedirs(backup_dir, exist_ok=True)

    state = {}
    if os.path.exists(state_path):
        try:
            state = json.load(open(state_path, encoding="utf-8"))
        except Exception:
            state = {}

    sdp = load_sdp()
    browser = sdp.find_browser()

    # 收集待处理文件（02 + 05，仅文本重建且能抽 URL）
    targets = []
    for top in ("02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）",
                "05_重点典型案例"):
        base = os.path.join(root, top)
        if not os.path.isdir(base):
            continue
        for dp, dn, fn in os.walk(base):
            if any(x in dp for x in ["_recover", "_case_excluded", "_backup"]):
                continue
            for f in fn:
                if not f.lower().endswith(".pdf"):
                    continue
                rel = os.path.relpath(os.path.join(dp, f), root)
                if state.get(rel) == "done":
                    continue
                p = os.path.join(dp, f)
                if not is_reconstructed(p):
                    continue
                targets.append((rel, p))

    print("[recover] 待处理（文本重建）文件数：%d" % len(targets))
    if args.limit:
        targets = targets[:args.limit]
        print("[recover] --limit=%d，仅处理前 %d 个" % (args.limit, len(targets)))

    stats = {"done": 0, "no_url": 0, "fetch_fail": 0, "skip_clean": 0}
    processed = 0
    for rel, p in targets:
        url = extract_url(p)
        if not url:
            state[rel] = "no_url"
            stats["no_url"] += 1
            print("  [no_url] %s" % rel)
            continue
        # 备份
        if not args.no_backup:
            bk = os.path.join(backup_dir, rel)
            os.makedirs(os.path.dirname(bk), exist_ok=True)
            if not os.path.exists(bk):
                try:
                    import shutil
                    shutil.copy2(p, bk)
                except Exception as e:
                    print("  [warn] 备份失败 %s: %s" % (rel, e))
        # 打印到临时文件
        tmp = os.path.join(tempfile.mkdtemp(prefix="recover_"), "out.pdf")
        ok = False
        for attempt in (1, 2):
            try:
                r = subprocess.run([PY, SDP, "html2pdf", "--url", url, "--out", tmp],
                                   capture_output=True, timeout=200)
                if r.returncode == 0 and os.path.exists(tmp) and os.path.getsize(tmp) > 1024:
                    ok = True
                    break
            except Exception as e:
                print("  [warn] 打印异常 %s 尝试%d: %s" % (rel, attempt, e))
            if attempt < 2:
                import time
                time.sleep(12)
        if ok and is_clean(tmp):
            try:
                import shutil as _sh
                _sh.move(tmp, p)
                state[rel] = "done"
                stats["done"] += 1
                print("  [done] %s  (%d bytes)" % (rel, os.path.getsize(p)))
            except Exception as e:
                state[rel] = "fetch_fail"
                stats["fetch_fail"] += 1
                print("  [fetch_fail] 替换失败 %s: %s" % (rel, e))
        else:
            state[rel] = "fetch_fail"
            stats["fetch_fail"] += 1
            print("  [fetch_fail] 打印未产出干净 PDF: %s (url=%s)" % (rel, url))
        processed += 1
        json.dump(state, open(state_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        if args.delay:
            import time
            time.sleep(args.delay)

    print("\n== recover 汇总 ==")
    print("处理:", processed, "成功还原:", stats["done"],
          "无URL跳过:", stats["no_url"], "打印失败保留:", stats["fetch_fail"])
    print("状态文件:", state_path)


if __name__ == "__main__":
    main()
