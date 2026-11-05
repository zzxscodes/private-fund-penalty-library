# -*- coding: utf-8 -*-
"""并发 OCR 预填缓存（2026-10-30）。

为什么要有这个脚本（实测数据）：
  * 库内 244 件扫描件、807 页。`verify --ocr-scan all` 是**单进程串行** OCR
    （单实例共享，开线程会 OOM），实测约 **2 小时**才跑完——验收前等不起。
  * 而 OCR 结果**与进程无关、只与文件身份有关**（`ocrlib.cache_key`）。所以先用
    多进程把缓存填满，verify 再跑就是"缓存命中"，几秒出结果。
  * 并发度别开太大：每进程一个 RapidOCR 实例约占 0.6~0.9 GB，20 核 / 16 GB 机器上
    6 并发是安全档（实测 807 页约 12~20 分钟）。

用法：
    python ocr_prefill.py --root "<库根>" --cache-dir "<缓存目录>" --jobs 6
    python ocr_prefill.py --root "<库根>" --cache-dir "<缓存>" --jobs 6 --only-missing
"""
import os
import sys
import time
import argparse
from concurrent.futures import ProcessPoolExecutor, as_completed

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ocrlib  # noqa: E402

SCAN_DIRS = ("01_", "02_", "03_")


def collect(root):
    """收集库内 01/02/03 三个案例目录下**无文字层**的 PDF（有文字层的不用 OCR）。"""
    out = []
    for name in sorted(os.listdir(root)):
        if not name.startswith(SCAN_DIRS):
            continue
        d0 = os.path.join(root, name)
        if not os.path.isdir(d0):
            continue
        for dp, _dn, fs in os.walk(d0):
            for f in sorted(fs):
                if not f.lower().endswith(".pdf"):
                    continue
                fp = os.path.join(dp, f)
                try:
                    if os.path.getsize(fp) < 3000:
                        continue
                except OSError:
                    continue
                if not ocrlib.has_text_layer(fp):
                    out.append(fp)
    return out


def work(args):
    """子进程任务：OCR 一份并写缓存。返回 (路径, 页OCR字符数, 秒)。"""
    fp, cache_dir = args
    t0 = time.time()
    try:
        t = ocrlib.ocr_text(fp, cache_dir=cache_dir)
    except Exception as e:
        return (fp, -1, time.time() - t0, str(e)[:80])
    return (fp, len(t or ""), time.time() - t0, "")


def main():
    ap = argparse.ArgumentParser(description="并发 OCR 预填缓存（供 verify --ocr-scan all 命中）")
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--cache-dir", required=True, help="OCR 文本缓存目录（必须与 verify --ocr-cache 同一个）")
    ap.add_argument("--jobs", type=int, default=6, help="并发进程数（默认 6；每进程约 0.6~0.9GB）")
    ap.add_argument("--only-missing", action="store_true",
                    help="只处理缓存里还没有的文件（默认也是跳过已缓存的，此参数仅为显式表达）")
    a = ap.parse_args()

    if not os.path.isdir(a.root):
        sys.exit("库根不存在：%s" % a.root)
    os.makedirs(a.cache_dir, exist_ok=True)

    files = collect(a.root)
    todo = [fp for fp in files
            if not os.path.exists(ocrlib.cache_path(fp, a.cache_dir))]
    print("扫描件 %d 份 ｜ 已有缓存 %d ｜ 待 OCR %d ｜ 并发 %d"
          % (len(files), len(files) - len(todo), len(todo), a.jobs))
    if not todo:
        print("全部命中缓存，无需 OCR。")
        return

    t0 = time.time()
    done = 0
    fail = []
    empty = []
    try:
        with ProcessPoolExecutor(max_workers=a.jobs) as ex:
            futs = [ex.submit(work, (fp, a.cache_dir)) for fp in todo]
            for fu in as_completed(futs):
                fp, n, sec, err = fu.result()
                done += 1
                if err:
                    fail.append((fp, err))
                elif n == 0:
                    empty.append(fp)
                if done % 10 == 0 or done == len(todo):
                    el = time.time() - t0
                    eta = el / done * (len(todo) - done)
                    print("  %d/%d ｜ 已用 %.0fs ｜ 预计剩余 %.0fs ｜ %s"
                          % (done, len(todo), el, eta, os.path.basename(fp)[:34]), flush=True)
    except KeyboardInterrupt:
        print("\n中断（已完成的缓存保留，可重跑续跑）")

    print("完成 %d 份 ｜ 耗时 %.0fs ｜ 空结果 %d ｜ 失败 %d"
          % (done, time.time() - t0, len(empty), len(fail)))
    for fp, err in fail[:10]:
        print("  [FAIL]", os.path.basename(fp), err)
    if fail:
        sys.exit(1)


if __name__ == "__main__":
    main()
