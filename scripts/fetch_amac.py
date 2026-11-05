# -*- coding: utf-8 -*-
"""中基协「纪律处分 → 受处分机构 / 受处分人员」全量抓取管道（v2 固化版，2026-10-26）。

用法：
    python fetch_amac.py --out items.json                       # 只抓列表（两栏成对、全部分页）
    python fetch_amac.py --out items.json --max-pages 2         # 试跑：每栏只翻 2 页
    python fetch_amac.py --out items.json --since 2026-01-01    # 只保留挂网日 >= since
    python fetch_amac.py --out items.json --download-dir dl     # 顺带下载决定书 PDF
    python fetch_amac.py --out items.json --download-dir dl --text-dir texts --ocr-cache .ocr_cache

产出 JSON：
    {"fetched_at": "...", "columns": {"scfjg": {...}, "scfry": {...}}, "items": [...]}
    item = {column, title, url, publish, local(可选), text(可选)}

铁律（详见 SKILL.md 0.4.71）：
  1. 两栏必须成对抓 —— v2 一行=一个案例（机构+高管合并），漏抓「受处分人员」等于半个案例；
  2. 列表混排协会动态，必须过滤 /hydj/ 并要求 URL 含 /zlgl/jlcf/，否则正则 0 命中；
  3. 挂网发布日期内嵌在 URL 的 /P0(\\d{8})，优先取它；取不到标「待补」，禁止拿落款冒充当发布日；
  4. 拼绝对地址一律用 urllib.parse.urljoin —— 手写会拼出 https://https:/... 表现为全部 502；
  5. 中基协决定书全是扫描件（无文字层），正文必须 OCR（R 通道），并复用 --ocr-cache 避免重跑。
"""
import sys, os, re, json, time, ssl, hashlib, argparse
import urllib.request
from urllib.parse import urljoin

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

COLS = {
    "scfjg": ("受处分机构", "https://www.amac.org.cn/zlgl/jlcf/scfjg/"),
    "scfry": ("受处分人员", "https://www.amac.org.cn/zlgl/jlcf/scfry/"),
}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}
ITEM_RE = re.compile(r'href="([^"]+?P0\d+\.pdf)"[^>]*>(?:\s*<[^>]+>\s*)*([^<]{4,200})')
PUB_RE = re.compile(r"/P0(\d{8})")


# ---------------------------------------------------------------- HTTP / 缓存
def http(url, binary=False, timeout=40, retry=3):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    for i in range(retry):
        try:
            b = urllib.request.urlopen(urllib.request.Request(url, headers=UA),
                                       timeout=timeout, context=ctx).read()
            return b if binary else b.decode("utf-8", "ignore")
        except Exception as e:
            if i == retry - 1:
                print("  [WARN] 取不到 %s (%s)" % (url, str(e)[:80]))
                return None
            time.sleep(1.5 * (i + 1))


def cached_get(url, cache_dir, binary=False, ttl_days=7):
    """列表页缓存：中基协改版慢，重跑必须复用，避免重复触发频控。"""
    if not cache_dir:
        return http(url, binary=binary)
    os.makedirs(cache_dir, exist_ok=True)
    p = os.path.join(cache_dir, hashlib.sha1(url.encode()).hexdigest() + (".bin" if binary else ".html"))
    if os.path.exists(p) and (time.time() - os.path.getmtime(p)) < ttl_days * 86400:
        return open(p, "rb").read() if binary else open(p, encoding="utf-8").read()
    d = http(url, binary=binary)
    if d is not None:
        open(p, "wb").write(d if isinstance(d, bytes) else d.encode("utf-8"))
    return d


# ---------------------------------------------------------------- 列表解析
def parse_items(html, base, column):
    out = []
    for m in ITEM_RE.finditer(html or ""):
        href, title = m.group(1), re.sub(r"\s+", " ", m.group(2)).strip()
        u = urljoin(base, href)                     # 铁律 4：绝不手写拼接
        if "/zlgl/jlcf/" not in u or "/hydj/" in u:  # 铁律 2：剔除协会动态
            continue
        mm = PUB_RE.search(u)
        pub = "%s-%s-%s" % (mm.group(1)[:4], mm.group(1)[4:6], mm.group(1)[6:]) if mm else ""
        out.append({"column": column, "title": title, "url": u, "publish": pub or "待补"})
    return out


def crawl_column(key, base, max_pages, cache_dir):
    idx = cached_get(base + "index.html", cache_dir)
    if not idx:
        return [], 0, False
    m = re.search(r"createPageHTML\((\d+)", idx)
    total = int(m.group(1)) if m else 1
    pages = total if not max_pages else min(total, max_pages)
    items = parse_items(idx, base, key)
    for p in range(1, pages):
        h = cached_get("%sindex_%d.html" % (base, p), cache_dir)
        if h:
            items += parse_items(h, base, key)
    seen, uniq = set(), []
    for it in items:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        uniq.append(it)
    uniq.sort(key=lambda x: x["publish"], reverse=True)
    return uniq, total, pages >= total


# ---------------------------------------------------------------- 正文（PDF → 文本）
OCR_MAX_PX = 2000
_OCR = None


def _get_ocr():
    """返回 OCR 引擎；本机没装 rapidocr 时返回 None（**不许抛异常**）。

    可迁移性（2026-10-27 实测）：换机器后 `rapidocr_onnxruntime` 往往没装，
    旧实现直接 ModuleNotFoundError 崩掉整个抓取（中基协决定书全是扫描件，必然走到 OCR）。
    正确行为是降级返回空正文并打警告，让流程继续——正文后续可由 --ocr-cache 或
    verify 阶段补做，而不是让一轮抓取前功尽弃。
    """
    global _OCR
    if _OCR is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except Exception:
            print("  [WARN] 未安装 rapidocr_onnxruntime，扫描件正文无法 OCR（可 pip install "
                  "rapidocr-onnxruntime，或复用 --ocr-cache）", file=sys.stderr)
            return None
        _OCR = RapidOCR()
    return _OCR


def _cache_keys(fp):
    """OCR 缓存键：**内容哈希优先，兼容旧 mtime 键**。

    为什么必须改（2026-10-27 实测）：旧键 = sha1(文件名|大小|mtime)。
    滑动窗口每轮重新下载决定书 → mtime 变成"刚才" → **旧缓存 100% 不命中**，
    于是每轮都要实跑 OCR（全库首次约 70 分钟），而内容其实一字未变。
    改成内容哈希后，同一份文书哪怕重新下载也能命中；查找时两个键都试，
    保证既有存量缓存（mtime 键）不会被作废。
    """
    h = hashlib.sha1()
    try:
        with open(fp, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        content = "c|" + h.hexdigest()
    except OSError:
        content = ""
    st = os.stat(fp)
    legacy = hashlib.sha1(("v2red|%s|%d|%d" % (os.path.basename(fp), st.st_size,
                                               int(st.st_mtime))).encode()).hexdigest()
    return [k for k in (content, legacy) if k]


def _cache_path(fp, ocr_cache):
    """返回 (写入路径, 候选查找路径列表)。"""
    ks = _cache_keys(fp)
    return os.path.join(ocr_cache, ks[0] + ".txt"), [os.path.join(ocr_cache, k + ".txt")
                                                     for k in ks]


def pdf_text(fp, ocr_cache=None, dpi=200):
    """有文字层直接取；无文字层走 R 通道 OCR（与 verify_library 同缓存键，可互相命中）。"""
    try:
        import fitz
    except ImportError:
        import pymupdf as fitz
    d = fitz.open(fp)
    try:
        raw = "\n".join(pg.get_text() or "" for pg in d)
        if len(re.sub(r"\s", "", raw)) >= 200:
            return raw, "text"
    finally:
        d.close()
    if ocr_cache:
        os.makedirs(ocr_cache, exist_ok=True)
        cp, cands = _cache_path(fp, ocr_cache)
        for c in cands:
            if os.path.exists(c):
                return open(c, encoding="utf-8").read(), "ocr-cache"
    ocr = _get_ocr()
    if ocr is None:
        return "", "no-ocr"          # 降级：不崩，交给调用方决定
    import numpy as np
    out = []
    d = fitz.open(fp)
    try:
        for pg in d:
            zoom = dpi / 72.0
            w = pg.rect.width * zoom
            if w > OCR_MAX_PX:
                zoom *= OCR_MAX_PX / w
            pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if img.ndim == 3 and img.shape[2] >= 3:          # 铁律 5：R 通道滤红章
                R = np.ascontiguousarray(img[:, :, 0])
                img = np.stack([R, R, R], axis=2)
            r, _ = ocr(img)
            if r:
                out.append(" ".join(t[1] for t in r))
    finally:
        d.close()
    txt = "\n".join(out)
    if ocr_cache:
        try:
            open(_cache_path(fp, ocr_cache)[0], "w", encoding="utf-8").write(txt)
        except Exception:
            pass
    return txt, "ocr"


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="中基协纪律处分全量抓取（受处分机构 + 受处分人员）")
    ap.add_argument("--out", required=True, help="输出 JSON 路径")
    ap.add_argument("--max-pages", type=int, default=0, help="每栏最多翻几页（0=全部）")
    ap.add_argument("--cache-dir", default="", help="列表页缓存目录（强烈建议复用旧目录）")
    ap.add_argument("--since", default="", help="只保留挂网日 >= since")
    ap.add_argument("--until", default="", help="只保留挂网日 <= until")
    ap.add_argument("--limit", type=int, default=0, help="每栏最多保留 N 条（调试用）")
    ap.add_argument("--download-dir", default="", help="下载决定书 PDF 到该目录")
    ap.add_argument("--text-dir", default="", help="抽取正文到该目录（自动 OCR 扫描件）")
    ap.add_argument("--ocr-cache", default="", help="OCR 文本缓存目录（与 verify_library 共用）")
    a = ap.parse_args()

    result = {"fetched_at": time.strftime("%Y-%m-%d %H:%M:%S"), "columns": {}, "items": [],
              "window": [a.since or "", a.until or ""]}
    for key, (label, base) in COLS.items():
        items, total, full = crawl_column(key, base, a.max_pages, a.cache_dir)
        # 2026-10-28：补 pages_fetched，供 assert_crawl_complete.py 做「页数是否翻全」断言
        pages_fetched = min(total, a.max_pages) if getattr(a, "max_pages", 0) else total
        result["columns"][key] = {"label": label, "total_pages": total, "full": full,
                                  "pages_fetched": pages_fetched,
                                  "items": len(items),
                                  # 2026-10-28：记录「平台最新一条的挂网日」——否则窗口内 0 条
                                  # 无法区分「本期真没发」与「抓取被限流/结构变了」。
                                  "latest_published": next(
                                      (x["publish"] for x in items if x["publish"] != "待补"), "")}
        print("%s（%s）：总页 %d，抓取 %s，条目 %d"
              % (label, key, total, "全部" if full else "部分(试跑)", len(items)))
        if a.since:
            items = [x for x in items if x["publish"] != "待补" and x["publish"] >= a.since]
        if a.until:
            items = [x for x in items if x["publish"] != "待补" and x["publish"] <= a.until]
        if a.limit:
            items = items[:a.limit]
        result["items"] += items

    # 下载
    if a.download_dir:
        os.makedirs(a.download_dir, exist_ok=True)
        ok = 0
        for it in result["items"]:
            fn = os.path.basename(it["url"])
            fp = os.path.join(a.download_dir, fn)
            if os.path.exists(fp) and os.path.getsize(fp) > 1024:
                it["local"] = fp
                ok += 1
                continue
            b = http(it["url"], binary=True)
            if b and b[:4] == b"%PDF":
                open(fp, "wb").write(b)
                it["local"] = fp
                ok += 1
            else:
                print("  [WARN] 下载失败 %s" % it["title"][:40])
        print("下载完成：%d/%d" % (ok, len(result["items"])))

    # 正文
    if a.text_dir:
        os.makedirs(a.text_dir, exist_ok=True)
        n = 0
        for it in result["items"]:
            fp = it.get("local")
            if not fp or not os.path.exists(fp):
                continue
            txt, how = pdf_text(fp, a.ocr_cache or None)
            tp = os.path.join(a.text_dir, os.path.splitext(os.path.basename(fp))[0] + ".txt")
            open(tp, "w", encoding="utf-8").write(txt)
            it["text"] = tp
            it["text_how"] = how
            n += 1
            print("  正文 %s（%s，%d 字）" % (os.path.basename(tp), how, len(txt)))
        print("正文抽取完成：%d 份" % n)

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    json.dump(result, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    # 批次预览：同一挂网日的机构件与人员件要配成对（v2 一行=一个案例）
    print("\n--- 按挂网批次预览（最新 10 批）---")
    batch = {}
    for it in result["items"]:
        batch.setdefault(it["publish"], {"scfjg": 0, "scfry": 0})
        batch[it["publish"]][it["column"]] += 1
    for d in sorted(batch, reverse=True)[:10]:
        print("  %s  机构 %d ／ 人员 %d" % (d, batch[d]["scfjg"], batch[d]["scfry"]))
    print("\n写出：%s（条目 %d）" % (a.out, len(result["items"])))


if __name__ == "__main__":
    main()
