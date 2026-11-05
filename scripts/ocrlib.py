# -*- coding: utf-8 -*-
"""OCR 唯一事实源（2026-10-30 收敛）。

为什么要有这个模块：
  * 原来「扫描件 OCR」有**三套实现**：`verify_library.ocr_pdf`（校验用）、
    `fetch_csrc/doc_text` 里的红通道渲染、`amac_lookup` 的末页 OCR。
    三套的 dpi / 红通道 / 最大像素 / 缓存键全是各自手写，改一处忘两处 → 结论对不上。
  * 更贵的是**缓存键**：`verify` 用 `v2red|<basename>|<size>|<mtime>|` 做 sha1，
    并发预填脚本若自己也拼一遍，哪怕差一个分隔符，缓存永远命不中，等于白跑 2 小时。
  * 所以本模块是唯一实现：`cache_key()` / `ocr_pdf()` 都从这里出，其他脚本一律 import。

红通道（2026-09-29 实测，根因级）：决定书落款是**黑字压红章**，普通 RGB→灰度时红章变深灰
把字糊掉，末页整段落款会直接消失。取 R 通道（红章被推到背景、黑字保留）后可读出正确日期。
缓存键带 `v2red|` 前缀：改渲染方式必须换前缀让旧缓存失效（0.4.15）。
"""
import os
import hashlib

OCR_DPI = 200
OCR_MAX_PX = 2000          # 高分辨率扫描件按固定 dpi 会渲染出超大图，OCR 变慢甚至返回空
CACHE_TAG = "v2red|"       # 渲染方案版本前缀；改红通道/dpi 必须同步改这里

_OCR = None                # 单实例（同进程内复用，避免每次重建模型）


def cache_key(fp):
    """OCR 缓存键：文件身份 = 文件名 + 大小 + mtime（不用内容哈希，避免每轮读全文件）。"""
    st = os.stat(fp)
    raw = f"{CACHE_TAG}{os.path.basename(fp)}|{st.st_size}|{int(st.st_mtime)}"
    return hashlib.sha1(raw.encode()).hexdigest()


def cache_path(fp, cache_dir):
    return os.path.join(cache_dir, cache_key(fp) + ".txt")


def _engine():
    global _OCR
    if _OCR is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR = RapidOCR()
    return _OCR


def _page_img(pg, dpi, red):
    import numpy as np
    import fitz
    zoom = dpi / 72.0
    w = pg.rect.width * zoom
    if w > OCR_MAX_PX:
        zoom *= OCR_MAX_PX / w
    pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if red and img.ndim == 3 and img.shape[2] >= 3:
        R = np.ascontiguousarray(img[:, :, 0])
        img = np.stack([R, R, R], axis=2)
    return img


def ocr_pdf(path, dpi=OCR_DPI, red=True, pages=None):
    """串行 OCR 一份 PDF，返回全文（页间用换行分隔）。

    pages=None 表示全部页；需要提速时可传页序号列表（如 [0, 1, -1, -2]）。
    并发交给**多进程**（每进程一个模型实例），不要在本函数里开线程——
    onnxruntime 多线程共享同一实例会 OOM 且结果错乱。
    """
    import fitz
    eng = _engine()
    out = []
    d = fitz.open(path)
    try:
        n = d.page_count
        idxs = range(n) if pages is None else [i % n for i in pages]
        for i in idxs:
            r, _ = eng(_page_img(d[i], dpi, red))
            if r:
                out.append(" ".join(t[1] for t in r))
    finally:
        d.close()
    return "\n".join(out)


def ocr_text(path, cache_dir=None, pages=None):
    """带缓存的 OCR：命中直接读文本，未命中才跑模型并落缓存。"""
    if cache_dir:
        p = cache_path(path, cache_dir)
        if os.path.exists(p):
            try:
                return open(p, encoding="utf-8").read()
            except Exception:
                pass
    t = ocr_pdf(path, pages=pages)
    if cache_dir:
        try:
            os.makedirs(cache_dir, exist_ok=True)
            open(cache_path(path, cache_dir), "w", encoding="utf-8").write(t or "")
        except Exception:
            pass
    return t


def has_text_layer(path, min_chars=60):
    """有文字层（原生/可选中文本）返回 True；扫描件返回 False。"""
    try:
        import fitz
        d = fitz.open(path)
        t = "".join(p.get_text() for p in d)
        d.close()
    except Exception:
        return False
    return len("".join(t.split())) >= min_chars
