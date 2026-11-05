# -*- coding: utf-8 -*-
"""中基协/证监局决定书落款日期提取器（红通道滤章 OCR）。

用法：
    python sign_date.py <pdf> [<pdf2> ...]            # 打印 文件<TAB>落款日期<TAB>定位方案
    python sign_date.py --json out.json <pdf> ...     # 同时输出 JSON

要点（详见 SKILL.md 0.1.3 / 0.1.4）：
  1. 落款压红章 -> 取 R 通道构造灰度图（红章亮、黑字暗），可修复「2026年4月0日」这类吞字；
  2. 定位顺序：末页中部偏右 -> 末页全页 -> 倒数第 2 页全页；
  3. 支持中文数字日期「二〇二六年四月三十日」；
  4. 只返回「最后一个」匹配到的日期，仍需按 SKILL.md 用「文号年份 + 挂网日上界 + 批次聚类」复核。
"""
import sys, os, re, json
import numpy as np
import fitz

OCR = None


def get_ocr():
    global OCR
    if OCR is None:
        from rapidocr_onnxruntime import RapidOCR
        OCR = RapidOCR()
    return OCR


def red_channel(img):
    """红通道滤章：R 通道里红章接近白（亮）、黑字暗 => 印章被推到背景。"""
    if img.ndim == 3 and img.shape[2] >= 3:
        R = np.ascontiguousarray(img[:, :, 0])
        return np.stack([R, R, R], axis=2)
    return img


MAX_PX = 2000   # 渲染宽度上限：高分辨率扫描件按固定 dpi 渲染会出超大图，OCR 变慢甚至失败


def render(doc, pno, clip=None, dpi=300):
    pg = doc[pno]
    w_in = pg.rect.width / 72.0 or 8.27          # 页面物理宽度（英寸）
    dpi = max(72, min(dpi, int(MAX_PX / w_in)))  # 按页面尺寸封顶，避免超大图
    pix = pg.get_pixmap(dpi=dpi, clip=clip) if clip else pg.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    return img[:, :, :3] if img.shape[2] >= 3 else np.stack([img[:, :, 0]] * 3, axis=2)


def ocr_text(ocr, img):
    r, _ = ocr(red_channel(img))
    return " ".join(t[1] for t in r) if r else ""


CN = {"〇": 0, "零": 0, "O": 0, "o": 0, "0": 0, "。": 0, "一": 1, "二": 2, "三": 3, "四": 4,
      "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def cn_num(s):
    """解析中文数字：支持「二〇二六」「十五」「二十」「三十一」及阿拉伯数字。"""
    s = s.strip()
    if s.isdigit():
        return int(s)
    if all(c in CN for c in s):                      # 二〇二六 这类逐位读
        return int("".join(str(CN[c]) for c in s))
    if "十" in s:                                     # 十五 / 二十 / 三十一
        a, _, b = s.partition("十")
        return (CN.get(a, 1) if a else 1) * 10 + (CN.get(b, 0) if b else 0)
    return None


# 容错正则：年/月/日三个组件之间允许 0~4 个杂字符（OCR 常插入乱码）
DATE_PATS = [
    re.compile(r"([〇零Oo0。一二三四五六七八九]{4})\s*年.{0,4}?([一二三四五六七八九十]{1,3}|\d{1,2})\s*月.{0,4}?([一二三四五六七八九十]{1,3}|\d{1,2})\s*日"),
    re.compile(r"(20\d{2})\s*[.．,，、\-—/／\s°º•·。､]{0,3}\s*年.{0,4}?(\d{1,2})\s*月.{0,4}?(\d{1,2})\s*日"),
    re.compile(r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})"),
]


def today_iso():
    import datetime
    return datetime.date.today().isoformat()


def parse_dates(text, max_date=None):
    """返回文本里所有可解析日期（ISO 字符串列表，按出现顺序）。

    max_date（2026-10-15 新增，见 SKILL 0.4.54）：过滤掉晚于该上界的候选。
    决定书正文常出现「请在 2026 年 12 月 31 日前完成整改」这类**未来履行期限**，
    它排在落款之前/之后都可能被 OCR 先扫到，从而把落款顶替掉。
    默认 None 保持历史行为；`find_sign()` 传入今天作上限。
    """
    out, t = [], re.sub(r"\s+", "", text)
    # OCR 噪声：落款常被识别成「2026年8月25／日」「2026年8月25/日」（日与「日」之间插斜杠），
    # 导致日期模式匹配不到。实测 2 例扫描件复核决定书（金佩君/南京建邺三正厚德）。
    t = re.sub(r"(?<=\d)[／/|](?=日)", "", t)
    text = re.sub(r"(?<=\d)[／/|](?=日)", "", text)
    for src in (text, t):
        for p in DATE_PATS:
            for y, m, d in p.findall(src):
                yy, mm, dd = cn_num(y), cn_num(m), cn_num(d)
                if yy and mm and dd and 2015 <= yy <= 2100 and 1 <= mm <= 12 and 1 <= dd <= 31:
                    s = "%04d-%02d-%02d" % (yy, mm, dd)
                    if max_date and s > max_date:      # 未来日期（整改期限等）不可能是落款
                        continue
                    if s not in out:
                        out.append(s)
    return out


def find_sign(path, dpi_mid=300, dpi_full=250, max_date="__today__"):
    """返回 (date|None, plan, raw_text)。date 取定位区域内最后一个日期。

    max_date（2026-10-15 新增，SKILL 0.4.54）：落款候选的上界，默认取今天。
    决定书里的「整改截止日 / 报告提交期限」常在 OCR 文本中排在落款附近，
    不设上界会被当成落款返回（实测沪证监决〔2026〕249号 被取成 2026-12-31，
    真落款为 2026-10-12）。传 None 可关闭该过滤。
    """
    if max_date == "__today__":
        max_date = today_iso()
    ocr = get_ocr()
    d = fitz.open(path)
    try:
        n = len(d)
        plans = [("pL_mid", n - 1, None, dpi_mid)]
        r = d[n - 1].rect
        plans[0] = ("pL_mid", n - 1, fitz.Rect(r.width * 0.25, r.height * 0.10,
                                               r.width, r.height * 0.75), dpi_mid)
        plans.append(("pL_bottom", n - 1, fitz.Rect(0, r.height * 0.55,
                                                    r.width, r.height), dpi_full))
        plans.append(("pL_full", n - 1, None, dpi_full))
        if n >= 2:
            plans.append(("p2_full", n - 2, None, dpi_full))
        for name, pno, clip, dpi in plans:
            try:
                raw = ocr_text(ocr, render(d, pno, clip, dpi))
            except Exception:
                continue
            ds = parse_dates(raw, max_date)
            if ds:
                return ds[-1], name, raw[:200]
        return None, "none", ""
    finally:
        d.close()


def main():
    args = [a for a in sys.argv[1:]]
    out_json = None
    if args and args[0] == "--json":
        out_json, args = args[1], args[2:]
    res = []
    for p in args:
        if not os.path.exists(p):
            print("%s\tERR\tnot-found" % p)
            continue
        try:
            dt, plan, raw = find_sign(p)
        except Exception as e:
            dt, plan, raw = None, "exc", str(e)[:120]
        print("%s\t%s\t%s" % (os.path.basename(p), dt or "NONE", plan))
        res.append({"path": p, "date": dt, "plan": plan, "raw": raw})
    if out_json:
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
