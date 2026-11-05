# -*- coding: utf-8 -*-
"""技能依赖自举：检测 → 缺失自动安装 → 冒烟验证（可迁移性的第一道闸）。

为什么要有它（用户 2026-10-27 要求「确保代码内容执行依赖问题自动解决」）：
  本技能搬到新机器后，最容易踩的不是逻辑 bug，而是**依赖缺失**：
    * 缺 pymupdf   → 正文抽取/OCR/转 PDF 全废
    * 缺 openpyxl  → 台账读写全废
    * 缺 rapidocr  → 中基协扫描件（几乎全是）取不到正文 → 分类全落 06_其他
    * 缺 playwright→ 北交所抓不到、网页转 PDF 无兜底
    * 缺中文字体   → 生成的决定书 PDF 全是方块/乱码
  这些缺失在旧实现里大多被 try/except 吞掉，表现为「抓到了 0 条」「PDF 没生成」
  这种看不出根因的现象。本脚本把它们**前置成一次显式检查 + 自动安装**。

分档：
  required  必须；装不上就别跑流水线（会静默产出残缺库）
  optional  可选；装不上只降级（会打印影响面），不阻断

用法：
  python ensure_deps.py              # 检查 + 自动 pip 安装缺失项
  python ensure_deps.py --check-only # 只报告不安装（CI / 只读环境用）
  python ensure_deps.py --with-ocr   # 连 OCR 一起装（rapidocr 体积较大）
"""
import os
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

# (pip 包名, import 名, 档位, 缺了会怎样)
DEPS = [
    ("pymupdf", "fitz", "required", "PDF 正文抽取 / OCR / 生成决定书 PDF 全部不可用"),
    ("openpyxl", "openpyxl", "required", "台账读写全部失败"),
    ("requests", "requests", "optional", "部分抓取走 urllib 兜底，速度较慢"),
    ("rapidocr-onnxruntime", "rapidocr_onnxruntime", "ocr",
     "中基协决定书几乎全是扫描件 → 取不到正文 → 分类退化成 06_其他"),
    ("playwright", "playwright", "browser",
     "北交所抓取不可用；网页转 PDF 失去兜底（受管终端下 CLI 打印常失败）"),
    ("beautifulsoup4", "bs4", "laws", "06 配套法规抓取（网页型法规正文清洗）不可用"),
    ("python-docx", "docx", "laws", "06 配套法规无法排版为 Word 留存"),
]

PIP = [sys.executable, "-m", "pip"]


def _importable(name):
    import importlib
    try:
        importlib.import_module(name)
        return True
    except Exception:
        return False


def _install(pkg):
    print("  安装中：%s ..." % pkg)
    r = subprocess.run(PIP + ["install", pkg, "--quiet"],
                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if r.returncode != 0:
        msg = (r.stdout or b"").decode("utf-8", "ignore").strip().splitlines()
        print("  安装失败：%s" % (msg[-1] if msg else "未知错误"))
        return False
    return True


def _check_font():
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import cjk_pdf
        return True, cjk_pdf.find_cjk_font()
    except Exception as e:
        return False, str(e).splitlines()[0]


def _browser_ok():
    """playwright 的 pip 包装上 ≠ 浏览器二进制在（2026-10-27 实测踩坑：
    `playwright install` 没跑过 → BSE 浏览器启动失败 → 北交所静默 0 条）。
    executable_path 只解析路径不启动浏览器，代价极低。"""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            ep = p.chromium.executable_path
            return bool(ep) and os.path.exists(ep)
    except Exception:
        return False


def _install_browser():
    print("  安装中：playwright chromium 浏览器二进制 ...")
    ok = True
    for target in ("chromium", "chromium-headless-shell"):
        r = subprocess.run([sys.executable, "-m", "playwright", "install", target],
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        if r.returncode != 0:
            msg = (r.stdout or b"").decode("utf-8", "ignore").strip().splitlines()
            print("  安装失败 %s：%s" % (target, msg[-1] if msg else "未知错误"))
            ok = False
    return ok


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="技能依赖自举")
    ap.add_argument("--check-only", action="store_true", help="只检查不安装")
    ap.add_argument("--with-ocr", action="store_true", help="一并安装 OCR（体积大）")
    ap.add_argument("--with-browser", action="store_true", help="一并安装 playwright")
    ap.add_argument("--with-laws", action="store_true", help="一并安装 06 法规抓取所需依赖")
    a = ap.parse_args(argv)

    print("=" * 66)
    print("技能依赖自举（Python %s）" % sys.version.split()[0])
    print("=" * 66)

    missing_required, degraded = [], []
    for pkg, mod, tier, impact in DEPS:
        if tier == "ocr" and not a.with_ocr:
            tier = "optional"
        if tier == "browser" and not a.with_browser:
            tier = "optional"
        if tier == "laws" and not a.with_laws:
            tier = "optional"
        if _importable(mod):
            print("  [OK]   %-24s %s" % (mod, tier))
            continue
        print("  [MISS] %-24s %s —— %s" % (mod, tier, impact))
        if tier == "required":
            missing_required.append(pkg)
        else:
            degraded.append((pkg, impact))
        if not a.check_only:
            if _install(pkg):
                print("  [FIXED] %s 安装完成" % pkg)
                if tier == "required":
                    missing_required.remove(pkg)
                else:
                    degraded = [x for x in degraded if x[0] != pkg]
            elif tier == "required":
                print("  [FATAL] 必需依赖装不上，请先手工安装：pip install %s" % pkg)

    # 字体（PDF 中文可读性的命门，单独查）
    ok, info = _check_font()
    print("  [%s] %-24s %s" % ("OK" if ok else "MISS", "CJK 字体", info if ok else info))
    if not ok:
        degraded.append(("CJK 字体", "生成的决定书 PDF 中文不可读/不可检索"))

    # 浏览器二进制（playwright pip 包之外的第二道，北交所抓取的命门）
    if _importable("playwright"):
        ok_b = _browser_ok()
        print("  [%s] %-24s %s" % ("OK" if ok_b else "MISS", "chromium 二进制",
                                   "已就绪" if ok_b else "playwright install chromium 未执行"))
        if not ok_b and not a.check_only:
            if _install_browser() and _browser_ok():
                print("  [FIXED] chromium 安装完成")
            else:
                degraded.append(("chromium 二进制", "北交所抓取不可用；网页转 PDF 失去兜底"))
        elif not ok_b:
            degraded.append(("chromium 二进制", "北交所抓取不可用；网页转 PDF 失去兜底"))

    print("-" * 66)
    if missing_required:
        print("FAIL 必需依赖仍缺失：%s" % "、".join(missing_required))
        print("  请先安装后再跑流水线，否则会静默产出残缺库。")
        return 1
    if degraded:
        print("WARN 以下能力将降级（不阻断）：")
        for pkg, impact in degraded:
            print("  - %s：%s" % (pkg, impact))
        if any(p == "rapidocr-onnxruntime" for p, _ in degraded):
            print("  →  OCR 缺失时请用 `--ocr-cache <既有缓存目录>` 复用历史结果，"
                  "或跑 `python ensure_deps.py --with-ocr` 装上。")
    else:
        print("PASS 依赖齐备，可正常执行流水线。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
