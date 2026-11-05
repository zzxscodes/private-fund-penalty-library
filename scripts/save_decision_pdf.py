# -*- coding: utf-8 -*-
"""原始决定书 PDF 留存工具。所有方式最终都产出真实 PDF 文件，不留存网页。

用法:
    python save_decision_pdf.py download   --url "<PDF直链>" --out "<输出.pdf>"
    python save_decision_pdf.py html2pdf   --url "<网页URL>" --out "<输出.pdf>"
    python save_decision_pdf.py htmlfile2pdf --file "<本地.html>" --out "<输出.pdf>"
    python save_decision_pdf.py images2pdf --out "<输出.pdf>" "<图1>" "<图2>" ...
    python save_decision_pdf.py text2pdf   --file "<文字.txt>" --out "<输出.pdf>" [--title "<标题>"]

说明:
    - download:    直接下载官方 PDF（带浏览器 UA）。
    - html2pdf:    用本机 Edge/Chrome 无头模式把网页打印为 PDF，保留官方页面原貌。
    - htmlfile2pdf:把本地 HTML（如已抓取的公告正文）打印为 PDF。
    - images2pdf:  把扫描件/公告图片按给定顺序合并为一个 PDF（依赖 Pillow），保留原始影像。
    - text2pdf:    纯文字决定书按原文排版包一层 HTML 后打印为 PDF（中文自动用系统黑体）。
    每一步都会校验输出文件真实存在且非空，失败以非零退出码结束。
"""
import argparse
import html
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

def browser_candidates():
    """浏览器候选路径（可迁移，2026-10-27）。

    顺序：① 环境变量 CHROME_PATH/BROWSER_PATH（显式指定优先）
          ② 各平台常见安装位置（Windows / Linux / macOS）
          ③ PATH 里的 msedge.exe、chrome.exe、google-chrome、chromium
    硬编码的 Windows 默认安装路径只是**候选之一**，找不到不会直接报错。
    """
    out = []
    for k in ("CHROME_PATH", "BROWSER_PATH"):
        v = os.environ.get(k)
        if v:
            out.append(v)
    if os.name == "nt":
        # 实测：部分机器上 Edge 无头打印不产出文件，Chrome 稳定可用，故 Chrome 优先
        out += [
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        local = os.environ.get("LOCALAPPDATA") or ""
        if local:
            out.append(os.path.join(local, "Google", "Chrome", "Application", "chrome.exe"))
    else:
        out += [
            "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium", "/usr/bin/chromium-browser",
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ]
    return [p for p in out if p]


def ok_or_die(out: Path) -> None:
    if out.exists() and out.stat().st_size > 1024:
        print(f"OK: {out} ({out.stat().st_size} bytes)")
        return
    print(f"ERROR: 输出文件未生成或过小: {out}", file=sys.stderr)
    sys.exit(1)


def find_browser() -> str:
    for p in browser_candidates():
        if Path(p).exists():
            return p
    for name in ("msedge.exe", "chrome.exe", "google-chrome", "chromium", "chromium-browser"):
        found = shutil.which(name)
        if found:
            return found
    print("ERROR: 未找到 Edge/Chrome，无法执行网页转 PDF", file=sys.stderr)
    sys.exit(2)


def print_via_playwright(browser: str, target: str, out: Path) -> bool:
    """CLI 无头打印失败时的兜底：用 playwright 承载同一浏览器再 page.pdf()。

    2026-10-08 实测：某些环境（沙箱/受管终端）下直接 spawn chrome.exe
    `--headless --print-to-pdf=...` **进程根本没起来**（无 stderr、无产物），
    但同一个 chrome.exe 由 playwright 启动时却完全正常。
    ⇒ 不能只靠 CLI 一条路；必须在 CLI 失败后自动降级到 playwright，
    否则会误判成「网页转 PDF 不可用」而留下空壳/缺件。
    """
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        return False
    try:
        with sync_playwright() as p:
            br = p.chromium.launch(executable_path=browser, headless=True,
                                   args=["--no-sandbox", "--disable-gpu"])
            pg = br.new_page()
            pg.goto(target, wait_until="networkidle", timeout=90000)
            pg.wait_for_timeout(2000)
            pg.pdf(path=str(out), format="A4", print_background=True)
            br.close()
    except Exception as e:
        print(f"  [playwright 兜底失败] {repr(e)[:160]}", file=sys.stderr)
        return False
    return out.exists() and out.stat().st_size > 1024


def print_to_pdf(browser: str, target: str, out: Path) -> None:
    profile = Path(tempfile.mkdtemp(prefix="pdfprofile_"))
    try:
        cmd = [
            browser, "--headless", "--disable-gpu", "--no-first-run",
            f"--user-data-dir={profile}",
            "--no-pdf-header-footer",
            f"--print-to-pdf={out}",
            "--virtual-time-budget=20000",
            target,
        ]
        subprocess.run(cmd, timeout=180, capture_output=True)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if out.exists() and out.stat().st_size > 1024:
        ok_or_die(out)
        return
    print("  [warn] CLI 无头打印未产出文件，改用 playwright 兜底…", file=sys.stderr)
    if print_via_playwright(browser, target, out):
        ok_or_die(out)
        return
    ok_or_die(out)


def cmd_download(url: str, out: Path) -> None:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=120) as resp, open(out, "wb") as f:
        shutil.copyfileobj(resp, f)
    # 简单校验 PDF 魔数
    with open(out, "rb") as f:
        head = f.read(5)
    if head != b"%PDF-":
        print("ERROR: 下载内容不是 PDF（可能是防盗链页面），请改用 html2pdf", file=sys.stderr)
        out.unlink(missing_ok=True)
        sys.exit(1)
    ok_or_die(out)


def cmd_images2pdf(out: Path, images: list) -> None:
    try:
        from PIL import Image
    except ImportError:
        print("ERROR: 缺少 Pillow，请在托管 Python 隔离 venv 中安装: pip install Pillow", file=sys.stderr)
        sys.exit(2)
    pages = []
    for p in images:
        im = Image.open(p)
        if im.mode != "RGB":
            im = im.convert("RGB")
        pages.append(im)
    if not pages:
        print("ERROR: 未提供图片", file=sys.stderr)
        sys.exit(2)
    pages[0].save(out, save_all=True, append_images=pages[1:])
    ok_or_die(out)


def cmd_text2pdf(txt: Path, out: Path, title: str) -> None:
    body = html.escape(txt.read_text(encoding="utf-8"))
    page = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        f"<title>{html.escape(title)}</title>"
        "<style>body{font-family:'Microsoft YaHei',SimSun,sans-serif;margin:2.5cm;"
        "line-height:1.8;font-size:12pt;} pre{white-space:pre-wrap;font-family:inherit;}"
        "h1{font-size:16pt;text-align:center;}</style></head><body>"
        f"<h1>{html.escape(title)}</h1><pre>{body}</pre></body></html>"
    )
    tmp = Path(tempfile.mkdtemp(prefix="text2pdf_")) / "page.html"
    tmp.write_text(page, encoding="utf-8")
    print_to_pdf(find_browser(), str(tmp), out)


def main() -> None:
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("download")
    p.add_argument("--url", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("html2pdf")
    p.add_argument("--url", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("htmlfile2pdf")
    p.add_argument("--file", required=True)
    p.add_argument("--out", required=True)

    p = sub.add_parser("images2pdf")
    p.add_argument("--out", required=True)
    p.add_argument("images", nargs="+")

    p = sub.add_parser("text2pdf")
    p.add_argument("--file", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--title", default="决定书")

    args = ap.parse_args()
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    if args.cmd == "download":
        cmd_download(args.url, out)
    elif args.cmd == "html2pdf":
        print_to_pdf(find_browser(), args.url, out)
    elif args.cmd == "htmlfile2pdf":
        print_to_pdf(find_browser(), str(Path(args.file).resolve()), out)
    elif args.cmd == "images2pdf":
        cmd_images2pdf(out, args.images)
    elif args.cmd == "text2pdf":
        cmd_text2pdf(Path(args.file), out, args.title)


if __name__ == "__main__":
    main()
