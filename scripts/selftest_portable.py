# -*- coding: utf-8 -*-
"""技能可迁移性自检（portability self-test）。

为什么要有这个脚本（2026-10-27 用户明确要求）：
  本技能要能整包拷贝给别人、在另一台机器上正常跑。迁移失败的头号原因是
  **脚本/文档里写死了开发机的绝对路径**（盘符、用户名、某次会话的工作空间目录），
  换机器后路径不存在 → import 失败、抓取 0 条、或更糟：静默写到错误位置。

  这类问题没法靠"记得别写"来避免，必须做成可执行的检查。本脚本扫描技能全目录，
  凡命中禁止模式且未显式标注豁免的行，一律报错。

检查项：
  A. 用户主目录绝对路径（Windows: X:\\Users\\<name>；POSIX: /Users/<name>、/home/<name>）
  B. 会话工作空间目录（YYYY-MM-DD-HH-MM-SS 形态）
  C. 开发机自建盘符工作目录（如 X:\\Workbuddy、X:\\WorkBuddy）
  D. 临时抓取目录（tmp_crawl / tmp_pilot / tmp_rebuild）作为**代码里的默认路径**
  E. 脚本内 sys.path 必须相对自身目录，不得插入绝对路径

豁免方式（二者其一）：
  在该行行尾加注释 `# portable-ok`，或在 Markdown 行内加 `<!-- portable-ok -->`。
  仅用于"文档里举例说明本机现状"这类确实需要写路径的场合，代码里原则上不给豁免。

用法：
    python selftest_portable.py                 # 扫描本脚本所在技能的根目录
    python selftest_portable.py --root <技能根>  # 显式指定
退出码：0=可迁移；1=发现硬编码，必须修。
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(HERE)

OK_MARK = "portable-ok"

# ---------------------------------------------------------------- 禁止模式
PATTERNS = [
    ("A-用户主目录", re.compile(
        r"(?i)(?:[A-Za-z]:[\\/](?:Users|Documents and Settings)[\\/][^\\/\s\"']+"
        r"|/(?:Users|home)/[A-Za-z0-9._-]+)")),
    ("B-会话工作空间", re.compile(r"\b20\d{2}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}\b")),
    ("C-自建盘符工作目录", re.compile(r"(?i)[A-Za-z]:[\\/]Work[Bb]uddy")),
    # D：只禁【绝对路径】形态的临时目录。文档里写 `tmp_crawl/xxx` 这种相对名是
    # 可迁移的（相对当前工作空间解析），不必禁；只有写成绝对路径的才是本机私有。
    ("D-临时抓取目录绝对路径", re.compile(
        r"(?i)(?:[A-Za-z]:[\\/]|/(?:Users|home)/)[^\s\"'`]*?"
        r"(?:tmp_crawl|tmp_pilot|tmp_rebuild)")),
]
# 标准程序安装目录不属于"本机私有路径"，允许出现
ALLOW = re.compile(
    r"(?i)(?:[A-Za-z]:[\\/]Program Files(?: \(x86\))?[\\/]"
    r"|[A-Za-z]:[\\/]Windows[\\/])")

SCAN_EXT = (".py", ".md", ".json", ".txt", ".yaml", ".yml")


def _is_exempt(line: str) -> bool:
    return OK_MARK in line


def scan_file(path: str):
    """返回 [(行号, 类别, 行内容)]。"""
    hits = []
    try:
        with open(path, encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
    except OSError:
        return hits
    for i, raw in enumerate(lines, 1):
        line = raw.rstrip("\n")
        if _is_exempt(line):
            continue
        for tag, pat in PATTERNS:
            m = pat.search(line)
            if not m:
                continue
            frag = m.group(0)
            if ALLOW.search(frag) or ALLOW.search(line):
                continue
            hits.append((i, tag, line.strip()[:160]))
            break
    return hits


def check_syspath(path: str):
    """脚本里 sys.path.insert 必须是相对自身目录的写法。"""
    bad = []
    try:
        with open(path, encoding="utf-8", errors="ignore") as f:
            for i, raw in enumerate(f, 1):
                line = raw.rstrip("\n")
                if "sys.path" not in line or _is_exempt(line):
                    continue
                if re.search(r"(?i)[A-Za-z]:[\\/]|/(?:Users|home)/", line):
                    bad.append((i, "E-绝对路径sys.path", line.strip()[:160]))
    except OSError:
        pass
    return bad


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="技能可迁移性自检")
    ap.add_argument("--root", default=SKILL_ROOT, help="技能根目录")
    a = ap.parse_args(argv)
    root = os.path.abspath(a.root)

    total = 0
    files = 0
    for dp, dns, fns in os.walk(root):
        dns[:] = [d for d in dns if d not in ("__pycache__", ".git")]
        for fn in sorted(fns):
            if not fn.lower().endswith(SCAN_EXT):
                continue
            p = os.path.join(dp, fn)
            rel = os.path.relpath(p, root).replace("\\", "/")
            files += 1
            hits = scan_file(p)
            if p.lower().endswith(".py"):
                hits += check_syspath(p)
            for ln, tag, txt in hits:
                total += 1
                print(f"[FAIL] {rel}:{ln}  {tag}\n       {txt}")

    print("-" * 66)
    if total == 0:
        print(f"PASS 可迁移性自检通过：扫描 {files} 个文件，未发现硬编码本机路径。")
        return 0
    print(f"FAIL 发现 {total} 处硬编码本机路径（扫描 {files} 个文件）。")
    print("  修法：改成相对路径 / 命令行参数 / 环境变量；确实要写本机路径则在该行加 "
          f"`{OK_MARK}` 标注。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
