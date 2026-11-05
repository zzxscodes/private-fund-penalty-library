# -*- coding: utf-8 -*-
"""D3 补录：把《证券投资顾问业务暂行规定》《关于加强对利用"荐股软件"从事证券
投资咨询业务监管的暂行规定》生成为 06 目录下的 Word，并登记到「配套法规依据清单」。

为什么走这条路（2026-10-30 实测）：
  * 中国证监会法规库 neris.csrc.gov.cn 是前端渲染页，urllib 直接取只拿到 230 字导航壳，
    正文取不到；北大法宝/公众号等二手源不作权威依据。
  * 上海证监局公开的《中国证监会决定修改的规范性文件》（附件1）**同时载有这两部法规的
    2020 年修订后全文**（PDF 原生文字层，可直接抽取），是证监会系统内的权威公开源。
  * 抽取时按标题与"第X条 / 一、二、…"边界切片，只取目标法规段落，避免把同 PDF 的
    其它法规混进来；再做「断行合并」，还原段落（PDF 每行 30 字左右，不合并会碎成一行一段）。

用法：python fetch_laws_advisory.py
"""
import os, re, sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

# fetch_laws 在 **import 期** 就 argparse 并强校验 root（空值直接 error 退出），
# 所以必须先把本脚本的 root 解析出来、回填 sys.argv，再 import。顺序不能倒。
_ROOT = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] and not sys.argv[1].startswith("-") else ""
if not _ROOT:
    sys.exit("用法：python fetch_laws_advisory.py <06_以上各类所有案例的配套法规依据 目录>")
sys.argv = ["fetch_laws.py", _ROOT]
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fetch_laws as F  # noqa: E402

SRC = ("https://www.csrc.gov.cn/shanghai/c105565/c1555523/1555523/files/"
       "%E9%99%84%E4%BB%B61_%E4%B8%AD%E5%9B%BD%E8%AF%81%E7%9B%91%E4%BC%9A%E5%86%B3%E5%AE%9A"
       "%E4%BF%AE%E6%94%B9%E7%9A%84%E8%A7%84%E8%8C%83%E6%80%A7%E6%96%87%E4%BB%B6.pdf")
SUB = "02_证监会部门规章与规范性文件"


def merge_paras(lines):
    """把 PDF 的硬断行合并成自然段：行尾无句读且下一行不是条文/项号开头 → 续行。"""
    out = []
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        if not out:
            out.append(ln)
            continue
        start_new = re.match(r"^(第[一二三四五六七八九十百]+条|[一二三四五六七八九十]+、|"
                             r"（[一二三四五六七八九十]+）|\([一二三四五六七八九十]+\)|附则|总则)", ln)
        prev_end = out[-1].endswith(("。", "；", "：", "！", "？", "”", "）", "》"))
        if start_new or prev_end:
            out.append(ln)
        else:
            out[-1] += ln
    # 排版校正：PDF 抽取把「第一条」与条文正文粘成「第一条为了规范…」，
    # 法规原文惯例是「第一条　为了规范…」，补一个全角空格，便于阅读与引用。
    fixed = []
    for ln in out:
        ln = re.sub(r"^(第[一二三四五六七八九十百]+条)(?=[^\s　])", r"\1　", ln)
        ln = re.sub(r"^(第[一二三四五六七八九十]+章)(?=[^\s　])", r"\1　", ln)
        fixed.append(ln)
    return fixed


def main():
    root = _ROOT
    outdir = os.path.join(root, SUB)
    os.makedirs(outdir, exist_ok=True)

    print("下载权威源 PDF ...", flush=True)
    raw = F.fetch_pdf_text(SRC)
    lines = F.clean_lines(raw.split("\n"))
    print("  行数 %d" % len(lines))

    # 边界（已实测锁定）：446 标题 / 447-627 法规一 / 628-629 法规二标题 / 630-703 法规二正文
    i1 = next(i for i, l in enumerate(lines) if l.strip() == "证券投资顾问业务暂行规定")
    i2 = next(i for i, l in enumerate(lines) if l.startswith("关于加强对利用") and "荐股软件" in l)
    i2end = next(i for i, l in enumerate(lines) if l.startswith("九、本规定自2013"))
    seg1 = merge_paras(lines[i1 + 1:i2])
    seg2 = merge_paras(lines[i2 + 2:i2end + 1])   # 标题占两行，正文从"为进一步规范…"起

    laws = [
        (dict(name="证券投资顾问业务暂行规定", no="证监会公告〔2020〕66号（2010年10月12日证监会公告〔2010〕27号发布，2020年两次修正）",
              org="中国证券监督管理委员会", date="2020-10-30", eff="2020-10-30", status="现行有效",
              site="中国证监会（上海证监局公开的《中国证监会决定修改的规范性文件》附件1）", url=SRC), seg1),
        (dict(name='关于加强对利用“荐股软件”从事证券投资咨询业务监管的暂行规定', no="证监会公告〔2012〕40号",
              org="中国证券监督管理委员会", date="2012-12-05", eff="2013-01-01", status="现行有效",
              site="中国证监会（上海证监局公开的《中国证监会决定修改的规范性文件》附件1）", url=SRC), seg2),
    ]
    for law, seg in laws:
        dest = os.path.join(outdir, law["name"] + ".docx")
        blocks = F.structure(seg)
        F.build_docx(dest, law, blocks)
        chars = sum(len(t) for _k, t in blocks)
        print("  [OK] %s ｜ 段 %d ｜ 正文 %d 字 ｜ %d 字节"
              % (law["name"], len(blocks), chars, os.path.getsize(dest)))
        if chars < 1500:
            raise SystemExit("正文过短，疑似切片失败：%s" % law["name"])
    print("完成 →", outdir)


if __name__ == "__main__":
    main()
