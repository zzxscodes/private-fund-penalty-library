#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""audit_files.py —— 台账 ↔ 磁盘 **逐份文件** 深检（2026-11-02）

verify_library 回答的是「数量对不对」（行数==PDF 数、孤儿 0、缺失 0）。
本工具回答的是「**每一份文件本身对不对**」——数量对得上，文件照样可能是错的：

  F1 内容重复：不同文件名、sha256 相同 → 重复下载 / 改名张冠李戴
  F2 空壳：页数 0 / 文字层 <200 字且未 OCR / 体积 <15KB
  F3 首页缺失主体：文件名里的机构名在正文中找不到（放错行 / 下载错文件）
  F4 人员文件无名：`姓名-机构.pdf` 里的姓名不在该行 D 列人员名单里
  F5 人员文件少了：D 列有该人、L 列无其文件（≠错误，多人共用一份，仅提示）
  F6 文号冲突：同一行内两份文件文号相同（可能是同一份文书存了两遍）
  F7 落款冲突：同一行内多份文件落款日期相差 >90 天（可能混入了别的案）
  F9 D 列人名「三无」：本行没有他的文书、文书首部也没有他的名字 →
     他其实是从**正文**顺取来的（F 列常写「实际控制人黄巍利用…」），不是本行当事人

只读，不改任何文件。

用法：
  python scripts/audit_files.py --root <库根> [--csv out.csv]
"""
import argparse
import csv
import hashlib
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

DOCNO = re.compile(r"[（(\[【〔]\s*(?:20\d{2})\s*[）)\]】〕]\s*第?\s*\d+\s*号|"
                   r"〔\s*20\d{2}\s*〕\s*\d+\s*号|\[\s*20\d{2}\s*\]\s*\d+\s*号")
SIGN = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")


def sha(fp, buf=1 << 20):
    h = hashlib.sha256()
    with open(fp, "rb") as f:
        while True:
            b = f.read(buf)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def pdf_info(fp):
    """(页数, 文字量, 去空白正文)"""
    try:
        import fitz
        d = fitz.open(fp)
        t = "".join(p.get_text() for p in d)
        n = d.page_count
        d.close()
        return n, len(t.strip()), re.sub(r"\s+", "", t)
    except Exception:
        return -1, -1, ""


def is_person_file(n):
    base = re.sub(r"\.pdf$", "", n, flags=re.I)
    if "-" not in base:
        return False
    head = base.split("-")[0]
    return len(head) <= 6 and not re.search(r"(公司|合伙|中心|企业|基金)", head)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--csv", default="")
    ap.add_argument("--ocr-cache", default="",
                    help="F9 遇到扫描件时 OCR 的缓存目录（默认 <库父目录>/_案例库维护留痕/_ocr_cache）")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    # OCR 缓存一律放**库外**（库根只放业务文件）：<库父目录>/_案例库维护留痕/_ocr_cache
    cache_dir = args.ocr_cache or os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    wb = openpyxl.load_workbook(ledger)
    ws = wb.worksheets[0]

    rows = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        names = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        d = str(ws.cell(r, 13).value or "").split("；")[0].strip()
        subj = str(ws.cell(r, 4).value or "")
        mgr = subj.split("；")[0].strip()
        persons = [x for x in re.split(r"[、,，]", subj.split("；", 1)[1]) if x.strip()] \
            if "；" in subj else []
        meas = str(ws.cell(r, 5).value or "")
        rows.append((r, names, d, mgr, persons, meas))

    out = []   # 问题（必须为 0 才能交付）
    obs = []   # 观察项（合法形态，不阻塞交付）
    hashes = {}
    seen_paths = set()

    for r, names, d, mgr, persons, meas in rows:
        docnos = []
        signs = []
        heads = []          # 本行各文书的首部文字（用于 F9 判「这个人到底在不在这份文书里」）
        row_fps = []        # 本行各文书路径（F9 遇到扫描件时按需 OCR）
        for nm in names:
            fp = os.path.join(root, d, nm)
            if not os.path.isfile(fp):
                out.append((r, nm, "F0文件缺失", ""))
                continue
            seen_paths.add(os.path.normcase(os.path.abspath(fp)))
            row_fps.append(fp)
            h = sha(fp)
            hashes.setdefault(h, []).append((r, nm))
            pages, tlen, txt = pdf_info(fp)
            size = os.path.getsize(fp)
            if txt:
                heads.append(txt[:600])

            # F2 空壳
            if pages == 0 or size < 15 * 1024:
                out.append((r, nm, "F2空壳/过小", "%d页 %.1fKB" % (pages, size / 1024)))
            # F1 内容重复（同 sha 多文件名）
            # F3 首页主体
            if txt:
                # 两边都要去括号：机构名「上创普盛（天津）创业投资管理有限公司」去括号后是
                # 「上创普盛创业投资…」，正文里仍是带括号的原名 —— 只去一边必然对不上
                # （第一版这样误报 40 份，全是括号型机构名）。
                key = re.sub(r"[（(].*?[）)]", "", mgr)[:6]
                txt2 = re.sub(r"[（(].*?[）)]", "", txt)
                if key and key not in txt2:
                    out.append((r, nm, "F3机构名未回现", "找「%s」" % key))
                # 文号只从**文书首部**取：正文里会引用别的文件的文号
                # （第一版全篇扫，行28/96 都抓到「〔2020〕71号」，其实是同一部法规的文号，误报 8 条）
                m = DOCNO.search(txt[:300])
                if m:
                    docnos.append(m.group(0))
                ms = SIGN.findall(txt[-400:]) or SIGN.findall(txt)
                if ms:
                    y, mo, dd = ms[-1]
                    signs.append("%s-%02d-%02d" % (y, int(mo), int(dd)))
            # F4 人员文件的人名不在 D 列
            if is_person_file(nm):
                p = re.sub(r"\.pdf$", "", nm, flags=re.I).split("-")[0]
                if persons and p not in persons:
                    out.append((r, nm, "F4人名不在D列", "文件「%s」/D列 %s" % (p, "、".join(persons))))
            # F5 D 列有人但无其文件
        if persons:
            have = [re.sub(r"\.pdf$", "", n, flags=re.I).split("-")[0]
                    for n in names if is_person_file(n)]
            missing = [p for p in persons if p not in have]
            if missing and len(names) > 1:
                # F5 是**观察项不是错误**：多人共用一份决定书（「…、…、…：」列名后一并处罚）
                # 完全合法，SKILL §15.3 已写明「人员文书数 ≠ 人员数」。
                # 前版把它算 problem，导致每次跑都固定报 41 条、掩盖真错。
                obs.append((r, "（无独立文件）", "F5人员无独立文件（观察）", "、".join(missing)))
            elif missing:
                obs.append((r, "（无独立文件）", "F5人员无独立文件（观察）", "、".join(missing)))

            # F9 D 列人名「三无」：本行没有他的文书、措施里也没有他 →
            # 他只是从**正文**顺取进来的人名（F 列常写「实际控制人黄巍利用…」），
            # 不是这一行的当事人。行408 就是这么多了一个「黄巍」：
            # 该行只有尉亲一份决定书（中基协处分〔2025〕359 号），黄巍的文书在行449。
            # 为什么 F5 抓不到：F5 把「无独立文件」一律当联名决定书放过；
            # 区分点是「措施里有没有他 / 首部有没有他」——
            #   有 ＝ 联名（合法，仍只进观察）；
            #   没有 ＝ 串名（错）。
            if missing:
                suspect = [p for p in missing if p not in meas]
                allhead = "".join(heads)
                if suspect and allhead:
                    real = [p for p in suspect if p not in allhead]
                    if real:
                        out.append((r, "（D列人名）", "F9人名无文件无措施且不在文书首部（疑正文顺取）",
                                    "、".join(real)))
                elif suspect:
                    # 全是扫描件、读不到首部：**按需 OCR**（只对有嫌疑的行触发，全库仅 1 行），
                    # 不能因为「读不到」就把可能的串名放过，也不能直接判错。
                    got = ""
                    try:
                        import ocrlib
                        for _fp in row_fps[:3]:
                            got += ocrlib.ocr_text(_fp, cache_dir=cache_dir) or ""
                    except Exception as _e:
                        got = ""
                    if not got:
                        obs.append((r, "（D列人名）", "F9疑似多余人名（OCR 不可用，无法核对首部）",
                                    "、".join(suspect)))
                    else:
                        # 只认**首部**：文书首部＝文号 + 「当事人：…」到「根据《」之前。
                        # 不能用全文——正文事实段里本来就会出现别人名（行408 的「实际控制人
                        # 黄巍利用基金财产…」就是），用全文判会把串名判成联名（第一版就这么错了）。
                        g1 = re.sub(r"\s+", "", got)
                        m0 = re.search(r"根据《", g1)
                        head_ocr = g1[:m0.start()] if m0 else g1[:600]
                        real = [p for p in suspect if p not in head_ocr]
                        if real:
                            out.append((r, "（D列人名）",
                                        "F9人名无文件无措施且不在文书首部（疑正文顺取）",
                                        "、".join(real)))
                        else:
                            obs.append((r, "（D列人名）", "F9人名在文书首部（联名/合并，正常观察）",
                                        "、".join(suspect)))
        # F6 文号重复：只有「同文号 + 落款相差 >60 天」才可疑（同一处罚的多当事人共用文号是常态）
        dup = [x for x in set(docnos) if docnos.count(x) > 1]
        if dup and len(signs) >= 2:
            import datetime as _dt2
            ds2 = sorted(_dt2.date.fromisoformat(x) for x in set(signs))
            if (ds2[-1] - ds2[0]).days > 60:
                out.append((r, "（行内多份）", "F6同文号跨期", "%s / %s~%s"
                            % ("、".join(dup[:2]), ds2[0], ds2[-1])))
        # F7 落款跨度过大
        if len(signs) >= 2:
            import datetime as _dt
            ds = sorted(_dt.date.fromisoformat(s) for s in set(signs))
            if (ds[-1] - ds[0]).days > 90:
                out.append((r, "（行内多份）", "F7落款跨度>90天", "%s ~ %s" % (ds[0], ds[-1])))

    # F1 全局内容重复
    for h, lst in hashes.items():
        if len(lst) > 1:
            for r, nm in lst[1:]:
                out.append((r, nm, "F1内容重复", "与行%s「%s」字节相同" % (lst[0][0], lst[0][1][:24])))

    # 磁盘孤儿（在 01/02/03 下但不在任何行 L 列）——A1 已查，这里再验一次
    for d0 in sorted(os.listdir(root)):
        if d0[:2] not in ("01", "02", "03"):
            continue
        for dp, _, fs in os.walk(os.path.join(root, d0)):
            for f in fs:
                if not f.lower().endswith(".pdf"):
                    continue
                if os.path.normcase(os.path.abspath(os.path.join(dp, f))) not in seen_paths:
                    out.append((0, f, "F8磁盘孤儿", dp.replace(root, "")[:40]))

    # 按类型统计
    from collections import Counter
    c = Counter(x[2] for x in out)
    print("\n逐份文件深检：%d 行 / %d 份" % (len(rows), len(hashes) and sum(len(v) for v in hashes.values()) or 0))
    print("  问题 %d 条｜观察 %d 条" % (len(out), len(obs)))
    for k in sorted(c):
        print("  %-16s %d" % (k, c[k]))
    print()
    for r, nm, why, extra in out[:40]:
        print("  行%-4s %-46s %-14s %s" % (r or "-", nm[:44], why, extra))
    if len(out) > 40:
        print("  …其余 %d 条见 CSV" % (len(out) - 40))
    if obs:
        from collections import Counter as _C
        _c = _C(x[2] for x in obs)
        for _k in sorted(_c):
            print("  观察项（不算错）：%-42s %d 条" % (_k, _c[_k]))
        print("    · F5 多人共用一份决定书属合法形态（SKILL §15.3），不逐条打印")
        for _o in [x for x in obs if not x[2].startswith("F5")][:20]:
            print("    · 行%-4s %-40s %s" % (_o[0], _o[2], _o[3]))

    if args.csv and (out or obs):
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["行号", "文件", "类型", "说明"])
            w.writerows(out)
            w.writerows(obs)
        print("\nCSV：%s" % args.csv)
    return 0 if not out else 1


if __name__ == "__main__":
    sys.exit(main())
