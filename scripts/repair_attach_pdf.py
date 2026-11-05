# -*- coding: utf-8 -*-
"""附件型决定书「空壳 PDF」就地修复（2026-10-30）。

背景（实测缺陷，必须修）：
  证监局约 1 成决定书正文不在网页内联 HTML，而在网页挂的 .docx/.doc/.wps/.pdf 附件里；
  网页本身只有「索引号 / 发文日期 / 名称 / 文号」元信息卡片 + 一个附件链接。
  旧实现无差别把**网页**打印成 PDF → 产出「只有导航框架、正文零内容」的空壳 PDF：
  文字层 320~410 字（全是导航栏与元信息），不含任何「经查/当事人/我局决定」标记。
  这批文件在库内看起来正常（有页数、有大小、能打开），**只有通读才发现没正文**，
  属「严重的内容缺失」，不是格式差异。

修法（与旧格式归档同批案例的形态对齐）：
  附件型 → 取附件本体转 PDF（.pdf 直接取、.docx/.doc/.wps 走 Word 导出，保留官方版式），
  正文、文号、落款与官方原件完全一致；不用文本重建、不加网页封面。

用法：
    python repair_attach_pdf.py --root "<案例库根目录>"            # 只体检，不落盘
    python repair_attach_pdf.py --root "<案例库根目录>" --apply    # 执行修复（先自动备份）
    [--only-rows 3,64,81]   只修指定台账行（调试用）
    [--limit N]             最多修 N 份

退出码：0 = 无需修复或全部修复成功；1 = 仍有空壳 / 修复失败。
"""
import argparse, json, os, re, shutil, sys, time
from datetime import datetime
import ledger_style as LS  # noqa: E402  运维目录唯一真源（maint_dir/work_dir/...）

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import fetch_csrc as F  # noqa: E402

LEDGER = "04_综合索引台账【总表】.xlsx"
SHEET = "综合索引台账"


def _split(v):
    """多值列拆分：**只按 `；` / `;`**，绝不按 `、`（源目录名本身含「、」）。"""
    if v is None:
        return []
    s = str(v).strip()
    if not s:
        return []
    return [x.strip() for x in re.split(r"[；;]", s) if x.strip()]


def read_ledger(root):
    from openpyxl import load_workbook
    p = os.path.join(root, LEDGER)
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    head = [str(x or "").strip() for x in rows[0]]
    idx = {h: i for i, h in enumerate(head)}

    def col(*names):
        for n in names:
            if n in idx:
                return idx[n]
        return None

    ci_seq = col("序号")
    ci_sub = col("被处罚机构/人", "被处罚主体", "主体")
    ci_files = col("案例文件", "决定书文件名", "文件名")
    ci_paths = col("存放路径")
    ci_urls = col("原文链接", "链接")
    out = []
    for r in rows[1:]:
        if not r or all(x in (None, "") for x in r):
            continue
        g = lambda i: (r[i] if i is not None and i < len(r) else "")
        out.append({"excel_row": len(out) + 2,
                    "seq": g(ci_seq), "subject": str(g(ci_sub) or ""),
                    "files": _split(g(ci_files)), "paths": _split(g(ci_paths)),
                    "urls": _split(g(ci_urls))})
    return out, head


def is_shell(path):
    """空壳判定：有文字层（≥60 字）却没有任何决定书正文标记。
    无文字层（<60 字）= 扫描图片型决定书，正文在图片里，不在此列。"""
    try:
        import fitz
        d = fitz.open(path)
        t = "".join(p.get_text() for p in d)
        n = d.page_count
        d.close()
    except Exception:
        return False, "无法打开"
    if n < 1:
        return True, "0 页"
    body = re.sub(r"\s+", "", t)
    if len(body) < 60:
        return False, "无文字层（扫描件）"
    if any(k in body for k in F.BODY_MARK):
        return False, "有正文"
    return True, "文字层 %d 字但无正文标记" % len(body)


def rebuild(url, out_pdf, workdir):
    """按 URL 重新生成决定书 PDF。返回 (ok, how, note)。"""
    m = re.search(r"/(c\d{5,})/", url)
    mid = m.group(1) if m else None
    kind = payload = ext = None
    for attempt in (1, 2, 3):
        kind, payload, ext = F.get_content(mid, url)
        if kind:
            break
        time.sleep(2.0 * attempt)
    if not kind:
        return False, "", "正文接口无返回"
    if kind == "attach":
        fn, text = F.save_attach(payload, ext, url, workdir, None)
        if not fn or not os.path.exists(fn):
            return False, "", "附件下载失败"
        if ext == "pdf":
            shutil.copy2(fn, out_pdf)
            how = "attach-pdf"
        else:
            ok, how = F.doc_to_pdf(fn, out_pdf)
        if not F.pdf_has_body(out_pdf):
            return False, how, "附件转 PDF 后仍无正文（附件可能是送达公告/名单）"
        return True, how, ""
    # 内联正文 → 官网原页面打印
    import importlib.util as _ilu
    spec = _ilu.spec_from_file_location("save_decision_pdf",
                                        os.path.join(HERE, "save_decision_pdf.py"))
    sdm = _ilu.module_from_spec(spec)
    spec.loader.exec_module(sdm)
    br = sdm.find_browser()
    sdm.print_to_pdf(br, url, out_pdf)
    if not F.pdf_has_body(out_pdf):
        return False, "web", "网页打印后仍无正文"
    return True, "web", ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--only-rows", default="")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--backup-dir", default="")
    a = ap.parse_args()
    root = os.path.abspath(a.root)
    if not os.path.isdir(root):
        sys.exit("库根不存在：%s" % root)

    rows, head = read_ledger(root)
    only = set(int(x) for x in a.only_rows.split(",") if x.strip().isdigit())
    workdir = os.path.join(LS.work_dir(root), "repair_attach")
    os.makedirs(workdir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = a.backup_dir or os.path.join(LS.maint_dir(root), "_attach_repair_bak_" + stamp)

    cands = []
    for r in rows:
        if only and int(r["seq"] or 0) not in only:
            continue
        for i, fn in enumerate(r["files"]):
            d = r["paths"][i] if i < len(r["paths"]) else (r["paths"][0] if r["paths"] else "")
            u = r["urls"][i] if i < len(r["urls"]) else (r["urls"][0] if r["urls"] else "")
            if not fn.lower().endswith(".pdf"):
                continue
            fp = os.path.join(root, d, fn) if d else os.path.join(root, fn)
            if not os.path.exists(fp):
                continue
            shell, why = is_shell(fp)
            if shell:
                cands.append({"row": r["seq"], "subject": r["subject"], "file": fn,
                              "path": fp, "rel": os.path.join(d, fn), "url": u, "why": why})

    print("台账 %d 行 ｜ 检出空壳 PDF %d 份" % (len(rows), len(cands)))
    if not cands:
        print("无需修复。")
        return 0
    if a.limit:
        cands = cands[:a.limit]

    stats = {"ok": 0, "fail": 0, "skip": 0, "how": {}}
    fails = []
    for k, c in enumerate(cands, 1):
        tag = "[%d/%d] 行%s %s" % (k, len(cands), c["row"], c["file"][:42])
        if not c["url"]:
            print(tag, "→ SKIP 无原文链接"); stats["skip"] += 1
            fails.append({**c, "err": "无原文链接"}); continue
        tmp = os.path.join(workdir, F.hashlib_sha1(c["url"]) + ".pdf")
        if os.path.exists(tmp):
            os.remove(tmp)
        try:
            ok, how, note = rebuild(c["url"], tmp, workdir)
        except Exception as e:
            ok, how, note = False, "", "%s: %s" % (type(e).__name__, str(e)[:80])
        if not ok:
            print(tag, "→ FAIL", note); stats["fail"] += 1
            fails.append({**c, "err": note}); continue
        if not a.apply:
            print(tag, "→ OK(%s) 待落盘" % how); stats["ok"] += 1
            stats["how"][how] = stats["how"].get(how, 0) + 1
            continue
        dst_bak = os.path.join(bak, c["rel"])
        os.makedirs(os.path.dirname(dst_bak), exist_ok=True)
        if not os.path.exists(dst_bak):
            shutil.copy2(c["path"], dst_bak)
        shutil.move(tmp, c["path"])
        print(tag, "→ OK(%s) %d 字节" % (how, os.path.getsize(c["path"])))
        stats["ok"] += 1
        stats["how"][how] = stats["how"].get(how, 0) + 1

    print("\n== 汇总 ==  %s" % json.dumps(stats, ensure_ascii=False))
    if a.apply and stats["ok"]:
        print("原文件已备份至：%s" % bak)
    if fails:
        print("\n未修复 %d 份：" % len(fails))
        for f in fails[:30]:
            print("  行%s %s —— %s" % (f["row"], f["file"][:44], f["err"]))
        rp = os.path.join(LS.work_dir(root), "repair_attach_failed.json")
        json.dump(fails, open(rp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("清单：%s" % rp)
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
