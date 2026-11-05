# -*- coding: utf-8 -*-
"""已入库案例「平台更新」回查（用户 2026-10-28 关注点）。

为什么需要：
  滑动窗口只抓「新挂网」条目，且入库判重只在「同(来源,主体,落款)」命中时才并入原行。
  这导致两类"更新"被漏掉 / 误处理：
    ① 平台对旧案「原地替换/更正」文书（同链接、新内容，如补登申辩采纳说明）
       → 发布日不在前向窗口 + 同键跳过，永远发现不了；
    ② 旧案后续出「复议/补充/变更决定」（新文号、新链接、新日期）
       → 当前会新增重复行，而非更新原案。
  本脚本把"对已入库案例回查平台"变成可执行的固定能力。

两种模式：
  【默认】链接回查：对库内全部（或 --sample 抽样）案例的原文链接做
    「可达性 + 内容哈希」比对。404/异常 → 死链候选；哈希变化 → 内容疑似更新候选。
    首跑只建基线（不报更新）。全部网络操作节流 + 容错，绝不崩、绝不损坏台账。
  【--deep】在默认基础上，额外重扫平台（宽窗口，复用 run_update 的 crawl→build），
    对抓到的每个条目做「(来源,主体)后续决定检测」：命中已入库同主体且带后续信号
    （变更/撤销/更正/补充/复议/维持/诉讼/针对原决定）→ 并入原行；歧义 → 候选。

用法：
  python update_scan.py --root "<库根>"                       # 默认：全量链接回查
  python update_scan.py --root "<库根>" --sample 5            # 仅抽样 5 条（调试）
  python update_scan.py --root "<库根>" --deep --start 2020-01-01  # 深度：重扫+合并
  python update_scan.py --root "<库根>" --deep --dry          # 深度试运行（只报候选，不写台账）
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from datetime import date
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import openpyxl                              # noqa: E402
import ledger_style as LS                     # noqa: E402
import ledger as LG                           # noqa: E402

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

# —— 与 run_update.stage_ingest 保持一致的更新识别规则（固定规则，不复用模型）——
FOLLOWUP_RE = re.compile(
    r"(变更|撤销|更正|补充|修改|重新作出).{0,12}决定"
    r"|原.{0,8}决定书?.{0,14}(变更|撤销|更正|补充|修改)"
    r"|本决定.{0,8}(替代|取代|撤销|变更).{0,8}原决定"
    r"|(复议|复核).{0,14}(决定|维持|撤销|变更)"
    r"|维持.{0,6}原(决定|处罚)"
    r"|针对.{0,18}(原|上述|前述).{0,18}决定"
    r"|行政(复议|诉讼).{0,10}(决定|维持|撤销)")
CONTENTCHANGE_RE = re.compile(
    r"(变更|撤销|更正|补充|修改|重新作出).{0,12}决定"
    r"|原.{0,8}决定书?.{0,14}(变更|撤销|更正|补充|修改)"
    r"|本决定.{0,8}(替代|取代|撤销|变更).{0,8}原决定")
SUPERSEDE_FIELDS = ("title", "excerpt", "text", "grounds", "measures", "defense", "category")


# ----------------------------------------------------------------- 台账读取
def _open(root):
    wb = openpyxl.load_workbook(os.path.join(root, LS.LEDGER_NAME))
    return wb, wb[LS.SHEET_CASE]


def _split_multi(s):
    return [x.strip() for x in str(s or "").split("；") if x.strip()]


def _cell(ws, r, k):
    v = ws.cell(r, LS.COL[k] + 1).value
    return str(v).strip() if v is not None else ""


def load_cases(root):
    """读台账 → 案例列表（含行号、来源、主体、落款、链接、文件名、路径、申辩、类别）。"""
    _, ws = _open(root)
    cases = []
    for r in range(2, ws.max_row + 1):
        src = _cell(ws, r, "source")
        sub = _cell(ws, r, "subject")
        if not src or not sub:
            continue
        cases.append({
            "row": r, "source": src, "subject": sub,
            "decided": _cell(ws, r, "decided"),
            "urls": _split_multi(_cell(ws, r, "urls")),
            "filenames": _split_multi(_cell(ws, r, "filenames")),
            "paths": _split_multi(_cell(ws, r, "paths")),
            "defense": _cell(ws, r, "defense"),
            "category": _cell(ws, r, "category"),
        })
    return cases


# ----------------------------------------------------------------- 链接回查
def _norm_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _fetch(url, timeout=30):
    """GET 原文链接，返回 (ok, bytes_or_error)。任何异常都返回 ok=False，绝不抛出。"""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return True, r.read(8_000_000)
    except Exception as e:                       # 404/超时/DNS/SSL/解码…全部兜底
        return False, ("%s:%s" % (type(e).__name__, e))


def recheck_urls(cases, work, sample=0, delay=0.25):
    """对案例的原文链接做可达性 + 内容哈希比对。返回统计 + 候选列表。"""
    state_path = os.path.join(work, "updscan_state.json")
    state = {}
    if os.path.exists(state_path):
        try:
            state = json.load(open(state_path, encoding="utf-8"))
        except Exception:
            state = {}
    new_state = {}
    dead, changed, checked = [], [], 0
    pool = cases if not sample else cases[:sample]

    for c in pool:
        for u in c["urls"]:
            if not u or u.startswith("待补"):
                continue
            checked += 1
            ok, payload = _fetch(u, timeout=30)
            if not ok:
                new_state[u] = {"status": "error", "err": str(payload)[:120]}
                dead.append({"row": c["row"], "subject": c["subject"], "url": u,
                             "issue": "链接不可达/异常", "detail": str(payload)[:120]})
                time.sleep(delay)
                continue
            h = _norm_hash(payload)
            prev = state.get(u)
            if prev is None:
                new_state[u] = {"status": "ok", "hash": h}     # 首跑建基线
            else:
                new_state[u] = {"status": "ok", "hash": h}
                if prev.get("status") != "ok":
                    dead.append({"row": c["row"], "subject": c["subject"], "url": u,
                                 "issue": "链接恢复（此前不可达）", "detail": ""})
                elif prev.get("hash") != h:
                    changed.append({"row": c["row"], "subject": c["subject"], "url": u,
                                    "issue": "内容疑似更新（哈希变化）", "detail": ""})
            time.sleep(delay)

    json.dump(new_state, open(state_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    baseline_new = sum(1 for u in new_state if u not in state)
    return {"checked": checked, "dead_links": dead, "content_changed": changed,
            "baseline_new": baseline_new}, new_state


# ----------------------------------------------------------------- 深度回查
def _looks_followup(rec):
    if str(rec.get("supersedes") or "").strip():
        return True
    blob = " ".join(str(rec.get(k) or "") for k in SUPERSEDE_FIELDS)
    return bool(FOLLOWUP_RE.search(blob))


def _is_content_change(rec):
    if str(rec.get("supersedes") or "").strip():
        return True
    blob = " ".join(str(rec.get(k) or "") for k in SUPERSEDE_FIELDS)
    return bool(CONTENTCHANGE_RE.search(blob))


def _merge_into_row(ws, r, rec, src_root, dry, refresh_core=True):
    """把抓到的新文书并入已入库行：拷贝文件、追加文件名/路径/链接、按信号刷新核心列、记备注。"""
    C = LS.COL
    dst_dir = os.path.join(src_root, rec.get("source_dir", ""), rec.get("category", ""))
    os.makedirs(dst_dir, exist_ok=True)

    def _multi(col, add):
        old = _split_multi(_cell(ws, r, col))
        for x in add:
            if x and x not in old:
                old.append(x)
        ws.cell(r, C[col] + 1).value = "；".join(old)

    nm, pt, ur = [], [], []
    for f in rec.get("files") or []:
        src, name = f.get("src") or "", f.get("name") or ""
        if not name:
            continue
        dst = os.path.join(dst_dir, name)
        try:
            if src.startswith("http"):
                if not os.path.exists(dst):
                    _download(src, dst)
            elif src and os.path.isfile(src):
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
            else:
                continue
        except Exception as e:
            print("    [WARN] 拷贝失败 %s: %s" % (name, e))
            continue
        nm.append(name)
        pt.append(os.path.dirname(os.path.relpath(dst, src_root)).replace("\\", "/"))
        ur.append((rec.get("urls") or [""])[0])
    _multi("filenames", nm)
    _multi("paths", pt)
    _multi("urls", ur)
    old_dec = _cell(ws, r, "decided")
    if refresh_core:
        for k in ("measures", "grounds", "evidence", "legal_basis"):
            v = "；".join(x for x in (rec.get(k) or []) if x)
            if v:
                ws.cell(r, C[k] + 1).value = v
    if rec.get("date_decided"):
        ws.cell(r, C["decided"] + 1).value = rec["date_decided"]
    if rec.get("publish_date"):
        ws.cell(r, C["publish"] + 1).value = rec["publish_date"]
    kind = "内容变更" if refresh_core else "后续决定（维持/复议/诉讼，原决定保留）"
    note = "原落款%s → 本决定（%s）更新；新落款%s" % (old_dec, kind, rec.get("date_decided"))
    dcol = C["defense"] + 1
    prev = _cell(ws, r, "defense")
    ws.cell(r, dcol).value = (prev + "；" + note) if prev and prev != "无" else note
    print("  ~ 合并更新 [行%d] %s：追加 %d 份，落款 %s → %s（%s）%s"
          % (r, rec.get("subject"), len(nm), old_dec, rec.get("date_decided"), kind,
             "（dry）" if dry else ""))


def _download(url, dst):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=90) as r:
        b = r.read()
    if len(b) < 1024:
        raise IOError("下载过小：%s" % url)
    open(dst, "wb").write(b)
    return True


def deep_scan(root, work, start, end, dry):
    """--deep：重扫平台（宽窗口，复用 run_update crawl→build）→ 读 records.json
    → 对每个条目做(来源,主体)后续决定检测 → 合并/候选。"""
    os.makedirs(work, exist_ok=True)
    print("  [deep] crawl→build（宽窗口 %s ~ %s）…" % (start, end))
    rc = subprocess.run(
        [sys.executable, os.path.join(HERE, "run_update.py"),
         "--root", root, "--work-dir", work, "--stages", "crawl,build",
         "--start", start, "--end", end, "--no-laws", "--no-cleanup"])
    if rc.returncode != 0:
        print("  [deep][FATAL] crawl→build 失败，跳过深度回查")
        return {"error": "crawl_build_failed"}
    recs = json.load(open(os.path.join(work, "records.json"), encoding="utf-8"))
    print("  [deep] 抓到 %d 条，做(来源,主体)后续决定检测…" % len(recs))

    wb, ws = _open(root)
    cases = load_cases(root)
    rows_by_sm = {}
    for c in cases:
        rows_by_sm.setdefault((c["source"], c["subject"]), []).append(c["row"])
    merged, candidates = 0, []
    for rec in recs:
        src = str(rec.get("source") or "").strip()
        sub = str(rec.get("subject") or "").strip()
        if not src or not sub:
            continue
        same = [r for r in rows_by_sm.get((src, sub), [])
                if _cell(ws, r, "decided") != str(rec.get("date_decided") or "").strip()]
        if not same:
            continue                                    # 库内无同主体 → 属新增案（由日常 ingest 处理）
        if _looks_followup(rec):
            if len(same) == 1:
                if not dry:
                    _merge_into_row(ws, same[0], rec, root, dry, _is_content_change(rec))
                    merged += 1
                else:
                    print("  [dry] 将合并 [行%d] %s" % (same[0], sub))
                    merged += 1
            else:
                candidates.append({"subject": sub, "date_decided": rec.get("date_decided"),
                                   "source": src, "rows": same,
                                   "reason": "同(来源,主体)多条候选，待人工确认"})
        else:
            try:
                nd = str(rec.get("date_decided") or "")[:10]
                od = _cell(ws, same[0], "decided")[:10]
                near = (len(nd) == 10 and len(od) == 10 and
                        abs((date.fromisoformat(nd) - date.fromisoformat(od)).days) <= 120)
            except Exception:
                near = False
            if near:
                candidates.append({"subject": sub, "date_decided": rec.get("date_decided"),
                                   "source": src, "rows": same,
                                   "reason": "同(来源,主体)且落款相近但无后续信号，疑似更新，待人工确认"})
    if merged and not dry:
        wb.save(os.path.join(root, LS.LEDGER_NAME))
        LG.cmd_sort(Path(os.path.abspath(root)))
    return {"merged": merged, "candidates": candidates}


# ----------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description="已入库案例·平台更新回查")
    ap.add_argument("--root", required=True)
    ap.add_argument("--work", default="", help="中间产物目录，默认 <root>_updscan")
    ap.add_argument("--deep", action="store_true", help="额外重扫平台做后续决定检测+合并")
    ap.add_argument("--start", default="2020-01-01", help="--deep 重扫起点")
    ap.add_argument("--end", default="", help="--deep 重扫终点，默认今天")
    ap.add_argument("--sample", type=int, default=0, help="仅抽样前 N 条案例回查（调试）")
    ap.add_argument("--delay", type=float, default=0.25, help="链接回查节流间隔（秒）")
    ap.add_argument("--no-net", action="store_true", help="跳过网络回查（仅结构校验）")
    ap.add_argument("--dry", action="store_true", help="--deep 试运行：只报候选/合并，不写台账")
    ap.add_argument("--json", default="", help="报告 JSON 路径")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    work = os.path.abspath(a.work or (root + "_updscan"))
    os.makedirs(work, exist_ok=True)
    end = a.end or date.today().isoformat()

    cases = load_cases(root)
    print("台账案例：%d 条" % len(cases))

    report = {"root": root, "cases": len(cases), "mode": ("deep" if a.deep else "url-recheck")}

    if not a.no_net:
        rr, _state = recheck_urls(cases, work, sample=a.sample, delay=a.delay)
        report["url_recheck"] = {"checked": rr["checked"],
                                 "dead_links": rr["dead_links"],
                                 "content_changed": rr["content_changed"]}
        print("\n== 链接回查 ==")
        print("  检查 %d 条链接 ｜ 死链/不可达 %d ｜ 内容疑似更新 %d"
              % (rr["checked"], len(rr["dead_links"]), len(rr["content_changed"])))
        for d in rr["dead_links"][:20]:
            print("   · 死链 [%d] %s %s" % (d["row"], d["subject"][:30], d["issue"]))
        for d in rr["content_changed"][:20]:
            print("   · 更新 [%d] %s %s" % (d["row"], d["subject"][:30], d["issue"]))
        if rr["dead_links"] or rr["content_changed"]:
            json.dump({"dead_links": rr["dead_links"],
                       "content_changed": rr["content_changed"]},
                      open(os.path.join(work, "update_candidates.json"), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            print("  候选明细 → %s" % os.path.join(work, "update_candidates.json"))
    else:
        print("  （--no-net，跳过网络回查）")

    if a.deep:
        dr = deep_scan(root, work, a.start, end, a.dry)
        report["deep"] = dr
        print("\n== 深度回查 ==")
        print("  合并更新 %d 行 ｜ 待人工候选 %d 条"
              % (dr.get("merged", 0), len(dr.get("candidates", []))))
        for c in dr.get("candidates", [])[:20]:
            print("   · 候选 [%s] %s：%s" % (c.get("source"), c.get("subject", "")[:30], c.get("reason")))

    out = a.json or os.path.join(work, "updscan_report.json")
    json.dump(report, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\n报告 → %s" % out)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.exit(main())
