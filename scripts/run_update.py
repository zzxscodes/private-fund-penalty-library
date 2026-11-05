# -*- coding: utf-8 -*-
"""案例库滑动窗口增量更新驱动（v3 格式，一年窗口可持续滚动）。

为什么需要它（用户 2026-10-27 明确）：
  「我想要的一年时间范围是之前那样滑动窗口的，或者说不断更新加入当天扩展的，
   **而不是用旧库去重建案例库**。」
  用户进一步说明了**为什么要滑动窗口**：
  「技能在使用者使用时一般是获取**从不同类别官网最早可获取案例到当日**的案例，
   所以作为定时任务去更新维护案例库时就要有这样的能力；本来这个案例库技能就是一个
   按时间或其他条件**收集、筛选、分类、存放、更新、维护**的业务技能。」
⇒ 本技能的价值不在"一次性建库"，而在**可持续维护**。两种用法：

  ① 首次建库（--full，或台账为空时自动）：抓取窗口起点放到极早（1990-01-01），
     把各平台**最早可获取**的文书一路抓到当日 —— 即"从最早到当日"的全量收集。
  ② 日常增量（定时任务）：只抓**上次最新挂网日往前 overlap 天**到当日的新增，
     入库后把落款日滚出窗口的行移出。这就是"滑动窗口/每天扩展"的能力。

  无论哪种，都**绝不复制旧库、绝不重建台账**；旧库/归档库/冻结快照一律只读。

窗口口径（重要）：
  * **库窗口**按决定书【落款日（处罚时间）】裁剪 = [--days] 天往前（默认 365）；
    --full 时起点取到最早，等于不做裁剪。
  * **抓取窗口**：全量模式 = [库窗口起点, 当日]；增量模式 = [上次最新挂网日 - overlap, 当日]。
    抓回来的记录再按 date_decided 落回库窗口，超窗的不入库（登记到不收录清单）。

六个阶段（可单独重跑，幂等）：
  crawl     三信源增量抓取（证监局 / 中基协 / 沪深北交易所）
  build     筛选 → 分类 → 机构+人员合并 → 抽三核心列 → 命名
  ingest    落 PDF 到 01/02/03/<子类>/ + 台账增量追加（判重后只加新案）
  laws      06 配套法规自动抓取（缺件才下载，幂等；首次建库无需手工单跑）
  finalize  06 配套法规清单回填 + 05 重点典型案例交叉归档
  verify    verify_library.py 全量校验
  retire    把落款日超出窗口的台账行 + PDF（含 05 副本）移出库外备份区
  cleanup   终态整洁收尾：库内非白名单杂物移出库外留痕区（只移不删）

用法（全部路径走命令行参数，脚本内无任何本机绝对路径）：
  python run_update.py --root "<案例库根>"                       # 一年窗口，跑全流程
  python run_update.py --root "<案例库根>" --days 365 --end 2026-10-27
  python run_update.py --root "<案例库根>" --stages crawl,build  # 只跑到构建
  python run_update.py --root "<案例库根>" --stages ingest,finalize,verify,retire
  python run_update.py --root "<案例库根>" --no-exchange --lag-days 45
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import date, timedelta

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

PY = sys.executable or "python"
# 2026-10-28 用户要求 #3：首次建库不能再让使用者单独跑 fetch_laws.py。
#  ⇒ 新增 laws 阶段（06 目录为空/缺件时自动补下载，幂等：已存在即跳过）。
# 2026-10-28 用户要求 #1：技能执行后库内终态必须干净。
#  ⇒ 新增 cleanup 阶段（收尾把非白名单杂物移出库外留痕区，只移不删）。
ALL_STAGES = ["crawl", "build", "ingest", "laws", "finalize", "verify", "retire", "cleanup"]
# updscan 是"已入库案例平台更新回查"阶段（2026-10-28 新增）：默认不进每日流水线，
# 由 --updscan / --deep-updscan 显式开启，避免每日重扫平台/回查链接拖慢日常入库。
VALID_STAGES = ALL_STAGES + ["updscan"]


def _sh(args, cwd=None):
    """串行跑子脚本：失败即抛，避免"某一步静默失败还继续往下跑"。"""
    print("\n$ " + " ".join('"%s"' % a if " " in a else a for a in args))
    r = subprocess.run([PY] + args, cwd=cwd or HERE)
    if r.returncode != 0:
        raise SystemExit("[run_update] 子步骤失败(%d)：%s" % (r.returncode, args[-1]))


def _dump(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    json.dump(obj, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def _load(path, default):
    if not os.path.exists(path):
        return default
    try:
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return default


def d(s):
    return date.fromisoformat(s)


EARLIEST = "1990-01-01"      # 足够早：让平台把「最早可获取」的文书一路返回出来


def earliest_decided(root):
    """读台账里最早的【落款日】——判断库是否已覆盖到窗口起点。

    与 latest_publish 配套使用，防止下面这个坑（2026-10-27 实测）：
    库里只有上次 30 天试建的 42 行，最新挂网日 = 2026-10-23，于是增量锚点把
    **365 天窗口的抓取压成了 11 天**（2026-10-16 ~ 10-27），一年里其余文书整批漏掉，
    而日志看上去完全正常（36 个局都"抓到了"，只是每局只有个位数）。
    ⇒ 锚点只有在「库已覆盖到窗口起点」时才可信。
    """
    import ledger_style as LS
    from openpyxl import load_workbook

    path = os.path.join(root, LS.LEDGER_NAME)
    if not os.path.exists(path):
        return ""
    try:
        ws = load_workbook(path, read_only=True)[LS.SHEET_CASE]
    except Exception:
        return ""
    best = ""
    for row in ws.iter_rows(min_row=2, values_only=True):
        v = str(row[LS.COL["decided"]] or "").strip()[:10]
        if re.match(r"^\d{4}-\d{2}-\d{2}$", v) and (not best or v < best):
            best = v
    return best


def latest_publish(root):
    """读台账里最新的【挂网发布日期】——增量抓取的起点锚点。

    取不到（台账为空/全是"待补"）返回 ""，调用方据此判定"这是首次建库，要抓全历史"。
    """
    import ledger_style as LS
    from openpyxl import load_workbook

    path = os.path.join(root, LS.LEDGER_NAME)
    if not os.path.exists(path):
        return ""
    try:
        ws = load_workbook(path, read_only=True)[LS.SHEET_CASE]
    except Exception:
        return ""
    best = ""
    for row in ws.iter_rows(min_row=2, values_only=True):
        v = str(row[LS.COL["publish"]] or "").strip()[:10]
        if re.match(r"^\d{4}-\d{2}-\d{2}$", v) and v > best:
            best = v
    return best


# ----------------------------------------------------------------- crawl
def stage_crawl(root, work, win_lib, win_crawl, args):
    os.makedirs(work, exist_ok=True)
    cs, ce = win_crawl
    ls, le = win_lib
    failed_src = []          # 抓取脚本异常退出的来源：绝不能让它"沿用上一轮产物"蒙混过关

    # 1) 证监局：36 局 × 行政监管措施/行政处罚决定，按【挂网日】窗口
    if not args.no_csrc:
        _sh([os.path.join(HERE, "fetch_csrc.py"),
             "--start", cs, "--end", ce,
             "--out", os.path.join(work, "csrc.json"),
             "--cache-dir", os.path.join(work, "cache_csrc"),
             "--attach-dir", os.path.join(work, "dl_csrc"),
             "--text-dir", os.path.join(work, "texts_csrc")]
            + (["--bureaus", args.bureaus] if args.bureaus else []))

    # 2) 中基协：受处分机构 + 受处分人员两栏
    if not args.no_amac:
        _sh([os.path.join(HERE, "fetch_amac.py"),
             "--out", os.path.join(work, "amac.json"),
             "--since", cs, "--until", ce,
             "--cache-dir", os.path.join(work, "cache_amac"),
             "--download-dir", os.path.join(work, "dl_amac"),
             "--text-dir", os.path.join(work, "texts_amac")]
            + (["--ocr-cache", args.ocr_cache] if args.ocr_cache else [])
            + (["--max-pages", str(args.max_pages)] if args.max_pages else []))

    # 3) 交易所：上交所/深交所/北交所
    if not args.no_exchange:
        try:
            _sh([os.path.join(HERE, "fetch_exchange.py"),
                 "--start", cs, "--end", ce,
                 "--out", os.path.join(work, "exch.json"),
                 "--cache-dir", os.path.join(work, "cache_exch")])
        except SystemExit as e:
            # 交易所频控/浏览器缺失不应拖垮整轮：记警告后继续（03 天然命中率极低）。
            # 但必须登记成"失败"，否则断言会读到上一轮的 exch.json，
            # 把上轮的频控当成这轮的结论（2026-10-28 实测误报 3 项）。
            failed_src.append("交易所")
            print("[WARN] 交易所抓取失败，跳过 03：%s" % e)

    # 失败来源落盘，交给断言判 FAIL（宁可整轮中止，也不无声沿用旧产物）
    fp = os.path.join(work, "crawl_failed.json")
    if failed_src:
        _dump(fp, failed_src)
    elif os.path.exists(fp):
        os.remove(fp)

    # 抓取产物路径必须是绝对路径，否则 build 换目录后找不到文件（2026-10-27 实测坑）
    for name in ("csrc.json", "amac.json"):
        p = os.path.join(work, name)
        if not os.path.exists(p):
            continue
        obj = _load(p, {"items": []})
        fixed = 0
        for it in obj.get("items", []):
            for k in ("pdf", "local", "text"):
                v = it.get(k)
                if not v or os.path.isabs(v) or v.startswith("http"):
                    continue
                cand = os.path.normpath(os.path.join(work, v))
                if os.path.exists(cand):
                    it[k] = cand
                    fixed += 1
        if fixed:
            _dump(p, obj)
        print("  %s：%d 条（路径转绝对 %d 处）" % (name, len(obj.get("items", [])), fixed))

    # 抓取完整性自证（2026-10-28，固定规则 + 零额外网络）：
    # 抓取层只负责记录事实（平台总数/翻了第几页/为什么停），本步骤用固定规则判定。
    # 不完整就**中止**——绝不把"少抓到"当成"本期无新增"继续入库（这是最隐蔽的数据缺失）。
    man_new = os.path.join(work, "crawl_manifest.json")
    man_prev = os.path.join(work, "crawl_manifest.prev.json")
    if os.path.exists(man_new) and not os.path.exists(man_prev):
        shutil.copy2(man_new, man_prev)      # 上一轮的自证报告 → 作环比基准
    cmd = ["assert_crawl_complete.py", "--work", work, "--out", man_new]
    if os.path.exists(man_prev):
        cmd += ["--prev", man_prev]
    if getattr(args, "max_pages", 0):
        cmd += ["--max-pages", str(args.max_pages)]
    # 本轮显式跳过的来源必须告诉断言，否则它会读到 work 目录里上一轮的残留产物，
    # 把"上轮交易所频控"当成"本轮交易所失败"（2026-10-28 实测误报 2 项 FAIL）。
    skipped = [n for n, flag in (("证监局", args.no_csrc), ("中基协", args.no_amac),
                                 ("交易所", args.no_exchange)) if flag]
    if skipped:
        cmd += ["--skip", ",".join(skipped)]
    try:
        _sh(cmd)
    except SystemExit:
        print("[FATAL] 抓取完整性断言未通过：本轮抓取不完整，已中止（未入库）。\n"
              "        请修好抓取后重跑；切勿把『少抓』当成『本期无新增』。", file=sys.stderr)
        raise
    return 0


# ----------------------------------------------------------------- build
def stage_build(root, work, win_lib, args):
    ls, le = win_lib
    _sh([os.path.join(HERE, "build_pipeline.py"),
         "--csrc-json", os.path.join(work, "csrc.json"),
         "--amac-json", os.path.join(work, "amac.json"),
         "--exch-json", os.path.join(work, "exch.json"),
         "--out", os.path.join(work, "records.json"),
         "--excluded-out", os.path.join(work, "excluded.json"),
         "--audit-out", os.path.join(work, "audit.json")]
        + (["--no-exch"] if args.no_exchange else []))

    # 按【落款日】落回库窗口：抓取窗口多留了 lag，超窗的不入库，但要登记，
    # 否则下一轮会反复抓同一批、反复剔除（看似"每轮都有新增"其实是噪音）。
    recs = _load(os.path.join(work, "records.json"), [])
    tol = 7       # fwrq 字段口径不稳（安徽=挂网日、江苏=落款日），给 7 天容差
    lo = (d(ls) - timedelta(days=tol)).isoformat()
    keep, out_win = [], []
    for r in recs:
        dd = (r.get("date_decided") or "").strip()[:10]
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", dd):
            keep.append(r)                       # 取不到日期先收，交校验兜底
            continue
        if lo <= dd <= le:
            keep.append(r)
        else:
            out_win.append({"source": r.get("source", ""), "subject": r.get("subject", ""),
                            "date_decided": dd, "url": (r.get("urls") or [""])[0],
                            "reason": "窗口外（落款日不在 %s~%s）" % (ls, le)})
    _dump(os.path.join(work, "records.json"), keep)
    excl = _load(os.path.join(work, "excluded.json"), [])
    excl += out_win
    _dump(os.path.join(work, "excluded.json"), excl)
    print("  窗口裁剪：保留 %d 条，超窗 %d 条（库窗口 %s~%s）" % (len(keep), len(out_win), ls, le))
    return 0


# ----------------------------------------------------------------- ingest
def _download(url, dst):
    import urllib.request
    req = urllib.request.Request(url, headers={
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")})
    with urllib.request.urlopen(req, timeout=90) as r:
        b = r.read()
    if len(b) < 1024:
        raise IOError("下载过小：%s" % url)
    open(dst, "wb").write(b)
    return True


def stage_ingest(root, work, args):
    import init_library as IL
    import ledger
    import ledger_style as LS
    from openpyxl import load_workbook
    from pathlib import Path

    root_p = Path(root)
    IL.create_dirs(root_p)
    IL.create_ledger(root_p)
    IL.create_log(root_p)

    recs = _load(os.path.join(work, "records.json"), [])
    if not recs:
        print("  无可入库记录")
        return 0

    wb = load_workbook(os.path.join(root, LS.LEDGER_NAME))
    ws = wb[LS.SHEET_CASE]
    have = ledger._existing_keys(ws)

    # ================= 2026-10-28 新增：同一事项「重新下发决定书」的合并能力 =================
    # 用户明确场景：当事人申辩后，监管重新下发变更/撤销/补充决定书。库内应"更新原案例的
    # 文件与台账"，而不是多出一行（A3 也抓不到，因为处罚时间不同），更不能因同名文件
    # `if not os.path.exists(dst)` 而沿用旧文书、新决定书根本不落盘。
    C = LS.COL
    # 后续决定信号（任一即视为「针对已入库同一事项的更新」→ 并入原行，不新增重复行）
    FOLLOWUP_RE = re.compile(
        r"(变更|撤销|更正|补充|修改|重新作出).{0,12}决定"
        r"|原.{0,8}决定书?.{0,14}(变更|撤销|更正|补充|修改)"
        r"|本决定.{0,8}(替代|取代|撤销|变更).{0,8}原决定"
        r"|(复议|复核).{0,14}(决定|维持|撤销|变更)"
        r"|维持.{0,6}原(决定|处罚)"
        r"|针对.{0,18}(原|上述|前述).{0,18}决定"
        r"|行政(复议|诉讼).{0,10}(决定|维持|撤销)")
    # 内容变更信号（只有这些才刷新三核心列「措施/事由/证据/依据」）。
    # ⚠ 2026-10-30 用户最终裁定（复核/维持类）：**原处分的核心内容必须保留**，
    #   复核（复议）决定只刷新「处罚时间=最终决定日」并在「申辩理由」列追加一条**明确区分的
    #   复核结论**（如 【复核决定·2026-08-25】本会维持原决定），**绝不允许**用"维持《决定书》…"
    #   这种复核文书的空壳表述去覆盖原处分的实质处罚（撤销登记/公开谴责/罚款）。
    #   ⇒ 复核/维持/诉讼 不进 CONTENTCHANGE_RE；只有「撤销/变更/更正/补充」类后续决定才 refresh_core。
    CONTENTCHANGE_RE = re.compile(
        r"(变更|撤销|更正|补充|修改|重新作出).{0,12}决定"
        r"|原.{0,8}决定书?.{0,14}(变更|撤销|更正|补充|修改)"
        r"|本决定.{0,8}(替代|取代|撤销|变更).{0,8}原决定")
    SUPERSEDE_FIELDS = ("title", "excerpt", "text", "grounds", "measures", "defense", "category")

    def _cell(r, k):
        v = ws.cell(r, C[k] + 1).value
        return str(v).strip() if v is not None else ""

    rows_by_sm = {}
    rows_by_org = {}              # 同机构索引（subject 分号前），用于「同机构不同人员并入原行」
    decided0 = {}                 # 每行的「原始」落款日（合并会改写单元格，near 判定必须用原始值）
    for r in range(2, ws.max_row + 1):
        s, m = _cell(r, "source"), _cell(r, "subject")
        if s and m:
            rows_by_sm.setdefault((s, m), []).append(r)
            _org = m.split("；")[0].strip()
            if _org:
                rows_by_org.setdefault((s, _org), []).append(r)
            decided0[r] = _cell(r, "decided")

    def _looks_followup(rec):
        """识别「针对已入库同一事项的后续决定/更新」。不看裸「撤销」（否则「撤销管理人登记」误判）。"""
        if str(rec.get("supersedes") or "").strip():
            return True
        blob = " ".join(str(rec.get(k) or "") for k in SUPERSEDE_FIELDS)
        return bool(FOLLOWUP_RE.search(blob))

    def _is_content_change(rec):
        """后续决定是否实质替换内容：是则刷新三核心列；复议/维持/诉讼否（只追加、保留原决定）。"""
        if str(rec.get("supersedes") or "").strip():
            return True
        blob = " ".join(str(rec.get(k) or "") for k in SUPERSEDE_FIELDS)
        return bool(CONTENTCHANGE_RE.search(blob))

    def _unique_name(d, name):
        base, ext = os.path.splitext(name)
        cand, i = name, 2
        while os.path.exists(os.path.join(d, cand)):
            cand = "%s（%d）%s" % (base, i, ext)
            i += 1
        return cand

    def _merge_into_row(r, rec, nm, pt, ur, refresh_core=True):
        """并入已有行：追加文件/路径/链接（去重）。

        2026-10-30 用户最终裁定（复核/维持类，refresh_core=False）：**原处分的核心内容
        （措施/事由/证据/依据）必须保留**，复核（复议）决定只刷新「处罚时间=最终决定日」并在
        「申辩理由」列追加一条**明确区分的复核结论**（如 【复核决定·2026-08-25】本会维持原决定），
        绝不用"维持《决定书》…"覆盖原处分实质处罚（撤销登记/公开谴责/罚款）。
        只有「撤销/变更/更正/补充」类后续决定（refresh_core=True）才真正替换三核心列。
        """
        def _multi(col, add):
            old = [x.strip() for x in str(ws.cell(r, C[col] + 1).value or "").split("；") if x.strip()]
            for x in add:
                if x and x not in old:
                    old.append(x)
            ws.cell(r, C[col] + 1).value = "；".join(old)

        _multi("filenames", nm)
        _multi("paths", pt)
        _multi("urls", ur)
        old_dec = _cell(r, "decided")
        if refresh_core:
            for k in ("measures", "grounds", "evidence", "legal_basis"):
                if k == "measures":
                    v = LS.format_measures_list(rec.get(k) or [])
                elif k == "grounds":
                    v = LS.normalize_grounds_items("；".join(x for x in (rec.get(k) or []) if x))
                else:
                    v = "；".join(x for x in (rec.get(k) or []) if x)
                if v:
                    ws.cell(r, C[k] + 1).value = v
        if rec.get("date_decided"):
            ws.cell(r, C["decided"] + 1).value = rec["date_decided"]
        if rec.get("publish_date"):
            ws.cell(r, C["publish"] + 1).value = rec["publish_date"]
        if rec.get("category"):
            ws.cell(r, C["category"] + 1).value = rec["category"]
        # 申辩列：追加区分标注（原处分内容始终保留，复核/维持只在此处注明最终结果）
        if refresh_core:
            kind = "内容变更（撤销/变更/更正/补充，已替换三核心列）"
            note = "原落款%s → 本决定（%s）更新；新落款%s" % (old_dec, kind, rec.get("date_decided"))
        else:
            dd = rec.get("date_decided") or old_dec
            kind = "复核/复议维持（原处分内容保留）"
            note = "【复核决定·%s】本会经复核维持原决定，原处分措施继续有效；复核决定书已留存于案例文件。" % dd
        dcol = C["defense"] + 1
        prev = str(ws.cell(r, dcol).value or "").strip()
        ws.cell(r, dcol).value = (prev + "\n\n" + note) if prev and prev != "无" else note
        print("  ~ 合并更新 [行%d] %s：追加文件 %d 份，落款 %s → %s（%s）"
              % (r, rec.get("subject"), len(nm), old_dec, rec.get("date_decided"), kind))

    def _merge_person_into_row(r, rec, new_person, nm, pt, ur):
        """把「同机构另一人员（或机构补罚）」的决定书并入该机构已有行（2026-10-30 用户指出的边界）。

        与 _merge_into_row 的区别：这里**不覆盖**机构原决定，而是做并集——
          · 主体：机构；张三 → 机构；张三、李四（追加人员）
          · 措施/事由/证据/依据/文件/路径/链接：并集追加
          · 落款：保留机构原日期（不覆盖），被并入文书的落款记入「申辩理由」列备查
        保证「一行=一个案例（机构+高管合并，措施分述）」在增量更新时同样成立。
        """
        cur = _cell(r, "subject")
        if "；" in cur:
            org, persons = cur.split("；", 1)
        else:
            org, persons = cur, ""
        plist = [x.strip() for x in persons.replace("、", "；").split("；") if x.strip()]
        if new_person and new_person not in plist:
            plist.append(new_person)
        ws.cell(r, C["subject"] + 1).value = (org + ("；" + "、".join(plist) if plist else ""))

        def _union(col, add):
            old = [x.strip() for x in str(ws.cell(r, C[col] + 1).value or "").split("；") if x.strip()]
            for x in add:
                x = (x or "").strip()
                if x and x not in old:
                    old.append(x)
            ws.cell(r, C[col] + 1).value = "；".join(old)

        # measures：拆现有（兼容「；」与换行）+ 追加 + 按主体重排分行（唯一真源）
        cur_m = str(ws.cell(r, C["measures"] + 1).value or "")
        mitems = [x.strip() for x in re.split(r"[；\n]", cur_m) if x.strip()]
        for x in (rec.get("measures") or []):
            x = (x or "").strip()
            if x and x not in mitems:
                mitems.append(x)
        ws.cell(r, C["measures"] + 1).value = LS.format_measures_list(mitems)
        _union("grounds", rec.get("grounds") or [])
        # 并入后重新列表化（幂等）：保持「一点一行、序号统一一、」
        ws.cell(r, C["grounds"] + 1).value = LS.normalize_grounds_items(
            ws.cell(r, C["grounds"] + 1).value)
        _union("evidence", rec.get("evidence") or [])
        _union("legal_basis", rec.get("legal_basis") or [])
        _union("filenames", nm)
        _union("paths", pt)
        _union("urls", ur)
        if rec.get("publish_date") and not _cell(r, "publish"):
            ws.cell(r, C["publish"] + 1).value = rec["publish_date"]
        old_dec = _cell(r, "decided")
        note = "同机构并入：%s（本决定落款%s；机构原落款%s）" % (
            new_person or "机构补罚", rec.get("date_decided") or "未载", old_dec or "未载")
        dcol = C["defense"] + 1
        prev = str(ws.cell(r, dcol).value or "").strip()
        ws.cell(r, dcol).value = (prev + "；" + note) if prev and prev != "无" else note
        print("  ~ 并入同机构 [行%d] %s += %s（措施/文件/链接并集，机构落款保留 %s）"
              % (r, org, new_person or "机构补罚", old_dec))

    batch, skipped, n_file, n_dl, missing = [], 0, 0, 0, []
    merged, merge_candidates = 0, []
    for rec in recs:
        keys = ledger._keys_of({
            "source": rec.get("source", ""),
            "subject": rec.get("subject", ""),
            "date_decided": rec.get("date_decided", ""),
            "filenames": [f["name"] for f in rec.get("files", [])],
            "urls": rec.get("urls", []),
        })
        dst_dir = os.path.join(root, rec["source_dir"], rec["category"])
        os.makedirs(dst_dir, exist_ok=True)
        ulist = rec.get("urls") or []

        def _place(rename_on_conflict=False):
            """本条的文书落盘；rename_on_conflict=True 时同名但不同文书自动加（N）后缀。"""
            nm, pt, ur = [], [], []
            f_inc = d_inc = 0
            for i, f in enumerate(rec.get("files") or []):
                src, name = f.get("src") or "", f.get("name") or ""
                if rename_on_conflict and name and os.path.exists(os.path.join(dst_dir, name)):
                    name = _unique_name(dst_dir, name)
                dst = os.path.join(dst_dir, name)
                try:
                    if src.startswith("http"):
                        if not os.path.exists(dst):
                            _download(src, dst)
                            d_inc += 1
                    elif src and os.path.isfile(src):
                        # 血的教训（0.4.52）：src 必须拼 basename，否则会把整个目录搬走
                        if not os.path.exists(dst):
                            shutil.copy2(src, dst)
                        f_inc += 1
                    else:
                        missing.append(name)
                except Exception as e:
                    missing.append("%s（%s）" % (name, e))
                nm.append(name)
                pt.append(os.path.dirname(os.path.relpath(dst, root)).replace("\\", "/"))
                ur.append(ulist[i] if i < len(ulist) else (ulist[-1] if ulist else ""))
            return nm, pt, ur, f_inc, d_inc

        # ============ 台账更新判定（2026-10-30 用户最终裁定）============
        # 裁定：①「同一机构·同一处罚时间」= 同一案件（同案多份/多人同日）→ 并入同一行；
        #       ②「同一机构·不同处罚时间」= 机构二次处罚（无论人员/机构、无论先后顺序）→ 一律【分行】；
        #       ③ 跨来源（证监局 vs 交易所等不同监管机关）= 不同案件 → 分行。
        # 顺序：①同主体后续决定(内容刷新) → ②精确重复跳过 → ③同机构同落日并入 → ④新增一行。
        src_s = str(rec.get("source") or "").strip()
        sub_s = str(rec.get("subject") or "").strip()
        rec_decided = str(rec.get("date_decided") or "").strip()
        rec_org = sub_s.split("；")[0].strip()
        rec_person = sub_s.split("；", 1)[1].strip() if "；" in sub_s else ""
        same_sm = [r for r in rows_by_sm.get((src_s, sub_s), [])
                   if _cell(r, "decided") != rec_decided]

        # ① 同主体 + 后续决定信号（变更/撤销/复议/维持…）→ 内容刷新式合并（三核心列可覆盖）
        if same_sm and _looks_followup(rec):
            if len(same_sm) == 1:
                nm, pt, ur, fi, di = _place(rename_on_conflict=True)
                n_file += fi
                n_dl += di
                _merge_into_row(same_sm[0], rec, nm, pt, ur, refresh_core=_is_content_change(rec))
                merged += 1
                continue
            merge_candidates.append({
                "subject": rec.get("subject"), "date_decided": rec.get("date_decided"),
                "source": src_s, "rows": same_sm,
                "reason": "同(来源,主体)有多条候选，无法确定并入哪一条，待人工确认"})
            continue

        # ② 精确重复（同主体+同落款+同文件/链接）→ 跳过（不重复入库）
        if keys & have:
            skipped += 1
            continue

        # ③ 同一机构 + 同一处罚时间 = 同一案件（同案多份决定/多人同日）→ 并入同一行。
        #    不同处罚时间（机构二次处罚）不并入 → 走 ④ 另起一行（2026-10-30 用户最终裁定「二次处罚分行」）。
        if rec_org and rec_decided:
            same_case = [r for r in rows_by_org.get((src_s, rec_org), [])
                         if _cell(r, "decided") == rec_decided]
            if same_case:
                target = same_case[0]
                nm, pt, ur, fi, di = _place(rename_on_conflict=True)
                n_file += fi
                n_dl += di
                _merge_person_into_row(target, rec, rec_person, nm, pt, ur)
                merged += 1
                continue

        names, paths, urls, _f, _d = _place()
        n_file += _f
        n_dl += _d
        batch.append({
            "date_decided": rec.get("date_decided", ""),
            "publish_date": rec.get("publish_date") or "待补",
            "subject": rec.get("subject", ""),
            "measures": LS.format_measures_list(rec.get("measures") or []),
            "grounds": LS.normalize_grounds_items("；".join(rec.get("grounds") or [])),
            "evidence": "；".join(rec.get("evidence") or []),
            "legal_basis": "；".join(rec.get("legal_basis") or []),
            "defense": rec.get("defense") or "无",
            "category": rec.get("category", ""),
            "source": rec.get("source", ""),
            "filenames": names, "paths": paths, "urls": urls,
            "typical": "是" if rec.get("typical") else "否",
        })

    if batch:
        added = ledger._append_cases(root_p, wb, ws, batch)
    else:
        added = 0
        if merged:
            # 2026-10-28 修：本轮若「只合并、无新增」，batch 为空 → _append_cases 不执行，
            # 而它的 wb.save 也只在 added>0 时调用 → 合并改动会全部留在内存里丢失
            # （实测：PDF 已复制、控制台也打印了"合并更新"，但台账 decided/申辩列纹丝不动）。
            wb.save(os.path.join(root, LS.LEDGER_NAME))
    print("  新增 %d 行，合并更新 %d 行，判重跳过 %d 条，复制 %d 份，下载 %d 份，缺件 %d"
          % (added, merged, skipped, n_file, n_dl, len(missing)))
    if merge_candidates:
        _dump(os.path.join(work, "merge_candidates.json"), merge_candidates)
        print("  ⚠ 疑似「重新下发」但同机构有多条候选、未自动合并：%d 条 → %s"
              % (len(merge_candidates), os.path.join(work, "merge_candidates.json")))
    if missing:
        _dump(os.path.join(work, "missing_files.json"), missing)
        print("  缺件明细 → %s" % os.path.join(work, "missing_files.json"))
    ledger.cmd_sort(root_p)
    return 0


# ----------------------------------------------------------------- laws
def _law_dir(root):
    """定位 06 配套法规目录（名字以 06_ 开头，找不到就按约定名建）。"""
    for nm in sorted(os.listdir(root)):
        p = os.path.join(root, nm)
        if os.path.isdir(p) and nm.startswith("06"):
            return p
    return os.path.join(root, "06_以上各类所有案例的配套法规依据")


def _law_count(law_dir):
    n = 0
    for _, _, fns in os.walk(law_dir):
        n += sum(1 for f in fns if os.path.splitext(f)[1].lower() in (".pdf", ".doc", ".docx"))
    return n


def stage_laws(root, work, args):
    """06 配套法规自动抓取（2026-10-28 用户要求 #3）。

    旧流程要求使用者在首次建库时**手工单独跑一次** fetch_laws.py，
    漏跑就会得到「06 目录空、D3 一堆缺失」的半成品库。
    现在并入流水线：入库后、回填法规清单前自动补。

    fetch_laws.py 本身是幂等的（目标文件已存在且 >1KB 即跳过），
    所以日常增量跑它几乎零网络开销；只有缺件时才真正下载。
    """
    if args.no_laws:
        print("  （--no-laws，跳过法规抓取）")
        return 0
    law_dir = _law_dir(root)
    before = _law_count(law_dir)
    os.makedirs(law_dir, exist_ok=True)
    print("  06 目录：%s（现有 %d 部）" % (law_dir, before))
    if before == 0:
        print("  → 首次建库：自动下载配套法规原文（40 部，含官方 PDF/Word 与网页转 Word）")
    try:
        _sh([os.path.join(HERE, "fetch_laws.py"), law_dir,
             "--tmp", os.path.join(work, "_lawsrc")])
    except SystemExit:
        print("  !! fetch_laws.py 以非零码退出（部分法规下载失败），已继续本轮；"
              "失败清单见上，下轮会自动重试（幂等）")
    after = _law_count(law_dir)
    print("  06 法规：%d → %d 部" % (before, after))
    return 0


# ----------------------------------------------------------------- finalize
def stage_finalize(root, work, args):
    """06 法规清单回填 + 05 交叉归档。法规原文由 fetch_laws.py 单独维护，
    本步只做「已留存法规 → 清单」的统计回填，不联网（联网补抓见 fetch_laws.py）。"""
    _sh([os.path.join(HERE, "finalize_library.py"), "--root", root,
         "--laws", "--typical", "--laws-out", os.path.join(work, "laws.json")])
    return 0


# ----------------------------------------------------------------- verify
def stage_verify(root, work, win_lib, args):
    """校验必须带上本轮窗口，否则 A2「窗口外记录」会按脚本默认窗口误报。"""
    out = os.path.join(work, "verify_report.json")
    # 校验「发现问题」是**正常输出**而非流程失败（退出码 1 表示有 problems），
    # 不能因此中断本轮（2026-10-27：曾导致 finalize 之后的收尾动作全被跳过）。
    try:
        _sh([os.path.join(HERE, "verify_library.py"), "--root", root, "--json", out,
             "--start", win_lib[0], "--end", win_lib[1], "--window-by", "decided"]
            + (["--ocr-cache", args.ocr_cache] if args.ocr_cache else []))
    except SystemExit:
        print("  （校验脚本以非零码退出＝报告存在问题，已继续本轮）")
    rep = _load(out, {})
    # problems 在报告里是**列表**（每条一句话），历史版本曾是 dict；
    # 两种都要能打印——否则收尾阶段直接 AttributeError 崩掉（2026-10-28 实测）。
    probs = rep.get("problems") or []
    if isinstance(probs, dict):
        items = ["%s: %s" % (k, v) for k, v in probs.items()]
    else:
        items = [str(x) for x in probs]
    if items:
        print("\n== 校验 problems（%d 项）==" % len(items))
        for x in items:
            print("  · " + x)
    else:
        print("\n== 校验未发现 problems ==")
    return 0


# ----------------------------------------------------------------- retire
def stage_retire(root, work, win_lib, args):
    """把落款日超出窗口的行 + PDF（含 05 副本）移出库外备份区，台账删行。

    硬约束（0.4.52）：移动 src 必须拼 basename 且先 isfile 断言，禁止 shutil.move(目录)。
    """
    import ledger_style as LS
    from openpyxl import load_workbook
    from pathlib import Path

    ls, le = win_lib
    root_p = Path(root)
    wb = load_workbook(os.path.join(root, LS.LEDGER_NAME))
    ws = wb[LS.SHEET_CASE]
    C = LS.COL
    backup = Path(args.backup_dir) if args.backup_dir else Path(LS.case_archive_dir(root))
    backup.mkdir(parents=True, exist_ok=True)

    moved, rows_out, orphan = 0, [], 0
    keep_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(x not in (None, "") for x in row[1:]):
            continue
        dd = str(row[C["decided"]] or "").strip()[:10]
        pv = str(row[C["publish"]] or "").strip()[:10]
        # 2026-10-29 修：窗口判据取 max(落款日, 挂网日)，不再单看落款。
        # 根因：库保留窗口（按落款 365 天）比抓取窗口（按挂网 365 天）更窄，
        # 证监局文书又普遍滞后刊发 9~23 天 → 落款已滚出窗口、挂网仍在窗口内的案例
        # 本轮被 retire 移出，下一轮增量又按挂网日抓回来 → 无限「移出/抓回」震荡。
        # 反例（2026-10-29 实测）：河南智慧和信 落款 2025-10-28 / 挂网 2025-10-30，
        # 窗口起点 2025-10-29 —— 只按落款判就该出库，但它恰恰是本窗口新收的典型案例。
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", dd):
            keep_rows.append(row)
            continue
        pv_ok = pv if re.match(r"^\d{4}-\d{2}-\d{2}$", pv) else dd
        if max(dd, pv_ok) >= ls:
            keep_rows.append(row)
            continue
        # 2026-10-29 修：多文件行必须按「第 i 个文件 ↔ 第 i 个目录」定位。
        # 旧版对所有文件都拼 pts[0]，多来源/多分类的行（如中基协 3 份附件跨目录）
        # 除第一份外全部 isfile 失败 → 计 orphan 不移动 → 台账删了行、PDF 变孤儿。
        fns = [x.strip() for x in str(row[C["filenames"]] or "").split("；") if x.strip()]
        pts = [x.strip() for x in str(row[C["paths"]] or "").split("；") if x.strip()]
        for i, fn in enumerate(fns):
            rel = pts[i] if i < len(pts) else (pts[0] if pts else "")
            src_pdf = os.path.join(root, rel, os.path.basename(fn)) if rel else ""
            if src_pdf and os.path.isfile(src_pdf):
                shutil.move(src_pdf, os.path.join(str(backup), os.path.basename(fn)))
                moved += 1
            else:
                orphan += 1
        rows_out.append(row)

    if not rows_out:
        print("  无超窗行（窗口起点 %s）" % ls)
        return 0

    # 05 副本联动移出（典型案副本随主件一起出库，避免孤儿副本）
    # 2026-10-29 修：① 由 os.remove 改为移入备份区（原实现等于直接删文件，不可逆）；
    # ② 匹配改为 basename 精确相等，旧版 `fn in 案例文件串` 是子串匹配，
    #    「杭州合信投资管理有限公司.pdf」会命中「王晓勇-杭州合信投资管理有限公司.pdf」。
    out_names = set()
    for r in rows_out:
        for fn in str(r[C["filenames"]] or "").split("；"):
            if fn.strip():
                out_names.add(os.path.basename(fn.strip()))
    for dp, _, fns in os.walk(os.path.join(root, "05_重点典型案例")):
        for fn in fns:
            if fn in out_names:
                try:
                    shutil.move(os.path.join(dp, fn), os.path.join(str(backup), fn))
                    moved += 1
                except OSError:
                    pass

    # 台账：清值后重排（原地写值不会缩 max_row，必须整表重写，见 0.4.x 幽灵行教训）
    for r_i in range(ws.max_row, 1, -1):
        ws.delete_rows(r_i)
    wrap = set(LS.CASE_WRAP_COLS)
    for r in keep_rows:
        ws.append(list(r))
        rr = ws.max_row
        for i in range(1, len(r) + 1):   # 重写后重新套用冻结样式，否则整表失去边框/换行
            LS.data_cell(ws.cell(row=rr, column=i), wrap=(i in wrap))
    LS.renumber(ws)
    wb.save(os.path.join(root, LS.LEDGER_NAME))
    print("  移出超窗 %d 行 / %d 份文件（备份区 %s，孤儿 %d）" % (len(rows_out), moved, backup, orphan))
    return 0


# ----------------------------------------------------------------- 不收录清单
def _sync_list_header(md):
    """回写清单文件头的「条数」与「最后更新」。

    2026-10-28 实测缺陷：`update_excluded_md` 只追加行、从不回写表头，
    文件头长期停在旧数字（本轮实际 4392 条，表头还写 4361），而 SKILL 0.3 明确要求
    「追加一行 + 同步把条数加一」。表头是给人看的唯一计数口径，漂移了就等于骗人。
    做法：按实际表格行数（`| 数字 |` 开头）重写，不依赖增量累加（累加会因去重而错）。
    """
    try:
        txt = open(md, encoding="utf-8").read()
    except Exception:
        return
    n = len(re.findall(r"(?m)^\| \d+ \|", txt))
    txt2 = re.sub(r"(?m)^(- 条数：\*\*)\d+(\*\*)", r"\g<1>%d\g<2>" % n, txt)
    txt2 = re.sub(r"(?m)^(- 最后更新：).*$", r"\g<1>%s" % date.today().isoformat(), txt2)
    if txt2 != txt:
        open(md, "w", encoding="utf-8").write(txt2)


def update_excluded_md(root, work):
    """把本轮剔除条目登记进「不收录清单.md」（记录类文件一律 .md，用户 2026-10-13 要求）。"""
    excl = _load(os.path.join(work, "excluded.json"), [])
    if not excl:
        return 0
    md = os.path.join(root, "不收录清单.md")
    seen = set()
    if os.path.exists(md):
        for line in open(md, encoding="utf-8"):
            seen.add(line.strip())
    lines = []
    if not os.path.exists(md):
        lines += ["# 不收录清单", "",
                  "| 序号 | 来源 | 标题/主体 | 原文链接 | 不收录原因 |",
                  "| --- | --- | --- | --- | --- |", ""]
    n = len([l for l in seen if l.startswith("| ") and re.match(r"^\| \d+ \|", l)])
    added = 0
    for e in excl:
        row = "| %d | %s | %s | %s | %s |" % (
            n + added + 1, e.get("source", ""),
            str(e.get("title") or e.get("subject", "")).replace("|", "\\|")[:60],
            e.get("url", ""), str(e.get("reason", "")).replace("|", "\\|"))
        if row in seen:
            continue
        lines.append(row)
        seen.add(row)
        added += 1
    if added:
        with open(md, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        _sync_list_header(md)          # 表头条数必须与实际行数一致
    print("  不收录清单新增 %d 条 → %s" % (added, md))
    return added


# ----------------------------------------------------------------- cleanup
def stage_cleanup(root, work, args):
    """终态整洁收尾（2026-10-28 用户要求 #1）。

    库内终态只允许留 6 个分类目录 + 4 个记录文件；其余（台账 .bak_*、
    清单备份、遗留的 _backup_out_of_window / _quarantine_orphan / _*.json、
    分类目录里误放的 md/log/xlsx）一律**移动**到库外留痕区，只移不删。
    """
    arc = os.path.join(os.path.dirname(root), "_案例库维护留痕")
    try:
        _sh([os.path.join(HERE, "cleanup_library.py"), "--root", root,
             "--archive", os.path.join(arc, os.path.basename(root), date.today().isoformat()),
             "--apply"])
    except SystemExit:
        print("  !! cleanup 阶段异常（不影响本轮结果），可手工复核："
              "python cleanup_library.py --root ...")
    return 0


# ----------------------------------------------------------------- updscan
def stage_updscan(root, work, args):
    """已入库案例「平台更新」回查（2026-10-28 用户关注点）。

    默认轻量：对库内全部案例的原文链接做可达性 + 内容哈希比对（原地更新/死链检测）。
    --deep-updscan：额外重扫平台并合并后续决定（变更/复议/补充等）。
    网络操作全程节流+容错，绝不损坏台账；歧义写入 update_candidates.json。
    """
    cmd = [os.path.join(HERE, "update_scan.py"), "--root", root,
           "--work", os.path.join(work, "updscan")]
    if getattr(args, "deep_updscan", False):
        cmd += ["--deep", "--start", getattr(args, "updscan_start", "2020-01-01")]
    s = getattr(args, "updscan_sample", 0) or 0
    if s:
        cmd += ["--sample", str(s)]
    try:
        _sh(cmd)
    except SystemExit:
        print("  [updscan] 以非零码退出，已继续本轮（详见报告）")
    return 0


# ----------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(description="案例库滑动窗口增量更新")
    ap.add_argument("--root", required=True, help="案例库根目录（必填，脚本不内置任何路径）")
    ap.add_argument("--days", type=int, default=365, help="库窗口天数（按落款日），默认 365")
    ap.add_argument("--end", default="", help="窗口终点 YYYY-MM-DD，默认今天")
    ap.add_argument("--start", default="", help="窗口起点，给了就忽略 --days")
    ap.add_argument("--full", action="store_true",
                    help="全量模式：库窗口起点取到最早（各平台最早可获取 → 当日），用于首次建库")
    ap.add_argument("--overlap-days", type=int, default=7,
                    help="增量模式在「上次最新挂网日」基础上回退的天数（抗挂网滞后），默认 7")
    ap.add_argument("--lag-days", type=int, default=0,
                    help="（兼容旧参数，等价于 --overlap-days）")
    ap.add_argument("--work-dir", default="",
                    help="中间产物目录。**默认放在库外同级**（库根有 9 项白名单，"
                         "中间产物落进库内会算脏，收尾还要清理）")
    ap.add_argument("--ocr-cache", default="", help="OCR 文本缓存目录（强烈建议复用）")
    ap.add_argument("--backup-dir", default="", help="超窗文件备份区，默认 <root> 同级")
    ap.add_argument("--stages", default=",".join(ALL_STAGES))
    ap.add_argument("--bureaus", default="", help="只抓指定证监局（拼音逗号分隔）")
    ap.add_argument("--max-pages", type=int, default=0, help="中基协每栏最多翻几页（0=全部）")
    ap.add_argument("--no-csrc", action="store_true")
    ap.add_argument("--no-amac", action="store_true")
    ap.add_argument("--no-exchange", action="store_true")
    ap.add_argument("--no-laws", action="store_true",
                    help="跳过 06 配套法规自动抓取（默认 laws 阶段会自动补缺件）")
    ap.add_argument("--no-cleanup", action="store_true",
                    help="跳过终态整洁收尾（默认 cleanup 阶段会把杂物移出库外留痕区）")
    ap.add_argument("--updscan", action="store_true",
                    help="（已默认开启）对已入库案例做平台更新回查（轻量：原文链接可达性+内容哈希比对）；可显式指定")
    ap.add_argument("--no-updscan", action="store_true",
                    help="关闭「随增量执行自动开启」的轻量回查（默认增量即跑，覆盖 0.4.99 用户要求；--full 首建自动跳过）")
    ap.add_argument("--deep-updscan", action="store_true",
                    help="对已入库案例做深度回查（重扫平台+合并后续决定）；与 --updscan 互斥时取深度")
    ap.add_argument("--updscan-sample", type=int, default=0,
                    help="updscan 仅抽样前 N 条案例（调试用）")
    ap.add_argument("--updscan-start", default="2020-01-01",
                    help="--deep-updscan 重扫起点，默认 2020-01-01")
    args = ap.parse_args(argv)

    root = os.path.abspath(args.root)
    # 工作目录默认放在运维主目录（库父目录/_案例库维护留痕/<库名>/_work）：库根只放交付物，绝不落在库内
    work = os.path.abspath(args.work_dir or LS.work_dir(root))
    os.makedirs(root, exist_ok=True)
    os.makedirs(work, exist_ok=True)

    end = args.end or date.today().isoformat()
    if args.full:
        start = args.start or EARLIEST
    else:
        # 窗口口径统一走 winlib（唯一事实源）。旧版此处 `end-(days-1)` 与
        # verify_library 的 `today.year-1` 差 1 天，导致 verify 判「在窗内」
        # 而 retire 判「超窗」互删，见 winlib.py 头部事故记录。
        import winlib
        start = args.start or winlib.window(end, args.days)[0]
    win_lib = (start, end)

    # ---- 抓取窗口（三种情形，顺序不能乱）----
    # 血的教训（2026-10-27 实测）：**库还没覆盖到窗口起点时绝不能用增量锚点**。
    # 现役库只有上次试建的 42 行、最新挂网日很近 → 锚点把 365 天窗口压成 11 天，
    # 36 个局每局只抓到个位数，日志却一切正常。故必须先判「库是否已覆盖到起点」。
    overlap = args.overlap_days or args.lag_days
    last_pub = latest_publish(root)
    first_dec = earliest_decided(root)
    if args.full:
        # 全量库不裁窗口，任何早于锚点的文书都已在库 → 锚点可用（否则每天重爬全历史）
        crawl_start = ((d(last_pub) - timedelta(days=overlap)).isoformat()
                       if last_pub else start)
        mode = ("全量·增量（锚点 %s，回退 %d 天）" % (last_pub, overlap) if last_pub
                else "全量（最早可获取 → 当日）")
    elif not last_pub or not first_dec:
        crawl_start = start
        mode = "首次建库（库窗口 %s → 当日全扫）" % start
    elif first_dec > start:
        # 库最早落款晚于窗口起点 = 历史部分还没补齐，必须按库窗口全扫
        crawl_start = start
        mode = "补齐历史（库最早落款 %s 晚于窗口起点 %s，按库窗口全扫）" % (first_dec, start)
    else:
        # 库已覆盖到窗口起点 → 只需抓锚点之后新增的（抗挂网滞后回退 overlap 天）
        crawl_start = max(start, (d(last_pub) - timedelta(days=overlap)).isoformat())
        mode = "增量（锚点 %s，回退 %d 天）" % (last_pub, overlap)
    win_crawl = (crawl_start, end)

    stages = [s.strip() for s in args.stages.split(",") if s.strip()]
    unknown = [s for s in stages if s not in VALID_STAGES]
    if unknown:
        raise SystemExit("未知阶段：%s（可选 %s）" % (unknown, "/".join(VALID_STAGES)))
    # updscan（已入库案例平台更新回查）：默认随「增量执行」自动开启，覆盖 0.4.99 用户要求
    # 「检查已创建库的所有相关案例」。轻量 url-recheck 回查全部已入库链接（可达性+内容哈希，
    # 捕获原地替换/更正+死链）；--deep-updscan 额外重扫平台（周期性/每月一次，避开每日重爬频控）。
    # 首次全量建库(--full)刚抓完，回查无意义且重复下载，故自动跳过；--no-updscan 可关闭轻量回查。
    if "updscan" not in stages:
        if args.deep_updscan or (not args.no_updscan and not args.full):
            stages.append("updscan")

    print("=" * 66)
    print("案例库滑动窗口增量更新")
    print("  库根    : %s" % root)
    print("  库窗口  : %s ~ %s（%d 天，按落款日）" % (start, end, (d(end) - d(start)).days + 1))
    print("  抓窗口  : %s ~ %s  【%s】" % (win_crawl[0], end, mode))
    print("  阶段    : %s" % " → ".join(stages))
    print("=" * 66)

    for s in stages:
        print("\n########## 阶段 %s ##########" % s)
        if s == "crawl":
            stage_crawl(root, work, win_lib, win_crawl, args)
        elif s == "build":
            stage_build(root, work, win_lib, args)
        elif s == "ingest":
            stage_ingest(root, work, args)
        elif s == "laws":
            stage_laws(root, work, args)
        elif s == "finalize":
            stage_finalize(root, work, args)
        elif s == "verify":
            stage_verify(root, work, win_lib, args)
        elif s == "retire":
            stage_retire(root, work, win_lib, args)
        elif s == "cleanup":
            if args.no_cleanup:
                print("  （--no-cleanup，跳过终态整洁收尾）")
            else:
                stage_cleanup(root, work, args)
        elif s == "updscan":
            stage_updscan(root, work, args)

    if "build" in stages:
        update_excluded_md(root, work)

    print("\n完成：%s" % " → ".join(stages))
    print("  台账行数与 PDF 数请务必核对（0.4.x 幽灵行教训）：")
    print("    python scripts/verify_library.py --root \"%s\"" % root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
