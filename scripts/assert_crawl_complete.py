# -*- coding: utf-8 -*-
"""抓取完整性断言（固定规则 · 零额外网络 · 可解释退出码）。

为什么要有它（2026-10-28 用户要求）：
  「一定要确保收集文件完整，不论是在代码层面确保还是在校验层面确保，还是二者皆有，
    在满足这一点的同时也要确保开销和固定化的原则。」
  ⇒ A6 平台对账（--online）要逐局联网重拉，开销大；日常不能靠它兜底。
  ⇒ 正确做法：抓取时**顺手记下自证事实**（平台声明总数、实际翻了第几页、为什么停），
    再由本脚本用**固定规则**判定是否完整。零额外网络，结果可复现，不依赖模型判断。

输入（run_update.py stage_crawl 的产物，缺哪个就跳过哪个来源）：
  <work>/csrc.json    证监局：stats.channels[] = {bureau_cn, col, platform_total,
                        pages_fetched, items_window, terminated_by}
  <work>/amac.json    中基协：columns{key} = {label, total_pages, pages_fetched, full, items}
  <work>/exch.json    交易所：stats = {sse_total, szse_total, bse_total,
                        latest{sse,szse,bse}, latest_seen, window, warnings[]}

固定规则（FAIL 阻断 / WARN 仅提示）：
  R1 已跑来源条目为 0                        → 能证明「平台最新一条 < 窗口起点」则 WARN，
                                                 否则 FAIL（0 条≠无新增）
  R2 证监局终止原因 ∉ {start_boundary, short_page}  → FAIL（http_fail/json_error/empty_results
                                                 = 翻页被意外中断，必然漏抓）
  R3 中基协 full != true（未翻全）            → FAIL；若显式 --max-pages 截断 → WARN
  R4 交易所**逐所**判定（不看总数）            → 某所 0 条且拿不到该所最新日期 → FAIL；
                                                 最新日期 < 窗口起点 → WARN（真无新增）
  R5 证监局正文失败率 > 20%                    → WARN（err/window_items）
  R6 与上轮 manifest 比较：同来源条目骤降 > 50% → 窗口相同则 FAIL，窗口不同仅 WARN
  R7 产物缺自证指标（旧版脚本产物）             → WARN（提示重跑抓取以启用断言）
  R8 stats.failed_items 非空                    → FAIL（列表抓到了、正文没取到 = 条目被丢弃 =
                                                 漏案；必须补抓。2026-10-28 广东局 9 条事故）
  R9 work/crawl_failed.json 里的来源             → FAIL（抓取脚本异常退出时，目录里往往还躺着
                                                 上一轮产物；照读就是"用旧结论判新抓取"）

用法：
  python assert_crawl_complete.py --work <work目录> [--prev <上轮manifest.json>]
                                  [--out <本轮manifest.json>] [--max-pages N]
  # 退出码 0 = PASS（可带 WARN）；1 = 存在 FAIL（本次抓取不完整，禁止继续入库）
"""
import argparse
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

OK_TERMINATORS = {"start_boundary", "short_page"}


def _load(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _pct(a, b):
    return 0.0 if not b else 100.0 * a / b


def check_csrc(obj, fails, warns, notes):
    if not obj:
        return None
    chans = ((obj.get("stats") or {}).get("channels") or [])
    items = obj.get("items") or []
    if not chans:
        warns.append("证监局：产物无 channels 自证指标（旧版抓取脚本），请重跑以启用断言")
        notes["csrc"] = {"items": len(items), "selfproof": False}
        return len(items)
    bad = [c for c in chans if c.get("terminated_by") not in OK_TERMINATORS]
    for c in bad[:8]:
        fails.append("证监局 %s·%s 翻页被中断：terminated_by=%s（窗口内 %d 条，平台总 %s）"
                     % (c.get("bureau_cn"), c.get("col"), c.get("terminated_by"),
                        c.get("items_window", 0), c.get("platform_total")))
    if len(bad) > 8:
        fails.append("证监局另有 %d 个栏目翻页被中断（共 %d 个）" % (len(bad) - 8, len(bad)))
    err = int((obj.get("stats") or {}).get("err") or 0)
    win = int((obj.get("stats") or {}).get("window_items") or 0)
    if win and _pct(err, win) > 20:
        warns.append("证监局正文失败率 %.1f%%（%d/%d），可能影响三核心列抽取"
                     % (_pct(err, win), err, win))
    # R8（2026-10-28 新增，真实事故驱动）：正文抓取失败 = 条目被丢弃 = 漏案，
    # 必须 FAIL 阻断。实测：广东局 9 条私募决定书因正文接口连续失败被静默 continue，
    # 库凭空少 9 个案例，而旧脚本只打一行 WARN、退出码 0，整轮"看起来正常"。
    failed = (obj.get("stats") or {}).get("failed_items") or []
    if failed:
        for x in failed[:10]:
            fails.append("证监局 %s·%s 决定书未取到正文（%s）：%s"
                         % (x.get("bureau_cn"), x.get("col"), x.get("reason"),
                            str(x.get("title"))[:40]))
        if len(failed) > 10:
            fails.append("证监局另有 %d 条正文抓取失败未入库（共 %d 条）"
                         % (len(failed) - 10, len(failed)))
    notes["csrc"] = {"items": len(items), "channels": len(chans),
                     "bad_channels": len(bad), "selfproof": True,
                     "failed_items": len(failed),
                     "err_rate": round(_pct(err, win), 1)}
    return len(items)


def check_amac(obj, fails, warns, notes, max_pages):
    if not obj:
        return None
    cols = obj.get("columns") or {}
    items = obj.get("items") or []
    if not cols:
        warns.append("中基协：产物无 columns 自证指标（旧版抓取脚本），请重跑以启用断言")
        notes["amac"] = {"items": len(items), "selfproof": False}
        return len(items)
    for key, c in cols.items():
        if c.get("full") is True:
            continue
        msg = ("中基协 %s 未翻全：抓到 %s/%s 页（条目 %s）"
               % (c.get("label") or key, c.get("pages_fetched", "?"),
                  c.get("total_pages", "?"), c.get("items", "?")))
        if max_pages:
            warns.append(msg + "；因 --max-pages=%d 显式限制" % max_pages)
        else:
            fails.append(msg)
    notes["amac"] = {"items": len(items), "columns": len(cols), "selfproof": True}
    return len(items)


VENUES = (("sse", "上交所"), ("szse", "深交所"), ("bse", "北交所"))


def check_exch(obj, fails, warns, notes, crawl_start=""):
    """交易所按【所】判，不能按总数判。

    2026-10-28 实测：深交所被频控整段返回 0 条时，上交所/北交所照常有数据，
    "总数 0" 与 "某所 0" 是两回事；而交易所自律措施本来就极稀疏，
    日常增量命中 0 条是常态。所以判据必须是**逐所**：
      该所有频控告警 且 拿不到该所"平台最新日期" → 无法证明 → FAIL
      该所最新日期早于抓取窗口起点              → 真无新增 → WARN
    """
    if not obj:
        return None
    st = obj.get("stats") or {}
    win = (obj.get("window") or st.get("window") or [""])[0] or crawl_start
    latest = st.get("latest") or {}
    wl = st.get("warnings") or []
    by_venue = {}
    for w in wl:
        by_venue.setdefault(w.get("venue"), []).append(w)
    totals, per = {}, {}
    for v, cn in VENUES:
        cnt = int(st.get("%s_total" % v) or 0)
        totals[v] = cnt
        lv = latest.get(v) or ""
        ws = by_venue.get(v) or []
        per[v] = {"items": cnt, "latest": lv, "warnings": len(ws)}
        if cnt:
            if ws:
                warns.append("%s 本轮 %d 条，但出现 %d 次频控告警，需复核是否漏抓"
                             % (cn, cnt, len(ws)))
            continue
        if lv and win and lv < win:
            warns.append("%s 本轮 0 条（平台最新 %s < 窗口起点 %s），已核实为真无新增%s"
                         % (cn, lv, win, "（其间有 %d 次频控探测）" % len(ws) if ws else ""))
        elif lv and ws:
            # 有数据但最新一条落在窗口内 → 说明窗口里本该有，只是没命中私募，无需阻断
            warns.append("%s 本轮 0 条，平台最新 %s 在窗口内但无私募命中，需复核" % (cn, lv))
        elif ws:
            fails.append("%s 出现 %d 次频控告警（%s），且拿不到该所最新日期——"
                         "零条=无法区分真无数据与限流"
                         % (cn, len(ws), "、".join(sorted({w.get("kind", "?") for w in ws}))))
        else:
            fails.append("%s 本轮 0 条且无「平台最新日期」自证（旧版产物或抓取未完成），"
                         "不能视为『本期无新增』" % cn)
    notes["exchange"] = {"items": sum(totals.values()), "warnings": len(wl),
                         "per_venue": per, "selfproof": bool(latest)}
    # 深交所逐栏目「声明 vs 实抓」对账（2026-10-28 新增，替代旧的"检索式栏目"播报）。
    # 旧版因 CATALOGID 少了 `_snapshot` 后缀，长期只能拿到真实数据的 1/5 却毫无征兆。
    # 现在每个栏目都把接口声明的 recordcount 与实际抓到的条数一起带出，
    # 声明 > 实抓 直接判 FAIL —— 这是唯一能证明"拿全了"的信号。
    for cid, v in (st.get("szse_cats") or {}).items():
        dec, got = int(v.get("declared") or 0), int(v.get("got") or 0)
        nm = v.get("name") or cid
        if dec and got < dec:
            fails.append("深交所「%s」声明 %d 条但只抓到 %d 条（分页未翻完/中途限流），"
                         "不能视为已拿全" % (nm, dec, got))
        elif dec:
            notes.setdefault("szse_cats", {})[cid] = {"name": nm, "declared": dec, "got": got}
    # 深交所附件正文下载：列表完整性（声明==实抓）已在上面保证；正文 PDF 若因静态域名
    # 对本 IP 临时封锁而没下完，属于"可恢复暂挂"，记 WARN 交给定时任务下轮续下，
    # 不以 FAIL 阻断入库（否则会连证监局/中基协的新增一起被卡住）。
    sb = st.get("szse_body") or {}
    sb_blocked = int(sb.get("blocked") or 0)
    sb_err = int(sb.get("dl_err") or 0)
    if sb_blocked:
        warns.append("深交所静态域名对本 IP 临时封锁：本轮 %d 个候选附件未下载（已暂停、"
                     "下轮自动续下），列表完整性不受影响" % sb_blocked)
    elif sb_err:
        warns.append("深交所 %d 个附件下载失败（非封锁类，下轮自动续下）" % sb_err)
    return sum(totals.values())


def window_of(csrc, amac, exch):
    """抓取窗口 [start, end]：优先取产物里记录的 window。"""
    for o in (csrc, exch, amac):
        if not o:
            continue
        w = o.get("window") or (o.get("stats") or {}).get("window")
        if w and isinstance(w, list) and len(w) == 2 and w[0]:
            return w
    return None


def _latest_seen(name, csrc, amac, exch=None):
    """平台侧「最新一条」的挂网日：用来证明窗口内 0 条到底是真无新增还是漏抓。"""
    try:
        if name == "证监局" and csrc:
            ds = [c.get("latest_seen") or "" for c in
                  ((csrc.get("stats") or {}).get("channels") or [])]
            return max([d for d in ds if d], default="")
        if name == "中基协" and amac:
            ds = [c.get("latest_published") or "" for c in (amac.get("columns") or {}).values()]
            return max([d for d in ds if d], default="")
        if name == "交易所" and exch:
            st = exch.get("stats") or {}
            ds = [st.get("latest_seen") or ""] + [v for v in (st.get("latest") or {}).values()]
            return max([d for d in ds if d], default="")
    except Exception:
        pass
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", required=True, help="本轮抓取产物目录")
    ap.add_argument("--prev", default="", help="上轮 manifest.json（环比骤降检查）")
    ap.add_argument("--out", default="", help="输出本轮 manifest.json")
    ap.add_argument("--max-pages", type=int, default=0, help="中基协显式页数上限")
    ap.add_argument("--skip", default="",
                    help="本轮未跑的来源，逗号分隔（如 '交易所'）：跳过检查。"
                         "不加这个参数时，--no-exchange 会把上一轮残留的 exch.json "
                         "当成『本轮产物』来判定（2026-10-28 实测误报 2 项 FAIL）。")
    a = ap.parse_args()

    fails, warns, notes = [], [], {}
    totals = {}
    skip = set(x.strip() for x in a.skip.split(",") if x.strip())

    csrc = _load(os.path.join(a.work, "csrc.json"))
    amac = _load(os.path.join(a.work, "amac.json"))
    exch = _load(os.path.join(a.work, "exch.json"))
    # 抓取脚本异常退出（run_update 会写 crawl_failed.json）：**不能**退化成"跳过"，
    # 因为目录里往往还躺着上一轮的产物——照读就会出现"用旧结论判新抓取"的假报告。
    failed_src = _load(os.path.join(a.work, "crawl_failed.json")) or []
    if failed_src:
        for name in failed_src:
            fails.append("%s 抓取脚本异常退出（本轮无有效产物），目录里的同名文件是上一轮残留，"
                         "不可作为本次结论" % name)
    if "交易所" in skip or "交易所" in failed_src:
        exch = None
    if "证监局" in skip or "证监局" in failed_src:
        csrc = None
    if "中基协" in skip or "中基协" in failed_src:
        amac = None
    if "交易所" in skip:
        print("  - 交易所：本轮显式跳过（--skip），忽略目录里可能存在的旧 exch.json")

    # 窗口要先算出来：交易所逐所判定需要「抓取窗口起点」做比较基准
    win = window_of(csrc, amac, exch)
    crawl_start = win[0] if win else ""

    for name, fn, obj in (("证监局", check_csrc, csrc),
                          ("中基协", lambda o, f, w, n: check_amac(o, f, w, n, a.max_pages), amac),
                          ("交易所", lambda o, f, w, n: check_exch(o, f, w, n, crawl_start), exch)):
        if obj is None:
            if name not in skip:
                print("  - %s：无产物（未跑该来源，跳过）" % name)
            continue
        cnt = fn(obj, fails, warns, notes)
        totals[name] = cnt
        print("  - %s：%s 条" % (name, cnt))

    # R1 已跑来源条目为 0 —— 但「0」不等于「漏」：
    #   2026-10-28 实测：日常增量窗口只有 8 天，中基协按批次发布，本轮确实 0 条新增，
    #   旧判据直接 FAIL，等于每个安静的更新日都被卡住。
    #   修正：能证明「平台最新一条早于本次抓取窗口起点」就判 0 条正常（记 note），
    #   否则仍 FAIL（无法区分真无新增与限流/结构变化）。
    #   交易所已在 check_exch 里逐所判过，这里不重复判（否则会按"总数 0"再 FAIL 一次）。
    for name, cnt in totals.items():
        if cnt or name == "交易所":
            continue
        latest = _latest_seen(name, csrc, amac, exch)
        if latest and crawl_start and latest < crawl_start:
            notes.setdefault("zero_ok", []).append(
                "%s 本轮 0 条，平台最新一条为 %s（早于抓取窗口起点 %s）→ 判为真无新增"
                % (name, latest, crawl_start))
            warns.append("%s 本轮 0 条新增（平台最新 %s < 窗口起点 %s），已核实为真无新增"
                         % (name, latest, crawl_start))
        else:
            fails.append("%s 抓取条目为 0，且平台最新日期%s —— 不能视为『本期无新增』"
                         % (name, ("为 %s、不早于窗口起点" % latest) if latest
                            else "未知（无自证指标）"))

    # R6 环比
    window = win
    prev = _load(a.prev) if a.prev else None
    if prev:
        prev_t = prev.get("totals") or {}
        same_window = prev.get("window") == window and window is not None
        for name, cnt in totals.items():
            pv = prev_t.get(name)
            if not pv or pv <= 0:
                continue
            if cnt < pv * 0.5:
                msg = "环比骤降：%s 上轮 %d → 本轮 %d（-%.0f%%）" % (
                    name, pv, cnt, 100 * (1 - cnt / pv))
                (fails if same_window else warns).append(
                    msg + ("；同一窗口内骤降=必然漏抓" if same_window else "；窗口已滚动，仅提示"))

    rep = {"work": os.path.abspath(a.work), "window": window, "totals": totals,
           "notes": notes, "fails": fails, "warns": warns}
    if a.out:
        try:
            with open(a.out, "w", encoding="utf-8") as f:
                json.dump(rep, f, ensure_ascii=False, indent=1)
            print("  自证报告已写入 %s" % a.out)
        except Exception as e:
            warns.append("自证报告写出失败：%s" % e)

    print("\n" + "=" * 60)
    if warns:
        print("WARN %d 项（不阻断，但需知晓）：" % len(warns))
        for w in warns:
            print("   · " + w)
    if fails:
        print("FAIL %d 项（本次抓取不完整，禁止继续入库）：" % len(fails))
        for x in fails:
            print("   ✗ " + x)
        print("=" * 60)
        return 1
    print("PASS 抓取完整性自证通过（%d 项提示）" % len(warns))
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
