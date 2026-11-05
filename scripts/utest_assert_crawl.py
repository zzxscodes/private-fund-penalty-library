# -*- coding: utf-8 -*-
"""assert_crawl_complete.py 的固定单测（2026-10-28 新增）。

为什么需要：这份断言的判据在 2026-10-28 一天内被真实事故改了四次——
  R1 从「0 条即 FAIL」改成「能自证平台最新 < 窗口起点就 WARN」；
  R4 从「三所总数」改成「逐所」；
  新增 R8（正文失败即 FAIL）、R9（抓取脚本崩溃不得沿用上轮产物）；
  新增 --skip（主动跳过不算失败）。
判据是"拦住漏案"的唯一闸门，改错了要么天天误杀、要么静默放行，
所以必须有可复跑的用例。

用法：python utest_assert_crawl.py   （全部通过退出码 0）
"""
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, "assert_crawl_complete.py")
WIN = ["2026-10-20", "2026-10-28"]

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)


def build(case, files):
    d = tempfile.mkdtemp(prefix="utest_assert_")
    for name, obj in files.items():
        with open(os.path.join(d, name), "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False)
    return d


def run(d, extra=()):
    p = subprocess.run([sys.executable, TARGET, "--work", d] + list(extra),
                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    return p.returncode, (p.stdout or b"").decode("utf-8", "ignore")


def exch(sse=0, szse=0, bse=0, latest=None, warnings=None):
    return {"window": WIN,
            "stats": {"sse_total": sse, "szse_total": szse, "bse_total": bse,
                      "latest": latest or {}, "warnings": warnings or []}}


def csrc(total=1, chans=None):
    return {"window": WIN,
            "items": [{"title": "x"}] * total,
            "stats": {"channels": chans if chans is not None else
                      [{"bureau_cn": "广东", "col": "行政监管措施", "platform_total": 9,
                        "pages_fetched": 1, "terminated_by": "start_boundary",
                        "items_window": total, "latest_seen": "2026-10-27"}]}}


CASES = []


def case(name, expect, files, extra=()):
    CASES.append((name, expect, files, extra))


# 1. 正常：三源都有数据，翻页正常终止 → PASS
case("正常抓取", 0, {"csrc.json": csrc()})

# 2. R2：翻页被中断（json_error）→ FAIL
case("翻页被中断", 1, {"csrc.json": csrc(chans=[
    {"bureau_cn": "新疆", "col": "行政监管措施", "platform_total": 0,
     "pages_fetched": 1, "terminated_by": "json_error", "items_window": 0}])})

# 3. R8：列表抓到但正文失败（广东 9 条事故型）→ FAIL
case("正文失败", 1, {"csrc.json": {
    "window": WIN, "items": [{"title": "x"}],
    "stats": {"channels": [{"bureau_cn": "广东", "col": "行政监管措施",
                            "platform_total": 155, "pages_fetched": 8,
                            "terminated_by": "short_page", "items_window": 1,
                            "latest_seen": "2026-10-27"}],
              "failed_items": [{"bureau_cn": "广东", "col": "行政监管措施",
                                "reason": "正文接口无返回", "title": "关于对某公司…"}]}}})

# 4. R1 修正：0 条但能自证「平台最新 < 窗口起点」→ PASS（真无新增）
case("零条可自证", 0, {"amac.json": {
    "window": WIN, "items": [],
    "columns": {"scfjg": {"label": "受处分机构", "total_pages": 35, "pages_fetched": 35,
                          "full": True, "items": 0, "latest_published": "2026-10-16"}}}})

# 5. R1：0 条且无自证 → FAIL
case("零条无自证", 1, {"amac.json": {"window": WIN, "items": [], "columns": {}}})

# 6. R4 逐所：三所 0 条但都有最新日期且早于窗口 → PASS
case("交易所逐所可自证", 0, {"exch.json": exch(
    latest={"sse": "2026-09-30", "szse": "2026-08-28", "bse": "2026-09-11"},
    warnings=[{"venue": "szse", "kind": "throttle_probe", "tag": "x"}])})

# 7. R4 逐所：深交所 0 条 + 频控 + 拿不到该所最新日期 → FAIL（另两所不背锅）
case("单所频控不可自证", 1, {"exch.json": exch(
    latest={"sse": "2026-09-30", "bse": "2026-09-11"},
    warnings=[{"venue": "szse", "kind": "throttle_probe", "tag": "x"}])})

# 8. 检索式栏目：分页恒 0 属结构，只 WARN 不 FAIL
case("检索式栏目仅提示", 0, {"exch.json": {
    "window": WIN,
    "stats": {"sse_total": 0, "szse_total": 0, "bse_total": 0,
              "latest": {"sse": "2026-09-30", "szse": "2026-08-28", "bse": "2026-09-11"},
              "search_only": {"ZQ_JGCS": {"name": "债券监管措施", "tab": "tab1",
                                          "probe_queries": 4, "probe_hits": 0}},
              "warnings": []}}})

# 9. R9：脚本崩溃 → FAIL（哪怕目录里躺着一份"看起来正常"的旧产物）
case("脚本崩溃不得沿用旧产物", 1,
     {"exch.json": exch(latest={"sse": "2026-09-30", "szse": "2026-08-28",
                                "bse": "2026-09-11"}),
      "crawl_failed.json": ["交易所"]})

# 10. --skip：主动跳过不算失败（与 R9 严格区分）
case("主动跳过不算失败", 0,
     {"exch.json": exch(latest={"sse": "2026-09-30", "szse": "2026-08-28",
                                "bse": "2026-09-11"})},
     extra=("--skip", "交易所"))


def main():
    bad = 0
    for name, expect, files, extra in CASES:
        d = build(name, files)
        rc, out = run(d, extra)
        ok = (rc == expect)
        print("%s %-22s 期望退出码=%d 实际=%d" % ("PASS" if ok else "FAIL", name, expect, rc))
        if not ok:
            bad += 1
            print("     输出：\n" + "\n".join("       " + x for x in out.splitlines()[-12:]))
    print("-" * 60)
    print("通过 %d / 失败 %d" % (len(CASES) - bad, bad))
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
