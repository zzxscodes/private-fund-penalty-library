# -*- coding: utf-8 -*-
"""交付前置自检总入口 —— 一条命令证明「每一格都查过、且都对过文件」。

用户 2026-11-02 的要求（原话）：
    「台账又是需求中重要的内容又是检查校验的重点和校验案例库的重要依据，所以技能一定要确保
      能够在每次执行之后能够通过检查每一格的内容和对照每一格的内容对应的实际文件
      来解决可能存在的已知和未知问题。」

这句话拆成三个可验证的命题，本脚本一件件证：
  ① **每一格**都查过   → `--coverage` 打印 15 列 × 检查器矩阵，任一一列没有检查器就 FAIL；
  ② **每一格都对过实际文件** → 至少一条「文件对照」判据（Z11/Z13/Z14/F3/F4），否则该列只能内部自洽；
  ③ **已知 + 未知**都覆盖 → 已知走 audit_cells/columns/filenames/semantics，未知走 Z11~Z14 离群探测。

用法：
    python run_selfcheck.py --root <库根> --start <起> --end <止>    # 跑全套（不含 OCR，约 10 分钟）
    python run_selfcheck.py ... --ocr                                 # 扫描件也 OCR 后判（更慢、更准）
    python run_selfcheck.py ... --report out.md                       # 落一份报告
    python run_selfcheck.py --coverage                                # 只打印覆盖矩阵自证

**--start/--end 必传**：窗口是业务参数（建库/上次增量固定的），不是「今天-365 天」。
不传会被拒（EXIT=2），因为 verify 的默认窗口与库窗口差一天就会误报 A2「窗口外记录」。

返回码：0 = 全部通过（问题 0）；1 = 有问题格；2 = 参数不全或检查器覆盖不全（自检本身失效）。
"""
import argparse
import os
import re
import subprocess
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# 15 列 × 检查器矩阵 —— 这张表就是「每一格都查过」的证明。
# 每列至少要有 1 条「形态」判据 + 1 条「文件对照」判据，缺一即视为自检失效。
# ---------------------------------------------------------------------------
COVERAGE = [
    ("A 序号",      ["audit_semantics:S1", "audit_cells:X"],                      ["verify_library:A1（行数↔PDF 数）"]),
    ("B 处罚时间",  ["audit_semantics:S2", "audit_cells:X"],                      ["verify_library:B4/B6（落款抽取）"]),
    ("C 挂网日",    ["audit_semantics:S2"],                                       ["verify_library:B4（与平台挂网日对账）"]),
    ("D 被处罚机构/人", ["audit_semantics:S3", "audit_cells:X"],                  ["verify_library:B5（主体名回现）", "audit_files:F3/F4"]),
    ("E 处罚措施",  ["audit_semantics:S4/S14", "audit_cells:X"],                  ["verify_library:B8（措施回现）"]),
    ("F 核心处罚事由", ["audit_semantics:S5", "audit_cells:X28/X29", "check_grounds_complete:Z14"],
                                                                                  ["check_fields_source:Z11（6-gram 回现）", "check_grounds_complete:Z14（条目完整性）"]),
    ("G 核心证据材料", ["audit_semantics:S6", "audit_cells:X"],                   ["check_evidence_col:Z13（占位漏抽/内容串味）"]),
    ("H 核心处罚依据", ["audit_semantics:S7", "audit_cells:X31", "audit_basis_comments:Z15",
                        "check_basis_attribution:Z16"],
                                                                                  ["verify_library:D1（法规清单）", "check_fields_source:Z11(--all-cols)", "audit_basis_comments:Z15（批注↔06 法条正文）",
                                                                                   "check_basis_attribution:Z16（法规归属↔文书原文）"]),
    ("I 申辩理由",  ["audit_semantics:S8", "audit_cells:X32"],                    ["check_fields_source:Z11(--all-cols)"]),
    ("J 违规类型",  ["audit_semantics:S9", "audit_columns"],                      ["verify_library:C1b/C5（分类目录有文件）"]),
    ("K 机构来源",  ["audit_semantics:S10", "audit_columns"],                     ["verify_library:A3（判重带来源）"]),
    ("L 案例文件",  ["audit_filenames:E1~E6"],                                    ["verify_library:A1/A2（孤儿/缺失）", "audit_files:F1/F2/F8"]),
    ("M 存放路径",  ["audit_semantics:S11"],                                      ["verify_library:C2（路径末级=违规类型）", "audit_files:F8（磁盘孤儿）"]),
    ("N 原文链接",  ["audit_semantics:S12"],                                      ["verify_library:A6（平台对账，按需）"]),
    ("O 是否典型",  ["audit_semantics:S13"],                                      ["verify_library:C9（典型副本存在）"]),
]


def print_coverage():
    print("=" * 78)
    print("检查器覆盖矩阵（证明「每一格都查过、都对过文件」）")
    print("=" * 78)
    bad = []
    for col, form, file_ in COVERAGE:
        ok = bool(form) and bool(file_)
        print("  %-16s 形态: %-38s 文件: %s%s"
              % (col, "、".join(form) or "—", "、".join(file_) or "—", "" if ok else "   ← 缺口"))
        if not ok:
            bad.append(col)
    print("\n共 %d 列：%d 列同时具备「形态判据 + 文件对照」" % (len(COVERAGE), len(COVERAGE) - len(bad)))
    if bad:
        print("✗ 以下列存在覆盖缺口，自检不成立：%s" % "、".join(bad))
        return 2
    print("✓ 15 列全部具备「形态判据 + 文件对照」——不存在只靠内部自洽的列")
    return 0


def run(name, cmd, timeout=3600):
    t0 = datetime.now()
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=timeout)
        out = (p.stdout or "") + (p.stderr or "")
        rc = p.returncode
    except subprocess.TimeoutExpired:
        out, rc = "TIMEOUT", 124
    return {"name": name, "cmd": cmd, "rc": rc, "out": out,
            "sec": (datetime.now() - t0).total_seconds()}


def num(pat, txt, default=None):
    m = re.search(pat, txt)
    return int(m.group(1)) if m else default


def summarize(steps):
    """把各工具输出归一成 (问题数, 观察数, 结论一句话)。"""
    rows = []
    for s in steps:
        o, n = s["out"], s["name"]
        if n == "verify_library":
            # verify 的问题是一串「· xxx」列表，没有 problems=N 字样；
            # 旧写法只认 `problems=N` → 一有问题就显示「?」（2026-11-03 实测）
            prob = num(r"problems\s*[=:]\s*(\d+)", o)
            if prob is None:
                seg = o.rsplit("结论：存在问题", 1)[-1] if "结论：存在问题" in o else ""
                prob = len(re.findall(r"^\s+·\s", seg, re.M))
                if prob == 0 and "全部校验通过" in o:
                    prob = 0
            rows.append(("① 数量/对账", n, s["rc"], prob, "EXIT=%d" % s["rc"]))
        elif n == "audit_cells":
            prob = num(r"问题格合计\s*(\d+)", o, 0)
            obs = num(r"观察项\s*(\d+)", o, 0)
            rows.append(("② 逐格形态", n, s["rc"], prob, "观察 %d" % obs))
        elif n == "audit_columns":
            rows.append(("② 逐列", n, s["rc"], 0 if "全部列通过" in o else 1, "全部列通过" if "全部列通过" in o else "见输出"))
        elif n == "audit_filenames":
            prob = num(r"问题\s*(\d+)\s*条", o)
            obs = num(r"观察\s*(\d+)\s*条", o, 0)
            rows.append(("② 命名", n, s["rc"], prob, "E1~E6｜观察 %d" % obs))
        elif n == "audit_semantics":
            prob = 0 if "未发现语义异常" in o else num(r"^\s*列 [A-O]\s+(\d+) 条", o, 1)
            rows.append(("② 逐格语义", n, s["rc"], prob, "S1~S14"))
        elif n == "audit_files":
            prob = num(r"问题\s*(\d+)\s*条", o, 0)
            obs = num(r"观察\s*(\d+)\s*条", o, 0)
            rows.append(("③ 逐份文件", n, s["rc"], prob, "观察 %d" % obs))
        elif n == "check_fields_source":
            prob = num(r"线索\s*[:：]?\s*(\d+)", o)
            rows.append(("③ 台账↔文书回现", n, s["rc"], prob, "Z11"))
        elif n == "check_evidence_col":
            prob = num(r"漏抽）\s*[:：]\s*(\d+)", o, 0) + num(r"（串味）\s*[:：]\s*(\d+)", o, 0)
            rows.append(("③ 证据列双向", n, s["rc"], prob, "Z13"))
        elif n == "check_grounds_complete":
            prob = num(r"疑似漏条\s*[:：]\s*(\d+)", o, 0)
            obs = num(r"观察项（「此外…还关注到」非违规条目）\s*[:：]\s*(\d+)", o, 0)
            rows.append(("③ 事由完整性", n, s["rc"], prob, "Z14｜观察 %d" % obs))
        elif n == "audit_basis_comments":
            prob = num(r"问题\s*(\d+)\s*｜观察", o)
            obs = num(r"｜观察\s*(\d+)", o, 0)
            rows.append(("③ 依据批注", n, s["rc"], prob, "Z15｜观察 %d" % obs))
        elif n == "check_basis_attribution":
            prob = num(r"问题\s*(\d+)\s*条｜观察", o)
            obs = num(r"｜观察\s*(\d+)\s*条", o, 0)
            rows.append(("③ 依据归属", n, s["rc"], prob, "Z16｜观察 %d" % obs))
        elif n == "utest_audit_cells":
            rows.append(("② 检查器自测", n, s["rc"], 0 if "PASS" in o else 1, "正例9/反例6"))
        elif n == "selftest_portable":
            rows.append(("④ 可迁移性", n, s["rc"], 0 if "PASS" in o else 1, "无硬编码路径"))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default="")
    ap.add_argument("--start", default="")
    ap.add_argument("--end", default="")
    ap.add_argument("--ocr", action="store_true")
    ap.add_argument("--report", default="")
    ap.add_argument("--coverage", action="store_true")
    ap.add_argument("--only", default="", help="只跑匹配的步骤（正则）")
    a = ap.parse_args()

    if a.coverage or not a.root:
        return print_coverage()

    # 窗口是**业务参数**，不是代码默认值。不传就按 verify 的默认「今天-365 天」判 A2，
    # 而本库窗口是建库/上次增量固定的（如 2025-10-31~2026-10-30）——两者差一天，
    # A2 就会报「窗口外记录」，把正确的台账判成错的（2026-11-02 实测）。
    if not (a.start and a.end):
        print("✗ 缺少窗口参数 --start / --end")
        print("  本库窗口是建库时固定的，不是「今天-365 天」；不传就会误报 A2 窗口外记录。")
        print("  正确用法：run_selfcheck.py --root <库根> --start 2025-10-31 --end 2026-10-30")
        return 2

    root = os.path.abspath(a.root)
    py = sys.executable
    S = HERE

    steps_spec = [
        ("verify_library", [py, os.path.join(S, "verify_library.py"), "--root", root], False),
        ("audit_cells", [py, os.path.join(S, "audit_cells.py"), "--root", root], True),
        ("audit_columns", [py, os.path.join(S, "audit_columns.py"), "--root", root], True),
        ("audit_filenames", [py, os.path.join(S, "audit_filenames.py"), "--root", root], True),
        ("audit_semantics", [py, os.path.join(S, "audit_semantics.py"), "--root", root], True),
        ("audit_files", [py, os.path.join(S, "audit_files.py"), "--root", root], True),
        ("check_fields_source", [py, os.path.join(S, "check_fields_source.py"), "--root", root], False),
        ("check_evidence_col", [py, os.path.join(S, "check_evidence_col.py"), "--root", root], False),
        ("check_grounds_complete", [py, os.path.join(S, "check_grounds_complete.py"), "--root", root], False),
        ("audit_basis_comments", [py, os.path.join(S, "audit_basis_comments.py"), "--root", root], False),
        ("check_basis_attribution", [py, os.path.join(S, "check_basis_attribution.py"), "--root", root], False),
        ("utest_audit_cells", [py, os.path.join(S, "utest_audit_cells.py")], True),
        ("selftest_portable", [py, os.path.join(S, "selftest_portable.py")], True),
    ]
    if a.start and a.end:
        i = next(i for i, x in enumerate(steps_spec) if x[0] == "verify_library")
        steps_spec[i][1].extend(["--start", a.start, "--end", a.end])
    if a.ocr:
        for nm in ("check_evidence_col", "check_grounds_complete"):
            i = next(i for i, x in enumerate(steps_spec) if x[0] == nm)
            steps_spec[i][1].append("--ocr")

    only = re.compile(a.only) if a.only else None

    steps = []
    for nm, cmd, _fast in steps_spec:
        if only and not only.search(nm):
            continue
        print("▶ %s ..." % nm, flush=True)
        steps.append(run(nm, cmd))
        print("   rc=%d  %.1fs" % (steps[-1]["rc"], steps[-1]["sec"]), flush=True)

    rows = summarize(steps)
    print("\n" + "=" * 78)
    print("自检汇总（%s）" % datetime.now().strftime("%Y-%m-%d %H:%M"))
    print("=" * 78)
    print("%-14s %-24s %-6s %-6s %s" % ("阶段", "工具", "退出码", "问题数", "说明"))
    total = 0
    for stage, nm, rc, prob, note in rows:
        p = "?" if prob is None else prob
        if isinstance(prob, int):
            total += prob
        print("%-14s %-24s %-6d %-6s %s" % (stage, nm, rc, p, note))
    print("-" * 78)
    print("问题合计：%s｜阶段覆盖 %d/%d" % (total, len(rows), len(steps_spec)))

    cov = print_coverage()

    verdict = 0 if (total == 0 and all(s["rc"] == 0 for s in steps) and cov == 0) else 1
    print("\n结论：%s" % ("✓ 全部通过，可交付" if verdict == 0 else "✗ 未通过，见上方问题"))

    if a.report:
        with open(a.report, "w", encoding="utf-8") as fh:
            fh.write("# 交付前自检报告（自动生成）\n\n")
            fh.write("- 生成时间：%s\n- 库根：`%s`\n- 结论：**%s**\n\n"
                     % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), root,
                        "全部通过" if verdict == 0 else "未通过"))
            fh.write("## 一、分段结果\n\n| 阶段 | 工具 | 退出码 | 问题数 | 说明 |\n|---|---|---|---|---|\n")
            for stage, nm, rc, prob, note in rows:
                fh.write("| %s | `%s` | %d | %s | %s |\n" % (stage, nm, rc, prob, note))
            fh.write("\n## 二、检查器覆盖矩阵\n\n| 列 | 形态判据 | 文件对照 |\n|---|---|---|\n")
            for col, form, file_ in COVERAGE:
                fh.write("| %s | %s | %s |\n" % (col, "、".join(form), "、".join(file_)))
            fh.write("\n## 三、原始输出\n\n")
            for s in steps:
                fh.write("### %s（rc=%d, %.1fs）\n\n```\n%s\n```\n\n"
                         % (s["name"], s["rc"], s["sec"], s["out"][-4000:]))
        print("报告：%s" % a.report)

    return verdict


if __name__ == "__main__":
    sys.exit(main())
