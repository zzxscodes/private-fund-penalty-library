# -*- coding: utf-8 -*-
"""extract_core 全量回归：三核心列 + 处罚措施 空 0、复议/诉讼模板不误入。

用法：
  python selftest_extract.py <tmp_pilot_dir>
  （默认取脚本同级的 ../../ 工作空间下的 tmp_pilot；找不到则用 CWD）

退出码 0 = 全部通过；非 0 = 存在空列或模板误入（CI/定时任务可据此报警）。
"""
import json
import os
import re
import sys
import glob

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import extract_core as E  # noqa: E402


def _load_csrc(tmp):
    p = os.path.join(tmp, "csrc.json")
    if not os.path.exists(p):
        return []
    d = json.load(open(p, encoding="utf-8"))
    return d.get("items", [])


def _load_amac(tmp):
    out = []
    for p in glob.glob(os.path.join(tmp, "texts_amac", "*.txt")):
        try:
            out.append((os.path.basename(p), open(p, encoding="utf-8", errors="ignore").read()))
        except Exception:
            pass
    return out


LEAK = re.compile(r"如不服|行政复议|人民法院|汇交|将罚款|将罚没款")


def main():
    tmp = sys.argv[1] if len(sys.argv) > 1 else None
    if not tmp:
        for cand in (os.path.join(os.path.dirname(HERE), "tmp_pilot"),
                     os.path.join(os.getcwd(), "tmp_pilot")):
            if os.path.isdir(cand):
                tmp = cand
                break
    if not tmp or not os.path.isdir(tmp):
        print("找不到 tmp_pilot 目录；请传入路径。", file=sys.stderr)
        return 2

    csrc = _load_csrc(tmp)
    amac = _load_amac(tmp)

    stats = {"cs_empty": 0, "am_empty": 0, "cs_leak": 0, "am_leak": 0, "cs_n": 0, "am_n": 0}
    fails = []

    for it in csrc:
        t = it.get("text") or ""
        if not t:
            continue
        stats["cs_n"] += 1
        r = E.extract_csrc(t)
        if not (r["measures"] and r["legal_basis"] and r["evidence"] and r["grounds"]):
            stats["cs_empty"] += 1
            fails.append(("CSRC_EMPTY", it.get("wh", "<e>"),
                          {k: (v or "")[:40] for k, v in r.items()}))
        if LEAK.search(r["measures"] or ""):
            stats["cs_leak"] += 1
            fails.append(("CSRC_LEAK", it.get("wh", "<e>"), r["measures"][:60]))

    for name, t in amac:
        stats["am_n"] += 1
        r = E.extract_amac(t)
        if not (r["measures"] and r["legal_basis"] and r["evidence"] and r["grounds"]):
            stats["am_empty"] += 1
            fails.append(("AMAC_EMPTY", name, {k: (v or "")[:40] for k, v in r.items()}))
        if LEAK.search(r["measures"] or ""):
            stats["am_leak"] += 1
            fails.append(("AMAC_LEAK", name, r["measures"][:60]))

    # ---- 契约层校验：规则 → 升格待模型确认 → 收口（无模型填值） ----
    # 本数据集规则产 0 空，故收口后必须 0 个 PENDING_MODEL、0 个诚实标记残余（不该误伤）。
    pend = 0
    for it in csrc:
        t = it.get("text") or ""
        if not t:
            continue
        fin = E.finalize(E.mark_pending(E.extract_csrc(t)))
        pend += sum(1 for v in fin.values() if v == E.PENDING_MODEL)
    for name, t in amac:
        fin = E.finalize(E.mark_pending(E.extract_amac(t)))
        pend += sum(1 for v in fin.values() if v == E.PENDING_MODEL)
    stats["pending"] = pend

    print("=" * 60)
    print("回归结果：")
    print("  证监局 %d 份 → 空列 %d，模板误入 %d" % (stats["cs_n"], stats["cs_empty"], stats["cs_leak"]))
    print("  中基协 %d 份 → 空列 %d，模板误入 %d" % (stats["am_n"], stats["am_empty"], stats["am_leak"]))
    print("  契约层：升格后残留 PENDING_MODEL = %d（应为 0）" % pend)
    print("=" * 60)
    if fails:
        print("失败样本：")
        for f in fails[:30]:
            print(" ", f)
        return 1
    if pend:
        print("⚠️ 契约层发现残留待模型确认 %d 处，需排查规则遗漏。" % pend)
        return 1
    print("✅ 全部通过：三核心列 + 处罚措施 空 0，复议/诉讼/缴款模板 0 误入；契约层 0 残留。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
