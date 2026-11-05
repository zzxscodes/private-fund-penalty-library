# -*- coding: utf-8 -*-
"""写入期日期断言单测（2026-10-30）。

背景（用户裁定）：「如果处理时确保没问题那其实没必要去验证复核的，除了某些必要的在
处理过程没有确保」——所以把「事后全量 OCR 复扫处罚时间」前移成「写入期硬断言」：

  * 中基协 / 交易所决定书按惯例必有落款 → 抽不到就是抽取失败 → 写「待核」，**禁止**拿挂网日顶替；
  * 证监局官网内联正文确有 38 条只写检查/整改日期、无落款 → 允许取平台发文日期，但必须留痕。

本测试守住两条底线：
  1. `build_pipeline` 里**不允许**再出现 `extract_sign_date(...) or pub` 这种静默回落（正则扫描防回写）；
  2. `date_or_hold` 三种分支（sign / publish / hold）行为正确，且中基协抽不到落款时返回「待核」。

用法：python utest_build_date.py     # 全绿退出码 0
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_pipeline as BP  # noqa: E402

OK = FAIL = 0


def check(name, cond, detail=""):
    global OK, FAIL
    if cond:
        OK += 1
        print("  [PASS] %s" % name)
    else:
        FAIL += 1
        print("  [FAIL] %s %s" % (name, detail))


def main():
    print("== 1. 防回写：禁止 `extract_sign_date(...) or pub` 静默回落 ==")
    src = open(os.path.join(HERE, "build_pipeline.py"), encoding="utf-8").read()
    bad = re.findall(r"extract_sign_date\([^)]*\)\s*or\s+\w+", src)
    check("无 `extract_sign_date(...) or <fallback>` 写法", not bad, str(bad[:3]))
    check("date_or_hold 已定义", "def date_or_hold(" in src)

    print("\n== 2. date_or_hold 三分支 ==")
    # 正文档末有落款 → sign
    t1 = "经查，你公司存在以下违规……\n中国证券投资基金业协会\n2026年5月26日"
    d, s = BP.date_or_hold(t1, "2026-07-30", "中基协", "t1")
    check("中基协有落款 → sign 2026-05-26", (d, s) == ("2026-05-26", "sign"), f"{d}/{s}")

    # 中基协抽不到落款 → hold（待核），绝不等于挂网日
    t2 = "经查，你公司存在以下违规……（扫描件 OCR 全糊）"
    d2, s2 = BP.date_or_hold(t2, "2026-07-30", "中基协", "t2")
    check("中基协无落款 → hold 待核", (d2, s2) == (BP.DATE_HOLD, "hold"), f"{d2}/{s2}")
    check("hold 时绝不返回挂网日", d2 != "2026-07-30", d2)

    # 交易所同样必须有落款
    d3, s3 = BP.date_or_hold(t2, "2026-07-30", "交易所", "t3")
    check("交易所无落款 → hold", (d3, s3) == (BP.DATE_HOLD, "hold"), f"{d3}/{s3}")

    # 证监局无落款 → 允许取平台发文日，但来源标 publish（留痕）
    d4, s4 = BP.date_or_hold(t2, "2026-07-30", "证监局", "t4")
    check("证监局无落款 → publish 2026-07-30", (d4, s4) == ("2026-07-30", "publish"), f"{d4}/{s4}")

    # 证监局有落款 → 仍以落款为准
    d5, s5 = BP.date_or_hold(t1, "2026-07-30", "证监局", "t5")
    check("证监局有落款 → sign 优先", (d5, s5) == ("2026-05-26", "sign"), f"{d5}/{s5}")

    print("\n== 3. 留痕：hold / publish 都必须进 AUDIT ==")
    n0 = len(BP.AUDIT)
    BP.date_or_hold(t2, "2026-07-30", "中基协", "audit-hold")
    BP.date_or_hold(t2, "2026-07-30", "证监局", "audit-pub")
    new = BP.AUDIT[n0:]
    check("新增 2 条审计记录", len(new) == 2, str(len(new)))
    check("hold 记录 reason=date_hold", any(x.get("reason") == "date_hold" for x in new), str(new[:2]))
    check("publish 记录 reason=date_publish", any(x.get("reason") == "date_publish" for x in new))

    print("\n== 4. OCR 杂字符落款仍能抽到（B9 回归）==")
    for txt, want in [
        ("中国证券投资基金业协会\n2026,年5月。26日", "2026-05-26"),
        ("中国证券投资基金业协会\n2025°年11月25日", "2025-11-25"),
        ("中国证券投资基金业协会\n2025年12月-24日", "2025-12-24"),
    ]:
        got = BP.extract_sign_date(txt)
        check("OCR 杂字符 %-22s → %s" % (txt.split("\n")[-1], want), got == want, got)

    print("\n%s ｜ 通过 %d ｜ 失败 %d" % ("全绿" if FAIL == 0 else "有失败", OK, FAIL))
    sys.exit(0 if FAIL == 0 else 1)


if __name__ == "__main__":
    main()
