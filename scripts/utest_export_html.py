# -*- coding: utf-8 -*-
"""网页台账的浏览器回归测试（utest：只读，不改库、不改台账）。

不是数据校验——数据校验是 `verify_library.py` / `run_selfcheck.py` 的活。
这里只回答一个问题：**生成的 HTML 在真浏览器里点得动吗**。
台账是对的、网页打不开，用户一样拿不到东西；这两件事必须分开验。

覆盖：渲染条数、facet 计数自洽、来源切换、勾选筛选、字段限定搜索、高亮、
详情各段是否齐全、法规标签跳筛选、滚动到底、`#L444` 直达、排序、CSV（BOM/列数）、
键盘、说明弹层、窄屏、全程无 JS 报错。

用法：
    python utest_export_html.py --root <库根>            # 跑全部断言
    python utest_export_html.py --root <库根> --shots d  # 另存截图供人工看
退出码 0 全过 / 1 有失败。
"""

from __future__ import annotations

import argparse
import io
import os
import shutil
import sys
import tempfile
from urllib.parse import quote

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("缺少依赖：pip install playwright && playwright install chromium")
    sys.exit(2)

import export_basis_html as EX


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="网页台账浏览器回归测试")
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--shots", help="截图输出目录（可选，给人眼看）")
    ap.add_argument("--keep", action="store_true", help="保留临时生成的 html")
    a = ap.parse_args(argv)

    tmp = tempfile.mkdtemp(prefix="wbhtml_")
    out = os.path.join(tmp, "web.html")
    rc = EX.main(["--root", os.path.abspath(a.root), "--out", out])
    if rc != 0 or not os.path.exists(out):
        print("生成失败，后面不用测了")
        return 1
    url = "file:///" + quote(out.replace("\\", "/"))

    fails, passes = [], []

    def ok(cond, msg, extra=""):
        (passes if cond else fails).append(
            ("PASS " if cond else "FAIL ") + msg + ("  " + str(extra) if extra != "" else ""))

    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1680, "height": 1050})
            errs = []
            pg.on("pageerror", lambda e: errs.append(str(e)))
            pg.on("console", lambda m: errs.append("console:" + m.text) if m.type == "error" else None)
            pg.goto(url)
            pg.wait_for_timeout(900)

            total = EX.collect(os.path.abspath(a.root))[0]
            N = len(total)

            ok(not errs, "加载无 JS 报错", errs[:3])
            ok(pg.eval_on_selector_all(".row", "e=>e.length") == N, "列表渲染 %d 行" % N)

            tot = pg.evaluate("""() => [...document.querySelectorAll('.fg[data-dim="k"] .nu')]
                .reduce((a,e)=>a+ +e.textContent,0)""")
            ok(tot == N, "违规类型 facet 计数合计 = 总数", tot)

            seg = pg.evaluate("""() => [...document.querySelectorAll('.seg[data-src] b')]
                .map(e=>+e.textContent)""")
            ok(sum(seg[1:]) == N, "来源分项之和 = 总数", seg)

            # 来源切换
            for src, exp in pg.evaluate("""() => [...document.querySelectorAll('.seg[data-src]')]
                    .slice(1).map(e=>[e.dataset.src, +e.querySelector('b').textContent])"""):
                pg.click('.seg[data-src="%s"]' % src)
                pg.wait_for_timeout(320)
                ok(pg.eval_on_selector_all(".row", "e=>e.length") == exp,
                   "切到「%s」→ %d 条" % (src, exp))
            pg.click('.seg[data-src=""]')
            pg.wait_for_timeout(300)

            # 勾选一个违规类型，列表条数应等于该 facet 数字
            first = pg.evaluate("""() => {
                const o = document.querySelector('.fg[data-dim="k"] .opt');
                const n = o.querySelector('.nu').textContent;
                o.querySelector('input').click();
                return {v: o.dataset.v, n: +n};
            }""")
            pg.wait_for_timeout(420)
            cnt = pg.eval_on_selector_all(".row", "e=>e.length")
            ok(cnt == first["n"], "勾选「%s」→ 与 facet 数字一致" % first["v"], "%d/%d" % (cnt, first["n"]))

            # 加第二个维度，验证维度间是「并且」
            pg.evaluate("""() => document.querySelector('.fg[data-dim="acts"] .opt input').click()""")
            pg.wait_for_timeout(420)
            both = pg.eval_on_selector_all(".row", "e=>e.length")
            ok(both <= cnt, "两个维度叠加后条数不增加（并且关系）", "%d → %d" % (cnt, both))

            pg.click("#bClear")
            pg.wait_for_timeout(360)
            ok(pg.eval_on_selector_all(".row", "e=>e.length") == N, "清空条件 → 回到总数")

            # 字段限定搜索
            pg.fill("#q", "来源:中基协 类型:募集")
            pg.wait_for_timeout(460)
            nq = pg.eval_on_selector_all(".row", "e=>e.length")
            ok(0 < nq < N, "字段限定搜索（来源:… 类型:…）生效", nq)
            st = pg.inner_text("#stat")
            ok("共 %d 条" % N in st and "当前" in st, "状态栏给出总数与当前数", st)

            # 关键词 + 高亮
            pg.fill("#q", "适当性")
            pg.wait_for_timeout(460)
            ok(pg.eval_on_selector_all(".row", "e=>e.length") > 0, "关键词有结果")
            ok(pg.eval_on_selector_all("#dinner mark", "e=>e.length") > 0, "详情内命中高亮")
            pg.fill("#q", "")
            pg.wait_for_timeout(400)

            # 详情完整性
            pg.hover('.row[data-i="0"]')
            pg.wait_for_timeout(320)
            dt = pg.inner_text("#dinner")
            for key in ["核心处罚依据（法条全文）", "处罚措施（机构+人员）", "核心处罚事由",
                        "核心证据材料", "申辩理由", "涉案法规", "依据摘录（台账原文）",
                        "挂网滞后", "存放路径", "案例文件"]:
                ok(key in dt, "详情含「%s」" % key)
            # 2026-11-04 用户明确：网页里不许出现 Excel 内部说法（"批注装不下"这类），
            # 它是给台账使用者看的，不是给维护者看的。整页扫一遍，不只是详情。
            body = pg.inner_text("body")
            bad_words = [w for w in ("批注", "悬停", "装不下", "超框", "csrc", "**")
                         if w in body]
            ok(not bad_words, "整页不出现技术词/未渲染符号", bad_words)

            # 法规标签 → 筛选
            ok(pg.eval_on_selector_all(".lawtag", "e=>e.length") > 0, "法规标签渲染")
            pg.click(".lawtag")
            pg.wait_for_timeout(460)
            badge = pg.evaluate("""() => { const b=document.querySelector('.fg[data-dim="laws"] .badge');
                return b.style.display==='none' ? 0 : +b.textContent; }""")
            ok(badge == 1, "点法规标签 → 该维度选中 1 项", badge)
            pg.click("#bClear")
            pg.wait_for_timeout(360)

            # 滚动到底
            pg.hover("#dscroll")
            for _ in range(30):
                pg.mouse.wheel(0, 900)
            pg.wait_for_timeout(420)
            sc = pg.evaluate("""() => { const e=document.getElementById('dscroll');
                return [e.scrollTop, e.scrollHeight - e.clientHeight]; }""")
            ok(sc[0] > 0, "右栏滚轮可滚", sc)
            ok(abs(sc[0] - sc[1]) < 8, "右栏能滚到最底", sc)

            # 最长条目（批注框装不下的那几条之一）
            longest = max(total, key=lambda d: len(d["text"]))["r"]
            pg.evaluate("(r)=>{location.hash='#L'+r}", longest)
            pg.reload()
            pg.wait_for_timeout(1100)
            ok(str(longest) in pg.inner_text("#dtitle"), "#L%d 直达最长条目" % longest,
               pg.inner_text("#dtitle")[:26])
            hh = pg.evaluate("""() => { const e=document.getElementById('dscroll');
                return [e.scrollHeight, e.clientHeight]; }""")
            ok(hh[0] > hh[1] * 1.5, "该条内容远高于可视区（确实需要滚）", hh)

            # 排序
            pg.select_option("#sort", "d-asc")
            pg.wait_for_timeout(420)
            d0 = pg.inner_text('.row[data-i="0"] .dt')
            pg.select_option("#sort", "d-desc")
            pg.wait_for_timeout(420)
            d1 = pg.inner_text('.row[data-i="0"] .dt')
            ok(d0 <= d1, "按处罚时间排序可切换", "%s / %s" % (d0, d1))
            pg.select_option("#sort", "r-asc")
            pg.wait_for_timeout(400)
            ok(pg.inner_text('.row[data-i="0"] .rno').strip().isdigit(), "排序后行号仍在且是 Excel 行号")

            # CSV
            with pg.expect_download() as dl:
                pg.click("#bExp")
            raw = io.open(dl.value.path(), "rb").read()
            ok(raw[:3] == b"\xef\xbb\xbf", "CSV 带 BOM（Excel 打开不乱码）")
            body = raw.decode("utf-8-sig").split("\r\n")
            ok(body[0].count(",") == 19, "CSV 20 列", body[0][:60])
            ok(len(body) - 1 == N, "CSV 行数 = 条目数", "%d/%d" % (len(body) - 1, N))

            # 键盘
            pg.click("body")
            pg.keyboard.press("ArrowDown")
            pg.wait_for_timeout(260)
            r1 = pg.inner_text("#dtitle")
            pg.keyboard.press("ArrowDown")
            pg.wait_for_timeout(260)
            ok(r1 != pg.inner_text("#dtitle"), "↓ 切换下一条")

            # 说明弹层
            pg.click("#bHelp")
            pg.wait_for_timeout(260)
            ok(pg.is_visible("#help .box"), "说明弹层可打开")
            pg.keyboard.press("Escape")
            pg.wait_for_timeout(220)
            ok(not pg.is_visible("#help .box"), "Esc 关闭说明")

            # 窄屏不塌
            pg.set_viewport_size({"width": 1366, "height": 768})
            pg.wait_for_timeout(500)
            ok(pg.eval_on_selector_all(".row", "e=>e.length") > 0, "1366×768 仍正常")
            box = pg.evaluate("""() => { const l=document.getElementById('list').getBoundingClientRect();
                const d=document.querySelector('.detail').getBoundingClientRect();
                return [l.width, d.width]; }""")
            ok(box[0] > 300 and box[1] > 400, "1366 下三栏未被挤塌", box)

            if a.shots:
                os.makedirs(a.shots, exist_ok=True)
                pg.screenshot(path=os.path.join(a.shots, "web_1366.png"))
                pg.set_viewport_size({"width": 1680, "height": 1050})
                pg.wait_for_timeout(400)
                pg.screenshot(path=os.path.join(a.shots, "web_1680.png"))

            ok(not errs, "全程无 JS 报错", errs[:3])
            b.close()
    finally:
        if not a.keep:
            shutil.rmtree(tmp, ignore_errors=True)

    print("\n".join(passes))
    if fails:
        print("\n".join(fails))
    print("\n%s  失败 %d / 通过 %d" % ("全部通过" if not fails else "有失败", len(fails), len(passes)))
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
