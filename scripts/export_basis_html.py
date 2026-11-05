# -*- coding: utf-8 -*-
"""把台账导成**单文件网页版台账**（检索台）：15 列全量 + 法条全文 + 多维筛选。

## 定位

这不是「处罚事由的备选展示方式」，而是台账的**第二个形态**：

  * Excel 台账：适合留存、盖章、交付、逐格校对；
  * 网页台账：适合**查**——按来源/违规类型/涉案法规/地区/措施/年度多条件交叉筛，
    搜任意关键词，一条条滚着读全文，查完直接导出 CSV 写报告。
    508 行在 Excel 里靠筛选箭头一格一格点，交叉条件一多就没法用；
    这里所有维度同时可见、都带条数，勾两个就知道还剩多少条。

## 为什么"能滚"这件事非要换载体

Excel 里「鼠标停在批注上滚轮看全文」是平台不支持的：
`commentsDrawing1.vml` 里每个批注的 `<ClientData ObjectType="Note">` 只有
`MoveWithCells / SizeWithCells / Anchor / AutoFill / Row / Column`（全文 `scroll` 出现 0 次），
`Comment.Shape` 也没有 `ScrollBars`；且 5 条最长案例的正文一屏装不下（最长需 1482px）。
浏览器里这两条限制同时消失。

## 数据口径（与台账严格一致，网页不另造事实）

| 网页字段 | 来源 | 说明 |
|---|---|---|
| 15 个内容列 | 台账 A~O 列 | 原样搬运，不改写、不补全 |
| 核心处罚依据（法条全文） | H 列批注正文 | 由 `annotate_basis.py` 从 06 目录法规原文生成 |
| 依据摘录（台账原文） | H 列单元格 | 保留，放元信息里做对照 |
| 处罚措施标签 | M 列文本关键词 | 只做分类标签，不改原文；原文全段照显 |
| 涉案法规 | H 列摘录里的《…》 | 走 `lawtext.ALIAS/norm_name` 归一，与写批注同一套映射 |
| 辖区 | 原文链接 csrc.gov.cn 一级路径 | 映射表 `ledger_style.BUREAU_CN`（唯一真源） |
| 挂网滞后 | 挂网日 − 落款日 | 仅展示，不作筛选（属观察项，不构成问题） |

**不给中基协/交易所编辖区**：它们的链接不在 csrc 下，辖区留空显示「—」。
宁可有空缺，也不按主体名猜注册地。

## 交互

    鼠标扫过列表      右侧立即换（不用点）
    右栏滚轮          可滚到底，看完整法条
    /                 聚焦搜索
    ↑ ↓               上一条 / 下一条
    1~9 或搜索框输行号  直达台账第 N 行（Excel 行号，与 H1 批注说明书一致）
    地址栏            `#L444` 直达；筛选条件写进 hash，可收藏/分享

## 用法

    python export_basis_html.py --root <库根>
    python export_basis_html.py --root <库根> --out <输出.html>

默认输出 **`<库根>/04_综合索引台账【总表】.html`**——与台账 xlsx 同名同目录。

2026-11-04 用户定：它是**台账的第二种形态**，不是附件，所以放库根并列展示。
这条改变了落点，连带两处必须同步，否则会被自己的规则清掉/报错：
  · `verify_library.py` C10 库根白名单要放行 `04_*.html`；
  · `cleanup_library.py` 的 `ROOT_KEEP_FILES` 要收它，否则 run_update 的清理阶段会把它搬走。
  · `annotate_basis.py` 写进 H1 表头批注的说明书里指向它，文案要与路径一致。

尺寸口径只有一处：调 `basis_comment.fit_size()`，与写批注同一套逻辑。
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import sys

from openpyxl import load_workbook

import basis_comment
import lawtext as LT
import ledger_style as LS

LEDGER = "04_综合索引台账【总表】.xlsx"
SHEET = "综合索引台账"
# 与台账正文同名、同放库根（2026-11-04 用户要求）：
# 用户在资源管理器里看到的是「同一份台账的两种形态」，不是两个分散的文件。
OUT_NAME = "04_综合索引台账【总表】.html"

_CSRC_SEG = re.compile(r"csrc\.gov\.cn/([a-z0-9]+)/")
_BOOK = re.compile(r"《([^》]{2,80})》")


# ---------------------------------------------------------------- 派生字段
# 注意：以下三条规则都只做「分类/索引」，不改台账任何原文。
# 网页里每一列显示的都是单元格原样，标签只用于筛选。

def measure_tags(m: str) -> list[str]:
    """处罚措施 → 标签。一条案例可命中多个（机构+人员合并在一格）。"""
    t = []
    if "警示函" in m:
        t.append("出具警示函")
    if "责令改正" in m:
        t.append("责令改正")
    if "诚信档案" in m or "诚信记录" in m:
        t.append("记入诚信档案")
    if "公开谴责" in m:
        t.append("公开谴责")
    if "警告" in m:                      # 「警示函」不含「警告」二字，不会串
        t.append("警告")
    if "暂停" in m and ("备案" in m or "受理" in m):
        t.append("暂停备案")
    if "撤销" in m or "注销" in m:
        t.append("撤销/注销登记")
    if "罚款" in m or "没收" in m:
        t.append("罚款/没收")
    if "监管警示" in m or "书面警示" in m:
        t.append("监管警示/书面警示")
    if "禁入" in m:
        t.append("市场禁入")
    if "黑名单" in m:
        t.append("加入黑名单")
    if "通报批评" in m or "监管谈话" in m or "约见" in m:
        t.append("通报批评/监管谈话")
    return t


def laws_of(basis: str) -> list[str]:
    """H 列摘录 → 涉案法规规范名列表（复用 lawtext 的 ALIAS，不另写一份映射）。

    简称（《条例》《暂行办法》）的消解沿用 lawtext 的口径：**用同一格里前面出现过的
    全称来定**——这是唯一可靠的办法，猜不出就不猜（返回 None，该项丢弃）。
    """
    out, fulls = [], []
    for raw in _BOOK.findall(basis or ""):
        n = LT.norm_name(raw)
        if n in LT.ALIAS:
            c = LT.ALIAS[n]
        elif n in LT.AMBIGUOUS:
            cand = [f for f in fulls if n in f]
            c = cand[-1] if cand else None
        else:
            c = n
        if not c:
            continue
        if c not in out:
            out.append(c)
        if c not in LT.ALIAS and c not in LT.AMBIGUOUS:
            fulls.append(c)
    return out


def region_of(link: str, src: str) -> str:
    """→ 辖区中文名；非证监局（中基协/交易所）留空，不按主体名猜。"""
    if src != "证监局":
        return ""
    for one in (link or "").split("；"):
        m = _CSRC_SEG.search(one)
        if m:
            return LS.BUREAU_CN.get(m.group(1), "")
    return ""


def _days(a: str, b: str):
    """b − a 的天数（都按 YYYY-MM-DD）；解析不了返回 None。"""
    try:
        da = _dt.datetime.strptime(a[:10], "%Y-%m-%d")
        db = _dt.datetime.strptime(b[:10], "%Y-%m-%d")
    except (ValueError, TypeError):
        return None
    return (db - da).days


def _s(v) -> str:
    """单元格值 → 干净字符串（日期统一 YYYY-MM-DD）。"""
    if v is None:
        return ""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def collect(root: str) -> tuple[list[dict], int]:
    """→ (条目列表, 超框条数)。条目按台账行序。"""
    path = os.path.join(root, LEDGER)
    if not os.path.exists(path):
        raise SystemExit("找不到台账：%s" % path)
    wb = load_workbook(path)
    if SHEET not in wb.sheetnames:
        raise SystemExit("台账里没有工作表 %r" % SHEET)
    ws = wb[SHEET]

    items, n_over = [], 0
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=8)          # H 列
        cm = cell.comment
        if cm is None:
            continue
        text = cm.text or ""
        if not text.strip():
            continue
        _w, h = basis_comment.fit_size(text)
        over = h >= basis_comment.BOX_MAX_H      # 与写批注时同一个判据
        n_over += bool(over)
        g = lambda c: _s(ws.cell(row=r, column=c).value)   # noqa: E731
        rec = {
            "r": r,
            "no": g(1),      # 序号
            "d": g(2),       # 处罚时间
            "p": g(3),       # 挂网发布日期
            "s": g(4),       # 被处罚机构/人
            "m": g(5),       # 处罚措施（机构+人员）
            "g": g(6),       # 核心处罚事由（F）
            "e": g(7),       # 核心证据材料（G）
            "h": g(8),       # 核心处罚依据（H，台账摘录）
            "defn": g(9),    # 申辩理由（I）
            "k": g(10),      # 违规类型
            "src": g(11),    # 机构来源
            "file": g(12),   # 案例文件
            "path": g(13),   # 存放路径
            "link": g(14),   # 原文链接
            "v": g(15),      # 是否典型案例
            "text": text,    # 法条全文（H 列批注）—— 网页"核心处罚依据"显示的就是它
            "over": over,
        }
        # —— 派生索引字段（只用于筛选，不改上面任何一列的内容）——
        rec["acts"] = measure_tags(rec["m"])
        rec["laws"] = laws_of(rec["h"])
        rec["reg"] = region_of(rec["link"], rec["src"])
        lag = _days(rec["d"], rec["p"])
        if lag is not None:
            rec["lag"] = lag
        items.append(rec)
    wb.close()
    return items, n_over


TPL = r"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>私募监管处罚案例库 · 台账检索台</title>
<style>
:root{
  --fg:#1f2328; --dim:#69707a; --line:#e3e7ec; --bg:#fff; --soft:#f7f8fa;
  --accent:#8a5a00; --warn:#b42318; --sel:#fff9e8; --selb:#e8c766; --blue:#0b5cad;
  --h1:50px; --h2:40px; --fs:14.5px;
}
*{box-sizing:border-box}
html,body{height:100%}
body{margin:0;background:var(--bg);color:var(--fg);overflow:hidden;
 font:var(--fs)/1.75 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
button{font-family:inherit}
/* ---------- 顶栏 ---------- */
header{height:var(--h1);display:flex;gap:12px;align-items:center;padding:0 16px;
 border-bottom:1px solid var(--line);background:#fff;position:relative;z-index:9}
h1{font-size:14px;margin:0;font-weight:600;white-space:nowrap;letter-spacing:.2px}
h1 span{font-weight:400;color:#9aa4b2;font-size:12px;margin-left:7px}
.sw{flex:1 1 auto;min-width:160px;position:relative;display:flex}
#q{flex:1 1 auto;padding:7px 30px 7px 11px;border:1px solid var(--line);border-radius:7px;
 font-size:13px;outline:none;background:#fcfcfd;font-family:inherit;color:var(--fg)}
#q:focus{border-color:#9aa4b2;box-shadow:0 0 0 3px rgba(154,164,178,.15);background:#fff}
#qx{position:absolute;right:4px;top:50%;transform:translateY(-50%);border:0;background:none;
 color:#9aa4b2;cursor:pointer;font-size:15px;padding:2px 6px;display:none}
#q:not(:placeholder-shown)+#qx{display:block}
.tools{display:flex;gap:6px;align-items:center}
.btn{border:1px solid var(--line);background:#fff;color:var(--dim);border-radius:6px;
 padding:5px 9px;font-size:12px;cursor:pointer;white-space:nowrap}
.btn:hover{border-color:#c4cbd4;color:var(--fg);background:#fcfcfd}
.btn.on{background:var(--sel);border-color:var(--selb);color:var(--accent)}
/* ---------- 来源条 ---------- */
.srcbar{height:var(--h2);display:flex;align-items:center;gap:6px;padding:0 16px;
 border-bottom:1px solid var(--line);background:var(--soft);overflow-x:auto;font-size:12px}
.srcbar .lbl{color:#9aa4b2;margin-right:2px;white-space:nowrap}
.seg{border:1px solid transparent;background:none;color:var(--dim);border-radius:5px;
 padding:3px 8px;cursor:pointer;white-space:nowrap;font-size:12px}
.seg:hover{background:#eef1f5}
.seg.on{background:#fff;border-color:var(--selb);color:var(--accent);font-weight:600;
 box-shadow:0 1px 2px rgba(16,24,40,.06)}
.seg b{font-weight:600;color:#9aa4b2;margin-left:3px;font-variant-numeric:tabular-nums}
.seg.on b{color:var(--accent)}
#stat{margin-left:auto;color:var(--dim);white-space:nowrap;font-variant-numeric:tabular-nums}
/* ---------- 三栏 ---------- */
.app{display:flex;height:calc(100vh - var(--h1) - var(--h2))}
/* 筛选栏 */
#filters{flex:0 0 244px;overflow-y:auto;border-right:1px solid var(--line);background:var(--soft);
 padding-bottom:40px}
#filters::-webkit-scrollbar,.list::-webkit-scrollbar,.dscroll::-webkit-scrollbar{width:11px;height:11px}
#filters::-webkit-scrollbar-track,.list::-webkit-scrollbar-track,.dscroll::-webkit-scrollbar-track{background:#f2f4f7}
#filters::-webkit-scrollbar-thumb,.list::-webkit-scrollbar-thumb,.dscroll::-webkit-scrollbar-thumb{
 background:#c3cbd6;border-radius:6px;border:3px solid #f2f4f7}
#filters::-webkit-scrollbar-thumb:hover,.list::-webkit-scrollbar-thumb:hover,
.dscroll::-webkit-scrollbar-thumb:hover{background:#96a1b0}
.fg{border-bottom:1px solid #e9edf2}
.fgh{width:100%;display:flex;align-items:center;gap:6px;padding:8px 12px;border:0;background:none;
 cursor:pointer;font-size:12.5px;font-weight:600;color:#38404a;text-align:left}
.fgh:hover{background:#eef1f5}
.fgh .caret{margin-left:auto;color:#b3bcc7;font-size:10px;transition:transform .12s}
.fg.closed .caret{transform:rotate(-90deg)}
.fgh .badge{font-size:10.5px;color:#fff;background:var(--accent);border-radius:8px;
 padding:0 5px;line-height:1.6;font-weight:600}
.fg.closed .fgb{display:none}
.fgb{padding:0 8px 8px}
.fsearch{width:100%;border:1px solid var(--line);border-radius:5px;padding:3px 7px;font-size:11.5px;
 margin-bottom:4px;outline:none;font-family:inherit;background:#fff}
.opt{display:flex;align-items:center;gap:6px;padding:3px 6px;border-radius:5px;cursor:pointer;
 font-size:12px;line-height:1.5;position:relative;overflow:hidden;color:#3a424c}
.opt:hover{background:#eef1f5}
.opt input{position:relative;z-index:2;margin:0;flex:0 0 auto;accent-color:var(--accent)}
.opt .nm{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
 position:relative;z-index:1}
.opt .nu{flex:0 0 auto;color:#9aa4b2;font-size:11px;font-variant-numeric:tabular-nums;
 position:relative;z-index:1}
.opt .bar{position:absolute;left:0;top:0;bottom:0;background:#e8ebf0;border-radius:5px;z-index:0}
.opt input:checked ~ .bar{background:#f7e6b8}
.opt.on .nm{font-weight:600;color:var(--accent)}
.opt.zero{opacity:.38}
.more{width:100%;border:0;background:none;color:var(--blue);font-size:11.5px;cursor:pointer;
 padding:4px;text-align:left}
.more:hover{text-decoration:underline}
/* 中栏 */
.mid{flex:0 1 36%;min-width:320px;max-width:640px;display:flex;flex-direction:column;
 border-right:1px solid var(--line)}
.sortbar{flex:0 0 auto;display:flex;align-items:center;gap:8px;padding:6px 12px;
 border-bottom:1px solid var(--line);font-size:11.5px;color:var(--dim);background:#fff}
.sortbar select{border:1px solid var(--line);border-radius:5px;padding:2px 4px;font-size:11.5px;
 background:#fff;color:#38404a;font-family:inherit;outline:none}
.sortbar .k{color:#9aa4b2}
.sortbar .sp{margin-left:auto}
.list{flex:1 1 auto;overflow-y:auto;background:var(--soft)}
.row{padding:7px 12px 8px 10px;cursor:pointer;border-bottom:1px solid #edf0f4;
 border-left:3px solid transparent}
.row:hover{background:#fff}
.row.on{background:var(--sel);border-left-color:var(--selb)}
.r1{display:flex;align-items:baseline;gap:8px}
.rno{flex:0 0 34px;text-align:right;color:#aab3be;font-size:11.5px;
 font-variant-numeric:tabular-nums}
.row.on .rno{color:var(--accent);font-weight:700}
.dt{flex:0 0 auto;color:#8a929c;font-size:11.5px;font-variant-numeric:tabular-nums}
.su{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
 font-size:13px;color:#2b323b}
.row.on .su{font-weight:600;color:#1a1f26}
.star{flex:0 0 auto;color:#c99a12;font-size:12px}
.r2{display:flex;align-items:center;gap:5px;margin-left:42px;margin-top:3px;font-size:10.5px;
 line-height:1.7;color:#9aa4b2;overflow:hidden}
.pill{flex:0 0 auto;border-radius:3px;padding:0 4px;line-height:1.65;border:1px solid currentColor}
.pill.k1{color:#8a5a00;background:#fdf6e3}
.pill.k2{color:#1f6f8b;background:#eaf4f8}
.pill.k3{color:#7a4a9c;background:#f3ecf8}
.pill.k4{color:#2f6b3f;background:#e9f4ec}
.pill.k5{color:#a34d17;background:#fbeee4}
.pill.k6{color:#5a6672;background:#eef1f5}
.pill.k7{color:#8a3b52;background:#fbeef1}
.pill.k8{color:#3f5a8a;background:#ecf0f8}
.pill.k9{color:#5d6b2f;background:#f2f5e6}
.pill.src{color:#5a6672;background:#eef1f5}
.r2 .pill{flex:0 1 auto;min-width:0;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pill.tj{color:var(--warn);background:#fdeceb}
.ell{flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
/* 右栏 */
.detail{flex:1 1 auto;min-width:0;display:flex;flex-direction:column;background:#fff}
.dhead{flex:0 0 auto;display:flex;gap:8px;align-items:center;padding:7px 18px;
 border-bottom:1px solid var(--line);background:#fff}
.dhead .t{font-weight:600;font-size:13px;overflow:hidden;text-overflow:ellipsis;
 white-space:nowrap;flex:1 1 auto;min-width:0}
.dhead .pos{flex:0 0 auto;font-size:11px;color:#9aa4b2;font-variant-numeric:tabular-nums}
.dhead .nav{flex:0 0 auto;display:flex;gap:4px}
.dscroll{flex:1 1 auto;overflow-y:auto}
.dinner{max-width:960px;margin:0 auto;padding:18px 26px 110px}
.meta{display:grid;grid-template-columns:auto 1fr;gap:3px 14px;font-size:12.5px;color:var(--dim);
 margin:0 0 6px;background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:11px 14px}
.meta dt{margin:0;white-space:nowrap;color:#8a929c}
.meta dd{margin:0;color:#38404a;word-break:break-word}
h2{font-size:12.5px;font-weight:600;color:var(--accent);letter-spacing:.5px;margin:22px 0 8px;
 padding-left:9px;border-left:3px solid var(--selb);display:flex;align-items:center;gap:8px}
h2 .h2c{margin-left:auto;font-weight:400;color:#8a929c;font-size:10.5px;letter-spacing:0}
.txt{white-space:pre-wrap;word-break:break-word;margin:0;font-family:inherit;font-size:inherit}
pre.txt{font-size:calc(var(--fs) - .5px);line-height:1.92}
.tagrow{display:flex;flex-wrap:wrap;gap:6px;margin:0}
.lawtag{font-size:11.5px;border:1px solid #d8dfe8;border-radius:12px;padding:1px 9px;
 background:#fbfcfd;color:#3f4a57;cursor:pointer}
.lawtag:hover{border-color:var(--selb);background:var(--sel);color:var(--accent)}
.rel{display:flex;flex-direction:column;gap:1px;font-size:12px}
.rel a{display:flex;gap:8px;padding:3px 6px;border-radius:5px;cursor:pointer;color:#38404a;
 text-decoration:none;border:0}
.rel a:hover{background:var(--soft);color:var(--accent)}
.rel .rn{color:#aab3be;font-variant-numeric:tabular-nums;flex:0 0 34px;text-align:right}
.rel .rd{color:#8a929c;flex:0 0 auto;font-variant-numeric:tabular-nums}
.rel .rs{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hint{color:var(--dim);font-size:12px;margin:8px 0 0}
.empty{color:#9aa4b2;padding:44px 8px;text-align:center;line-height:2}
a{color:var(--blue);text-decoration:none;border-bottom:1px solid #cfe0f2}
a:hover{border-bottom-color:var(--blue)}
mark{background:#ffe9a3;color:inherit;border-radius:2px;padding:0 1px}
/* 说明弹层 */
#help{position:fixed;inset:0;background:rgba(20,24,30,.38);display:none;z-index:99;
 align-items:center;justify-content:center}
#help.on{display:flex}
#help .box{background:#fff;border-radius:12px;max-width:720px;width:calc(100% - 60px);
 max-height:82vh;overflow:auto;padding:22px 26px;box-shadow:0 18px 50px rgba(16,24,40,.22)}
#help h3{margin:0 0 12px;font-size:15px}
#help p,#help li{font-size:12.8px;line-height:1.9;color:#3a424c;margin:6px 0}
#help code{background:var(--soft);border:1px solid var(--line);border-radius:4px;
 padding:0 4px;font-size:11.8px}
#help table{border-collapse:collapse;font-size:12.3px;margin:6px 0 10px}
#help td{border:1px solid var(--line);padding:3px 8px;color:#3a424c}
#help td:first-child{color:#8a929c;white-space:nowrap}
#help .cl{margin-top:14px;text-align:right}
@media (max-width:1180px){ #filters{flex-basis:210px} .mid{flex-basis:40%} }
@media (max-width:980px){
  .app{flex-direction:column;height:auto;overflow:auto}
  body{overflow:auto}
  #filters{flex:0 0 auto;max-height:none;border-right:0;border-bottom:1px solid var(--line)}
  .mid{flex:0 0 auto;max-width:none;border-right:0;border-bottom:1px solid var(--line)}
  .list{max-height:46vh}.dscroll{max-height:none}
  .detail{min-height:60vh}
}
@media print{
  header,.srcbar,#filters,.mid,#help{display:none}
  body{overflow:visible}.app{display:block;height:auto}
  .dscroll{overflow:visible;max-height:none}.dinner{padding:0;max-width:none}
}
</style></head><body>
<header>
  <h1>私募监管处罚案例库<span>台账检索台</span></h1>
  <div class="sw">
    <input id="q" type="search" placeholder="搜机构 / 法规 / 条款 / 关键词（空格＝并且，支持 来源:中基协 这类写法）">
    <button id="qx" title="清空搜索">×</button>
  </div>
  <div class="tools">
    <button class="btn" id="bExp" title="把当前筛选结果导成 CSV（Excel 可直接打开）">导出 CSV</button>
    <button class="btn" id="bFs1" title="缩小正文">A−</button>
    <button class="btn" id="bFs2" title="放大正文">A+</button>
    <button class="btn" id="bHelp" title="口径、用法、快捷键">说明</button>
  </div>
</header>
<div class="srcbar" id="srcbar"><span class="lbl">来源</span><span id="stat"></span></div>
<div class="app">
  <aside id="filters"></aside>
  <section class="mid">
    <div class="sortbar">
      <span class="k">排序</span>
      <select id="sort">
        <option value="r-asc">台账行号（新在上）</option>
        <option value="d-desc">处罚时间 新→旧</option>
        <option value="d-asc">处罚时间 旧→新</option>
        <option value="p-desc">挂网日期 新→旧</option>
        <option value="s-asc">主体名称</option>
      </select>
      <span class="sp" id="midinfo"></span>
      <button class="btn" id="bClear" title="清空全部筛选条件">清空条件</button>
    </div>
    <div class="list" id="list"></div>
  </section>
  <section class="detail">
    <div class="dhead">
      <span class="t" id="dtitle">把鼠标移到左边任一行</span>
      <span class="pos" id="dpos"></span>
      <span class="nav">
        <button class="btn" id="bPrev" title="上一条（↑）">↑</button>
        <button class="btn" id="bNext" title="下一条（↓）">↓</button>
        <button class="btn" id="bCopy" title="复制整条为 Markdown，可直接贴进报告">复制</button>
        <button class="btn" id="bPrint" title="只打印当前这一条">打印</button>
      </span>
    </div>
    <div class="dscroll" id="dscroll">
      <div class="dinner" id="dinner">
        <p class="empty">左边是台账里的全部案例。<br>
        鼠标扫过任一行，这里立刻显示该条的完整内容与法条全文。<br>
        左栏可按来源、违规类型、涉案法规、地区、处罚措施、年度交叉筛选。</p>
      </div>
    </div>
  </section>
</div>
<div id="help"><div class="box">
  <h3>这是一份什么文件</h3>
  <p>它是 <code>04_综合索引台账【总表】.xlsx</code> 的网页形态：15 列内容与台账逐列一致，
  并且把「核心处罚依据」对应的法条全文展开在右侧，可以直接读到条文末尾。</p>
  <p>所有正文都来自台账本身和 06 目录的法规原文，不新增、不改写、不补全。
  筛选用的标签（处罚措施、涉案法规、地区）只是索引，不改动任何原文。</p>
  <h3>怎么用</h3>
  <ul>
    <li><b>筛</b>：左栏勾选，维度之间是「并且」，维度内是「或者」。每项后面的数字是<b>其余条件不变时</b>该项还能筛出多少条，勾之前就知道会剩几条。</li>
    <li><b>搜</b>：搜索框空格分词＝全部命中；也可以用字段限定：<code>来源:中基协</code>、<code>类型:募集</code>、<code>法规:适当性</code>、<code>地区:深圳</code>、<code>措施:公开谴责</code>、<code>年:2025</code>。</li>
    <li><b>读</b>：鼠标扫过列表即切换；右侧滚轮可读到底。点击「涉案法规」里的名称＝按该法规筛选。</li>
    <li><b>取</b>：<code>复制</code> 得到整条 Markdown，可直接贴进报告；<code>导出 CSV</code> 导出当前筛选结果，Excel 打开不乱码。</li>
    <li><b>定位</b>：地址栏 <code>#L444</code> 直达台账第 444 行；筛选条件也会写进地址栏，可收藏或转发。</li>
  </ul>
  <h3>快捷键</h3>
  <table>
    <tr><td><code>/</code></td><td>聚焦搜索框</td></tr>
    <tr><td><code>↑</code> <code>↓</code></td><td>上一条 / 下一条</td></tr>
    <tr><td><code>Esc</code></td><td>先清搜索，再按一次清全部筛选</td></tr>
    <tr><td><code>A−</code> <code>A+</code></td><td>正文大小（法条全文随之缩放）</td></tr>
  </table>
  <h3>数据口径</h3>
  <ul>
    <li>收录区间：落款日 2025-10-31 ~ 2026-10-30，超出窗口的旧案例已移出库。</li>
    <li>「处罚措施」按 M 列文字归类，一条可带多个标签；措施原文整段照显，以原文为准。</li>
    <li>「涉案法规」按依据摘录里的《…》归类，简称归到对应全称。</li>
    <li>「地区」取自证监局网站的辖区，中基协与交易所的案例没有这一项，留空不显示，也不按机构名称推测注册地。</li>
    <li>「挂网滞后」＝挂网日 − 落款日，各地刊发节奏不同，可用来看公告时效。</li>
  </ul>
  <div class="cl"><button class="btn" id="bHelpClose">关闭</button></div>
</div></div>
<script>
const DATA = __DATA__;
const $ = id => document.getElementById(id);
const esc = s => (s==null?'':String(s)).replace(/[&<>"]/g,
  c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

/* ============ 维度定义 ============ */
const kName = k => { const i = String(k).indexOf('_'); return i<0 ? String(k) : String(k).slice(i+1); };
const kNo   = k => { const i = String(k).indexOf('_'); return i<0 ? '' : String(k).slice(0,i); };
const DIMS = {
  src : {label:'机构来源',        vals: d => d.src ? [d.src] : []},
  k   : {label:'违规类型',        vals: d => d.k ? [d.k] : []},
  laws: {label:'涉案法规',        vals: d => (d.laws && d.laws.length) ? d.laws : []},
  acts: {label:'处罚措施',        vals: d => (d.acts && d.acts.length) ? d.acts : []},
  reg : {label:'地区',            vals: d => d.reg ? [d.reg] : []},
  yr  : {label:'处罚年度',        vals: d => d.d ? [d.d.slice(0,4)] : []},
  v   : {label:'是否典型案例',    vals: d => [d.v === '是' ? '是' : '否']},
  ind : {label:'是否涉及个人责任', vals: d => [/高管（|从业人员/.test(d.m||'') ? '含个人' : '仅机构']},
};
const DIM_ORDER = ['k','laws','acts','reg','yr','v','ind'];
const DIM_LABEL = {src:'机构来源',k:'违规类型',laws:'涉案法规',acts:'处罚措施',
                   reg:'地区',yr:'年度',v:'典型案例',ind:'个人责任'};
const SEL = {};
for (const d in DIMS) SEL[d] = new Set();

/* 预计算搜索串 */
DATA.forEach(d => {
  d._k = kName(d.k); d._kno = kNo(d.k);
  d._h = [d.r, d.no, d.d, d.p, d.s, d.m, d.g, d.e, d.h, d.defn, d.k, d._k,
          d.src, d.reg || '', d.file, d.path, d.v,
          (d.acts||[]).join(' '), (d.laws||[]).join(' ')].join(' ').toLowerCase();
});

/* 维度选项全集（频次降序，固定顺序） */
const ALL = {};
function tally(fn){
  const m = {};
  DATA.forEach(d => fn(d).forEach(v => { m[v] = (m[v]||0) + 1; }));
  return m;
}
ALL.k    = tally(d => DIMS.k.vals(d));
ALL.laws = tally(d => DIMS.laws.vals(d));
ALL.acts = tally(d => DIMS.acts.vals(d));
ALL.reg  = tally(d => DIMS.reg.vals(d));
ALL.yr   = tally(d => DIMS.yr.vals(d));
ALL.src  = tally(d => DIMS.src.vals(d));
ALL.v    = tally(d => DIMS.v.vals(d));
ALL.ind  = tally(d => DIMS.ind.vals(d));
const byFreq = m => Object.keys(m).sort((a,b) => (m[b]-m[a]) || a.localeCompare(b,'zh'));
const OPTIONS = {
  k   : byFreq(ALL.k),
  laws: byFreq(ALL.laws),
  acts: byFreq(ALL.acts),
  reg : byFreq(ALL.reg),
  yr  : Object.keys(ALL.yr).sort().reverse(),
  v   : ['是','否'],
  ind : ['含个人','仅机构'],
};

/* ============ 查询 ============ */
const ALIAS = {'来源':'src','src':'src','机构来源':'src',
  '类型':'k','分类':'k','k':'k','违规类型':'k',
  '法规':'laws','法条':'laws','law':'laws','laws':'laws',
  '措施':'acts','acts':'acts','处罚措施':'acts',
  '辖区':'reg','地区':'reg','reg':'reg',
  '年':'yr','年度':'yr','year':'yr','yr':'yr',
  '典型':'v','v':'v'};
function parseQ(text){
  const plain = [], dims = {};
  String(text||'').trim().split(/\s+/).filter(Boolean).forEach(t => {
    const i = t.search(/[:：]/);
    if (i > 0){
      const dim = ALIAS[t.slice(0,i)];
      const val = t.slice(i+1).trim().toLowerCase();
      if (dim && val){ (dims[dim] = dims[dim] || []).push(val); return; }
    }
    plain.push(t.toLowerCase());
  });
  return {plain, dims};
}
/* skip：这个维度不参与（这样 facet 数字回答的是「勾了它会剩几条」） */
function pass(d, q, skip){
  if (q.plain.length && !q.plain.every(k => d._h.includes(k))) return false;
  for (const dim in q.dims){
    if (dim === skip) continue;
    const vals = (DIMS[dim] ? DIMS[dim].vals(d) : []).map(v => v.toLowerCase());
    if (!q.dims[dim].every(need => vals.some(v => v.includes(need)))) return false;
  }
  for (const dim in SEL){
    if (dim === skip) continue;
    const s = SEL[dim];
    if (s.size && !DIMS[dim].vals(d).some(v => s.has(v))) return false;
  }
  return true;
}
const currentQ = () => parseQ(qEl.value);

/* ============ 状态 ============ */
let view = [], selIdx = -1, curR = -1, sortKey = 'r-asc';
const qEl = $('q'), listEl = $('list'), midEl = $('midinfo'), dInner = $('dinner');

/* ============ 顶部来源条 ============ */
function buildSrcbar(){
  const kinds = [''].concat(byFreq(ALL.src));
  $('srcbar').innerHTML = '<span class="lbl">来源</span>' + kinds.map(v =>
      `<button class="seg${v===''?' on':''}" data-src="${esc(v)}">${v===''?'全部':esc(v)} <b>0</b></button>`
    ).join('') + '<span id="stat"></span>';
  $('srcbar').querySelectorAll('.seg[data-src]').forEach(b => b.addEventListener('click', () => {
    SEL.src.clear();
    if (b.dataset.src) SEL.src.add(b.dataset.src);
    apply();
  }));
}

/* ============ 左栏筛选 ============ */
const MAXSHOW = {k:99, laws:10, acts:99, reg:12, yr:99, v:99, ind:99};
const expanded = {laws:false, reg:false};
function optLabel(dim, v){
  return dim === 'k'
    ? `<span style="color:#8a929c">${esc(kNo(v))}</span> ${esc(kName(v))}`
    : esc(v);
}
function buildFilters(){
  const host = $('filters');
  host.innerHTML = DIM_ORDER.map(dim => {
    const opts = OPTIONS[dim];
    const cap = expanded[dim] ? opts.length : Math.min(opts.length, MAXSHOW[dim]);
    const items = opts.slice(0, cap).map(v =>
      `<label class="opt" data-dim="${dim}" data-v="${esc(v)}">
         <input type="checkbox" value="${esc(v)}">
         <span class="bar"></span>
         <span class="nm" title="${esc(v)}">${optLabel(dim,v)}</span>
         <span class="nu">0</span>
       </label>`).join('');
    const more = opts.length > MAXSHOW[dim]
      ? `<button class="more" data-dim="${dim}">${expanded[dim] ? '收起' : '展开全部 '+opts.length+' 项'}</button>` : '';
    return `<div class="fg" data-dim="${dim}">
      <button class="fgh">${DIMS[dim].label}<span class="badge"></span><span class="caret">▾</span></button>
      <div class="fgb">
        ${opts.length > 14 ? `<input class="fsearch" data-dim="${dim}" placeholder="在这个维度里找…">` : ''}
        <div class="opts">${items}</div>${more}
      </div></div>`;
  }).join('');

  host.querySelectorAll('.opt input').forEach(inp => {
    inp.addEventListener('change', e => {
      const lab = e.target.closest('.opt'), dim = lab.dataset.dim, v = lab.dataset.v;
      if (e.target.checked) SEL[dim].add(v); else SEL[dim].delete(v);
      lab.classList.toggle('on', e.target.checked);
      apply();
    });
  });
  host.querySelectorAll('.fgh').forEach(b =>
    b.addEventListener('click', () => b.parentElement.classList.toggle('closed')));
  host.querySelectorAll('.more').forEach(b =>
    b.addEventListener('click', () => {
      expanded[b.dataset.dim] = !expanded[b.dataset.dim];
      buildFilters(); syncChecks(); apply();
    }));
  host.querySelectorAll('.fsearch').forEach(inp =>
    inp.addEventListener('input', () => {
      const kw = inp.value.trim().toLowerCase(), wrap = inp.parentElement;
      wrap.querySelectorAll('.opt').forEach(el => {
        el.style.display = (!kw || el.dataset.v.toLowerCase().includes(kw)) ? '' : 'none';
      });
      const m = wrap.querySelector('.more'); if (m) m.style.display = kw ? 'none' : '';
    }));
  DIM_ORDER.forEach((d,i) => {
    if (i >= 3) host.querySelector(`.fg[data-dim="${d}"]`).classList.add('closed');
  });
  syncChecks();
}
function syncChecks(){
  DIM_ORDER.forEach(dim => {
    const box = $('filters').querySelector(`.fg[data-dim="${dim}"]`);
    if (!box) return;
    box.querySelectorAll('.opt').forEach(o => {
      const on = SEL[dim].has(o.dataset.v);
      o.querySelector('input').checked = on;
      o.classList.toggle('on', on);
    });
    if (SEL[dim].size) box.classList.remove('closed');
  });
  $('srcbar').querySelectorAll('.seg[data-src]').forEach(b =>
    b.classList.toggle('on', (b.dataset.src || '') === ([...SEL.src][0] || '')));
}

/* ============ facet 计数 ============ */
function updateFacets(q){
  const cache = {};
  DIM_ORDER.forEach(dim => {
    const m = {};
    DATA.forEach(d => {
      if (!pass(d, q, dim)) return;
      DIMS[dim].vals(d).forEach(v => { m[v] = (m[v]||0) + 1; });
    });
    cache[dim] = m;
  });
  DIM_ORDER.forEach(dim => {
    const m = cache[dim];
    const box = $('filters').querySelector(`.fg[data-dim="${dim}"]`);
    if (!box) return;
    let max = 1;
    for (const v in m) if (m[v] > max) max = m[v];
    box.querySelectorAll('.opt').forEach(el => {
      const v = el.dataset.v, c = m[v] || 0;
      el.querySelector('.nu').textContent = c;
      el.querySelector('.bar').style.width = (c / max * 100).toFixed(1) + '%';
      el.classList.toggle('zero', c === 0 && !SEL[dim].has(v));
    });
    const n = SEL[dim].size, badge = box.querySelector('.badge');
    badge.style.display = n ? '' : 'none';
    badge.textContent = n;
  });
  const srcCnt = {}; let srcAll = 0;
  DATA.forEach(d => {
    if (!pass(d, q, 'src')) return;
    srcAll++;
    srcCnt[d.src] = (srcCnt[d.src] || 0) + 1;
  });
  $('srcbar').querySelectorAll('.seg').forEach(b => {
    const v = b.dataset.src;
    b.querySelector('b').textContent = v === '' ? srcAll : (srcCnt[v] || 0);
  });
  // 是否处于「筛过」的状态：只看结果数，不依赖条件是勾选来的还是搜索来的
  const filtered = view.length !== DATA.length;
  $('stat').textContent = `共 ${DATA.length} 条` + (filtered ? ` · 当前 ${view.length} 条` : '');
  midEl.textContent = filtered ? `筛出 ${view.length} 条` : '';
}

/* ============ 列表 ============ */
const SORTERS = {
  'r-asc' : (a,b) => a.r - b.r,
  'd-desc': (a,b) => String(b.d||'').localeCompare(String(a.d||'')) || a.r - b.r,
  'd-asc' : (a,b) => String(a.d||'').localeCompare(String(b.d||'')) || a.r - b.r,
  'p-desc': (a,b) => String(b.p||'').localeCompare(String(a.p||'')) || a.r - b.r,
  's-asc' : (a,b) => String(a.s||'').localeCompare(String(b.s||''),'zh') || a.r - b.r,
};
function renderList(){
  const q = currentQ();
  view = DATA.filter(d => pass(d, q));
  view.sort(SORTERS[sortKey] || SORTERS['r-asc']);
  listEl.innerHTML = view.length ? view.map((d,i) => {
    const acts = d.acts || [];
    return `<div class="row" data-i="${i}" id="L${d.r}">`+
      `<div class="r1"><span class="rno">${d.r}</span><span class="dt">${esc(d.d)}</span>`+
      `<span class="su" title="${esc(d.s)}">${esc(d.s) || '（无主体）'}</span>`+
      (d.v === '是' ? '<span class="star" title="典型案例">★</span>' : '') + `</div>`+
      `<div class="r2"><span class="pill k${parseInt(d._kno,10)||9}">${esc(d._kno)} ${esc(d._k)}</span>`+
      `<span class="pill src">${esc(d.src)}${d.reg ? ' · ' + esc(d.reg) : ''}</span>`+
      `<span class="ell">${esc(acts[0]||'')}${acts.length>1?' +'+(acts.length-1):''}</span>`+
      `</div></div>`;
  }).join('') : '<p class="empty">没有匹配的条目。<br>点上方「清空条件」或放宽筛选。</p>';

  selIdx = -1;
  updateFacets(q);
  if (!view.length){
    curR = -1;
    $('dtitle').textContent = '没有匹配的条目';
    dInner.innerHTML = '<p class="empty">当前条件筛不出案例。<br>左栏每项后面的数字会告诉你哪一项归零了。</p>';
    saveHash(); return;
  }
  const keep = view.findIndex(d => d.r === curR);
  show(keep >= 0 ? keep : 0, keep < 0);
  saveHash();
}
function show(i, scroll){
  if (!view.length) return;
  i = Math.max(0, Math.min(view.length - 1, i));
  if (selIdx >= 0){
    const prev = listEl.querySelector(`.row[data-i="${selIdx}"]`);
    if (prev) prev.classList.remove('on');
  }
  selIdx = i;
  const d = view[i];
  curR = d.r;
  const el = listEl.querySelector(`.row[data-i="${i}"]`);
  if (el){
    el.classList.add('on');
    if (scroll !== false) el.scrollIntoView({block:'nearest'});
  }
  $('dpos').textContent = `${i + 1} / ${view.length}`;
  renderDetail(d);
  saveHash();
}

/* ============ 详情 ============ */
const metaRow = (k,v) => v ? `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>` : '';
function sect(title, body){
  return body
    ? `<h2>${esc(title)}</h2><p class="txt">${esc(body)}</p>`
    : `<h2>${esc(title)}</h2><p class="txt" style="color:#9aa4b2">（空）</p>`;
}
function relOf(d){
  const core = String(d.s||'').split('；')[0].trim();
  if (!core) return '';
  const same = DATA.filter(x => x.r !== d.r && String(x.s||'').includes(core));
  if (!same.length) return '';
  return `<h2>同一主体的其他记录<span class="h2c">共 ${same.length} 条</span></h2><div class="rel">` +
    same.slice(0,12).map(x => {
      const a = x.acts || [];
      return `<a data-goto="${x.r}"><span class="rn">${x.r}</span><span class="rd">${esc(x.d)}</span>`+
        `<span class="rs">${esc(a[0]||'')}${a.length>1?' 等'+a.length+' 项':''}</span></a>`;
    }).join('') + '</div>';
}
function renderDetail(d){
  $('dtitle').textContent = `第 ${d.r} 行 · ${d.s || '（无主体）'}`;
  let h = '<dl class="meta">';
  h += metaRow('序号', d.no) + metaRow('处罚时间', d.d) + metaRow('挂网发布日期', d.p);
  h += metaRow('被处罚机构/人', d.s) + metaRow('违规类型', d.k) + metaRow('机构来源', d.src);
  h += metaRow('地区', d.reg);
  if (d.v && d.v !== '否') h += metaRow('是否典型案例', d.v);
  if (typeof d.lag === 'number') h += metaRow('挂网滞后', d.lag + ' 天');
  h += metaRow('案例文件', d.file) + metaRow('存放路径', d.path);
  h += metaRow('依据摘录（台账原文）', d.h);
  h += '</dl>';

  const acts = d.acts || [];
  h += `<h2>处罚措施（机构+人员）${acts.length ? `<span class="h2c">${esc(acts.join(' · '))}</span>` : ''}</h2>`;
  h += `<p class="txt">${d.m ? esc(d.m) : '<span style="color:#9aa4b2">（空）</span>'}</p>`;
  h += sect('核心处罚事由', d.g);
  h += sect('核心证据材料', d.e);
  h += sect('申辩理由', d.defn);

  h += `<h2>核心处罚依据（法条全文）<span class="h2c">${String(d.text||'').replace(/\s/g,'').length} 字</span></h2>`;
  h += `<pre class="txt">${esc(d.text)}</pre>`;

  if (d.laws && d.laws.length){
    h += `<h2>涉案法规<span class="h2c">点击可按该法规筛选</span></h2><div class="tagrow">` +
      d.laws.map(l => `<span class="lawtag" data-law="${esc(l)}">${esc(l)}</span>`).join('') + '</div>';
  }
  h += relOf(d);

  if (d.link){
    const links = String(d.link).split('；').filter(Boolean);
    h += `<h2>原文链接<span class="h2c">${links.length} 份文书</span></h2><p class="hint">` +
      links.map((u,i) => `<a href="${esc(u)}" target="_blank" rel="noreferrer" title="${esc(u)}">${
        links.length === 1 ? '打开原文' : '第 ' + (i + 1) + ' 份'}</a>`).join('　') +
      '</p>';
  }
  dInner.innerHTML = h;
  $('dscroll').scrollTop = 0;
  highlight();
}
function highlight(){
  const q = currentQ();
  if (!q.plain.length) return;
  dInner.querySelectorAll('.txt').forEach(node => {
    let out = esc(node.textContent);
    for (const k of q.plain){
      const re = new RegExp(k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      out = out.replace(re, m => `<mark>${m}</mark>`);
    }
    node.innerHTML = out;
  });
}

/* ============ 交互 ============ */
listEl.addEventListener('mouseover', e => {
  const row = e.target.closest('.row');
  if (!row) return;
  const i = +row.dataset.i;
  if (i !== selIdx) show(i, false);
});
listEl.addEventListener('click', e => {
  const row = e.target.closest('.row');
  if (row){ show(+row.dataset.i, false); saveHash(); }
});
dInner.addEventListener('click', e => {
  const lt = e.target.closest('.lawtag');
  if (lt){
    SEL.laws.add(lt.dataset.law);
    const box = $('filters').querySelector('.fg[data-dim="laws"]');
    if (box && !expanded.laws){
      expanded.laws = true; buildFilters();
    }
    syncChecks();
    if (box) $('filters').querySelector('.fg[data-dim="laws"]').classList.remove('closed');
    apply();
    return;
  }
  const go = e.target.closest('[data-goto]');
  if (go) gotoRow(+go.dataset.goto, true);
});

let tq = null;
qEl.addEventListener('input', () => {
  clearTimeout(tq);
  tq = setTimeout(apply, 110);
});
$('qx').addEventListener('click', () => { qEl.value = ''; apply(); qEl.focus(); });

$('sort').addEventListener('change', e => { sortKey = e.target.value; renderList(); saveHash(); });
$('bClear').addEventListener('click', () => {
  for (const d in SEL) SEL[d].clear();
  qEl.value = '';
  syncChecks();
  apply();
});
let fs = 14.5;
const setFs = v => { fs = Math.max(11.5, Math.min(21, v));
  document.documentElement.style.setProperty('--fs', fs + 'px'); };
$('bFs1').addEventListener('click', () => setFs(fs - 1));
$('bFs2').addEventListener('click', () => setFs(fs + 1));
$('bPrev').addEventListener('click', () => show(selIdx - 1));
$('bNext').addEventListener('click', () => show(selIdx + 1));
$('bPrint').addEventListener('click', () => window.print());
$('bCopy').addEventListener('click', async () => {
  const d = view[selIdx];
  if (!d) return;
  const txt = [
    `## 第 ${d.r} 行　${d.d}　${d.s}`,
    `- 违规类型：${d.k}`,
    `- 来源：${d.src}${d.reg ? '　辖区：' + d.reg : ''}`,
    `- 是否典型案例：${d.v}`,
    `\n**处罚措施（机构+人员）**\n\n${d.m}`,
    d.g ? `\n**核心处罚事由**\n\n${d.g}` : '',
    d.e ? `\n**核心证据材料**\n\n${d.e}` : '',
    d.defn ? `\n**申辩理由**\n\n${d.defn}` : '',
    (d.laws && d.laws.length) ? `\n**涉案法规**\n\n${d.laws.join('；')}` : '',
    `\n**核心处罚依据（法条全文）**\n\n${d.text}`,
    d.link ? `\n原文：${String(d.link).split('；')[0]}` : '',
  ].filter(Boolean).join('\n');
  await putText(txt);
  flash($('bCopy'), '已复制');
});
async function putText(t){
  try { await navigator.clipboard.writeText(t); }
  catch(_){
    const ta = document.createElement('textarea');
    ta.value = t; ta.style.position = 'fixed'; ta.style.left = '-9999px';
    document.body.appendChild(ta); ta.select();
    document.execCommand('copy'); ta.remove();
  }
}
function flash(btn, txt){
  const old = btn.textContent;
  btn.textContent = txt; btn.classList.add('on');
  setTimeout(() => { btn.textContent = old; btn.classList.remove('on'); }, 1300);
}
$('bExp').addEventListener('click', () => {
  const cols = [['r','台账行号'],['no','序号'],['d','处罚时间'],['p','挂网发布日期'],
    ['s','被处罚机构/人'],['m','处罚措施'],['g','核心处罚事由'],['e','核心证据材料'],
    ['h','核心处罚依据（台账摘录）'],['defn','申辩理由'],['k','违规类型'],['src','机构来源'],
    ['reg','地区'],['acts','措施标签'],['laws','涉案法规'],['file','案例文件'],
    ['path','存放路径'],['link','原文链接'],['v','是否典型案例'],['text','法条全文']];
  const cell = v => {
    const s = String(v == null ? '' : v).replace(/"/g, '""');
    return /[",\n\r\t]/.test(s) ? `"${s}"` : s;
  };
  const rows = [cols.map(c => cell(c[1])).join(',')];
  view.forEach(d => rows.push(cols.map(c => {
    const v = d[c[0]];
    return cell(Array.isArray(v) ? v.join('；') : v);
  }).join(',')));
  const blob = new Blob(['\ufeff' + rows.join('\r\n')], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `台账检索结果_${view.length}条.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  flash($('bExp'), `已导出 ${view.length} 条`);
});
$('bHelp').addEventListener('click', () => $('help').classList.add('on'));
$('bHelpClose').addEventListener('click', () => $('help').classList.remove('on'));
$('help').addEventListener('click', e => { if (e.target === $('help')) $('help').classList.remove('on'); });

window.addEventListener('hashchange', () => {
  // 用户手改 #L444 或粘贴别人给的筛选链接时，跟着走（不重复处理自己写的 hash）
  const raw = decodeURIComponent(location.hash.slice(1));
  if (/^\d+$/.test(raw) || /^L\d+$/.test(raw)){
    const r = +(raw[0] === 'L' ? raw.slice(1) : raw);
    if (r !== curR) gotoRow(r, true);
  }
});

document.addEventListener('keydown', e => {
  const inInput = /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName);
  if (e.key === '/' && !inInput){ e.preventDefault(); qEl.focus(); qEl.select(); return; }
  if (e.key === 'Escape'){
    if ($('help').classList.contains('on')){ $('help').classList.remove('on'); return; }
    if (qEl.value){ qEl.value = ''; apply(); return; }
    $('bClear').click(); return;
  }
  if (inInput) return;
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp'){
    e.preventDefault();
    show(selIdx + (e.key === 'ArrowDown' ? 1 : -1));
  }
});

/* ============ 地址栏状态 ============ */
let th = null;
function saveHash(){
  clearTimeout(th);
  th = setTimeout(() => {
    const p = [];
    const qv = qEl.value.trim();
    if (qv) p.push('q=' + encodeURIComponent(qv));
    for (const dim in SEL) if (SEL[dim].size)
      p.push(dim + '=' + [...SEL[dim]].map(encodeURIComponent).join(','));
    if (sortKey !== 'r-asc') p.push('sort=' + sortKey);
    if (curR > 0) p.push('L=' + curR);
    try { history.replaceState(null, '', p.length ? '#' + p.join('&') : location.pathname + location.search); }
    catch(_){}
  }, 150);
}
function parseHash(){
  const raw = decodeURIComponent(location.hash.slice(1));
  if (!raw) return {L:0};
  if (/^\d+$/.test(raw)) return {L:+raw};
  if (/^L\d+$/.test(raw)) return {L:+raw.slice(1)};
  const sp = new URLSearchParams(raw);
  if (sp.get('q')) qEl.value = sp.get('q');
  for (const dim in SEL){
    const v = sp.get(dim);
    if (v) v.split(',').filter(Boolean).forEach(x => SEL[dim].add(x));
  }
  if (sp.get('sort') && SORTERS[sp.get('sort')]){
    sortKey = sp.get('sort'); $('sort').value = sortKey;
  }
  return {L: sp.get('L') ? +sp.get('L') : 0};
}
function gotoRow(r, keepFilters){
  let i = view.findIndex(d => d.r === r);
  if (i < 0 && !keepFilters){
    for (const d in SEL) SEL[d].clear();
    qEl.value = '';
    syncChecks();
    renderList();
    i = view.findIndex(d => d.r === r);
  }
  if (i < 0){
    $('dtitle').textContent = `第 ${r} 行不在当前筛选结果里`;
    return false;
  }
  show(i, true);
  const el = listEl.querySelector(`.row[data-i="${i}"]`);
  if (el) el.scrollIntoView({block:'center'});
  return true;
}

/* ============ 启动 ============ */
function apply(){
  renderList();
  saveHash();
}
buildSrcbar();
buildFilters();
const H = parseHash();
syncChecks();
renderList();
if (H.L) gotoRow(H.L, true);
</script></body></html>
"""


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="导出台账为单文件网页台账（多维筛选 + 法条全文）")
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--out", help="输出 html 路径（默认放 <库父>/_案例库维护留痕/）")
    a = ap.parse_args(argv)

    root = os.path.abspath(a.root)
    items, n_over = collect(root)
    if not items:
        raise SystemExit("台账里没有任何批注，先跑 annotate_basis.py --apply")

    out = a.out or os.path.join(root, OUT_NAME)
    os.makedirs(os.path.dirname(out), exist_ok=True)

    payload = json.dumps(items, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")
    doc = TPL.replace("__DATA__", payload)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(doc)

    n_law = sum(1 for d in items if d["laws"])
    n_reg = sum(1 for d in items if d["reg"])
    size = os.path.getsize(out)
    print("条目 %d 条｜批注框装不下 %d 条｜有涉案法规 %d 条｜有辖区 %d 条｜输出 %.1f KB" % (
        len(items), n_over, n_law, n_reg, size / 1024))
    print("→ %s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
