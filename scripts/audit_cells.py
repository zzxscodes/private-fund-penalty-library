#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""audit_cells.py —— 台账**每一格**内容体检（2026-11-02，逐格，不只逐列）

与另外两个体检器的分工：
  verify_library.py   要 PDF + 网络，面向交付验收（数量/对账/孤儿/正文回现）
  audit_columns.py    只看列级规则，秒级（空值、占位、枚举、日期）
  audit_cells.py      **本文件**：把 508×15 个格子逐个过形态检查——空格污染、
                      信用代码残片、未分点、网页导语、叠字、括号不配对、首字标点……

设计取舍（为什么这么判）：
  ① 空格只算 **半角空格 U+0020 / 全角 U+3000 / 制表符**，**换行 \\n 一律不算**——
     措施列、事由列的分条换行是**要的**，把 \\n 当污染会把正常分条也报出来（第一版实测误报 183 条）。
  ② 判「未分点」不能只看有没有「一、」：单条违规本来就不该编号。所以只在
     **长度 > 40 且含 ≥2 个并列分隔** 时才判「疑似多条未拆分」，单条短句归观察项。
  ③ 判「网页导语」用「等问题」「等行为」「如下」——这是证监局网页摘要的固定尾巴，
     决定书正文不会这么收尾。命中即说明 F 列取的是网页摘要而不是文书事实段。
  ④ 信用代码残片 `\\d{15,18}[A-Z]?\\s*[）)]` —— 正常文书里代码在**括号内**且前面有主体名，
     出现在句首、后面紧跟「）」的，一定是抽取时把「（统一社会信用代码：」切掉了。

输出：按「列 × 问题类型」的计数矩阵 + 每类采样（默认每类前 8 条），最后给总判定。
用法：python scripts/audit_cells.py --root <库根> [--show 12] [--col F]
"""
import argparse
import collections
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
import openpyxl  # noqa: E402

import ledger_style as LS  # noqa: E402

NAMES = ["A序号", "B处罚时间", "C挂网日期", "D被处罚机构人", "E处罚措施", "F核心处罚事由",
         "G核心证据材料", "H核心处罚依据", "I申辩理由", "J违规类型", "K机构来源",
         "L案例文件", "M存放路径", "N原文链接", "O是否典型案例"]

# ---------- 通用形态 ----------
SP_GAP = re.compile(r"(?<=[\u4e00-\u9fa5，。；：、）》】”’])[ \t\u3000]+(?=[\u4e00-\u9fa5《（【“‘])")
SP_STRAY = re.compile(r"[ \t\u3000]{2,}|\t")
FULL_SP = re.compile(r"\u3000")
DUP = re.compile(r"([\u4e00-\u9fa5]{2})\1")
DUP_OK = {"常常", "渐渐", "刚刚", "仅仅", "屡屡", "往往", "明明", "统统", "统统"}
PAREN_OPEN = re.compile(r"[（(]")
PAREN_CLOSE = re.compile(r"[）)]")
NUM_MARK = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|[（(]\d+[）)]|\d+[.、])", re.M)
PAGE = re.compile(r"[—–]\s*\d{1,3}\s*[—–]|第\s*\d{1,3}\s*页")
CODE = re.compile(r"(?:^|[。；\n])\s*\d{15,18}[A-Z0-9]?[ \t]*[）)]")
HALF_PUNCT = re.compile(r"[,;:]")
WEB_SHELL = re.compile(r"中国政府网|京ICP备|京公网安备|网站识别码|联系我们\s*\|?\s*法律声明|"
                       r"当前位置：首页|请输入关键字|主办单位：|版权所有：|证监局介绍")
FW_ALNUM = re.compile(r"[０-９Ａ-Ｚａ-ｚ]")
LEAD_PUNCT = re.compile(r"^\s*[，。、；：）)\]】]")
TAIL_SEMI = re.compile(r"[；;]\s*$")
# 页脚/印发信息特征（F/G/H/I 列出现即串味）
FOOTER = re.compile(r"证监会市场二司|法治司[；;]|办公室\s*20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*印发|"
                    r"抄送[:：]|MERGEFORMAT|主办单位[:：]")
# 行首中文序号（X28 用）：一、二、…十、；不含「（一）」「1.」——那两类不会出现"回退"歧义
CN_IDX = re.compile(r"\n\s*([一二三四五六七八九十]+)、")
_CN_D = "零一二三四五六七八九"


def _cn2int(s):
    """中文数字 → int（1~99）：一=1 十=10 十二=12 二十=20 二十三=23。"""
    if not s:
        return 0
    if s == "十":
        return 10
    if "十" in s:
        a, _, b = s.partition("十")
        return (_CN_D.index(a) if a in _CN_D and a else 1) * 10 + (_CN_D.index(b) if b in _CN_D and b else 0)
    return _CN_D.index(s) if s in _CN_D else 0
ENUM_J = None  # 运行期从数据里取合法集合

J_SETS = {
    "中基协": {"01_募集与适当性违规", "02_登记备案、信息报送、信披违规",
             "03_投资运作、关联交易、利益输送", "04_内控治理、人员资质、风控缺位",
             "05_基金财产挪用、混同、保本保收益", "06_其他行政违规"},
    "证监局": {"01_募集与适当性违规", "02_登记备案、信息报送、信披违规",
             "03_投资运作、关联交易、利益输送", "04_内控治理、人员资质、风控缺位",
             "05_基金财产挪用、混同、保本保收益", "06_其他行政违规"},
    "交易所": {"01_异常交易、对倒、价格操纵", "02_产品交易风控、持仓限制违规",
             "03_信息报送、交易行为自律监管", "04_其他自律监管"},
}


def S(v):
    return "" if v is None else str(v)


def check_cell(col, s, row):
    """返回该格命中的问题码列表"""
    out = []
    if not s:
        return out
    body = s  # A/O/K 等短值不查空格
    if col in (3, 4, 5, 6, 7, 8):  # D~I 长文本
        if SP_GAP.search(body):
            out.append("X1半角空格夹中文")
        if SP_STRAY.search(body):
            out.append("X2连续空格/制表符")
        if FULL_SP.search(body):
            out.append("X3全角空格")
        if DUP.search(body):
            hits = [m.group(0) for m in DUP.finditer(body) if m.group(0) not in DUP_OK]
            if hits:
                out.append("X4叠字")
        if PAGE.search(body):
            out.append("X5页码标记")
        if LEAD_PUNCT.match(body):
            out.append("X6首字是标点/右括号")
        if len(PAREN_OPEN.findall(body)) != len(PAREN_CLOSE.findall(body)):
            out.append("X7括号不配对")
        if CODE.search(body):
            out.append("X8信用代码残片")
        if WEB_SHELL.search(body):
            out.append("X26网页外壳混入")
        if FW_ALNUM.search(body):
            out.append("X27全角英数")
        for m in HALF_PUNCT.finditer(body):
            a = body[m.start() - 1] if m.start() else ""
            b = body[m.end()] if m.end() < len(body) else ""
            if not (a.isdigit() and b.isdigit()):
                out.append("X9半角标点")
                break

    if col == 4:  # E 措施
        if re.search(r"不明确|不详|未知|待定|待补", body):
            out.append("X10措施占位")
        if TAIL_SEMI.search(body):
            out.append("X11措施以分号结尾")
        if body.count("管理人：") > 1:
            out.append("X12重复「管理人：」段")

    if col == 5:  # F 事由
        marks = NUM_MARK.findall(body)
        lines = [x for x in body.split("\n") if x.strip()]
        if marks and len(lines) < len(marks):
            out.append("X20序号未换行")
        if not marks and len(body) > 40:
            parts = [p for p in body.split("、") if p.strip()]
            bad = re.compile(r"未|违|不|擅自|挪用|侵占|虚构|隐瞒|超|承诺|混同|让渡|委托|"
                             r"从事|缺失|缺位|不当|虚假|误导|将|向|通过|使用|利用")
            hit = sum(1 for p in parts if bad.search(p))
            avg = sum(len(p) for p in parts) / len(parts) if parts else 0
            is_narrative = any(("，" in p or "," in p) for p in parts)
            if (not is_narrative and len(parts) >= 3 and hit >= 2
                    and hit / len(parts) >= 0.5 and avg >= 6):
                out.append("X21疑似多条未拆分")
            else:
                out.append("X21'单条整段（正常观察）")
        if re.search(r"等问题[。\s]*$|等行为[。\s]*$|如下[：:\s]*$", body):
            out.append("X22网页导语句")
        if re.search(r"^(?:经查|经查明|我局发现|经我局|根据相关行政监管措施决定书)", body):
            out.append("X23叙述引导词开头")
        if marks and not re.search(r"[。；;]\s*$", body):
            out.append("X24末条无句末标点")
        if re.search(r"以上事实|上述事实|我局认为|本局认为|本所认为|我会认为|协会认为|"
                     r"经审理查明|足以认定|我局决定|本局决定|本所决定|现决定|决定对你", body):
            out.append("X25事由越界")
        # X29 页脚/印发信息（2026-11-02）：证监局网页另存 PDF 的页脚是
        # 「证监会市场二司、法治司；中国证券投资基金业协会。广东证监局办公室2026年3月16日印发。」
        # 它**确实出现在该文书里**（Z11 回现率 100%），但位置是页脚，不是事由（行119/312/314/365）。
        if FOOTER.search(body):
            out.append("X29页脚串味")
        # X28 序号回退（2026-11-02）：一份决定书里有**两段并列枚举**（先「一是…六是」，
        # 再「此外，经查…存在以下违规行为：一是…二是」），抽取器把第二段的「一是」当新一段
        # 的第一条 → 台账成了「一、二、三、四、五、六、一、」。形态完全合法（非空、有编号、
        # 枚举值都对），除了「序号回退」这一条没有任何断言能发现（行36/344/408/449 实测）。
        if CN_IDX.findall("\n" + body):
            nums = [_cn2int(x) for x in CN_IDX.findall("\n" + body)]
            nums = [n for n in nums if n]
            if len(nums) >= 3 and any(nums[i] <= max(nums[:i]) for i in range(1, len(nums))):
                out.append("X28序号回退")

    if col == 8:  # I 申辩
        # X32 申辩越界（2026-11-02）：行344 从「审理意见」一路抽到「处分决定 + 复核申请 + 抄送」，
        # 264 字里有 220 字不是申辩。申辩列只装「当事人说了什么」，不装「协会怎么认定、怎么处」。
        if re.search(r"决定作出以下|如对(?:以上|本)[^。]{0,12}有异议|提出书面复核申请|"
                     r"抄送[:：]|鉴于以上基本事实|印发", body):
            out.append("X32申辩越界")
        if FOOTER.search(body):
            out.append("X29页脚串味")

    if col == 7:  # H 依据
        if re.search(r"未载明|不明确|另见|同上", body):
            out.append("X30依据占位")
        if not re.search(r"第[一二三四五六七八九十百零\d]+(?:\.[\d]+)*条", body):
            out.append("X31无条款号")

    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--show", type=int, default=8)
    ap.add_argument("--col", default="", help="只看某列（A~O）")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    ledger = args.ledger or os.path.join(root, "04_综合索引台账【总表】.xlsx")
    if not os.path.isfile(ledger):
        cands = [f for f in os.listdir(root)
                 if f.endswith(".xlsx") and not f.startswith("~$")]
        ledger = os.path.join(root, cands[0])
    ws = openpyxl.load_workbook(ledger, data_only=True).worksheets[0]
    rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    rows = [r for r in rows if S(r[0]).strip()]

    matrix = collections.Counter()          # (col, code) -> n
    samples = collections.defaultdict(list)  # code -> [(row, seq, col, snippet)]
    for i, r in enumerate(rows, start=2):
        for col in range(15):
            s = S(r[col])
            for code in check_cell(col, s, i):
                matrix[(col, code)] += 1
                if len(samples[code]) < args.show:
                    m = (SP_GAP.search(s) or SP_STRAY.search(s) or DUP.search(s)
                         or CODE.search(s) or PAGE.search(s) or re.search(r".{0,18}", s))
                    p = m.start() if m else 0
                    snip = s[max(0, p - 14):p + 30].replace("\n", "⏎")
                    samples[code].append((i, S(r[0]), NAMES[col], snip))

    # 列专属之外，还要查枚举与格式
    enum_bad = []
    for i, r in enumerate(rows, start=2):
        k = S(r[10]).strip()
        j = S(r[9]).strip()
        if k not in J_SETS:
            enum_bad.append((i, S(r[0]), "K机构来源", k))
        elif j not in J_SETS[k]:
            enum_bad.append((i, S(r[0]), "J违规类型", "%s / %s" % (k, j)))
        if S(r[14]).strip() not in ("是", "否"):
            enum_bad.append((i, S(r[0]), "O是否典型", S(r[14])))
        if not re.match(r"^https?://", S(r[13]).strip().split("；")[0]):
            enum_bad.append((i, S(r[0]), "N原文链接", S(r[13])[:50]))
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", S(r[1]).strip()):
            enum_bad.append((i, S(r[0]), "B处罚时间", S(r[1])))
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", S(r[2]).strip()):
            enum_bad.append((i, S(r[0]), "C挂网日期", S(r[2])))

    print("=" * 78)
    print("台账逐格内容体检 —— %d 行 × 15 列 = %d 格" % (len(rows), len(rows) * 15))
    print("=" * 78)

    if not matrix and not enum_bad:
        print("\n未发现形态问题。")
        return 0

    codes = sorted({c for _, c in matrix})
    if args.col:
        keep = {i for i, n in enumerate(NAMES) if n.startswith(args.col)}
        codes = [c for c in codes if any((col in keep) for (col, cc) in matrix if cc == c)]

    hdr = "问题类型".ljust(24) + "".join(n[0].rjust(5) for n in NAMES) + "   合计"
    print("\n" + hdr)
    print("-" * len(hdr))
    total = 0
    for code in codes:
        line = code.ljust(24)
        sub = 0
        for col in range(15):
            v = matrix.get((col, code), 0)
            line += (str(v) if v else "·").rjust(5)
            sub += v
        line += str(sub).rjust(7)
        print(line)
        total += sub

    print("\n【采样】")
    for code in codes:
        if not samples[code]:
            continue
        print("\n-- %s（%d 格）" % (code, sum(matrix.get((c, code), 0) for c in range(15))))
        for it in samples[code]:
            print("   行%-4s 序%-4s %-8s %s" % (it[0], it[1], it[2], it[3]))

    if enum_bad:
        print("\n【枚举/格式异常】%d 处" % len(enum_bad))
        for it in enum_bad[:args.show * 2]:
            print("   行%-4s 序%-4s %-10s %s" % it)

    # 带 ' 的码是观察项（原文本身如此），不计入 problems
    n_obs = sum(v for (c, code), v in matrix.items() if "'" in code)
    n_err = sum(v for (c, code), v in matrix.items() if "'" not in code) + len(enum_bad)
    print("\n" + "=" * 78)
    print("问题格合计 %d（不含观察项）+ 观察项 %d + 枚举/格式 %d" % (n_err, n_obs, len(enum_bad)))
    print("=" * 78)
    return 1 if (n_err + len(enum_bad)) else 0


if __name__ == "__main__":
    sys.exit(main())
