# -*- coding: utf-8 -*-
"""决定书主体名清洗（SKILL 0.4.14 落地实现）。

背景：2026-09-30 实测，中基协决定书 OCR 正文「当事人：深圳信钰基金管理有限公司根据《中华人
民共和国证券…」无逗号/换行分隔，旧正则 `([^\\s，,。;\\n]{2,60})` 会把「根据《…》」整段吞进
主体名，导致台账主体列 / 台账文件名列 / 实体文件名三处同时污染。

对策：在冒号之后按「停词 + 停字符」双向截断，并做合法性校验（机构后缀 / 人名长度）。
本模块任何正则改动都必须先跑 `python subject_clean.py --selftest`。
"""
import re

STOP_CHARS = "，,。、；;：《<「『【[】] \t\n\r\"'"
# 注意：（( 不列为停字符——「东方融投（福州）投资管理有限公司」等主体名自带括号
STOP_TOKENS = [
    "根据", "不服", "违反", "涉嫌", "存在", "经查", "经查明", "经我会", "本机",
    "申请人", "当事人", "被处罚人", "被处分人", "本会", "我会", "协会",
    "以下", "简称",  # 「正夫君泽（北京）资本管理有限公司（以下简称正夫君泽）」截断用
    "提出", "申请", "登记", "注册", "系", "并", "未", "向", "对", "其",
    "于20", "在20", "自20",  # 只拦截「于2026年…」式日期短语；裸「于/在」会误伤人名为「石在」
]
NAME_RE = re.compile(r"[\u4e00-\u9fa5A-Za-z0-9·（）()\-]")

# ---------------------------------------------------------------- 非人名碎片词表
# 2026-10-29 库内实测（用户实拍第 2 类缺陷「PDF 命名出现不明确」的同类问题）：
# `is_valid_subject` 里「自然人 = 2~8 个汉字」这条太宽，于是正文里的时间/角色/动作碎片
# 被当成高管姓名写进台账 D 列，并拼出「职位-上海通赛投资有限公司.pdf」
# 「签章-广东横琴均成…pdf」「至今-前海闽荣…pdf」「离职-…」这类文件名。
# 实测污染 9 行：至今/之一/以前/之后/挂名/职位/签章/离职/的请求/不知情。
# ⇒ 单一事实源在此；build_pipeline / verify_library / content_repair 一律 import 本表，
#   禁止各自维护一份（否则新碎片只会被其中一处拦住，另一处继续放行）。
NOISE_TOKENS = frozenset({
    "至今", "以前", "之后", "之前", "以上", "以下", "其中", "之一", "之二",
    "挂名", "职位", "职务", "签章", "签名", "盖章", "离职", "任职", "在任", "离任",
    "的请求", "不知情", "知悉", "相关", "有关", "等", "等等", "当事人", "申请人",
    "有关人员", "相关人员", "上述", "不详", "未知", "不明", "不明确", "待补", "缺位",
    "待核", "无", "空缺", "空岗", "未设", "未配",
})
# 尾部虚词：若「X+虚词」与「X」同时出现在同一行的主体列表里，判「X+虚词」是解析碎片
# （实测「杜胜于」与「杜胜」并存，真名是杜胜，多出的「于」是介词）。
TAIL_PARTICLES = ("于", "的", "等", "和", "及", "在", "对", "为", "与", "并", "将", "向")
# 碎片首字（虚词开头必非人名；**不收**于/向/时/成/和 等真实姓氏）
_FRAG_HEAD = frozenset("的了是在对为与并将以其")
# 噪声词根：允许「词根 + 1~2 字」形态（离职后、任职期间、挂名期间…）
_NOISE_ROOTS = ("离职", "离任", "任职", "在任", "挂名", "职位", "职务", "签章", "签名",
                "盖章", "至今", "以前", "以后", "期间")


def is_noise_token(s: str) -> bool:
    """该词是否为「非人名碎片」（含全角/半角空白归一）。

    2026-11-02 实测补强：原实现只做**精确匹配**，于是「离职后」「的安排」这类
    「词根+虚词后缀」形态直接穿透（汕头宜华一行主体被写成
    「肖伟昊、袁榕、离职后、的安排、林浩」）。现补两条形态规则：
      ① 以虚词开头（的/了/是/在/对/为/与/并/将/以/其）→ 必为解析碎片；
      ② 以噪声词根开头且只多 1~2 字（离职后 / 任职期间 / 挂名期间）→ 碎片。
    反例守卫：真实的姓（于、向、时、成、和）**不进**虚词首字表，避免误伤。
    """
    t = (s or "").strip().strip("　 ")
    if t in NOISE_TOKENS:
        return True
    if t[:1] in _FRAG_HEAD:
        return True
    for root in _NOISE_ROOTS:
        if t.startswith(root) and 0 < len(t) - len(root) <= 2:
            return True
    return False


def drop_noise_tokens(names, keep=frozenset()):
    """从姓名列表里剔除碎片词 + 「X+虚词」型重复碎片，保持原顺序、去重。

    `keep`：行内真实的机构名集合——机构名一律不参与剔除（防止把
    「…合伙企业（有限合伙）」里恰好以「等/为」结尾的机构误删）。
    """
    out, seen = [], set()
    keep = set(keep or ())
    lower = {(n or "").strip() for n in names if (n or "").strip()}
    for raw in names:
        n = (raw or "").strip()
        if not n or n in seen:
            continue
        if n not in keep:
            if is_noise_token(n):
                continue
            # 「X+虚词」且 X 也在本行 → 判碎片
            if len(n) >= 3 and n[-1] in TAIL_PARTICLES and n[:-1] in lower:
                continue
        seen.add(n)
        out.append(n)
    return out


# ---------------------------------------------------------------- 角色前缀剥离
# 2026-11-02 实测（用户人工检验前自查发现）：中基协「对从业人员」的决定书正文写
# 「当事人：尉亲，曾登记为浙江优策投资管理有限公司……」/「实际负责深圳市券期资本
# 管理有限公司」，抽取时把**角色短语**连同机构名一起当成机构全称，于是台账 D 列首段
# 变成「曾登记为浙江优策投资管理有限公司」「实际负责浙江优策投资管理有限公司」，
# 文件名也跟着变成「尉亲-曾登记为浙江优策投资管理有限公司.pdf」。
# 危害不止难看：机构名被污染后，同一机构在中基协/证监局的多份文书**不再归为同一主体**，
# 「同机构多案加日期」这类规则全部失效（浙江优策 4 行被拆成 3 个"不同机构"）。
# 反例守卫：只剥多字角色短语，**不收**单字「原」（「原阳县…」是真地名）。
ROLE_PREFIX = ("曾登记为", "曾登记", "实际负责", "实际控制", "时任", "原任", "现任",
               "登记为", "曾任", "时任为")
# 剥离后剩余部分至少要是这个长度——实测真实被污染的机构名剩余 12~13 字
# （浙江优策投资管理有限公司 / 深圳市券期资本管理有限公司）；
# 阈值设 8 可挡住「时任投资有限公司」这类短名误剥。
_ROLE_MIN_REST = 8


def strip_role_prefix(org: str) -> str:
    """剥掉机构名前面的「曾登记为 / 实际负责 / 时任…」角色短语。"""
    s = (org or "").strip()
    for p in ROLE_PREFIX:
        if s.startswith(p) and len(s) - len(p) >= _ROLE_MIN_REST:
            return s[len(p):].strip()
    return s


def clean_subject(raw: str, maxlen: int = 45) -> str:
    """从「当事人：xxx」冒号后的原始串中截出主体全名。失败返回 ''。"""
    if not raw:
        return ""
    s = raw.strip()
    # 先整段砍掉「（以下简称××）」式同位语——否则会残留半个左括号导致合法性校验失败
    s = re.split(r"[（(](?:以下|下称|简称)", s)[0].strip()
    out = []
    for i, ch in enumerate(s):
        if ch in STOP_CHARS:
            break
        stop = False
        for tk in STOP_TOKENS:
            if s.startswith(tk, i):
                stop = True
                break
        if stop:
            break
        if not NAME_RE.match(ch):
            break
        out.append(ch)
        if len(out) >= maxlen:
            break
    name = "".join(out).strip("的了和与、,，")
    return name if is_valid_subject(name) else ""


def is_valid_subject(name: str) -> bool:
    """机构必须有合法后缀；自然人限 2~4 个汉字（含复姓/少数民族长名放宽到 8）。

    2026-10-29 加严：碎片词（NOISE_TOKENS）一律不合法——旧版「2~8 个汉字即人名」
    把「至今/职位/签章/离职」判成高管姓名，污染台账 D 列与文件名。
    """
    if not name:
        return False
    if re.search(r"(有限公司|有限责任公司|股份公司|合伙企业|管理中心|中心|事务所"
                 r"|工作室|公司|企业|基金|资管|有限合伙)$", name):
        return 3 <= len(name) <= 45
    if is_noise_token(name):
        return False
    if re.fullmatch(r"[\u4e00-\u9fa5·]{2,8}", name):
        return True
    return False


def title_subject(title: str) -> str:
    """从中基协标题括号里取官方全名：纪律处分决定书（xxx）/ 〔xxx〕/ (xxx)

    括号可嵌套：〔正夫君泽（北京）资本管理有限公司〕——内层（）是主体名的一部分，
    早期版本用 `[^（）()〔〕【】]{2,60}` 会把「正夫君泽」截成主体名（2026-09-30 实测）。
    """
    if not title:
        return ""
    for pat in (r"〔(.+?)〕", r"【(.+?)】", r"\[(.+?)\]", r"（(.+?)）", r"\((.+?)\)",
                r"（((?:[^（）]|（[^（）]*）)+)）", r"\(((?:[^()]|\([^()]*\))+)\)"):
        m = re.search(pat, title)
        if m:
            name = m.group(1).strip()
            if is_valid_subject(name):
                return name
    # 兜底：标题里裸写的人名/机构名（去掉「纪律处分决定书」等前缀与右侧括号）
    name = re.sub(r"^[^（(〔【\[]*[（(〔【\[]?", "", title)
    name = re.sub(r"[）)〕】\]].*$", "", name).strip()
    return name if is_valid_subject(name) else ""


def strip_brackets(s: str) -> str:
    return re.sub(r"[（）()〔〕\[\]【】]", "", s or "")


def pick_subject(body_name: str, title_name: str) -> tuple:
    """正文名 + 标题名 → (最终名, 判定说明)。"""
    b, t = (body_name or "").strip(), (title_name or "").strip()
    if b and t:
        bs, ts = strip_brackets(b), strip_brackets(t)
        if bs == ts or bs in ts or ts in bs:
            return (t if len(t) >= len(b) else b), "一致"
        return t, "不一致→取标题"
    if t:
        return t, "仅标题"
    if b:
        return b, "仅正文"
    return "", "都缺"


# ---------------- 自检 ----------------
def _selftest():
    cases = [
        # (输入原始串, 期望输出)
        ("当事人：深圳信钰基金管理有限公司根据《中华人民共和国证券投资基金法》", "深圳信钰基金管理有限公司"),
        ("申请人：台州沃源私募基金管理有限公司申请人不服中国证券投资基金业协会", "台州沃源私募基金管理有限公司"),
        ("当事人：东方融投（福州）投资管理有限公司\n经查，你公司", "东方融投（福州）投资管理有限公司"),
        ("当事人：海南易凡私募基金管理有限公司根据《基金法》、《私募", "海南易凡私募基金管理有限公司"),
        ("当事人：石在，男，1970年出生", "石在"),
        ("当事人：深圳市前海腾创投资有限公司根据《中华人民共和国", "深圳市前海腾创投资有限公司"),
        ("申请人：金佩君，女，1986年2月出生", "金佩君"),
        ("当事人：", ""),
        ("", ""),
        ("申请人不服本会作出的决定", ""),
        # 括号嵌套：不得在「（以下简称…）」处截断，也不得丢掉（北京）
        ("当事人：正夫君泽（北京）资本管理有限公司（以下简称正夫君泽）根据《基金法》",
         "正夫君泽（北京）资本管理有限公司"),
        # 停字符《 生效
        ("当事人：海南易凡私募基金管理有限公司（以下简称", "海南易凡私募基金管理有限公司"),
    ]
    ok = True
    for raw, exp in cases:
        got = clean_subject(raw.split("：", 1)[1] if "：" in raw else raw)
        flag = "OK " if got == exp else "FAIL"
        if got != exp:
            ok = False
        print(f"  {flag} {raw[:34]:<36} -> {got!r} (期望 {exp!r})")
    # 反例：不得把正文整句当主体
    bad = clean_subject("深圳信钰基金管理有限公司根据《中华人民共和国证券投资基金法》"
                        "（以下简称《基金法》、《私募投资基金监督管理暂行办法》")
    assert bad == "深圳信钰基金管理有限公司", bad
    assert is_valid_subject("深圳信钰基金管理有限公司")
    assert not is_valid_subject("根据《中华人民共和国证券")
    assert not is_valid_subject("")
    # 标题括号嵌套（2026-09-30 实测：〔正夫君泽（北京）资本管理有限公司〕）
    assert title_subject("纪律处分决定书〔正夫君泽（北京）资本管理有限公司〕") == \
        "正夫君泽（北京）资本管理有限公司", title_subject("纪律处分决定书〔正夫君泽（北京）资本管理有限公司〕")
    assert title_subject("纪律处分复核决定书（徐铭崎）") == "徐铭崎"
    assert title_subject("纪律处分决定书（深圳前海玖瀛资产管理有限公司）") == "深圳前海玖瀛资产管理有限公司"
    assert title_subject("纪律处分决定书（逄颖颖)") == "逄颖颖"
    assert title_subject("") == ""
    print("  OK  标题解析校验通过")
    print("  OK  反例校验通过")

    # ---------------- 碎片词（2026-10-29 库内实测 9 行污染）----------------
    # 正常：真名/机构名必须放行
    for n in ("汪磊", "张伟民", "王晓群", "东方融投（福州）投资管理有限公司"):
        assert is_valid_subject(n), n
    # 反例：正文碎片不得被当成人名（旧版「2~8 汉字即人名」全部放行）
    for n in ("至今", "职位", "签章", "离职", "以前", "之后", "挂名", "之一", "的请求", "不知情"):
        assert not is_valid_subject(n), n
    # 边界：「X+虚词」型重复碎片（实测「杜胜于」与「杜胜」并存，真名杜胜）
    assert drop_noise_tokens(["至今", "杜胜于", "徐著峰", "杜胜"]) == ["徐著峰", "杜胜"]
    assert drop_noise_tokens(["李晓晴", "李薇薇", "离职", "的请求", "焦健"]) == ["李晓晴", "李薇薇", "焦健"]
    # 反例：机构名即便以虚词结尾也不得被删（keep 参数）
    keep = {"上海通赛投资有限公司"}
    assert drop_noise_tokens(["职位", "上海通赛投资有限公司"], keep=keep) == ["上海通赛投资有限公司"]
    print("  OK  碎片词校验通过（10 反例 + 2 边界 + 1 机构豁免）")

    # ---------------- 角色前缀（2026-11-02 实测 4 行污染）----------------
    assert strip_role_prefix("曾登记为浙江优策投资管理有限公司") == "浙江优策投资管理有限公司"
    assert strip_role_prefix("实际负责浙江优策投资管理有限公司") == "浙江优策投资管理有限公司"
    assert strip_role_prefix("曾登记为杭州合信投资管理有限公司") == "杭州合信投资管理有限公司"
    assert strip_role_prefix("实际负责深圳市券期资本管理有限公司") == "深圳市券期资本管理有限公司"
    # 反例：干净机构名一字不动；单字「原」不剥（原阳县是真地名）
    for n in ("浙江优策投资管理有限公司", "原阳县金泉投资管理有限公司", "时任投资有限公司"):
        assert strip_role_prefix(n) == n, n
    print("  OK  角色前缀校验通过（4 正例 + 3 反例）")
    return ok


if __name__ == "__main__":
    import sys
    print("subject_clean 自检：")
    sys.exit(0 if _selftest() else 1)
