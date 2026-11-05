# -*- coding: utf-8 -*-
"""AMAC 管理人名册离线倒查（P3 决定性阳性路径之一）。

== 用户裁定（2026-10-29，决定性标准）==
  【只要被处罚主体在中基协完成「私募基金管理人登记」，即认定私募机构、必须收录】；
  这包含两类管理人：
    · 在册（active）私募基金管理人；
    · 已注销/已撤销登记（cancelled）的私募基金管理人（即「被处罚私募被注销」情形——
      能被注销/撤销登记的，必然曾是已登记管理人，仍属本库范畴）。

== 为什么要下载名册离线倒查（而非每案联网查）==
  用户明确要求：把 AMAC 集中公示的管理人名册下载下来，离线倒查每个被处罚主体名，
  省去不断联网查询；且【更新案例库时必须重新下载】（名册会变）。

== 为什么【不】用产品（fund）名册 ==
  用户明确：「有登记产品但是从未有登记为管理人的，不是案例库要的案例」。
  「在册 + 已注销管理人」名册查询本身已排除产品-only 实体，故产品名册不参与收录判定。

== 可移植性（原则1）==
  名册落地路径由 --out 指定（默认技能目录下 amac_registry/），绝不硬编码本机绝对路径；
  读取路径（AmacRegistry）只依赖 json，无需 Playwright，可在任意机器加载既有名册。

== 端点（已实测，2026-10-29）==
  在册：   POST /amac-infodisc/api/pof/manager?rand=..&page=P&size=1000   （size=1000 可用，约19页）
  已注销： POST /amac-infodisc/api/cancelled/manager?rand=..&page=P&size=20  （size 上限20，约1269页/25374条）
  （纯 HTTP 被 WAF 拦截，必须经 Playwright 浏览器会话内 page.request 调用）

用法：
    python amac_lookup.py --out DIR            # 下载两大名册到 DIR（更新库时加 --force 重下）
    from amac_lookup import AmacRegistry
    reg = AmacRegistry(DIR)                     # 加载本地名册（不联网）
    reg.is_registered_manager("北京博晟投资管理有限公司")   # -> 'active' / 'cancelled' / None
"""
import sys, os, re, json, time, argparse
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

BASE = "https://gs.amac.org.cn"
EP_ACTIVE = "/amac-infodisc/api/pof/manager"
EP_CANCELLED = "/amac-infodisc/api/cancelled/manager"
NAME_KEY = "managerName"          # 两个端点记录里主体名字段都叫 managerName
DEFAULT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "amac_registry")

# 主体名规范化：去掉常见公司后缀，做全角->半角，便于「文书名 ↔ 名册名」容差匹配
_SUFFIXES = ["股份有限公司", "有限责任公司", "（特殊普通合伙）", "(特殊普通合伙)",
             "合伙企业（有限合伙）", "合伙企业(有限合伙)", "（有限合伙）", "(有限合伙)",
             "有限公司", "合伙企业", "股份公司", "集团", "公司"]


def _fw_to_hw(s):
    """全角字符转半角（名册与文书可能用不同括号/数字）。"""
    out = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            out.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def _norm(name):
    """规范化主体名：去空白 + 全角转半角。"""
    return _fw_to_hw((name or "").strip()).replace(" ", "")


def _core(name):
    """去掉尾部公司/合伙后缀后的核心名（容差匹配用，只剥一层最长后缀）。"""
    n = _norm(name)
    for suf in sorted(_SUFFIXES, key=len, reverse=True):
        if n.endswith(suf) and len(n) > len(suf):
            return n[: -len(suf)]
    return n


class AmacRegistry:
    """加载本地 AMAC 管理人名册，提供离线倒查接口。"""

    def __init__(self, directory):
        self.dir = directory
        self.active = self._load_json(os.path.join(directory, "manager_active.json"))
        self.cancelled = self._load_json(os.path.join(directory, "manager_cancelled.json"))
        # 关键（2026-10-30 实测修 bug）：在册名册记录字段是 managerName，已注销名册记录字段是
        # orgName（结构不同，曾因取 managerName 全为 None 导致 25374 条已注销记录【完全未入索引】，
        # P3 闸门只能查到在册、查不到已注销——用户担心的「先被处罚、后主动注销登记」场景漏收）。
        # 统一把 orgName 归一到 managerName，后续索引/查询只认 managerName。
        for r in self.cancelled:
            if not r.get("managerName") and r.get("orgName"):
                r["managerName"] = r["orgName"]
        # 倒查索引：norm名 -> status；core名 -> status（core 命中优先级低，若冲突优先 active）
        self._idx = {}
        self._build_index(self.active, "active")
        self._build_index(self.cancelled, "cancelled")

    @staticmethod
    def _load_json(fp):
        if not os.path.exists(fp):
            return []
        try:
            return json.load(open(fp, encoding="utf-8"))
        except Exception:
            return []

    def _build_index(self, records, status):
        for r in records:
            nm = r.get(NAME_KEY)
            if not nm:
                continue
            for key in (_norm(nm), _core(nm)):
                if not key:
                    continue
                # 冲突时 active 优先（在册比已注销更说明是私募管理人）
                if key not in self._idx or (status == "active" and self._idx[key] != "active"):
                    self._idx[key] = status

    def is_registered_manager(self, name):
        """返回 'active' / 'cancelled' / None（未收录或名册未加载）。"""
        if not name:
            return None
        key = _norm(name)
        if key in self._idx:
            return self._idx[key]
        core = _core(name)
        if core and core in self._idx:
            return self._idx[core]
        return None

    def has_data(self):
        return bool(self.active) or bool(self.cancelled)

    def summary(self):
        return {"active": len(self.active), "cancelled": len(self.cancelled),
                "indexed": len(self._idx)}


# ----------------------------------------------------------------- 下载（Playwright 浏览器会话）
def _page_request(pg, endpoint, page, size):
    """POST 名册 API 单页。返回 (json_or_None, status)。异常（含 ECONNRESET）返回 (None, -1)。"""
    import random
    url = BASE + endpoint
    q = "?rand=%s&page=%d&size=%d" % (round(random.random(), 16), page, size)
    try:
        r = pg.request.post(url + q, data="{}", headers={"Content-Type": "application/json"})
    except Exception as e:
        return None, -1
    if r.status != 200:
        return None, r.status
    try:
        j = r.json()
    except Exception:
        return None, 0
    return j, 200


def _save(out_fp, recs):
    json.dump(recs, open(out_fp, "w", encoding="utf-8"), ensure_ascii=False, indent=0)


def _download_endpoint(pg, endpoint, size, out_fp, max_pages=100000, delay=0.15,
                      session_refresh=None, refresh_every=150):
    """分页下载一个端点，增量落盘（每页写盘，断点可续）。返回 (total, got)。

    健壮性（2026-10-29 实测：翻到 ~891 页时服务器 ECONNRESET 断流）：
      · 异常/非200 计入连续错误，退避重试；连续失败过多则保存进度并退出（可重跑续传）；
      · 每 page 增量写盘，崩了也保留已抓部分；重跑时从已有记录数推算续传起点（无缺页）；
      · 每 refresh_every 页调用 session_refresh() 刷新浏览器会话，规避长连接被掐。
    """
    recs = []
    if os.path.exists(out_fp):                  # 断点续传：装载已有记录
        try:
            recs = json.load(open(out_fp, encoding="utf-8"))
        except Exception:
            recs = []
    # 续传起点：已抓页数 = ceil(len/size)；若末页恰好满页则 len%size==0 从 len//size 续，
    # 若末页不满则末页已完整、从 len//size+1 续（避免重复拉取最后一个不满页造成重复记录）。
    page = (len(recs) + size - 1) // size
    total = None
    consecutive_err = 0
    while page < max_pages:
        if session_refresh and page and page % refresh_every == 0:
            try:
                session_refresh()
            except Exception:
                pass
        j, st = _page_request(pg, endpoint, page, size)
        if st != 200 or not j:
            consecutive_err += 1
            print("    [WARN] %s page=%d 状态 %s，重试(%d)" % (endpoint, page, st, consecutive_err))
            if consecutive_err >= 6:
                print("    [ERR] %s 连续失败，保存进度并退出（已得 %d，可重跑续传）" % (endpoint, len(recs)))
                _save(out_fp, recs)
                return total, len(recs)
            time.sleep(2.0 * consecutive_err)
            continue
        consecutive_err = 0
        content = j.get("content") or []
        recs += content
        if total is None:
            total = j.get("totalElements")
            pages = j.get("totalPages") or 0
            print("    %s 总数 %s，共 %d 页，size=%d（续传起点 page=%d）"
                  % (endpoint, total, pages, size, page))
        if page % 50 == 0:                      # 周期性落盘，降低崩损
            _save(out_fp, recs)
        page += 1
        if (total is not None and len(recs) >= total) or not content or page >= (j.get("totalPages") or 10**9):
            break
        time.sleep(delay)
    _save(out_fp, recs)                          # 末态落盘
    return total, len(recs)


def download(out_dir=None, force=False, headless=True):
    """下载在册 + 已注销管理人名册到 out_dir。返回 AmacRegistry。

    force=False 时若名册已存在则跳过下载（直接加载）；force=True 强制重下（更新库时用）。
    已下载部分支持断点续传（重跑从已有记录续传）。
    """
    out_dir = os.path.abspath(out_dir or DEFAULT_DIR)
    os.makedirs(out_dir, exist_ok=True)
    fp_active = os.path.join(out_dir, "manager_active.json")
    fp_cancelled = os.path.join(out_dir, "manager_cancelled.json")

    need_active = force or not os.path.exists(fp_active)
    need_cancelled = force or not os.path.exists(fp_cancelled)

    try:
        from playwright.sync_api import sync_playwright
    except Exception as e:
        print("[ERR] 未安装 playwright，无法下载名册：%s" % e)
        return AmacRegistry(out_dir)

    with sync_playwright() as p:
        b = p.chromium.launch(headless=headless)
        pg = b.new_page()
        def refresh():
            # 刷新会话：重新打开参考页，规避长连接被服务器掐断（ECONNRESET）
            pg.goto(BASE + "/amac-infodisc/res/pof/manager/index.html",
                    wait_until="domcontentloaded", timeout=60000)
            pg.wait_for_timeout(1500)
        # 先打开在册页拿会话
        pg.goto(BASE + "/amac-infodisc/res/pof/manager/index.html",
                wait_until="domcontentloaded", timeout=60000)
        pg.wait_for_timeout(2000)
        if need_active:
            print("[下载] 在册私募基金管理人（size=1000）...")
            _download_endpoint(pg, EP_ACTIVE, 1000, fp_active, refresh_every=10**9)
        else:
            print("[跳过] 在册名册已存在（--force 可重下）")
        if need_cancelled:
            # cancelled 端点 size 上限 20，需翻多页；长连接易断，开启会话周期刷新 + 断点续传
            print("[下载] 已注销私募基金管理人（size=20，约1269页）...")
            _download_endpoint(pg, EP_CANCELLED, 20, fp_cancelled, delay=0.08,
                              session_refresh=refresh, refresh_every=120)
        else:
            print("[跳过] 已注销名册已存在（--force 可重下）")
        b.close()
    reg = AmacRegistry(out_dir)
    print("[完成] 名册摘要：", reg.summary())
    return reg


def decide(text, subject="", title="", registry=None):
    """综合收录判定（P1/P2 文本 + P3 名册闸门），返回 (是否收录, 判定依据)。

    这是 build / screen / recon 管道的统一收录入口（单一事实源）：
      1) 先按 classify_filter.classify_pf 得到 (inc, why, strength)；
      2) 若 registry 已加载数据：
         · hard 收（AMAC_REG 登记编号/备案、CANCEL_REG 注销登记、SUBJECT_PE 文书直接指认、
           正文点名机构全称 step 5.5）→ 直接收，压倒名册（名称变异导致名册未命中也收）；
         · soft 收（_PF_NAME+PF_LAWS / PF_WEAK+_PF_NAME）但名册【未命中】
           主体为管理人 → 推翻为剔除（排除「上市公司顺带引用私募法/提及第三方私募」假阳性）；
         · 文本剔除 但名册命中（active/cancelled）→ 推翻为收录（P3 压倒文本，
           覆盖「先被处罚、后被注销」及文本未载明登记的主体）；
      3) 若 registry 未加载（可移植首跑/离线）→ 退化为纯文本 classify_pf，流程不中断。
    """
    from classify_filter import classify_pf
    inc, why, strength = classify_pf(text, subject, title)
    if registry is not None and registry.has_data():
        reg = registry.is_registered_manager(subject)
        if inc and strength == "soft" and reg is None:
            return False, "软证据疑似私募但AMAC管理人名册未命中（主体非登记管理人）"
        if (not inc) and reg is not None:
            return True, "AMAC管理人名册倒查命中（%s）" % reg
    return inc, why


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="下载/加载 AMAC 私募基金管理人名册（在册+已注销）")
    ap.add_argument("--out", default=DEFAULT_DIR, help="名册落地目录（默认技能下 amac_registry/）")
    ap.add_argument("--force", action="store_true", help="强制重新下载（更新案例库时用）")
    ap.add_argument("--headless", action="store_true", default=True, help="无头浏览器")
    ap.add_argument("--no-download", action="store_true",
                    help="只加载本地名册并打印摘要，不下载")
    a = ap.parse_args()
    if a.no_download:
        reg = AmacRegistry(os.path.abspath(a.out))
        print("名册摘要：", reg.summary())
        # 抽样自测
        for t in ["北京博晟投资管理有限公司", "天津暖流企业管理咨询股份有限公司",
                  "中国工商银行股份有限公司", "某不存在的有限公司"]:
            print("  %s -> %s" % (t, reg.is_registered_manager(t)))
    else:
        download(a.out, force=a.force, headless=a.headless)
