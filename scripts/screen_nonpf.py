# -*- coding: utf-8 -*-
"""screen_nonpf.py v4 — 以「中基协登记」为唯一决定性阳性标准的私募筛选审计。

判定逻辑与 classify_filter.is_private_fund 同源（单一事实源，禁止本地副本）：
  被处罚主体在中国证券投资基金业协会完成私募基金管理人登记（或私募基金产品完成备案）
  → 即认定为私募机构、必须收录；名称像不像私募一律不误杀（特例：企业管理咨询/
    网络科技/投资管理 命名的真实私募管理人）。
  仅当主体名明确为「绝不会是中基协私募管理人」的持牌金融机构（证券/期货/银行/信托/
  保险/公募/基金销售）才剔除。

用法：
  python screen_nonpf.py --root "<库根>" [--use-ledger] [--src 02|03|all]
"""
import argparse, os, fitz
from classify_filter import is_private_fund  # 唯一事实源，与 build 管道一致
# P3 名册闸门：与 build 管道统一走 decide（文本判定 + AMAC 管理人名册复核）
from amac_lookup import decide, AmacRegistry, DEFAULT_DIR as _AMAC_REG_DIR


def extract_text(p):
    try:
        d = fitz.open(p)
        t = "".join(pg.get_text() for pg in d)
        d.close()
        return t
    except Exception:
        return ""


def classify(subject, text, registry=None):
    """与 decide 同源；返回 (状态, 依据)。状态：pf=收录 / nonpf=剔除。"""
    ok, why = decide(text, subject, "", registry)
    return ("pf" if ok else "nonpf", why)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--use-ledger", action="store_true")
    ap.add_argument("--src", default="all")
    ap.add_argument("--registry-dir", default=_AMAC_REG_DIR,
                   help="AMAC 管理人名册目录；缺失则退化为纯文本判定")
    ap.add_argument("--no-registry", action="store_true", help="关闭 P3 名册闸门")
    args = ap.parse_args()
    root = args.root
    registry = None
    if not args.no_registry:
        try:
            registry = AmacRegistry(os.path.abspath(args.registry_dir))
        except Exception:
            registry = None
    if args.use_ledger:
        import openpyxl
        wb = openpyxl.load_workbook(os.path.join(root, "04_综合索引台账【总表】.xlsx"), data_only=True)
        ws = wb.active
        total = pf = nonpf = 0
        finds = []
        for ri in range(2, ws.max_row + 1):
            src = ws.cell(row=ri, column=11).value or ""
            if args.src != "all" and args.src not in src:
                continue
            if src not in ("证监局", "交易所"):
                continue
            subj = (ws.cell(row=ri, column=4).value or "")
            basis = (ws.cell(row=ri, column=7).value or "")
            reason = (ws.cell(row=ri, column=6).value or "")
            text = subj + "\n" + basis + "\n" + reason
            v, why = classify(subj, text, registry)
            total += 1
            if v == "pf":
                pf += 1
            else:
                nonpf += 1
                finds.append((subj, src, why))
        print("[screen] 台账扫描 %d | 收录(PF)=%d 非私募(剔除)=%d" % (total, pf, nonpf))
        for s, src, why in finds:
            print("  [剔除] %s << %s" % (s, why))
        return
    srcs = (["02_各地证监局行政监管（警示函、责令改正、行政处罚、市场禁入）",
             "03_交易所监管措施（沪深交易所、北交所，私募产品交易、异常交易纪律）"]
            if args.src == "all" else [args.src])
    total = pf = nonpf = 0
    finds = []
    for top in srcs:
        base = os.path.join(root, top)
        if not os.path.isdir(base):
            continue
        for dp, dn, fn in os.walk(base):
            if any(x in dp for x in ["_recover", "_case_excluded", "_backup"]):
                continue
            for f in fn:
                if not f.lower().endswith(".pdf"):
                    continue
                p = os.path.join(dp, f)
                t = extract_text(p)
                subj = f[:-4]  # 文件名(主体)最干净，用于名称判定；法规/登记判定走全文 t
                v, why = classify(subj, t, registry)
                total += 1
                if v == "pf":
                    pf += 1
                else:
                    nonpf += 1
                    finds.append((os.path.relpath(p, root), why))
    print("[screen] PDF扫描 %d | 收录(PF)=%d 非私募(剔除)=%d" % (total, pf, nonpf))
    for path, why in finds:
        print("  %s  << %s" % (path, why))


if __name__ == "__main__":
    main()
