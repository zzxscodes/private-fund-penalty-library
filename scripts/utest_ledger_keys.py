# -*- coding: utf-8 -*-
"""ledger._existing_keys / _keys_of 单元测试（v3 格式，SKILL 要求四类用例全覆盖）。

v2（2026-10-26）去文号化：案例键 = (source, manager, date_decided)，
文件键 = 每个 filename，链接键 = 每个 url（K/L/M 多值列按「；」拆分）。

v3（2026-10-27）新增「申辩理由」列后，列序变为 15 列。
⚠ 教训：造测试行**不要写死下标**，必须按 ledger_style.COL 组装，
  否则台账一升级列序，测试就会因为取错列而假失败（2026-10-27 实测 4 条假失败）。
"""
import sys, os, importlib.util
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
SK = os.path.dirname(os.path.abspath(__file__))      # 可迁移：不写死盘符
spec = importlib.util.spec_from_file_location("lg", os.path.join(SK, "ledger.py"))
lg = importlib.util.module_from_spec(spec); spec.loader.exec_module(lg)


class FakeWS:
    """只实现 iter_rows(min_row=2, values_only=True) 的假工作表。
    行结构长度取自 ledger_style.CASE_HEADERS（跟着冻结规格走）。"""

    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, min_row=2, values_only=False, max_col=None):
        for r in self.rows:
            yield tuple(r)


def _row(manager, decided, source, fns, urls):
    """按 ledger_style.COL 的**当前**列序构造一行（列变更时自动跟随，不写死下标）。"""
    vals = {
        "seq": 1, "decided": decided, "publish": "2026-10-20", "subject": manager,
        "measures": "管理人：警告", "grounds": "事由", "evidence": "证据",
        "legal_basis": "依据", "defense": "无", "category": "01_x", "source": source,
        "filenames": "；".join(fns), "paths": "；".join("p/" + f for f in fns),
        "urls": "；".join(urls), "typical": "否",
    }
    row = [None] * len(lg.LS.CASE_HEADERS)
    for k, i in lg.LS.COL.items():
        row[i] = vals.get(k)
    return row


CASES = [
    ("正常-同局同管理人同日判重",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "浙江证监局", "manager": "A私募基金管理有限公司",
      "date_decided": "2026-10-12", "filenames": ["f1.pdf"], "urls": ["http://u1"]}, True),
    ("正常-文件名命中也判重（同案例补充文件）",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "浙江证监局", "manager": "B无关主体",
      "date_decided": "2020-01-01", "filenames": ["f1.pdf"], "urls": []}, True),
    ("正常-链接命中也判重",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "浙江证监局", "manager": "B无关主体",
      "date_decided": "2020-01-01", "filenames": ["f9.pdf"], "urls": ["http://u1"]}, True),
    ("边界-同局同日不同管理人不算重复",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "浙江证监局", "manager": "B私募基金管理有限公司",
      "date_decided": "2026-10-12", "filenames": ["f2.pdf"], "urls": ["http://u2"]}, False),
    ("边界-跨局同管理人同日不算重复",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "河北证监局", "manager": "A私募基金管理有限公司",
      "date_decided": "2026-10-12", "filenames": ["f2.pdf"], "urls": ["http://u2"]}, False),
    ("边界-多值列拆分后任一文件命中即判重",
     [_row("A私募基金管理有限公司", "2026-10-12", "中基协",
           ["A私募基金管理有限公司.pdf", "时任合规风控负责人张春阳-A私募基金管理有限公司.pdf"],
           ["http://u1", "http://u2"])],
     {"source": "中基协", "manager": "X",
      "date_decided": "2020-01-01", "filenames": ["时任合规风控负责人张春阳-A私募基金管理有限公司.pdf"],
      "urls": []}, True),
    ("反例-完全无关不判重",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "江苏证监局", "manager": "C私募基金管理有限公司",
      "date_decided": "2026-09-01", "filenames": ["f3.pdf"], "urls": ["http://u3"]}, False),
    ("反例-空记录不判重",
     [_row("A私募基金管理有限公司", "2026-10-12", "浙江证监局", ["f1.pdf"], ["http://u1"])],
     {"source": "", "manager": "", "date_decided": "", "filenames": [], "urls": []}, False),
]

fail = 0
for name, rows, rec, expect in CASES:
    ws = FakeWS(rows)
    ks = lg._keys_of(rec) & lg._existing_keys(ws)
    got = bool(ks)
    ok = (got == expect)
    fail += (not ok)
    print(("  PASS " if ok else "  FAIL ") + name + f"  期望判重={expect} 实际={got}")
print("\n单元测试:", "全部通过" if fail == 0 else f"{fail} 条失败")
sys.exit(1 if fail else 0)
