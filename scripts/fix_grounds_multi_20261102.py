# -*- coding: utf-8 -*-
"""F 列「事由条目漏抽」批量修复（2026-11-02）。

症状：F 只写了第一条，原文还有「二是 / 二、」后续条目整段丢失。
根因：证监局分支旧版把「违反了」当事实段终点，遇到第一条的定性句就截断
      （已在 extract_core._grounds_csrc 根治，本脚本是存量修复）。

安全边界（四条硬约束，任一不满足该行就跳过）：
  1. 文书必须有文字层（<200 字视为扫描件，交 OCR 流程，不在本脚本处理）；
  2. **新条目数必须多于旧条目数**——本脚本只做「补漏」，不做「改写」；
  3. 每一条新事由都要能在原文回现（6-gram 覆盖率 ≥0.85），防止串案；
  4. 新文本长度 ≤1500 字（防空转把整篇塞进来）。

用法：
    python fix_grounds_multi_20261102.py --root <库根>            # 干跑，打印旧新对照
    python fix_grounds_multi_20261102.py --root <库根> --apply    # 落盘（自动备份台账）
"""
import argparse
import os
import re
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

import openpyxl

import extract_core as X
import ledger_style as LS
from check_fields_source import best_cover, is_person_file, pick_inst_file
from refill_from_inst import _to_full_punc

LEDGER = "04_综合索引台账【总表】.xlsx"
MARK = re.compile(r"(?m)^\s*[一二三四五六七八九十]+、")
MAXLEN = 1500
COVER_MIN = 0.85


def _unify_style(text, subj=""):
    """把任一路径抽出来的 F 收敛成同一种写法：一条一行、去掉条目内的违规定性句。

    为什么需要：证监局路径（_grounds_csrc）已经剥掉「，违反了《X》第Y条」，
    中基协路径（extract_amac）不会——不统一的话，库里会出现两种事由写法，
    且依据句被塞进事由列（依据本该只在 H 列）。口径一致才算修完。
    """
    lines = [x for x in str(text).split("\n") if x.strip()]
    if len(lines) > 1 and not MARK.match(lines[0]):
        lines = lines[1:]          # 丢掉引导句前言
    out = []
    for ln in lines:
        m = X._item_cut_pos(ln)
        body = ln[:m] if m >= 0 else ln
        body = body.strip()
        # 切掉「…国彤万和。上述行为违反了」留下的**裸主体名**残尾（行245 实测）
        if subj and body.rstrip("。 ").endswith(subj):
            body = body.rstrip("。 ")[: -len(subj)]
        body = LS.tidy_grounds_item(body)
        if len(re.sub(r"^\s*[一二三四五六七八九十]+、", "", body).strip("。；;，, ")) >= 4:
            out.append(body)
    return LS.normalize_grounds_items("\n".join(out)).strip()


def count_items(f):
    return len(MARK.findall(str(f)))


def ground_from(txt, source):
    """按来源选抽取器 —— 这是本轮最大的一条教训：

    中基协决定书的段落结构是「一、基本事实 / 二、申辩意见 / 三、审理意见与纪律处分决定」，
    事实段止于「二、申辩意见」；证监局/交易所止于「我局决定/现决定」。
    拿证监局那条边界去切中基协文书，会把申辩意见 + 审理意见整段并进事由
    （行45 实测：抽出来 13 条，里面全是申辩和审理内容）。
    所以必须先看 K 列来源再选抽取器，不能让「一个函数包打天下」。
    """
    src = str(source or "")
    if "中基协" in src or "amac" in src.lower():
        try:
            return X.extract(txt, "amac").get("grounds", "") or ""
        except Exception:
            return ""
    return X._grounds_csrc(txt)


_PUNC = re.compile(r"[\s\u3000，。；：、,.;:（）()《》“”\"'‘’…—–\-|｜]")


def _skeleton_len(s):
    """去标点后的字数——用于「补漏只能变长」的护栏（标点差异不算实质变化）。"""
    return len(_PUNC.sub("", str(s)))


def read_text(path, cache_dir=None, ocr=False):
    import fitz

    d = fitz.open(path)
    try:
        t = re.sub(r"\s+", "", "".join(p.get_text() for p in d))
    finally:
        d.close()
    if len(t) >= 200:
        return t
    if not ocr:
        return t
    # 中基协纪律处分决定书大量是扫描件，不 OCR 就整批判不了（77 行实测）
    try:
        import ocrlib

        return re.sub(r"\s+", "", ocrlib.ocr_text(path, cache_dir=cache_dir))
    except Exception:
        return t


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--ledger", default="")
    ap.add_argument("--rows", default="")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--ocr", action="store_true", help="扫描件补 OCR 后再判")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    ledger = a.ledger or os.path.join(root, LEDGER)
    cache_dir = os.path.join(os.path.dirname(root), "_案例库维护留痕", "_ocr_cache")
    only = {int(x) for x in re.findall(r"\d+", a.rows)} if a.rows else None

    wb = openpyxl.load_workbook(ledger)
    ws = wb.active

    plan, skip = [], []
    for r in range(2, ws.max_row + 1):
        if only and r not in only:
            continue
        old = str(ws.cell(r, 6).value or "").strip()
        if not old:
            continue
        rel = str(ws.cell(r, 13).value or "").strip()
        files = [x.strip() for x in re.split(r"[；;]", str(ws.cell(r, 12).value or "")) if x.strip()]
        if not files:
            continue
        inst = pick_inst_file(files)
        prim = inst or (files[0] if len(files) == 1 else "")
        if not prim:
            skip.append((r, "无机构文书，且人员文书不唯一"))
            continue
        p = os.path.join(root, rel, prim)
        if not os.path.isfile(p):
            skip.append((r, "文件缺失"))
            continue
        txt = read_text(p, cache_dir, a.ocr)
        if len(txt) < 200:
            skip.append((r, "无文字层（待 OCR）"))
            continue

        new = ground_from(txt, str(ws.cell(r, 11).value or ""))
        if not new:
            continue
        subj = str(ws.cell(r, 4).value or "").split("；")[0].strip()
        new = LS.strip_page_noise(new)
        new = LS.normalize_grounds(new, subj, split_parallel=True)
        new = _unify_style(new, subj)
        new = _to_full_punc(new)
        new = "\n".join(LS.tidy_grounds_item(x) if x.strip() else x for x in new.split("\n")).strip()

        if new.strip() == old.strip():
            continue
        # 旧值是「整段无序号」时不动：中基协部分案例的事实段本身是连续叙述，
        # 改成条目化会丢掉叙述里的第一条（行111 实测：旧的「未办理基金备案」在新值里没了）。
        if count_items(old) == 0:
            skip.append((r, "旧值为整段无序号，条目化会丢原句，需人工"))
            continue
        # 只补漏，不改写
        if count_items(new) <= count_items(old):
            skip.append((r, "条目数未增加（%d→%d）" % (count_items(old), count_items(new))))
            continue
        # 护栏：补漏只能变长。变短说明某条被 ITEM_CUT_VIOL 削掉了（行179 实测：
        # 「资产证明文件不符合《私募办法》对合格投资者的认定标准」在「不符合」处被切断）。
        # 比较时先把旧值里的**前言行**去掉——F 列不该有「XX公司存在以下问题。」这种引导句，
        # 去掉前言会让总字数变少，但那是修正而不是丢失（行130 实测）。
        old_body = "\n".join(x for x in str(old).split("\n") if MARK.match(x)) or str(old)
        if _skeleton_len(new) < _skeleton_len(old_body) * 0.98:
            skip.append((r, "新文本比旧文本短（%d→%d 字），疑似条目被削" % (_skeleton_len(old_body), _skeleton_len(new))))
            continue
        if len(new) > MAXLEN:
            skip.append((r, "新文本过长 %d 字" % len(new)))
            continue
        # 护栏：中基协决定书的「一、基本事实 / 二、纪律处分决定 / 三、申辩意见」是**段落标题**，
        # 不是违规条目。OCR 路径若把它们当条目，整行内容会变成结构目录（行280/351 实测）。
        if re.search(r"(?m)^\s*[一二三四五六七八九十]+、\s*(?:基本事实|纪律处分决定|申辩意见|协会认为|审理意见|复核意见)",
                     new):
            skip.append((r, "抽到的是段落标题（基本事实/处分决定），非违规条目"))
            continue
        # 护栏：悬空序号「二。」「三、」——末尾有个空标题，说明切分错位
        if re.search(r"(?m)^\s*[一二三四五六七八九十]+[、.]?\s*。?\s*$", new) or re.search(r"[、.]\s*$", new.strip()):
            skip.append((r, "末尾有悬空序号，切分错位"))
            continue
        m_dig = LS.CJK_DIGIT_CJK.search(new)
        if m_dig:
            skip.append((r, "中文夹数字噪声「%s」，需人工" % m_dig.group(0)))
            continue
        # 序号必须严格递增不重复（行251 实测：原文是「二、…一是…二是…」的子项结构，
        # 拆完出现两个「二、」，属 X28 类错，先跳过交人工，别把错的写进去）
        nums = re.findall(r"(?m)^\s*([一二三四五六七八九十]+)、", new)
        if nums and nums != list("一二三四五六七八九十"[:len(nums)]):
            skip.append((r, "序号重复/回退 %s，需人工" % "".join(nums)))
            continue
        bad = []
        # 回现比对前先抹平两件「台账已归一、原文没有」的差异，否则全是假线索：
        #   ① 页码噪声（原文「部分—2—所管理」，台账已剥成「部分所管理」——行122）；
        #   ② 称呼置换（原文「你公司」，台账写成机构全称——行3）。
        txt_cmp = LS.strip_page_noise(txt)
        if subj:
            txt_cmp = txt_cmp.replace(subj, "公司")
        for seg in new.split("\n"):
            seg2 = re.sub(r"^\s*[一二三四五六七八九十]+、", "", seg)
            if subj:
                seg2 = seg2.replace(subj, "公司")
            if len(seg2) >= 12 and best_cover(seg2, txt_cmp) < COVER_MIN:
                bad.append(seg2[:24])
        if bad:
            skip.append((r, "条目回现不足：%s" % "／".join(bad)[:60]))
            continue
        plan.append((r, prim, old, new))

    print("=" * 78)
    print("F 列事由漏抽修复（%s）" % ("落盘" if a.apply else "干跑"))
    print("=" * 78)
    print("待修 %d 行｜跳过 %d 行" % (len(plan), len(skip)))
    for r, fn, old, new in plan:
        print("\n行%-4d %s  [%d 条 → %d 条]" % (r, fn[:38], count_items(old), count_items(new)))
        print("  旧：%s" % old[:150].replace("\n", " / "))
        print("  新：%s" % new[:400].replace("\n", " / "))
    if skip:
        print("\n--- 跳过 ---")
        for r, why in skip[:40]:
            print(" 行%-4d %s" % (r, why))

    if not a.apply:
        print("\n（干跑结束，加 --apply 落盘）")
        return 0
    if not plan:
        print("\n无待修行")
        return 0

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ledger.replace(".xlsx", "_bak_grounds_%s.xlsx" % stamp)
    shutil.copy2(ledger, bak)
    for r, _fn, _old, new in plan:
        ws.cell(r, 6).value = new
    wb.save(ledger)
    wb.close()
    print("\n✓ 已写入 %d 行｜备份：%s" % (len(plan), bak))
    return 0


if __name__ == "__main__":
    sys.exit(main())
