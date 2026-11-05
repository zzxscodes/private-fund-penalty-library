#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""make_macro_workbook.py —— 给台账正本生成「宏版孪生体」(.xlsm)

为什么要分两步
--------------
本机（2026-11-04 实测）Excel 用 COM `Workbooks.Open` 打开稍大的文件一律返回
`RPC_E_CALL_REJECTED`，所以不能「打开台账 → 另存为 xlsm」。但 `Workbooks.Add()`
可以、`VBProject` 可以、`SaveAs(xlsm)` 也可以。

于是：
  步骤 A（Excel）：建一个 2 个工作表的空白簿 → 注入 VBA 模块与 ThisWorkbook 事件
                  → 另存为临时 .xlsm，只为拿它的 `xl/vbaProject.bin`。
  步骤 B（Python）：把这个 vbaProject.bin 移植进台账正本的包，
                  补 `[Content_Types].xml` 与 `xl/_rels/workbook.xml.rels`。

好处：台账的 sheet/comments/VML **全程不经 Excel**，字节级不变——
批注框的 660pt 尺寸与锚点一个都不会被 Excel 重算。

用法
----
    python make_macro_workbook.py --root <案例库根目录>
    python make_macro_workbook.py --root <库> --bas <自己的模块.bas>
    python make_macro_workbook.py --root <库> --vba-src <已建好的宏源.xlsm>   # 复用，不启动 Excel

产物：<库>/04_综合索引台账【总表】.xlsm（同名正本的孪生体）
校验：verify_library 的 C10b 会比对两者四个部件是否逐字节相同，过期即报问题。

依赖：pywin32 + 本机可用 Excel。VBA 工程访问需要 `AccessVBOM=1`，
本脚本自己写、自己还原，不依赖外部手工设置。
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile

LEDGER_BASENAME = "04_综合索引台账【总表】"
SEC_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"

THISWB_CODE = """
Private Sub Workbook_Open()
    On Error Resume Next
    modLawPanel.LogLine "Workbook_Open 触发 " & Format$(Now, "hh:nn:ss") & _
                        " Excel=" & Application.Version & " 文件=" & ThisWorkbook.Name
    Application.StatusBar = "法条悬停已启用：把鼠标停在 H 列任一有批注的格子上约 1 秒，会弹出可滚动的法条全文"
    modLawPanel.StartHover
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.StatusBar = False
    modLawPanel.StopHover
End Sub
Private Sub Workbook_Deactivate()
    On Error Resume Next
    modLawPanel.StopHover
End Sub
Private Sub Workbook_Activate()
    On Error Resume Next
    modLawPanel.LogLine "Workbook_Activate " & Format$(Now, "hh:nn:ss")
    modLawPanel.StartHover
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    modLawPanel.OnSelectCell Sh, Target
End Sub
""".replace("\n", "\r\n")

# —— UserForm 窗体代码：窗内 TextBox 是真控件，滚轮由它自己接 ——
FORM_CODE = """
Public LastMove As Double

Private Sub UserForm_Activate()
    Me.txtLaw.SetFocus
    LastMove = Timer
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LastMove = Timer
    Me.txtLaw.SetFocus
End Sub

Private Sub txtLaw_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LastMove = Timer
End Sub
""".replace("\n", "\r\n")

PARTS_TO_COMPARE = ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml",
                    "xl/comments/comment1.xml", "xl/drawings/commentsDrawing1.vml")


def log(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- VBA 工程信任
def set_access_vbom(value):
    """VBA 工程访问开关。实测它会被系统/Excel 异常退出时改回 0，
    所以在**建 Excel 实例之前**写、结束后还原，中间不留空档。"""
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SEC_KEY, 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, value)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SEC_KEY, 0, winreg.KEY_READ) as k:
        return winreg.QueryValueEx(k, "AccessVBOM")[0]


def module_code(bas_path):
    """读 .bas，去掉 `Attribute VB_Name`（那是给「导入文件」用的，AddFromString 不需要）。"""
    raw = None
    for enc in ("gbk", "utf-8-sig", "utf-8"):
        try:
            with open(bas_path, "r", encoding=enc) as f:
                raw = f.read()
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise RuntimeError("读不出 %s 的编码（试过 gbk/utf-8）" % bas_path)
    lines = raw.split("\n")
    if lines and lines[0].startswith("Attribute VB_Name"):
        lines = lines[1:]
    return "\r\n".join(ln.rstrip("\r") for ln in lines)


# ---------------------------------------------------------------- 步骤 A
def build_vba_source(bas_path, out_xlsm):
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError("缺 pywin32（%s）——先跑 ensure_deps.py" % e)

    code = module_code(bas_path)
    log("A. 模块源码 %d 字符，来自 %s" % (len(code), os.path.basename(bas_path)))
    log("   AccessVBOM 设为 1，读回 = %s" % set_access_vbom(1))
    pythoncom.CoInitialize()
    xl = None
    ok = False
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.Interactive = False
        xl.EnableEvents = False

        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < 2:      # 台账有 2 个表，文档模块数要对得上
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))

        vp = wb.VBProject
        comp = vp.VBComponents.Add(1)       # 1 = 标准模块
        comp.Name = "modLawPanel"
        comp.CodeModule.AddFromString(code)
        tw = vp.VBComponents("ThisWorkbook")
        tw.CodeModule.AddFromString(THISWB_CODE)
        log("   注入完成：modLawPanel %d 行 / ThisWorkbook %d 行"
            % (comp.CodeModule.CountOfLines, tw.CodeModule.CountOfLines))

        # —— UserForm：窗内 TextBox 是独立控件，滚轮才有人接 ——
        frm = vp.VBComponents.Add(3)          # 3 = vbext_ct_MSForm
        frm.Name = "frmLaw"
        frm.Designer.Caption = "法条全文"
        _tb = frm.Designer.Controls.Add("Forms.TextBox.1", "txtLaw", True)
        _tb.Left = 2
        _tb.Top = 2
        _tb.Width = 400
        _tb.Height = 500
        _tb.MultiLine = True
        _tb.WordWrap = True
        _tb.ScrollBars = 2                    # fmScrollBarsVertical
        _tb.Font.Name = "Microsoft YaHei"
        _tb.Font.Size = 10
        _tb.BackColor = 16777164              # 浅黄
        frm.CodeModule.AddFromString(FORM_CODE)
        log("   UserForm 已建：frmLaw / txtLaw %dx%d ScrollBars=%d"
            % (_tb.Width, _tb.Height, _tb.ScrollBars))

        # 编译校验：AddFromString 不做完整编译，拼错的 Sub 名/非法语句要到运行时才炸。
        # 跑一次只打日志、不弹窗的 SelfTest，能把「整工程编译失败」提前到构建期暴露。
        try:
            xl.Run("modLawPanel.SelfTest")
            log("   编译校验：SelfTest 运行成功（整工程可编译）")
        except Exception as e:
            log("   编译校验失败（这个文件不要交付）:", str(e)[:200])
            raise

        if os.path.exists(out_xlsm):
            os.remove(out_xlsm)
        wb.SaveAs(out_xlsm, 52)             # 52 = xlOpenXMLWorkbookMacroEnabled
        # Close 常被拒（0x80010001），不影响已落盘文件，以文件为准
        try:
            wb.Close(SaveChanges=False)
        except Exception as e:
            log("   Close 被拒（忽略，文件已落盘）:", str(e)[:60])
        ok = os.path.getsize(out_xlsm) > 4096
        log("   宏源已生成：%s  %d 字节" % (os.path.basename(out_xlsm), os.path.getsize(out_xlsm)))
    finally:
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        log("   AccessVBOM 改回 0，读回 = %s" % set_access_vbom(0))
    return ok


# ---------------------------------------------------------------- 步骤 B
def transplant(src_xlsx, dst_xlsm, vba_src_xlsm):
    with zipfile.ZipFile(vba_src_xlsm) as z:
        vba_bin = z.read("xl/vbaProject.bin")
        ct_src = z.read("[Content_Types].xml").decode("utf-8")
        rels_src = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    # 从宏源学它自己的写法（Excel 用 <Default Extension="bin">，不是 Override）
    m_ct = (re.search(r'<Override[^>]*vbaProject\.bin[^>]*/>', ct_src)
            or re.search(r'<Default[^>]*Extension="bin"[^>]*/>', ct_src))
    m_wbct = re.search(r'PartName="/xl/workbook\.xml" ContentType="([^"]+)"', ct_src)
    m_rel = re.search(r'<Relationship[^>]*vbaProject[^>]*/>', rels_src)
    if not (m_ct and m_wbct and m_rel):
        raise RuntimeError("宏源结构不符合预期，停止移植")
    log("B. vbaProject.bin %d 字节" % len(vba_bin))
    log("   workbook 主类型 ->", m_wbct.group(1))

    with zipfile.ZipFile(src_xlsx) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'(<Override PartName="/xl/workbook\.xml" ContentType=")[^"]+(")',
                r"\g<1>%s\g<2>" % m_wbct.group(1), ct)
    if "vbaProject" not in ct:
        ct = ct.replace("</Types>", m_ct.group(0) + "</Types>")
    data["[Content_Types].xml"] = ct.encode("utf-8")

    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "vbaProject" not in rels:
        used = set(re.findall(r'Id="(rId\d+)"', rels))
        n = 1
        while "rId%d" % n in used:
            n += 1
        src_id = re.search(r'Id="([^"]+)"', m_rel.group(0)).group(1)
        new_rel = m_rel.group(0).replace('Id="%s"' % src_id, 'Id="rId%d"' % n)
        rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
        data["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    data["xl/vbaProject.bin"] = vba_bin

    if os.path.exists(dst_xlsm):
        # 备份必须放**库外**：库根有白名单（verify_library C10），
        # 直接在库根留 .bak 会被判「非白名单条目」。
        bak_dir = os.path.join(os.path.dirname(os.path.abspath(src_xlsx)),
                               "..", "_案例库维护留痕", "台账备份")
        bak_dir = os.path.abspath(bak_dir)
        os.makedirs(bak_dir, exist_ok=True)
        bak = os.path.join(bak_dir, "%s.bak_%d" % (os.path.basename(dst_xlsm), int(time.time())))
        shutil.move(dst_xlsm, bak)
        log("   旧宏版已备份 -> %s" % bak)

    with zipfile.ZipFile(dst_xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm in list(names) + ["xl/vbaProject.bin"]:
            zout.writestr(nm, data[nm])
    log("   移植完成：%s  %d 字节" % (os.path.basename(dst_xlsm), os.path.getsize(dst_xlsm)))


def verify_pair(xlsx, xlsm):
    """立刻自证：四个部件必须逐字节相同，否则这批产物不能用。"""
    import hashlib
    bad = []
    with zipfile.ZipFile(xlsx) as zx, zipfile.ZipFile(xlsm) as zm:
        names = set(zm.namelist())
        if "xl/vbaProject.bin" not in names:
            raise RuntimeError("产物里没有 vbaProject.bin")
        for p in PARTS_TO_COMPARE:
            if p not in names:
                continue
            if hashlib.sha256(zx.read(p)).hexdigest() != hashlib.sha256(zm.read(p)).hexdigest():
                bad.append(p)
    if bad:
        raise RuntimeError("宏版与正本不一致：%s" % ", ".join(bad))
    log("   自证通过：工作表/批注/批注 VML 与正本逐字节相同，宏部件已就位")


def main():
    ap = argparse.ArgumentParser(description="给台账正本生成宏版孪生体 .xlsm")
    ap.add_argument("--root", required=True, help="案例库根目录")
    ap.add_argument("--bas", default=None,
                    help="VBA 模块文件（默认用技能自带 references/modLawPanel.bas）")
    ap.add_argument("--vba-src", default=None,
                    help="复用已建好的宏源 .xlsm（给了就不启动 Excel，本机 Excel 不可用时的后路）")
    ap.add_argument("--xlsx-name", default=LEDGER_BASENAME + ".xlsx")
    ap.add_argument("--xlsm-name", default=LEDGER_BASENAME + ".xlsm")
    a = ap.parse_args()

    root = os.path.abspath(a.root)
    xlsx = os.path.join(root, a.xlsx_name)
    xlsm = os.path.join(root, a.xlsm_name)
    if not os.path.isfile(xlsx):
        log("找不到台账正本：%s" % xlsx)
        return 1

    bas = a.bas or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "references", "modLawPanel.bas")
    bas = os.path.abspath(bas)

    tmpdir = None
    vba_src = a.vba_src
    try:
        if vba_src:
            log("A. 复用宏源：%s" % vba_src)
            if not os.path.isfile(vba_src):
                log("   宏源不存在：%s" % vba_src)
                return 1
        else:
            if not os.path.isfile(bas):
                log("找不到 VBA 模块：%s" % bas)
                return 1
            tmpdir = tempfile.mkdtemp(prefix="macrowb_")
            vba_src = os.path.join(tmpdir, "vba_source.xlsm")
            if not build_vba_source(bas, vba_src):
                log("步骤 A 失败：Excel 未能生成宏源。")
                log("  若本机 Excel 不可用，可在有 Excel 的机器上跑一次，再用 --vba-src 复用。")
                return 1

        transplant(xlsx, xlsm, vba_src)
        verify_pair(xlsx, xlsm)
    except Exception as e:
        log("失败：%s %s" % (type(e).__name__, e))
        return 1
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)

    log("完成：%s" % xlsm)
    log("提示：正本更新后必须重跑本脚本，否则 verify_library 的 C10b 会报「宏版已过期」。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
