Attribute VB_Name = "modLawPanel"
Option Explicit

' =====================================================================
'  H 列法条面板：鼠标**停留**即出现，可上下滚动读全文
'
'  为什么必须用轮询：Excel 没有「鼠标悬停单元格」事件。
'  SelectionChange 只在**点击**时触发，所以上一版是「点一下才弹」，
'  不是用户要的「停下就弹」。这里改用 Application.OnTime 每 250ms
'  轮询一次鼠标坐标，用 RangeFromPoint 反查指针下的单元格。
'
'  安全性：只读批注文本，不修改任何单元格、不修改批注、不写台账文件。
'  只在内存里创建一个浮动文本框，关闭工作簿即消失。
'  不依赖任何系统设置（不建窗体、不碰 VBProject 信任）。
' =====================================================================

#If VBA7 Then
    Private Type POINTAPI
        x As Long
        y As Long
    End Type
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#Else
    Private Type POINTAPI
        x As Long
        y As Long
    End Type
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
#End If

' ---- 可调参数 ------------------------------------------------------
' **间隔不能小于 1000ms**：实测 250ms 时 Application.OnTime 会退化成「立即执行」，
' 1.8MB 日志 / 750 次每秒空转把 Excel 打满（2026-11-04 实测，可见/隐藏实例都一样）。
' Excel 的 OnTime 粒度就是 1 秒，小于 1 秒的时间点一律当「已到期」立刻执行 → 自旋。
Private Const POLL_MS As Long = 1000       ' 轮询间隔（毫秒），最小 1000
Private Const DWELL_POLLS As Long = 1      ' 同一格连续命中 N 次才弹（配合 POLL_MS=1s → 停 1 秒弹）
Private Const LEAVE_POLLS As Long = 2      ' 指针离开 N 次（=2s）才收，防误收
Private Const MOVE_TOLERANCE As Long = 3   ' 指针位移超过这么多像素就重新计时
Private Const AUTO_CLOSE_SEC As Long = 90  ' 面板最长驻留秒数，到点自动收
' 指针在面板矩形内 → 视为「还在看」，不收。
' 若你的机器上出现「滚动时面板消失」，把它改成 False（退化为只看格子）。
Private Const KEEP_ON_PANEL As Boolean = True

Private Const PANEL_NAME As String = "LawPanel"
' 窗体是否已弹出（UserForm 模式）
Public gPanelUp As Boolean

Private Type RECT_T
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT_T) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT_T) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If


' 批注长度阈值：短批注原生悬停就能看全，不必弹面板
Private Const AUTO_MIN_CHARS As Long = 1400

' 面板尺寸（像素）
Private Const PANEL_W As Double = 640
Private Const PANEL_MAX_H As Double = 760

' ---- 运行状态 ------------------------------------------------------
Private gRunning As Boolean
Private gHoverKey As String        ' 当前指针停留的 H 格 "表名!行号"
Private gDwell As Long
Private gLeave As Long
Private gShownKey As String        ' 当前已弹出面板对应的 "表名!行号"
Private gShownSheet As Worksheet
Private gShownAt As Double         ' 弹出时刻（Timer 秒）
Private gTicks As Long             ' 轮询次数（自检用）
Private gLastX As Long             ' 上一 tick 的指针坐标（判断"停着没动"）
Private gLastY As Long
Private gLastAt As Double          ' 上一 tick 的时刻（防自旋保险用）
Public gDebug As Boolean           ' True 时把每次 tick 写日志


''' <summary>启动轮询（Workbook_Open 调用）。重复调用无副作用。</summary>
Public Sub StartHover()
    ' 自愈：gRunning 可能是"上一次运行留下的真"，但定时器早就死了（宏被重置 /
    ' OnTime 丢约 / 用户中断）。判据＝超过 5 秒没 tick 就重新约，不再盲信标志位。
    If gRunning Then
        If gLastAt > 0 Then
            If (Timer - gLastAt) < 5 Then Exit Sub
        End If
    End If
    gRunning = True
    gTicks = 0
    gLastAt = 0
    LogLine "StartHover POLL=" & POLL_MS & "ms  file=" & ThisWorkbook.Name
    On Error Resume Next
    Application.StatusBar = "【法条悬停已启用】把鼠标停在 H 列某格上约 1 秒，左侧弹出可滚动面板"
    Application.OnTime EarliestTime:=Now + POLL_MS / 86400000#, _
                       Procedure:="modLawPanel.HoverTick"
End Sub


''' <summary>停止轮询并收面板（BeforeClose / Deactivate 调用）。</summary>
Public Sub StopHover()
    gRunning = False
    On Error Resume Next
    Application.StatusBar = False
    gDwell = 0
    gLeave = 0
    gHoverKey = ""
    On Error Resume Next
    If Not gShownSheet Is Nothing Then HideLawPanel gShownSheet
    Set gShownSheet = Nothing
    gShownKey = ""
End Sub


''' <summary>轮询主体：每 POLL_MS 跑一次，跑完自己续约。</summary>
Public Sub HoverTick()
    On Error Resume Next

    If Not gRunning Then Exit Sub

    ' —— 防自旋保险（2026-11-04 实测事故）——
    ' OnTime 一旦把 POLL_MS 当"已到期"处理，就会变成每秒几百次的死循环，把 Excel 打满。
    ' 这里量一下两次 tick 的真实间隔：小于 POLL_MS 的一半，就当作异常，
    ' 直接拉长到 1 秒再续约，并且这一轮什么也不做。
    Dim gap As Double
    gap = Timer - gLastAt
    If gLastAt > 0 Then
        If gap < POLL_MS / 2000# And gap >= 0 Then
            gLastAt = Timer
            Application.OnTime EarliestTime:=Now + 1 / 86400#, _
                               Procedure:="modLawPanel.HoverTick"
            Exit Sub
        End If
    End If
    gLastAt = Timer

    gTicks = gTicks + 1

    ' —— 窗体已经弹出：只判断鼠标还在不在窗体上，离开就收 ——
    ' 这一段必须放在"仅 Excel 为前台才处理"之前：窗体一弹出 Excel 就不是前台了，
    ' 放后面会导致窗体永远关不掉。
    If gPanelUp Then
        If Not PointerOnPanel() Then
            If Timer - frmLaw.LastMove > 0.8 Then HideLawPanel Application.ActiveSheet
        Else
            frmLaw.LastMove = Timer
        End If
        Err.Clear
        If gRunning Then
            Application.OnTime EarliestTime:=Now + POLL_MS / 86400000#, _
                               Procedure:="modLawPanel.HoverTick"
        End If
        Exit Sub
    End If

    If gDebug Then LogLine "tick " & gTicks

    ' 只在 Excel 是前台窗口、且不在编辑状态时才判断，避免打断输入
    If Application.Interactive Then
        If Application.WindowState <> xlMinimized Then
            If GetForegroundWindow() = Application.Hwnd Then
                If Application.Ready Then
                    InspectPointer
                End If
            End If
        End If
    End If

    ' 超时自动收
    If Len(gShownKey) > 0 Then
        If Timer - gShownAt > AUTO_CLOSE_SEC Then
            If Not gShownSheet Is Nothing Then HideLawPanel gShownSheet
            Set gShownSheet = Nothing
            gShownKey = ""
        End If
    End If

    Err.Clear
    If gRunning Then
        Application.OnTime EarliestTime:=Now + POLL_MS / 86400000#, _
                           Procedure:="modLawPanel.HoverTick"
    End If
End Sub


''' <summary>看指针此刻在哪：在 H 格上停够时间就弹，离开就收。</summary>
Private Sub InspectPointer()
    Dim pt As POINTAPI
    Dim obj As Object
    Dim cell As Range
    Dim key As String
    Dim txt As String

    If GetCursorPos(pt) = 0 Then Exit Sub

    ' 指针动了就重新计时——"停 1 秒"才算停留，手划过不弹
    If Abs(pt.x - gLastX) > MOVE_TOLERANCE Or Abs(pt.y - gLastY) > MOVE_TOLERANCE Then
        gDwell = 0
        gHoverKey = ""
        gLastX = pt.x
        gLastY = pt.y
    End If

    ' ① 指针压在自己弹出的面板上 → 保持不动（否则滚轮滚到一半面板就没了）
    If KEEP_ON_PANEL Then
        If PointInPanel(pt.x, pt.y) Then
            gLeave = 0
            gDwell = 0
            Exit Sub
        End If
    End If

    ' ② 反查指针下的对象
    Set obj = Nothing
    On Error Resume Next
    Set obj = Application.ActiveWindow.RangeFromPoint(pt.x, pt.y)
    On Error GoTo 0
    If obj Is Nothing Then
        CountLeave
        Exit Sub
    End If
    If Not TypeOf obj Is Range Then
        ' 落在图形 / OLE 上（含本面板）→ 不算离开
        gLeave = 0
        Exit Sub
    End If

    Set cell = obj.Cells(1, 1)

    ' ③ 不是「有长批注的 H 格」→ 记一次离开
    If cell.Column <> 8 Then
        gDwell = 0
        CountLeave
        Exit Sub
    End If
    If cell.Comment Is Nothing Then
        gDwell = 0
        CountLeave
        Exit Sub
    End If
    txt = cell.Comment.Text
    If Len(txt) <= AUTO_MIN_CHARS Then
        gDwell = 0
        CountLeave
        Exit Sub
    End If

    ' ④ 停在同一个 H 格上累计次数
    key = cell.Worksheet.Name & "!" & cell.Row
    If key <> gHoverKey Then
        gHoverKey = key
        gDwell = 0
        LogLine "hover H" & cell.Row & " 批注字数=" & Len(txt) & " 待弹=" & (Len(txt) > AUTO_MIN_CHARS)
    End If
    gDwell = gDwell + 1
    gLeave = 0

    ' ⑤ 停够时间且还没弹过 → 弹
    If key <> gShownKey Then
        If gDwell >= DWELL_POLLS Then
            ShowLawPanel cell.Worksheet, txt, cell.Address(False, False)
            On Error Resume Next
            Application.StatusBar = "已弹出 " & cell.Address(False, False) & " 的法条面板（左侧，滚轮可翻）"
            gShownKey = key
            Set gShownSheet = cell.Worksheet
            gShownAt = Timer
        End If
    End If
End Sub


''' <summary>累计离开次数，够了就收面板。</summary>
Private Sub CountLeave()
    If Len(gShownKey) = 0 Then Exit Sub
    gLeave = gLeave + 1
    If gLeave >= LEAVE_POLLS Then
        On Error Resume Next
        If Not gShownSheet Is Nothing Then HideLawPanel gShownSheet
        Set gShownSheet = Nothing
        gShownKey = ""
        gLeave = 0
    End If
End Sub


''' <summary>指针是否落在面板矩形内（磅 → 屏幕像素，按窗口文档区原点换算）。</summary>
Private Function PointInPanel(ByVal sx As Long, ByVal sy As Long) As Boolean
    Dim ob As OLEObject
    Dim w As Window
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long

    If gShownSheet Is Nothing Then Exit Function
    On Error Resume Next
    Set ob = gShownSheet.OLEObjects(PANEL_NAME)
    If ob Is Nothing Then Exit Function
    Set w = Application.ActiveWindow
    If w Is Nothing Then Exit Function

    x1 = w.PointsToScreenPixelsX(ob.Left)
    y1 = w.PointsToScreenPixelsY(ob.Top)
    x2 = w.PointsToScreenPixelsX(ob.Left + ob.Width)
    y2 = w.PointsToScreenPixelsY(ob.Top + ob.Height)

    PointInPanel = (sx >= x1 And sx <= x2 And sy >= y1 And sy <= y2)
    On Error GoTo 0
End Function


''' <summary>右侧弹出一个可滚动的法条面板。</summary>
Public Sub ShowLawPanel(ByVal Sh As Worksheet, ByVal Txt As String, ByVal Addr As String)
    ' 2026-11-04 第十二版：改用 UserForm 独立窗口。
    ' 嵌入在表上的 ActiveX 文本框拿不到焦点 → 滚轮滚的是工作表（用户实测"只能拖不能滚"）。
    ' 独立窗体是顶层窗口，窗内 TextBox 有焦点，滚轮才会被它接收。
    On Error GoTo Failed
    With frmLaw
        .Caption = "法条全文  " & Addr & "   ·   滚轮上下翻 · 鼠标移开自动关"
        .Width = 650
        .Height = 700
        .txtLaw.Width = 620
        .txtLaw.Height = 630
        .txtLaw.Text = Txt
        .LastMove = Timer
        .Show vbModeless
        .txtLaw.SetFocus
        .LastMove = Timer
    End With
    gPanelUp = True
    gShownAt = Timer
    LogLine "FormOK " & Addr & " 字数=" & Len(Txt)
    Exit Sub
Failed:
    gPanelUp = False
    LogLine "ShowLawPanel FAIL " & Err.Number & " (" & Err.Description & ")"
End Sub


''' <summary>移除面板（拖等，没有也不报错）。</summary>
Public Sub HideLawPanel(ByVal Sh As Worksheet)
    On Error Resume Next
    If gPanelUp Then
        Unload frmLaw
        gPanelUp = False
    End If
    Err.Clear
End Sub


''' <summary>鼠标当前是否落在窗体矩形内（窗体是独立顶层窗口，类名 ThunderDFrame）。</summary>
Private Function PointerOnPanel() As Boolean
    Dim p As POINTAPI
    Dim r As RECT_T
    Dim h As LongPtr
    On Error Resume Next
    If Not gPanelUp Then Exit Function
    #If VBA7 Then
        h = FindWindow("ThunderDFrame", vbNullString)
    #Else
        h = FindWindow("ThunderDFrame", vbNullString)
    #End If
    If h = 0 Then Exit Function
    If GetCursorPos(p) = 0 Then Exit Function
    If GetWindowRect(h, r) = 0 Then Exit Function
    PointerOnPanel = (p.x >= r.Left And p.x <= r.Right And p.y >= r.Top And p.y <= r.Bottom)
End Function


''' <summary>手动弹当前格（宏列表里可直接点的入口）。</summary>
Public Sub ShowCurrentCellLaw()
    Dim c As Range
    Set c = ActiveCell
    If c.Comment Is Nothing Then
        MsgBox "当前单元格没有批注。", vbInformation
        Exit Sub
    End If
    ShowLawPanel ActiveSheet, c.Comment.Text, c.Address(False, False)
End Sub


''' <summary>手动收起（宏列表入口）。</summary>
Public Sub HideCurrentPanel()
    On Error Resume Next
    HideLawPanel ActiveSheet
    If Not gShownSheet Is Nothing Then HideLawPanel gShownSheet
    Set gShownSheet = Nothing
    gShownKey = ""
End Sub


''' <summary>点击兜底：选中 H 格时也弹（不改原逻辑，防轮询被禁时全无反应）。</summary>
Public Sub OnSelectCell(ByVal Sh As Worksheet, ByVal Target As Range)
    Dim txt As String
    On Error Resume Next
    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub
    If Target.Column <> 8 Then Exit Sub
    If Target.Comment Is Nothing Then Exit Sub
    txt = Target.Comment.Text
    If Len(txt) <= AUTO_MIN_CHARS Then Exit Sub
    ShowLawPanel Sh, txt, Target.Address(False, False)
    gShownKey = Sh.Name & "!" & Target.Row
    Set gShownSheet = Sh
    gShownAt = Timer
End Sub


' =====================================================================
'  自检与日志（交付前用得到，正常使用不看也不影响）
' =====================================================================


' =====================================================================
'  一键诊断：在真 Excel 里跑一次，把真实环境写进日志
'  用法：Alt+F8 → 选 modLawPanel.DiagAll → 执行
' =====================================================================
Public Sub DiagAll()
    Dim pt As POINTAPI
    Dim s As String
    Dim ob As OLEObject
    Dim ws As Worksheet
    Dim r As Object
    Dim t0 As Double
    Dim t1 As Double
#If VBA7 Then
    Dim fg As LongPtr
#Else
    Dim fg As Long
#End If

    On Error Resume Next
    Set ws = ActiveSheet

    s = "=== DiagAll " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    s = s & vbCrLf & "File=" & ThisWorkbook.Name
    s = s & vbCrLf & "Version=" & Application.Version & "  Build=" & Application.Build
    s = s & vbCrLf & "Sheet=" & ws.Name
    s = s & vbCrLf & "HoverRunning=" & gRunning & "  Ticks=" & gTicks
    s = s & vbCrLf & "EnableEvents=" & Application.EnableEvents
    s = s & vbCrLf & "Interactive=" & Application.Interactive
    s = s & vbCrLf & "Win UsableW/H=" & Format$(Application.ActiveWindow.UsableWidth, "0") _
        & "/" & Format$(Application.ActiveWindow.UsableHeight, "0") _
        & "  Zoom=" & Application.ActiveWindow.Zoom

    If GetCursorPos(pt) <> 0 Then
        s = s & vbCrLf & "GetCursorPos=OK " & pt.x & "," & pt.y
    Else
        s = s & vbCrLf & "GetCursorPos=FAIL"
    End If

    ' ① 面板能不能建
    HideLawPanel ws
    t0 = Timer
    ShowLawPanel ws, String$(600, "测"), "DIAG"
    s = s & vbCrLf & "CreatePanel ms=" & Format$(Int((Timer - t0) * 1000), "0")
    s = s & vbCrLf & "OLEObjects=" & ws.OLEObjects.Count
    Set ob = Nothing
    Set ob = ws.OLEObjects(PANEL_NAME)
    If Not ob Is Nothing Then
        s = s & vbCrLf & "Panel L/T/W/H=" & Format$(ob.Left, "0.0") & "/" & _
            Format$(ob.Top, "0.0") & "/" & Format$(ob.Width, "0.0") & "/" & _
            Format$(ob.Height, "0.0")
        s = s & vbCrLf & "Panel Visible=" & ob.Visible & "  ScrollBars=" & ob.Object.ScrollBars
    Else
        s = s & vbCrLf & "Panel=未创建（ShowLawPanel 失败，Err 被吞）"
    End If

    ' ② 指针能不能反查到格子
    Set r = Nothing
    Set r = Application.ActiveWindow.RangeFromPoint(pt.x, pt.y)
    If r Is Nothing Then
        s = s & vbCrLf & "RangeFromPoint=Nothing"
    ElseIf TypeOf r Is Range Then
        s = s & vbCrLf & "RangeFromPoint=Range " & r.Address(False, False) & " col=" & r.Column
    Else
        s = s & vbCrLf & "RangeFromPoint=Shape/其他对象"
    End If

    ' ③ 前台判定（悬停逻辑的开关）
    fg = GetForegroundWindow()
    s = s & vbCrLf & "Foreground=" & fg & "  AppHwnd=" & Application.Hwnd & _
        "  Same=" & (fg = Application.Hwnd)

    LogLine s
    HideLawPanel ws
    t1 = Timer
    If gDebug Then LogLine "DiagAll 用时 " & Format$(Int((t1 - t0) * 1000), "0") & "ms"
    MsgBox "诊断已写入日志：" & vbCrLf & Environ$("TEMP") & "\lawpanel.log" & vbCrLf & vbCrLf & _
           "把这个文件发我（或把内容贴给我）即可。", vbInformation
End Sub

''' <summary>把环境能力写进日志：API 能否调用、面板能否创建。</summary>
Public Sub SelfTest()
    Dim pt As POINTAPI
    Dim s As String

    On Error Resume Next
    s = "=== SelfTest " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="

    If GetCursorPos(pt) <> 0 Then
        s = s & vbCrLf & "GetCursorPos=OK " & pt.x & "," & pt.y
    Else
        s = s & vbCrLf & "GetCursorPos=FAIL"
    End If
    s = s & vbCrLf & "ForegroundWindow=" & GetForegroundWindow()
    s = s & vbCrLf & "App.Hwnd=" & Application.Hwnd
    s = s & vbCrLf & "OLEObjects=" & ActiveSheet.OLEObjects.Count
    If ActiveSheet.OLEObjects.Count > 0 Then
        s = s & vbCrLf & "Panel=" & ActiveSheet.OLEObjects(PANEL_NAME).Width & "x" & _
            ActiveSheet.OLEObjects(PANEL_NAME).Height & " pt"
        s = s & vbCrLf & "ScrollBars=" & ActiveSheet.OLEObjects(PANEL_NAME).Object.ScrollBars
        s = s & vbCrLf & "PanelTextLen=" & Len(ActiveSheet.OLEObjects(PANEL_NAME).Object.Text)
    End If
    s = s & vbCrLf & "VBA7=" & CStr(IsVBA7())
    LogLine s
End Sub


''' <summary>跑一轮轮询自检：启动 → 数 tick → 停止，结果写日志。</summary>
Public Sub DebugOn()
    gDebug = True
    LogLine "=== DebugOn ==="
End Sub


Public Sub TickTest()
    gDebug = True
    LogLine "=== TickTest 开始 ==="
    StartHover
End Sub


Public Sub TickStop()
    LogLine "=== TickStop tick=" & gTicks & " ==="
    gDebug = False
    StopHover
End Sub


Private Function IsVBA7() As Boolean
#If VBA7 Then
    IsVBA7 = True
#Else
    IsVBA7 = False
#End If
End Function


''' <summary>追加一行到 %TEMP%\lawpanel.log（路径写死会因机器不同失败，故用 TEMP）。</summary>
Public Sub LogLine(ByVal Msg As String)
    Dim f As Integer
    On Error Resume Next
    If FileLen(Environ$("TEMP") & "\lawpanel.log") > 300000 Then
        Kill Environ$("TEMP") & "\lawpanel.log"
    End If
    f = FreeFile
    Open Environ$("TEMP") & "\lawpanel.log" For Append As #f
    Print #f, Msg
    Close #f
End Sub

''' <summary>【一次定论】打开工作簿就建一个大号可滚文本框。
' 鼠标移进去滚轮能滚 = Excel 里这条路成立；滚不动 = 这条路也不通。
' 不依赖悬停轮询、不依赖 OnTime，只依赖「宏能不能跑」。</summary>
Public Sub ShowScrollTest()
    Dim ob As OLEObject
    Dim vis As Range
    Dim pw As Double, ph As Double, px As Double, py As Double
    Dim txt As String
    Dim i As Long
    Dim st As String
    Dim gotFocus As Boolean

    On Error GoTo Failed

    st = "VisibleRange"
    Set vis = Application.ActiveWindow.VisibleRange

    st = "size"
    pw = 640
    If pw > vis.Width * 0.45 Then pw = vis.Width * 0.45
    ph = vis.Height - 40
    If ph > 640 Then ph = 640
    If ph < 320 Then ph = 320
    px = vis.Left + vis.Width - pw - 24
    If px < vis.Left Then px = vis.Left
    py = vis.Top + 20

    st = "remove old"
    HideLawPanel Application.ActiveSheet

    st = "add"
    Set ob = Application.ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", _
        Left:=px, Top:=py, Width:=pw, Height:=ph)
    ob.Name = PANEL_NAME
    ob.Placement = 3
    ob.PrintObject = False

    st = "props"
    ob.Object.MultiLine = True
    ob.Object.WordWrap = True
    ob.Object.ScrollBars = 2
    ob.Object.Font.Name = "Microsoft YaHei"
    ob.Object.Font.Size = 10
    ob.Object.ForeColor = 0
    ob.Object.BackColor = 16777164

    txt = "【滚动测试】把鼠标箭头移到我这个框里面，滚轮上下滚。" & vbCrLf & _
          "内容动了 = Excel 里做得到。一动不动 = 这条路也不行。" & vbCrLf & _
          "点表格任一格可以关掉我。" & vbCrLf & String(56, "-") & vbCrLf
    For i = 1 To 100
        txt = txt & "第 " & i & " 行：滚动我，看我动不动。" & vbCrLf
    Next i

    st = "settext"
    ob.Object.Text = txt

    st = "focus"
    On Error Resume Next
    ob.Activate
    ob.Object.SetFocus
    gotFocus = (Err.Number = 0)
    Err.Clear
    On Error GoTo Failed

    LogLine "ScrollTest OK L=" & ob.Left & " T=" & ob.Top & " W=" & ob.Width & _
            " H=" & ob.Height & " ScrollBars=" & ob.Object.ScrollBars & _
            " 焦点=" & gotFocus & " 字数=" & Len(ob.Object.Text)
    Exit Sub

Failed:
    LogLine "ScrollTest FAIL " & Err.Number & " (" & Err.Description & ") @ " & st
End Sub
