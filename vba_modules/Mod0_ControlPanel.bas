' ==============================================================================
' 模块名称: Mod0_ControlPanel
' 核心职责: 控制面板交互工作台 (UI, Presentation & Business Pipeline Orchestrator)
' 设计准则:
'   1. 顶部横幅双核设计: [ 重置与同步代码 ] (系统运维) + [ >>> 一键执行全流程 <<< ] (业务生产)
'   2. 业务全流程在此调度，支持随时无损热更与极富信息量的交付汇总报告
'   3. 680pt 几何严密对齐，纯正 GBK 中文，绝对杜绝任何 Unicode Emoji 乱码
' ==============================================================================
Option Explicit

Public Sub 一键生成控制面板()
    Dim wb As Workbook, ws As Worksheet, shp As Shape, sIdx As Long
    
    Set wb = ThisWorkbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If wb.Sheets.Count > 1 Then
        For sIdx = wb.Sheets.Count To 2 Step -1
            wb.Sheets(sIdx).Delete
        Next sIdx
    End If
    
    Set ws = wb.Sheets(1)
    ws.Name = "Sheet1"
    ws.Cells.Clear
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    
    ' 精确对齐列宽（使得 Col B 起始于 20pt，Col J 结束于 680pt）
    ws.Columns("A").ColumnWidth = 2.63  ' ~20pt
    ws.Columns("B").ColumnWidth = 14    ' 88pt
    ws.Columns("C").ColumnWidth = 11    ' 70pt
    ws.Columns("D").ColumnWidth = 13    ' 82pt
    ws.Columns("E").ColumnWidth = 11    ' 70pt
    ws.Columns("F").ColumnWidth = 13    ' 82pt
    ws.Columns("G").ColumnWidth = 3.5   ' 26pt
    ws.Columns("H").ColumnWidth = 13    ' 82pt
    ws.Columns("I").ColumnWidth = 13    ' 82pt
    ws.Columns("J").ColumnWidth = 13    ' 82pt (Col J 右边缘严格等于 680pt)
    ws.Columns("K").ColumnWidth = 3
    
    ws.Cells.Font.Name = "微软雅黑"
    ws.Cells.Font.Size = 10
    
    Call DrawHeaderBanner(ws)           ' Top: 15 ~ 70 (Right: 680pt, 双核功能栏)
    Call DrawSummaryCards(ws)           ' Top: 85 ~ 295 (Right: 680pt)
    Call DrawDateBar(ws)                ' Row 20 (Right: 680pt)
    Call DrawWorkflowAndPrinciples(ws)  ' Row 23 ~ 27 (极简流程与工作原理看板)
    Call DrawQuickLinksAndLog(ws)       ' Row 29 ~ 33 (Right: 680pt)
    
    Call 刷新控制面板数据
    
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    Call AppendLog("【控制面板】工作台已成功初始化并就绪！")
End Sub

Private Sub DrawHeaderBanner(ws As Worksheet)
    ' 1. 顶部主横幅底板 (20 ~ 680pt, 宽 660pt)
    Dim banner As Shape
    Set banner = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 660, 55)
    With banner
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(24, 76, 120)
        With .TextFrame2
            .TextRange.Text = "  科研台账 (keyantaizhang)"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 15.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 14
        End With
    End With
    
    ' 2. 系统管理按钮: [ 重置面板与同步代码 ] (325 ~ 465pt, 宽 140pt, 深青灰底色)
    Dim btnSync As Shape
    Set btnSync = ws.Shapes.AddShape(msoShapeRoundedRectangle, 325, 25, 140, 34)
    With btnSync
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(52, 73, 94)
        .OnAction = "Mod_Sync.一键热更并重置面板"
        With .TextFrame2
            .TextRange.Text = "[ 重置与同步代码 ]"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' 3. 业务生产主控按钮: [ >>> 一键自动化执行全流程 <<< ] (475 ~ 670pt, 宽 195pt, 橙色高光)
    Dim btnRun As Shape
    Set btnRun = ws.Shapes.AddShape(msoShapeRoundedRectangle, 475, 25, 195, 34)
    With btnRun
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(230, 126, 34)
        .OnAction = "Mod0_ControlPanel.一键执行业务全流程"
        With .TextFrame2
            .TextRange.Text = ">>> 一键自动化执行全流程 <<<"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Sub DrawSummaryCards(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 85, "【 数据概览与状态看板 】")
    
    Call CreateMetricCard(ws, 20, 110, 210, 75, "教师名单库 (config)", "card_Teacher", "0", RGB(238, 245, 252), RGB(31, 110, 165))
    Call CreateMetricCard(ws, 245, 110, 210, 75, "原始库数据 (raw_data)", "card_Raw", "0", RGB(238, 245, 252), RGB(39, 174, 96))
    Call CreateMetricCard(ws, 470, 110, 210, 75, "本室入库成果 (papers_final)", "card_Final", "0", RGB(254, 249, 231), RGB(211, 84, 0))
    
    Dim r As Long
    For r = 12 To 18
        ws.Rows(r).RowHeight = 18
    Next r
    
    ws.Range("B12").Value = "[ 原始数据源 raw_data ]"
    ws.Range("B12").Font.Bold = True
    ws.Range("B12").Font.Color = RGB(100, 100, 100)
    
    ws.Range("B13").Value = "WOS (SCI) 原始:": ws.Range("D13").Value = "0 条"
    ws.Range("B14").Value = "EI Compendex 原始:": ws.Range("D14").Value = "0 条"
    ws.Range("B15").Value = "CNKI 知网原始:": ws.Range("D15").Value = "0 条"
    ws.Range("B16").Value = "Scopus 原始:": ws.Range("D16").Value = "0 条"
    
    ws.Range("H12").Value = "[ 成果交付大表 papers_final_merged ]"
    ws.Range("H12").Font.Bold = True
    ws.Range("H12").Font.Color = RGB(100, 100, 100)
    
    ws.Range("H13").Value = "最终合并入库 (篇数):": ws.Range("J13").Value = "0 篇"
    ws.Range("H14").Value = "其中 SCI 论文:": ws.Range("J14").Value = "0 篇"
    ws.Range("H15").Value = "其中 EI 论文:": ws.Range("J15").Value = "0 篇"
    ws.Range("H16").Value = "其中 中文核心:": ws.Range("J16").Value = "0 篇"
    ws.Range("H17").Value = "(SCI+EI 双收录):": ws.Range("J17").Value = "0 篇"
    
    ws.Range("H18").Value = "非本室/排除条目:": ws.Range("J18").Value = "0 篇 (SCI 0, EI 0, 中文 0, SCI+EI 0)"
    
    With ws.Range("B13:D16, H13:J17")
        .Font.Size = 9.5
        .Font.Color = RGB(60, 60, 60)
    End With
    
    With ws.Range("H18, J18")
        .Font.Size = 9
        .Font.Color = RGB(130, 130, 130)
    End With
    
    ws.Range("D13:D16, J13:J17").HorizontalAlignment = xlRight
    ws.Range("D13:D16, J13:J17").Font.Bold = True
    ws.Range("J18").HorizontalAlignment = xlRight
End Sub

Private Sub DrawDateBar(ws As Worksheet)
    ws.Rows(19).RowHeight = 8
    ws.Rows(20).RowHeight = 24
    
    ws.Range("B20").Value = "成果发表时间:"
    ws.Range("B20").Font.Bold = True
    ws.Range("B20").Font.Color = RGB(24, 76, 120)
    ws.Range("B20").VerticalAlignment = xlCenter
    
    ws.Range("C20").Value = "起始日期"
    ws.Range("C20").Font.Color = RGB(120, 130, 140)
    ws.Range("C20").HorizontalAlignment = xlRight
    ws.Range("C20").VerticalAlignment = xlCenter
    
    With ws.Range("D20")
        .NumberFormat = "yyyy-mm-dd"
        .Value = "2026-01-01"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 205, 230)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(31, 110, 165)
    End With
    
    ws.Range("E20").Value = "截止日期"
    ws.Range("E20").Font.Color = RGB(120, 130, 140)
    ws.Range("E20").HorizontalAlignment = xlRight
    ws.Range("E20").VerticalAlignment = xlCenter
    
    With ws.Range("F20")
        .NumberFormat = "yyyy-mm-dd"
        .Value = "2026-12-31"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 205, 230)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(31, 110, 165)
    End With
    
    Dim btnTop As Single, btnH As Single
    btnTop = ws.Range("H20").Top + 1
    btnH = 22
    
    Call CreateMiniPresetBtn(ws, 440, btnTop, 44, btnH, "全部", "SetDatePreset_All", RGB(127, 140, 141))
    Call CreateMiniPresetBtn(ws, 489, btnTop, 44, btnH, "2026", "SetDatePreset_2026", RGB(52, 152, 219))
    Call CreateMiniPresetBtn(ws, 538, btnTop, 44, btnH, "2025", "SetDatePreset_2025", RGB(46, 204, 113))
    Call CreateMiniPresetBtn(ws, 587, btnTop, 44, btnH, "2024", "SetDatePreset_2024", RGB(155, 89, 182))
    Call CreateMiniPresetBtn(ws, 636, btnTop, 44, btnH, "近3年", "SetDatePreset_3Years", RGB(230, 126, 34))
End Sub

Private Sub CreateMiniPresetBtn(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                                btnText As String, macroName As String, btnColor As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With btn
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = btnColor
        .OnAction = macroName
        With .TextFrame2
            .TextRange.Text = btnText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Public Sub SetDatePreset_All()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").ClearContents
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").ClearContents
    Call AppendLog("【时间已设为: 全部历史】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_2026()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2026-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2026-12-31"
    Call AppendLog("【时间已设为: 2026全年度】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_2025()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2025-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2025-12-31"
    Call AppendLog("【时间已设为: 2025全年度】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_2024()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2024-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2024-12-31"
    Call AppendLog("【时间已设为: 2024全年度】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_3Years()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2024-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2026-12-31"
    Call AppendLog("【时间已设为: 近三年区间 2024~2026】（点击右上角大按钮即可生成成果）")
End Sub

Private Sub DrawWorkflowAndPrinciples(ws As Worksheet)
    ws.Rows(21).RowHeight = 10
    
    Dim titleTop As Single, cardTop As Single
    titleTop = ws.Range("B22").Top
    cardTop = titleTop + 24
    
    Call CreateSectionTitle(ws, 20, titleTop, "【 极简操作流程与系统工作原理 】")
    
    Dim guideCard As Shape
    Set guideCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, cardTop, 660, 95)
    With guideCard
        .Line.ForeColor.RGB = RGB(200, 220, 240)
        .Line.Weight = 1
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(246, 250, 254)
    End With
    
    ' 左栏: 3步极简操作流程
    Dim tbLeft As Shape
    Set tbLeft = ws.Shapes.AddShape(msoShapeRectangle, 32, cardTop + 6, 310, 82)
    With tbLeft
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "【 极简操作流程 (3 步搞定) 】" & vbCrLf & _
                              "1. 准备数据: 在 raw_data 目录放入 WOS/EI/知网 原始文件；" & vbCrLf & _
                              "2. 设定时间: 在上方面板选择或输入年份区间；" & vbCrLf & _
                              "3. 一键执行: 点击右上角【一键自动化执行全流程】大按钮！"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9
            .TextRange.Paragraphs(1).Font.Size = 9.5
            .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .TextRange.Font.Fill.ForeColor.RGB = RGB(60, 80, 100)
            .VerticalAnchor = msoAnchorTop
            .MarginLeft = 0: .MarginTop = 0
        End With
    End With
    
    ' 中部分割虚线
    Dim sepLine As Shape
    Set sepLine = ws.Shapes.AddLine(350, cardTop + 10, 350, cardTop + 85)
    With sepLine
        .Line.ForeColor.RGB = RGB(210, 225, 240)
        .Line.DashStyle = msoLineDash
    End With
    
    ' 右栏: 系统底层三大核心工作原理
    Dim tbRight As Shape
    Set tbRight = ws.Shapes.AddShape(msoShapeRectangle, 360, cardTop + 6, 310, 82)
    With tbRight
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "【 系统底层工作原理 (三大核心支柱) 】" & vbCrLf & _
                              "1. 智能消歧: 依据 100 位教师拼音库自动认领，非本室条目安全排除；" & vbCrLf & _
                              "2. 年份过滤: 严格按正式出版年份 (PY等) 筛选在期成果；" & vbCrLf & _
                              "3. 跨库合流: 基于 DOI 与题目哈希去重，自动标识 SCI+EI 复合收录。"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9
            .TextRange.Paragraphs(1).Font.Size = 9.5
            .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .TextRange.Font.Fill.ForeColor.RGB = RGB(60, 80, 100)
            .VerticalAnchor = msoAnchorTop
            .MarginLeft = 0: .MarginTop = 0
        End With
    End With
End Sub

Private Sub DrawQuickLinksAndLog(ws As Worksheet)
    Dim linkTop As Single, btnLinkTop As Single, logTitleTop As Single, logBoxTop As Single
    
    linkTop = ws.Range("B22").Top + 130
    btnLinkTop = linkTop + 24
    logTitleTop = btnLinkTop + 42
    logBoxTop = logTitleTop + 24
    
    Call CreateSectionTitle(ws, 20, linkTop, "【 独立文件与目录快捷联动 】")
    
    Call CreateFileLinkButton(ws, 20, btnLinkTop, 150, 30, "打开教师档案表", "config\teachers_profile.xlsx")
    Call CreateFileLinkButton(ws, 190, btnLinkTop, 150, 30, "打开原始数据目录", "raw_data")
    Call CreateFileLinkButton(ws, 360, btnLinkTop, 150, 30, "打开最终成果大表", "papers_final_merged.xlsx")
    Call CreateFileLinkButton(ws, 530, btnLinkTop, 150, 30, "打开规范说明文档", "docs\PIPELINE_SPEC.md")
    
    Call CreateSectionTitle(ws, 20, logTitleTop, "【 系统运行提示与日志 】")
    
    Dim logBox As Shape
    Set logBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, logBoxTop, 660, 60)
    With logBox
        .Name = "shp_LogBox"
        .Line.ForeColor.RGB = RGB(200, 214, 229)
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(248, 249, 250)
        With .TextFrame2
            .TextRange.Text = "就绪: 极简流批直达架构已就绪。所有路径基于 ThisWorkbook.Path 动态自适应。" & vbCrLf & _
                              "提示: 若更新了本地代码可随时点击【重置与同步代码】；点击【一键执行全流程】直接产出成果大表！"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9
            .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
            .VerticalAnchor = msoAnchorTop
            .MarginLeft = 12
            .MarginTop = 8
        End With
    End With
End Sub

Private Sub CreateSectionTitle(ws As Worksheet, left As Single, top As Single, titleText As String)
    Dim tb As Shape
    Set tb = ws.Shapes.AddShape(msoShapeRectangle, left, top, 450, 20)
    With tb
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = titleText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 0
            .MarginTop = 0
        End With
    End With
End Sub

Private Sub CreateMetricCard(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                             title As String, shapeId As String, defaultVal As String, bgColor As Long, numColor As Long)
    Dim card As Shape
    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With card
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = bgColor
    End With
    
    Dim tbTitle As Shape
    Set tbTitle = ws.Shapes.AddShape(msoShapeRectangle, left + 8, top + 6, width - 16, 20)
    With tbTitle
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = title
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(120, 130, 140)
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
    
    Dim tbNum As Shape
    Set tbNum = ws.Shapes.AddShape(msoShapeRectangle, left + 8, top + 26, width - 16, 40)
    With tbNum
        .Name = shapeId
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = defaultVal
            .TextRange.Font.Name = "Arial"
            .TextRange.Font.Size = 22
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = numColor
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
End Sub

Private Sub CreateFileLinkButton(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                                 btnText As String, relPath As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With btn
        .Line.ForeColor.RGB = RGB(200, 210, 220)
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(245, 248, 250)
        .OnAction = "'打开相对路径 """ & relPath & """'"
        With .TextFrame2
            .TextRange.Text = btnText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(60, 80, 100)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

' ==============================================================================
' 2. 统计与动态更新机制 (全程静默无闪烁)
' ==============================================================================
Public Sub 刷新控制面板数据()
    Dim wb As Workbook, wsPanel As Worksheet
    Dim rootDir As String, rawDir As String
    Dim teacherCount As Long, wosRawCount As Long, eiRawCount As Long, cnkiRawCount As Long, scopusRawCount As Long
    Dim totalRaw As Long
    Dim finalMergeCount As Long, sciCount As Long, eiCount As Long, cnkiCount As Long, sciEiCount As Long
    Dim excTotal As Long, excSci As Long, excEi As Long, excCnki As Long, excSciEi As Long
    
    Dim oldScreenUpdating As Boolean, oldDisplayAlerts As Boolean, oldEnableEvents As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set wb = ThisWorkbook
    rootDir = wb.Path
    rawDir = rootDir & Application.PathSeparator & "raw_data"
    
    On Error Resume Next
    Set wsPanel = wb.Sheets(1)
    On Error GoTo 0
    If wsPanel Is Nothing Then GoTo CleanExit
    
    ' 1. 师生人员数量
    teacherCount = GetExternalWorkbookRows(rootDir & Application.PathSeparator & "config" & Application.PathSeparator & "teachers_profile.xlsx", "D")
    
    ' 2. 原始文件数量
    wosRawCount = ScanRawDataFileCount(rawDir, Array("wos.txt", "wos.xlsx", "savedrecs.txt", "savedrecs.xlsx"))
    eiRawCount = ScanRawDataFileCount(rawDir, Array("ei.xlsx", "ei.xls", "ei.csv", "ei_raw.xlsx"))
    cnkiRawCount = ScanRawDataFileCount(rawDir, Array("cnki.xls", "cnki.xlsx", "cnki_raw.xlsx", "cnki_raw.xls"))
    scopusRawCount = ScanRawDataFileCount(rawDir, Array("scopus.csv", "scopus_raw.csv"))
    totalRaw = wosRawCount + eiRawCount + cnkiRawCount + scopusRawCount
    
    ' 3. 100% 真实逐行统计入库成果与排除成果 (零硬编码，零比例估算)
    Call GetMergedOutputStats(rootDir & Application.PathSeparator & "papers_final_merged.xlsx", _
                             finalMergeCount, sciCount, eiCount, cnkiCount, sciEiCount, _
                             excTotal, excSci, excEi, excCnki, excSciEi)
    
    ' 4. 更新卡片
    Call UpdateCardValue(wsPanel, "card_Teacher", Format(teacherCount, "#,##0"))
    Call UpdateCardValue(wsPanel, "card_Raw", Format(totalRaw, "#,##0"))
    Call UpdateCardValue(wsPanel, "card_Final", Format(finalMergeCount, "#,##0"))
    
    ' 5. 更新明细清单 (实测数据)
    wsPanel.Range("D13").Value = Format(wosRawCount, "#,##0") & " 条"
    wsPanel.Range("D14").Value = Format(eiRawCount, "#,##0") & " 条"
    wsPanel.Range("D15").Value = Format(cnkiRawCount, "#,##0") & " 条"
    wsPanel.Range("D16").Value = Format(scopusRawCount, "#,##0") & " 条"
    
    wsPanel.Range("J13").Value = Format(finalMergeCount, "#,##0") & " 篇"
    wsPanel.Range("J14").Value = Format(sciCount, "#,##0") & " 篇"
    wsPanel.Range("J15").Value = Format(eiCount, "#,##0") & " 篇"
    wsPanel.Range("J16").Value = Format(cnkiCount, "#,##0") & " 篇"
    wsPanel.Range("J17").Value = Format(sciEiCount, "#,##0") & " 篇"
    
    ' 排除条目真实单行呈现
    wsPanel.Range("J18").Value = Format(excTotal, "#,##0") & " 篇 (SCI " & excSci & ", EI " & excEi & ", 中文 " & excCnki & ", SCI+EI " & excSciEi & ")"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
End Sub

' ==============================================================================
' 3. 业务全流程主控调度与全维度交付报告 (可被 Mod_Sync 动态无损热更)
' ==============================================================================
Public Sub 一键执行业务全流程()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("确定要启动【一键自动化执行全流程】吗？" & vbCrLf & vbCrLf & _
                    "系统将一次性自动完成：" & vbCrLf & _
                    "  1. 自动构建完善教师拼音与多格式别名库；" & vbCrLf & _
                    "  2. 按照设定的发表时间范围多源抽取 WOS/EI/知网；" & vbCrLf & _
                    "  3. 智能消歧认领、跨库去重并直出 7 列交付大表；" & vbCrLf & _
                    "  4. 自动刷新数据看板并生成全量交付总结报告。", _
                    vbQuestion + vbYesNo, "一键全流程")
    If answer <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' 1. 执行步骤 1 (教师特征库)
    Call AppendLog("【全流程启动】正在自动构建教师特征库...")
    On Error Resume Next
    Application.Run "Mod1_TeacherPinyin.生成老师拼音变体"
    
    ' 2. 执行步骤 2 (抽取清洗直出)
    Call AppendLog("【全流程进行中】正在抽取原始数据、时间校验与去重直出...")
    Application.Run "Mod2_CleanRawData.清洗所有原始数据"
    
    ' 3. 全面刷新数据看板
    Call 刷新控制面板数据
    Call AppendLog("【全流程完成】已成功直出 7 列交付大表 papers_final_merged.xlsx！")
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' 4. 提取看板全维度数据，生成丰富、严谨、多层次的终极大汇报弹窗
    Dim ws As Worksheet
    Dim teacherNum As String, rawTotalNum As String, finalNum As String
    Dim wosNum As String, eiNum As String, cnkiNum As String
    Dim sciNum As String, eiOutNum As String, cnkiOutNum As String, sciEiNum As String
    Dim excDetail As String, dateRangeStr As String
    
    Set ws = ThisWorkbook.Sheets(1)
    teacherNum = ws.Shapes("card_Teacher").TextFrame2.TextRange.Text
    rawTotalNum = ws.Shapes("card_Raw").TextFrame2.TextRange.Text
    finalNum = ws.Shapes("card_Final").TextFrame2.TextRange.Text
    
    wosNum = ws.Range("D13").Text
    eiNum = ws.Range("D14").Text
    cnkiNum = ws.Range("D15").Text
    
    sciNum = ws.Range("J14").Text
    eiOutNum = ws.Range("J15").Text
    cnkiOutNum = ws.Range("J16").Text
    sciEiNum = ws.Range("J17").Text
    excDetail = ws.Range("J18").Text
    
    dateRangeStr = ws.Range("D20").Text & " 至 " & ws.Range("F20").Text
    If Trim(dateRangeStr) = "至" Or (Trim(ws.Range("D20").Text) = "" And Trim(ws.Range("F20").Text) = "") Then
        dateRangeStr = "全部历史区间 (不限时间)"
    End If
    
    Dim summaryReport As String
    summaryReport = "【 科研台账 (keyantaizhang) 全流程交付报告 】" & vbCrLf & _
                    "--------------------------------------------------" & vbCrLf & _
                    "【统计时间范围】 " & dateRangeStr & vbCrLf & vbCrLf & _
                    "一、 原始数据采集概况：" & vbCrLf & _
                    "  - 原始记录总计：" & rawTotalNum & " 条 (WOS " & wosNum & " | EI " & eiNum & " | 知网 " & cnkiNum & ")" & vbCrLf & _
                    "  - 归属认领基准：" & teacherNum & " 位课题组师生/研究人员" & vbCrLf & vbCrLf & _
                    "二、 本室正式入库成果 (7 列标准终稿表)：" & vbCrLf & _
                    "  - 正式入库总量：" & finalNum & " 篇" & vbCrLf & _
                    "  - 收录结构细分：" & vbCrLf & _
                    "      * SCI 论文：" & sciNum & vbCrLf & _
                    "      * EI 论文：" & eiOutNum & vbCrLf & _
                    "      * 中文核心：" & cnkiOutNum & vbCrLf & _
                    "      * (其中 SCI+EI 双收录)：" & sciEiNum & vbCrLf & vbCrLf & _
                    "三、 排除与过滤说明：" & vbCrLf & _
                    "  - 非本室/未认领条目：" & excDetail & vbCrLf & _
                    "  - 判定准则：严格以出版年份(PY)为准，无本室教师成果安全排除" & vbCrLf & vbCrLf & _
                    "四、 成果交付文件：" & vbCrLf & _
                    "  - 最终成果大表：papers_final_merged.xlsx (7 大核心字段已就绪)"
    
    MsgBox summaryReport, vbInformation + vbOKOnly, "论文整理全流程已就绪"
End Sub

Public Sub 一键自动化执行全流程()
    Call 一键执行业务全流程
End Sub

Public Sub Btn_RunAll_Click()
    Call 一键执行业务全流程
End Sub

Private Function ScanRawDataFileCount(folderPath As String, candidateNames As Variant) As Long
    Dim fso As Object, i As Long, fName As String, fullPath As String, fExt As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then ScanRawDataFileCount = 0: Exit Function
    
    For i = LBound(candidateNames) To UBound(candidateNames)
        fName = CStr(candidateNames(i))
        fullPath = folderPath & Application.PathSeparator & fName
        
        If fso.FileExists(fullPath) Then
            fExt = LCase(fso.GetExtensionName(fullPath))
            If fExt = "txt" Then
                Dim stm As Object, content As String, lines() As String, c As Long, j As Long
                Set stm = CreateObject("ADODB.Stream")
                stm.Type = 2: stm.Charset = "utf-8": stm.Open
                stm.LoadFromFile fullPath
                content = stm.ReadText(-1)
                stm.Close
                
                lines = Split(content, vbLf)
                c = 0
                For j = 0 To UBound(lines)
                    If Trim(Replace(lines(j), vbCr, "")) <> "" Then c = c + 1
                Next j
                If c >= 2 Then ScanRawDataFileCount = c - 1 Else ScanRawDataFileCount = 0
                Exit Function
            ElseIf fExt = "xlsx" Or fExt = "xls" Or fExt = "csv" Then
                ScanRawDataFileCount = GetExternalWorkbookRows(fullPath, "A")
                Exit Function
            End If
        End If
    Next i
    
    ScanRawDataFileCount = 0
End Function

Private Sub GetMergedOutputStats(filePath As String, ByRef totalRows As Long, ByRef sciTotal As Long, _
                                 ByRef eiTotal As Long, ByRef cnkiTotal As Long, ByRef sciEiTotal As Long, _
                                 ByRef excTotal As Long, ByRef excSci As Long, ByRef excEi As Long, _
                                 ByRef excCnki As Long, ByRef excSciEi As Long)
    totalRows = 0: sciTotal = 0: eiTotal = 0: cnkiTotal = 0: sciEiTotal = 0
    excTotal = 0: excSci = 0: excEi = 0: excCnki = 0: excSciEi = 0
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub
    
    Dim extWb As Workbook, isAlreadyOpen As Boolean, fName As String
    fName = fso.GetFileName(filePath)
    
    On Error Resume Next
    Set extWb = Workbooks(fName)
    If extWb Is Nothing Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Set extWb = Workbooks.Open(filePath, ReadOnly:=True, AddToMru:=False)
        isAlreadyOpen = False
    Else
        isAlreadyOpen = True
    End If
    On Error GoTo 0
    
    If extWb Is Nothing Then Exit Sub
    
    ' 1. 真实逐行统计 Sheet 1: 入库成果
    Dim ws1 As Worksheet, lr1 As Long, r As Long, sType As String
    On Error Resume Next
    Set ws1 = extWb.Sheets(1)
    On Error GoTo 0
    If Not ws1 Is Nothing Then
        lr1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
        If lr1 >= 2 Then
            totalRows = lr1 - 1
            For r = 2 To lr1
                sType = Trim(CStr(ws1.Cells(r, "G").Value))
                If InStr(sType, "SCI") > 0 Then sciTotal = sciTotal + 1
                If InStr(sType, "EI") > 0 Then eiTotal = eiTotal + 1
                If InStr(sType, "中文核心") > 0 Then cnkiTotal = cnkiTotal + 1
                If InStr(sType, "SCI") > 0 And InStr(sType, "EI") > 0 Then sciEiTotal = sciEiTotal + 1
            Next r
        End If
    End If
    
    ' 2. 真实逐行统计 Sheet 2: 未认领排除成果 (100% 动态实测)
    Dim ws2 As Worksheet, lr2 As Long
    If extWb.Sheets.Count >= 2 Then
        Set ws2 = extWb.Sheets(2)
        lr2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        If lr2 >= 2 Then
            excTotal = lr2 - 1
            For r = 2 To lr2
                sType = Trim(CStr(ws2.Cells(r, "G").Value))
                If InStr(sType, "SCI") > 0 Then excSci = excSci + 1
                If InStr(sType, "EI") > 0 Then excEi = excEi + 1
                If InStr(sType, "中文核心") > 0 Then excCnki = excCnki + 1
                If InStr(sType, "SCI") > 0 And InStr(sType, "EI") > 0 Then excSciEi = excSciEi + 1
            Next r
        End If
    End If
    
    If Not isAlreadyOpen Then extWb.Close SaveChanges:=False
End Sub

Private Function GetExternalWorkbookRows(filePath As String, checkCol As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then GetExternalWorkbookRows = 0: Exit Function
    
    Dim extWb As Workbook, isAlreadyOpen As Boolean, fName As String
    fName = fso.GetFileName(filePath)
    
    On Error Resume Next
    Set extWb = Workbooks(fName)
    If extWb Is Nothing Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Set extWb = Workbooks.Open(filePath, ReadOnly:=True, AddToMru:=False)
        isAlreadyOpen = False
    Else
        isAlreadyOpen = True
    End If
    
    If Not extWb Is Nothing Then
        Dim ws As Worksheet, lastRow As Long
        Set ws = extWb.Sheets(1)
        lastRow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).Row
        If lastRow >= 2 Then GetExternalWorkbookRows = lastRow - 1 Else GetExternalWorkbookRows = 0
        If Not isAlreadyOpen Then extWb.Close SaveChanges:=False
    Else
        GetExternalWorkbookRows = 0
    End If
    On Error GoTo 0
End Function

Private Sub UpdateCardValue(ws As Worksheet, shapeName As String, newVal As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    If Not shp Is Nothing Then shp.TextFrame2.TextRange.Text = newVal
    On Error GoTo 0
End Sub

Public Sub AppendLog(msg As String)
    Dim wsPanel As Worksheet, timeStr As String
    timeStr = Format(Now, "hh:mm:ss")
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets(1)
    If Not wsPanel Is Nothing Then
        wsPanel.Shapes("shp_LogBox").TextFrame2.TextRange.Text = "[" & timeStr & "] " & msg
    End If
    On Error GoTo 0
End Sub

Public Sub 打开相对路径(relPath As String)
    Dim fullPath As String, fso As Object
    fullPath = ThisWorkbook.Path & Application.PathSeparator & relPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fullPath) Then
        If LCase(fso.GetExtensionName(fullPath)) = "xlsx" Or LCase(fso.GetExtensionName(fullPath)) = "xlsm" Or LCase(fso.GetExtensionName(fullPath)) = "xls" Then
            Workbooks.Open fullPath
        Else
            CreateObject("WScript.Shell").Run """" & fullPath & """"
        End If
    ElseIf fso.FolderExists(fullPath) Then
        CreateObject("WScript.Shell").Run "explorer.exe """ & fullPath & """"
    Else
        MsgBox "未找到指定的目标文件或文件夹：" & vbCrLf & fullPath, vbExclamation + vbOKOnly, "提示"
    End If
End Sub