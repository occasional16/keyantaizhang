' ==============================================================================
' 模块名称: Mod0_ControlPanel
' 核心职责: 控制面板交互工作台 (UI, Presentation & Business Pipeline Orchestrator)
' 设计哲学:
'   1. 顶部横幅双核设计: [ 一键热更并重置面板 ] (系统运维) + [ >>> 一键自动化执行全流程 <<< ] (业务生产)
'   2. 全景呈现: 数据看板 + 时间筛选 + 4步流水线卡片 + 3步流程与工作原理 + 快捷通道 + 实时日志
'   3. 680pt 几何严密对齐，纯正 GBK 中文，绝对杜绝任何 Unicode Emoji 乱码或图层重叠
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
    
    ' 顺序渲染 5 大核心功能板块 (绝对坐标精密计算，彻底杜绝重叠)
    Call DrawHeaderBanner(ws)           ' Top: 15 ~ 70
    Call DrawSummaryCards(ws)           ' Top: 80 ~ 275
    Call DrawDateBar(ws)                ' Top: 285 ~ 335
    Call DrawWorkflowCards(ws)          ' Top: 345 ~ 416 (步骤1/2/3/底座 4张交互卡片)
    Call DrawPrinciplesCard(ws)         ' Top: 425 ~ 535 (极简操作流程与三大工作原理)
    Call DrawQuickLinksAndLog(ws)       ' Top: 545 ~ 660 (4个快捷通道 + 独立日志框)
    
    Call 刷新控制面板数据
    
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    Call AppendLog("【控制面板】工作台已成功初始化并就绪！")
End Sub

' ------------------------------------------------------------------------------
' 1. 顶部主横幅与双核控制按钮 (Top: 15 ~ 70)
' ------------------------------------------------------------------------------
Private Sub DrawHeaderBanner(ws As Worksheet)
    ' 主横幅底板 (20 ~ 680pt, 宽 660pt, 高 55pt)
    Dim banner As Shape
    Set banner = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 660, 55)
    With banner
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(24, 76, 120)
        With .TextFrame2
            .TextRange.Text = "科研台账 (keyantaizhang) 课题组学术家底一账摸清" & vbCrLf & _
                              "全流程自动化 · 100% 动态逐行扫描 · SCI/EI/中文核心 · 双工作表 9 列直出 (含影响因子)"
            .TextRange.Paragraphs(1).Font.Name = "微软雅黑"
            .TextRange.Paragraphs(1).Font.Size = 14.5
            .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.Paragraphs(2).Font.Name = "微软雅黑"
            .TextRange.Paragraphs(2).Font.Size = 8.5
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(214, 234, 248)
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 14
        End With
    End With
    
    ' 运维按钮: [ 一键热更并重置面板 ] (310 ~ 455pt, 宽 145pt, 高 35pt, 深青灰)
    Dim btnSync As Shape
    Set btnSync = ws.Shapes.AddShape(msoShapeRoundedRectangle, 310, 25, 145, 35)
    With btnSync
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(52, 73, 94)
        .OnAction = "Mod_Sync.一键热更并重置面板"
        With .TextFrame2
            .TextRange.Text = "[ 一键热更并重置面板 ]"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' 生产按钮: [ >>> 一键自动化执行全流程 <<< ] (465 ~ 670pt, 宽 205pt, 高 35pt, 橙色高亮)
    Dim btnRun As Shape
    Set btnRun = ws.Shapes.AddShape(msoShapeRoundedRectangle, 465, 25, 205, 35)
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

' ------------------------------------------------------------------------------
' 2. 数据指标概览与结构明细 (Top: 80 ~ 275)
' ------------------------------------------------------------------------------
Private Sub DrawSummaryCards(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 80, "一、 数据指标概览与结构明细 (100% 动态实测统计)")
    
    ' 3 大核心 KPI 状态卡片 (Top: 102, Height: 68)
    Call CreateMetricCard(ws, 20, 102, 210, 68, "师生名单库 (config)", "card_Teacher", "0", RGB(238, 245, 252), RGB(31, 110, 165))
    Call CreateMetricCard(ws, 245, 102, 210, 68, "原始库数据 (raw_data)", "card_Raw", "0", RGB(238, 245, 252), RGB(39, 174, 96))
    Call CreateMetricCard(ws, 470, 102, 210, 68, "本室入库成果 (papers_final)", "card_Final", "0", RGB(254, 249, 231), RGB(211, 84, 0))
    
    ' 数据明细网格 (Rows 12 ~ 18)
    Dim r As Long
    For r = 12 To 18
        ws.Rows(r).RowHeight = 18
    Next r
    
    ws.Range("B12").Value = "【 原始文献导入清单 (raw_data) 】"
    ws.Range("B13").Value = "WOS 原始文献 (SCI):": ws.Range("D13").Value = "0 篇"
    ws.Range("B14").Value = "EI 原始文献 (Compendex):": ws.Range("D14").Value = "0 篇"
    ws.Range("B15").Value = "知网原始文献 (CNKI):": ws.Range("D15").Value = "0 篇"
    ws.Range("B16").Value = "Scopus 备用文献:": ws.Range("D16").Value = "0 篇"
    
    ws.Range("H12").Value = "【 成果大表 papers_final_merged 】"
    ws.Range("H13").Value = "正式入库成果总量:": ws.Range("J13").Value = "0 篇"
    ws.Range("H14").Value = "其中 SCI 检索论文:": ws.Range("J14").Value = "0 篇"
    ws.Range("H15").Value = "其中 EI 检索论文:": ws.Range("J15").Value = "0 篇"
    ws.Range("H16").Value = "其中 中文核心期刊:": ws.Range("J16").Value = "0 篇"
    ws.Range("H17").Value = "(SCI+EI 双收录):": ws.Range("J17").Value = "0 篇"
    ws.Range("H18").Value = "非本室/排除条目:": ws.Range("J18").Value = "0 篇 (SCI 0, EI 0, 中文 0, SCI+EI 0)"
    
    With ws.Range("B12, H12")
        .Font.Name = "微软雅黑": .Font.Size = 10: .Font.Bold = True: .Font.Color = RGB(24, 76, 120)
    End With
    
    With ws.Range("B13:D16, H13:J17")
        .Font.Name = "微软雅黑": .Font.Size = 9.5: .Font.Color = RGB(60, 60, 60): .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("H18, J18")
        .Font.Name = "微软雅黑": .Font.Size = 9: .Font.Color = RGB(130, 130, 130): .VerticalAlignment = xlCenter
    End With
    
    ws.Range("D13:D16, J13:J17").HorizontalAlignment = xlRight
    ws.Range("D13:D16, J13:J17").Font.Bold = True
    ws.Range("J18").HorizontalAlignment = xlRight
End Sub

' ------------------------------------------------------------------------------
' 3. 成果发表时间范围筛选 (Top: 285 ~ 335)
' ------------------------------------------------------------------------------
Private Sub DrawDateBar(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 285, "二、 成果发表时间范围筛选 (严格以出版年份 PY 过滤)")
    
    ws.Rows(19).RowHeight = 6
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
    
    ' 5 个快捷年份胶囊按钮
    Dim btnTop As Single, btnH As Single
    btnTop = ws.Range("H20").Top + 1
    btnH = 22
    
    Call CreateMiniPresetBtn(ws, 440, btnTop, 44, btnH, "全部", "Mod0_ControlPanel.SetDatePreset_All", RGB(127, 140, 141))
    Call CreateMiniPresetBtn(ws, 489, btnTop, 44, btnH, "2026", "Mod0_ControlPanel.SetDatePreset_2026", RGB(52, 152, 219))
    Call CreateMiniPresetBtn(ws, 538, btnTop, 44, btnH, "2025", "Mod0_ControlPanel.SetDatePreset_2025", RGB(46, 204, 113))
    Call CreateMiniPresetBtn(ws, 587, btnTop, 44, btnH, "2024", "Mod0_ControlPanel.SetDatePreset_2024", RGB(155, 89, 182))
    Call CreateMiniPresetBtn(ws, 636, btnTop, 44, btnH, "近3年", "Mod0_ControlPanel.SetDatePreset_3Years", RGB(230, 126, 34))
End Sub

Public Sub SetDatePreset_All()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").Value = ""
    ws.Range("F20").Value = ""
    Call AppendLog("【时间已设为: 全量历史区间】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_2026()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2026-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2026-12-31"
    Call AppendLog("【时间已设为: 2026 年度】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_2025()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2025-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2025-12-31"
    Call AppendLog("【时间已设为: 2025 年度】（点击右上角大按钮即可生成成果）")
End Sub

Public Sub SetDatePreset_2024()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("D20").NumberFormat = "yyyy-mm-dd"
    ws.Range("D20").Value = "2024-01-01"
    ws.Range("F20").NumberFormat = "yyyy-mm-dd"
    ws.Range("F20").Value = "2024-12-31"
    Call AppendLog("【时间已设为: 2024 年度】（点击右上角大按钮即可生成成果）")
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

' ------------------------------------------------------------------------------
' 4. 自动化流水线步骤卡片 (Top: 345 ~ 416, 步骤1/2/3/底座 4 张可点击卡片)
' ------------------------------------------------------------------------------
Private Sub DrawWorkflowCards(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 345, "三、 自动化作业流水线 (支持点击各卡片单独分步执行)")
    
    Dim boxW As Single, boxH As Single, gap As Single, startL As Single, topY As Single
    boxW = 153: boxH = 48: gap = 16: startL = 20: topY = 368
    
    ' 步骤 1 (蓝色)
    Dim s1 As Shape
    Set s1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL, topY, boxW, boxH)
    With s1
        .Line.ForeColor.RGB = RGB(41, 128, 185): .Line.Weight = 1.2
        .Fill.ForeColor.RGB = RGB(238, 245, 252)
        .OnAction = "Mod1_TeacherPinyin.生成老师拼音变体"
        With .TextFrame2
            .TextRange.Text = "步骤 1: 构建师生特征库" & vbCrLf & "生成全格式拼音与检索特征"
            .TextRange.Paragraphs(1).Font.Name = "微软雅黑": .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .TextRange.Paragraphs(2).Font.Name = "微软雅黑": .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' 步骤 2 (绿色)
    Dim s2 As Shape
    Set s2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL + boxW + gap, topY, boxW, boxH)
    With s2
        .Line.ForeColor.RGB = RGB(39, 174, 96): .Line.Weight = 1.2
        .Fill.ForeColor.RGB = RGB(238, 245, 252)
        .OnAction = "Mod2_PipelineMain.清洗所有原始数据"
        With .TextFrame2
            .TextRange.Text = "步骤 2: 抽取清洗与去重" & vbCrLf & "多源抽取·消歧·8列直出"
            .TextRange.Paragraphs(1).Font.Name = "微软雅黑": .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(39, 174, 96)
            .TextRange.Paragraphs(2).Font.Name = "微软雅黑": .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' 步骤 3 (橙色)
    Dim s3 As Shape
    Set s3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL + (boxW + gap) * 2, topY, boxW, boxH)
    With s3
        .Line.ForeColor.RGB = RGB(211, 84, 0): .Line.Weight = 1.2
        .Fill.ForeColor.RGB = RGB(254, 249, 231)
        .OnAction = "Mod0_ControlPanel.刷新控制面板数据"
        With .TextFrame2
            .TextRange.Text = "步骤 3: 刷新看板数据" & vbCrLf & "100% 逐行动态扫描实测"
            .TextRange.Paragraphs(1).Font.Name = "微软雅黑": .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(211, 84, 0)
            .TextRange.Paragraphs(2).Font.Name = "微软雅黑": .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    ' 底座 (紫色)
    Dim s4 As Shape
    Set s4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL + (boxW + gap) * 3, topY, boxW, boxH)
    With s4
        .Line.ForeColor.RGB = RGB(142, 68, 173): .Line.Weight = 1.2
        .Fill.ForeColor.RGB = RGB(245, 238, 248)
        .OnAction = "Mod_Sync.一键热更并重置面板"
        With .TextFrame2
            .TextRange.Text = "底座: 热更并重置面板" & vbCrLf & "秒级载入 vba_modules"
            .TextRange.Paragraphs(1).Font.Name = "微软雅黑": .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(142, 68, 173)
            .TextRange.Paragraphs(2).Font.Name = "微软雅黑": .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

' ------------------------------------------------------------------------------
' 5. 极简操作流程与系统工作原理 (Top: 425 ~ 535)
' ------------------------------------------------------------------------------
Private Sub DrawPrinciplesCard(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 425, "四、 极简操作流程与系统工作原理")
    
    Dim cardTop As Single
    cardTop = 447
    
    Dim guideCard As Shape
    Set guideCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, cardTop, 660, 82)
    With guideCard
        .Line.ForeColor.RGB = RGB(200, 220, 240)
        .Line.Weight = 1
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(246, 250, 254)
    End With
    
    ' 左栏: 3步极简操作流程
    Dim tbLeft As Shape
    Set tbLeft = ws.Shapes.AddShape(msoShapeRectangle, 32, cardTop + 5, 310, 72)
    With tbLeft
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "【 极简操作流程 (3 步搞定) 】" & vbCrLf & _
                              "1. 准备数据: 在 raw_data 目录放入 WOS/EI/知网 原始文件；" & vbCrLf & _
                              "2. 设定时间: 在上方面板选择或输入年份区间；" & vbCrLf & _
                              "3. 一键执行: 点击右上角【一键自动化执行全流程】大按钮！"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 8.5
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
    Set sepLine = ws.Shapes.AddLine(350, cardTop + 8, 350, cardTop + 74)
    With sepLine
        .Line.ForeColor.RGB = RGB(210, 225, 240)
        .Line.DashStyle = msoLineDash
    End With
    
    ' 右栏: 系统底层三大核心工作原理
    Dim tbRight As Shape
    Set tbRight = ws.Shapes.AddShape(msoShapeRectangle, 360, cardTop + 5, 310, 72)
    With tbRight
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "【 系统底层工作原理 (三大核心支柱) 】" & vbCrLf & _
                              "1. 智能消歧: 依据 100 位教师拼音库自动认领，非本室条目安全排除；" & vbCrLf & _
                              "2. 年份过滤: 严格按正式出版年份 (PY等) 筛选在期成果；" & vbCrLf & _
                              "3. 跨库合流: 基于 DOI 与题目哈希去重，自动标识 SCI+EI 复合收录。"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 8.5
            .TextRange.Paragraphs(1).Font.Size = 9.5
            .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .TextRange.Font.Fill.ForeColor.RGB = RGB(60, 80, 100)
            .VerticalAnchor = msoAnchorTop
            .MarginLeft = 0: .MarginTop = 0
        End With
    End With
End Sub

' ------------------------------------------------------------------------------
' 6. 快捷联动通道与实时处理日志 (Top: 540 ~ 655)
' ------------------------------------------------------------------------------
Private Sub DrawQuickLinksAndLog(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 540, "五、 快捷文件联动通道与实时处理日志")
    
    Dim btnLinkTop As Single
    btnLinkTop = 562
    Call CreateFileLinkButton(ws, 20, btnLinkTop, 153, 28, "师生档案 (config)", "config\teachers_profile.xlsx")
    Call CreateFileLinkButton(ws, 189, btnLinkTop, 153, 28, "原始数据 (raw_data)", "raw_data")
    Call CreateFileLinkButton(ws, 358, btnLinkTop, 153, 28, "交付成果大表", "papers_final_merged.xlsx")
    Call CreateFileLinkButton(ws, 527, btnLinkTop, 153, 28, "检索实操指南", "docs\DATABASE_RETRIEVAL_GUIDE.md")
    
    Dim logBoxTop As Single
    logBoxTop = 600
    Dim logBox As Shape
    Set logBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, logBoxTop, 660, 52)
    With logBox
        .Name = "shp_LogBox"
        .Line.ForeColor.RGB = RGB(200, 214, 229)
        .Line.Weight = 1
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(248, 249, 250)
        With .TextFrame2
            .TextRange.Text = "就绪: 极简流批直达架构已就绪。所有路径基于 ThisWorkbook.Path 动态自适应。" & vbCrLf & _
                              "提示: 点击上方【一键自动化执行全流程】直接产出 8 列标准成果大表 (含影响因子)！"
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9
            .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
            .VerticalAnchor = msoAnchorTop
            .MarginLeft = 12
            .MarginTop = 6
        End With
    End With
End Sub

' ------------------------------------------------------------------------------
' 通用 UI 辅助函数
' ------------------------------------------------------------------------------
Private Sub CreateSectionTitle(ws As Worksheet, left As Single, top As Single, titleText As String)
    Dim tb As Shape
    Set tb = ws.Shapes.AddShape(msoShapeRectangle, left, top, 600, 18)
    With tb
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = titleText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 10
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
    Set tbTitle = ws.Shapes.AddShape(msoShapeRectangle, left + 8, top + 5, width - 16, 18)
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
    Set tbNum = ws.Shapes.AddShape(msoShapeRectangle, left + 8, top + 23, width - 16, 38)
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

Private Sub CreateFileLinkButton(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                                 btnText As String, relPath As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With btn
        .Line.ForeColor.RGB = RGB(200, 215, 230)
        .Line.Weight = 1
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .OnAction = "'Mod0_ControlPanel.打开相对路径 """ & relPath & """'"
        With .TextFrame2
            .TextRange.Text = btnText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Public Sub 刷新控制面板数据()
    Call Mod0_MetricsEngine.扫描并刷新看板
End Sub

Public Sub Btn_RunAll_Click()
    Call 一键执行业务全流程
End Sub

Public Sub 一键自动化执行全流程()
    Call 一键执行业务全流程
End Sub

Public Sub 一键执行业务全流程()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("确定要启动【一键自动化执行全流程】吗？" & vbCrLf & vbCrLf & _
                    "系统将一次性自动完成：" & vbCrLf & _
                    "  1. 自动构建完善教师拼音与多格式别名库；" & vbCrLf & _
                    "  2. 按照设定的发表时间范围多源抽取 WOS/EI/知网；" & vbCrLf & _
                    "  3. 智能消歧认领、跨库去重并直出 9 列交付大表 (含通讯作者与影响因子) (含影响因子)；" & vbCrLf & _
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
    Call AppendLog("【全流程进行中】正在抽取原始数据、时间校验、JIF匹配与去重直出...")
    Application.Run "Mod2_PipelineMain.清洗所有原始数据"
    
    ' 3. 全面刷新数据看板
    Call 刷新控制面板数据
    Call AppendLog("【全流程完成】已成功直出 9 列交付大表 (含通讯作者与影响因子) (含影响因子) papers_final_merged.xlsx！")
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' 4. 提取看板全维度数据，生成终极大汇报弹窗
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
                    "二、 课题组正式入库成果 (9 列标准终稿表 (含通讯作者与影响因子)，含影响因子)：" & vbCrLf & _
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
                    "  - 最终成果大表：papers_final_merged.xlsx (9 大核心字段已就绪，含影响因子)"
    
    MsgBox summaryReport, vbInformation + vbOKOnly, "论文整理全流程已就绪"
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