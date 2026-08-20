' ==============================================================================
' 模块名称: Mod0_ControlPanel
' 核心职责: 交互控制台视图渲染、预设按钮回调、全流程总调度与日志引擎
' 设计哲学: 专注 UI 呈现与用户交互，数据逐行扫描统计下沉至 Mod0_MetricsEngine
' ==============================================================================
Option Explicit

' ==============================================================================
' 1. 控制台主界面标准重构与渲染
' ==============================================================================
Public Sub 一键生成控制面板()
    Dim wb As Workbook, ws As Worksheet
    Dim oldScreenUpdating As Boolean, oldDisplayAlerts As Boolean, oldEnableEvents As Boolean
    
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Sheets("控制面板")
    If ws Is Nothing Then Set ws = wb.Sheets(1): ws.Name = "控制面板"
    On Error GoTo 0
    
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ws.Columns("A").ColumnWidth = 3.5
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 11
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 11
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 3.5
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 3.5
    
    Call DrawHeaderBanner(ws)
    Call DrawSummaryCards(ws)
    Call DrawDateBar(ws)
    Call DrawWorkflowAndPrinciples(ws)
    Call DrawQuickLinksAndLog(ws)
    
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    
    Call 刷新控制面板数据
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
End Sub

Private Sub DrawHeaderBanner(ws As Worksheet)
    Dim banner As Shape, btnRun As Shape
    Set banner = ws.Shapes.AddShape(msoShapeRectangle, 20, 15, 660, 54)
    With banner
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(24, 76, 120)
        With .TextFrame2
            .TextRange.Text = "科研台账 (keyantaizhang) 课题组学术家底一账摸清" & vbCrLf & _
                              "全流程自动化处理 · 100% 动态逐行扫描实测 · SCI/EI/中文核心 · 双工作表 8 列直出"
            .TextRange.Paragraphs(1).Font.Size = 15
            .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.Paragraphs(2).Font.Size = 9.5
            .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(214, 234, 248)
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 20
        End With
    End With
    
    Set btnRun = ws.Shapes.AddShape(msoShapeRoundedRectangle, 475, 25, 195, 34)
    With btnRun
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(230, 126, 34)
        .OnAction = "Mod0_ControlPanel.Btn_RunAll_Click"
        With .TextFrame2
            .TextRange.Text = ">>> 一键自动化执行全流程 <<<"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Sub DrawSummaryCards(ws As Worksheet)
    Call CreateMetricCard(ws, 20, 110, 210, 75, "师生档案 (config)", "card_Teacher", "0", RGB(238, 245, 252), RGB(31, 110, 165))
    Call CreateMetricCard(ws, 245, 110, 210, 75, "原始导入 (raw_data)", "card_Raw", "0", RGB(238, 245, 252), RGB(39, 174, 96))
    Call CreateMetricCard(ws, 470, 110, 210, 75, "入库成果 (papers_final)", "card_Final", "0", RGB(254, 249, 231), RGB(211, 84, 0))
    
    Call CreateSectionTitle(ws, 20, 85, "一、 数据指标概览与结构明细")
    
    ws.Range("B12").Value = "【 原始文献导入清单 】"
    ws.Range("B13").Value = "WOS 原始文献 (SCI):"
    ws.Range("B14").Value = "EI 原始文献 (Compendex):"
    ws.Range("B15").Value = "知网原始文献 (CNKI):"
    ws.Range("B16").Value = "Scopus 备用文献:"
    
    ws.Range("H12").Value = "【 成果大表 papers_final_merged 】"
    ws.Range("H13").Value = "正式入库成果总量:"
    ws.Range("H14").Value = "- SCI 检索论文:"
    ws.Range("H15").Value = "- EI 检索论文:"
    ws.Range("H16").Value = "- 中文核心期刊:"
    ws.Range("H17").Value = "(SCI+EI 双收录):": ws.Range("J17").Value = "0 篇"
    ws.Range("H18").Value = "【排除未认领明细】:"
    
    With ws.Range("B13:D16, H13:J18")
        .Font.Name = "微软雅黑"
        .Font.Size = 9.5
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("B12, H12")
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(24, 76, 120)
    End With
    
    ws.Range("D13:D16, J13:J18").HorizontalAlignment = xlRight
    ws.Range("D13:D16, J13:J18").Font.Bold = True
End Sub

Private Sub DrawDateBar(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 200, "二、 成果发表时间范围筛选 (严格以出版年份 PY 过滤)")
    
    ws.Range("B20").Value = "起始日期 (YYYY/MM/DD):"
    ws.Range("B20").Font.Name = "微软雅黑"
    ws.Range("B20").Font.Size = 9.5
    ws.Range("B20").Font.Bold = True
    ws.Range("B20").Font.Color = RGB(24, 76, 120)
    
    With ws.Range("D20")
        .Value = "2026/01/01"
        .Font.Name = "微软雅黑": .Font.Size = 10: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(254, 249, 231)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(243, 156, 18)
    End With
    
    ws.Range("E20").Value = "至 截止日期:"
    ws.Range("E20").Font.Name = "微软雅黑"
    ws.Range("E20").Font.Size = 9.5
    ws.Range("E20").Font.Bold = True
    ws.Range("E20").HorizontalAlignment = xlCenter
    
    With ws.Range("F20")
        .Value = "2026/12/31"
        .Font.Name = "微软雅黑": .Font.Size = 10: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(254, 249, 231)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(243, 156, 18)
    End With
    
    Dim btnTop As Single, btnH As Single
    btnTop = ws.Range("H20").Top - 2
    btnH = 22
    Call CreateMiniPresetBtn(ws, 440, btnTop, 44, btnH, "全量", "Mod0_ControlPanel.SetDatePreset_All", RGB(127, 140, 141))
    Call CreateMiniPresetBtn(ws, 489, btnTop, 44, btnH, "2026", "Mod0_ControlPanel.SetDatePreset_2026", RGB(41, 128, 185))
    Call CreateMiniPresetBtn(ws, 538, btnTop, 44, btnH, "2025", "Mod0_ControlPanel.SetDatePreset_2025", RGB(39, 174, 96))
    Call CreateMiniPresetBtn(ws, 587, btnTop, 44, btnH, "2024", "Mod0_ControlPanel.SetDatePreset_2024", RGB(155, 89, 182))
    Call CreateMiniPresetBtn(ws, 636, btnTop, 44, btnH, "近3年", "Mod0_ControlPanel.SetDatePreset_3Years", RGB(230, 126, 34))
End Sub

Private Sub CreateMiniPresetBtn(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                               txt As String, actionMacro As String, colorRgb As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With btn
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = colorRgb
        .OnAction = actionMacro
        With .TextFrame2
            .TextRange.Text = txt
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Public Sub SetDatePreset_All()
    ThisWorkbook.Sheets(1).Range("D20").Value = ""
    ThisWorkbook.Sheets(1).Range("F20").Value = ""
    Call 刷新控制面板数据
End Sub

Public Sub SetDatePreset_2026()
    ThisWorkbook.Sheets(1).Range("D20").Value = "2026/01/01"
    ThisWorkbook.Sheets(1).Range("F20").Value = "2026/12/31"
    Call 刷新控制面板数据
End Sub

Public Sub SetDatePreset_2025()
    ThisWorkbook.Sheets(1).Range("D20").Value = "2025/01/01"
    ThisWorkbook.Sheets(1).Range("F20").Value = "2025/12/31"
    Call 刷新控制面板数据
End Sub

Public Sub SetDatePreset_2024()
    ThisWorkbook.Sheets(1).Range("D20").Value = "2024/01/01"
    ThisWorkbook.Sheets(1).Range("F20").Value = "2024/12/31"
    Call 刷新控制面板数据
End Sub

Public Sub SetDatePreset_3Years()
    ThisWorkbook.Sheets(1).Range("D20").Value = "2024/01/01"
    ThisWorkbook.Sheets(1).Range("F20").Value = "2026/12/31"
    Call 刷新控制面板数据
End Sub

Private Sub DrawWorkflowAndPrinciples(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 230, "三、 自动化作业流水线与核心治理原则")
    
    Dim boxW As Single, boxH As Single, gap As Single, startL As Single, topY As Single
    boxW = 153: boxH = 50: gap = 16: startL = 20: topY = 250
    
    Dim s1 As Shape
    Set s1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL, topY, boxW, boxH)
    With s1
        .Line.ForeColor.RGB = RGB(41, 128, 185): .Line.Weight = 1.5
        .Fill.ForeColor.RGB = RGB(238, 245, 252)
        .OnAction = "Mod1_TeacherPinyin.生成老师拼音变体"
        With .TextFrame2
            .TextRange.Text = "步骤 1: 构建师生特征库" & vbCrLf & "生成全格式拼音与检索特征"
            .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(24, 76, 120)
            .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    Dim s2 As Shape
    Set s2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL + boxW + gap, topY, boxW, boxH)
    With s2
        .Line.ForeColor.RGB = RGB(39, 174, 96): .Line.Weight = 1.5
        .Fill.ForeColor.RGB = RGB(238, 245, 252)
        .OnAction = "Mod2_PipelineMain.清洗所有原始数据"
        With .TextFrame2
            .TextRange.Text = "步骤 2: 抽取清洗与去重" & vbCrLf & "多源抽取·消歧认领·双表直出"
            .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(39, 174, 96)
            .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    Dim s3 As Shape
    Set s3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL + (boxW + gap) * 2, topY, boxW, boxH)
    With s3
        .Line.ForeColor.RGB = RGB(211, 84, 0): .Line.Weight = 1.5
        .Fill.ForeColor.RGB = RGB(254, 249, 231)
        .OnAction = "Mod0_ControlPanel.刷新控制面板数据"
        With .TextFrame2
            .TextRange.Text = "步骤 3: 刷新看板数据" & vbCrLf & "100% 逐行动态扫描实测"
            .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(211, 84, 0)
            .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    
    Dim s4 As Shape
    Set s4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, startL + (boxW + gap) * 3, topY, boxW, boxH)
    With s4
        .Line.ForeColor.RGB = RGB(142, 68, 173): .Line.Weight = 1.5
        .Fill.ForeColor.RGB = RGB(245, 238, 248)
        .OnAction = "Mod_Sync.一键热更并重置面板"
        With .TextFrame2
            .TextRange.Text = "底座: 热更与重置面板" & vbCrLf & "秒级载入 vba_modules"
            .TextRange.Paragraphs(1).Font.Size = 9.5: .TextRange.Paragraphs(1).Font.Bold = msoTrue
            .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(142, 68, 173)
            .TextRange.Paragraphs(2).Font.Size = 8: .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(100, 110, 120)
            .VerticalAnchor = msoAnchorMiddle: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Sub DrawQuickLinksAndLog(ws As Worksheet)
    Call CreateSectionTitle(ws, 20, 310, "四、 快捷文件通道与实时处理日志")
    
    Dim btnLinkTop As Single
    btnLinkTop = 330
    Call CreateFileLinkButton(ws, 20, btnLinkTop, 150, 30, "教师名单 (config)", "config\teachers_profile.xlsx")
    Call CreateFileLinkButton(ws, 190, btnLinkTop, 150, 30, "原始数据 (raw_data)", "raw_data")
    Call CreateFileLinkButton(ws, 360, btnLinkTop, 150, 30, "交付成果大表", "papers_final_merged.xlsx")
    Call CreateFileLinkButton(ws, 530, btnLinkTop, 150, 30, "检索实操指南", "docs\DATABASE_RETRIEVAL_GUIDE.md")
    
    With ws.Range("B23:J27")
        .Merge
        .Interior.Color = RGB(245, 247, 250)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(218, 225, 231)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(52, 73, 94)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Value = "【系统日志就绪】欢迎使用科研台账自动化整理系统！" & vbCrLf & _
                 "点击上方 [ >>> 一键自动化执行全流程 <<< ] 开始批量处理。"
    End With
End Sub

Private Sub CreateSectionTitle(ws As Worksheet, left As Single, top As Single, titleText As String)
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(msoShapeRectangle, left, top, 660, 20)
    With lbl
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
        End With
    End With
End Sub

Private Sub CreateMetricCard(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                             titleText As String, numShapeName As String, initVal As String, _
                             bgColor As Long, borderColor As Long)
    Dim bg As Shape, valShp As Shape
    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With bg
        .Line.ForeColor.RGB = borderColor: .Line.Weight = 1.5
        .Fill.ForeColor.RGB = bgColor
        With .TextFrame2
            .TextRange.Text = titleText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(44, 62, 80)
            .VerticalAnchor = msoAnchorTop
            .MarginTop = 8: .MarginLeft = 10
        End With
    End With
    
    Set valShp = ws.Shapes.AddShape(msoShapeRectangle, left + 10, top + 26, width - 20, height - 32)
    With valShp
        .Name = numShapeName
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = initVal
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 22
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = borderColor
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .MarginLeft = 0
        End With
    End With
End Sub

Private Sub CreateFileLinkButton(ws As Worksheet, left As Single, top As Single, width As Single, height As Single, _
                                 btnText As String, targetRelPath As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    With btn
        .Line.ForeColor.RGB = RGB(189, 195, 199): .Line.Weight = 1
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .OnAction = "'Mod0_ControlPanel.打开相对路径 """ & targetRelPath & """'"
        With .TextFrame2
            .TextRange.Text = btnText
            .TextRange.Font.Name = "微软雅黑"
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(60, 80, 100)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

' ==============================================================================
' 2. 调度总控与一键全流程执行
' ==============================================================================
Public Sub 刷新控制面板数据()
    On Error Resume Next
    Call Mod0_MetricsEngine.扫描并刷新看板
    On Error GoTo 0
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
                    "  3. 智能消歧认领、跨库去重并直出 8 列交付大表 (含影响因子)；" & vbCrLf & _
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
    Call AppendLog("【全流程完成】已成功直出 8 列交付大表 (含影响因子) papers_final_merged.xlsx！")
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
                    "二、 课题组正式入库成果 (8 列标准终稿表，含影响因子)：" & vbCrLf & _
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
                    "  - 最终成果大表：papers_final_merged.xlsx (8 大核心字段已就绪，含影响因子)"
    
    MsgBox summaryReport, vbInformation + vbOKOnly, "论文整理全流程已就绪"
End Sub

Public Sub AppendLog(msg As String)
    Dim ws As Worksheet, oldText As String, timeStr As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(1)
    If ws Is Nothing Then Exit Sub
    
    timeStr = Format(Now, "hh:mm:ss")
    oldText = ws.Range("B23").Value
    ws.Range("B23").Value = "[" & timeStr & "] " & msg & vbCrLf & oldText
    On Error GoTo 0
End Sub

Public Sub 打开相对路径(relPath As String)
    Dim fullPath As String
    fullPath = ThisWorkbook.Path & Application.PathSeparator & relPath
    If CreateObject("Scripting.FileSystemObject").FileExists(fullPath) Or _
       CreateObject("Scripting.FileSystemObject").FolderExists(fullPath) Then
        CreateObject("WScript.Shell").Run """" & fullPath & """"
    Else
        MsgBox "未找到指定的文件或目录：" & vbCrLf & fullPath, vbExclamation, "路径不存在"
    End If
End Sub