' ==============================================================================
' 模块名称: Mod2_PipelineMain
' 核心职责: 科研台账多源清洗总控、记录聚合去重、双工作表 9 列直出与格式化引擎
' ==============================================================================
Option Explicit

Public Sub 清洗所有原始数据()
    Dim rootDir As String, rawDir As String, configPath As String, outPath As String
    Dim fso As Object, dictTeachers As Object, dictIF As Object
    Dim dictDoi As Object, dictTitle As Object, dictRecs As Object
    Dim totalRaw As Long, totalMerged As Long, acceptedCount As Long, excludedCount As Long, outOfDateCount As Long
    Dim wosCount As Long, eiCount As Long, cnkiCount As Long
    Dim wbOut As Workbook, wsOut As Worksheet, wsPanel As Worksheet
    Dim filterStart As Date, filterEnd As Date, hasStart As Boolean, hasEnd As Boolean
    Dim sStart As String, sEnd As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rootDir = ThisWorkbook.Path
    rawDir = rootDir & Application.PathSeparator & "raw_data"
    configPath = rootDir & Application.PathSeparator & "config" & Application.PathSeparator & "teachers_profile.xlsx"
    outPath = rootDir & Application.PathSeparator & "papers_final_merged.xlsx"
    
    If Not fso.FileExists(configPath) Then
        MsgBox "未找到教师主档案文件：" & vbCrLf & configPath, vbCritical, "缺少配置文件"
        Exit Sub
    End If
    
    If Not fso.FolderExists(rawDir) Then
        MsgBox "未找到原始数据目录 raw_data/ ！", vbCritical, "缺少原始目录"
        Exit Sub
    End If
    
    ' 读取控制台上的日期范围设定 (D20 与 F20)
    hasStart = False: hasEnd = False
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets(1)
    If Not wsPanel Is Nothing Then
        sStart = Trim(CStr(wsPanel.Range("D20").Value))
        sEnd = Trim(CStr(wsPanel.Range("F20").Value))
        If sStart <> "" Then
            If IsDate(sStart) Then filterStart = CDate(sStart): hasStart = True
        End If
        If sEnd <> "" Then
            If IsDate(sEnd) Then filterEnd = CDate(sEnd): hasEnd = True
        End If
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 0. 加载字典库
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "正在加载期刊影响因子字典 (journal_if.xlsx)..."
    On Error GoTo 0
    Set dictIF = LoadJournalIfDictionary(rootDir & Application.PathSeparator & "config" & Application.PathSeparator & "journal_if.xlsx")
    
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "正在加载教师匹配特征库..."
    On Error GoTo 0
    Set dictTeachers = LoadTeacherAliasDictionary(configPath)
    If dictTeachers.Count = 0 Then
        MsgBox "教师特征库为空！请先执行【步骤 1: 完善教师拼音与检索特征库】。", vbExclamation, "特征库未就绪"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Set dictDoi = CreateObject("Scripting.Dictionary")
    Set dictTitle = CreateObject("Scripting.Dictionary")
    Set dictRecs = CreateObject("Scripting.Dictionary")
    outOfDateCount = 0
    
    ' 1. 读取并合并 WOS (带日期过滤)
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "正在抽取 WOS (SCI) 原始数据并校验出版日期..."
    On Error GoTo 0
    wosCount = IngestWosData(rawDir, dictTeachers, dictDoi, dictTitle, dictRecs, hasStart, filterStart, hasEnd, filterEnd, outOfDateCount)
    
    ' 2. 读取并合并 EI (带日期过滤)
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "正在抽取 EI 原始数据并校验出版日期..."
    On Error GoTo 0
    eiCount = IngestEiData(rawDir, dictTeachers, dictDoi, dictTitle, dictRecs, hasStart, filterStart, hasEnd, filterEnd, outOfDateCount)
    
    ' 3. 读取并合并 CNKI (带日期过滤)
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "正在抽取 CNKI 原始数据并校验出版日期..."
    On Error GoTo 0
    cnkiCount = IngestCnkiData(rawDir, dictTeachers, dictDoi, dictTitle, dictRecs, hasStart, filterStart, hasEnd, filterEnd, outOfDateCount)
    
    totalRaw = wosCount + eiCount + cnkiCount
    totalMerged = dictRecs.Count
    
    ' 4. 双工作表直出 papers_final_merged.xlsx (Sheet1: 课题组入库成果, Sheet2: 未认领排除成果)
    Dim wsExc As Worksheet
    Set wbOut = GetOrCreateOutputWorkbook(outPath, wsOut, wsExc)
    
    If wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row >= 2 Then
        wsOut.Range("A2:I" & wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row).ClearContents
    End If
    If wsExc.Cells(wsExc.Rows.Count, "A").End(xlUp).Row >= 2 Then
        wsExc.Range("A2:I" & wsExc.Cells(wsExc.Rows.Count, "A").End(xlUp).Row).ClearContents
    End If
    
    Dim kVar As Variant, rec As Variant, rawAuthorStr As String, corrStr As String
    acceptedCount = 0
    excludedCount = 0
    
    For Each kVar In dictRecs.Keys
        rec = dictRecs(kVar)
        corrStr = ""
        If UBound(rec) >= 8 Then corrStr = CStr(rec(8))
        
        If Trim(CStr(rec(4))) <> "" Then
            ' 入库成果 -> 写入 Sheet 1 (9 列)
            acceptedCount = acceptedCount + 1
            wsOut.Cells(acceptedCount + 1, 1).Value = acceptedCount
            wsOut.Cells(acceptedCount + 1, 2).Value = rec(0)
            wsOut.Cells(acceptedCount + 1, 3).Value = rec(1)
            wsOut.Cells(acceptedCount + 1, 4).Value = rec(2)
            wsOut.Cells(acceptedCount + 1, 5).Value = rec(3)
            wsOut.Cells(acceptedCount + 1, 6).Value = rec(4)
            wsOut.Cells(acceptedCount + 1, 7).Value = corrStr
            wsOut.Cells(acceptedCount + 1, 8).Value = rec(6)
            wsOut.Cells(acceptedCount + 1, 9).Value = GetJournalIfValue(CStr(rec(1)), dictIF)
        Else
            ' 未认领/排除成果 -> 写入 Sheet 2 (9 列)
            excludedCount = excludedCount + 1
            rawAuthorStr = ""
            If UBound(rec) >= 7 Then rawAuthorStr = CStr(rec(7))
            
            wsExc.Cells(excludedCount + 1, 1).Value = excludedCount
            wsExc.Cells(excludedCount + 1, 2).Value = rec(0)
            wsExc.Cells(excludedCount + 1, 3).Value = rec(1)
            wsExc.Cells(excludedCount + 1, 4).Value = rec(2)
            wsExc.Cells(excludedCount + 1, 5).Value = rec(3)
            wsExc.Cells(excludedCount + 1, 6).Value = rawAuthorStr
            wsExc.Cells(excludedCount + 1, 7).Value = corrStr
            wsExc.Cells(excludedCount + 1, 8).Value = rec(6)
            wsExc.Cells(excludedCount + 1, 9).Value = GetJournalIfValue(CStr(rec(1)), dictIF)
        End If
    Next kVar
    
    Call FormatOutputSheet(wsOut, acceptedCount + 1)
    Call FormatOutputSheet(wsExc, excludedCount + 1)
    wbOut.Close SaveChanges:=True
    
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 5. 刷新控制台看板与日志
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "【整理完成】课题组正式入库 " & acceptedCount & " 篇，未认领排除 " & excludedCount & " 篇，跨期过滤 " & outOfDateCount & " 篇。"
    On Error GoTo 0
End Sub

Public Sub AddOrMergeRecord(ByVal title As String, ByVal journal As String, ByVal vol As String, ByVal issue As String, _
                            ByVal matchedAuthors As String, ByVal corrAuthor As String, ByVal rawAuthors As String, ByVal doi As String, ByVal srcType As String, _
                            dictDoi As Object, dictTitle As Object, dictRecs As Object)
    Dim normDoi As String, normT As String
    Dim recId As Long, oldRec As Variant, mergedType As String
    
    normDoi = LCase(Trim(doi))
    normT = NormalizeTitleForMatch(title)
    
    recId = 0
    If normDoi <> "" Then
        If dictDoi.Exists(normDoi) Then recId = dictDoi(normDoi)
    End If
    If recId = 0 And normT <> "" Then
        If dictTitle.Exists(normT) Then recId = dictTitle(normT)
    End If
    
    If recId > 0 Then
        oldRec = dictRecs(recId)
        mergedType = MergeSourceTypes(CStr(oldRec(6)), srcType)
        oldRec(6) = mergedType
        If Trim(CStr(oldRec(4))) = "" And Trim(matchedAuthors) <> "" Then oldRec(4) = matchedAuthors
        If Trim(CStr(oldRec(2))) = "" And Trim(vol) <> "" Then oldRec(2) = vol
        If Trim(CStr(oldRec(3))) = "" And Trim(issue) <> "" Then oldRec(3) = issue
        If UBound(oldRec) >= 7 Then
            If Trim(CStr(oldRec(7))) = "" And Trim(rawAuthors) <> "" Then oldRec(7) = rawAuthors
        End If
        If UBound(oldRec) >= 8 Then
            oldRec(8) = Mod3_Field_Author.MergeCorrespondingAuthors(CStr(oldRec(8)), corrAuthor)
        End If
        dictRecs(recId) = oldRec
    Else
        recId = dictRecs.Count + 1
        Dim newRec(8) As Variant
        newRec(0) = title
        newRec(1) = journal
        newRec(2) = vol
        newRec(3) = issue
        newRec(4) = matchedAuthors
        newRec(5) = doi
        newRec(6) = srcType
        newRec(7) = rawAuthors
        newRec(8) = corrAuthor
        dictRecs(recId) = newRec
        
        If normDoi <> "" Then dictDoi(normDoi) = recId
        If normT <> "" Then dictTitle(normT) = recId
    End If
End Sub

Public Function MergeSourceTypes(oldType As String, newType As String) As String
    Dim typesDict As Object, p() As String, i As Long
    Set typesDict = CreateObject("Scripting.Dictionary")
    p = Split(oldType, "+")
    For i = LBound(p) To UBound(p)
        If Trim(p(i)) <> "" Then typesDict(Trim(p(i))) = 1
    Next i
    typesDict(newType) = 1
    
    Dim res As String
    res = ""
    If typesDict.Exists("SCI") Then res = res & IIf(res = "", "", "+") & "SCI"
    If typesDict.Exists("EI") Then res = res & IIf(res = "", "", "+") & "EI"
    If typesDict.Exists("中文核心") Then res = res & IIf(res = "", "", "+") & "中文核心"
    If res = "" Then res = newType
    MergeSourceTypes = res
End Function

Public Function GetOrCreateOutputWorkbook(outPath As String, ByRef wsOut As Worksheet, ByRef wsExc As Worksheet) As Workbook
    Dim fso As Object, wb As Workbook, headers As Variant, c As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    headers = Array("序号", "论文题目", "期刊名称", "卷", "期", "作者", "通讯作者", "收录类型", "影响因子")
    
    If fso.FileExists(outPath) Then
        On Error Resume Next
        Set wb = Workbooks(fso.GetFileName(outPath))
        If wb Is Nothing Then Set wb = Workbooks.Open(outPath)
        On Error GoTo 0
    Else
        Set wb = Workbooks.Add
        wb.SaveAs outPath
    End If
    
    ' Sheet 1: 课题组入库成果
    Set wsOut = wb.Sheets(1)
    wsOut.Name = "课题组入库成果"
    For c = 0 To UBound(headers)
        wsOut.Cells(1, c + 1).Value = headers(c)
    Next c
    
    ' Sheet 2: 未认领排除成果
    If wb.Sheets.Count < 2 Then
        Set wsExc = wb.Sheets.Add(After:=wsOut)
    Else
        Set wsExc = wb.Sheets(2)
    End If
    wsExc.Name = "未认领排除成果"
    For c = 0 To UBound(headers)
        wsExc.Cells(1, c + 1).Value = headers(c)
    Next c
    
    Set GetOrCreateOutputWorkbook = wb
End Function

Public Sub FormatOutputSheet(ws As Worksheet, lastRow As Long)
    With ws.Range("A1:I1")
        .Font.Name = "微软雅黑": .Font.Size = 11: .Font.Bold = True: .Font.Color = RGB(24, 76, 120)
        .Interior.Color = RGB(238, 245, 252): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    If lastRow >= 2 Then
        With ws.Range("A2:I" & lastRow)
            .Font.Name = "微软雅黑": .Font.Size = 10: .VerticalAlignment = xlCenter
        End With
        ws.Range("A2:A" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("D2:E" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("H2:I" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("B2:C" & lastRow).HorizontalAlignment = xlLeft
        ws.Range("F2:G" & lastRow).HorizontalAlignment = xlLeft
    End If
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 55
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 8
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 35
    ws.Columns("G").ColumnWidth = 28
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 12
    ws.Activate
    ActiveWindow.DisplayGridlines = True
End Sub