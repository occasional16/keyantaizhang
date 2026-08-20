' ==============================================================================
' 模块名称: Mod2_CleanRawData
' 核心职责: 多源原始数据抽取、日期范围过滤、消歧认领、去重与 7 列直出
' 日期规则:
'   - 严格以正式出版年份 (PY / Publication year / Year-年) 与出版日期为准
'   - 若仅有年份，只要年份在指定区间内即视为符合条件
' 成果表头: 序号 | 论文题目 | 期刊名称 | 卷 | 期 | 作者 | 收录类型 (严格 7 列)
' ==============================================================================
Option Explicit

Public Sub 清洗所有原始数据()
    Dim rootDir As String, rawDir As String, configPath As String, outPath As String
    Dim fso As Object, dictTeachers As Object
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
    
    ' 读取控制台上的日期范围设定 (D20 与 I20)
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
        wsOut.Range("A2:G" & wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row).ClearContents
    End If
    If wsExc.Cells(wsExc.Rows.Count, "A").End(xlUp).Row >= 2 Then
        wsExc.Range("A2:G" & wsExc.Cells(wsExc.Rows.Count, "A").End(xlUp).Row).ClearContents
    End If
    
    Dim kVar As Variant, rec As Variant, rawAuthorStr As String
    acceptedCount = 0
    excludedCount = 0
    
    For Each kVar In dictRecs.Keys
        rec = dictRecs(kVar)
        If Trim(CStr(rec(4))) <> "" Then
            ' 入库成果 -> 写入 Sheet 1
            acceptedCount = acceptedCount + 1
            wsOut.Cells(acceptedCount + 1, 1).Value = acceptedCount
            wsOut.Cells(acceptedCount + 1, 2).Value = rec(0)
            wsOut.Cells(acceptedCount + 1, 3).Value = rec(1)
            wsOut.Cells(acceptedCount + 1, 4).Value = rec(2)
            wsOut.Cells(acceptedCount + 1, 5).Value = rec(3)
            wsOut.Cells(acceptedCount + 1, 6).Value = rec(4)
            wsOut.Cells(acceptedCount + 1, 7).Value = rec(6)
        Else
            ' 未认领/排除成果 -> 写入 Sheet 2
            excludedCount = excludedCount + 1
            rawAuthorStr = ""
            If UBound(rec) >= 7 Then rawAuthorStr = CStr(rec(7))
            
            wsExc.Cells(excludedCount + 1, 1).Value = excludedCount
            wsExc.Cells(excludedCount + 1, 2).Value = rec(0)
            wsExc.Cells(excludedCount + 1, 3).Value = rec(1)
            wsExc.Cells(excludedCount + 1, 4).Value = rec(2)
            wsExc.Cells(excludedCount + 1, 5).Value = rec(3)
            wsExc.Cells(excludedCount + 1, 6).Value = rawAuthorStr
            wsExc.Cells(excludedCount + 1, 7).Value = rec(6)
        End If
    Next kVar
    
    Call FormatOutputSheet(wsOut, acceptedCount + 1)
    Call FormatOutputSheet(wsExc, excludedCount + 1)
    wbOut.Save
    
    ' 5. 刷新控制台看板与日志
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "【整理完成】在期入库 " & acceptedCount & " 篇，非本室 " & excludedCount & " 篇，超期排除 " & outOfDateCount & " 篇！"
    On Error GoTo 0
End Sub

Private Function LoadTeacherAliasDictionary(configPath As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    Dim wb As Workbook, isAlreadyOpen As Boolean, ws As Worksheet, lastRow As Long, r As Long
    Dim cName As String, allAlias As String, arrAlias() As String, i As Long, aItem As String
    
    On Error Resume Next
    Set wb = Workbooks("teachers_profile.xlsx")
    If Not wb Is Nothing Then isAlreadyOpen = True Else Set wb = Workbooks.Open(configPath, ReadOnly:=True): isAlreadyOpen = False
    On Error GoTo 0
    If wb Is Nothing Then Set LoadTeacherAliasDictionary = d: Exit Function
    
    Set ws = wb.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    For r = 2 To lastRow
        cName = Trim(ws.Cells(r, "D").Value)
        allAlias = Trim(ws.Cells(r, "N").Value)
        If cName <> "" Then
            d(LCase(cName)) = cName
            If allAlias <> "" Then
                arrAlias = Split(allAlias, ";")
                For i = LBound(arrAlias) To UBound(arrAlias)
                    aItem = LCase(Trim(arrAlias(i)))
                    If aItem <> "" Then
                        d(aItem) = cName
                        d(Replace(aItem, ".", "")) = cName
                        d(Replace(Replace(aItem, ".", ""), ",", "")) = cName
                        d(Replace(Replace(Replace(aItem, ".", ""), ",", ""), " ", "")) = cName
                    End If
                Next i
            End If
        End If
    Next r
    
    If Not isAlreadyOpen And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadTeacherAliasDictionary = d
End Function

' ------------------------------------------------------------------------------
' 辅助: 判定论文出版日期是否在设定范围内
' ------------------------------------------------------------------------------
Public Function IsPaperInDateRange(pubYearStr As String, pubDateStr As String, _
                                   hasStart As Boolean, filterStart As Date, _
                                   hasEnd As Boolean, filterEnd As Date) As Boolean
    If Not hasStart And Not hasEnd Then
        IsPaperInDateRange = True
        Exit Function
    End If
    
    Dim pStart As Date, pEnd As Date
    Call ParsePaperDateSpan(pubYearStr, pubDateStr, pStart, pEnd)
    
    If pStart = #1/1/1900# Then
        IsPaperInDateRange = True ' 无法解析日期的默认保留
        Exit Function
    End If
    
    If hasStart Then
        If pEnd < filterStart Then
            IsPaperInDateRange = False
            Exit Function
        End If
    End If
    
    If hasEnd Then
        If pStart > filterEnd Then
            IsPaperInDateRange = False
            Exit Function
        End If
    End If
    
    IsPaperInDateRange = True
End Function

Private Sub ParsePaperDateSpan(pubYearStr As String, pubDateStr As String, ByRef pStart As Date, ByRef pEnd As Date)
    Dim y As Long, m As Long, d As Long
    Dim sYear As String, sDate As String
    
    pStart = #1/1/1900#: pEnd = #1/1/1900#
    sYear = Trim(pubYearStr)
    sDate = LCase(Trim(pubDateStr))
    
    y = 0: m = 0: d = 0
    
    ' 提取 4 位年份
    If Len(sYear) >= 4 Then
        If IsNumeric(Left(sYear, 4)) Then y = CLng(Left(sYear, 4))
    End If
    If y = 0 And Len(sDate) >= 4 Then
        Dim regY As Object, matchesY As Object
        Set regY = CreateObject("VBScript.RegExp")
        regY.Pattern = "\b(19\d\d|20\d\d)\b"
        If regY.Test(sDate) Then
            Set matchesY = regY.Execute(sDate)
            y = CLng(matchesY(0).Value)
        End If
    End If
    
    If y = 0 Then Exit Sub
    
    ' 提取月份 (英文缩写或数字)
    If InStr(sDate, "jan") > 0 Then m = 1
    If InStr(sDate, "feb") > 0 Then m = 2
    If InStr(sDate, "mar") > 0 Then m = 3
    If InStr(sDate, "apr") > 0 Then m = 4
    If InStr(sDate, "may") > 0 Then m = 5
    If InStr(sDate, "jun") > 0 Then m = 6
    If InStr(sDate, "jul") > 0 Then m = 7
    If InStr(sDate, "aug") > 0 Then m = 8
    If InStr(sDate, "sep") > 0 Then m = 9
    If InStr(sDate, "oct") > 0 Then m = 10
    If InStr(sDate, "nov") > 0 Then m = 11
    If InStr(sDate, "dec") > 0 Then m = 12
    
    ' 提取 YYYY-MM-DD
    Dim regD As Object, matchesD As Object
    Set regD = CreateObject("VBScript.RegExp")
    regD.Pattern = "(\d{4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})"
    If regD.Test(sDate) Then
        Set matchesD = regD.Execute(sDate)
        m = CLng(matchesD(0).SubMatches(1))
        d = CLng(matchesD(0).SubMatches(2))
    End If
    
    If m > 0 And d > 0 Then
        On Error Resume Next
        pStart = DateSerial(y, m, d)
        pEnd = pStart
        On Error GoTo 0
        If pStart <> #1/1/1900# Then Exit Sub
    End If
    
    If m > 0 Then
        On Error Resume Next
        pStart = DateSerial(y, m, 1)
        pEnd = DateSerial(y, m + 1, 0)
        On Error GoTo 0
        If pStart <> #1/1/1900# Then Exit Sub
    End If
    
    ' 仅年份: 覆盖当年 1月1日 至 12月31日
    pStart = DateSerial(y, 1, 1)
    pEnd = DateSerial(y, 12, 31)
End Sub

Private Sub AddOrMergeRecord(ByVal title As String, ByVal journal As String, ByVal vol As String, ByVal issue As String, _
                             ByVal matchedAuthors As String, ByVal rawAuthors As String, ByVal doi As String, ByVal srcType As String, _
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
        dictRecs(recId) = oldRec
    Else
        recId = dictRecs.Count + 1
        Dim newRec(7) As Variant
        newRec(0) = title
        newRec(1) = journal
        newRec(2) = vol
        newRec(3) = issue
        newRec(4) = matchedAuthors
        newRec(5) = doi
        newRec(6) = srcType
        newRec(7) = rawAuthors
        dictRecs(recId) = newRec
        
        If normDoi <> "" Then dictDoi(normDoi) = recId
        If normT <> "" Then dictTitle(normT) = recId
    End If
End Sub

Private Function MergeSourceTypes(oldType As String, newType As String) As String
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

Private Function IngestWosData(rawDir As String, dictTeachers As Object, _
                               dictDoi As Object, dictTitle As Object, dictRecs As Object, _
                               hasStart As Boolean, filterStart As Date, hasEnd As Boolean, filterEnd As Date, _
                               ByRef outOfDateCount As Long) As Long
    Dim fso As Object, txtFile As String, xlsxFile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    txtFile = rawDir & Application.PathSeparator & "WOS.txt"
    If Not fso.FileExists(txtFile) Then txtFile = rawDir & Application.PathSeparator & "savedrecs.txt"
    xlsxFile = rawDir & Application.PathSeparator & "WOS.xlsx"
    If Not fso.FileExists(xlsxFile) Then xlsxFile = rawDir & Application.PathSeparator & "savedrecs.xlsx"
    
    Dim count As Long: count = 0
    
    If fso.FileExists(txtFile) Then
        Dim stm As Object, content As String, lines() As String, headers() As String
        Dim tiIdx As Long, soIdx As Long, vlIdx As Long, isIdx As Long, auIdx As Long, diIdx As Long, pyIdx As Long, pdIdx As Long
        Dim i As Long, j As Long, parts() As String
        Dim rawTitle As String, rawSO As String, rawVL As String, rawIS As String, rawAU As String, rawDOI As String, rawPY As String, rawPD As String
        
        tiIdx = -1: soIdx = -1: vlIdx = -1: isIdx = -1: auIdx = -1: diIdx = -1: pyIdx = -1: pdIdx = -1
        
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2: stm.Charset = "utf-8": stm.Open
        stm.LoadFromFile txtFile
        content = stm.ReadText(-1)
        stm.Close
        
        lines = Split(content, vbLf)
        If UBound(lines) < 1 Then IngestWosData = 0: Exit Function
        
        headers = Split(Replace(lines(0), vbCr, ""), vbTab)
        For j = LBound(headers) To UBound(headers)
            Select Case Trim(headers(j))
                Case "TI", "Article Title": tiIdx = j
                Case "SO", "Source Title": soIdx = j
                Case "VL", "Volume": vlIdx = j
                Case "IS", "Issue": isIdx = j
                Case "AU", "Authors", "Author Full Names": If auIdx = -1 Then auIdx = j
                Case "DI", "DOI": diIdx = j
                Case "PY", "Publication Year": pyIdx = j
                Case "PD", "Publication Date": pdIdx = j
            End Select
        Next j
        
        For i = 1 To UBound(lines)
            Dim lText As String
            lText = Replace(lines(i), vbCr, "")
            If Trim(lText) <> "" Then
                parts = Split(lText, vbTab)
                rawTitle = "": rawSO = "": rawVL = "": rawIS = "": rawAU = "": rawDOI = "": rawPY = "": rawPD = ""
                If tiIdx >= 0 And tiIdx <= UBound(parts) Then rawTitle = parts(tiIdx)
                If soIdx >= 0 And soIdx <= UBound(parts) Then rawSO = parts(soIdx)
                If vlIdx >= 0 And vlIdx <= UBound(parts) Then rawVL = parts(vlIdx)
                If isIdx >= 0 And isIdx <= UBound(parts) Then rawIS = parts(isIdx)
                If auIdx >= 0 And auIdx <= UBound(parts) Then rawAU = parts(auIdx)
                If diIdx >= 0 And diIdx <= UBound(parts) Then rawDOI = parts(diIdx)
                If pyIdx >= 0 And pyIdx <= UBound(parts) Then rawPY = parts(pyIdx)
                If pdIdx >= 0 And pdIdx <= UBound(parts) Then rawPD = parts(pdIdx)
                
                If Trim(rawTitle) <> "" Then
                    count = count + 1
                    If IsPaperInDateRange(rawPY, rawPD, hasStart, filterStart, hasEnd, filterEnd) Then
                        Call AddOrMergeRecord(CleanPaperTitle(rawTitle), ConvertToTitleCase(rawSO), _
                                              Trim(rawVL), Trim(rawIS), ExtractLabAuthors(rawAU, dictTeachers), _
                                              Trim(rawAU), Trim(rawDOI), "SCI", dictDoi, dictTitle, dictRecs)
                    Else
                        outOfDateCount = outOfDateCount + 1
                    End If
                End If
            End If
        Next i
        IngestWosData = count
        Exit Function
    ElseIf fso.FileExists(xlsxFile) Then
        Dim wbRaw As Workbook, wsRaw As Worksheet, lr As Long, r As Long, c As Long
        Dim tiCol As Long, soCol As Long, vlCol As Long, isCol As Long, auCol As Long, diCol As Long, pyCol As Long, pdCol As Long
        Set wbRaw = Workbooks.Open(xlsxFile, ReadOnly:=True)
        Set wsRaw = wbRaw.Sheets(1)
        For c = 1 To wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
            Select Case Trim(wsRaw.Cells(1, c).Value)
                Case "Article Title", "TI": tiCol = c
                Case "Source Title", "SO": soCol = c
                Case "Volume", "VL": vlCol = c
                Case "Issue", "IS": isCol = c
                Case "Authors", "Author Full Names", "AU": If auCol = 0 Then auCol = c
                Case "DOI", "DI": diCol = c
                Case "Publication Year", "PY": pyCol = c
                Case "Publication Date", "PD": pdCol = c
            End Select
        Next c
        lr = wsRaw.Cells(wsRaw.Rows.Count, IIf(tiCol > 0, tiCol, 1)).End(xlUp).Row
        For r = 2 To lr
            rawTitle = "": rawSO = "": rawVL = "": rawIS = "": rawAU = "": rawDOI = "": rawPY = "": rawPD = ""
            If tiCol > 0 Then rawTitle = wsRaw.Cells(r, tiCol).Value
            If soCol > 0 Then rawSO = wsRaw.Cells(r, soCol).Value
            If vlCol > 0 Then rawVL = wsRaw.Cells(r, vlCol).Value
            If isCol > 0 Then rawIS = wsRaw.Cells(r, isCol).Value
            If auCol > 0 Then rawAU = wsRaw.Cells(r, auCol).Value
            If diCol > 0 Then rawDOI = wsRaw.Cells(r, diCol).Value
            If pyCol > 0 Then rawPY = wsRaw.Cells(r, pyCol).Value
            If pdCol > 0 Then rawPD = wsRaw.Cells(r, pdCol).Value
            If Trim(rawTitle) <> "" Then
                count = count + 1
                If IsPaperInDateRange(rawPY, rawPD, hasStart, filterStart, hasEnd, filterEnd) Then
                    Call AddOrMergeRecord(CleanPaperTitle(rawTitle), ConvertToTitleCase(rawSO), _
                                          Trim(rawVL), Trim(rawIS), ExtractLabAuthors(rawAU, dictTeachers), _
                                          Trim(rawAU), Trim(rawDOI), "SCI", dictDoi, dictTitle, dictRecs)
                Else
                    outOfDateCount = outOfDateCount + 1
                End If
            End If
        Next r
        wbRaw.Close SaveChanges:=False
        IngestWosData = count
        Exit Function
    End If
    IngestWosData = 0
End Function

Private Function IngestEiData(rawDir As String, dictTeachers As Object, _
                              dictDoi As Object, dictTitle As Object, dictRecs As Object, _
                              hasStart As Boolean, filterStart As Date, hasEnd As Boolean, filterEnd As Date, _
                              ByRef outOfDateCount As Long) As Long
    Dim fso As Object, eiFile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    eiFile = rawDir & Application.PathSeparator & "EI.xlsx"
    If Not fso.FileExists(eiFile) Then eiFile = rawDir & Application.PathSeparator & "EI.xls"
    If Not fso.FileExists(eiFile) Then eiFile = rawDir & Application.PathSeparator & "EI.csv"
    If Not fso.FileExists(eiFile) Then eiFile = rawDir & Application.PathSeparator & "ei_raw.xlsx"
    If Not fso.FileExists(eiFile) Then IngestEiData = 0: Exit Function
    
    Dim wbRaw As Workbook, wsRaw As Worksheet, lr As Long, r As Long, c As Long
    Dim tiCol As Long, soCol As Long, vlCol As Long, isCol As Long, auCol As Long, diCol As Long, pyCol As Long, pdCol As Long
    Dim rawTitle As String, rawSO As String, rawVL As String, rawIS As String, rawAU As String, rawDOI As String, rawPY As String, rawPD As String
    Dim count As Long: count = 0
    
    Set wbRaw = Workbooks.Open(eiFile, ReadOnly:=True)
    Set wsRaw = wbRaw.Sheets(1)
    For c = 1 To wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
        Select Case Trim(wsRaw.Cells(1, c).Value)
            Case "Title": tiCol = c
            Case "Source": soCol = c
            Case "Volume": vlCol = c
            Case "Issue": isCol = c
            Case "Author": auCol = c
            Case "DOI": diCol = c
            Case "Publication year": pyCol = c
            Case "Issue date": pdCol = c
        End Select
    Next c
    
    lr = wsRaw.Cells(wsRaw.Rows.Count, IIf(tiCol > 0, tiCol, 1)).End(xlUp).Row
    For r = 2 To lr
        rawTitle = "": rawSO = "": rawVL = "": rawIS = "": rawAU = "": rawDOI = "": rawPY = "": rawPD = ""
        If tiCol > 0 Then rawTitle = wsRaw.Cells(r, tiCol).Value
        If soCol > 0 Then rawSO = wsRaw.Cells(r, soCol).Value
        If vlCol > 0 Then rawVL = wsRaw.Cells(r, vlCol).Value
        If isCol > 0 Then rawIS = wsRaw.Cells(r, isCol).Value
        If auCol > 0 Then rawAU = wsRaw.Cells(r, auCol).Value
        If diCol > 0 Then rawDOI = wsRaw.Cells(r, diCol).Value
        If pyCol > 0 Then rawPY = wsRaw.Cells(r, pyCol).Value
        If pdCol > 0 Then rawPD = wsRaw.Cells(r, pdCol).Value
        
        If Trim(rawTitle) <> "" Then
            count = count + 1
            If IsPaperInDateRange(rawPY, rawPD, hasStart, filterStart, hasEnd, filterEnd) Then
                Call AddOrMergeRecord(CleanPaperTitle(rawTitle), ConvertToTitleCase(rawSO), _
                                      Trim(rawVL), Trim(rawIS), ExtractLabAuthors(rawAU, dictTeachers), _
                                      Trim(rawAU), Trim(rawDOI), "EI", dictDoi, dictTitle, dictRecs)
            Else
                outOfDateCount = outOfDateCount + 1
            End If
        End If
    Next r
    wbRaw.Close SaveChanges:=False
    IngestEiData = count
End Function

Private Function IngestCnkiData(rawDir As String, dictTeachers As Object, _
                                dictDoi As Object, dictTitle As Object, dictRecs As Object, _
                                hasStart As Boolean, filterStart As Date, hasEnd As Boolean, filterEnd As Date, _
                                ByRef outOfDateCount As Long) As Long
    Dim fso As Object, cnkiFile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    cnkiFile = rawDir & Application.PathSeparator & "CNKI.xls"
    If Not fso.FileExists(cnkiFile) Then cnkiFile = rawDir & Application.PathSeparator & "CNKI.xlsx"
    If Not fso.FileExists(cnkiFile) Then cnkiFile = rawDir & Application.PathSeparator & "cnki_raw.xlsx"
    If Not fso.FileExists(cnkiFile) Then cnkiFile = rawDir & Application.PathSeparator & "cnki_raw.xls"
    If Not fso.FileExists(cnkiFile) Then IngestCnkiData = 0: Exit Function
    
    Dim wbRaw As Workbook, wsRaw As Worksheet, lr As Long, r As Long, c As Long
    Dim tiCol As Long, soCol As Long, vlCol As Long, isCol As Long, auCol As Long, diCol As Long, pyCol As Long, pdCol As Long
    Dim rawTitle As String, rawSO As String, rawVL As String, rawIS As String, rawAU As String, rawDOI As String, rawPY As String, rawPD As String
    Dim count As Long: count = 0
    
    Set wbRaw = Workbooks.Open(cnkiFile, ReadOnly:=True)
    Set wsRaw = wbRaw.Sheets(1)
    For c = 1 To wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
        Select Case Trim(wsRaw.Cells(1, c).Value)
            Case "Title-题名", "题名", "Title": tiCol = c
            Case "Source-文献来源", "文献来源", "Source": soCol = c
            Case "Volume-卷", "卷", "Volume": vlCol = c
            Case "Period-期", "期", "Issue": isCol = c
            Case "Author-作者", "作者", "Author": auCol = c
            Case "DOI-DOI", "DOI": diCol = c
            Case "Year-年", "年", "Year": pyCol = c
            Case "PubTime-发表时间", "发表时间": pdCol = c
        End Select
    Next c
    
    lr = wsRaw.Cells(wsRaw.Rows.Count, IIf(tiCol > 0, tiCol, 1)).End(xlUp).Row
    For r = 2 To lr
        rawTitle = "": rawSO = "": rawVL = "": rawIS = "": rawAU = "": rawDOI = "": rawPY = "": rawPD = ""
        If tiCol > 0 Then rawTitle = wsRaw.Cells(r, tiCol).Value
        If soCol > 0 Then rawSO = wsRaw.Cells(r, soCol).Value
        If vlCol > 0 Then rawVL = wsRaw.Cells(r, vlCol).Value
        If isCol > 0 Then rawIS = wsRaw.Cells(r, isCol).Value
        If auCol > 0 Then rawAU = wsRaw.Cells(r, auCol).Value
        If diCol > 0 Then rawDOI = wsRaw.Cells(r, diCol).Value
        If pyCol > 0 Then rawPY = wsRaw.Cells(r, pyCol).Value
        If pdCol > 0 Then rawPD = wsRaw.Cells(r, pdCol).Value
        
        If Trim(rawTitle) <> "" Then
            count = count + 1
            If IsPaperInDateRange(rawPY, rawPD, hasStart, filterStart, hasEnd, filterEnd) Then
                Call AddOrMergeRecord(CleanPaperTitle(rawTitle), Trim(rawSO), _
                                      Trim(rawVL), Trim(rawIS), ExtractLabAuthors(rawAU, dictTeachers), _
                                      Trim(rawAU), Trim(rawDOI), "中文核心", dictDoi, dictTitle, dictRecs)
            Else
                outOfDateCount = outOfDateCount + 1
            End If
        End If
    Next r
    wbRaw.Close SaveChanges:=False
    IngestCnkiData = count
End Function

Private Function NormalizeTitleForMatch(ByVal t As String) As String
    Dim res As String, i As Long, ch As String, chCode As Long
    t = LCase(Trim(t))
    res = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        chCode = AscW(ch)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or (chCode >= &H4E00 And chCode <= &H9FA5) Then
            res = res & ch
        End If
    Next i
    NormalizeTitleForMatch = res
End Function

Private Function CleanPaperTitle(ByVal t As String) As String
    t = Replace(t, vbCr, " "): t = Replace(t, vbLf, " "): t = Replace(t, vbTab, " "): t = Trim(t)
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    Do While Right(t, 1) = ".": t = Trim(Left(t, Len(t) - 1)): Loop
    CleanPaperTitle = t
End Function

Public Function ConvertToTitleCase(ByVal s As String) As String
    s = Trim(s)
    If s = "" Then ConvertToTitleCase = "": Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If AscW(Mid(s, i, 1)) >= &H4E00 Then ConvertToTitleCase = s: Exit Function
    Next i
    Dim words() As String, w As String, wLower As String, wUpper As String, res As String
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    words = Split(s, " ")
    For i = LBound(words) To UBound(words)
        w = Trim(words(i)): wLower = LCase(w): wUpper = UCase(w)
        If IsAcronym(wUpper) Then
            w = wUpper
        ElseIf (i > LBound(words) And i < UBound(words)) And IsMinorWord(wLower) Then
            w = wLower
        Else
            If Len(w) = 1 Then
                w = UCase(w)
            ElseIf Len(w) > 1 Then
                If InStr(w, "-") > 0 Then
                    Dim subW() As String, k As Long
                    subW = Split(w, "-")
                    For k = LBound(subW) To UBound(subW)
                        If Len(subW(k)) > 0 Then
                            If IsAcronym(UCase(subW(k))) Then
                                subW(k) = UCase(subW(k))
                            ElseIf k > 0 And IsMinorWord(LCase(subW(k))) Then
                                subW(k) = LCase(subW(k))
                            Else
                                subW(k) = UCase(Left(subW(k), 1)) & LCase(Mid(subW(k), 2))
                            End If
                        End If
                    Next k
                    w = Join(subW, "-")
                Else
                    w = UCase(Left(w, 1)) & LCase(Mid(w, 2))
                End If
            End If
        End If
        If res = "" Then res = w Else res = res & " " & w
    Next i
    ConvertToTitleCase = res
End Function

Private Function IsMinorWord(w As String) As Boolean
    Select Case w
        Case "of", "in", "and", "the", "on", "for", "to", "at", "by", "with", "from", _
             "into", "onto", "via", "a", "an", "as", "nor", "but", "part", "section"
            IsMinorWord = True
        Case Else: IsMinorWord = False
    End Select
End Function

Private Function IsAcronym(w As String) As Boolean
    Select Case w
        Case "IEEE", "ASME", "ACM", "CFD", "AI", "PINN", "FEM", "CNC", "MEMS", "NEMS", _
             "IC", "LED", "OLED", "UV", "NMR", "SEM", "TEM", "AFM", "XPS", "XRD", _
             "PANI", "PDMS", "UHMWPE", "PBO", "SOFC", "DES", "LES", "RANS", _
             "1D", "2D", "3D", "4D", "5D", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"
            IsAcronym = True
        Case Else: IsAcronym = False
    End Select
End Function

Private Function ExtractLabAuthors(rawAuthorsStr As String, dictTeachers As Object) As String
    If Trim(rawAuthorsStr) = "" Then ExtractLabAuthors = "": Exit Function
    Dim sClean As String: sClean = CleanEiAffiliationTags(rawAuthorsStr)
    Dim arr() As String, i As Long, auth As String, authNorm As String, authClean As String, cTeacher As String
    Dim seenTeachers As Object, res As String
    Set seenTeachers = CreateObject("Scripting.Dictionary")
    arr = Split(sClean, ";")
    For i = LBound(arr) To UBound(arr)
        auth = Trim(arr(i))
        If auth <> "" Then
            authNorm = LCase(auth)
            authClean = Replace(Replace(Replace(authNorm, ".", ""), ",", ""), " ", "")
            cTeacher = ""
            If dictTeachers.Exists(authNorm) Then
                cTeacher = dictTeachers(authNorm)
            ElseIf dictTeachers.Exists(Replace(authNorm, ".", "")) Then
                cTeacher = dictTeachers(Replace(authNorm, ".", ""))
            ElseIf dictTeachers.Exists(authClean) Then
                cTeacher = dictTeachers(authClean)
            End If
            If cTeacher <> "" Then
                If Not seenTeachers.Exists(cTeacher) Then
                    seenTeachers(cTeacher) = 1
                    If res = "" Then res = auth Else res = res & "; " & auth
                End If
            End If
        End If
    Next i
    ExtractLabAuthors = res
End Function

Private Function CleanEiAffiliationTags(ByVal s As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True: reg.IgnoreCase = True
    reg.Pattern = "\(\s*[\d\s,]+\s*\)"
    CleanEiAffiliationTags = reg.Replace(s, "")
End Function

Private Function GetOrCreateOutputWorkbook(outPath As String, ByRef wsOut As Worksheet, ByRef wsExc As Worksheet) As Workbook
    Dim fso As Object, wb As Workbook, headers As Variant, c As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    headers = Array("序号", "论文题目", "期刊名称", "卷", "期", "作者", "收录类型")
    
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

Private Sub FormatOutputSheet(ws As Worksheet, lastRow As Long)
    With ws.Range("A1:G1")
        .Font.Name = "微软雅黑": .Font.Size = 11: .Font.Bold = True: .Font.Color = RGB(24, 76, 120)
        .Interior.Color = RGB(238, 245, 252): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    If lastRow >= 2 Then
        With ws.Range("A2:G" & lastRow)
            .Font.Name = "微软雅黑": .Font.Size = 10: .VerticalAlignment = xlCenter
        End With
        ws.Range("A2:A" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("D2:E" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("G2:G" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("B2:C" & lastRow).HorizontalAlignment = xlLeft
        ws.Range("F2:F" & lastRow).HorizontalAlignment = xlLeft
    End If
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 55
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 8
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 35
    ws.Columns("G").ColumnWidth = 16
    ws.Activate
    ActiveWindow.DisplayGridlines = True
End Sub