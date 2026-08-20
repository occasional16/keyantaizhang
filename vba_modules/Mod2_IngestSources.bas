' ==============================================================================
' 模块名称: Mod2_IngestSources
' 核心职责: 多源异构文献原始文件解析器 (WOS / EI / CNKI 官方文件高精度抽取)
' ==============================================================================
Option Explicit

Public Function IngestWosData(rawDir As String, dictTeachers As Object, _
                              dictDoi As Object, dictTitle As Object, dictRecs As Object, _
                              hasStart As Boolean, filterStart As Date, hasEnd As Boolean, filterEnd As Date, _
                              ByRef outOfDateCount As Long) As Long
    Dim fso As Object, wosFile As String, isTxt As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    wosFile = rawDir & Application.PathSeparator & "WOS.txt"
    isTxt = True
    
    If Not fso.FileExists(wosFile) Then
        wosFile = rawDir & Application.PathSeparator & "savedrecs.txt"
    End If
    If Not fso.FileExists(wosFile) Then
        wosFile = rawDir & Application.PathSeparator & "WOS.xlsx"
        isTxt = False
    End If
    If Not fso.FileExists(wosFile) Then
        wosFile = rawDir & Application.PathSeparator & "savedrecs.xlsx"
        isTxt = False
    End If
    If Not fso.FileExists(wosFile) Then IngestWosData = 0: Exit Function
    
    Dim count As Long: count = 0
    
    If isTxt Then
        Dim stm As Object, content As String, lines() As String, headers() As String
        Dim r As Long, c As Long, tiCol As Long, soCol As Long, vlCol As Long, isCol As Long, auCol As Long, diCol As Long, pyCol As Long, pdCol As Long
        Dim fields() As String
        Dim rawTitle As String, rawSO As String, rawVL As String, rawIS As String, rawAU As String, rawDOI As String, rawPY As String, rawPD As String
        
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2: stm.Charset = "utf-8": stm.Open
        stm.LoadFromFile wosFile
        content = stm.ReadText(-1)
        stm.Close
        
        lines = Split(content, vbLf)
        If UBound(lines) < 1 Then IngestWosData = 0: Exit Function
        
        headers = Split(Replace(lines(0), vbCr, ""), vbTab)
        tiCol = -1: soCol = -1: vlCol = -1: isCol = -1: auCol = -1: diCol = -1: pyCol = -1: pdCol = -1
        For c = LBound(headers) To UBound(headers)
            Select Case Trim(headers(c))
                Case "TI", "Article Title", "Title": tiCol = c
                Case "SO", "Source Title", "Publication Name": soCol = c
                Case "VL", "Volume": vlCol = c
                Case "IS", "Issue": isCol = c
                Case "AU", "Authors", "Author": auCol = c
                Case "DI", "DOI": diCol = c
                Case "PY", "Publication Year", "Year": pyCol = c
                Case "PD", "Publication Date": pdCol = c
            End Select
        Next c
        
        For r = 1 To UBound(lines)
            If Trim(Replace(lines(r), vbCr, "")) <> "" Then
                fields = Split(Replace(lines(r), vbCr, ""), vbTab)
                rawTitle = "": rawSO = "": rawVL = "": rawIS = "": rawAU = "": rawDOI = "": rawPY = "": rawPD = ""
                If tiCol >= 0 And tiCol <= UBound(fields) Then rawTitle = fields(tiCol)
                If soCol >= 0 And soCol <= UBound(fields) Then rawSO = fields(soCol)
                If vlCol >= 0 And vlCol <= UBound(fields) Then rawVL = fields(vlCol)
                If isCol >= 0 And isCol <= UBound(fields) Then rawIS = fields(isCol)
                If auCol >= 0 And auCol <= UBound(fields) Then rawAU = fields(auCol)
                If diCol >= 0 And diCol <= UBound(fields) Then rawDOI = fields(diCol)
                If pyCol >= 0 And pyCol <= UBound(fields) Then rawPY = fields(pyCol)
                If pdCol >= 0 And pdCol <= UBound(fields) Then rawPD = fields(pdCol)
                
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
        Next r
        IngestWosData = count
        Exit Function
    Else
        ' WOS Excel 备用解析
        Dim wbRaw As Workbook, wsRaw As Worksheet, lr As Long
        Set wbRaw = Workbooks.Open(wosFile, ReadOnly:=True)
        Set wsRaw = wbRaw.Sheets(1)
        For c = 1 To wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
            Select Case Trim(wsRaw.Cells(1, c).Value)
                Case "TI", "Article Title", "Title": tiCol = c
                Case "SO", "Source Title", "Publication Name": soCol = c
                Case "VL", "Volume": vlCol = c
                Case "IS", "Issue": isCol = c
                Case "AU", "Authors", "Author": auCol = c
                Case "DI", "DOI": diCol = c
                Case "PY", "Publication Year", "Year": pyCol = c
                Case "PD", "Publication Date": pdCol = c
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

Public Function IngestEiData(rawDir As String, dictTeachers As Object, _
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

Public Function IngestCnkiData(rawDir As String, dictTeachers As Object, _
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
            Case "Source-文献来源", "文献来源", "期刊": soCol = c
            Case "Volume-卷", "卷": vlCol = c
            Case "Period-期", "期": isCol = c
            Case "Author-作者", "作者": auCol = c
            Case "DOI-DOI", "DOI": diCol = c
            Case "Year-年", "年": pyCol = c
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