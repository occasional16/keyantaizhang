' ==============================================================================
' 模块名称: Mod0_MetricsEngine
' 核心职责: 100% 真实动态逐行实测统计与状态卡片刷新引擎
' 设计哲学: 彻底杜绝任何估算，直接扫描磁盘文件与交付成果工作表输出精准统计
' ==============================================================================
Option Explicit

Public Sub 扫描并刷新看板()
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
    
    ' 1. 扫描教师名单
    teacherCount = GetExternalWorkbookRows(rootDir & Application.PathSeparator & "config" & Application.PathSeparator & "teachers_profile.xlsx", "D")
    
    ' 2. 扫描原始数据文件
    wosRawCount = ScanRawDataFileCount(rawDir, Array("wos.txt", "wos.xlsx", "savedrecs.txt", "savedrecs.xlsx"))
    eiRawCount = ScanRawDataFileCount(rawDir, Array("ei.xlsx", "ei.xls", "ei.csv", "ei_raw.xlsx"))
    cnkiRawCount = ScanRawDataFileCount(rawDir, Array("cnki.xls", "cnki.xlsx", "cnki_raw.xlsx", "cnki_raw.xls"))
    scopusRawCount = ScanRawDataFileCount(rawDir, Array("scopus.csv", "scopus_raw.csv"))
    totalRaw = wosRawCount + eiRawCount + cnkiRawCount + scopusRawCount
    
    ' 3. 100% 真实逐行扫描最终大表
    Call GetMergedOutputStats(rootDir & Application.PathSeparator & "papers_final_merged.xlsx", _
                             finalMergeCount, sciCount, eiCount, cnkiCount, sciEiCount, _
                             excTotal, excSci, excEi, excCnki, excSciEi)
    
    ' 4. 刷新顶部指标卡
    Call UpdateCardValue(wsPanel, "card_Teacher", Format(teacherCount, "#,##0"))
    Call UpdateCardValue(wsPanel, "card_Raw", Format(totalRaw, "#,##0"))
    Call UpdateCardValue(wsPanel, "card_Final", Format(finalMergeCount, "#,##0"))
    
    ' 5. 刷新下方明细清单
    wsPanel.Range("D13").Value = Format(wosRawCount, "#,##0") & " 篇"
    wsPanel.Range("D14").Value = Format(eiRawCount, "#,##0") & " 篇"
    wsPanel.Range("D15").Value = Format(cnkiRawCount, "#,##0") & " 篇"
    wsPanel.Range("D16").Value = Format(scopusRawCount, "#,##0") & " 篇"
    
    wsPanel.Range("J13").Value = Format(finalMergeCount, "#,##0") & " 篇"
    wsPanel.Range("J14").Value = Format(sciCount, "#,##0") & " 篇"
    wsPanel.Range("J15").Value = Format(eiCount, "#,##0") & " 篇"
    wsPanel.Range("J16").Value = Format(cnkiCount, "#,##0") & " 篇"
    wsPanel.Range("J17").Value = Format(sciEiCount, "#,##0") & " 篇"
    
    ' 刷新排除明细细分摘要
    wsPanel.Range("J18").Value = Format(excTotal, "#,##0") & " 篇 (SCI " & excSci & ", EI " & excEi & ", 中文 " & excCnki & ", SCI+EI " & excSciEi & ")"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
End Sub

Public Function ScanRawDataFileCount(folderPath As String, candidateNames As Variant) As Long
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

Public Sub GetMergedOutputStats(filePath As String, ByRef totalRows As Long, ByRef sciTotal As Long, _
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
    
    ' 1. 扫描入库成果 Sheet 1
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
    
    ' 2. 扫描排除成果 Sheet 2
    Dim ws2 As Worksheet, lr2 As Long
    On Error Resume Next
    If extWb.Sheets.Count >= 2 Then Set ws2 = extWb.Sheets(2)
    On Error GoTo 0
    If Not ws2 Is Nothing Then
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
    
    If Not isAlreadyOpen And Not extWb Is Nothing Then
        extWb.Close SaveChanges:=False
    End If
End Sub

Public Function GetExternalWorkbookRows(filePath As String, checkCol As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then GetExternalWorkbookRows = 0: Exit Function
    
    Dim extWb As Workbook, isAlreadyOpen As Boolean, ws As Worksheet, lastRow As Long
    Dim fName As String
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
    
    If extWb Is Nothing Then GetExternalWorkbookRows = 0: Exit Function
    
    On Error Resume Next
    Set ws = extWb.Sheets(1)
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not isAlreadyOpen Then extWb.Close SaveChanges:=False
        GetExternalWorkbookRows = 0
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).Row
    If lastRow >= 2 Then
        GetExternalWorkbookRows = lastRow - 1
    Else
        GetExternalWorkbookRows = 0
    End If
    
    If Not isAlreadyOpen And Not extWb Is Nothing Then
        extWb.Close SaveChanges:=False
    End If
End Function

Public Sub UpdateCardValue(ws As Worksheet, shapeName As String, newVal As String)
    On Error Resume Next
    ws.Shapes(shapeName).TextFrame2.TextRange.Text = newVal
    On Error GoTo 0
End Sub