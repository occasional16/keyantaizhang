' ==============================================================================
' 模块名称: Mod3_Field_JournalIF
' 核心职责: 期刊名称 Title Case 规范化与期刊影响因子 (JIF) 高速匹配引擎
' ==============================================================================
Option Explicit

Public Function LoadJournalIfDictionary(configPath As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then
        Set LoadJournalIfDictionary = d
        Exit Function
    End If
    
    Dim wb As Workbook, ws As Worksheet, isAlreadyOpen As Boolean
    Dim lastRow As Long, dataArr As Variant, r As Long
    Dim jName As String, ifVal As String, normKey As String
    
    On Error Resume Next
    Set wb = Workbooks(fso.GetFileName(configPath))
    If wb Is Nothing Then
        Set wb = Workbooks.Open(configPath, ReadOnly:=True, AddToMru:=False)
        isAlreadyOpen = False
    Else
        isAlreadyOpen = True
    End If
    On Error GoTo 0
    
    If wb Is Nothing Then
        Set LoadJournalIfDictionary = d
        Exit Function
    End If
    
    On Error Resume Next
    Set ws = wb.Sheets(1)
    On Error GoTo 0
    If ws Is Nothing Then
        If Not isAlreadyOpen Then wb.Close SaveChanges:=False
        Set LoadJournalIfDictionary = d
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        dataArr = ws.Range("A2:G" & lastRow).Value2
        For r = 1 To UBound(dataArr, 1)
            jName = Trim(CStr(dataArr(r, 1)))
            ifVal = Trim(CStr(dataArr(r, 7)))
            If jName <> "" And ifVal <> "" And UCase(ifVal) <> "N/A" Then
                normKey = NormalizeJournalForIfMatch(jName)
                If normKey <> "" Then
                    If Not d.Exists(normKey) Then
                        d(normKey) = ifVal
                    End If
                End If
            End If
        Next r
    End If
    
    If Not isAlreadyOpen And Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    
    Set LoadJournalIfDictionary = d
End Function

Public Function NormalizeJournalForIfMatch(ByVal jName As String) As String
    Dim s As String
    s = UCase(Trim(jName))
    If s = "" Then NormalizeJournalForIfMatch = "": Exit Function
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    s = Replace(s, ":", "")
    s = Replace(s, ";", "")
    s = Replace(s, "'", "")
    s = Replace(s, """", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, "&", "")
    s = Replace(s, "_", "")
    NormalizeJournalForIfMatch = s
End Function

Public Function GetJournalIfValue(ByVal journalName As String, dictIF As Object) As Variant
    If dictIF Is Nothing Then GetJournalIfValue = "": Exit Function
    If dictIF.Count = 0 Then GetJournalIfValue = "": Exit Function
    
    Dim normKey As String
    normKey = NormalizeJournalForIfMatch(journalName)
    If normKey <> "" Then
        If dictIF.Exists(normKey) Then
            Dim v As String
            v = dictIF(normKey)
            If IsNumeric(v) Then
                GetJournalIfValue = CDbl(v)
            Else
                GetJournalIfValue = v
            End If
            Exit Function
        End If
    End If
    GetJournalIfValue = ""
End Function

Public Function ConvertToTitleCase(ByVal s As String) As String
    Dim words() As String, i As Long, w As String, res As String
    s = Trim(s)
    If s = "" Then ConvertToTitleCase = "": Exit Function
    
    ' 过滤中文期刊：包含中文字符直接返回
    If Len(s) <> LenB(StrConv(s, vbFromUnicode)) Then
        ConvertToTitleCase = s
        Exit Function
    End If
    
    s = Replace(s, "-", " - ")
    s = Replace(s, ":", " : ")
    s = Replace(s, "/", " / ")
    
    words = Split(s, " ")
    res = ""
    
    For i = LBound(words) To UBound(words)
        w = Trim(words(i))
        If w <> "" Then
            If IsAcronym(w) Then
                w = UCase(w)
            ElseIf i > LBound(words) And IsMinorWord(LCase(w)) Then
                w = LCase(w)
            Else
                w = UCase(Left(w, 1)) & LCase(Mid(w, 2))
            End If
            If res = "" Then res = w Else res = res & " " & w
        End If
    Next i
    
    res = Replace(res, " - ", "-")
    res = Replace(res, " : ", ": ")
    res = Replace(res, " / ", "/")
    ConvertToTitleCase = res
End Function

Public Function IsMinorWord(w As String) As Boolean
    Dim minors As Variant, m As Variant
    minors = Array("a", "an", "the", "and", "but", "or", "for", "nor", "on", "at", "to", "from", "by", "of", "in", "with")
    For Each m In minors
        If w = m Then IsMinorWord = True: Exit Function
    Next m
    IsMinorWord = False
End Function

Public Function IsAcronym(w As String) As Boolean
    Dim acronyms As Variant, a As Variant
    acronyms = Array("IEEE", "ASME", "ACM", "AIP", "IOP", "SPIE", "MDPI", "SCI", "EI", "CNKI", "USA", "UK", "CFD", "FEA", "MEMS", "NEMS")
    For Each a In acronyms
        If UCase(w) = a Then IsAcronym = True: Exit Function
    Next a
    IsAcronym = False
End Function