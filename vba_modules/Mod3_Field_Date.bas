' ==============================================================================
' 模块名称: Mod3_Field_Date
' 核心职责: 正式出版年与日期跨度解析、区间有效性与跨年发表严格判定
' ==============================================================================
Option Explicit

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
        IsPaperInDateRange = True ' 无法解析时间时保守保留
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

Public Sub ParsePaperDateSpan(pubYearStr As String, pubDateStr As String, ByRef pStart As Date, ByRef pEnd As Date)
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
    
    ' 提取月份 (支持英文与数字)
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
    
    If m = 0 Then
        Dim regM As Object, matchesM As Object
        Set regM = CreateObject("VBScript.RegExp")
        regM.Pattern = "\b([1-9]|1[0-2])\b"
        If regM.Test(sDate) Then
            Set matchesM = regM.Execute(sDate)
            m = CLng(matchesM(0).Value)
        End If
    End If
    
    ' 提取日期
    If m > 0 Then
        Dim regD As Object, matchesD As Object
        Set regD = CreateObject("VBScript.RegExp")
        regD.Pattern = "\b([1-9]|[12]\d|3[01])\b"
        If regD.Test(sDate) Then
            Set matchesD = regD.Execute(sDate)
            d = CLng(matchesD(0).Value)
            If d = m And matchesD.Count > 1 Then d = CLng(matchesD(1).Value)
        End If
    End If
    
    If m > 0 And d > 0 Then
        On Error Resume Next
        pStart = DateSerial(y, m, d)
        pEnd = DateSerial(y, m, d)
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
    
    ' 仅有年份: 覆盖当年 1月1日 至 12月31日
    pStart = DateSerial(y, 1, 1)
    pEnd = DateSerial(y, 12, 31)
End Sub