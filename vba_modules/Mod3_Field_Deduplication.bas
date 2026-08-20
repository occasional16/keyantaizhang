' ==============================================================================
' 模块名称: Mod3_Field_Deduplication
' 核心职责: 题目清洗、归一化字符提取与跨库主备去重键生成
' ==============================================================================
Option Explicit

Public Function CleanPaperTitle(ByVal t As String) As String
    t = Trim(t)
    Do While Right(t, 1) = "."
        t = Trim(Left(t, Len(t) - 1))
    Loop
    CleanPaperTitle = t
End Function

Public Function NormalizeTitleForMatch(ByVal t As String) As String
    Dim s As String, i As Long, ch As String, res As String
    s = LCase(Trim(t))
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or (AscW(ch) > 255) Then
            res = res & ch
        End If
    Next i
    NormalizeTitleForMatch = res
End Function