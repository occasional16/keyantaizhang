' ==============================================================================
' 模块名称: Mod3_Field_Author
' 核心职责: 作者字段深度清洗、机构角标剥离、别名库载入与本组师生消歧认领
' ==============================================================================
Option Explicit

Public Function LoadTeacherAliasDictionary(configPath As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    Dim wb As Workbook, isAlreadyOpen As Boolean, ws As Worksheet, lastRow As Long, r As Long
    Dim cName As String, allAlias As String, arrAlias() As String, i As Long, aItem As String
    
    On Error Resume Next
    Set wb = Workbooks("teachers_profile.xlsx")
    If Not wb Is Nothing Then
        isAlreadyOpen = True
    Else
        Set wb = Workbooks.Open(configPath, ReadOnly:=True, AddToMru:=False)
        isAlreadyOpen = False
    End If
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

Public Function ExtractLabAuthors(rawAuthorsStr As String, dictTeachers As Object) As String
    If dictTeachers Is Nothing Then ExtractLabAuthors = "": Exit Function
    If dictTeachers.Count = 0 Then ExtractLabAuthors = "": Exit Function
    
    Dim s As String, arr() As String, i As Long, auth As String
    Dim pureAuth As String, cTeacher As String, res As String
    Dim seenTeachers As Object
    Set seenTeachers = CreateObject("Scripting.Dictionary")
    
    s = CleanEiAffiliationTags(rawAuthorsStr)
    s = Replace(s, "，", ";")
    s = Replace(s, "、", ";")
    s = Replace(s, " and ", "; ")
    s = Replace(s, " AND ", "; ")
    
    arr = Split(s, ";")
    res = ""
    
    For i = LBound(arr) To UBound(arr)
        auth = Trim(arr(i))
        If auth <> "" Then
            pureAuth = LCase(auth)
            cTeacher = ""
            If dictTeachers.Exists(pureAuth) Then
                cTeacher = dictTeachers(pureAuth)
            ElseIf dictTeachers.Exists(Replace(pureAuth, ".", "")) Then
                cTeacher = dictTeachers(Replace(pureAuth, ".", ""))
            ElseIf dictTeachers.Exists(Replace(Replace(pureAuth, ".", ""), ",", "")) Then
                cTeacher = dictTeachers(Replace(Replace(pureAuth, ".", ""), ",", ""))
            ElseIf dictTeachers.Exists(Replace(Replace(Replace(pureAuth, ".", ""), ",", ""), " ", "")) Then
                cTeacher = dictTeachers(Replace(Replace(Replace(pureAuth, ".", ""), ",", ""), " ", ""))
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

Public Function CleanEiAffiliationTags(ByVal s As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True: reg.IgnoreCase = True
    reg.Pattern = "\(\s*[\d\s,]+\s*\)"
    CleanEiAffiliationTags = reg.Replace(s, "")
End Function