' ==============================================================================
' 模块名称: Mod3_Field_Author
' 核心职责: 作者与通讯作者字段清洗、机构角标剥离、别名库消歧与本组成员提取
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
            cTeacher = MatchTeacherName(pureAuth, dictTeachers)
            
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

Public Function ExtractLabCorrespondingAuthors(rawStr As String, srcType As String, dictTeachers As Object) As String
    If dictTeachers Is Nothing Then ExtractLabCorrespondingAuthors = "": Exit Function
    If dictTeachers.Count = 0 Then ExtractLabCorrespondingAuthors = "": Exit Function
    
    Dim s As String, res As String
    Dim seenTeachers As Object
    Set seenTeachers = CreateObject("Scripting.Dictionary")
    res = ""
    s = Trim(rawStr)
    If s = "" Then ExtractLabCorrespondingAuthors = "": Exit Function
    
    Select Case srcType
        Case "SCI"
            ' WOS 格式: "Jin, J (corresponding author), Tsinghua Univ..." 或 "Tian, Y; Ma, LR (corresponding author)..."
            Dim regWos As Object, matchesWos As Object, authorBlock As String
            Set regWos = CreateObject("VBScript.RegExp")
            regWos.IgnoreCase = True: regWos.Global = False
            regWos.Pattern = "^(.*?)\s*\(\s*corresponding author\s*\)"
            
            If regWos.Test(s) Then
                Set matchesWos = regWos.Execute(s)
                authorBlock = matchesWos(0).SubMatches(0)
            Else
                ' 备用直接提取逗号或分号前的人名
                authorBlock = s
            End If
            
            Dim arrWos() As String, iW As Long, authW As String, cTeacherW As String
            arrWos = Split(Replace(authorBlock, " and ", "; "), ";")
            For iW = LBound(arrWos) To UBound(arrWos)
                authW = Trim(arrWos(iW))
                If InStr(authW, ",") > 0 Or InStr(authW, " ") > 0 Then
                    cTeacherW = MatchTeacherName(LCase(authW), dictTeachers)
                    If cTeacherW <> "" Then
                        If Not seenTeachers.Exists(cTeacherW) Then
                            seenTeachers(cTeacherW) = 1
                            If res = "" Then res = authW & "(" & cTeacherW & ")" Else res = res & "; " & authW & "(" & cTeacherW & ")"
                        End If
                    End If
                End If
            Next iW
            
        Case "EI"
            ' EI 格式: "Lu, Xinchun(xclu@tsinghua.edu.cn)" 或多位以分号分隔
            Dim regEmail As Object, cleanEiStr As String
            Set regEmail = CreateObject("VBScript.RegExp")
            regEmail.Global = True: regEmail.IgnoreCase = True
            regEmail.Pattern = "\s*\([^\)]*?\)"
            cleanEiStr = regEmail.Replace(s, "")
            
            Dim arrEi() As String, iE As Long, authE As String, cTeacherE As String
            arrEi = Split(Replace(cleanEiStr, " and ", "; "), ";")
            For iE = LBound(arrEi) To UBound(arrEi)
                authE = Trim(arrEi(iE))
                If authE <> "" Then
                    cTeacherE = MatchTeacherName(LCase(authE), dictTeachers)
                    If cTeacherE <> "" Then
                        If Not seenTeachers.Exists(cTeacherE) Then
                            seenTeachers(cTeacherE) = 1
                            If res = "" Then res = authE & "(" & cTeacherE & ")" Else res = res & "; " & authE & "(" & cTeacherE & ")"
                        End If
                    End If
                End If
            Next iE
            
        Case "中文核心"
            ' 知网导师或中文通讯作者格式: "邵天敏" -> "邵天敏(导师)"
            Dim arrCnki() As String, iC As Long, authC As String, cTeacherC As String
            arrCnki = Split(Replace(Replace(Replace(s, "，", ";"), "、", ";"), " ", ";"), ";")
            For iC = LBound(arrCnki) To UBound(arrCnki)
                authC = Trim(arrCnki(iC))
                If authC <> "" Then
                    cTeacherC = MatchTeacherName(LCase(authC), dictTeachers)
                    If cTeacherC <> "" Then
                        If Not seenTeachers.Exists(cTeacherC) Then
                            seenTeachers(cTeacherC) = 1
                            If res = "" Then res = cTeacherC & "(导师)" Else res = res & "; " & cTeacherC & "(导师)"
                        End If
                    End If
                End If
            Next iC
    End Select
    
    ExtractLabCorrespondingAuthors = res
End Function

Private Function MatchTeacherName(pureAuth As String, dictTeachers As Object) As String
    If dictTeachers.Exists(pureAuth) Then
        MatchTeacherName = dictTeachers(pureAuth)
        Exit Function
    End If
    If dictTeachers.Exists(Replace(pureAuth, ".", "")) Then
        MatchTeacherName = dictTeachers(Replace(pureAuth, ".", ""))
        Exit Function
    End If
    If dictTeachers.Exists(Replace(Replace(pureAuth, ".", ""), ",", "")) Then
        MatchTeacherName = dictTeachers(Replace(Replace(pureAuth, ".", ""), ",", ""))
        Exit Function
    End If
    If dictTeachers.Exists(Replace(Replace(Replace(pureAuth, ".", ""), ",", ""), " ", "")) Then
        MatchTeacherName = dictTeachers(Replace(Replace(Replace(pureAuth, ".", ""), ",", ""), " ", ""))
        Exit Function
    End If
    MatchTeacherName = ""
End Function

Public Function CleanEiAffiliationTags(ByVal s As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True: reg.IgnoreCase = True
    reg.Pattern = "\(\s*[\d\s,]+\s*\)"
    CleanEiAffiliationTags = reg.Replace(s, "")
End Function
Public Function MergeCorrespondingAuthors(oldCA As String, newCA As String) As String
    Dim d As Object, allArr() As String, i As Long, item As String, res As String
    Dim cName As String, leftP As Long, rightP As Long, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    
    If Trim(oldCA) = "" And Trim(newCA) = "" Then MergeCorrespondingAuthors = "": Exit Function
    If Trim(oldCA) = "" Then MergeCorrespondingAuthors = Trim(newCA): Exit Function
    If Trim(newCA) = "" Then MergeCorrespondingAuthors = Trim(oldCA): Exit Function
    
    allArr = Split(oldCA & "; " & newCA, ";")
    For i = LBound(allArr) To UBound(allArr)
        item = Trim(allArr(i))
        If item <> "" Then
            leftP = InStr(item, "(")
            rightP = InStr(item, ")")
            If leftP > 0 And rightP > leftP Then
                cName = Mid(item, leftP + 1, rightP - leftP - 1)
                If Not d.Exists(cName) Then
                    d(cName) = item
                Else
                    If Len(item) > Len(CStr(d(cName))) Then
                        d(cName) = item
                    End If
                End If
            Else
                If Not d.Exists(item) Then d(item) = item
            End If
        End If
    Next i
    
    res = ""
    For Each k In d.Keys
        If res = "" Then res = CStr(d(k)) Else res = res & "; " & CStr(d(k))
    Next k
    MergeCorrespondingAuthors = res
End Function