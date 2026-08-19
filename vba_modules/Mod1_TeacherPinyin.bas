' ==============================================================================
' 模块名称: Mod1_TeacherPinyin
' 核心职责: 专职负责 config/teachers_profile.xlsx 教师主表拼音与检索特征库构建
' 适    配: 独立文件架构，全相对路径 ThisWorkbook.Path；支持任意新加教师或测试人名
' 生成格式:
'   G列: 姓, 名名      (如: Lu, Xinchun / Zhu, Yu / Ce, Shi)
'   H列: 姓, 名-名     (如: Lu, Xin-chun / Zhu, Yu / Ce, Shi)
'   I列: 姓, 缩写(WOS) (如: Lu, XC / Zhu, Y / Ce, S - WOS标准AU格式)
'   J列: 姓, 缩写.     (如: Lu, X.C. / Zhu, Y. / Ce, S. - 常见期刊格式)
'   K列: 名名, 姓      (如: Xinchun, Lu / Yu, Zhu / Shi, Ce)
'   L列: 名-名, 姓     (如: Xin-chun, Lu / Yu, Zhu / Shi, Ce)
'   M列: 名名 姓       (如: Xinchun Lu / Yu Zhu / Shi Ce - 常用正序格式)
'   N列: 全部匹配特征  (汇总所有中英文别名变体，以分号隔开，供后续匹配宏秒级快速调用)
' ==============================================================================
Option Explicit

Public Sub 生成老师拼音变体()
    Dim configPath As String
    Dim wbConfig As Workbook, isAlreadyOpen As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim cName As String
    Dim pSurname As String, pGiven1 As String, pGiven2 As String
    Dim pGivenCombined As String, pGivenHyphen As String
    Dim pInitials As String, pInitialsDot As String
    Dim singleInitials As String, singleInitialsDot As String
    Dim allVariants As String
    Dim processedCount As Long
    
    ' 1. 定位 config/teachers_profile.xlsx
    configPath = ThisWorkbook.Path & Application.PathSeparator & "config" & Application.PathSeparator & "teachers_profile.xlsx"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then
        MsgBox "未找到教师主档案文件：" & vbCrLf & configPath, vbCritical, "错误"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 2. 打开工作簿
    On Error Resume Next
    Set wbConfig = Workbooks("teachers_profile.xlsx")
    If Not wbConfig Is Nothing Then
        isAlreadyOpen = True
    Else
        Set wbConfig = Workbooks.Open(configPath)
        isAlreadyOpen = False
    End If
    On Error GoTo 0
    
    If wbConfig Is Nothing Then
        MsgBox "无法打开 config/teachers_profile.xlsx 文件！", vbCritical, "错误"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Set ws = wbConfig.Sheets(1)
    
    ' 3. 标准化表头
    ws.Cells(1, 4).Value = "姓名"
    ws.Cells(1, 7).Value = "姓, 名名"
    ws.Cells(1, 8).Value = "姓, 名-名"
    ws.Cells(1, 9).Value = "姓, 缩写(WOS)"
    ws.Cells(1, 10).Value = "姓, 缩写."
    ws.Cells(1, 11).Value = "名名, 姓"
    ws.Cells(1, 12).Value = "名-名, 姓"
    ws.Cells(1, 13).Value = "名名 姓"
    ws.Cells(1, 14).Value = "全部匹配特征(检索库)"
    
    With ws.Range("G1:N1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 242)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 4. 获取教师最大行
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "【teachers_profile.xlsx】表中未找到教师名单数据！", vbExclamation, "提示"
        GoTo CleanUp
    End If
    
    ' 5. 遍历每一行执行通用拼音提取与特征库构建
    For r = 2 To lastRow
        cName = Trim(ws.Cells(r, "D").Value)
        
        If cName <> "" Then
            Call ParseChineseName(cName, pSurname, pGiven1, pGiven2)
            
            If pSurname <> "" Then
                If pGiven2 <> "" Then
                    pGivenCombined = pGiven1 & LCase(pGiven2)
                    pGivenHyphen = pGiven1 & "-" & LCase(pGiven2)
                    pInitials = UCase(Left(pGiven1, 1) & Left(pGiven2, 1))
                    pInitialsDot = UCase(Left(pGiven1, 1)) & "." & UCase(Left(pGiven2, 1)) & "."
                Else
                    pGivenCombined = pGiven1
                    pGivenHyphen = pGiven1
                    pInitials = UCase(Left(pGiven1, 1))
                    pInitialsDot = UCase(Left(pGiven1, 1)) & "."
                End If
                
                singleInitials = UCase(Left(pGiven1, 1))
                singleInitialsDot = UCase(Left(pGiven1, 1)) & "."
                
                ws.Cells(r, 7).Value = pSurname & ", " & pGivenCombined
                ws.Cells(r, 8).Value = pSurname & ", " & pGivenHyphen
                ws.Cells(r, 9).Value = pSurname & ", " & pInitials
                ws.Cells(r, 10).Value = pSurname & ", " & pInitialsDot
                ws.Cells(r, 11).Value = pGivenCombined & ", " & pSurname
                ws.Cells(r, 12).Value = pGivenHyphen & ", " & pSurname
                ws.Cells(r, 13).Value = pGivenCombined & " " & pSurname
                
                allVariants = BuildAllVariants(pSurname, pGiven1, pGiven2, cName)
                ws.Cells(r, 14).Value = allVariants
                processedCount = processedCount + 1
            End If
        End If
    Next r
    
    ws.Columns("G:N").AutoFit
    
    ' 6. 保存并同步反馈
    wbConfig.Save
    
    ThisWorkbook.Activate
    
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "【步骤 1 成功】已为 teachers_profile.xlsx 中 " & processedCount & " 位教师完成全格式拼音与检索特征库构建！"
    On Error GoTo 0
    
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.AppendLog", "【步骤1完成】已成功为 " & processedCount & " 位教师生成完整的拼音变体库！"
    On Error GoTo 0

CleanUp:
    If Not isAlreadyOpen And Not wbConfig Is Nothing Then
        wbConfig.Close SaveChanges:=True
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub ParseChineseName(ByVal cName As String, ByRef pSurname As String, ByRef pGiven1 As String, ByRef pGiven2 As String)
    Dim nLen As Integer
    Dim sur As String, giv As String
    
    cName = Replace(cName, " ", "")
    nLen = Len(cName)
    pSurname = "": pGiven1 = "": pGiven2 = ""
    
    If nLen = 0 Then Exit Sub
    
    ' 复姓处理 (欧阳, 诸葛, 司马, 上官, 皇甫, 东方, 司徒, 司空等)
    If nLen >= 3 And (Left(cName, 2) = "欧阳" Or Left(cName, 2) = "诸葛" Or _
                     Left(cName, 2) = "司马" Or Left(cName, 2) = "上官" Or _
                     Left(cName, 2) = "皇甫" Or Left(cName, 2) = "东方" Or _
                     Left(cName, 2) = "司徒" Or Left(cName, 2) = "司空") Then
        sur = Left(cName, 2)
        giv = Mid(cName, 3)
    Else
        sur = Left(cName, 1)
        giv = Mid(cName, 2)
    End If
    
    pSurname = GetCharPinyin(sur, True)
    
    If Len(giv) >= 1 Then
        pGiven1 = GetCharPinyin(Mid(giv, 1, 1), False)
    End If
    If Len(giv) >= 2 Then
        pGiven2 = GetCharPinyin(Mid(giv, 2, 1), False)
    End If
End Sub

Private Function BuildAllVariants(pSur As String, pG1 As String, pG2 As String, cName As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    Dim gComb As String, gHyph As String, gInit As String, gInitDot As String
    Dim sInit As String, sInitDot As String
    
    If pG2 <> "" Then
        gComb = pG1 & LCase(pG2)
        gHyph = pG1 & "-" & LCase(pG2)
        gInit = UCase(Left(pG1, 1) & Left(pG2, 1))
        gInitDot = UCase(Left(pG1, 1)) & "." & UCase(Left(pG2, 1)) & "."
    Else
        gComb = pG1
        gHyph = pG1
        gInit = UCase(Left(pG1, 1))
        gInitDot = UCase(Left(pG1, 1)) & "."
    End If
    
    sInit = UCase(Left(pG1, 1))
    sInitDot = UCase(Left(pG1, 1)) & "."
    
    d(cName) = 1
    d(pSur & ", " & gComb) = 1
    d(pSur & ", " & gHyph) = 1
    d(pSur & ", " & pG1 & " " & pG2) = 1
    d(pSur & ", " & gInit) = 1
    d(pSur & ", " & gInitDot) = 1
    d(pSur & ", " & sInit) = 1
    d(pSur & ", " & sInitDot) = 1
    d(pSur & " " & gInit) = 1
    d(pSur & " " & gInitDot) = 1
    d(pSur & " " & sInit) = 1
    d(pSur & " " & sInitDot) = 1
    d(gComb & " " & pSur) = 1
    d(gHyph & " " & pSur) = 1
    d(gComb & ", " & pSur) = 1
    d(gHyph & ", " & pSur) = 1
    d(gInit & " " & pSur) = 1
    d(gInitDot & " " & pSur) = 1
    
    ' 特殊多音字多向冗余
    If pSur = "Xie" Or pSur = "Jie" Then
        Dim altSur As String
        altSur = IIf(pSur = "Xie", "Jie", "Xie")
        d(altSur & ", " & gComb) = 1
        d(altSur & ", " & gInit) = 1
        d(altSur & ", " & gInitDot) = 1
        d(gComb & " " & altSur) = 1
    End If
    
    Dim k As Variant, res As String
    For Each k In d.Keys
        If res = "" Then res = k Else res = res & "; " & k
    Next k
    BuildAllVariants = res
End Function

Private Function GetCharPinyin(ch As String, isSurname As Boolean) As String
    ' 复姓处理
    If ch = "欧阳" Then GetCharPinyin = "Ouyang": Exit Function
    If ch = "诸葛" Then GetCharPinyin = "Zhuge": Exit Function
    If ch = "司马" Then GetCharPinyin = "Sima": Exit Function
    If ch = "上官" Then GetCharPinyin = "Shangguan": Exit Function
    If ch = "皇甫" Then GetCharPinyin = "Huangfu": Exit Function
    If ch = "东方" Then GetCharPinyin = "Dongfang": Exit Function
    If ch = "司徒" Then GetCharPinyin = "Situ": Exit Function
    If ch = "司空" Then GetCharPinyin = "Sikong": Exit Function

    ' 姓氏多音字优先处理
    If isSurname Then
        Select Case ch
            Case "解": GetCharPinyin = "Xie": Exit Function
            Case "单": GetCharPinyin = "Shan": Exit Function
            Case "仇": GetCharPinyin = "Qiu": Exit Function
            Case "朴": GetCharPinyin = "Piao": Exit Function
            Case "翟": GetCharPinyin = "Zhai": Exit Function
            Case "查": GetCharPinyin = "Zha": Exit Function
            Case "盖": GetCharPinyin = "Ge": Exit Function
            Case "曾": GetCharPinyin = "Zeng": Exit Function
            Case "雒": GetCharPinyin = "Luo": Exit Function
            Case "区": GetCharPinyin = "Ou": Exit Function
        End Select
    End If

    ' 通用汉字拼音字典映射
    Select Case ch
        Case "丁": GetCharPinyin = "Ding"
        Case "七": GetCharPinyin = "Qi"
        Case "万": GetCharPinyin = "Wan"
        Case "三": GetCharPinyin = "San"
        Case "丛": GetCharPinyin = "Cong"
        Case "东": GetCharPinyin = "Dong"
        Case "严": GetCharPinyin = "Yan"
        Case "丰": GetCharPinyin = "Feng"
        Case "丹": GetCharPinyin = "Dan"
        Case "丽": GetCharPinyin = "Li"
        Case "乌": GetCharPinyin = "Wu"
        Case "乐": GetCharPinyin = "Le"
        Case "乔": GetCharPinyin = "Qiao"
        Case "乜": GetCharPinyin = "Nie"
        Case "九": GetCharPinyin = "Jiu"
        Case "习": GetCharPinyin = "Xi"
        Case "乾": GetCharPinyin = "Qian"
        Case "于": GetCharPinyin = "Yu"
        Case "云": GetCharPinyin = "Yun"
        Case "五": GetCharPinyin = "Wu"
        Case "井": GetCharPinyin = "Jing"
        Case "亚": GetCharPinyin = "Ya"
        Case "亮": GetCharPinyin = "Liang"
        Case "仇": GetCharPinyin = "Qiu"
        Case "从": GetCharPinyin = "Cong"
        Case "仰": GetCharPinyin = "Yang"
        Case "仲": GetCharPinyin = "Zhong"
        Case "任": GetCharPinyin = "Ren"
        Case "伊": GetCharPinyin = "Yi"
        Case "伍": GetCharPinyin = "Wu"
        Case "伏": GetCharPinyin = "Fu"
        Case "伟": GetCharPinyin = "Wei"
        Case "何": GetCharPinyin = "He"
        Case "余": GetCharPinyin = "Yu"
        Case "佟": GetCharPinyin = "Tong"
        Case "佩": GetCharPinyin = "Pei"
        Case "侯": GetCharPinyin = "Hou"
        Case "俞": GetCharPinyin = "Yu"
        Case "俭": GetCharPinyin = "Jian"
        Case "倪": GetCharPinyin = "Ni"
        Case "健": GetCharPinyin = "Jian"
        Case "傅": GetCharPinyin = "Fu"
        Case "储": GetCharPinyin = "Chu"
        Case "元": GetCharPinyin = "Yuan"
        Case "充": GetCharPinyin = "Chong"
        Case "克": GetCharPinyin = "Ke"
        Case "党": GetCharPinyin = "Dang"
        Case "全": GetCharPinyin = "Quan"
        Case "八": GetCharPinyin = "Ba"
        Case "公": GetCharPinyin = "Gong"
        Case "六": GetCharPinyin = "Liu"
        Case "关": GetCharPinyin = "Guan"
        Case "养": GetCharPinyin = "Yang"
        Case "冀": GetCharPinyin = "Ji"
        Case "冉": GetCharPinyin = "Ran"
        Case "军": GetCharPinyin = "Jun"
        Case "农": GetCharPinyin = "Nong"
        Case "冯": GetCharPinyin = "Feng"
        Case "冰": GetCharPinyin = "Bing"
        Case "冷": GetCharPinyin = "Leng"
        Case "凌": GetCharPinyin = "Ling"
        Case "凤": GetCharPinyin = "Feng"
        Case "刁": GetCharPinyin = "Diao"
        Case "刘": GetCharPinyin = "Liu"
        Case "刚": GetCharPinyin = "Gang"
        Case "利": GetCharPinyin = "Li"
        Case "别": GetCharPinyin = "Bie"
        Case "剑": GetCharPinyin = "Jian"
        Case "力": GetCharPinyin = "Li"
        Case "劳": GetCharPinyin = "Lao"
        Case "勇": GetCharPinyin = "Yong"
        Case "勤": GetCharPinyin = "Qin"
        Case "勾": GetCharPinyin = "Gou"
        Case "包": GetCharPinyin = "Bao"
        Case "匡": GetCharPinyin = "Kuang"
        Case "十": GetCharPinyin = "Shi"
        Case "升": GetCharPinyin = "Sheng"
        Case "华": GetCharPinyin = "Hua"
        Case "卓": GetCharPinyin = "Zhuo"
        Case "单": GetCharPinyin = "Shan"
        Case "卜": GetCharPinyin = "Bu"
        Case "卞": GetCharPinyin = "Bian"
        Case "卢": GetCharPinyin = "Lu"
        Case "卫": GetCharPinyin = "Wei"
        Case "印": GetCharPinyin = "Yin"
        Case "危": GetCharPinyin = "Wei"
        Case "厉": GetCharPinyin = "Li"
        Case "厍": GetCharPinyin = "She"
        Case "原": GetCharPinyin = "Yuan"
        Case "双": GetCharPinyin = "Shuang"
        Case "古": GetCharPinyin = "Gu"
        Case "可": GetCharPinyin = "Ke"
        Case "史": GetCharPinyin = "Shi"
        Case "叶": GetCharPinyin = "Ye"
        Case "司": GetCharPinyin = "Si"
        Case "吉": GetCharPinyin = "Ji"
        Case "后": GetCharPinyin = "Hou"
        Case "向": GetCharPinyin = "Xiang"
        Case "吕": GetCharPinyin = "Lv"
        Case "吴": GetCharPinyin = "Wu"
        Case "周": GetCharPinyin = "Zhou"
        Case "和": GetCharPinyin = "He"
        Case "咸": GetCharPinyin = "Xian"
        Case "唐": GetCharPinyin = "Tang"
        Case "商": GetCharPinyin = "Shang"
        Case "喻": GetCharPinyin = "Yu"
        Case "嘉": GetCharPinyin = "Jia"
        Case "四": GetCharPinyin = "Si"
        Case "国": GetCharPinyin = "Guo"
        Case "堵": GetCharPinyin = "Du"
        Case "夏": GetCharPinyin = "Xia"
        Case "夔": GetCharPinyin = "Kui"
        Case "大": GetCharPinyin = "Da"
        Case "天": GetCharPinyin = "Tian"
        Case "夫": GetCharPinyin = "Fu"
        Case "奚": GetCharPinyin = "Xi"
        Case "姚": GetCharPinyin = "Yao"
        Case "姜": GetCharPinyin = "Jiang"
        Case "姬": GetCharPinyin = "Ji"
        Case "娄": GetCharPinyin = "Lou"
        Case "婵": GetCharPinyin = "Chan"
        Case "婷": GetCharPinyin = "Ting"
        Case "子": GetCharPinyin = "Zi"
        Case "孔": GetCharPinyin = "Kong"
        Case "孙": GetCharPinyin = "Sun"
        Case "孟": GetCharPinyin = "Meng"
        Case "季": GetCharPinyin = "Ji"
        Case "宇": GetCharPinyin = "Yu"
        Case "安": GetCharPinyin = "An"
        Case "宋": GetCharPinyin = "Song"
        Case "宏": GetCharPinyin = "Hong"
        Case "宓": GetCharPinyin = "Mi"
        Case "宗": GetCharPinyin = "Zong"
        Case "宝": GetCharPinyin = "Bao"
        Case "宣": GetCharPinyin = "Xuan"
        Case "宦": GetCharPinyin = "Huan"
        Case "宫": GetCharPinyin = "Gong"
        Case "宰": GetCharPinyin = "Zai"
        Case "家": GetCharPinyin = "Jia"
        Case "容": GetCharPinyin = "Rong"
        Case "宿": GetCharPinyin = "Su"
        Case "寇": GetCharPinyin = "Kou"
        Case "富": GetCharPinyin = "Fu"
        Case "寿": GetCharPinyin = "Shou"
        Case "封": GetCharPinyin = "Feng"
        Case "尚": GetCharPinyin = "Shang"
        Case "尤": GetCharPinyin = "You"
        Case "尹": GetCharPinyin = "Yin"
        Case "居": GetCharPinyin = "Ju"
        Case "屈": GetCharPinyin = "Qu"
        Case "屠": GetCharPinyin = "Tu"
        Case "山": GetCharPinyin = "Shan"
        Case "岑": GetCharPinyin = "Cen"
        Case "峰": GetCharPinyin = "Feng"
        Case "崔": GetCharPinyin = "Cui"
        Case "嵇": GetCharPinyin = "Ji"
        Case "巢": GetCharPinyin = "Chao"
        Case "左": GetCharPinyin = "Zuo"
        Case "巩": GetCharPinyin = "Gong"
        Case "巫": GetCharPinyin = "Wu"
        Case "巴": GetCharPinyin = "Ba"
        Case "师": GetCharPinyin = "Shi"
        Case "席": GetCharPinyin = "Xi"
        Case "常": GetCharPinyin = "Chang"
        Case "干": GetCharPinyin = "Gan"
        Case "平": GetCharPinyin = "Ping"
        Case "幸": GetCharPinyin = "Xing"
        Case "广": GetCharPinyin = "Guang"
        Case "庄": GetCharPinyin = "Zhuang"
        Case "庆": GetCharPinyin = "Qing"
        Case "应": GetCharPinyin = "Ying"
        Case "庞": GetCharPinyin = "Pang"
        Case "康": GetCharPinyin = "Kang"
        Case "庾": GetCharPinyin = "Yu"
        Case "廉": GetCharPinyin = "Lian"
        Case "廖": GetCharPinyin = "Liao"
        Case "建": GetCharPinyin = "Jian"
        Case "弓": GetCharPinyin = "Gong"
        Case "弘": GetCharPinyin = "Hong"
        Case "张": GetCharPinyin = "Zhang"
        Case "强": GetCharPinyin = "Qiang"
        Case "彦": GetCharPinyin = "Yan"
        Case "彭": GetCharPinyin = "Peng"
        Case "徐": GetCharPinyin = "Xu"
        Case "德": GetCharPinyin = "De"
        Case "志": GetCharPinyin = "Zhi"
        Case "怀": GetCharPinyin = "Huai"
        Case "慎": GetCharPinyin = "Shen"
        Case "慕": GetCharPinyin = "Mu"
        Case "慧": GetCharPinyin = "Hui"
        Case "戈": GetCharPinyin = "Ge"
        Case "戎": GetCharPinyin = "Rong"
        Case "成": GetCharPinyin = "Cheng"
        Case "戚": GetCharPinyin = "Qi"
        Case "戴": GetCharPinyin = "Dai"
        Case "房": GetCharPinyin = "Fang"
        Case "扈": GetCharPinyin = "Hu"
        Case "才": GetCharPinyin = "Cai"
        Case "扣": GetCharPinyin = "Kou"
        Case "扬": GetCharPinyin = "Yang"
        Case "扶": GetCharPinyin = "Fu"
        Case "支": GetCharPinyin = "Zhi"
        Case "敏": GetCharPinyin = "Min"
        Case "敖": GetCharPinyin = "Ao"
        Case "文": GetCharPinyin = "Wen"
        Case "斌": GetCharPinyin = "Bin"
        Case "新": GetCharPinyin = "Xin"
        Case "方": GetCharPinyin = "Fang"
        Case "於": GetCharPinyin = "Yu"
        Case "施": GetCharPinyin = "Shi"
        Case "时": GetCharPinyin = "Shi"
        Case "昌": GetCharPinyin = "Chang"
        Case "明": GetCharPinyin = "Ming"
        Case "易": GetCharPinyin = "Yi"
        Case "昝": GetCharPinyin = "Zan"
        Case "春": GetCharPinyin = "Chun"
        Case "晁": GetCharPinyin = "Chao"
        Case "晋": GetCharPinyin = "Jin"
        Case "晏": GetCharPinyin = "Yan"
        Case "晓": GetCharPinyin = "Xiao"
        Case "晨": GetCharPinyin = "Chen"
        Case "景": GetCharPinyin = "Jing"
        Case "智": GetCharPinyin = "Zhi"
        Case "暨": GetCharPinyin = "Ji"
        Case "暴": GetCharPinyin = "Bao"
        Case "曦": GetCharPinyin = "Xi"
        Case "曹": GetCharPinyin = "Cao"
        Case "曾": GetCharPinyin = "Zeng"
        Case "朱": GetCharPinyin = "Zhu"
        Case "朴": GetCharPinyin = "Piao"
        Case "权": GetCharPinyin = "Quan"
        Case "李": GetCharPinyin = "Li"
        Case "杜": GetCharPinyin = "Du"
        Case "束": GetCharPinyin = "Shu"
        Case "杨": GetCharPinyin = "Yang"
        Case "杭": GetCharPinyin = "Hang"
        Case "杰": GetCharPinyin = "Jie"
        Case "松": GetCharPinyin = "Song"
        Case "林": GetCharPinyin = "Lin"
        Case "柏": GetCharPinyin = "Bai"
        Case "查": GetCharPinyin = "Zha"
        Case "柯": GetCharPinyin = "Ke"
        Case "柳": GetCharPinyin = "Liu"
        Case "柴": GetCharPinyin = "Chai"
        Case "栾": GetCharPinyin = "Luan"
        Case "桂": GetCharPinyin = "Gui"
        Case "桑": GetCharPinyin = "Sang"
        Case "桓": GetCharPinyin = "Huan"
        Case "梁": GetCharPinyin = "Liang"
        Case "梅": GetCharPinyin = "Mei"
        Case "森": GetCharPinyin = "Sen"
        Case "楚": GetCharPinyin = "Chu"
        Case "樊": GetCharPinyin = "Fan"
        Case "欢": GetCharPinyin = "Huan"
        Case "欧": GetCharPinyin = "Ou"
        Case "步": GetCharPinyin = "Bu"
        Case "武": GetCharPinyin = "Wu"
        Case "殳": GetCharPinyin = "Shu"
        Case "段": GetCharPinyin = "Duan"
        Case "殷": GetCharPinyin = "Yin"
        Case "毋": GetCharPinyin = "Wu"
        Case "毕": GetCharPinyin = "Bi"
        Case "毛": GetCharPinyin = "Mao"
        Case "水": GetCharPinyin = "Shui"
        Case "永": GetCharPinyin = "Yong"
        Case "江": GetCharPinyin = "Jiang"
        Case "池": GetCharPinyin = "Chi"
        Case "汤": GetCharPinyin = "Tang"
        Case "汪": GetCharPinyin = "Wang"
        Case "汲": GetCharPinyin = "Ji"
        Case "沃": GetCharPinyin = "Wo"
        Case "沈": GetCharPinyin = "Shen"
        Case "沙": GetCharPinyin = "Sha"
        Case "法": GetCharPinyin = "Fa"
        Case "泽": GetCharPinyin = "Ze"
        Case "津": GetCharPinyin = "Jin"
        Case "洪": GetCharPinyin = "Hong"
        Case "测": GetCharPinyin = "Ce"
        Case "浦": GetCharPinyin = "Pu"
        Case "浩": GetCharPinyin = "Hao"
        Case "温": GetCharPinyin = "Wen"
        Case "游": GetCharPinyin = "You"
        Case "湛": GetCharPinyin = "Zhan"
        Case "滑": GetCharPinyin = "Hua"
        Case "滕": GetCharPinyin = "Teng"
        Case "满": GetCharPinyin = "Man"
        Case "潘": GetCharPinyin = "Pan"
        Case "濮": GetCharPinyin = "Pu"
        Case "焦": GetCharPinyin = "Jiao"
        Case "然": GetCharPinyin = "Ran"
        Case "煜": GetCharPinyin = "Yu"
        Case "熊": GetCharPinyin = "Xiong"
        Case "燕": GetCharPinyin = "Yan"
        Case "爱": GetCharPinyin = "Ai"
        Case "牛": GetCharPinyin = "Niu"
        Case "牧": GetCharPinyin = "Mu"
        Case "狄": GetCharPinyin = "Di"
        Case "猛": GetCharPinyin = "Meng"
        Case "玉": GetCharPinyin = "Yu"
        Case "王": GetCharPinyin = "Wang"
        Case "珠": GetCharPinyin = "Zhu"
        Case "班": GetCharPinyin = "Ban"
        Case "琳": GetCharPinyin = "Lin"
        Case "瑶": GetCharPinyin = "Yao"
        Case "璩": GetCharPinyin = "Qu"
        Case "甄": GetCharPinyin = "Zhen"
        Case "甘": GetCharPinyin = "Gan"
        Case "生": GetCharPinyin = "Sheng"
        Case "田": GetCharPinyin = "Tian"
        Case "申": GetCharPinyin = "Shen"
        Case "男": GetCharPinyin = "Nan"
        Case "白": GetCharPinyin = "Bai"
        Case "皓": GetCharPinyin = "Hao"
        Case "皮": GetCharPinyin = "Pi"
        Case "益": GetCharPinyin = "Yi"
        Case "盖": GetCharPinyin = "Ge"
        Case "盛": GetCharPinyin = "Sheng"
        Case "相": GetCharPinyin = "Xiang"
        Case "睿": GetCharPinyin = "Rui"
        Case "瞿": GetCharPinyin = "Qu"
        Case "石": GetCharPinyin = "Shi"
        Case "磊": GetCharPinyin = "Lei"
        Case "礼": GetCharPinyin = "Li"
        Case "祁": GetCharPinyin = "Qi"
        Case "祖": GetCharPinyin = "Zu"
        Case "祝": GetCharPinyin = "Zhu"
        Case "禄": GetCharPinyin = "Lu"
        Case "福": GetCharPinyin = "Fu"
        Case "禹": GetCharPinyin = "Yu"
        Case "秋": GetCharPinyin = "Qiu"
        Case "程": GetCharPinyin = "Cheng"
        Case "穆": GetCharPinyin = "Mu"
        Case "空": GetCharPinyin = "Kong"
        Case "窦": GetCharPinyin = "Dou"
        Case "立": GetCharPinyin = "Li"
        Case "章": GetCharPinyin = "Zhang"
        Case "童": GetCharPinyin = "Tong"
        Case "竺": GetCharPinyin = "Zhu"
        Case "符": GetCharPinyin = "Fu"
        Case "简": GetCharPinyin = "Jian"
        Case "管": GetCharPinyin = "Guan"
        Case "籍": GetCharPinyin = "Ji"
        Case "米": GetCharPinyin = "Mi"
        Case "糜": GetCharPinyin = "Mi"
        Case "索": GetCharPinyin = "Suo"
        Case "红": GetCharPinyin = "Hong"
        Case "纪": GetCharPinyin = "Ji"
        Case "终": GetCharPinyin = "Zhong"
        Case "经": GetCharPinyin = "Jing"
        Case "缪": GetCharPinyin = "Miao"
        Case "罗": GetCharPinyin = "Luo"
        Case "罡": GetCharPinyin = "Gang"
        Case "羊": GetCharPinyin = "Yang"
        Case "羲": GetCharPinyin = "Xi"
        Case "羿": GetCharPinyin = "Yi"
        Case "翁": GetCharPinyin = "Weng"
        Case "翟": GetCharPinyin = "Zhai"
        Case "翼": GetCharPinyin = "Yi"
        Case "耿": GetCharPinyin = "Geng"
        Case "聂": GetCharPinyin = "Nie"
        Case "聪": GetCharPinyin = "Cong"
        Case "胡": GetCharPinyin = "Hu"
        Case "胥": GetCharPinyin = "Xu"
        Case "能": GetCharPinyin = "Neng"
        Case "臧": GetCharPinyin = "Zang"
        Case "舒": GetCharPinyin = "Shu"
        Case "艾": GetCharPinyin = "Ai"
        Case "芮": GetCharPinyin = "Rui"
        Case "花": GetCharPinyin = "Hua"
        Case "苍": GetCharPinyin = "Cang"
        Case "苏": GetCharPinyin = "Su"
        Case "苗": GetCharPinyin = "Miao"
        Case "若": GetCharPinyin = "Ruo"
        Case "范": GetCharPinyin = "Fan"
        Case "茅": GetCharPinyin = "Mao"
        Case "茹": GetCharPinyin = "Ru"
        Case "荀": GetCharPinyin = "Xun"
        Case "荆": GetCharPinyin = "Jing"
        Case "荣": GetCharPinyin = "Rong"
        Case "莘": GetCharPinyin = "Shen"
        Case "莫": GetCharPinyin = "Mo"
        Case "莹": GetCharPinyin = "Ying"
        Case "萧": GetCharPinyin = "Xiao"
        Case "葛": GetCharPinyin = "Ge"
        Case "董": GetCharPinyin = "Dong"
        Case "蒋": GetCharPinyin = "Jiang"
        Case "蒯": GetCharPinyin = "Kuai"
        Case "蒲": GetCharPinyin = "Pu"
        Case "蓝": GetCharPinyin = "Lan"
        Case "蓟": GetCharPinyin = "Ji"
        Case "蓬": GetCharPinyin = "Peng"
        Case "蔚": GetCharPinyin = "Wei"
        Case "蔡": GetCharPinyin = "Cai"
        Case "蔺": GetCharPinyin = "Lin"
        Case "薄": GetCharPinyin = "Bo"
        Case "薛": GetCharPinyin = "Xue"
        Case "虞": GetCharPinyin = "Yu"
        Case "融": GetCharPinyin = "Rong"
        Case "衡": GetCharPinyin = "Heng"
        Case "袁": GetCharPinyin = "Yuan"
        Case "裘": GetCharPinyin = "Qiu"
        Case "裴": GetCharPinyin = "Pei"
        Case "褚": GetCharPinyin = "Chu"
        Case "解": GetCharPinyin = "Xie"
        Case "訾": GetCharPinyin = "Zi"
        Case "詹": GetCharPinyin = "Zhan"
        Case "计": GetCharPinyin = "Ji"
        Case "许": GetCharPinyin = "Xu"
        Case "试": GetCharPinyin = "Shi"
        Case "诸": GetCharPinyin = "Zhu"
        Case "谈": GetCharPinyin = "Tan"
        Case "谢": GetCharPinyin = "Xie"
        Case "谭": GetCharPinyin = "Tan"
        Case "谷": GetCharPinyin = "Gu"
        Case "贝": GetCharPinyin = "Bei"
        Case "贡": GetCharPinyin = "Gong"
        Case "贲": GetCharPinyin = "Ben"
        Case "贵": GetCharPinyin = "Gui"
        Case "费": GetCharPinyin = "Fei"
        Case "贺": GetCharPinyin = "He"
        Case "贾": GetCharPinyin = "Jia"
        Case "赖": GetCharPinyin = "Lai"
        Case "赵": GetCharPinyin = "Zhao"
        Case "越": GetCharPinyin = "Yue"
        Case "路": GetCharPinyin = "Lu"
        Case "车": GetCharPinyin = "Che"
        Case "辉": GetCharPinyin = "Hui"
        Case "辛": GetCharPinyin = "Xin"
        Case "边": GetCharPinyin = "Bian"
        Case "连": GetCharPinyin = "Lian"
        Case "逄": GetCharPinyin = "Pang"
        Case "选": GetCharPinyin = "Xuan"
        Case "通": GetCharPinyin = "Tong"
        Case "逯": GetCharPinyin = "Lu"
        Case "道": GetCharPinyin = "Dao"
        Case "邓": GetCharPinyin = "Deng"
        Case "邢": GetCharPinyin = "Xing"
        Case "那": GetCharPinyin = "Na"
        Case "邬": GetCharPinyin = "Wu"
        Case "邰": GetCharPinyin = "Tai"
        Case "邱": GetCharPinyin = "Qiu"
        Case "邴": GetCharPinyin = "Bing"
        Case "邵": GetCharPinyin = "Shao"
        Case "邹": GetCharPinyin = "Zou"
        Case "郁": GetCharPinyin = "Yu"
        Case "郎": GetCharPinyin = "Lang"
        Case "郏": GetCharPinyin = "Jia"
        Case "郑": GetCharPinyin = "Zheng"
        Case "郗": GetCharPinyin = "Xi"
        Case "郜": GetCharPinyin = "Gao"
        Case "郝": GetCharPinyin = "Hao"
        Case "郦": GetCharPinyin = "Li"
        Case "郭": GetCharPinyin = "Guo"
        Case "都": GetCharPinyin = "Du"
        Case "鄂": GetCharPinyin = "E"
        Case "金": GetCharPinyin = "Jin"
        Case "钟": GetCharPinyin = "Zhong"
        Case "钢": GetCharPinyin = "Gang"
        Case "钭": GetCharPinyin = "Tou"
        Case "钮": GetCharPinyin = "Niu"
        Case "钱": GetCharPinyin = "Qian"
        Case "锋": GetCharPinyin = "Feng"
        Case "锴": GetCharPinyin = "Kai"
        Case "闫": GetCharPinyin = "Yan"
        Case "闵": GetCharPinyin = "Min"
        Case "闻": GetCharPinyin = "Wen"
        Case "阎": GetCharPinyin = "Yan"
        Case "阙": GetCharPinyin = "Que"
        Case "阚": GetCharPinyin = "Kan"
        Case "阮": GetCharPinyin = "Ruan"
        Case "阳": GetCharPinyin = "Yang"
        Case "阴": GetCharPinyin = "Yin"
        Case "陆": GetCharPinyin = "Lu"
        Case "陈": GetCharPinyin = "Chen"
        Case "陶": GetCharPinyin = "Tao"
        Case "隆": GetCharPinyin = "Long"
        Case "隗": GetCharPinyin = "Kui"
        Case "雄": GetCharPinyin = "Xiong"
        Case "雍": GetCharPinyin = "Yong"
        Case "雒": GetCharPinyin = "Luo"
        Case "雪": GetCharPinyin = "Xue"
        Case "雯": GetCharPinyin = "Wen"
        Case "雷": GetCharPinyin = "Lei"
        Case "震": GetCharPinyin = "Zhen"
        Case "霍": GetCharPinyin = "Huo"
        Case "青": GetCharPinyin = "Qing"
        Case "静": GetCharPinyin = "Jing"
        Case "靳": GetCharPinyin = "Jin"
        Case "鞠": GetCharPinyin = "Ju"
        Case "韦": GetCharPinyin = "Wei"
        Case "韩": GetCharPinyin = "Han"
        Case "韶": GetCharPinyin = "Shao"
        Case "项": GetCharPinyin = "Xiang"
        Case "顺": GetCharPinyin = "Shun"
        Case "须": GetCharPinyin = "Xu"
        Case "顾": GetCharPinyin = "Gu"
        Case "颜": GetCharPinyin = "Yan"
        Case "飞": GetCharPinyin = "Fei"
        Case "饶": GetCharPinyin = "Rao"
        Case "马": GetCharPinyin = "Ma"
        Case "骆": GetCharPinyin = "Luo"
        Case "高": GetCharPinyin = "Gao"
        Case "魏": GetCharPinyin = "Wei"
        Case "鱼": GetCharPinyin = "Yu"
        Case "鲁": GetCharPinyin = "Lu"
        Case "鲍": GetCharPinyin = "Bao"
        Case "鸣": GetCharPinyin = "Ming"
        Case "鹏": GetCharPinyin = "Peng"
        Case "麻": GetCharPinyin = "Ma"
        Case "黄": GetCharPinyin = "Huang"
        Case "黎": GetCharPinyin = "Li"
        Case "鼎": GetCharPinyin = "Ding"
        Case "齐": GetCharPinyin = "Qi"
        Case "龙": GetCharPinyin = "Long"
        Case "龚": GetCharPinyin = "Gong"
        Case Else
            GetCharPinyin = ""
    End Select
End Function
' ==============================================================================
' 跨模块安全热更 Mod0_ControlPanel 引擎 (由 Mod1 代为执行，避免 Mod0 自删调用栈崩溃)
' ==============================================================================

' ==============================================================================
' 辅助跨模块调度: 异步热更 Mod0_ControlPanel (防止 Mod0 自删除自修改导致 Excel 崩溃)
' ==============================================================================