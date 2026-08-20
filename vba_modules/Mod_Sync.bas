' ==============================================================================
' 模块名称: Mod_Sync
' 核心职责: 独立全自动热更与重置引擎 (Dynamic Hot-Reload & Reset Bootloader)
' 设计哲学: 自身作为永久固化的底座，动态遍历磁盘上 vba_modules/*.bas 秒级载入内存
' ==============================================================================
Option Explicit

Public Sub 一键热更并重置面板()
    Dim wb As Workbook, vbp As Object
    Dim modulesDir As String, fso As Object, f As Object
    Dim fName As String, baseName As String, loadedDict As Object
    Dim comp As Object, i As Long
    
    Set wb = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set loadedDict = CreateObject("Scripting.Dictionary")
    modulesDir = wb.Path & Application.PathSeparator & "vba_modules"
    
    If Not fso.FolderExists(modulesDir) Then
        MsgBox "未找到本地代码目录：" & vbCrLf & modulesDir, vbCritical + vbOKOnly, "目录缺失"
        Exit Sub
    End If
    
    On Error Resume Next
    Set vbp = wb.VBProject
    If Err.Number <> 0 Then
        MsgBox "无法访问 VBA 工程！请确保已在 Excel 选项中勾选：【信任对 VBA 工程对象模型的访问】。", vbCritical + vbOKOnly, "权限受限"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 1. 动态遍历 vba_modules/*.bas 并重载
    For Each f In fso.GetFolder(modulesDir).Files
        fName = f.Name
        If LCase(fso.GetExtensionName(fName)) = "bas" Then
            baseName = fso.GetBaseName(fName)
            If StrComp(baseName, "Mod_Sync", vbTextCompare) <> 0 Then
                Call ReloadSingleModule(wb, baseName, f.Path)
                loadedDict(LCase(baseName)) = 1
            End If
        End If
    Next f
    
    ' 2. 清理已在磁盘删除的废弃标准模块
    For i = vbp.VBComponents.Count To 1 Step -1
        Set comp = vbp.VBComponents(i)
        If comp.Type = 1 Then ' 1 = vbext_ct_StdModule
            If StrComp(comp.Name, "Mod_Sync", vbTextCompare) <> 0 Then
                If Not loadedDict.Exists(LCase(comp.Name)) Then
                    vbp.VBComponents.Remove comp
                End If
            End If
        End If
    Next i
    
    ' 3. 热更新完成后，保存工作簿并重构面板与刷新数据
    wb.Save
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.一键生成控制面板"
    Application.Run "Mod0_ControlPanel.AppendLog", "【热更就绪】已全量动态同步本地最新代码并完成控制面板标准重构！"
    On Error GoTo 0
End Sub

Public Sub 一键热更所有模块()
    Call 一键热更并重置面板
End Sub


Private Sub ReloadSingleModule(wb As Workbook, modName As String, basPath As String)
    Dim vbp As Object, comp As Object, fso As Object, ts As Object, codeContent As String
    Set vbp = wb.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set ts = fso.OpenTextFile(basPath, 1, False, -2) ' -2 = TristateUseDefault (GBK / ANSI)
    codeContent = ts.ReadAll
    ts.Close
    
    On Error Resume Next
    Set comp = vbp.VBComponents(modName)
    On Error GoTo 0
    
    If comp Is Nothing Then
        Set comp = vbp.VBComponents.Add(1)
        comp.Name = modName
    End If
    
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub
