' ==============================================================================
' 模块名称: Mod_Sync
' 核心职责: 独立热更与重置引擎 (Hot-Reload & Reset Bootloader)
' 设计哲学: 自身作为永久固化的底座，专注将磁盘上的 Mod0、Mod1、Mod2 秒级载入内存
' ==============================================================================
Option Explicit

Public Sub 一键热更并重置面板()
    Dim wb As Workbook, vbp As Object
    Dim modulesDir As String, fso As Object
    Dim modNames As Variant, mName As Variant
    Dim basFile As String
    
    Set wb = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
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
    
    modNames = Array("Mod0_ControlPanel", "Mod1_TeacherPinyin", "Mod2_CleanRawData")
    
    For Each mName In modNames
        basFile = modulesDir & Application.PathSeparator & CStr(mName) & ".bas"
        If fso.FileExists(basFile) Then
            Call ReloadSingleModule(wb, CStr(mName), basFile)
        End If
    Next mName
    
    ' 热更新完成后，调用全新的 Mod0 重构面板并刷新数据
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.一键生成控制面板"
    Application.Run "Mod0_ControlPanel.AppendLog", "【热更就绪】已全量同步本地最新代码并完成控制面板标准重构！"
    On Error GoTo 0
End Sub

Public Sub 一键热更所有模块()
    Call 一键热更并重置面板
End Sub

Public Sub 一键自动化执行全流程()
    ' 自动路由至 Mod0 的业务全流程
    On Error Resume Next
    Application.Run "Mod0_ControlPanel.一键执行业务全流程"
    On Error GoTo 0
End Sub

Private Sub ReloadSingleModule(wb As Workbook, modName As String, basPath As String)
    Dim vbp As Object, comp As Object, fso As Object, ts As Object, codeContent As String
    Set vbp = wb.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set ts = fso.OpenTextFile(basPath, 1, False, -2)
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