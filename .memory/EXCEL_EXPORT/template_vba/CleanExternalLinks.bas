Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

'=============================================================================
' CleanExternalLinks — 清除工作簿中所有外部链接
'
' 对应 Python ZIP/XML 清理逻辑：
'   Step 1: 扫描外部链接源 + 损坏 defined names
'   Step 2: 含外部引用 [N] 的公式 → 数组批量转值（避免逐单元格 COM 调用）
'   Step 3: 删除指向外部工作簿或包含错误的 defined names
'   Step 4: BreakLink 断开所有外部链接源
'   Step 5: 清理隐藏名称（跳过内置名称和 _xlfn 前缀）
'   Step 6: 删除可能残留的外部链接工作表（占位，Excel 保存时自动处理）
'   Step 7: 输出清理报告
'
' 性能设计：
'   - 公式扫描/转换使用 Variant 数组批量处理，3 次 COM 调用覆盖整张表
'   - 每 200 行 yield 一次 DoEvents，防止 Windows 标记"无响应"
'   - StatusBar 实时进度
'=============================================================================

Private Type CleanStats
    ExtLinksFound    As Long
    NamesDeleted     As Long
    CellsConverted   As Long
    LinksBroken      As Long
    ErrorsEncountered As Long
End Type

Private Const ERR_TOKENS        As String = "#REF! #VALUE! #N/A #NAME? #DIV/0! #NULL! #NUM!"
Private Const DOEVENTS_INTERVAL As Long = 200   ' 每处理多少行 yield 一次

'=============================================================================
' 主入口
'=============================================================================
Public Sub CleanExternalLinks_Main(Optional ByVal wb As Workbook)
    Dim stats     As CleanStats
    Dim startTime As Long

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "没有可用的工作簿。", vbExclamation
        Exit Sub
    End If

    startTime = GetTickCount

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Step 1: 扫描
    Application.StatusBar = "CleanExternalLinks: 正在扫描外部链接..."
    ScanExternalLinks wb, stats

    If stats.ExtLinksFound = 0 Then
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "未检测到外部链接，无需清理。", vbInformation
        Exit Sub
    End If

    ' Step 2: 公式 → 值（数组批量）
    Application.StatusBar = "CleanExternalLinks: 正在转换含外部引用的公式为值..."
    ConvertExternalFormulas wb, stats

    ' Step 3: 删除损坏的 defined names
    Application.StatusBar = "CleanExternalLinks: 正在删除损坏的名称..."
    DeleteBadNames wb, stats

    ' Step 4: BreakLink
    Application.StatusBar = "CleanExternalLinks: 正在断开链接源..."
    BreakAllLinks wb, stats

    ' Step 5: 隐藏名称
    Application.StatusBar = "CleanExternalLinks: 正在清理隐藏名称..."
    CleanHiddenNames wb, stats

    ' Step 6
    DeleteExternalLinkSheets wb

    ' Step 7: 报告
    Application.StatusBar = False
    PrintReport stats, GetTickCount - startTime

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'=============================================================================
' Step 1: 扫描
'=============================================================================
Private Sub ScanExternalLinks(ByVal wb As Workbook, ByRef stats As CleanStats)
    Dim links As Variant
    Dim i     As Long

    links = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            stats.ExtLinksFound = stats.ExtLinksFound + 1
            Debug.Print "[Step1] 外部链接源 " & i & ": " & links(i)
        Next i
    End If

    Dim nm As Name
    For Each nm In wb.Names
        If IsBadNameRefersTo(nm.RefersTo) Then
            stats.ExtLinksFound = stats.ExtLinksFound + 1
            Debug.Print "[Step1] 损坏名称: " & nm.Name & " -> " & nm.RefersTo
        End If
    Next nm
End Sub

'=============================================================================
' Step 2: 数组批量转换公式为值（核心性能优化点）
'
' 关键：不再逐单元格 For Each，而是：
'   1. SpecialCells(xlCellTypeFormulas) 找到公式区域
'   2. 对每个区域，一次 COM 调用读取整个 .Formula 数组 + .Value 数组
'   3. 在 VBA 内存中遍历数组、标记外部引用、替换为值
'   4. 一次 COM 调用将修改后的数组写回
'
' 这样 1 万行的工作表从 ~3 万次 COM 调用降到 ~6 次。
'=============================================================================
Private Sub ConvertExternalFormulas(ByVal wb As Workbook, ByRef stats As CleanStats)
    Dim ws        As Worksheet
    Dim area      As Range
    Dim formulas  As Variant
    Dim values    As Variant
    Dim r         As Long
    Dim c         As Long
    Dim rowCount  As Long
    Dim colCount  As Long
    Dim converted As Long

    For Each ws In wb.Worksheets
        If ws.Type <> xlWorksheet Then GoTo NextSheet
        If ws.UsedRange Is Nothing Then GoTo NextSheet

        On Error Resume Next
        Set area = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0

        If area Is Nothing Then GoTo NextSheet

        Application.StatusBar = "CleanExternalLinks: 处理 " & ws.Name & " ..."

        ' SpecialCells 可能返回不连续区域，逐个 area 处理
        Dim subArea As Range
        For Each subArea In area.Areas
            rowCount = subArea.Rows.Count
            colCount = subArea.Columns.Count

            ' 单单元格直接处理
            If rowCount = 1 And colCount = 1 Then
                If HasExternalRef(CStr(subArea.Formula)) Then
                    On Error Resume Next
                    subArea.Value = subArea.Value
                    If Err.Number = 0 Then converted = converted + 1 Else stats.ErrorsEncountered = stats.ErrorsEncountered + 1
                    Err.Clear
                    On Error GoTo 0
                End If
                GoTo NextArea
            End If

            ' 大区域：数组批量处理
            ' 一次 COM 读公式
            formulas = subArea.Formula
            If IsEmpty(formulas) Then GoTo NextArea

            ' 先扫描是否存在外部引用（避免不必要的 Value 读取）
            Dim hasExt As Boolean
            hasExt = False
            For r = 1 To rowCount
                For c = 1 To colCount
                    If HasExternalRef(CStr(formulas(r, c))) Then
                        hasExt = True
                        Exit For
                    End If
                Next c
                If hasExt Then Exit For
            Next r

            If Not hasExt Then GoTo NextArea

            ' 一次 COM 读当前值
            values = subArea.Value

            ' 内存中替换
            For r = 1 To rowCount
                For c = 1 To colCount
                    If HasExternalRef(CStr(formulas(r, c))) Then
                        formulas(r, c) = values(r, c)
                        converted = converted + 1
                    End If
                Next c

                ' 定期 yield，防止假死
                If r Mod DOEVENTS_INTERVAL = 0 Then DoEvents
            Next r

            ' 一次 COM 写回
            subArea.Formula = formulas
NextArea:
        Next subArea

        If converted > 0 Then
            Debug.Print "[Step2] " & ws.Name & ": 转换 " & converted & " 个公式为值"
            stats.CellsConverted = stats.CellsConverted + converted
            converted = 0
        End If

NextSheet:
    Next ws
End Sub

'=============================================================================
' Step 3: 删除损坏的 defined names
'=============================================================================
Private Sub DeleteBadNames(ByVal wb As Workbook, ByRef stats As CleanStats)
    Dim nm       As Name
    Dim badNames As Collection
    Dim v        As Variant

    Set badNames = New Collection

    For Each nm In wb.Names
        If IsBadNameRefersTo(nm.RefersTo) Then
            badNames.Add nm.Name
        End If
    Next nm

    For Each v In badNames
        On Error Resume Next
        wb.Names(CStr(v)).Delete
        If Err.Number = 0 Then
            stats.NamesDeleted = stats.NamesDeleted + 1
            Debug.Print "[Step3] 删除名称: " & CStr(v)
        Else
            stats.ErrorsEncountered = stats.ErrorsEncountered + 1
            Debug.Print "[Step3] 删除失败: " & CStr(v) & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next v
End Sub

'=============================================================================
' Step 4: BreakLink
'=============================================================================
Private Sub BreakAllLinks(ByVal wb As Workbook, ByRef stats As CleanStats)
    Dim links As Variant
    Dim i     As Long

    links = wb.LinkSources(xlExcelLinks)
    If IsEmpty(links) Then Exit Sub

    For i = LBound(links) To UBound(links)
        Application.StatusBar = "CleanExternalLinks: 正在断开链接 " & i & "/" & (UBound(links) - LBound(links) + 1)
        On Error Resume Next
        wb.BreakLink links(i), xlExcelLinks
        If Err.Number = 0 Then
            stats.LinksBroken = stats.LinksBroken + 1
            Debug.Print "[Step4] 断开链接: " & links(i)
        Else
            stats.ErrorsEncountered = stats.ErrorsEncountered + 1
            Debug.Print "[Step4] 断开失败: " & links(i) & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        DoEvents
    Next i
End Sub

'=============================================================================
' Step 5: 清理隐藏名称
'=============================================================================
Private Sub CleanHiddenNames(ByVal wb As Workbook, ByRef stats As CleanStats)
    Dim nm       As Name
    Dim badNames As Collection
    Dim v        As Variant

    Set badNames = New Collection

    For Each nm In wb.Names
        If nm.Name Like "*!*" Then GoTo NextName2
        If nm.Name Like "_xlfn*" Then GoTo NextName2
        If nm.Name Like "_xlpm*" Then GoTo NextName2
        If nm.Name Like "_FilterDatabase*" Then GoTo NextName2

        If IsBadNameRefersTo(nm.RefersTo) Then
            badNames.Add nm.Name
        End If
NextName2:
    Next nm

    For Each v In badNames
        On Error Resume Next
        wb.Names(CStr(v)).Delete
        If Err.Number = 0 Then
            stats.NamesDeleted = stats.NamesDeleted + 1
            Debug.Print "[Step5] 删除隐藏名称: " & CStr(v)
        Else
            Debug.Print "[Step5] 跳过: " & CStr(v) & " (" & Err.Description & ")"
            Err.Clear
        End If
        On Error GoTo 0
    Next v
End Sub

'=============================================================================
' Step 6: 外部链接工作表（占位）
'=============================================================================
Private Sub DeleteExternalLinkSheets(ByVal wb As Workbook)
    Debug.Print "[Step6] 外部链接工作表将在保存时自动清理"
End Sub

'=============================================================================
' 辅助：公式是否包含外部引用
'=============================================================================
Private Function HasExternalRef(ByVal formulaText As String) As Boolean
    If Len(formulaText) = 0 Then Exit Function

    ' [N] 外部工作簿索引
    If InStr(formulaText, "[") > 0 Then
        HasExternalRef = True
        Exit Function
    End If

    ' 错误 token
    Dim tokens As Variant
    Dim i      As Long
    tokens = Split(ERR_TOKENS, " ")
    For i = LBound(tokens) To UBound(tokens)
        If InStr(formulaText, tokens(i)) > 0 Then
            HasExternalRef = True
            Exit Function
        End If
    Next i

    ' file:// 协议
    If InStr(formulaText, "file://") > 0 Then
        HasExternalRef = True
        Exit Function
    End If

    ' UNC \\server\...
    If formulaText Like "\\\\[a-zA-Z]*" Then
        HasExternalRef = True
        Exit Function
    End If

    HasExternalRef = False
End Function

'=============================================================================
' 辅助：defined name 是否损坏/含外部引用
'=============================================================================
Private Function IsBadNameRefersTo(ByVal refersTo As String) As Boolean
    If Len(refersTo) = 0 Then
        IsBadNameRefersTo = True
        Exit Function
    End If

    Dim tokens As Variant
    Dim i      As Long
    tokens = Split(ERR_TOKENS, " ")
    For i = LBound(tokens) To UBound(tokens)
        If InStr(refersTo, tokens(i)) > 0 Then
            IsBadNameRefersTo = True
            Exit Function
        End If
    Next i

    If InStr(refersTo, "[") > 0 Then
        IsBadNameRefersTo = True
        Exit Function
    End If

    If InStr(refersTo, "file://") > 0 Then
        IsBadNameRefersTo = True
        Exit Function
    End If

    If refersTo Like "\\\\[a-zA-Z]*" Then
        IsBadNameRefersTo = True
        Exit Function
    End If

    If refersTo Like "[A-Z]:\*" Then
        IsBadNameRefersTo = True
        Exit Function
    End If

    IsBadNameRefersTo = False
End Function

'=============================================================================
' Step 7: 报告
'=============================================================================
Private Sub PrintReport(ByRef stats As CleanStats, ByVal elapsedMs As Long)
    Dim msg As String

    msg = "=== 外部链接清理报告 ===" & vbCrLf & vbCrLf & _
          "检测到外部链接源/损坏名称: " & stats.ExtLinksFound & vbCrLf & _
          "公式转换为值: " & stats.CellsConverted & vbCrLf & _
          "删除 defined names: " & stats.NamesDeleted & vbCrLf & _
          "断开链接: " & stats.LinksBroken & vbCrLf & _
          "错误数: " & stats.ErrorsEncountered & vbCrLf & _
          "耗时: " & Format(elapsedMs / 1000, "0.00") & " 秒" & vbCrLf & vbCrLf & _
          "保存工作簿后，calcChain 将自动重建。" & vbCrLf & _
          "请验证：重新打开文件应不再弹出'更新链接'提示。"

    Debug.Print Replace(msg, vbCrLf, vbCrLf & "[Report] ")
    MsgBox msg, vbInformation, "清理完成"
End Sub

'=============================================================================
' 快捷入口
'=============================================================================
Public Sub CleanExternalLinks_ActiveWorkbook()
    CleanExternalLinks_Main ActiveWorkbook
End Sub
