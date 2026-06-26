VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rateLibrary
   Caption         =   "综合单价库"
   ClientHeight    =   8268.001
   ClientLeft      =   288
   ClientTop       =   1188
   ClientWidth     =   12864
   OleObjectBlob   =   "rateLibrary.frx":0000
   StartUpPosition =   1  '窗口所有者中心
End
Attribute VB_Name = "rateLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 综合单价库 - 简洁版，无 Win32 API 调用

Private Sub UserForm_Initialize()
    On Error Resume Next

    ' 设置表头标签
    With Label1
        .Caption = "  编号" & String(8, " ") & "名称" & String(72, " ") & "单位" & String(8, " ") & "单价"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleNone
    End With

    ' 设置ListBox为多列模式
    With ListBox1
        .Font.Name = "Consolas"
        .Font.Size = 9
        .IntegralHeight = False
        .ColumnCount = 4
        .ColumnWidths = "60;400;50;80"
    End With

    ' 初始化搜索控件
    TextBox1.Text = ""
    CommandButton1.Caption = "上一个"
    CommandButton2.Caption = "下一个"

    With Label2
        .Caption = "0/0"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .TextAlign = fmTextAlignLeft
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleNone
    End With

    TextBox1.Tag = "输入搜索内容后按回车键搜索"

    ' 初始化搜索管理器
    SearchManager.InitializeSearchManager ListBox1, Label2

    ' 加载数据
    Call LoadRateData

    ' 恢复窗口位置
    Call RestoreFormPosition
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call SaveFormPosition
    SearchManager.CleanupSearchManager
End Sub

' ====================================================================
' 控件事件处理
' ====================================================================

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox1.ListIndex < 0 Then Exit Sub

    Dim 编号 As String
    编号 = ListBox1.List(ListBox1.ListIndex, 0)

    If Not Selection Is Nothing Then
        Selection.Cells(1, 1).Value = 编号
    End If

    Unload Me
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyR And Shift = 2 Then  ' Ctrl+R: 清除搜索
        TextBox1.Text = ""
        SearchManager.ClearSearchResults
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call PerformSearch
    ElseIf KeyCode = vbKeyEscape Then
        TextBox1.Text = ""
        SearchManager.ClearSearchResults
        ListBox1.SetFocus
    End If
End Sub

Private Sub CommandButton1_Click()
    If Not SearchManager.GotoPreviousResult() Then
        MsgBox "无法执行操作", vbInformation, "提示"
        TextBox1.SetFocus
    End If
End Sub

Private Sub CommandButton2_Click()
    If Not SearchManager.GotoNextResult() Then
        MsgBox "无法执行操作", vbInformation, "提示"
        TextBox1.SetFocus
    End If
End Sub

' ====================================================================
' 核心功能方法
' ====================================================================

Private Sub PerformSearch()
    If Trim(TextBox1.Text) = "" Then
        MsgBox "请输入要搜索的内容", vbInformation, "搜索提示"
        Exit Sub
    End If

    Dim resultCount As Long
    resultCount = SearchManager.PerformNewSearch(TextBox1.Text)

    If resultCount = 0 Then
        MsgBox "未找到包含 '" & Trim(TextBox1.Text) & "' 的项目", vbInformation, "搜索结果"
        TextBox1.SelStart = 0
        TextBox1.SelLength = Len(TextBox1.Text)
    End If
End Sub

Public Sub EnsureItemVisibleInForm(itemIndex As Long)
    If itemIndex < 0 Or itemIndex >= ListBox1.ListCount Then Exit Sub

    Dim actualRowHeight As Single
    Dim visibleRows As Long

    actualRowHeight = 15
    visibleRows = Int(ListBox1.Height / actualRowHeight)
    If visibleRows < 1 Then visibleRows = 1

    Dim idealTopIndex As Long
    idealTopIndex = itemIndex - Int(visibleRows / 2)

    If idealTopIndex < 0 Then idealTopIndex = 0
    If idealTopIndex > ListBox1.ListCount - visibleRows Then
        idealTopIndex = ListBox1.ListCount - visibleRows
        If idealTopIndex < 0 Then idealTopIndex = 0
    End If

    ListBox1.ListIndex = itemIndex
    ListBox1.TopIndex = idealTopIndex

    Me.Repaint
    DoEvents
End Sub

' ====================================================================
' 数据加载
' ====================================================================

Private Sub LoadRateData()
    On Error Resume Next
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("RATE")
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        ListBox1.Clear
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim data1 As Variant, data2 As Variant
    data1 = ws.Range("A2:C" & lastRow).Value
    data2 = ws.Range("J2:J" & lastRow).Value

    Dim rows1 As Long, rows2 As Long
    If TypeName(data1) = "Variant()" Then
        rows1 = UBound(data1, 1)
    Else
        rows1 = 1
    End If
    If TypeName(data2) = "Variant()" Then
        rows2 = UBound(data2, 1)
    Else
        rows2 = 1
    End If

    Dim actualRows As Long
    actualRows = WorksheetFunction.Min(rows1, rows2)

    Dim resultData() As Variant
    ReDim resultData(1 To actualRows, 1 To 4)

    Dim i As Long
    For i = 1 To actualRows
        On Error Resume Next
        If TypeName(data1) = "Variant()" Then
            resultData(i, 1) = data1(i, 1)
            resultData(i, 2) = data1(i, 2)
            resultData(i, 3) = data1(i, 3)
        Else
            resultData(i, 1) = data1
            resultData(i, 2) = ws.Cells(2, 2).Value
            resultData(i, 3) = ws.Cells(2, 3).Value
        End If

        If TypeName(data2) = "Variant()" Then
            resultData(i, 4) = data2(i, 1)
        Else
            resultData(i, 4) = data2
        End If
        On Error GoTo ErrorHandler
    Next i

    ListBox1.Clear
    ListBox1.List = resultData

    Label2.Caption = "0/" & ListBox1.ListCount
    SearchManager.InitializeSearchManager ListBox1, Label2
    Exit Sub

ErrorHandler:
    ListBox1.Clear
    Debug.Print "数据加载错误: " & Err.Description
    On Error GoTo 0
End Sub

' ====================================================================
' 窗口位置管理 (使用注册表，无 API 调用)
' ====================================================================

Private Sub RestoreFormPosition()
    On Error Resume Next

    Dim savedLeft As Long, savedTop As Long
    savedLeft = GetSetting("RateLibrary", "Position", "Left", -1)
    savedTop = GetSetting("RateLibrary", "Position", "Top", -1)

    If savedLeft = -1 Or savedTop = -1 Then
        ' 首次打开，居中于 Excel 窗口
        Me.StartUpPosition = 1
    Else
        Me.Left = savedLeft
        Me.Top = savedTop
    End If
End Sub

Private Sub SaveFormPosition()
    On Error Resume Next
    SaveSetting "RateLibrary", "Position", "Left", Me.Left
    SaveSetting "RateLibrary", "Position", "Top", Me.Top
End Sub
