VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rateLibrary 
   Caption         =   "综合单价库"
   ClientHeight    =   8268.001
   ClientLeft      =   288
   ClientTop       =   1188
   ClientWidth     =   12864
   OleObjectBlob   =   "rateLibrary.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "rateLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' 移除了与VBE环境不兼容的Attribute属性
' VBA中，用代码设置窗体可调整大小的功能代码
#If Win64 Then ' 64位系统
    Private Declare PtrSafe Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal Hwnd As LongPtr, _
        ByVal nIndex As Long) _
        As LongPtr

    Private Declare PtrSafe Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal Hwnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr) _
        As LongPtr

    Private Declare PtrSafe Function FindWindow _
        Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) _
        As LongPtr

    Private Declare PtrSafe Function DrawMenuBar _
        Lib "user32" ( _
        ByVal Hwnd As LongPtr) _
        As Long
        
    Private Declare PtrSafe Function SetWindowPos _
        Lib "user32" ( _
        ByVal Hwnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) _
        As Long
    
    Private hWndForm As LongPtr
    Private FIstype As LongPtr
#Else ' 32位系统
    Private Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) _
        As Long

    Private Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) _
        As Long

    Private Declare Function FindWindow _
        Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) _
        As Long

    Private Declare Function DrawMenuBar _
        Lib "user32" ( _
        ByVal hwnd As Long) _
        As Long
        
    Private Declare Function SetWindowPos _
        Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) _
        As Long
    
    Private hWndForm As Long
    Private FIstype As Long
#End If

Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

' 鼠标滚轮支持标志
Private m_WheelEnabled As Boolean
Private m_ActivateCount As Long

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

' ====================================================================
' 窗体核心事件处理
' ====================================================================

Private Sub UserForm_Initialize()
    On Error Resume Next
    Debug.Print "UserForm_Initialize 开始"
    
    ' 设置窗口可调整大小
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    FIstype = GetWindowLong(hWndForm, GWL_STYLE)
    FIstype = FIstype Or WS_THICKFRAME
    SetWindowLong hWndForm, GWL_STYLE, FIstype
    DrawMenuBar hWndForm
    
    ' 恢复窗体位置（首次显示时居中，之后恢复到上次位置）
    Call RestoreFormPosition
    
    ' 设置表头标签（与ListBox列宽对应：60;400;50;80磅）
    With Label1
        .Caption = "  编号" & String(8, " ") & "名称" & String(72, " ") & "单位" & String(8, " ") & "单价"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' 初始化鼠标滚轮支持 - 增强版
    ' 添加详细调试信息和健壮的初始化逻辑
    On Error GoTo WheelInitError
    Debug.Print "=== 开始初始化鼠标滚轮支持 ==="
    Debug.Print "窗体标题: " & Me.Caption
    Debug.Print "ListBox1名称: " & ListBox1.Name
    Debug.Print "ListBox1句柄是否存在: " & (hWndForm <> 0)
    
    ' 确保ListBox1先获得焦点（这对滚轮初始化很重要）
    ListBox1.SetFocus
    Debug.Print "ListBox1焦点设置状态: " & ListBox1.Enabled
    
    ' 重置状态标志
    m_WheelEnabled = False
    
    ' 先尝试禁用任何可能存在的滚轮支持
    On Error Resume Next
    Call MouseWheelHandler.DisableMouseWheelSupport("初始化前清理")
    Debug.Print "清理旧的滚轮支持完成"
    On Error GoTo WheelInitError
    
    ' 调用MouseWheelHandler启用滚轮支持
    Debug.Print "尝试调用MouseWheelHandler.EnableMouseWheelSupport..."
    Dim initResult As Boolean
    initResult = MouseWheelHandler.EnableMouseWheelSupport(Me.Caption, Me)
    
    If initResult Then
        m_WheelEnabled = True
        Debug.Print "=== 鼠标滚轮支持初始化成功 ==="
        
        ' 设置活动ListBox引用（额外保险）
        MouseWheelHandler.SetActiveListBox ListBox1
        Debug.Print "ListBox引用已设置"
        
        ' 立即运行测试以验证初始化
        Debug.Print "运行初始化测试..."
        On Error Resume Next
        MouseWheelHandler.TestMouseWheelCapture
        On Error GoTo 0
    Else
        Debug.Print "=== 鼠标滚轮支持初始化失败，但不中断程序 ==="
    End If
    
    On Error GoTo 0
    GoTo WheelInitContinue
    
WheelInitError:
    Debug.Print "滚轮初始化错误: " & Err.Description & " (错误号: " & Err.Number & ")"
    On Error Resume Next
WheelInitContinue:
    On Error GoTo 0
    
    ' 设置ListBox为多列模式，固定列宽
    With ListBox1
        .Font.Name = "Consolas"
        .Font.Size = 9
        .IntegralHeight = False
        .ColumnCount = 4                    ' 4列：编号、名称、单位、单价
        .ColumnWidths = "60;400;50;80"      ' 固定每列宽度（单位：磅）
    End With
    
    ' 设置搜索相关控件
    TextBox1.Text = ""  ' 清空搜索框
    CommandButton1.Caption = "上一条"
    CommandButton2.Caption = "下一条"
    
    ' 设置Label2显示搜索结果统计
    With Label2
        .Caption = "0/0"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .TextAlign = fmTextAlignLeft
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' 为TextBox1添加提示（通过Tag属性存储提示信息）
    TextBox1.Tag = "输入搜索内容后按回车键搜索"
    
    ' 初始化搜索管理器
    SearchManager.InitializeSearchManager ListBox1, Label2
    
 ' [新增] 强制加载最新数据
    Call LoadRateData
    
    ' 移除Resize事件调用，保持原始设计尺寸
' UserForm_Resize - 已注释以保持原始VBE设计尺寸
End Sub

' 窗体激活时重新启用滚轮支持
Private Sub UserForm_Activate()
    On Error Resume Next
    
    m_ActivateCount = m_ActivateCount + 1
    Debug.Print "UserForm_Activate 调用第" & m_ActivateCount & "次"
    
    ' 重新启用滚轮支持
    If Not m_WheelEnabled Then
        Call EnableWheelSupport
    Else
        ' 重新应用滚轮支持，确保窗口子类化状态正确
        Call MouseWheelHandler.EnableMouseWheelSupport(Me.Caption, Me)
        Debug.Print "滚轮支持已重新应用"
    End If
    
    On Error GoTo 0
End Sub

' 窗体停用时处理滚轮支持
Private Sub UserForm_Deactivate()
    On Error Resume Next
    Debug.Print "UserForm_Deactivate 调用"
    
    ' 可选：可以在这里暂时禁用滚轮支持，但通常MouseWheelHandler会自动处理
    On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    
    ' 移除所有大小限制和控件调整逻辑
    ' 完全使用VBE设计器中设置的原始尺寸
    
End Sub

' 窗体关闭时保存位置
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 禁用滚轮支持并清理
    Call DisableWheelSupport("UserForm_QueryClose")
    Call SaveFormPosition
    
    ' 清理搜索管理器
    SearchManager.CleanupSearchManager
End Sub

' ====================================================================
' 控件事件处理
' ====================================================================

Private Sub ListBox1_Click()
    ' 确保ListBox有焦点，这样键盘事件能正常工作
    ListBox1.SetFocus
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox1.ListIndex < 0 Then Exit Sub
    
    Dim selectedText As String
    selectedText = ListBox1.List(ListBox1.ListIndex)
    
    ' 获取编号（第一列）
    Dim 编号 As String
    编号 = ListBox1.List(ListBox1.ListIndex, 0)  ' 直接获取第一列的值
    
    If Not Selection Is Nothing Then
        Selection.Cells(1, 1).Value = 编号
    End If
    
    Unload Me
End Sub

' Esc键关闭窗体和其他快捷键
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Debug.Print "检测到ESC键"
        Unload Me
    ' Ctrl+T 测试滚轮功能 (Shift=2表示Ctrl键)
    ElseIf KeyCode = vbKeyT And Shift = 2 Then
        Debug.Print "检测到Ctrl+T"
        FormDebugHelper.TestWheelFunction Me
    ' Ctrl+D 诊断功能
    ElseIf KeyCode = vbKeyD And Shift = 2 Then
        Debug.Print "检测到Ctrl+D"
        FormDebugHelper.DiagnoseWheelIssue Me
    ' Ctrl+S 手动滚动测试
    ElseIf KeyCode = vbKeyS And Shift = 2 Then
        Debug.Print "检测到Ctrl+S"
        FormDebugHelper.ManualTestScroll Me
    ' Ctrl+V 验证ListBox滚轮
    ElseIf KeyCode = vbKeyV And Shift = 2 Then
        Debug.Print "检测到Ctrl+V"
        ' 直接调用测试方法，不依赖FormDebugHelper
        Call TestWheelFunction
    ' Ctrl+B 强制重新初始化滚轮支持
    ElseIf KeyCode = vbKeyB And Shift = 2 Then
        Debug.Print "检测到Ctrl+B - 强制重新初始化滚轮支持"
        Call ForceReinitializeWheelSupport
    ' Ctrl+F 显示焦点状态
    ElseIf KeyCode = vbKeyF And Shift = 2 Then
        Debug.Print "检测到Ctrl+F"
        FormDebugHelper.ShowFocusStatus Me
    ' Ctrl+R 清除搜索并重置显示
    ElseIf KeyCode = vbKeyR And Shift = 2 Then
        Debug.Print "检测到Ctrl+R"
        TextBox1.Text = ""
        SearchManager.ClearSearchResults
        Debug.Print "搜索已清除，显示所有数据"
    End If
End Sub

Private Sub UserForm_Click()
    ' 点击窗体时也给ListBox设置焦点
    ListBox1.SetFocus
End Sub

' TextBox1 按回车键时进行搜索
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call PerformSearch
        ' 搜索后保持TextBox1焦点，方便连续搜索
    ElseIf KeyCode = vbKeyEscape Then
        ' ESC键清空搜索框并返回ListBox
        TextBox1.Text = ""
        SearchManager.ClearSearchResults
        ListBox1.SetFocus
    End If
End Sub

' TextBox1 获得焦点时显示提示
Private Sub TextBox1_Enter()
    ' 在状态栏或调试窗口显示使用提示
    Debug.Print "搜索提示：" & TextBox1.Tag
End Sub

' CommandButton1 点击时上一条
Private Sub CommandButton1_Click()
    If Not SearchManager.GotoPreviousResult() Then
        MsgBox "请先执行搜索", vbInformation, "提示"
        TextBox1.SetFocus
    End If
End Sub

' CommandButton2 点击时下一条
Private Sub CommandButton2_Click()
    If Not SearchManager.GotoNextResult() Then
        MsgBox "请先执行搜索", vbInformation, "提示"
        TextBox1.SetFocus
    End If
End Sub

' 优化的焦点管理方案
' VBA中滚轮事件捕获与焦点密切相关，我们采用以下策略：
' 1. 对于非交互控件（标签），鼠标悬停时让ListBox获得焦点
' 2. 对于交互控件（文本框、按钮），保持其正常的焦点行为
' 3. 同时通过MouseWheelHandler模块捕获全局滚轮事件

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 确保ListBox有焦点以接收键盘事件
    If Not ListBox1.Enabled Then Exit Sub
    ListBox1.SetFocus
End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 标签区域鼠标悬停时设置ListBox焦点，这是获取滚轮事件的基础方式
    If Not ListBox1.Enabled Then Exit Sub
    ListBox1.SetFocus
End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 标签区域鼠标悬停时设置ListBox焦点
    If Not ListBox1.Enabled Then Exit Sub
    ListBox1.SetFocus
End Sub

' TextBox1保留原有功能，不自动设置焦点
Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 保持默认行为，允许用户正常使用搜索框
    ' 注意：在TextBox中时滚轮事件可能无法直接传递到ListBox
    ' 这是VBA表单的限制，我们通过MouseWheelHandler尽量处理
End Sub

' 按钮保留原有功能，不自动设置焦点
Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 保持默认行为，允许用户正常点击按钮
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 保持默认行为，允许用户正常点击按钮
End Sub

' 添加UserForm_MouseMove事件，当鼠标在窗体空白区域时设置ListBox焦点
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 当鼠标在窗体空白区域时，确保ListBox有焦点以接收滚轮事件
    If Not ListBox1.Enabled Then Exit Sub
    ListBox1.SetFocus
End Sub

' VBA表单不原生支持UserForm_MouseWheel事件，已移除
' 滚轮支持将通过MouseWheelHandler模块的窗口子类化实现

' ====================================================================
' 核心功能方法
' ====================================================================

' 执行搜索 - 简化版，使用SearchManager模块
Private Sub PerformSearch()
    If Trim(TextBox1.Text) = "" Then
        MsgBox "请输入要搜索的内容", vbInformation, "搜索提示"
        Exit Sub
    End If
    
    Dim resultCount As Long
    resultCount = SearchManager.PerformNewSearch(TextBox1.Text)
    
    If resultCount = 0 Then
        MsgBox "未找到包含 '" & Trim(TextBox1.Text) & "' 的项目", vbInformation, "搜索结果"
        ' 选中所有文本便于重新输入
        TextBox1.SelStart = 0
        TextBox1.SelLength = Len(TextBox1.Text)
    End If
End Sub

' 确保指定项在ListBox中可见并居中显示（公共方法）
Public Sub EnsureItemVisibleInForm(itemIndex As Long)
    If itemIndex < 0 Or itemIndex >= ListBox1.ListCount Then Exit Sub
    
    ' 使用ListBox的实际行高计算可见行数，而不是固定值
    Dim actualRowHeight As Single
    Dim visibleRows As Long
    
    ' 使用VBE兼容的方式设置行高
    ' 直接使用默认行高值，确保兼容性
    actualRowHeight = 15 ' 在Excel VBA中，这是ListBox控件的标准默认行高
    
    ' 计算实际可见行数（考虑ListBox的实际高度）
    visibleRows = Int(ListBox1.Height / actualRowHeight)
    
    ' 确保至少有1行可见
    If visibleRows < 1 Then visibleRows = 1
    
    On Error GoTo 0
    
    ' 计算理想的TopIndex，使选中项尽量居中
    Dim idealTopIndex As Long
    idealTopIndex = itemIndex - Int(visibleRows / 2)
    
    ' 边界检查
    If idealTopIndex < 0 Then idealTopIndex = 0
    If idealTopIndex > ListBox1.ListCount - visibleRows Then
        idealTopIndex = ListBox1.ListCount - visibleRows
        ' 确保不小于0
        If idealTopIndex < 0 Then idealTopIndex = 0
    End If
    
    ' 关键：直接设置ListIndex然后再设置TopIndex，确保索引一致性
    ListBox1.ListIndex = itemIndex
    ListBox1.TopIndex = idealTopIndex
    
    ' 强制刷新显示
    Me.Repaint
    DoEvents
End Sub

' 公共方法：供标准模块调用的滚动功能（增强版）
Public Sub ScrollListBox(ByVal scrollLines As Long)
    On Error GoTo ScrollError
    
    Debug.Print "=== ScrollListBox被调用，滚动行数: " & scrollLines & " ==="
    
    ' 确保ListBox可用
    If Not ListBox1.Enabled Then
        Debug.Print "ListBox未启用，无法滚动"
        Exit Sub
    End If
    
    If ListBox1.ListCount = 0 Then
        Debug.Print "ListBox为空，无法滚动"
        Exit Sub
    End If
    
    ' 获取当前索引，确保有效
    Dim currentIndex As Long
    currentIndex = ListBox1.ListIndex
    If currentIndex = -1 Then currentIndex = 0
    
    Debug.Print "当前索引: " & currentIndex & ", ListBox总数: " & ListBox1.ListCount
    
    ' 计算新的索引
    Dim newIndex As Long
    newIndex = currentIndex + scrollLines
    
    ' 边界检查
    If newIndex < 0 Then newIndex = 0
    If newIndex >= ListBox1.ListCount Then newIndex = ListBox1.ListCount - 1
    
    ' 只有索引真正改变时才设置
    If newIndex <> currentIndex Then
        ' 先设置ListIndex
        ListBox1.ListIndex = newIndex
        
        ' 确保项可见并居中
        Call EnsureItemVisibleInForm(newIndex)
        
        Debug.Print "滚动成功: " & currentIndex & " -> " & newIndex & ", TopIndex: " & ListBox1.TopIndex
        
        ' 强制刷新
        Me.Repaint
        DoEvents
    Else
        Debug.Print "已到达边界，无法继续滚动"
    End If
    
    Exit Sub
    
ScrollError:
    Debug.Print "滚动错误: " & Err.Description & " (错误号: " & Err.Number & ")"
    On Error Resume Next
End Sub

' ====================================================================
' 滚轮支持功能
' ====================================================================

' 启用滚轮支持（增强版）
Private Sub EnableWheelSupport()
    On Error GoTo ErrorHandler
    
    If m_WheelEnabled Then
        Debug.Print "滚轮支持已经启用，跳过重复调用"
        Exit Sub
    End If
    
    Debug.Print "=== 开始启用滚轮支持 ==="
    
    ' 确保ListBox1获得焦点
    ListBox1.SetFocus
    Debug.Print "ListBox1焦点状态: " & ListBox1.Enabled
    
    ' 先清理旧的支持（如果有）
    On Error Resume Next
    MouseWheelHandler.DisableMouseWheelSupport ("EnableWheelSupport前清理")
    On Error GoTo ErrorHandler
    
    ' 调用MouseWheelHandler启用滚轮支持
    Debug.Print "调用MouseWheelHandler.EnableMouseWheelSupport..."
    Dim result As Boolean
    result = MouseWheelHandler.EnableMouseWheelSupport(Me.Caption, Me)
    
    If result Then
        m_WheelEnabled = True
        Debug.Print "滚轮支持已成功启用"
        
        ' 额外设置ListBox引用
        MouseWheelHandler.SetActiveListBox ListBox1
        Debug.Print "已设置活动ListBox引用"
        
        ' 测试滚轮捕获
        Debug.Print "执行滚轮捕获测试..."
        On Error Resume Next
        MouseWheelHandler.TestMouseWheelCapture
        On Error GoTo ErrorHandler
    Else
        Debug.Print "滚轮支持启用失败，但继续执行"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "滚轮启用出错: " & Err.Description & " (错误号: " & Err.Number & ")"
    On Error Resume Next
    ListBox1.SetFocus ' 出错时至少确保ListBox有焦点
End Sub

' 清理滚轮支持（增强版）
Private Sub DisableWheelSupport(Optional ByVal caller As String = "未知")
    On Error Resume Next
    
    Debug.Print "=== 开始禁用滚轮支持 - 调用来源: " & caller & " ==="
    
    ' 无论是否启用都尝试清理，确保资源释放
    Dim wasEnabled As Boolean
    wasEnabled = m_WheelEnabled
    
    ' 清理MouseWheelHandler
    MouseWheelHandler.DisableMouseWheelSupport caller
    
    ' 重置状态
    m_WheelEnabled = False
    
    If wasEnabled Then
        Debug.Print "滚轮支持已成功禁用 - 调用来源: " & caller
    Else
        Debug.Print "滚轮支持本来就是禁用状态 - 调用来源: " & caller
    End If
    
    On Error GoTo 0
End Sub

' 测试滚轮功能的辅助方法
Private Sub TestWheelFunction()
    On Error GoTo TestError
    
    Debug.Print "=== 开始测试滚轮功能 ==="
    
    ' 测试ListBox基本操作
    Debug.Print "ListBox1可见性: " & ListBox1.Visible
    Debug.Print "ListBox1启用状态: " & ListBox1.Enabled
    Debug.Print "ListBox1项目数量: " & ListBox1.ListCount
    
    ' 执行手动滚动测试
    Debug.Print "执行手动滚动测试..."
    
    ' 如果ListBox不为空，进行向上和向下的滚动测试
    If ListBox1.ListCount > 0 Then
        ' 保存当前索引
        Dim saveIndex As Long
        saveIndex = ListBox1.ListIndex
        If saveIndex = -1 Then saveIndex = 0
        
        Debug.Print "保存的索引: " & saveIndex
        
        ' 测试向下滚动
        Debug.Print "测试向下滚动..."
        Call ScrollListBox(3)
        
        ' 测试向上滚动
        Debug.Print "测试向上滚动..."
        Call ScrollListBox(-3)
        
        ' 恢复原来的索引
        If saveIndex >= 0 And saveIndex < ListBox1.ListCount Then
            ListBox1.ListIndex = saveIndex
            Call EnsureItemVisibleInForm(saveIndex)
            Debug.Print "已恢复到原来的索引"
        End If
    Else
        Debug.Print "ListBox为空，无法进行滚动测试"
    End If
    
    ' 测试MouseWheelHandler功能
    Debug.Print "测试MouseWheelHandler功能..."
    On Error Resume Next
    Dim isWheelCaptured As Boolean
    isWheelCaptured = MouseWheelHandler.isWheelCaptured
    Debug.Print "MouseWheelHandler捕获状态: " & isWheelCaptured
    
    ' 打印MouseWheelHandler内部状态
    Debug.Print "MouseWheelHandler内部状态信息:"
    MouseWheelHandler.PrintInternalState
    On Error GoTo TestError
    
    Debug.Print "=== 滚轮功能测试完成 ==="
    
    Exit Sub
    
TestError:
    Debug.Print "滚轮测试错误: " & Err.Description & " (错误号: " & Err.Number & ")"
    On Error Resume Next
End Sub

' 强制重新初始化滚轮支持
Private Sub ForceReinitializeWheelSupport()
    On Error GoTo ReinitError
    
    Debug.Print "=== 开始强制重新初始化滚轮支持 ==="
    
    ' 1. 首先完全清理现有的滚轮支持
    Debug.Print "步骤1: 清理现有滚轮支持..."
    Call DisableWheelSupport("ForceReinitializeWheelSupport清理阶段")
    
    ' 强制暂停，确保清理完成
    Debug.Print "等待清理完成..."
    DoEvents
    
    ' 2. 确保ListBox获得焦点
    Debug.Print "步骤2: 设置ListBox焦点..."
    ListBox1.SetFocus
    Debug.Print "ListBox焦点设置状态: " & ListBox1.Enabled
    
    ' 3. 重新初始化滚轮支持
    Debug.Print "步骤3: 重新初始化滚轮支持..."
    Call EnableWheelSupport
    
    ' 4. 验证重新初始化结果
    Debug.Print "步骤4: 验证重新初始化结果..."
    If m_WheelEnabled Then
        Debug.Print "? 滚轮支持重新初始化成功!"
        
        ' 额外的测试
        Call TestWheelFunction
    Else
        Debug.Print "? 滚轮支持重新初始化失败!"
        
        ' 尝试直接设置MouseHook
        Debug.Print "尝试直接重新注册鼠标钩子..."
        On Error Resume Next
        MouseWheelHandler.ReinitializeMouseHook
        Debug.Print "鼠标钩子重新注册尝试完成"
        On Error GoTo ReinitError
    End If
    
    Debug.Print "=== 强制重新初始化滚轮支持完成 ==="
    
    Exit Sub
    
ReinitError:
    Debug.Print "滚轮重新初始化错误: " & Err.Description & " (错误号: " & Err.Number & ")"
    On Error Resume Next
End Sub

' 检查滚轮是否启用（供调试模块使用）
Public Function IsWheelEnabled() As Boolean
    IsWheelEnabled = m_WheelEnabled
End Function

' === 新增：加载数据的通用方法 ===
Private Sub LoadRateData()
    On Error Resume Next
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRng As Range
    
    ' ==========================================================
    ' >>> 数据引用位置 - 重要：用户需要根据实际情况修改此处 <<<
    ' ==========================================================
    ' 1. 工作表名称 - 修改为实际存放定额数据的工作表
    Set ws = ThisWorkbook.Worksheets("RATE")
    
    If ws Is Nothing Then Exit Sub
    
    ' 2. 数据起始位置和范围确定 - 修改为实际数据所在位置
    ' 假设数据从第2行开始（第1行是标题），A列是最后一行依据
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        ListBox1.Clear
        Exit Sub
    End If
    
    ' 3. 数据列范围 - 修改为两个范围：A2:C(编号, 名称, 单位) 和 J2:J(单价)
    On Error GoTo ErrorHandler
    
    ' 先读取A2:C范围的数据
    Dim range1 As Range
    Set range1 = ws.Range("A2:C" & lastRow)
    
    ' 再读取J2:J范围的数据（单价列）
    Dim range2 As Range
    Set range2 = ws.Range("J2:J" & lastRow)
    
    ' 获取数据
    Dim data1 As Variant
    Dim data2 As Variant
    data1 = range1.Value
    data2 = range2.Value
    
    ' 验证数据行数是否匹配
    Dim rows1 As Long
    Dim rows2 As Long
    
    ' 处理单列数据的特殊情况（如果只有一行数据，UBound无法正常工作）
    If TypeName(data1) = "Variant()" Then
        rows1 = UBound(data1, 1)
    Else
        rows1 = 1 ' 只有一行数据
    End If
    
    If TypeName(data2) = "Variant()" Then
        rows2 = UBound(data2, 1)
    Else
        rows2 = 1 ' 只有一行数据
    End If
    
    ' 确保两个范围的数据行数相同
    Dim actualRows As Long
    actualRows = WorksheetFunction.Min(rows1, rows2)
    
    ' 准备4列的数组（ListBox需要4列数据）
    Dim resultData() As Variant
    ReDim resultData(1 To actualRows, 1 To 4)
    
    ' 合并数据
    Dim i As Long
    For i = 1 To actualRows
        ' 安全地复制数据，添加错误处理
        On Error Resume Next
        ' 复制A-C列数据到前3列
        If TypeName(data1) = "Variant()" Then
            resultData(i, 1) = data1(i, 1) ' 编号
            resultData(i, 2) = data1(i, 2) ' 名称
            resultData(i, 3) = data1(i, 3) ' 单位
        Else
            ' 只有一行数据的情况
            resultData(i, 1) = data1
            resultData(i, 2) = ws.Cells(2, 2).Value
            resultData(i, 3) = ws.Cells(2, 3).Value
        End If
        
        ' 复制J列数据到第4列（单价）
        If TypeName(data2) = "Variant()" Then
            resultData(i, 4) = data2(i, 1) ' 单价
        Else
            ' 只有一行数据的情况
            resultData(i, 4) = data2
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' 设置ListBox数据
    ListBox1.Clear ' 先清空ListBox
    ListBox1.List = resultData
    
    ' 重置错误处理
    On Error GoTo 0
    Exit Sub
    
ErrorHandler:
    ' 发生错误时的处理
    ListBox1.Clear
    If Err.Number <> 0 Then
        Debug.Print "数据加载错误: " & Err.Description
        ' 可以选择显示错误消息或继续执行
    End If
    On Error GoTo 0
    ' ==========================================================
    ' >>> 数据引用位置结束 <<<
    ' ==========================================================
    
    ' 更新 Label2 状态
    Label2.Caption = "0/" & ListBox1.ListCount
    
    ' 如果使用了 SearchManager，需要重置搜索缓存
    SearchManager.InitializeSearchManager ListBox1, Label2
End Sub


' ====================================================================
' 窗体位置管理
' ====================================================================

' 恢复窗体位置 (已修复定位 BUG)
Private Sub RestoreFormPosition()
    On Error Resume Next
    
    Dim regKey As String
    regKey = "Software\Excel\RateLibrary"
    
    ' 从注册表读取上次保存的位置
    Dim savedLeft As Long, savedTop As Long, savedWidth As Long, savedHeight As Long
    savedLeft = GetSetting("RateLibrary", "Position", "Left", -1)
    savedTop = GetSetting("RateLibrary", "Position", "Top", -1)
    
    ' [引用源: 38] 读取宽高
    savedWidth = GetSetting("RateLibrary", "Position", "Width", Me.Width)
    savedHeight = GetSetting("RateLibrary", "Position", "Height", Me.Height)
    
    ' 获取当前活动窗口的可视区域信息（比 Application 更准确）
    Dim winLeft As Long, winTop As Long, winWidth As Long, winHeight As Long
    winLeft = Application.ActiveWindow.Left
    winTop = Application.ActiveWindow.Top
    winWidth = Application.ActiveWindow.Width
    winHeight = Application.ActiveWindow.Height

    ' 判断是否需要重置位置：
    ' 1. 首次运行 (savedLeft = -1)
    ' 2. 上次保存的位置已经跑到了当前窗口的外面 (例如窗口变窄了)
    Dim needReset As Boolean
    needReset = (savedLeft = -1) Or (savedTop = -1)
    
    ' 如果保存的左边距 > 当前窗口宽度，说明窗口变小了，需要拉回来
    If savedLeft > (winWidth - 50) Then needReset = True
    If savedTop > (winHeight - 50) Then needReset = True
    
    If needReset Then
        ' 定位到Excel窗口的左侧，留出一点边距
        savedLeft = Application.Left + winLeft + 10 ' 左侧边距10像素
        savedTop = Application.Top + winTop + 50   ' 顶部稍微靠下，避开功能区
        
        ' 简单的边界保护，防止负数
        If savedLeft < 0 Then savedLeft = 0
        If savedTop < 0 Then savedTop = 0
        
        ' 使用VBE设计器中设置的原始窗体大小，不设置默认值
    End If
    
    ' 设置窗体位置和大小
    Me.Width = savedWidth
    Me.Height = savedHeight
    
    ' 使用 API 设置位置
    SetWindowPos hWndForm, 0, savedLeft, savedTop, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
End Sub

' 保存窗体位置
Private Sub SaveFormPosition()
    On Error Resume Next
    
    ' 获取当前窗体位置和大小
    SaveSetting "RateLibrary", "Position", "Left", Me.Left
    SaveSetting "RateLibrary", "Position", "Top", Me.Top
    SaveSetting "RateLibrary", "Position", "Width", Me.Width
    SaveSetting "RateLibrary", "Position", "Height", Me.Height
End Sub


