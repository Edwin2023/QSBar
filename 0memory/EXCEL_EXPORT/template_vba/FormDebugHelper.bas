Attribute VB_Name = "FormDebugHelper"
' 窗体调试辅助模块 - 处理调试和测试功能
' 将调试逻辑从窗体中分离，减少窗体代码复杂度

Option Explicit

' ====================================================================
' 滚轮测试功能
' ====================================================================

' 测试滚轮功能
Public Sub TestWheelFunction(ByRef form As Object)
    Debug.Print "=== 测试滚轮功能 ==="
    
    ' 检查是否已启用
    If MouseWheelHandler.IsMouseWheelEnabled() Then
        Debug.Print "? 滚轮支持已启用"
    Else
        Debug.Print "? 滚轮支持未启用"
        ' 尝试重新启用
        Debug.Print "尝试重新启用滚轮支持..."
        form.EnableWheelSupport
    End If
    
    ' 显示详细状态
    Debug.Print "窗体标题: " & form.Caption
    Debug.Print "ListBox项目数: " & form.ListBox1.ListCount
    Debug.Print "当前选中索引: " & form.ListBox1.ListIndex
    
    ' 手动测试滚动
    Debug.Print "执行手动向下滚动测试..."
    form.ScrollListBox 3
    
    Debug.Print "执行手动向上滚动测试..."
    form.ScrollListBox -3
    
    Debug.Print "=== 测试完成 ==="
    Debug.Print "现在请在ListBox区域内用鼠标滚轮滚动"
    Debug.Print "如果工作正常，您应该看到滚轮消息和滚动效果"
End Sub

' 诊断滚轮问题
Public Sub DiagnoseWheelIssue(ByRef form As Object)
    Debug.Print "=== 滚轮问题诊断 ==="
    
    ' 1. 检查MouseWheelHandler状态
    Debug.Print "1. MouseWheelHandler状态检查:"
    Debug.Print "   - IsActive: " & MouseWheelHandler.isActive()
    Debug.Print "   - IsEnabled: " & MouseWheelHandler.IsMouseWheelEnabled()
    
    ' 2. 检查窗体状态
    Debug.Print "2. 窗体状态检查:"
    Debug.Print "   - 窗体标题: " & form.Caption
    Debug.Print "   - 窗体可见: " & form.Visible
    Debug.Print "   - ListBox项目数: " & form.ListBox1.ListCount
    Debug.Print "   - 当前选中: " & form.ListBox1.ListIndex
    Debug.Print "   - TopIndex: " & form.ListBox1.TopIndex
    
    ' 3. 测试滚动方法
    Debug.Print "3. 测试ScrollListBox方法:"
    Dim oldIndex As Long, oldTopIndex As Long
    oldIndex = form.ListBox1.ListIndex
    oldTopIndex = form.ListBox1.TopIndex
    
    Debug.Print "   执行前: Index=" & oldIndex & ", TopIndex=" & oldTopIndex
    form.ScrollListBox 1
    Debug.Print "   执行后: Index=" & form.ListBox1.ListIndex & ", TopIndex=" & form.ListBox1.TopIndex
    
    If form.ListBox1.ListIndex <> oldIndex Or form.ListBox1.TopIndex <> oldTopIndex Then
        Debug.Print "   ? ScrollListBox方法工作正常"
    Else
        Debug.Print "   ? ScrollListBox方法无效果"
    End If
    
    Debug.Print "=== 诊断完成 ==="
    Debug.Print "快捷键测试："
    Debug.Print "- Ctrl+T: 滚轮功能测试"
    Debug.Print "- Ctrl+D: 诊断测试（当前命令）"
    Debug.Print "- Ctrl+S: 手动滚动测试"
    Debug.Print "- Ctrl+V: 验证ListBox滚轮"
    Debug.Print "- Ctrl+F: 显示焦点状态"
    Debug.Print "- Ctrl+R: 清除搜索重置显示"
    Debug.Print "- ESC: 关闭窗体"
End Sub

' 手动测试滚动
Public Sub ManualTestScroll(ByRef form As Object, Optional ByVal direction As Long = 1)
    Debug.Print "=== 手动测试滚动 ==="
    Debug.Print "方向: " & IIf(direction > 0, "向下", "向上")
    
    ' 模拟滚轮调用
    If direction > 0 Then
        form.ScrollListBox 3  ' 向下滚动
    Else
        form.ScrollListBox -3 ' 向上滚动
    End If
    
    Debug.Print "=== 手动滚动测试完成 ==="
End Sub

' 验证ListBox滚轮
Public Sub VerifyListBoxWheel(ByRef form As Object)
    Debug.Print "=== ListBox滚轮验证测试 ==="
    Debug.Print "请在ListBox区域（文字数据上）滚动鼠标滚轮"
    Debug.Print "观察以下信息："
    Debug.Print "1. 是否显示滚轮消息和TopIndex变化"
    Debug.Print "2. TopIndex是否改变: 当前TopIndex=" & form.ListBox1.TopIndex
    Debug.Print "3. 当前选中索引: " & form.ListBox1.ListIndex
    Debug.Print "4. ESC键是否正常工作"
    Debug.Print "=================================="
End Sub

' 显示焦点状态
Public Sub ShowFocusStatus(ByRef form As Object)
    Debug.Print "=== 当前焦点状态 ==="
    On Error Resume Next
    Dim focusInfo As String
    Dim errorOccurred As Boolean
    
    ' 安全地检测当前焦点控件
    Dim activeCtrl As Object
    Set activeCtrl = form.ActiveControl
    
    If Err.Number = 0 And Not activeCtrl Is Nothing Then
        focusInfo = "当前焦点控件: " & TypeName(activeCtrl)
        Select Case TypeName(activeCtrl)
            Case "ListBox"
                focusInfo = focusInfo & " (ListBox1)"
            Case "TextBox"
                focusInfo = focusInfo & " (TextBox1-搜索框)"
            Case "CommandButton"
                focusInfo = focusInfo & " (CommandButton)"
            Case Else
                focusInfo = focusInfo & " (其他控件)"
        End Select
    Else
        focusInfo = "无法确定焦点位置 (可能在窗体或出错)"
        If Err.Number <> 0 Then
            focusInfo = focusInfo & " - 错误: " & Err.Description
        End If
    End If
    
    Debug.Print focusInfo
    Debug.Print "ListBox项目数: " & form.ListBox1.ListCount
    Debug.Print "ListBox TopIndex: " & form.ListBox1.TopIndex
    Debug.Print "ListBox选中索引: " & form.ListBox1.ListIndex
    Debug.Print "搜索结果数: " & SearchManager.GetSearchResultCount()
    Debug.Print "========================"
    On Error GoTo 0
End Sub

' ====================================================================
' 窗体状态检查功能
' ====================================================================

' 显示窗体滚轮状态
Public Sub ShowWheelStatus(ByRef form As Object)
    Debug.Print "=== 滚轮状态检查 ==="
    Debug.Print "窗体滚轮启用状态: " & form.IsWheelEnabled()
    Debug.Print "MouseWheelHandler活动状态: " & MouseWheelHandler.isActive()
    Debug.Print "ListBox项目数: " & form.ListBox1.ListCount
    Debug.Print "当前选中索引: " & form.ListBox1.ListIndex
    Debug.Print "=== 状态检查完成 ==="
End Sub

' 测试滚轮效果
Public Sub TestWheelEffect(ByRef form As Object)
    Debug.Print "=== 测试滚轮效果 ==="
    
    If form.ListBox1.ListCount < 10 Then
        Debug.Print "ListBox项目太少，无法测试滚动效果"
        Exit Sub
    End If
    
    ' 保存当前状态
    Dim originalIndex As Long, originalTopIndex As Long
    originalIndex = form.ListBox1.ListIndex
    originalTopIndex = form.ListBox1.TopIndex
    
    Debug.Print "原始状态: Index=" & originalIndex & ", TopIndex=" & originalTopIndex
    
    ' 测试向下滚动
    Debug.Print "测试向下滚动..."
    form.ScrollListBox 5
    Debug.Print "向下滚动后: Index=" & form.ListBox1.ListIndex & ", TopIndex=" & form.ListBox1.TopIndex
    
    ' 等待一秒
    Application.Wait Now + TimeValue("00:00:01")
    
    ' 测试向上滚动
    Debug.Print "测试向上滚动..."
    form.ScrollListBox -5
    Debug.Print "向上滚动后: Index=" & form.ListBox1.ListIndex & ", TopIndex=" & form.ListBox1.TopIndex
    
    ' 恢复原始位置
    form.ListBox1.ListIndex = originalIndex
    form.ListBox1.TopIndex = originalTopIndex
    
    Debug.Print "=== 滚轮效果测试完成 ==="
End Sub

' 模拟滚轮测试
Public Sub TestSimulateWheel(ByRef form As Object)
    Debug.Print "=== 模拟滚轮测试 ==="
    
    ' 模拟向下滚动
    Debug.Print "模拟向下滚动..."
    form.ScrollListBox 3
    
    ' 模拟向上滚动
    Debug.Print "模拟向上滚动..."
    form.ScrollListBox -3
    
    Debug.Print "模拟测试完成。如果ListBox有滚动，说明ScrollListBox方法正常"
End Sub
