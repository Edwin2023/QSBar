Attribute VB_Name = "MouseWheelHandler"
' 安全的鼠标滚轮处理模块 - 单一模块解决方案
' 基于窗口子类化，集成了安全机制避免循环和错误

Option Explicit

#If Win64 Then
    Private Declare PtrSafe Function SetWindowLongPtr _
        Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal Hwnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr) _
        As LongPtr
        
    Private Declare PtrSafe Function CallWindowProc _
        Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As LongPtr, _
        ByVal Hwnd As LongPtr, _
        ByVal Msg As Long, _
        ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr) _
        As LongPtr
        
    Private Declare PtrSafe Function FindWindow _
        Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) _
        As LongPtr
        
    Private Declare PtrSafe Function FindWindowEx _
        Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, _
        ByVal hWndChildAfter As LongPtr, _
        ByVal lpszClass As String, _
        ByVal lpszWindow As String) _
        As LongPtr
        
    Private Declare PtrSafe Function GetWindowLongPtr _
        Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal Hwnd As LongPtr, _
        ByVal nIndex As Long) _
        As LongPtr

    Private m_OriginalProc As LongPtr
    Private m_FormHwnd As LongPtr
    Private m_ListBoxHwnd As LongPtr
    Private m_ListBoxOriginalProc As LongPtr
#Else
    Private Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) _
        As Long
        
    Private Declare Function CallWindowProc _
        Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As Long, _
        ByVal hWnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) _
        As Long
        
    Private Declare Function FindWindow _
        Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) _
        As Long
        
    Private Declare Function FindWindowEx _
        Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As Long, _
        ByVal hWndChildAfter As Long, _
        ByVal lpszClass As String, _
        ByVal lpszWindow As String) _
        As Long
        
    Private Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) _
        As Long

    Private m_OriginalProc As Long
    Private m_FormHwnd As Long
    Private m_ListBoxHwnd As Long
    Private m_ListBoxOriginalProc As Long
#End If

' 常量定义
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000

' 全局变量
Private m_WheelEnabled As Boolean
Private m_ActiveForm As Object
Private m_ActiveListBox As Object

' ====================================================================
' 公共接口函数
' ====================================================================

' 启用鼠标滚轮支持
Public Function EnableMouseWheelSupport(ByVal formCaption As String, ByRef formRef As Object) As Boolean
    On Error GoTo ErrorHandler
    
    If m_WheelEnabled Then
        EnableMouseWheelSupport = True
        Exit Function
    End If
    
    ' 保存窗体和ListBox引用
    Set m_ActiveForm = formRef
    Set m_ActiveListBox = formRef.ListBox1  ' 假设ListBox名称为ListBox1
    
    ' 查找窗体窗口 - 尝试多种方式
    m_FormHwnd = FindWindow("ThunderDFrame", formCaption)
    Debug.Print "查找ThunderDFrame窗口: " & formCaption & ", 句柄: " & Hex(m_FormHwnd)
    
    ' 如果ThunderDFrame找不到，尝试其他类名
    If m_FormHwnd = 0 Then
        m_FormHwnd = FindWindow(vbNullString, formCaption)
        Debug.Print "通过标题查找窗口: " & formCaption & ", 句柄: " & Hex(m_FormHwnd)
    End If
    
    ' Excel环境特殊处理：尝试获取当前活动窗口
    If m_FormHwnd = 0 Then
        ' 在Excel VBA环境中，可能需要不同的方法
        Debug.Print "尝试Excel VBA环境的窗口查找方法..."
        m_FormHwnd = FindWindow("ThunderDFrame", vbNullString)
        Debug.Print "Excel环境窗口句柄: " & Hex(m_FormHwnd)
    End If
    
    If m_FormHwnd <> 0 Then
        ' 子类化窗口
        #If Win64 Then
            m_OriginalProc = SetWindowLongPtr(m_FormHwnd, GWL_WNDPROC, AddressOf SafeWindowProc)
        #Else
            m_OriginalProc = SetWindowLong(m_FormHwnd, GWL_WNDPROC, AddressOf SafeWindowProc)
        #End If
        
        If m_OriginalProc <> 0 Then
            ' 尝试查找并子类化ListBox控件 - 遍历子窗口
            #If Win64 Then
                Dim childHwnd As LongPtr
            #Else
                Dim childHwnd As Long
            #End If
            childHwnd = FindWindowEx(m_FormHwnd, 0, vbNullString, vbNullString)
            While childHwnd <> 0
                Debug.Print "找到子窗口: " & Hex(childHwnd)
                ' 通常第一个子窗口就是ListBox，或者我们可以尝试所有子窗口
                If m_ListBoxHwnd = 0 Then m_ListBoxHwnd = childHwnd  ' 使用第一个找到的
                childHwnd = FindWindowEx(m_FormHwnd, childHwnd, vbNullString, vbNullString)
            Wend
            Debug.Print "选择ListBox窗口，句柄: " & Hex(m_ListBoxHwnd)
            
            ' 如果找到ListBox，也对它进行子类化
            If m_ListBoxHwnd <> 0 Then
                #If Win64 Then
                    m_ListBoxOriginalProc = SetWindowLongPtr(m_ListBoxHwnd, GWL_WNDPROC, AddressOf SafeWindowProc)
                #Else
                    m_ListBoxOriginalProc = SetWindowLong(m_ListBoxHwnd, GWL_WNDPROC, AddressOf SafeWindowProc)
                #End If
                
                If m_ListBoxOriginalProc <> 0 Then
                    Debug.Print "? ListBox也已子类化: " & Hex(m_ListBoxOriginalProc)
                    
                    ' ?? 尝试禁用ListBox的WS_VSCROLL样式来阻止内置滚轮
                    Call DisableListBoxInternalScroll(m_ListBoxHwnd)
                End If
            End If
            
            m_WheelEnabled = True
            EnableMouseWheelSupport = True
            Debug.Print "? 滚轮支持已启用，窗体过程: " & Hex(m_OriginalProc)
        Else
            Debug.Print "? 窗口子类化失败，原始过程为0"
            EnableMouseWheelSupport = False
        End If
    Else
        Debug.Print "? 未找到窗体窗口: " & formCaption
        EnableMouseWheelSupport = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "启用滚轮支持失败: " & Err.Description & " (错误号: " & Err.Number & ")"
    EnableMouseWheelSupport = False
End Function

' 设置活动ListBox
Public Sub SetActiveListBox(ByRef listBox As Object)
    Set m_ActiveListBox = listBox
End Sub

' 禁用滚轮支持并清理
Public Sub DisableMouseWheelSupport(Optional ByVal caller As String = "MouseWheelHandler")
    On Error Resume Next
    
    If m_WheelEnabled And m_FormHwnd <> 0 And m_OriginalProc <> 0 Then
        ' 恢复ListBox的原始窗口过程
        If m_ListBoxHwnd <> 0 And m_ListBoxOriginalProc <> 0 Then
            #If Win64 Then
                SetWindowLongPtr m_ListBoxHwnd, GWL_WNDPROC, m_ListBoxOriginalProc
            #Else
                SetWindowLong m_ListBoxHwnd, GWL_WNDPROC, m_ListBoxOriginalProc
            #End If
            Debug.Print "? ListBox子类化已恢复"
        End If
        
        ' 恢复窗体的原始窗口过程
        #If Win64 Then
            SetWindowLongPtr m_FormHwnd, GWL_WNDPROC, m_OriginalProc
        #Else
            SetWindowLong m_FormHwnd, GWL_WNDPROC, m_OriginalProc
        #End If
        
        m_WheelEnabled = False
        m_OriginalProc = 0
        m_ListBoxOriginalProc = 0
        m_FormHwnd = 0
        m_ListBoxHwnd = 0
        Set m_ActiveForm = Nothing
        Set m_ActiveListBox = Nothing
        Debug.Print "? 滚轮支持已禁用 - 来自: " & caller
    Else
        ' 记录重复禁用的情况
        If Not m_WheelEnabled Then
            Debug.Print "注意：滚轮支持已经被禁用 - 重复调用来自: " & caller
        End If
    End If
End Sub

' 检查是否已启用
Public Function IsMouseWheelEnabled() As Boolean
    IsMouseWheelEnabled = m_WheelEnabled
End Function

' 检查是否活动
Public Function isActive() As Boolean
    isActive = m_WheelEnabled
End Function

' 获取版本信息
Public Function GetVersion() As String
    GetVersion = "MouseWheelHandler v3.0 - 安全单模块解决方案"
End Function

' ====================================================================
' 安全的窗口过程函数
' ====================================================================
#If Win64 Then
Public Function SafeWindowProc(ByVal Hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function SafeWindowProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
    
    ' 增强错误处理和递归检查
    Static recursionGuard As Boolean
    
    ' 防止递归调用
    If recursionGuard Then
        ' 根据窗口句柄选择正确的原始过程
        If Hwnd = m_FormHwnd And m_OriginalProc <> 0 Then
            SafeWindowProc = CallWindowProc(m_OriginalProc, Hwnd, uMsg, wParam, lParam)
        ElseIf Hwnd = m_ListBoxHwnd And m_ListBoxOriginalProc <> 0 Then
            SafeWindowProc = CallWindowProc(m_ListBoxOriginalProc, Hwnd, uMsg, wParam, lParam)
        Else
            SafeWindowProc = 0
        End If
        Exit Function
    End If
    
    recursionGuard = True
    
    On Error GoTo ErrorHandler
    
    ' 检测滚轮消息 - 使用多种方式检测
    If uMsg = WM_MOUSEWHEEL Or uMsg = &H20A Or uMsg = &H20E Then  ' 包括水平滚轮
        Debug.Print "?? 滚轮消息! uMsg=" & Hex(uMsg) & ", wParam=" & Hex(wParam) & ", lParam=" & Hex(lParam)
    End If
    
    ' 处理所有滚轮消息 - 不管来源窗口
    If (uMsg = WM_MOUSEWHEEL Or uMsg = &H20A) And m_WheelEnabled And Not m_ActiveListBox Is Nothing Then
        Debug.Print "???? 滚轮消息处理！uMsg: " & Hex(uMsg) & ", wParam: " & Hex(wParam) & ", 来源窗口: " & Hex(Hwnd)
        
        ' 提取滚轮方向（更安全的方式）
        Dim wheelDelta As Long
        ' 在32位系统中，需要使用不同的方法提取高位字
        #If Win64 Then
            wheelDelta = CLng((wParam And &HFFFF0000) / 65536)
        #Else
            ' 32位系统的处理方式
            wheelDelta = CLng(wParam) \ 65536
        #End If
        
        ' 转换为有符号整数
        If wheelDelta > 32767 Then wheelDelta = wheelDelta - 65536
        
        Debug.Print "滚轮方向值: " & wheelDelta
        
        ' ?? 新策略：统一处理窗体滚轮，直接操作ListBox的TopIndex
        On Error Resume Next
        If Not m_ActiveListBox Is Nothing Then
            ' 显示消息来源（用于调试）
            Debug.Print "?? 滚轮消息来源: " & IIf(Hwnd = m_FormHwnd, "窗体", "ListBox") & " (句柄: " & Hex(Hwnd) & ")"
            
            ' ?? 关键：无论消息来自哪里，都直接操作ListBox的TopIndex
            If wheelDelta > 0 Then
                ' 向上滚动（正值）
                Debug.Print "?? 向上滚动TopIndex"
                If m_ActiveListBox.TopIndex > 0 Then
                    If m_ActiveListBox.TopIndex > 3 Then
                        m_ActiveListBox.TopIndex = m_ActiveListBox.TopIndex - 3
                    Else
                        m_ActiveListBox.TopIndex = 0
                    End If
                End If
            Else
                ' 向下滚动（负值）
                Debug.Print "?? 向下滚动TopIndex"
                Dim maxTopIndex As Long
                maxTopIndex = m_ActiveListBox.ListCount - 1
                If m_ActiveListBox.TopIndex < maxTopIndex Then
                    m_ActiveListBox.TopIndex = m_ActiveListBox.TopIndex + 3
                End If
            End If
            
            ' 显示当前TopIndex（用于验证）
            Debug.Print "?? 当前TopIndex: " & m_ActiveListBox.TopIndex & "/" & (m_ActiveListBox.ListCount - 1)
        Else
            Debug.Print "? ListBox引用为空，无法滚动"
        End If
        On Error GoTo ErrorHandler
        
        ' ?? 关键：返回0并且不调用原始过程，完全阻止默认处理
        recursionGuard = False
        SafeWindowProc = 0
        Exit Function
    End If
    
    ' 调用正确的原始窗口过程处理其他消息
    If Hwnd = m_FormHwnd And m_OriginalProc <> 0 Then
        SafeWindowProc = CallWindowProc(m_OriginalProc, Hwnd, uMsg, wParam, lParam)
    ElseIf Hwnd = m_ListBoxHwnd And m_ListBoxOriginalProc <> 0 Then
        SafeWindowProc = CallWindowProc(m_ListBoxOriginalProc, Hwnd, uMsg, wParam, lParam)
    Else
        SafeWindowProc = 0
    End If
    
    recursionGuard = False
    Exit Function
    
ErrorHandler:
    recursionGuard = False
    ' 出错时调用正确的原始窗口过程
    If Hwnd = m_FormHwnd And m_OriginalProc <> 0 Then
        SafeWindowProc = CallWindowProc(m_OriginalProc, Hwnd, uMsg, wParam, lParam)
    ElseIf Hwnd = m_ListBoxHwnd And m_ListBoxOriginalProc <> 0 Then
        SafeWindowProc = CallWindowProc(m_ListBoxOriginalProc, Hwnd, uMsg, wParam, lParam)
    Else
        SafeWindowProc = 0
    End If
End Function

' ====================================================================
' 自动清理 - 完全注释掉避免意外禁用
' ====================================================================
' 注意：Auto_Close在某些情况下会被意外调用，导致滚轮支持被错误禁用
' 因此完全注释掉这个方法，让窗体自己负责清理滚轮支持
'Public Sub Auto_Close()
'    Debug.Print "Auto_Close被调用 - 禁用滚轮支持"
'    Call DisableMouseWheelSupport("Auto_Close")
'End Sub

' ====================================================================
' 滚轮测试和诊断方法
' ====================================================================

' 测试滚轮消息捕获
Public Sub TestMouseWheelCapture()
    Debug.Print "=== 滚轮消息捕获测试 ==="
    Debug.Print "当前状态:"
    Debug.Print "- 窗口句柄: " & Hex(m_FormHwnd)
    Debug.Print "- 原始过程: " & Hex(m_OriginalProc)
    Debug.Print "- 滚轮启用: " & m_WheelEnabled
    Debug.Print "- 窗体引用: " & Not (m_ActiveForm Is Nothing)
    Debug.Print "- ListBox引用: " & Not (m_ActiveListBox Is Nothing)
    Debug.Print ""
    Debug.Print "请在窗体上滚动鼠标滚轮..."
    Debug.Print "如果消息被捕获，您会看到 ?? 滚轮消息"
End Sub

' 强制重新启用滚轮支持
Public Sub ForceRestartWheelSupport()
    Debug.Print "=== 强制重启滚轮支持 ==="
    
    ' 先禁用
    Call DisableMouseWheelSupport("强制重启")
    
    ' 等待一下
    Application.Wait Now + TimeValue("00:00:01")
    
    ' 重新启用
    If Not m_ActiveForm Is Nothing Then
        Call EnableMouseWheelSupport(m_ActiveForm.Caption, m_ActiveForm)
    End If
End Sub

' 禁用ListBox内置滚轮行为
#If Win64 Then
Private Sub DisableListBoxInternalScroll(ByVal Hwnd As LongPtr)
#Else
Private Sub DisableListBoxInternalScroll(ByVal Hwnd As Long)
#End If
    On Error Resume Next
    
    Debug.Print "?? 尝试禁用ListBox内置滚轮行为"
    
    ' 获取当前窗口样式
    #If Win64 Then
        Dim currentStyle As LongPtr
        currentStyle = GetWindowLongPtr(Hwnd, GWL_STYLE)
    #Else
        Dim currentStyle As Long
        currentStyle = GetWindowLong(Hwnd, GWL_STYLE)
    #End If
    
    Debug.Print "ListBox当前样式: " & Hex(currentStyle)
    
    ' 移除垂直和水平滚动条样式（这可能会影响内置滚轮）
    #If Win64 Then
        Dim newStyle As LongPtr
    #Else
        Dim newStyle As Long
    #End If
    newStyle = currentStyle And Not WS_VSCROLL And Not WS_HSCROLL
    
    If newStyle <> currentStyle Then
        #If Win64 Then
            SetWindowLongPtr Hwnd, GWL_STYLE, newStyle
        #Else
            SetWindowLong Hwnd, GWL_STYLE, newStyle
        #End If
        Debug.Print "? 已移除ListBox滚动条样式: " & Hex(newStyle)
    Else
        Debug.Print "ListBox样式无需更改"
    End If
End Sub

