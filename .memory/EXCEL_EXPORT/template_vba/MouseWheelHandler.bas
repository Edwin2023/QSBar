Attribute VB_Name = "MouseWheelHandler"
' 安全存根 - 原全局钩子/窗口子类化代码已移除
' 此模块保留空壳以确保向后兼容，不再实现任何 Win32 API 功能

Option Explicit

Public Function IsMouseWheelEnabled() As Boolean
    IsMouseWheelEnabled = False
End Function

Public Function isActive() As Boolean
    isActive = False
End Function

Public Function isWheelCaptured() As Boolean
    isWheelCaptured = False
End Function

Public Sub SetActiveListBox(ByRef listBox As Object)
    ' 空实现
End Sub

Public Sub DisableMouseWheelSupport(Optional ByVal caller As String = "")
    ' 空实现
End Sub

Public Sub PrintInternalState()
    Debug.Print "MouseWheelHandler: 已禁用 (安全模式)"
End Sub

Public Sub ReinitializeMouseHook()
    Debug.Print "MouseWheelHandler: 不执行任何操作 (安全模式)"
End Sub
