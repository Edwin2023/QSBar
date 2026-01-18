Attribute VB_Name = "M_Photo"

Sub 调整图片()  '忽略图表、自选图形、窗体控件，只选择剪贴画和图片
Dim times As Integer
    On Error GoTo Over
            
    times = Mid(CommandBars("报价工具栏").Controls(5).Caption, 3)
    Selection.Placement = xlMoveAndSize  '随单元格大小高度变化
    For Each obj In Selection.ShapeRange
        obj.Top = obj.TopLeftCell.Top + 4 '设置图片上边距
        obj.Left = obj.TopLeftCell.Left + 4 '设置左边距
        obj.LockAspectRatio = msoFalse  '调高度前锁定比例
        obj.Height = obj.TopLeftCell.Height * times - 8 '设置高，宽
        obj.width = obj.TopLeftCell.width - 8 '
        'obj.LockAspectRatio = msoFalse '调宽度前解除锁定比例
        If obj.width < 10 Then
            obj.width = obj.TopLeftCell.width * 2 - 8
        End If
    Next
Over:
End Sub


Sub 全选图片()  '忽略图表、自选图形、窗体控件，只选择剪贴画和图片
    On Error GoTo Over
    If ActiveSheet.ProtectContents Then MsgBox "工作表已保护,本程序拒绝执行！", 64, "提示": Exit Sub

    ActiveSheet.Shapes.SelectAll
    
Over:
End Sub



