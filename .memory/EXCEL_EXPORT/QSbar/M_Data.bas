'作用：1、清除超链接  2、可跨单元格粘死  3、更新单元格数据 ，让带公式的数值不会以字符串的形式展示公式，而是常规使用公式。 4、字符串转数值
Sub 数值化(Optional myrng As Range)

    On Error Resume Next ' 错误处理

    ' 1. 确定处理范围
    If myrng Is Nothing Then
        Set myrng = Intersect(Selection, ActiveSheet.UsedRange)
        If myrng Is Nothing Then Exit Sub ' 如果没有选择任何单元格，则退出
    End If

    ' 2. 优化处理方式，提高效率
    Dim dataArr As Variant ' 使用数组存储单元格数据
    dataArr = myrng.Value

    Dim i As Long, j As Long
    Dim val As Variant

    For i = 1 To UBound(dataArr, 1)
        For j = 1 To UBound(dataArr, 2)
            val = dataArr(i, j)

            ' 3. 清除超链接（一次性清除整个区域的超链接）
            myrng.Hyperlinks.Delete

            ' 4. 公式转数值、百分号处理、字符串转数值
            If IsNumeric(val) Then ' 如果是数值，则无需转换
                dataArr(i, j) = val
            ElseIf IsError(val) Then ' 如果是错误值，则保留
                dataArr(i, j) = val
            ElseIf Right(val, 1) = "%" Then ' 百分号处理
                dataArr(i, j) = CSng(Left(val, Len(val) - 1) / 100)
            ElseIf IsDate(val) Then '日期格式转换
                dataArr(i, j) = CDate(val)
            ElseIf IsEmpty(val) Then '如果单元格为空，则跳过
                dataArr(i, j) = val
            ElseIf Left(val, 1) = "=" Then ' 如果是公式，则保持公式格式
                dataArr(i, j) = val
            ElseIf IsNumeric(val) = False And IsError(val) = False And IsDate(val) = False And IsEmpty(val) = False And Left(val, 1) <> "=" Then
                ' 如果是字符串形式的数字，则转换为数值格式
                If IsNumeric(CDbl(val)) Then ' 检查是否可以转换为数值
                    dataArr(i, j) = CDbl(val)
                Else
                    dataArr(i, j) = val
                End If
            Else
                dataArr(i, j) = val
            End If


        Next j
    Next i

    myrng.Value = dataArr ' 将处理后的数据写回单元格

    ' 5. 设置数字格式（一次性设置整个区域的格式）
    With myrng
        .ShrinkToFit = True
        .NumberFormatLocal = " #,##0.00_ ;[红色] -#,##0.00_ ;_ """"??_ ;@"
    End With
    Debug.Print myrng.Address
    Set myrng = Nothing ' 释放对象变量

End Sub

Public Sub 强制刷新()
    Dim cell As Range
    ' 如果未指定区域，则默认使用当前工作表的已使用区域
    Set rng = ActiveSheet.UsedRange

    For Each cell In rng
        ' 如果单元格有公式，则强制刷新公式

            cell.formula = cell.formula

    Next cell
End Sub



Sub 会计格式2位小数()
On Error Resume Next
    With Intersect(Selection, ActiveSheet.UsedRange)
        .ShrinkToFit = True
        .NumberFormatLocal = " #,##0.00_ ;[红色] -#,##0.00_ ;_ """"??_ ;@"
    End With

End Sub
Sub 会计格式3位小数()
On Error Resume Next
    With Intersect(Selection, ActiveSheet.UsedRange)
        .ShrinkToFit = True
        .NumberFormatLocal = " #,##0.000_ ;[红色] -#,##0.000_ ;_ """"??_ ;@"
    End With

End Sub
Sub 会计格式整数()
On Error Resume Next
    With Intersect(Selection, ActiveSheet.UsedRange)
        .ShrinkToFit = True
        .NumberFormatLocal = " #,##0_ ;[红色] -#,##0_ ;_ """"??_ ;@"
    End With

End Sub

Sub 文本换行格式()
On Error Resume Next
    With Selection
        .NumberFormatLocal = "@"
        .WrapText = True
    End With

End Sub

Sub 亿万格式()
On Error Resume Next
    With Selection
        .ShrinkToFit = True
        .NumberFormatLocal = "[<=-100000000]-0!.00,,""亿"";[>=100000000]0!.00,,""亿"";0!.0,""万"""
    End With

End Sub


Sub 粘死所有外表链接()

    On Error Resume Next
    
    If ActiveSheet.ProtectContents Then MsgBox "工作表已保护,本程序拒绝执行！", 64, "提示": Exit Sub
    Dim cell As Range, FirstAddress As String, sht As Worksheet
    Dim s As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Worksheets    '遍历所有工作表
        With sht.UsedRange  '仅对已用数据区域进行操作
            '查找带有“=*'!”的单元格，该单元格引用了外部工作表数据
            Set cell = .Find("=*[*]*", LookIn:=xlFormulas, SearchOrder:=xlByRows, LookAt:=xlPart, MatchCase:=True)
            FirstAddress = cell.Address    '记录第一个查到的地址
            Do
                cell.Value = cell.Value   '将找到的单元格公式转换成值
                Set cell = .FindNext(cell)
            Loop Until cell Is Nothing Or cell.Address = FirstAddress
        End With
    Next sht    '进入下一个工作表
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


'将引用其他格的单元格地址锁定

Sub 地址锁定()

On Error Resume Next:
  
  Application.ScreenUpdating = False
  Dim myrng, rng2 As Range           '选定区域
  Dim ss, ts As String
  
  
  For Each myrng In Selection
       ss = myrng.formula
       ss = Right(ss, Len(ss) - 1)
       Set rng2 = Range(ss)
      
       If rng2.Address <> "" Then
          
           myrng.formula = "='" & rng2.Worksheet.Name & "'!" & rng2.Address

       End If
       
       Set rng2 = Nothing
  Next
  
  Application.ScreenUpdating = True
ISOVER:
End Sub

'清除整个工作簿中全部超链接

Sub 删除超链接()
    Dim sht As Worksheet
    Application.ScreenUpdating = False
    For Each sht In ActiveWorkbook.Worksheets  '遍历当前工作簿中所有工作表
        sht.UsedRange.Hyperlinks.Delete  '删除所有超级链接
    Next
    Application.ScreenUpdating = True
End Sub


'用对话框返回文件的绝对路径

Sub 文件名目录()
     'Declare a variable as a FileDialog object.
    Dim fd As FileDialog
    Dim myCell As Range

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    'Declare a variable to contain the path
    'of each selected item. Even though the path is aString,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd

        'Allow the user to select multiple files.
        .AllowMultiSelect = True

        'Use the Show method to display the File Picker dialog box and return the user's action.
        'If the user presses the button...
        If .Show = -1 Then

            Set myCell = Selection(1, 1)
            
            For Each vrtSelectedItem In .selectedItems
                myCell.Value = vrtSelectedItem
                Set myCell = myCell.Offset(1, 0)
                'vrtSelectedItem is aString that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example displays the path in a message box.
               ' MsgBox "Selected item's path: " & vrtSelectedItem

            Next
        'If the user presses Cancel...
        Else
        End If
    End With

    'Set the object variable to Nothing.
    Set fd = Nothing

End Sub
