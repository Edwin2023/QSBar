Attribute VB_Name = "M_Export"
Sub OUT本表()
    Dim file_name As String
    Dim save_path As String
    Dim new_wb As Workbook
    Dim original_calc As XlCalculation
    
    '保留原始计算模式
    original_calc = Application.Calculation
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    On Error GoTo ErrorHandler
    
    '--- 文件名处理 ---
    file_name = "(OUT" & Format(Date, "yyyy-mm-dd") & ")" & ActiveSheet.Name
    '替换非法字符
    file_name = Replace(file_name, ":", "-")
    file_name = Replace(file_name, "\", "-")
    file_name = Replace(file_name, "/", "-")
    file_name = Left(file_name, 50) & ".xlsx"  '限制文件名长度
    
    '--- 路径处理 ---
    save_path = Environ("USERPROFILE") & "\Desktop\"
    '确保路径结尾有反斜杠
    If Right(save_path, 1) <> "\" Then save_path = save_path & "\"
    
    '--- 复制工作表 ---
    ActiveSheet.Copy
    Set new_wb = ActiveWorkbook
    
    '--- 粘贴法保留格式 ---
    With new_wb.Sheets(1).UsedRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
    End With
    Application.CutCopyMode = False  '清除剪贴板
    
    '--- 文件覆盖处理 ---
    DoEvents
    '修正语法：删除多余右括号
    If Dir(save_path & file_name) <> "" Then Kill save_path & file_name
    
    '--- 安全保存 ---
    new_wb.SaveAs fileName:=save_path & file_name, _
                FileFormat:=xlOpenXMLWorkbook, _
                AccessMode:=xlExclusive, _
                ConflictResolution:=xlLocalSessionChanges
    
CleanUp:
    '恢复原始设置
    With Application
        .ScreenUpdating = True
        .Calculation = original_calc
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 1004
            MsgBox "无法保存文件，可能原因：" & vbCrLf & _
                   "1. 文件已被其他程序打开" & vbCrLf & _
                   "2. 路径不存在或无写入权限", vbCritical
        Case Else
            MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume CleanUp
End Sub


Sub 复制新表()
On Error Resume Next:

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
End With


Dim k, i As Integer

'获取文件名“（OUT日期）文件名”*必须放在Sheets.Copy前面
Dim markName, temp As String
markName = "(OUT" & Replace(Date, "/", "-") & ")"    '标记，加入日期，注意日期原格式2017/6/11会报错，要换/符号
temp = ActiveWorkbook.Path & "\" & markName & ActiveWorkbook.Name  '记下表格名称和路径
For i = 1 To 10     '去后缀（从右往左）
    If Left(Right(temp, i), 1) = "." Then Exit For
Next
temp = Left(temp, Len(temp) - i)


'1复制并数据处理
Sheets.Copy

'2保存
ActiveWorkbook.SaveAs fileName:=temp, FileFormat:=xlOpenXMLWorkbook


With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With


End Sub


Sub 全表粘死()
    On Error Resume Next

    Dim k As Integer
    Dim ws As Worksheet

    For k = 1 To Worksheets.Count
        Set ws = ActiveWorkbook.Worksheets(k)

        ' 清空剪贴板，防止粘贴错误
        Application.CutCopyMode = False

        ' 仅处理数据区域，避免使用整个UsedRange
        If Not ws.UsedRange Is Nothing Then
            ws.UsedRange.Copy

            ' 粘贴为数值
            ws.UsedRange.PasteSpecial Paste:=xlPasteValues

            ' 清理剪贴板
            Application.CutCopyMode = False
            DoEvents ' 释放内存
        End If
    Next k

    ' 最后清理剪贴板
    Application.CutCopyMode = False
    DoEvents
End Sub



Sub 删除所有表格隐藏列()

    Dim Counter As Integer
    Dim sht As Worksheet

    For i = 1 To Sheets.Count
        Set sht = Sheets(i)
        Counter = 0
        Do While Counter < sht.UsedRange.Columns.Count
            Counter = Counter + 1

            If sht.UsedRange.Columns(Counter).EntireColumn.Hidden = True Then
                sht.UsedRange.Columns(Counter).EntireColumn.Delete Shift:=xlToLeft
                Counter = Counter - 1
            End If
        Loop
    Next
    
End Sub


Sub 取消所有自动筛选()
    On Error Resume Next

    Dim sht As Worksheet
    For Each sht In Sheets
        ' 检查工作表是否有自动筛选
        If sht.AutoFilterMode Then
            sht.AutoFilterMode = False ' 关闭自动筛选
        End If
    Next sht

    On Error GoTo 0 ' 关闭错误忽略
End Sub




Sub 清除错误和外表名称()
    Dim NME As Name


    For Each NME In ActiveWorkbook.Names
        If (NME.Value Like "*REF*") Or (NME.Value Like "*:\*") Then
            NME.Delete
        End If
    Next

    For Each NME In ActiveWorkbook.Names
        If (NME.Value Like "*\\*") Or (NME.Value Like "*#N/A*") Then
            NME.Delete
        End If
    Next

'
'    For Each NME In Names
'        Debug.Print NME.value
'    Next

End Sub

'对比打印区域和UsedRange区域，交集保留，差集删除


Sub 删除打印区域以外的值()
    
    On Error Resume Next
    Dim sht As Worksheet
    Dim rng, OUTRNG, temp As Range      'RNG是检索值，OUTRN储存非交集区域，Temp储存最后筛选出来的range
    Dim PrintName As Name

    For i = 1 To ActiveWorkbook.Sheets.Count
        '1、通过打印名称获取打印区域，存到PrintName
        Set PrintName = Nothing
        For j = 1 To Sheets(i).Names.Count
            If Sheets(i).Names(j).Name Like "*Print_Area*" Then
                Set PrintName = Sheets(i).Names(j)
            End If
        Next
        
        '没有打印区域的就以Sheets(i).UsedRage作为打印区域名称
        If PrintName Is Nothing Then
            Sheets(i).Names.Add Name:="Print_Temp", RefersTo:=Sheets(i).UsedRange
            Set PrintName = Sheets(i).Names("Print_Temp")
        End If
        
'        Debug.Print "Name:" & PrintName.Name & "     Address:" & PrintName.RefersToRange.Address
        
        
        '2、比较使用区和打印区，保留交集，删除差集
        
           '2.1删除列
            Set OUTRNG = Nothing
           '求差集。没有交集（即差集），则Union起来
            For Each rng In Sheets(i).UsedRange.Columns
                If Intersect(rng, PrintName.RefersToRange) Is Nothing Then
                    If OUTRNG Is Nothing Then
                        Set OUTRNG = rng
                    Else
                        Set OUTRNG = Application.Union(OUTRNG, rng)
                    End If
                End If
            Next
            '删除差集
            If OUTRNG Is Nothing Then
'                Debug.Print "无多余值在非打印区！"
            Else
                OUTRNG.EntireColumn.Delete
            End If


           '2.2删除行
            Set OUTRNG = Nothing
           '求差集
            For Each rng In Sheets(i).UsedRange.Rows
                If Intersect(rng, PrintName.RefersToRange) Is Nothing Then
                    If OUTRNG Is Nothing Then
                        Set OUTRNG = rng
                    Else
                        Set OUTRNG = Application.Union(OUTRNG, rng)
                    End If
                End If
            Next
            '删除差集
            If OUTRNG Is Nothing Then
                Debug.Print "无多余值在非打印区！"
            Else
                OUTRNG.EntireRow.Delete
            End If
            '差集对象清空
            Set OUTRNG = Nothing
        
    Next

End Sub


Sub OUT制作()

'On Error Resume Next:

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .DisplayAlerts = False
End With

Dim k, i As Integer


'获取文件名“（OUT日期）文件名”*必须放在Sheets.Copy前面
Dim markName, temp As String
markName = "(OUT" & Replace(Date, "/", "-") & ")"    '标记，加入日期，注意日期原格式2017/6/11会报错，要换/符号
temp = ActiveWorkbook.Path & "\" & markName & ActiveWorkbook.Name  '记下表格名称和路径

'去文件名后缀（从右往左）
For i = 1 To 10
    If Left(Right(temp, i), 1) = "." Then Exit For
Next
temp = Left(temp, Len(temp) - i)


'1复制并数据处理

Sheets.Copy

Call 展开所有行
Call 取消所有自动筛选
Call 清除错误和外表名称
Call 全表粘死


Call 隐藏非1级列
Call 删除所有表格隐藏列
Call 删除打印区域以外的值

''删除自定义名称,非打印区域名称不删除。（执行这句条件格式就不好用了）
'For Each item In ActiveWorkbook.Names
'    If Not item.Name Like "*Print*" Then item.Delete
'Next


'2保存

'方案一：存原目录
ActiveWorkbook.SaveAs fileName:=temp, FileFormat:=xlOpenXMLWorkbook

'方案二：打开另存为对话框
'Application.Dialogs(5).Show 另存为对话框


With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .DisplayAlerts = True
End With
    
End Sub

Sub OUT内部全数据()
'在当前文件夹生成OUT文件，

On Error Resume Next:

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .DisplayAlerts = False
End With

Dim k, i As Integer


'获取文件名“（OUT日期）文件名”*必须放在Sheets.Copy前面
Dim markName, temp As String
markName = "(OUT内部全数据" & Replace(Date, "/", "-") & ")"    '标记，加入日期，注意日期原格式2017/6/11会报错，要换/符号
temp = ActiveWorkbook.Path & "\" & markName & ActiveWorkbook.Name  '记下表格名称和路径

'去文件名后缀（从右往左）
For i = 1 To 10
    If Left(Right(temp, i), 1) = "." Then Exit For
Next
temp = Left(temp, Len(temp) - i)


'1复制并数据处理
Call 展开所有行
Sheets.Copy
Call 全表粘死
Call 取消所有自动筛选
Call 清除错误和外表名称

Call 删除打印区域以外的值

''删除自定义名称,非打印区域名称不删除。（执行这句条件格式就不好用了）
'For Each item In ActiveWorkbook.Names
'    If Not item.Name Like "*Print*" Then item.Delete
'Next


'2保存

'方案一：存原目录
ActiveWorkbook.SaveAs fileName:=temp, FileFormat:=xlOpenXMLWorkbook

'方案二：打开另存为对话框
'Application.Dialogs(5).Show 另存为对话框


With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .DisplayAlerts = True
End With

End Sub


'Sub 隐藏非1级列()
''    On Error Resume Next
'
'    For k = 1 To Worksheets.Count
'        With ActiveWorkbook.Worksheets(k)
'            .Outline.ShowLevels ColumnLevels:=1
'        End With
'    Next
'
''        On Error GoTo 0 ' 恢复默认错误处理
'
'End Sub


Sub 隐藏非1级列()
    On Error Resume Next ' 忽略错误，继续执行

    Dim k As Integer
    Dim ws As Worksheet
    
    For k = 1 To Worksheets.Count
        Set ws = ActiveWorkbook.Worksheets(k)
        
        ' 输出当前的工作表名称和k值，以便调试
'        Debug.Print "当前k值: " & k & " 工作表名称: " & ws.Name
        
        ' 尝试执行隐藏非1级列的操作
        ws.Outline.ShowLevels ColumnLevels:=1

        ' 如果出错，输出错误信息
'        If Err.Number <> 0 Then
'            Debug.Print "错误发生在工作表: " & ws.Name & " 错误号: " & Err.Number & " 错误描述: " & Err.Description
'            Err.Clear ' 清除错误
'        End If
    Next k

    On Error GoTo 0 ' 恢复默认错误处理
End Sub

Sub 展开所有折叠列()
    On Error Resume Next

    For k = 1 To Worksheets.Count
        With ActiveWorkbook.Worksheets(k)
            .Outline.ShowLevels ColumnLevels:=3
        End With
    Next
    
End Sub


Sub 展开所有行()
    On Error Resume Next ' 忽略错误，继续执行

    Dim k As Integer

    ' 第二轮：展开所有行
    For k = 1 To Worksheets.Count
        With ActiveWorkbook.Worksheets(k)
            ' 如果该工作表有大纲结构，展开所有行
            .Outline.ShowLevels RowLevels:=4

            ' 如果出错，打印错误信息
            If Err.Number <> 0 Then
                Debug.Print "展开行出错，工作表: " & Worksheets(k).Name & " 错误号: " & Err.Number & " 错误描述: " & Err.Description
                Err.Clear ' 清除错误，继续执行
            End If
        End With
    Next k

'    On Error GoTo 0 ' 恢复默认错误处理
End Sub

Sub 展开所有行列()
    On Error Resume Next ' 忽略错误，继续执行

    Dim k As Integer
    ' 第一轮：展开所有列
    For k = 1 To Worksheets.Count
        With ActiveWorkbook.Worksheets(k)
            ' 尝试展开列，如果没有大纲结构则跳过
            .Outline.ShowLevels ColumnLevels:=4
            If Err.Number <> 0 Then
                Debug.Print "展开列出错，工作表: " & Worksheets(k).Name & " 错误号: " & Err.Number & " 错误描述: " & Err.Description
                Err.Clear ' 清除错误，继续执行
            End If
        End With
    Next k

    ' 第二轮：展开所有行
    For k = 1 To Worksheets.Count
        With ActiveWorkbook.Worksheets(k)
            ' 如果该工作表有大纲结构，展开所有行
            .Outline.ShowLevels RowLevels:=4

            ' 如果出错，打印错误信息
            If Err.Number <> 0 Then
                Debug.Print "展开行出错，工作表: " & Worksheets(k).Name & " 错误号: " & Err.Number & " 错误描述: " & Err.Description
                Err.Clear ' 清除错误，继续执行
            End If
        End With
    Next k

'    On Error GoTo 0 ' 恢复默认错误处理
End Sub



Sub showL1()
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
End Sub

Sub showL2()
    ActiveSheet.Outline.ShowLevels RowLevels:=2
End Sub
Sub showL3()
    ActiveSheet.Outline.ShowLevels RowLevels:=3
End Sub
Sub showL4()
    ActiveSheet.Outline.ShowLevels RowLevels:=4
End Sub




'当前工作簿另存，标记+文件名.xlsx
Sub 另存当前工作簿()

Dim Mark, FileTitle As String

'标记，加入日期。注意日期原格式2017/6/11会报错，要换/符号
Mark = "(O" & Replace(Date, "/", "-") & ")"
'路径 + 标记 + 当前文件名
FileTitle = ActiveWorkbook.Path & "\" & Mark & ActiveWorkbook.Name

'去后缀。这一步必须有，因为选格式保存时会自带后缀。
'由于instr只能从左往右，用循环的方式从右往从找10个值，直到发现"."退出，取左边的
For i = 1 To 10
    If Left(Right(FileTitle, i), 1) = "." Then Exit For
Next
FileTitle = Left(FileTitle, Len(FileTitle) - i)

'保存，xlOpenXMLWorkbook是xlsx格式
ActiveWorkbook.SaveAs fileName:=FileTitle, FileFormat:=xlOpenXMLWorkbook
    
End Sub



Sub 删除名称()

'On Error Resume Next
'
    For Each Item In ActiveWorkbook.Names
        If Not Item.Name Like "*Print*" Then Item.Delete 'Debug.Print item.Name 'item.Delete
    Next
'    While ActiveWorkbook.Names.Count > 3
'        ActiveWorkbook.Names(3).Delete
'    Wend
'Debug.Print ActiveWorkbook.Names(3).Name

End Sub
