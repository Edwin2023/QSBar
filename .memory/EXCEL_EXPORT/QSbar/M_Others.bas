Attribute VB_Name = "M_Others"

Sub 广联达处理()
On Error Resume Next
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
End With
    

    Call 改表格名("- ")
    Call 改表格名("-")
    Call 改表格名("_")
    Call 删除所有表格隐藏列
    For Each sht In Sheets
        Call 数值化(sht.UsedRange)
    Next


With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub


'修改表格名字，清除掉replaceChar之前的字符

Sub 改表格名(replaceChar As String)
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    Dim P As Integer
     
       '如果没有指定字符，默认-
       If replaceChar = "" Then replaceChar = "-"
       
       For Each sht In Sheets
           P = InStr(1, sht.Name, replaceChar)
           If P <> 0 Then
            sht.Name = Mid(sht.Name, P + 1, Len(sht.Name))
           End If
       Next
        
    Application.ScreenUpdating = True

End Sub


'修改表格名字，清除掉replaceChar之前的字符

Sub 改表名(old_word As String, new_word As String)
    
    Application.ScreenUpdating = False

    For Each sht In Sheets
        sht.Name = Replace(sht.Name, old_word, new_word)
    Next
     
    Application.ScreenUpdating = True

End Sub

Sub 广联达表名()
'一、改表名
    Call 改表名("- ", "")
    Call 改表名("-", "")
    Call 改表名("_", "")
    Call 改表名("(", "")
    Call 改表名(")", "")
    Call 改表名("BQ Quantity", "")
    Call 改表名("DIQS", "")
End Sub



'选择区域数据改成引用区相同数据的地址

Sub lockaddress()
    
    Dim rng1, rng2 As Range
    Dim cell1, cell2 As Range
    Dim addressValue As String
    
    Set rng1 = Intersect(Selection, Selection.Worksheet.UsedRange)      '选区
    Set rng2 = Application.InputBox("指定引用区:", Type:=8)             '引用区
    Set rng2 = Intersect(rng2, rng2.Worksheet.UsedRange)
    
    For Each cell1 In rng1
        For Each cell2 In rng2
            If cell1.Value = cell2.Value And cell1.Value <> "" Then
                addressValue = "=" & cell2.Parent.Name & "!" & cell2.Address
                cell1.formula = addressValue
                cell1.NumberFormatLocal = "G/通用格式"
'                Debug.Print cell1.value & " at " & addressValue
            End If
        Next
    Next


End Sub



Sub 向下复制()
'
Dim rng As Range
Dim i, num As Integer
 
Application.ScreenUpdating = False
 
On Error GoTo myErr

num = InputBox("请在选框中输入要复制的数量", "请输入复制数量")

If IsNumeric(num) Then
Set rng = Selection
    For i = 1 To num
       '复制插入
        rng.Copy
        rng.Offset(rng.Rows.Count * i).Insert Shift:=xlDown
       '增加打印水平分隔线，如果有打印区域
        If ActiveSheet.PageSetup.PrintArea <> "" Then
          ActiveSheet.HPageBreaks.Add Before:=rng.Offset(rng.Rows.Count * i)
        End If
    Next
End If

Application.ScreenUpdating = True
myErr:
End Sub



'是否有分页符
Function HasPageBreak(ByVal rng As Range) As Boolean
On Error GoTo myErr:
    If Not IsMissing(rng) Then  '如果有参数
        If TypeName(rng) = "Range" Then   '如果是Range
        
            If rng.Rows(1).PageBreak <> xlPageBreakNone Then
                HasPageBreak = True
                Exit Function
            Else
                HasPageBreak = False
                Exit Function
            End If
            
        End If
    End If
myErr:
End Function

Sub 插入分页符()

    Dim Target, Targets As Range
    
    '重置分页符
'    ActiveSheet.ResetAllPageBreaks

    '判断区域C列
    Set Targets = Intersect(Range("C:C"), ActiveSheet.UsedRange)
    
    For Each Target In Targets
        '判断条件包含"*To Collection*","*TO COLLECTION*"
        If Target.Value Like "*To Collection*" Or Target.Value Like "*TO COLLECTION*" Then
            Set Target = Target.Offset(2, 0).EntireRow
            ActiveSheet.HPageBreaks.Add Before:=Target
        End If
    Next

End Sub



'分页符位置修改

Sub 分页符加边框()
    
Dim pgbr, TargetRng As Range

'所有横向分页符
For Each pgbr In ActiveSheet.HPageBreaks
    '所有分页符位置按打印区域调整宽度
    Set TargetRng = pgbr.Location.Resize(1, ActiveSheet.Names("Print_Area").RefersToRange.Columns.Count)
    '上边框
    With TargetRng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
'        .ColorIndex = xlAutomatic
        .Weight = xlMedium
        .ThemeColor = 4
        .TintAndShade = -0.249946592608417
    End With
    '下边框（同一位置，保险起见）
    With TargetRng.Offset(-1, 0).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
'        .ColorIndex = xlAutomatic
        .Weight = xlMedium
        .ThemeColor = 0
        .TintAndShade = -0.249946592608417
    End With
Next

End Sub





'【选区】第k列有对应值，删除对应行

Sub 删除第k列特定值的行(k As Integer, Criteria As String, Optional setRng As Range)

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Dim rng As Range
    If setRng Is Nothing Then
        Set rng = Intersect(Selection, ActiveSheet.UsedRange)
    Else
        Set rng = setRng
    End If
        
    Dim lastRow, i, j As Long
    lastRow = rng.Rows.Count
    j = 0
    
    
    For i = lastRow To 1 Step -1
        If rng.Cells(i, k).Value Like Criteria Then
            rng.Rows(i).EntireRow.Delete
            j = j + 1
        End If
    Next

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub
