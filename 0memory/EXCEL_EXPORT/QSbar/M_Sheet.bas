Attribute VB_Name = "M_Sheet"


'插入“工作表目录”，并建立链接

Sub 工作表目录()
    On Error GoTo GetOut
    Application.ScreenUpdating = False

    Sheets.Add Before:=Worksheets(1)
    Sheets(1).Name = "INDEX"    '命名工作表

    '标题
    With Range("A1:B1")
        .Value = Array("编号", "目录")
        .Font.Bold = True
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.599993896298105
     End With
    
    '从第2个表，第2行开始
    rowNum = 2  '目录起始行
    
    For i = 2 To Worksheets.Count
        If Worksheets(i).Visible = xlSheetVisible Then  '仅处理可见表
            '写入目录行（自动跳过隐藏表）
            Cells(rowNum, 1).Value = rowNum - 1  '连续编号
            Cells(rowNum, 2).Value = "=HYPERLINK(""#'" & Worksheets(i).Name & "'!A1"", """ & Worksheets(i).Name & """)"
            rowNum = rowNum + 1  '行号自增
        End If
    Next
    
    '设置上下、左右都居中对齐
    Columns("A:A").HorizontalAlignment = xlCenter
    Columns("B:B").HorizontalAlignment = xlLeft
    Columns("A:C").AutoFit

    Application.ScreenUpdating = True
GetOut:
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''
'Sub 添加目录()
'    Dim sht As Worksheet, i As Integer
'    On Error Resume Next
'    '只EXCEL2010能用，因为2003会报错
'      If Application.Version = "14.0" Then
'      i = 0
'    ' MsgBox CommandBars("报价工具栏").Controls(1).Caption
'      CommandBars("报价工具栏").Controls("目录").Delete
'      Set myControl = CommandBars("报价工具栏").Controls.Add(msoControlComboBox, , , 1, True)  '产生一个组合框
'      With myControl
'          .Caption = "目录" '组合框控件的标题
'          .Text = ActiveSheet.Name '默认显示文件
'          .OnAction = "添加目录"
'          For Each sht In Sheets  '将工作表表追加到组合框中
'              .AddItem Text:=sht.Name, index:=i + 1
'              .DropDownWidth = 105  '设置给合框宽度
'              .OnAction = "into"  '在给合框回车时执行的宏
'              i = i + 1
'          Next sht
'      End With
'    End If
'End Sub
'Sub into()    '如果选择"刷新目录"则重新建立工具栏否则进入工作表
'On Error Resume Next
'    Sheets(myControl.Text).Select
'End Sub





Sub 汇总全表()
    
    '1）新建一汇总表
    Dim ws As Worksheet
    With ActiveWorkbook
        Set ws = .Sheets.Add(Before:=.Sheets(1))
        ws.Name = "汇总"
    End With

    '2）插入数据
    Dim TitleRow, endRow As Integer   '表头（插入点）
    endRow = Sheets(1).UsedRange.Rows.Count + Sheets(1).UsedRange.Row - 1   '初始化插入表的首尾位置
    TitleRow = endRow + 1
    
    
    
    
    For i = 2 To Sheets.Count
                Sheets(i).UsedRange.Copy
                Sheets(1).Cells(TitleRow + 1, 1).Insert Shift:=xlDown
                '重新设置表头表尾位置。
                endRow = Sheets(1).UsedRange.Rows.Count + Sheets(1).UsedRange.Row - 1
                TitleRow = endRow + 1
                
                
    
        If (BQ表格(i) <> "") Then
         '插入数据
            '如果找不到【，则是〖表头
            If Not (BQ表格(i) Like "【*") Then
                Sheets(1).Cells(TitleRow, 1).Value = "〖" & Sheets(BQ表格(i)).Name & "〗"   '制作表头
                '插入数据
                Sheets(BQ表格(i)).UsedRange.Copy
                Sheets(1).Cells(TitleRow + 1, 1).Insert Shift:=xlDown
        
                '数据分组。重新设置表头表尾位置。
                endRow = Sheets(1).UsedRange.Rows.Count + Sheets(1).UsedRange.Row - 1
                Sheets(1).Range(TitleRow + 1 & ":" & endRow).Rows.Group
                TitleRow = endRow + 2
            Else
                '如果找到【，则直接改下数值
                Sheets(1).Cells(TitleRow, 1).Value = BQ表格(i)
    
                '重新设置表头表尾位置。
                endRow = Sheets(1).UsedRange.Rows.Count + Sheets(1).UsedRange.Row - 1
                TitleRow = endRow + 1
    
            End If
        End If
    Next
    
'    '《调试》
'    For i = 0 To UBound(BQ表格)
'        If (BQ表格(i) <> "") Then
''            Debug.Print BQ表格(i)
'        End If
'    Next
    
End Sub






Sub 表格汇总()

Dim TitleRow, endRow As Integer

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
End With

'新建一汇总表
Sheets.Add Before:=Worksheets(1)

'初始化表尾，表尾位置
endRow = Sheets(1).UsedRange.Rows.Count + Sheets(1).UsedRange.Row - 1
TitleRow = endRow

For i = 2 To Sheets.Count
'    '做表头
'    With Sheets(1).Cells(TitleRow, 1)
'        .value = Sheets(i).Name
'        .HorizontalAlignment = xlLeft
'        With .EntireRow.Font
'            .Bold = True
'            .Name = "微软雅黑"
'            .Size = 11
'        End With
'        With .EntireRow.Interior
'            .ThemeColor = xlThemeColorAccent1
'            .TintAndShade = 0.599993896298105
'        End With
'    End With

    '插入数据
    Sheets(i).UsedRange.Copy
    Sheets(1).Cells(TitleRow + 1, 1).Insert Shift:=xlDown

    '更新表尾，分组，更新表头
    endRow = Sheets(1).UsedRange.Rows.Count + Sheets(1).UsedRange.Row - 1
'    Sheets(1).Range(TitleRow + 1 & ":" & EndRow).Rows.Group
    TitleRow = endRow

Next

'改名，收起表单
Sheets(1).Name = "表单汇总"
Call showL1

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With
    

End Sub










Sub 删除空白行(Optional setRng As Range)

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
        If Application.CountA(rng.Rows(i)) = 0 Then
            rng.Rows(i).EntireRow.Delete
            j = j + 1
        End If
    Next


    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub




Sub UnhideAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub


