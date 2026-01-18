Attribute VB_Name = "M_Format"

'区域第col列的值是符合条件值，like cond
'设置菜单（区域，列，判断条件值，级别）
Sub RowGroup(ByVal rng As Range, col As Integer, ByVal cond As String, lev As Integer)
On Error Resume Next
    If lev < 0 Then lev = 1
    If lev > 6 Then lev = 6

    For i = 1 To rng.Rows.Count
        '判断位置。如果第i列单元格有值 like cond，就按lev定级
        If rng.Cells(i, col).Value Like cond Then
            rng.Rows(i).OutlineLevel = lev
        End If
    Next

End Sub


'根据行等级设置样式

Sub RowStyle(rng As Range)

On Error Resume Next   '防止因为没定义Font属性而报错

'提速
With Application
    .ScreenUpdating = False              '屏幕更新
    .Calculation = xlCalculationManual   '手动计算
    .EnableEvents = False                '事件控制
End With

'设置样式
For i = 1 To rng.Rows.Count
    With rng.Rows(i)
        '1）字体和背景
        If .OutlineLevel = 1 Then       '判断条件是OutlineLevel属性
            .Interior.ThemeColor = xlThemeColorLight2
            .Interior.TintAndShade = -0.249977111117893
            .Font.ThemeColor = xlThemeColorDark1         '深蓝
'                .Font.Size = 11
            .Font.Bold = True
            .Font.Name = "Microsoft YaHei UI"
        End If

        If .OutlineLevel = 2 Then
            .Interior.ThemeColor = xlThemeColorLight2
            .Interior.TintAndShade = 0.799981688894314   '普蓝
'                .Font.Size = 10
            .Font.Bold = True
            .Font.Name = "Microsoft YaHei UI"
        End If

        If .OutlineLevel = 3 Then
            .Interior.ThemeColor = xlThemeColorAccent2
            .Interior.TintAndShade = 0.799981688894314   '淡橙
            .Font.Name = "Microsoft YaHei UI"
'            .Font.Size = 9
'            .Font.Bold = False
            
        End If
        
        If .OutlineLevel = 4 Then
            .Font.Name = "Microsoft YaHei UI"
'                .Font.Size = 9
        End If
        
        '2）边框
'            For j = 1 To rng.Columns.Count
'                With .Columns(j)
'                    '上
'                    If .Borders(xlEdgeTop).LineStyle <> xlDouble And .Borders(xlEdgeTop).Weight <> xlMedium Then
'                        .Borders(xlEdgeTop).Weight = xlHairline
'                    End If
'                    '下
'                    If .Borders(xlEdgeBottom).LineStyle <> xlDouble And .Borders(xlEdgeBottom).Weight <> xlMedium Then
'                        .Borders(xlEdgeBottom).Weight = xlHairline
'                    End If
'                    '左
'                    If .Borders(xlEdgeLeft).LineStyle <> xlDouble And .Borders(xlEdgeLeft).Weight <> xlMedium Then
'                        .Borders(xlEdgeLeft).Weight = xlThin
'                    End If
'                    '右
'                    If .Borders(xlEdgeRight).LineStyle <> xlDouble And .Borders(xlEdgeRight).Weight <> xlMedium Then
'                        .Borders(xlEdgeRight).Weight = xlThin
'                    End If
'                End With
'            Next
        
    End With
Next

'提速end
With Application
    .ScreenUpdating = True           '屏幕更新
    .EnableEvents = True             '事件控制
End With


End Sub

'清除所有样式
Sub ClearStyle(rng As Range)

    '清楚填充色，边框，字体色（分组不能清除，否则无法配合RowStyle使用
    With rng
        .Interior.pattern = xlNone
'        .Font.ColorIndex = xlAutomatic
'        .Font.Bold = False
'        .Borders(xlDiagonalDown).LineStyle = xlNone
'        .Borders(xlDiagonalUp).LineStyle = xlNone
'        .Borders(xlEdgeLeft).LineStyle = xlNone
'        .Borders(xlEdgeTop).LineStyle = xlNone
'        .Borders(xlEdgeBottom).LineStyle = xlNone
'        .Borders(xlEdgeRight).LineStyle = xlNone
'        .Borders(xlInsideVertical).LineStyle = xlNone
'        .Borders(xlInsideHorizontal).LineStyle = xlNone


        
    End With

End Sub



'判断是否设置单元格样式：字体颜色或填充色
Function IsStyled(rng As Range) As Boolean
    On Error Resume Next   '防止.Font没定义而报错
    Dim FontColor, InteriorColor As Integer
    Dim cell1 As Range
    
    '初始化
    FontColor = -1000       '随便取的不知道什么色
    InteriorColor = -4142   '空白色
    
    '先假设都未设置，返回False
    IsStyled = False
    
    '只要有一个单元格设置了字体颜色或填充色，返回True
    For Each cell1 In rng
    
        FontColor = cell1.Font.ThemeColor
        InteriorColor = cell1.Interior.ThemeColor
    
        If FontColor <> -1000 Or InteriorColor <> -4142 Then
            IsStyled = True
            Exit For
        End If
    
    Next
    
    
    '检查最后一个的样式
    Debug.Print cell1.Address & "：FontColor=" & FontColor & "   InteriorColor=" & InteriorColor
    

End Function

'按条件分组
Sub 设置分级()

    '1)参数设置
    Dim rng As Range
    Set rng = Intersect(Selection, ActiveSheet.UsedRange)  '选区
    
    Dim col As Integer
    col = 2   '判断列
    
    
    '2)条件分级
    For i = 1 To rng.Rows.Count
        lvl = 1
        V = rng.Cells(i, col).Value
        
        If V Like "*【*" Then
            rng.Rows(i).OutlineLevel = 1
            
        ElseIf V Like "*《*" Then
            rng.Rows(i).OutlineLevel = 2
            
        ElseIf V Like "*{*" Then
            rng.Rows(i).OutlineLevel = 3
        ElseIf V Like "*｛*" Then
            rng.Rows(i).OutlineLevel = 3
        Else
            rng.Rows(i).OutlineLevel = 4
        End If
    Next
    
    
End Sub




Sub 设置分级样式()

    Dim rng As Range
    
    '默认区域是ActiveSheet.UsedRange
'    Set rng = Intersect(Selection, ActiveSheet.UsedRange)
    Set rng = Selection

    '设置样式
    Call ClearStyle(rng)   '先要清除所有样式
    Call RowStyle(rng)     '设置样式

End Sub



'EXCEL自带组合功能不能实现多区域分组，该sub解决这一问题
'多区域可见行分组

Sub 多选区组合()

    With Selection.SpecialCells(xlCellTypeVisible)
    '要访问多区域数据必须先区域循环再行间循环，而不能直接行间循环
        For i = 1 To .Areas.Count
            For j = 1 To .Areas(i).Rows.Count
                .Areas(i).Rows(j).Group

            Next
        Next
    End With
    
End Sub

Sub 多选区取消组合()
On Error Resume Next
    '多区域可见行分组
    With Selection.SpecialCells(xlCellTypeVisible)
    '要访问多区域数据必须先区域循环再行间循环，而不能直接行间循环
        For i = 1 To .Areas.Count
            For j = 1 To .Areas(i).Rows.Count
                .Areas(i).Rows(j).Ungroup
            Next
        Next
    End With
End Sub


Sub 只选可见单元格()
    Selection.SpecialCells(xlCellTypeVisible).Select
End Sub
