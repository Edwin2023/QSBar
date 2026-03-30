Public xlApplication As myEvent


Sub Auto_Open()

    '添加自动事件功能
    Set xlApplication = New myEvent
    Set xlApplication.xlApp = Application
    
    Call 添加报价工具栏
 

    '添加快捷建
    Application.OnKey "^7", "文本换行格式"
    Application.OnKey "^8", "会计格式整数"
    Application.OnKey "^9", "会计格式2位小数"
    Application.OnKey "^0", "会计格式3位小数"
    Application.OnKey "^6", "亿万格式"

    
    Application.OnKey "^1", "showL1"
    Application.OnKey "^2", "showL2"
    Application.OnKey "^3", "showL3"
    Application.OnKey "^4", "showL4"
    Application.OnKey "^5", "只选可见单元格"
    
 
    Application.OnKey "^%1", "展开透视表"
    Application.OnKey "^%2", "折叠透视表"
 
    'Call 检测用户权限
    
End Sub

Sub 添加报价工具栏()
    '生成工具栏
     On Error Resume Next
     
     Application.CommandBars("报价工具栏").Delete
     Application.CommandBars("QS工具栏").Delete
     

     With Application.CommandBars.Add("报价工具栏", msoBarTop, , False)
         With .Controls.Add
             .FaceId = 320
             .OnAction = "BATCH_MODIFY"
             .Caption = "批量修改"
             .Style = msoButtonIcon
         End With
         With .Controls.Add
             .FaceId = 370
             .OnAction = "数值化"
             .Caption = "数值化"
             .Style = msoButtonIcon
         End With
         With .Controls.Add
             .FaceId = 225
             .OnAction = "地址锁定"
             .Caption = "地址锁定"
             .Style = msoButtonIcon
         End With
         With .Controls.Add(msoControlPopup)        'Application.CommandBars("报价工具栏").Controls(5)
             .FaceId = 218
             .Caption = "图X1"
             .OnAction = "调整图片"
             .TooltipText = "图片按单元格高度放大X倍"
             .Style = msoButtonIcon
                With .Controls.Add
                   .Caption = "X1"
                   .OnAction = "改图片调整名称(1)"
                   .Style = msoButtonCaption
                   .DescriptionText = "放大1倍"
                End With
                With .Controls.Add
                   .Caption = "X5"
                   .OnAction = "改图片调整名称(5)"
                   .Style = msoButtonCaption
                   .DescriptionText = "放大5倍"
                End With
                With .Controls.Add
                   .Caption = "X7"
                   .OnAction = "改图片调整名称(7)"
                   .Style = msoButtonCaption
                   .DescriptionText = "放大7倍"
                End With
                With .Controls.Add
                   .Caption = "X10"
                   .OnAction = "改图片调整名称(10)"
                   .Style = msoButtonCaption
                   .DescriptionText = "放大10倍"
                End With
                With .Controls.Add
                   .Caption = "选择全部图片"
                   .OnAction = "全选图片"
                   .Style = msoButtonCaption
                   .DescriptionText = "选择全部图片"
                End With
         End With
         With .Controls.Add
             .FaceId = 70
             .OnAction = "会计格式整数"
             .Caption = "会计格式整数"
             .Style = msoButtonIcon
         End With
         With .Controls.Add
             .FaceId = 72
             .OnAction = "会计格式2位小数"
             .Caption = "会计格式2位小数"
             .Style = msoButtonIcon
         End With
         With .Controls.Add
             .FaceId = 73
             .OnAction = "会计格式3位小数"
             .Caption = "会计格式3位小数"
             .Style = msoButtonIcon
         End With
         With .Controls.Add
            If Application.Calculation = xlCalculationManual Then
               .FaceId = 51
               .TooltipText = "手动计算"
            Else
               .FaceId = 69
               .TooltipText = "自动计算"
            End If
               .OnAction = "手动or自动计算"
               .Caption = "手动or自动计算"
               .Style = msoButtonIcon
          End With
        With .Controls.Add
            .FaceId = 230
            .OnAction = "分页边框"
            .Caption = "分页边框"
            .Style = msoButtonIcon
        End With
        With .Controls.Add
            .FaceId = 107
            .OnAction = "设置分级样式"
            .Caption = "设置分级样式"
            .Style = msoButtonIcon
        End With
        With .Controls.Add
            .FaceId = 581
            .OnAction = "设置分级"
            .Caption = "设置分级"
            .Style = msoButtonIcon
        End With
          
          With .Controls.Add(msoControlPopup, , , , True)
           .Caption = "报价处理"
               With .Controls.Add
                  .Caption = "OUT本表"
                  .OnAction = "OUT本表"
                  .FaceId = 175
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "OUT制作"
                  .OnAction = "OUT制作"
                  .FaceId = 175
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "OUT内部全数据"
                  .OnAction = "OUT内部全数据"
                  .FaceId = 175
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "表格汇总"
                  .OnAction = "表格汇总"
                  .FaceId = 175
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "广联达处理"
                  .OnAction = "广联达处理"
                  .FaceId = 175
                  .Style = msoButtonIconAndCaption
               End With
                With .Controls.Add
                  .Caption = "向下复制"
                  .OnAction = "向下复制"
                  .FaceId = 317
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "表格目录"
                  .OnAction = "工作表目录"
                  .FaceId = 12
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "文件名目录"
                  .OnAction = "文件名目录"
                  .FaceId = 12
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "清除超链接"
                  .OnAction = "删除超链接"
                  .FaceId = 449
                  .Style = msoButtonIconAndCaption
               End With
               With .Controls.Add
                  .Caption = "锁定有效性序列单元格"
                  .OnAction = "锁定有效性序列单元格"
                  .FaceId = 225
                  .Style = msoButtonIconAndCaption
               End With
            End With

         .Visible = True
      End With
    

End Sub


Sub 删除所有加载工具栏()
    On Error Resume Next
    For Each cell In Application.CommandBars
        cell.Delete
    Next
End Sub

Sub 手动or自动计算()
    
    If Application.Calculation = xlManual Then
        Application.Calculation = xlAutomatic
    Else
        Application.Calculation = xlManual
    End If
    
    Call 自动计算图标检查
    
End Sub

Sub 改图片调整名称(ByVal times As Integer)
    On Error Resume Next
    CommandBars("报价工具栏").Controls(5).Reset
    CommandBars("报价工具栏").Controls(5).Caption = "图X" & times
    Call 调整图片
End Sub

Sub 自动计算图标检查()
    On Error Resume Next
    CommandBars("报价工具栏").Controls("手动or自动计算").Reset
    With CommandBars("报价工具栏").Controls("手动or自动计算")
        If Application.Calculation = xlCalculationManual Then
           .FaceId = 51
           .TooltipText = "手动计算"
        Else
           .FaceId = 69
           .TooltipText = "自动计算"
        End If
           .OnAction = "手动or自动计算"
           .Caption = "手动or自动计算"
           .Style = msoButtonIcon
    End With
End Sub

Sub BATCH_MODIFY()
    Call BatchWindow.Show
End Sub



