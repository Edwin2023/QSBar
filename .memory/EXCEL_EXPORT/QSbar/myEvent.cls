VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents xlApp As Application
Attribute xlApp.VB_VarHelpID = -1

'Private Sub xlApp_AfterCalculate()
'    Call 更新分摊单元数据
'End Sub


Private Sub xlApp_WorkbookActivate(ByVal wb As Workbook)
'    Call 添加下料工具栏
    Call 自动计算图标检查
'    Call 更新分摊单元数据
End Sub

'Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
'    Call 更新自定义函数路径
'End Sub

