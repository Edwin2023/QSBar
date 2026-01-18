VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchWindow 
   Caption         =   "BATCH_MODIFY"
   ClientHeight    =   4990
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   5136
   OleObjectBlob   =   "BatchWindow.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BatchWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_colHandlers As Collection ' 必须声明在窗体模块顶部

Private Sub CenterName_Change()
      validationList
End Sub

Private Sub CommandButton2_Click()
    validationList
End Sub

Private Sub leftName_Change()
      validationList
End Sub

Private Sub rightName_Change()
      validationList
End Sub

Private Sub CommandButton1_Click()
'  BatchWindow.Hide
Unload Me ' 安全关闭当前窗体
End Sub


Private Sub OK_btn_Click()
    Dim Item As Long  ' 使用 Long 避免溢出
    Dim rng As Range
    Dim formula As String

    On Error Resume Next
    

    Item = 1
    For Each rng In Selection.SpecialCells(xlCellTypeVisible)
        formula = rng.formula
        If Left(formula, 1) = "=" Then
            rng.formula = "=" & SumUp(Mid(formula, 2), Item)
        Else
            rng.formula = SumUp(formula, Item)
        End If
        Item = Item + 1
    Next rng
    On Error GoTo 0  ' 恢复正常的错误处理

'    BatchWindow.Hide
    Unload Me ' 安全关闭当前窗体
End Sub

Private Sub update_Click()
    validationList
End Sub

Private Function SumUp(everyValue As String, Item As Long) As String
    Dim LASTVALUE As String

    ' Step 1: 中间取值，要判断 <原名称> 这个参数
    If InStr(1, CenterName.text, "<原名称>") > 0 Then
        LASTVALUE = Replace(CenterName.text, "<原名称>", everyValue)
    Else
        LASTVALUE = CenterName.text
    End If

    ' Step 2: 替换值1 和值2
    LASTVALUE = Replace(LASTVALUE, rpName1.text, rpValue1.text)
    LASTVALUE = Replace(LASTVALUE, rpName2.text, rpValue2.text)

    ' Step 3: 左右和中间值组合
    LASTVALUE = LeftName.text & LASTVALUE & RightName.text

    ' Step 4: 加入 # 序号功能代码
    ' 确保序号替换逻辑不会导致溢出
    If Item > 0 Then
        LASTVALUE = Replace(LASTVALUE, "#", Item)
    Else
        LASTVALUE = Replace(LASTVALUE, "#", "")
    End If

    SumUp = LASTVALUE
End Function

' 更新列表
Private Sub validationList()
    Dim allName As String
    Dim Item As Long
    Dim rng As Range
    Dim visibleCells As Range

    ' 获取所有可见单元格
    On Error Resume Next
    Set visibleCells = Selection.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleCells Is Nothing Then
        MsgBox "没有找到可见单元格。", vbExclamation
        Exit Sub
    End If

    ' 初始化 Item
    Item = 1
    For Each rng In visibleCells
        Dim formula As String
        formula = rng.formula
        If Left(formula, 1) = "=" Then
            allName = allName & "=" & SumUp(Mid(formula, 2), Item) & vbCrLf
        Else
            allName = allName & SumUp(formula, Item) & vbCrLf
        End If
        Item = Item + 1
    Next rng

    ' 将结果赋值给 NameList
    NameList.text = allName
End Sub

Private Sub UserForm_Activate()
    CenterName.text = "<原名称>"

    ' 清空所有 TextBox
    Dim mycontrol As MSForms.Control
    For Each mycontrol In Me.Controls
        If TypeName(mycontrol) = "TextBox" Then
            If mycontrol.Name <> "CenterName" Then
                mycontrol.text = ""
            End If
        End If
    Next mycontrol

    validationList
End Sub


Private Sub UserForm_Initialize()
    Set m_colHandlers = New Collection ' 初始化集合
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "TextBox" Then
            Dim objHandler As New clsKeyHandler
            Set objHandler.TextBoxGroup = ctl
'            Debug.Print ctl.Name
            m_colHandlers.Add objHandler ' 保持实例存活
        End If
    Next
    
End Sub

