VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeyHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents TextBoxGroup As MSForms.TextBox
Attribute TextBoxGroup.VB_VarHelpID = -1

Private Sub TextBoxGroup_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload TextBoxGroup.Parent
    ElseIf KeyCode = vbKeyReturn Then
'        BatchWindow.OK_btn_Click
        TextBoxGroup.Parent.OK_btn.Value = True ' 触发按钮的Click事件
    End If
End Sub
