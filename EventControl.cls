VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Browser As Browser

Private WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents lbl As MSForms.label
Attribute lbl.VB_VarHelpID = -1

Public Sub SetControl(ByRef Control As MSForms.Control)
    Select Case TypeName(Control)
    Case "CheckBox"
        Set chk = Control
        Set Browser = Control.Parent
    Case "TextBox"
        Set txt = Control
        Set Browser = Control.Parent
    Case "Label"
        Set lbl = Control
        Set Browser = Control.Parent
    Case Else
        Err.Raise 1000, "EventControl", "�Ή����ĂȂ���"
    End Select
End Sub


Private Sub ddd_Enter()

End Sub

Private Sub lbl_Click()
    Select Case lbl.Caption
        Case "�V�K�E�B���h�E"
            With UserForms.Add("Browser")
                .Show 0
            End With
        Case "����"
            Unload Browser
        Case Else
            MsgBox "�{�^���ɋ@�\�����ĂȂ���"
    End Select
End Sub
