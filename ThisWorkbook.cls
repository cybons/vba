VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private WithEvents Bkob As Workbook
Attribute Bkob.VB_VarHelpID = -1

Sub SetBook()
If Bkob Is Nothing Then
Set Bkob = ActiveWorkbook
Else
Set Bkob = Nothing
End If
End Sub
Sub UnsetBook()

End Sub
Private Sub Bkob_Activate()

If Browser.Visible = False Then
    PopupBrowser.Show 0
End If
End Sub

Private Sub Bkob_BeforeClose(Cancel As Boolean)
Unload Browser
Set Bkob = Nothing
End Sub


Public Sub active()
With UserForms.Add("Browser")
    .Show 0
End With

End Sub

