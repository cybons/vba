Attribute VB_Name = "Module4"
Option Explicit

'-----------------------------------------------------------
' �@�\: �J���Ă���WebBrowser����HTMLDocument�����ׂĕԂ�
' ����: Num/1-5�͈̔�
' �Ԃ�l: Collection
'-----------------------------------------------------------
Function GetHTMLDocuments() As Collection
Set GetHTMLDocuments = New Collection
Dim Form As UserForm, Control
For Each Form In UserForms
    For Each Control In Form.Controls
        If TypeName(Control) = "WebBrowser" Then
            GetHTMLDocuments.Add Control
        End If
    Next
Next
End Function

Sub InitNewWinodow(ByVal Number As Long)
If Number > 5 Or Number < 1 Then Call Err.Raise(600, "InitNewWinodow", "Window�̎w���1-5�Ŏw�肵�Ă�������")

UserForms("NewWindow" & Number).Show


End Sub

Function RunBrowser() As Browser
Set RunBrowser = Browser

End Function

Sub RunBrowswer2()

Browser.Show vbModeless

End Sub
Private Sub Auto_Open()
    Application.OnKey "+^I", "RunBrowswer2"
End Sub

