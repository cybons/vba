Attribute VB_Name = "Module1"
Option Explicit


Sub ttt()

Dim myForm As Browser, i As Long
    With UserForms.Add("Browser")
        Dim Control As Control
        With .Controls.Add("Forms.Label.1", _
                          "Label" & Format(i, "0#"))    'ÅcÅc(2)'
          .Height = 12
          .Width = 72
          .Top = 0
          .Left = i * 72
          .Caption = "Label" & Format(i, "0#")
          
        End With
        .WebView.Top = 12
        .Caption = "Popup"
        .Show 0
    End With
End Sub
Sub addControl()
Dim myForm As Browser, i As Long
    With UserForms.Add("Browser")
        Dim Control As Control
        For Each Control In .Controls
            If TypeName(Control) = "Label" Then
                'Control.Visible = False
            End If
        Next
        .WebView.Top = 12
        .Caption = "Popup"
        .Show 0
    End With
End Sub
Sub test2()
MsgBox 1
End Sub
Public Sub RunBrowser()
    With UserForms.Add("Browser")
        .Show 0
    End With
End Sub

