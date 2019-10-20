VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Header As Object
Public Body As Object

Private Sub Class_Initialize()
    Set Header = CreateObject("Scripting.Dictionary")
    Set Body = New Collection
End Sub

Public Function ConvertToStr(ByVal boundaryStr As String) As String

    ConvertToStr = DicToStr(Header) & "--" & boundaryStr & vbCrLf
    Dim part As Variant
    For Each part In Body
        ConvertToStr = ConvertToStr & _
            DicToStr(part.Header) & part.Body(1) & vbCrLf & "--" & boundaryStr & vbCrLf
    Next
    ConvertToStr = Left(ConvertToStr, Len(ConvertToStr) - 2) & "--"

End Function

Private Function DicToStr(ByVal dic As Object) As String

    Dim key As Variant
    For Each key In dic
        DicToStr = DicToStr & key & ": " & dic(key) & vbCrLf
    Next
    DicToStr = DicToStr & vbCrLf

End Function

