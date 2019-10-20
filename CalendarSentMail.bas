Attribute VB_Name = "CalendarSentMail"
Option Explicit
Sub Reset(ws As Worksheet)
    With ws
        .Columns(1).ColumnWidth = 20
        .Columns(2).ColumnWidth = 20
        .Columns(3).ColumnWidth = 40
        .Cells.RowHeight = 18.75
    End With
End Sub
Sub RamdomReset(ws As Worksheet)
    With ws
        .Columns(1).ColumnWidth = 80
        .Columns(2).ColumnWidth = 80
        .Columns(3).ColumnWidth = 160
        .Cells.RowHeight = 50
    End With
End Sub
Sub Make_Click()
    Call Reset(Worksheets("Sheet1"))
    
    
    Dim SendUser As Dictionary
    Set SendUser = CreateObject("Scripting.Dictionary")
    
    
    Dim List As Variant
    List = Worksheets("Sheet1").Range("A1").CurrentRegion.value
    
    Dim userName As String, KeyText As String
    Dim i As Long
    For i = LBound(List) To UBound(List)
        userName = List(i, 1)
        KeyText = GetSendFileName(List(i, 3), "受領テキスト")
        If Not SendUser.Exists(userName) Then
            SendUser.Add userName, New Collection
        End If
        
        SendUser.item(userName).Add KeyText
    Next
    
    Dim user As Variant
    For Each user In SendUser
        Call MakeText(user, SendUser.item(user))
    Next

End Sub
Sub MakeText(ByVal user As String, FileList As Collection)
Dim SavePath As String
With CreateObject("Scripting.FileSystemObject")
    SavePath = .BuildPath(ThisWorkbook.path, "tmp")
    SavePath = .BuildPath(SavePath, user & ".txt")
End With

Dim Body As String
Dim Line As Variant
Body = "本文１" & vbCrLf
Body = Body & String(72, "-") & vbCrLf
For Each Line In FileList
    Body = Body & Line & vbCrLf
Next
Body = Left(Body, Len(Body) - Len(vbCrLf))
Body = Body & String(72, "-") & vbCrLf
Body = Body & "本文２"

With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText Body
        .SaveToFile SavePath, 2
        .Close
End With
End Sub
Private Function GetSendFileName(ByVal text As String, StartKeyText As String) As Variant
    
    Dim Matches As Object
    With CreateObject("VBScript.RegExp")
        .Pattern = "[\s\S]+?(" & StartKeyText & ".+?\n)[\s\S]+"
        .IgnoreCase = False
        .Global = True
        Set Matches = .Execute(text)
    End With
    
    If Matches.Count = 1 Then
        GetSendFileName = Matches(0).submatches(0)
    End If
End Function
