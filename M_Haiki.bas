Attribute VB_Name = "M_Haiki"
Option Explicit
Const START_ROW As Long = 2
Dim EmailAddress As Dictionary
Enum Col
    社員番号 = 1
    HPC
    起票者
    管理者
    Email
End Enum
Sub MakeDraft()

Dim List As Variant
With [A1].CurrentRegion

    '見出しを除いて配列に放り込む
    List = .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0).value
    
End With

'Emailアドレスの辞書を初期化
Set EmailAddress = CreateObject("Scripting.Dictionary")

Dim key As String
Dim ApplicantUserList As Dictionary
Set ApplicantUserList = CreateObject("Scripting.Dictionary")

Dim i As Long
For i = LBound(List) To UBound(List)
    key = List(i, Col.起票者)
    
    If Not ApplicantUserList.Exists(key) Then
        ApplicantUserList.Add key, New Collection
    End If
    
    With New Form
        .Admin = List(i, Col.管理者)
        .HPC = List(i, Col.HPC)
        .AdminEmail = GetEmailAddress(List(i, Col.管理者))
        .ApplicantEmail = GetEmailAddress(key)
    
        ApplicantUserList.item(key).Add .Self
    End With

Next i

Call CreateMail(ApplicantUserList)
:
End Sub
Private Sub CreateMail(ApplicantUserList As Dictionary)
Dim Applicant As Variant
Dim HPC() As String
Dim ApplicantUserForm As Form
Dim cnt As Long, AdminUserList As Dictionary


For Each Applicant In ApplicantUserList
    
    '同じ起票者に送るHPCリストの個数
    cnt = ApplicantUserList.item(Applicant).Count
    ReDim HPC(cnt - 1)  '配列を初期化する
    
    '管理者の重複を排除するためにDictionaryを使う
    Set AdminUserList = CreateObject("Scripting.Dictionary")
    
    'HPCを配列に入れるためにカウンタを用意
    Dim i As Long
    i = 0
    For Each ApplicantUserForm In ApplicantUserList.item(Applicant)
        With ApplicantUserForm
            'HPCを配列に入れてく
            HPC(i) = .HPC
            i = i + 1
            
            
            '管理者は重複を排除。起票者と同じだと除外する
            If Not AdminUserList.Exists(.Admin) And .Admin <> Applicant Then
                AdminUserList.Add .Admin, .AdminEmail
            End If
        End With
    Next
    
    
    With New Mail
        .From = ""
        .mailTo.Add GetEmailAddress(Applicant)
        
        Dim item As Variant
        For Each item In AdminUserList
            .mailCC.Add AdminUserList.item(item)
        Next
        
        .mailBody = "hogehoge"
        .mailSubject = "foobar" & vbCrLf & Join(HPC, vbCrLf)
        .AttachedFile.Add "D:\Documents\Game.pdf"
        .AttachedFile.Add "D:\Documents\code.pdf"
        
        Dim Draft As String
        Draft = vbCrLf & .Create
        With CreateObject("ADODB.Stream")
            Dim SavePath As String
            SavePath = .BuildPath(ThisWorkbook.path, "tmp")
            SavePath = .BuildPath(SavePath, Applicant & ".txt")
            Call SaveTxt(Draft, SavePath)
        End With
    End With
    
Next
End Sub
Private Function GetEmailAddress(ByVal UserCode As String) As String
    If UserCode = "" Then Exit Function
    If Not EmailAddress.Exists(UserCode) Then
        GetEmailAddress = Cells(ActiveSheet.[C:D].Find(What:=UserCode).Row, 5)
    End If
End Function

Private Sub SaveTxt(str, path)
Dim Target As String
    Target = path
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText str
        .SaveToFile Target, 2
        .Close
    End With
End Sub

Private Function LastRow(param As Worksheet, column As Long) As Long
    With param
        LastRow = .Cells(Rows.Count, column).End(xlUp)
    End With
End Function

