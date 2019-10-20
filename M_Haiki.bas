Attribute VB_Name = "M_Haiki"
Option Explicit
Const START_ROW As Long = 2
Dim EmailAddress As Dictionary
Enum Col
    �Ј��ԍ� = 1
    HPC
    �N�[��
    �Ǘ���
    Email
End Enum
Sub MakeDraft()

Dim List As Variant
With [A1].CurrentRegion

    '���o���������Ĕz��ɕ��荞��
    List = .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0).value
    
End With

'Email�A�h���X�̎�����������
Set EmailAddress = CreateObject("Scripting.Dictionary")

Dim key As String
Dim ApplicantUserList As Dictionary
Set ApplicantUserList = CreateObject("Scripting.Dictionary")

Dim i As Long
For i = LBound(List) To UBound(List)
    key = List(i, Col.�N�[��)
    
    If Not ApplicantUserList.Exists(key) Then
        ApplicantUserList.Add key, New Collection
    End If
    
    With New Form
        .Admin = List(i, Col.�Ǘ���)
        .HPC = List(i, Col.HPC)
        .AdminEmail = GetEmailAddress(List(i, Col.�Ǘ���))
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
    
    '�����N�[�҂ɑ���HPC���X�g�̌�
    cnt = ApplicantUserList.item(Applicant).Count
    ReDim HPC(cnt - 1)  '�z�������������
    
    '�Ǘ��҂̏d����r�����邽�߂�Dictionary���g��
    Set AdminUserList = CreateObject("Scripting.Dictionary")
    
    'HPC��z��ɓ���邽�߂ɃJ�E���^��p��
    Dim i As Long
    i = 0
    For Each ApplicantUserForm In ApplicantUserList.item(Applicant)
        With ApplicantUserForm
            'HPC��z��ɓ���Ă�
            HPC(i) = .HPC
            i = i + 1
            
            
            '�Ǘ��҂͏d����r���B�N�[�҂Ɠ������Ə��O����
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

