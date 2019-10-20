Attribute VB_Name = "SubTools"
Option Explicit
Declare Function QueryPerformanceCounter Lib "kernel32" _
                           (x As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" _
                           (x As Currency) As Boolean
Dim Freq As Currency
Dim Overhead  As Currency
Dim Ctr1 As Currency, Ctr2 As Currency, result As Currency
'// �~���b�ȉ��̍����x�ŏ������Ԍv���֐�
'// ����    �F(IN)  �z��ϐ�
'// �߂�l  �FBoolean �������ς݁�True�A����������False
Public Sub SWStart()
    If QueryPerformanceCounter(Ctr1) Then
        QueryPerformanceCounter Ctr2
        QueryPerformanceFrequency Freq
'        Debug.Print "QueryPerformanceCounter minimum resolution: 1/" & _
'                    Freq * 10000; " sec"
'        Debug.Print "API Overhead: "; (Ctr2 - Ctr1) / Freq * 1000; "�~���b"
        Overhead = Ctr2 - Ctr1
    Else
        Err.Raise 513, "StopwatchError", "High-resolution counter not supported."
    End If
    QueryPerformanceCounter Ctr1
End Sub

Public Sub SWStop()
    QueryPerformanceCounter Ctr2
    result = (Ctr2 - Ctr1 - Overhead) / Freq * 1000
End Sub

Public Sub SWShow(Optional Caption As String)
    Debug.Print Caption & " " & result
End Sub


'// �z�񏉊�������֐�
'// ����    �F(IN)  �z��ϐ�
'// �߂�l  �FBoolean �������ς݁�True�A����������False
Function IsInitArray(ary()) As Boolean
    If Sgn(ary) <> 0 Then
        IsInitArray = True
    Else
        IsInitArray = False
    End If
End Function


'// �e�L�X�g���N���b�v�{�[�h�֓\��t��
'// ����    �F(IN)  �e�L�X�g
'// �߂�l  �FNothing

Sub PasteClipBoard(text As String)
 With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .text = text
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub


'// �w�肵���V�[�g������������
'// ����    �F(IN)  ���[�N�V�[�g�I�u�W�F�N�g
'// �߂�l  �FNothing
Public Sub InitializeWorkSheet(ws As Worksheet)
Dim PreviousActiveWindow As Window
Set PreviousActiveWindow = ActiveWindow

    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        PreviousActiveWindow.Activate
        With .Cells
            .Clear
            .UseStandardHeight = True
            .UseStandardWidth = True
        End With
    End With
End Sub
'// �w�肵���R���N�V������String�ɕϊ�����
'// ����    �F(IN)  ��؂蕶��
'// �߂�l  �FString
Private Function CollectionToString(myCol As Collection, Delimiter) As String
 
    Dim result  As String
    Dim item As Variant
    
    For Each item In myCol
        result = result & item & Delimiter
    Next
    
    result = Left(result, Len(result) - Len(Delimiter))
    CollectionToString = result
    
End Function

