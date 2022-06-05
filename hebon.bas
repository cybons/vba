Attribute VB_Name = "hebon"
Option Explicit
Function hebon(ByVal kana As String) As String
    Dim temp_data As String
    Dim temp_data2 As String
    Dim temp_data3 As String
    Dim temp_data4 As String
    Dim temp_data5 As String
    Dim temp_data6 As String
    Dim roma_data As String
    Dim result_data As String
    Dim sub_str2 As String
    Dim map As Dictionary
    Set map = initMap
    Dim i As Long
    For i = 1 To Len(kana)
        temp_data = Mid(kana, i, 1)
        temp_data2 = Mid(kana, i, 2)
        If Not map.Exists(temp_data2) Then
            If Not map.Exists(temp_data) Then
                roma_data = roma_data + temp_data
            Else
                roma_data = roma_data + map(temp_data)
            End If
        Else
            i = i + 1
            roma_data = roma_data + map(temp_data2)
        End If
    Next
    
    Dim k As Long
    For k = 1 To Len(roma_data)
        temp_data = Mid(roma_data, k, 1)
        temp_data2 = Mid(roma_data, k, 2)
        temp_data3 = Mid(roma_data, k, 3)
        temp_data4 = Mid(roma_data, k, 4)
        temp_data5 = Mid(roma_data, k, 5)
        temp_data6 = Mid(roma_data, k, 6)
        sub_str2 = Mid(temp_data2, 1, 2)
        If temp_data4 = "noue" Then
            k = k + 3
            temp_data = "noue"
        ElseIf temp_data6 = "touchi" Then
            k = k + 5
            temp_data = "touchi"
        Else
            
            Dim re As Object
            Set re = CreateObject("VBScript.RegExp")
            With re
                .Pattern = "/[a-z]/gi"       ''�����p�^�[����ݒ�
                .IgnoreCase = True          ''�啶���Ə���������ʂ��Ȃ�
                .Global = True
            End With
            If temp_data2 = "uu" Or temp_data2 = "ee" Or temp_data2 = "ou" Or temp_data2 = "oo" Then
                k = k + 1
            ElseIf temp_data = "��" Or temp_data = "�b" Then
                If temp_data3 = "��ch" Or temp_data3 = "�bch" Then
                    temp_data = "t"
                ElseIf re.test(sub_str2) Then
                    temp_data = sub_str2
                Else
                    temp_data = "tsu"
                End If
            End If
        End If
        result_data = result_data + temp_data
    Next
    
    hebon = result_data
End Function
Function initMap() As Dictionary
Static map As Dictionary
If Not map Is Nothing Then
    Set initMap = map
    Exit Function
End If
Set map = New Dictionary
With map
    .Add "��", "a"
    .Add "��", "i"
    .Add "��", "u"
    .Add "��", "e"
    .Add "��", "o"
    .Add "��", "ka"
    .Add "��", "ki"
    .Add "��", "ku"
    .Add "��", "ke"
    .Add "��", "ko"
    .Add "��", "sa"
    .Add "��", "shi"
    .Add "��", "su"
    .Add "��", "se"
    .Add "��", "so"
    .Add "��", "ta"
    .Add "��", "chi"
    .Add "��", "tsu"
    .Add "��", "te"
    .Add "��", "to"
    .Add "��", "na"
    .Add "��", "ni"
    .Add "��", "nu"
    .Add "��", "ne"
    .Add "��", "no"
    .Add "��", "ha"
    .Add "��", "hi"
    .Add "��", "fu"
    .Add "��", "he"
    .Add "��", "ho"
    .Add "��", "ma"
    .Add "��", "mi"
    .Add "��", "mu"
    .Add "��", "me"
    .Add "��", "mo"
    .Add "��", "ya"
    .Add "��", "yu"
    .Add "��", "yo"
    .Add "��", "ra"
    .Add "��", "ri"
    .Add "��", "ru"
    .Add "��", "re"
    .Add "��", "ro"
    .Add "��", "wa"
    .Add "��", "i"
    .Add "��", "e"
    .Add "��", "o"
    .Add "��", "n"
    .Add "��", "a"
    .Add "��", "i"
    .Add "��", "u"
    .Add "��", "e"
    .Add "��", "o"
    .Add "��", "ga"
    .Add "��", "gi"
    .Add "��", "gu"
    .Add "��", "ge"
    .Add "��", "go"
    .Add "��", "za"
    .Add "��", "ji"
    .Add "��", "zu"
    .Add "��", "ze"
    .Add "��", "zo"
    .Add "��", "da"
    .Add "��", "ji"
    .Add "��", "zu"
    .Add "��", "de"
    .Add "��", "do"
    .Add "��", "ba"
    .Add "��", "bi"
    .Add "��", "bu"
    .Add "��", "be"
    .Add "��", "bo"
    .Add "��", "pa"
    .Add "��", "pi"
    .Add "��", "pu"
    .Add "��", "pe"
    .Add "��", "po"
    .Add "����", "kya"
    .Add "����", "kyu"
    .Add "����", "kyo"
    .Add "����", "sha"
    .Add "����", "shu"
    .Add "����", "sho"
    .Add "����", "cha"
    .Add "����", "chu"
    .Add "����", "cho"
    .Add "����", "che"
    .Add "�ɂ�", "nya"
    .Add "�ɂ�", "nyu"
    .Add "�ɂ�", "nyo"
    .Add "�Ђ�", "hya"
    .Add "�Ђ�", "hyu"
    .Add "�Ђ�", "hyo"
    .Add "�݂�", "mya"
    .Add "�݂�", "myu"
    .Add "�݂�", "myo"
    .Add "���", "rya"
    .Add "���", "ryu"
    .Add "���", "ryo"
    .Add "����", "gya"
    .Add "����", "gyu"
    .Add "����", "gyo"
    .Add "����", "ja"
    .Add "����", "ju"
    .Add "����", "jo"
    .Add "�т�", "bya"
    .Add "�т�", "byu"
    .Add "�т�", "byo"
    .Add "�҂�", "pya"
    .Add "�҂�", "pyu"
    .Add "�҂�", "pyo"
    .Add "�A", "a"
    .Add "�C", "i"
    .Add "�E", "u"
    .Add "�G", "e"
    .Add "�I", "o"
    .Add "�J", "ka"
    .Add "�L", "ki"
    .Add "�N", "ku"
    .Add "�P", "ke"
    .Add "�R", "ko"
    .Add "�T", "sa"
    .Add "�V", "shi"
    .Add "�X", "su"
    .Add "�Z", "se"
    .Add "�\", "so"
    .Add "�^", "ta"
    .Add "�`", "chi"
    .Add "�c", "tsu"
    .Add "�e", "te"
    .Add "�g", "to"
    .Add "�i", "na"
    .Add "�j", "ni"
    .Add "�k", "nu"
    .Add "�l", "ne"
    .Add "�m", "no"
    .Add "�n", "ha"
    .Add "�q", "hi"
    .Add "�t", "fu"
    .Add "�w", "he"
    .Add "�z", "ho"
    .Add "�}", "ma"
    .Add "�~", "mi"
    .Add "��", "mu"
    .Add "��", "me"
    .Add "��", "mo"
    .Add "��", "ya"
    .Add "��", "yu"
    .Add "��", "yo"
    .Add "��", "ra"
    .Add "��", "ri"
    .Add "��", "ru"
    .Add "��", "re"
    .Add "��", "ro"
    .Add "��", "wa"
    .Add "��", "i"
    .Add "��", "e"
    .Add "��", "o"
    .Add "��", "n"
    .Add "�@", "a"
    .Add "�B", "i"
    .Add "�D", "u"
    .Add "�F", "e"
    .Add "�H", "o"
    .Add "�K", "ga"
    .Add "�M", "gi"
    .Add "�O", "gu"
    .Add "�Q", "ge"
    .Add "�S", "go"
    .Add "�U", "za"
    .Add "�W", "ji"
    .Add "�Y", "zu"
    .Add "�[", "ze"
    .Add "�]", "zo"
    .Add "�_", "da"
    .Add "�a", "ji"
    .Add "�d", "zu"
    .Add "�f", "de"
    .Add "�h", "do"
    .Add "�o", "ba"
    .Add "�r", "bi"
    .Add "�u", "bu"
    .Add "�x", "be"
    .Add "�{", "bo"
    .Add "�p", "pa"
    .Add "�s", "pi"
    .Add "�v", "pu"
    .Add "�y", "pe"
    .Add "�|", "po"
    .Add "�L��", "kya"
    .Add "�L��", "kyu"
    .Add "�L��", "kyo"
    .Add "�V��", "sha"
    .Add "�V��", "shu"
    .Add "�V��", "sho"
    .Add "�`��", "cha"
    .Add "�`��", "chu"
    .Add "�`��", "cho"
    .Add "�j��", "nya"
    .Add "�j��", "nyu"
    .Add "�j��", "nyo"
    .Add "�q��", "hya"
    .Add "�q��", "hyu"
    .Add "�q��", "hyo"
    .Add "�~��", "mya"
    .Add "�~��", "myu"
    .Add "�~��", "myo"
    .Add "����", "rya"
    .Add "����", "ryu"
    .Add "����", "ryo"
    .Add "�M��", "gya"
    .Add "�M��", "gyu"
    .Add "�M��", "gyo"
    .Add "�W��", "ja"
    .Add "�W��", "ju"
    .Add "�W��", "jo"
    .Add "�r��", "bya"
    .Add "�r��", "byu"
    .Add "�r��", "byo"
    .Add "�s��", "pya"
    .Add "�s��", "pyu"
    .Add "�s��", "pyo"
    .Add "�W�F", "jie"
    .Add "�`�F", "chie"
    .Add "�e�B", "tei"
    .Add "�f�B", "dei"
    .Add "�f��", "deyu"
    .Add "�t�@", "fua"
    .Add "�t�B", "fui"
    .Add "�t�F", "fue"
    .Add "�t�H", "fuo"
    .Add "���@", "bua"
    .Add "���B", "bui"
    .Add "��", "bu"
    .Add "���F", "bue"
    .Add "���H", "buo"
    .Add "�[", ""

End With
Set initMap = map
End Function
