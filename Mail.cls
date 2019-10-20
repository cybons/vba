VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public From As String
Public mailTo As Collection
Public mailCC As Collection
Public mailSubject As String
Public mailBody As String
Public AttachedFile As Collection
Private boundaryStr As String
Private Draft As Mime

'// 指定したコレクションをStringに変換する
'// 引数    ：(IN)  区切り文字
'// 戻り値  ：String
Private Function CollectionToString(myCol As Collection, Delimiter) As String
 
    Dim result  As String
    Dim item As Variant
    
    For Each item In myCol
        result = result & item & Delimiter
    Next
    
    result = Left(result, Len(result) - Len(Delimiter))
    CollectionToString = result
    
End Function


Private Sub Class_Initialize()
    Set Draft = New Mime
    Set mailTo = New Collection
    Set mailCC = New Collection
    Set AttachedFile = New Collection
    boundaryStr = GetUniqueKey
End Sub

Public Function Create()
'ヘッダー
    With Draft.Header
        .Add "To", CollectionToString(mailTo, ", ")
        .Add "Cc", CollectionToString(mailCC, ", ")
        .Add "From", From
        .Add "Subject", "=?utf-8?B?" & EncodeBase64("Str", mailSubject, False) & "?="
        .Add "MIME-Version", "1.0"
        .Add "Content-Type", "multipart/mixed; boundary=""" & boundaryStr & """"
    End With

    'ボディ(1) 本文
    Draft.Body.Add New Mime
    With Draft.Body(1)
        With .Header
            .Add "Content-Type", "text/plain; charset=""utf-8"""
            .Add "Content-Transfer-Encoding", "base64"
        End With
        .Body.Add EncodeBase64("Str", mailBody, False)
    End With

    Dim AttachedFileName As String
    Dim AttachedFilePath As Variant
    Dim i As Long: i = 2
    
    For Each AttachedFilePath In AttachedFile
    
        AttachedFileName = CreateObject("Scripting.FileSystemObject").GetFileName(AttachedFilePath)
    
        'ボディ(2) 添付ファイル
        Draft.Body.Add New Mime
        With Draft.Body(i)
            With .Header
                .Add "Content-Type", "application/octet-stream; name=""" & AttachedFileName & """"
                .Add "Content-Transfer-Encoding", "base64"
                .Add "Content-Disposition", "attachment; filename=""" & AttachedFileName & """"
            End With
            .Body.Add EncodeBase64("File", AttachedFilePath, False)
        End With
        i = i + 1
    Next
    Create = Draft.ConvertToStr(boundaryStr)
End Function
Private Function EncodeBase64( _
    ByVal tgtType As String, _
    ByVal Target As String, _
    ByVal urlSafe As Boolean) As String

    Dim base64 As Object
    Set base64 = CreateObject("MSXML2.DOMDocument").createElement("base64")
    base64.DataType = "bin.base64"

    With CreateObject("ADODB.Stream")
        Select Case tgtType
        Case "File"
            .Type = 1 'adTypeBinary
            .Open
            .LoadFromFile Target
        Case "Str"
            .Type = 2 'adTypeText
            .Charset = "UTF-8"
            .Open
            .WriteText Target
            .Position = 0
            .Type = 1 'adTypeBinary
            .Position = 3
        End Select
        base64.nodeTypedValue = .Read
        .Close
    End With

    EncodeBase64 = base64.text
    'URLの予約語を取り除く
    If urlSafe = True Then _
        EncodeBase64 = Replace(Replace(EncodeBase64, "+", "-"), "/", "_")

End Function
Function GetUniqueKey() As String
    Static n As Long
    Static storedStamp As String
    Dim timeStamp As String
    timeStamp = _
        RadixConversion(CLng(Date), 36) & "_" & _
        RadixConversion(Hour(Time) * 60 ^ 2 + Minute(Time) * 60 + Second(Time), 36)

    If storedStamp = timeStamp Then
        n = n + 1
    Else
        n = 0
    End If
    
    storedStamp = timeStamp
    GetUniqueKey = timeStamp & "#" & n
End Function

Public Function RadixConversion(ByVal num As Long, ByVal Radix As Long) As String
    Dim Quotient As Long
    Dim Remainder As Long
    Dim Answer As String
    Quotient = num
    Do
        Remainder = Quotient Mod Radix
        Quotient = Quotient \ Radix
        Answer = GetNumChar(Remainder) & Answer
    Loop Until Quotient = 0
    RadixConversion = Answer
End Function

Private Function GetNumChar(ByVal num As Long) As String
    Dim temp As Variant
    temp = Split("0 1 2 3 4 5 6 7 8 9 A B C D E F G H I J K L M N O P Q R S T U V W X Y Z")
    GetNumChar = temp(num)
End Function

