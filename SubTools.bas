Attribute VB_Name = "SubTools"
Option Explicit
Declare Function QueryPerformanceCounter Lib "kernel32" _
                           (x As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" _
                           (x As Currency) As Boolean
Dim Freq As Currency
Dim Overhead  As Currency
Dim Ctr1 As Currency, Ctr2 As Currency, result As Currency
'// ミリ秒以下の高精度で処理時間計測関数
'// 引数    ：(IN)  配列変数
'// 戻り値  ：Boolean 初期化済み＝True、未初期化＝False
Public Sub SWStart()
    If QueryPerformanceCounter(Ctr1) Then
        QueryPerformanceCounter Ctr2
        QueryPerformanceFrequency Freq
'        Debug.Print "QueryPerformanceCounter minimum resolution: 1/" & _
'                    Freq * 10000; " sec"
'        Debug.Print "API Overhead: "; (Ctr2 - Ctr1) / Freq * 1000; "ミリ秒"
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


'// 配列初期化判定関数
'// 引数    ：(IN)  配列変数
'// 戻り値  ：Boolean 初期化済み＝True、未初期化＝False
Function IsInitArray(ary()) As Boolean
    If Sgn(ary) <> 0 Then
        IsInitArray = True
    Else
        IsInitArray = False
    End If
End Function


'// テキストをクリップボードへ貼り付け
'// 引数    ：(IN)  テキスト
'// 戻り値  ：Nothing

Sub PasteClipBoard(text As String)
 With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .text = text
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub


'// 指定したシートを初期化する
'// 引数    ：(IN)  ワークシートオブジェクト
'// 戻り値  ：Nothing
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

