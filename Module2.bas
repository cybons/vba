Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'ウィンドウプロシージャのコール
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'ホットキーを登録する
Private Declare Function RegisterHotKey Lib "user32.dll" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'ホットキーを解除する
Private Declare Function UnregisterHotKey Lib "user32.dll" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'定数定義
Private Const GWL_WNDPROC = (-4) 'ウィンドウプロシージャ
Private Const MOD_CONTROL As Integer = &H2 'コントロールキー
Private hOldWndProc As Long 'オリジナルウィンドウプロシージャのハンドル
Private Const WM_HOTKEY = &H312 'ホットキーのウィンドウメッセージ
Private Const HK_EVENT_CTRL_C = &H100 'ホットキー番号1
Private Const HK_EVENT_CTRL_V = &H101 'ホットキー番号2

Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If msg = WM_HOTKEY Then
    If wParam = HK_EVENT_CTRL_C Then
        MsgBox "無効状態"
        
    End If
End If
'オリジナルプロシージャをコール
WndProc = CallWindowProc(hOldWndProc, hWnd, msg, wParam, lParam)
End Function

Public Sub Regist(hWnd)
Debug.Print "regi" & hWnd
Debug.Print "appli" & Application.hWnd
'ホットキーの登録
Debug.Print RegisterHotKey(hWnd, HK_EVENT_CTRL_C, MOD_CONTROL, vbKeyN)

'ウィンドウプロシージャの変更
hOldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub Unregist(hWnd)
'ウィンドウプロシージャを元に戻す
Call SetWindowLong(hWnd, GWL_WNDPROC, hOldWndProc)

'ホットキーの解除
Call UnregisterHotKey(hWnd, HK_EVENT_CTRL_C)
End Sub


