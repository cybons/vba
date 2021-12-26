Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'�E�B���h�E�v���V�[�W���̃R�[��
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'�z�b�g�L�[��o�^����
Private Declare Function RegisterHotKey Lib "user32.dll" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'�z�b�g�L�[����������
Private Declare Function UnregisterHotKey Lib "user32.dll" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'�萔��`
Private Const GWL_WNDPROC = (-4) '�E�B���h�E�v���V�[�W��
Private Const MOD_CONTROL As Integer = &H2 '�R���g���[���L�[
Private hOldWndProc As Long '�I���W�i���E�B���h�E�v���V�[�W���̃n���h��
Private Const WM_HOTKEY = &H312 '�z�b�g�L�[�̃E�B���h�E���b�Z�[�W
Private Const HK_EVENT_CTRL_C = &H100 '�z�b�g�L�[�ԍ�1
Private Const HK_EVENT_CTRL_V = &H101 '�z�b�g�L�[�ԍ�2

Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If msg = WM_HOTKEY Then
    If wParam = HK_EVENT_CTRL_C Then
        MsgBox "�������"
        
    End If
End If
'�I���W�i���v���V�[�W�����R�[��
WndProc = CallWindowProc(hOldWndProc, hWnd, msg, wParam, lParam)
End Function

Public Sub Regist(hWnd)
Debug.Print "regi" & hWnd
Debug.Print "appli" & Application.hWnd
'�z�b�g�L�[�̓o�^
Debug.Print RegisterHotKey(hWnd, HK_EVENT_CTRL_C, MOD_CONTROL, vbKeyN)

'�E�B���h�E�v���V�[�W���̕ύX
hOldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub Unregist(hWnd)
'�E�B���h�E�v���V�[�W�������ɖ߂�
Call SetWindowLong(hWnd, GWL_WNDPROC, hOldWndProc)

'�z�b�g�L�[�̉���
Call UnregisterHotKey(hWnd, HK_EVENT_CTRL_C)
End Sub


