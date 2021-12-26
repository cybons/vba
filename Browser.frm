VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Browser 
   Caption         =   "Browser"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "Browser.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Windows API�錾
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function WindowFromObject Lib "oleacc" Alias "WindowFromAccessibleObject" (ByVal pacc As Object, phwnd As LongPtr) As LongPtr
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private WithEvents XLApp As Excel.Application
Attribute XLApp.VB_VarHelpID = -1
Private mXLHwnd As LongPtr      'Excel's window handle
Private mhwndForm As LongPtr    'userform's window handle

Private Const WindowWidth As Long = 1920
Private Const WindowHeight As Long = 1080

Private keyHookHwnd As Long

Const GWL_HWNDPARENT As Long = -8   '�e�E�B���h�E�̃n���h��
Private EventList As Collection



Private Property Get hWnd() As Long
    WindowFromObject Me, hWnd
End Property
' �t�H�[�������T�C�Y�\�ɂ��邽�߂̐ݒ�
Public Sub FormSetting()
    Dim result As Long
    Dim hWnd As Long
    Dim Wnd_STYLE As Long
    
 
    hWnd = GetActiveWindow()
    Debug.Print "form" & hWnd
    
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE Or WS_THICKFRAME Or &H30000
    
 
    result = SetWindowLong(hWnd, GWL_STYLE, Wnd_STYLE)
    result = DrawMenuBar(hWnd)
End Sub





Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'If MsgBox("�E�B���h�E����܂�����낵���ł����H", vbYesNo) = vbNo Then
'    Cancel = True
'End If

If Label1.Visible = False Then
    Cancel = False
    Call Unregist(keyHookHwnd)
    Exit Sub
End If
    If CloseMode = 0 Then
        MsgBox "[�~]�{�^���ł͕����܂���B", 48
        Cancel = True
    End If
End Sub

Private Sub UserForm_Resize()
    Dim iHeight
    Dim iWidth
    
    '// Height��+36�͖ڎ����������B
    iWidth = Me.InsideWidth - WebView.Left * 2
    iHeight = Me.InsideHeight - WebView.Top * 2 + 36
    
    '// Width��Height�ɂ�0�ȉ��͐ݒ�s�̂��߃G���[�ɂȂ�̂łO�`�F�b�N
    If (iWidth > 0 And iHeight > 0) Then
        WebView.Width = iWidth
        WebView.Height = iHeight
    End If
    
    
End Sub
Private Sub UserForm_Activate()
    keyHookHwnd = hWnd
    Debug.Print "key" & hWnd

    Call Regist(keyHookHwnd)
Call Popup(GetActiveWindow())
    Call FormSetting
End Sub

Private Sub UserForm_Initialize()

    'Call SetForegroundWindow(hwnd)
    WebView.Navigate2 "www.google.com"
    
    Set EventList = New Collection
    Dim Con As Control
    Dim EC As EventControl
    For Each Con In Me.Controls
        Select Case TypeName(Con)
        Case "Label", "TextBox"
            Set EC = New EventControl
            EC.SetControl Con
            EventList.Add EC
        End Select
    Next
    Dim br
    
    'Excel2013(ver15)-  SDI
    If Val(Application.Version) >= 15 Then
        Set XLApp = Application
        '���[�U�[�t�H�[���̃n���h����Caption������擾
        mhwndForm = FindWindow("ThunderDFrame", Caption)
    End If
    'Call RunKey
End Sub
Private Sub Popup(hWnd As Long)
'Excel���őO�ʂɐݒ�(��ɍőO�ʂɐݒ肵�ċ����I�ɍőO�ʂɈړ������Ă���A�u��Ɂv���O��)
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
End Sub

Private Sub WebView_NewWindow2(ppDisp As Object, Cancel As Boolean)
    'Cancel = False
    
    Dim PopupBrowser As Browser, i As Long
    Set PopupBrowser = UserForms.Add("Browser")
    With PopupBrowser
        Dim Control As Control
        For Each Control In .Controls
            If TypeName(Control) = "Label" Then
                Control.Visible = False
            End If
        Next
        .WebView.Top = 0
    End With

    PopupBrowser.Show 0
    Set ppDisp = PopupBrowser.WebView
End Sub




'�� WindowActivate ---------------------------------------------
Private Sub XLApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    SetWindowLong mhwndForm, GWL_HWNDPARENT, Application.hWnd
    SetForegroundWindow mhwndForm
End Sub
 
'�� WorkbookBeforeClose --------------------------------------------
Private Sub XLApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    SetWindowLong mhwndForm, GWL_HWNDPARENT, 0&
End Sub




''
'Private Sub RunKey()
'    ' ** F1�L�[�𖳌��ɂ���
'    Application.OnKey "{F1}", ""
'    Application.OnKey "^n", "msg"
'End Sub
'Sub msg()
'MsgBox "test"
'End Sub

