Attribute VB_Name = "TrayIcon"
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205

Private TrayI As NOTIFYICONDATA

Public gammaStatus As String

Public Sub init()

TrayI.cbSize = Len(TrayI)

    TrayI.hWnd = frmMain.picHook.hWnd
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = frmMain.imgIconHold.Picture
    TrayI.szTip = "Your monitor's brightness is " & LCase(gammaStatus) & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TrayI

End Sub

Public Sub termin()

    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = frmMain.picHook.hWnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI

End Sub

