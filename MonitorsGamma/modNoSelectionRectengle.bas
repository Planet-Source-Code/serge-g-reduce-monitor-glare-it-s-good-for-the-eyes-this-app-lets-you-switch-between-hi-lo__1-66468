Attribute VB_Name = "ModGamma"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public lpPrevWndProc As Long
Private Const WM_SETFOCUS = &H7

Public Sub Hook(mHwnd As Long)
    lpPrevWndProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook(mHwnd As Long)
    SetWindowLong mHwnd, GWL_WNDPROC, lpPrevWndProc
End Sub

Function WindowProc(ByVal mHwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case WM_SETFOCUS
            Exit Function
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, mHwnd, uMsg, wParam, lParam)

End Function



