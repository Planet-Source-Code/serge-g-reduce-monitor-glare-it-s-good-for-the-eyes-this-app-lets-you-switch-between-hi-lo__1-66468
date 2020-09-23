VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Low"
   ClientHeight    =   1005
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1684
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      DownPicture     =   "frmMain.frx":21CE
      Height          =   375
      Left            =   75
      Picture         =   "frmMain.frx":2B68
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Exit Application"
      Top             =   450
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   795
      Picture         =   "frmMain.frx":3502
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Lower Brightness"
      Top             =   150
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   75
      Picture         =   "frmMain.frx":3F3C
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Restore Brightness"
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   0
      TabIndex        =   3
      Top             =   -60
      Width           =   1605
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   945
         Left            =   15
         Top             =   105
         Width           =   1575
      End
   End
   Begin VB.PictureBox picHook 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgIconHold 
      Height          =   240
      Left            =   720
      Picture         =   "frmMain.frx":4976
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image imgIconOff 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":53B0
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconOn 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":5DEA
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuRestoreItem 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeItem 
         Caption         =   "Change To High"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEndItem 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Ramp1(0 To 255, 0 To 2) As Integer
Private Ramp2(0 To 255, 0 To 2) As Integer
Private Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Private Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'''''''''''''''''''TopMost
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'''''''''''''''''''TopMost
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Dim isLow As Boolean
Dim actNow As Boolean

Private Sub Command1_Click()

   setHigh

End Sub

Private Sub Command2_Click()

    startUp

End Sub

Private Sub Command3_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    If App.PrevInstance = True Then End
   
    gammaStatus = "High"
   
    alwaysOnTop
    startUp
    actNow = False
    
    Dim ctl As Control

    For Each ctl In Me.Controls
        If (TypeOf ctl Is CheckBox) Or (TypeOf ctl Is CommandButton) Or (TypeOf ctl Is OptionButton) Then
            'If Len(ctl.Tag) <> 0 Then
                If ctl.Tag <> "unhook" Then
                    Hook ctl.hWnd
                End If
            'End If
        End If
    Next
    
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then
        Hide
        init
    ElseIf Me.WindowState = vbNormal Then
        Show
        If actNow = False Then
            waitTime 0.6
        End If
        actNow = False
        termin
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
   
   SetDeviceGammaRamp Me.hdc, Ramp1(0, 0)
   termin
   
End Sub
Public Function Int2Lng(IntVal As Integer) As Long
   
   CopyMemory Int2Lng, IntVal, 2
   
End Function
Public Function Lng2Int(Value As Long) As Integer
   
   CopyMemory Lng2Int, Value, 2
   
End Function

Private Sub alwaysOnTop()

    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub startUp()
    
    If isLow Then
        mnuChangeItem.Caption = "Change To High"
        Exit Sub
    Else
        mnuChangeItem.Caption = "Change To Low"
    End If
    
    'Me.Icon = imgOff.Picture
    Me.Caption = "Low"
    Me.Icon = ImageList1.ListImages(1).Picture
    imgIconHold.Picture = imgIconOn.Picture
    If Me.WindowState = vbMinimized Then termin: init
    
    Dim iCtr       As Integer
    Dim lVal       As Long
   
    GetDeviceGammaRamp Me.hdc, Ramp1(0, 0)
        For iCtr = 0 To 255
            lVal = Int2Lng(Ramp1(iCtr, 0))
            Ramp2(iCtr, 0) = Lng2Int(Int2Lng(Ramp1(iCtr, 0)) / 2)

            Ramp2(iCtr, 1) = Lng2Int(Int2Lng(Ramp1(iCtr, 1)) / 2)
            Ramp2(iCtr, 2) = Lng2Int(Int2Lng(Ramp1(iCtr, 2)) / 2)
        Next iCtr
    SetDeviceGammaRamp Me.hdc, Ramp2(0, 0)
    
    isLow = True
    gammaStatus = "Low"
    mnuChangeItem.Caption = "Change To High"
   
End Sub


Private Sub mnuChangeItem_Click()

    If LCase(mnuChangeItem.Caption) = "change to low" Then
        Command2_Click
    ElseIf LCase(mnuChangeItem.Caption) = "change to high" Then
        Command1_Click
    Else     'precaution
        If isLow = False Then
            Command2_Click
        Else
            Command1_Click
        End If
    End If

End Sub

Private Sub mnuEndItem_Click()

    Unload Me

End Sub

Private Sub mnuRestoreItem_Click()

    actNow = True
    Me.WindowState = vbNormal
    Show

End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        Me.WindowState = vbNormal
        Show
        'Me.SetFocus
    ElseIf Msg = WM_RBUTTONUP Then
        PopupMenu mnuPopUp
    End If
    
End Sub

Sub waitTime(sec As Long)

    Dim start As Long
    start = Timer
    Do While Timer < start + sec
        DoEvents
    Loop

End Sub

Sub setHigh()

    SetDeviceGammaRamp Me.hdc, Ramp1(0, 0)
   
    'Me.Icon = imgOn.Picture
    imgIconHold.Picture = imgIconOff.Picture
    If Me.WindowState = vbMinimized Then termin: init
    Me.Icon = ImageList1.ListImages(2).Picture
    Me.Caption = "High"
    isLow = False
    gammaStatus = "High"
    mnuChangeItem.Caption = "Change To Low"

End Sub
