VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "sysTray"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'http://support.microsoft.com/kb/176085
Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon
    Me.Icon = LoadPicture(NmeaRouterIcon)
'    Me.Show
'    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        If CurrentProfile <> "" Then
            .szTip = CurrentProfile & vbNullChar
        Else
        .szTip = "NmeaRouter " & App.Major & "." & App.Minor & "." & App.Revision & vbNullChar
        End If
    End With
'creates the SysTray icon
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'this procedure receives the callbacks from the System Tray icon.
Dim result As Long
Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
    Case WM_LBUTTONUP        '514 restore form window
'         Me.WindowState = vbNormal
'         result = SetForegroundWindow(Me.hwnd)
'         Me.Show
    Case WM_LBUTTONDBLCLK    '515 restore form window
'         Me.WindowState = vbNormal
'         result = SetForegroundWindow(Me.hwnd)
'         Me.Show
    Case WM_RBUTTONUP        '517 display popup menu
        result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mPopupSys
    End Select
End Sub

Private Sub Form_Resize()
       'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
       'this removes the icon from the system tray
Dim result As Long
    With frmRouter
        .WindowState = vbNormal
        result = SetForegroundWindow(.hwnd)
        .Show
        .Visible = True
    End With
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mPopExit_Click()
       'called when user clicks the popup menu Exit command
    Unload Me
    Unload frmRouter
End Sub

Private Sub mPopRestore_Click()
    Call RestoreFromSysTray
End Sub

'This has to be a separate PUBLIC routine as it may need to be called when
'frmRouter is first loaded. This is because when frmRouter Loads
'it requires initially to create the form to the systray
Public Sub RestoreFromSysTray()
Dim result As Long
    With frmRouter
        .WindowState = vbNormal
        result = SetForegroundWindow(.hwnd)
        .Show
        .Visible = True
        .UpdateMshRows      'causes variable length to be 0
    End With
    Unload Me
End Sub

