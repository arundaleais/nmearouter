VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Begin VB.Form frmRouter 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "NMEA AIS Router"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmRouter.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer VDOTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2280
      Top             =   1080
   End
   Begin VB.Timer GraphTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   480
   End
   Begin VB.Timer ClearStatusBarTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   2880
      Top             =   480
   End
   Begin VB.Timer StatusBarTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   480
   End
   Begin VB.Timer ReconnectTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3960
      Top             =   480
   End
   Begin VB.Timer SpeedTimer 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   4560
      Top             =   480
   End
   Begin VB.Timer StatsTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5160
      Top             =   480
   End
   Begin VB.Timer FileInputTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   6960
      Top             =   480
   End
   Begin VB.Timer DequeTimeout 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   480
   End
   Begin VB.Timer PollTimer 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   6360
      Top             =   480
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5610
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbRouter 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   360
         Left            =   600
         MaskColor       =   &H00000000&
         TabIndex        =   5
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   4
         ToolTipText     =   "Pause"
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   360
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   3
         ToolTipText     =   "Start"
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.TextBox txtTerm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   9015
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   8400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSockets 
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   953
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRoutes 
      Height          =   540
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   953
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuFileProfileSave 
         Caption         =   "Profile Save"
      End
      Begin VB.Menu MenuFileProfileSaveAs 
         Caption         =   "Profile Save As"
      End
      Begin VB.Menu MenuFileInput 
         Caption         =   "TestFileInput"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileTest 
         Caption         =   "Test"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuConfigure 
      Caption         =   "Configure"
      Begin VB.Menu MenuProfile 
         Caption         =   "Profile"
         Begin VB.Menu MenuProfileNew 
            Caption         =   "New"
         End
         Begin VB.Menu MenuProfileOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu MenuProfileSave 
            Caption         =   "Save"
         End
         Begin VB.Menu MenuProfileSaveAs 
            Caption         =   "Save As"
         End
         Begin VB.Menu MenuProfileDelete 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu MenuConfigureVCP 
         Caption         =   "VCP"
         Begin VB.Menu MenuVCPCreate 
            Caption         =   "Create"
         End
         Begin VB.Menu MenuVCPRemove 
            Caption         =   "Remove"
         End
      End
      Begin VB.Menu MenuSocket 
         Caption         =   "Connection"
         Begin VB.Menu MenuSocketNew 
            Caption         =   "New"
         End
         Begin VB.Menu MenuSocketOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu MenuSocketDelete 
            Caption         =   "Delete"
         End
         Begin VB.Menu MenuSocketRecorder 
            Caption         =   "Recorder"
            Begin VB.Menu MenuSocketRecorderCreate 
               Caption         =   "Create"
            End
            Begin VB.Menu MenuSocketRecorderRemove 
               Caption         =   "Remove"
            End
         End
      End
      Begin VB.Menu MenuRoute 
         Caption         =   "Route"
         Begin VB.Menu MenuRouteNew 
            Caption         =   "New"
         End
         Begin VB.Menu MenuRouteOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu MenuRouteDelete 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu MenuConfigureFilter 
         Caption         =   "Filters"
      End
      Begin VB.Menu MenuConfigureGraph 
         Caption         =   "Graph"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuConfigureExport 
         Caption         =   "Export"
      End
   End
   Begin VB.Menu MenuView 
      Caption         =   "View"
      Begin VB.Menu MenuViewInout 
         Caption         =   "Input/Output"
         Begin VB.Menu MenuViewInoutData 
            Caption         =   "Data"
            Checked         =   -1  'True
         End
         Begin VB.Menu MenuViewInoutSockets 
            Caption         =   "Connections"
            Checked         =   -1  'True
         End
         Begin VB.Menu MenuViewInoutRoutes 
            Caption         =   "Routes"
            Checked         =   -1  'True
         End
         Begin VB.Menu MenuViewInoutGraph 
            Caption         =   "Graph"
         End
      End
      Begin VB.Menu MenuViewConnections 
         Caption         =   "Connections"
      End
      Begin VB.Menu MenuViewRoutes 
         Caption         =   "Routes"
      End
      Begin VB.Menu MenuViewForwarding 
         Caption         =   "Forwarding"
      End
      Begin VB.Menu MenuViewWinsock 
         Caption         =   "TCP/IP Sockets"
      End
      Begin VB.Menu MenuViewComm 
         Caption         =   "Serial Sockets"
      End
      Begin VB.Menu MenuViewFile 
         Caption         =   "File Sockets"
      End
      Begin VB.Menu MenuViewTTY 
         Caption         =   "Display Sockets"
      End
      Begin VB.Menu MenuViewLog 
         Caption         =   "Event Log"
      End
      Begin VB.Menu MenuViewRegistry 
         Caption         =   "Registry"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuHelpWeb 
         Caption         =   "Help File"
      End
      Begin VB.Menu MenuHelpRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "About"
      End
   End
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
Attribute VB_Name = "frmRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright@ 2009-17 Neal Arundale (Arundale of Scarborough Ltd)

'Declarations for Setting ShowInTaskbar programatically
'see http://www.vbforums.com/showthread.php?357131-Set-Get-the-ShowInTaskbar-property-for-a-window-at-runtime

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000
Const SW_HIDE = 0
Const SW_NORMAL = 1
'=================================
'http://www.developerfusion.com/code/1607/counting-lines-in-a-multiline-textbox/
Private Declare Function SendMessageAsLong Lib "user32" _
     Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long
Const EM_GETLINECOUNT = 186
'================================

'ArrangeControls variables
' The percentage occupied by each pane
Private PercentageVertical As Single
Private PercentageHorizontal As Single

Const hgt1Min = 560 'Sockets when blank
Const hgt2Min = 560 'Routes
Const hgt3Min = 315 'Data

Public Clients As New Collection
'Public NewProfile As String

Private SelIdx As Long  'The last selected Idx on mshSockets
Private SelIdx1 As Long 'The last selected Idx1 on mshRoutes
Private SelIdx2 As Long 'The last selected Idx2 on mshRoutes


Private Sub MenuAddins_Click()
'    If FileExists(Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com\setupc.exe") Then
'        MenuConfigureVCP.Enabled = False
'    Else
'        MenuConfigureVCP.Enabled = True
'    End If

End Sub

Private Sub CreateVCPDriver()
Dim ret As Long
Dim Version As String
Dim VCPInstalledDir As String
Dim SetupDir As String
Dim SetupFile As String
Dim kb As String
Dim UACValue As Long

'Created when com0com Driver is installed (by com0com)
    VCPInstalledDir = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\com0com", "Install_Dir")
    If VCPInstalledDir <> "" Then
'        If FileExists(VCPInstalledDir & "\setupc.exe") Then
            Exit Sub
'        End If
    End If

'    SetupDir = Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com"
    SetupDir = App.Path & "\com0com"
    Version = GetVersion1
'rename driver setup file dependant on driver required
    Select Case Version
        Case Is = "Windows 2000", "Windows XP", "Windows Server 2003"
            SetupFile = "setup_com0com-3.0.0.0-i386-and-x64-unsigned.exe"
        Case Is = "Windows Vista", "Windows 7", "Windows 8"
            If Is64bit Then
                SetupFile = "setup_com0com_W7_x64_signed.exe"
            Else
                SetupFile = "setup_com0com-3.0.0.0-i386-and-x64-unsigned.exe"
'                SetupFile = "setup_com0com_W7_x86_signed.exe"
            End If
    End Select

'possibility of XP 64 bit but driver would not be signed
    If Is64bit Then
        Version = Version & " (64 bit)"
    Else
        Version = Version & " (32 bit)"
    End If
    
    If Version = "" Then
        MsgBox "Sorry - no driver available for " & Version, vbExclamation, "See Help for more information"
        Exit Sub
    End If
    
    ret = MsgBox("You are using " & Version & vbCrLf _
    & "Using install file " & SetupFile & vbCrLf & vbCrLf _
    & "This add-in creates a com0com VCP driver" & vbCrLf _
    & "It is not required unless you want an Output of NmeaRouter to be" & vbCrLf _
    & "sent to the Serial (COM) input of a program running on this PC," & vbCrLf _
    & "and the program has NO provision to input networked data (TCP or UDP)" & vbCrLf _
    & vbCrLf & "See Help for more information" _
    , vbOKCancel, "Virtual Comm Port Driver")    ', "http://web.arundale.co.uk/docs/ais/nmearouter.html")
    If ret = vbOK Then

WriteLog "Creating VCP Driver from " & SetupDir & "\" & SetupFile
'Save and disable UAC & New Hardware
        Call DisableUAC
'Install com0com Driver
'        kb = "cmd.exe /C """ & SetupFile & " /S /D=%ProgramFiles%\Arundale\NmeaRouter\com0com"""
        kb = "cmd.exe /C """ & SetupFile & " /S /D=" & App.Path & "\com0com"""
        ret = ExecCmd(kb)
        VCPInstalledDir = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\com0com", "Install_Dir")
        If VCPInstalledDir = "" Then
WriteLog "VCP Driver failed (" & SetupDir & "\" & SetupFile & ")"
            MsgBox "Create VCP Driver Failed" & vbCrLf & ErrorDescriptionDLL
        Else
WriteLog "VCP Driver installed to " & VCPInstalledDir & "\" & SetupFile
        End If

'restore UAC & reenable New Hardware
        Call RestoreUAC
    End If
End Sub

Private Sub Form_Initialize()
'Stop
End Sub

Private Sub MenuConfigureFilter_Click()
    Filtercfg.Show
End Sub

Private Sub MenuFileProfileSave_Click()
    Call Profilecfg.ProfileSave
End Sub

Private Sub MenuFileProfileSaveAs_Click()
    Call Profilecfg.ProfileSaveAs
End Sub

Private Sub MenuHelp_Click()
    If CheckKey(ActivationKey) < 0 Then
        MenuHelpRegister.Caption = "Register"
    Else            'registered
        MenuHelpRegister.Caption = "Contact Me"
    End If
End Sub

Private Sub MenuHelpRegister_Click()
Dim SerialNo As Long
    If GetSerialNo > 0 Then
        Load frmRegister
        frmRegister.Show vbModal
    Else
        MsgBox "Cannot get device serial no" & vbCrLf _
        & "Please contact me" & vbCrLf, vbExclamation, "Device Error"
    End If
End Sub

Private Sub MenuHelpWeb_Click()
    Call HttpSpawn("http://www.nmearouter.com/docs/ais/nmearouter.html")
End Sub

Private Sub MenuHelpAbout_Click()
        MsgBox "NmeaRouter " & App.Major & "." & App.Minor & "." & App.Revision & " " & Chr$(10) & "Copyright " & Chr$(169) & " 2012-17 Neal Arundale"
End Sub

Private Sub cmdPause_Click()
Dim Hidx As Long
Dim Pause As Boolean

    If cmdPause.Caption = "Pause" Then
WriteLog "Pause Command"
        cmdPause.Caption = "Continue"
        Call PauseAllFileInput
    Else
WriteLog "Continue Command"
        cmdPause.Caption = "Pause"
        Call ContinueAllFileInput
    End If

End Sub

Public Sub cmdStart_Click()
Dim Idx As Long
    
'Debug.Print "Start Command"
WriteLog "Start Command"
    cmdStart.Enabled = False
'    cmdStop.Enabled = True
    cmdPause.Caption = "Pause"
    txtTerm.Text = ""
'You cannot create a new clsDuplicateFilter
'here because it will destroy the settings loaded from the profile
'Create the Forwarding First
'Then Open the Handler
    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).State <> -1 Then
            Call Routecfg.CreateSocketForwards(Idx)
'State must be set to 11 to prevent File reading first record until timer is started
            Sockets(Idx).State = 11 'Trying to open
            Call OpenHandler(Idx)
        End If
    Next Idx
'Only when all the handlers have been created you can start
'File Input, otherwise if done sequentially the handler
'to which the file input is being routed may not be open.
    
    If IsSerialHandlerInUse = True Then
        PollTimer.Enabled = True
    End If
    
    cmdStop.Enabled = True
    Call StatsTimer_Timer   'update stats
    SpeedTimer.Enabled = True
    StatsTimer.Enabled = True
    If MenuViewInoutGraph.Checked = True Then
        If ExcelOpen = False Then
            Call CreateWorkbook
        End If
        GraphTimer.Enabled = True
    End If

    For Idx = 1 To UBound(Sockets)
        If IsRecordable(Idx) = True Then
            Call CreateRecorder(Idx)
        End If
    Next Idx
    
'This call the reconnect timer in a separate thread, which will
'open any handlers (including the TTY windows)
    ReconnectTimer.Interval = 500
    ReconnectTimer.Enabled = True 'Cannot be done by stats timer
'Start the file Input by enabling the input timer
'This starts a new thread, otherwise cmdStart will not exit until
'all the file input is completed
    If IsFileHandlerOpenForInput = True Then
        Call ContinueAllFileInput
    End If
    

End Sub

Public Sub cmdStop_Click()
Dim Idx As Long

WriteLog "Stop Command"
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    cmdPause.Caption = "Continue"
    
'Stop this first so Close is not re-opened
    ReconnectTimer.Enabled = False
    PollTimer.Enabled = False
    GraphTimer.Enabled = False
    VDOTimer.Enabled = False
'We have to save the window positions as the positions are reset
'when CmdStart is clicked
'Stop
    Call SaveWindows
'Debug.Print Height
'Remove the Forwarding first
'Then Close the Handler
    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).State <> -1 Then
            Call Routecfg.RemoveSocketForwards(Idx)
            Call CloseHandler(Idx)
        End If
    Next Idx
    
'I dont think you want to reset
'    SourceDuplicateFilter.Reset
'cannotset to nothis (destroys the profile inf0
'    Set SourceDuplicateFilter = Nothing
    Call RemoveStreamSockets
    
'Leave the stats time running so we can check no processing is done
'    StatsTimer.Enabled = False
    SpeedTimer.Enabled = False
    Call SpeedTimer_Timer   'Update the Speed
'Call again to clear the speed
    Call SpeedTimer_Timer   'Clear the speed
    Call StatsTimer_Timer   'update stats

End Sub

' We open the port and get things started here
Private Sub Form_Load()
Dim i As Long
Dim Cancel As Boolean
Dim ProfileKey As String
Dim WindowKey As String
Dim KeyValue As String
Dim Position() As String

    Me.Icon = LoadPicture(NmeaRouterIcon)

'Get the position of frmRouter - this MUST be done while the form
'is loading to prevent a resize & to ensure the ShowinTheTaskBar
'executes correctly
    ProfileKey = ROUTERKEY & "\Profiles\" & CurrentProfile
    WindowKey = ProfileKey
    KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "Window")
    Position = Split(KeyValue, ",")
'Ensure registry key exists - wont on initial load
    If UBound(Position) >= 3 Then
'if you use with frmRouter the .height causes a resize
'This stop ShowinTheTaskBar working properly
            
'If position (0) = -48000 then its been saved minimized
'in the registry. This will cause it to be restored to the
'Task bar and you cannot then change it back to "normal"
            If Position(0) > 0 Then
                Left = Position(0)
                Top = Position(1)
                Width = Position(2)
                Height = Position(3)   'causes a resize
            End If
    End If
    With mshSockets
        .AllowUserResizing = flexResizeColumns
        .FormatString = "<Port|<Connection|^Direction|<IP Address : Port|>S/min|^Protocol|^Enabled|>Sentences|^Graph|Idx|Hidx|^State|<Status"
        .ColWidth(0) = 0  'Port
        .ColWidth(1) = 1300 '<Comments
        .ColWidth(2) = 1200  'Direction
        .ColWidth(3) = 1700 'Address
        .ColWidth(4) = 600  'Speed
        .ColWidth(5) = 1000  'Protocol
        .ColWidth(6) = 800  'Enabled
        .ColWidth(7) = 1000 'Count
        .ColWidth(8) = 800 'Graph
If jnasetup = True Then
        .ColWidth(9) = 500  'Socket
        .ColWidth(10) = 500 'Handler Index
Else
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0  'State
End If
        .ColWidth(12) = 4000  'Status
        .Rows = 2
        .Width = 0
        For i = 0 To .Cols - 1
            .Width = .Width + .ColWidth(i) + 15
        Next i
    End With
    With mshRoutes
        .AllowUserResizing = flexResizeColumns
        .FormatString = "Route between|^Direction|^and|Enabled|Idx1|Idx2|Ridx"
        .ColWidth(0) = 1500  'Connection 1
        .ColWidth(1) = 700  'Direction
        .ColWidth(2) = 1500  'Connection 2
        .ColWidth(3) = 1000  'Enabled
If jnasetup = True Then
        .ColWidth(4) = 500  'Idx1
        .ColWidth(5) = 500  'Idx2
        .ColWidth(6) = 500  'Ridx
Else
        .ColWidth(4) = 0  'Idx1
        .ColWidth(5) = 0  'Idx2
        .ColWidth(6) = 0  'Ridx
End If
        .Rows = 2
        .Width = 0
        For i = 0 To .Cols - 1
            .Width = .Width + .ColWidth(i) + 15
        Next i
    End With
    
    
    With StatusBar
        .Panels(1).AutoSize = sbrContents
        .Panels(1).Bevel = sbrNoBevel
'        .Panels(1).Text = "This is the status bar"
'When reducing height keep status bar to the front
        .ZOrder 0
    End With
    
'frmRouter must have ShowinTheTaskBar set to True
'in Design Properties
'Always minimise to task bar - then resore it when profile is
'loaded (if required)
    If cmdSysTray = True Then
        SetWindowLong Me.hwnd, GWL_EXSTYLE, (GetWindowLong(hwnd, _
        GWL_EXSTYLE) And Not WS_EX_APPWINDOW)
        Me.Hide
        Me.WindowState = vbMinimized
     End If
    
'Sort out ShowinTheTaskBar while the form is loading
'and before anything else calls ArrangeControls
    If CurrentProfile = "" Then
'Here when first Run (no Registry)
        CurrentProfile = "Profile1"
    End If
    
'Check if installer not including com0com
'If Dir(Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com", vbDirectory) = "" Then
If Dir(App.Path & "\com0com", vbDirectory) = "" Then
    MenuConfigureVCP.Visible = False
    MenuConfigureVCP.Enabled = False
End If

'You still have to load the profile to create the initial profile
'and start correctly
'v59    Call LoadProfile(CurrentProfile, Cancel)

End Sub


Public Function ResetDisplay()
Dim Row As Long
Dim Col As Long

    With mshSockets
        Do Until .Rows = .FixedRows + 1
            .RemoveItem (.FixedRows)    'base 0
        Loop
'.Row & .Col must be set to clear the back color
        .Col = .FixedCols   'if no fixed start at col 0
        .Row = .Rows - .FixedRows   'Must set to clear cellbackcolor
        Do Until .Col = .Cols - .FixedCols  'Exit if only Col is fixed col
            .TextMatrix(.Rows - .FixedRows, .Col) = ""
            .CellBackColor = vbWhite
            If .Col = .Cols - 1 Then Exit Do    'Last Column
            .Col = .Col + 1                     '
        Loop
    End With
    
    With mshRoutes
        Do Until .Rows = .FixedRows + 1
            .RemoveItem (.FixedRows)    'base 0
        Loop
'.Row & .Col must be set to clear the back color
        .Col = .FixedCols   'if no fixed start at col 0
        .Row = .Rows - .FixedRows   'Must set to clear cellbackcolor
        Do Until .Col = .Cols - .FixedCols  'Exit if only Col is fixed col
            .TextMatrix(.Rows - .FixedRows, .Col) = ""
            .CellBackColor = vbWhite
            If .Col = .Cols - 1 Then Exit Do    'Last Column
            .Col = .Col + 1                     '
        Loop
    End With
    
    With StatusBar
    End With
    txtTerm.Text = ""
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dim ret As Integer
'Dim Reason As String

'    ret = vbOK      'default is Quit
'Select Case UnloadMode
'Case Is = vbFormControlMenu
'    Reason = "User clicked close(X)"
'    ret = MsgBox("Do you wish to Quit NmeaRouter", vbOKCancel, "NmeaRouter")
'Case Is = vbFormCode
'    Reason = "Unload invoked from code"
'Case Is = vbAppWindows
'    Reason = "Operationg environment session ending"
'Case Is = vbAppTaskManager
'    Reason = "Task Manager"
'Case Is = vbFormMDIForm
'    Reason = "MDI parent closing child"
'Case Is = vbFormOwner
'    Reason = "Owner closing"
'Case Else
'    Reason = "Unknown"
'End Select
'    WriteLog (Me.Name & ".Form_QueryUnload (UnloadMode=" & Reason & ")")
    If QueryQuit(UnloadMode) = vbOK Then
        Call CloseProfile(CurrentProfile)
'MUST NOT call END to exit program otherwise
'Exit in Batch or Task manager cannot force program exit
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    ArrangeControls
End Sub

'This is only used to test input from a file
'In particular Control characters
'Use Menu Editor to enable and make MenuFileInput Visible
Private Sub MenuFileInput_Click()
Dim FileCh As Integer
Dim FileName As String
Dim kb As String
Dim i As Long
Const TestComm = False  'True to test Comm
Const TestForwardData = True    'True to test a Connection
Const TestInputConnection = 7   'Set to IDX of connection to test
    If TestComm Then
        If Not (Comms(0) Is Nothing) Then
            Comms(0).CommOutput (Date & Time & vbCrLf)
        End If
    End If
    
If True Then
    FileCh = FreeFile
'Set to file name containg the test sentences
'Each sentence must be terminated with <013><010> Must be 3 digits
'to replace characters with a single byte character code
    FileName = Environ("APPDATA") & "\Arundale\NmeaComm\TestFile.txt"
    Open FileName For Input As #FileCh
    Do Until EOF(FileCh)
        Line Input #FileCh, kb
        i = 1
        Do
            i = InStr(i, kb, "<")
            If i > 0 And Mid$(kb, i + 4, 1) = ">" Then
                kb = Replace(kb, Mid$(kb, i, 5), Chr$(CInt(Mid$(kb, i + 1, 3))))
            End If
            i = i + 1 'skip < as no > 4 characters on
        Loop Until i = 1
        If TestComm Then
            If Not (Comms(0) Is Nothing) Then
                Comms(0).CommOutput (kb & vbCrLf)
            End If
        End If
        If TestForwardData Then
            Call ForwardData(kb, TestInputConnection)
'Stop
        End If
    Loop
    Close #FileCh  'old file being used as template
End If

End Sub

Private Sub MenuConfigure_Click()
'You cant add or delete a Route if only 1 socket
'If there are 2 sockets you must be able to either
'Add or Delete a route
    If SocketCount < 2 Then
        MenuRoute.Enabled = False
    Else
        MenuRoute.Enabled = True
    End If
End Sub

Private Sub MenuConfigureGraph_Click()
    Call Graphcfg.GraphOpen
End Sub

Private Sub MenuConfigureExport_Click()
Dim ret As Integer
    ret = MsgBox("All your settings will be exported to :-" & vbCrLf _
    & """" & TempPath & "Router.reg" & """" & vbCrLf _
    & "Please note this is a .reg file." & vbCrLf _
    & "If you click on the Exported file, you will asked if you wish" & vbCrLf _
    & "to run the Windows Registry Editor." & vbCrLf _
    & "If you do, all existing settings will be overwritten with the " _
    & "Exported settings", vbOKCancel, "Export Settings")
    If ret = vbOK Then
        Call OutputAllKeys(ROUTERKEY, TempPath & "Router.reg")
    End If

End Sub

Private Sub MenuFileTest_Click()
Dim Cancel As Boolean
Dim i As Long
Dim j As Long
Dim start As Long
Dim kb As String
Dim Idx As Long

'Stop
    If CurrentProfile <> "TestInput" Then
        Call CloseProfile(CurrentProfile)
        CurrentProfile = "TestInput"
        Idx = 1
        Sockets(Idx).DevName = "TestInput"
        Sockets(Idx).Direction = 0
        Sockets(Idx).Enabled = True
        Sockets(Idx).Handler = 3
        Call OpenHandler(Idx)
    Else
        Idx = DevNameToSocket(CurrentProfile)
    End If
    frmRouter.PollTimer.Enabled = True
    start = 32
    For i = 1 To 2000
        kb = i & vbTab
        For j = 1 To 70
            kb = kb & Chr$(start + j)
        Next j
        Call ForwardData(kb & vbCrLf, Idx)
        DoEvents
'        Start = Start + 1
        If start > 125 - 70 Then start = 32
    Next i
'    frmRouter.PollTimer.Enabled = True
'    Call Comms(Sockets(Idx).Hidx).CommOutput("")
'    Call CloseHandler(Idx)

End Sub

Private Sub MenuProfile_Click()
Dim ProfileCount As Long
Dim Profiles() As Variant

    Call Profilecfg.GetProfiles(ProfileCount, Profiles)
    If ProfileCount < MAX_PROFILES Then
        MenuProfileNew.Enabled = True
        MenuProfileSaveAs.Enabled = True
    Else
        MenuProfileNew.Enabled = False
        MenuProfileSaveAs.Enabled = False
    End If
    If ProfileCount > 0 Then
        MenuProfileOpen.Enabled = True
        MenuProfileSave.Enabled = True
        If ProfileCount < MAX_PROFILES Then
            MenuProfileSaveAs.Enabled = True
        Else
            MenuProfileSaveAs.Enabled = False
        End If
        MenuProfileDelete.Enabled = True
    Else
        MenuProfileOpen.Enabled = False
        MenuProfileSave.Enabled = False
        MenuProfileSaveAs.Enabled = True
        MenuProfileDelete.Enabled = False
    End If
End Sub

Private Sub MenuProfileNew_Click()
    Call Profilecfg.ProfileNew
    Call MakeFormsVisisble
    Call UpdateMshRows      'causes variable length to be 0
End Sub

Private Sub MenuProfileOpen_Click()
'    Loading = True
    Call Profilecfg.ProfileOpen
    Call MakeFormsVisisble
'    Loading = False
    Call UpdateMshRows      'causes variable length to be 0
End Sub

Private Sub MenuProfileSave_Click()
    Call Profilecfg.ProfileSave
End Sub

Private Sub MenuProfileSaveAs_Click()
    Call Profilecfg.ProfileSaveAs
End Sub

Private Sub MenuProfileDelete_Click()
    Call Profilecfg.ProfileDelete
    Call UpdateMshRows      'causes variable length to be 0
End Sub

Private Sub MenuRoute_Click()
'    If IsNewRoute < MAX_ROUTES Then
    If IsNewRoute = True Then
        MenuRouteNew.Enabled = True
    Else
        MenuRouteNew.Enabled = False
    End If
        
    If RouteCount > 0 Then
        MenuRouteDelete.Enabled = True
        MenuRouteOpen.Enabled = True
    Else
        MenuRouteDelete.Enabled = False
        MenuRouteOpen.Enabled = False
    End If
End Sub

Private Sub MenuRouteNew_Click()
    Call Routecfg.RouteNew
    Call UpdateMshRows
End Sub

Private Sub MenuRouteOpen_Click()
    Call Routecfg.RouteOpen(SelIdx1, SelIdx2)
    Call UpdateMshRows
End Sub

Private Sub MenuRouteDelete_Click()
    Call Routecfg.RouteDelete(SelIdx1, SelIdx2)
    Call UpdateMshRows
End Sub

Private Sub MenuSocket_Click()
    If SocketCount < MAX_SOCKETS Then
        MenuSocketNew.Enabled = True
    Else
        MenuSocketNew.Enabled = False
    End If
    If SocketCount > 0 Then
        MenuSocketOpen.Enabled = True
        MenuSocketDelete.Enabled = True
    Else
        MenuSocketOpen.Enabled = False
        MenuSocketDelete.Enabled = False
    End If
    
'Assume cant record, must have a selidx and be recordable
    MenuSocketRecorder.Enabled = False
    If SelIdx > 0 Then
        If IsRecordable(SelIdx) Then
            MenuSocketRecorder.Enabled = True
        End If
    End If
End Sub

Private Sub MenuSocketNew_Click()
    Call Socketcfg.SocketNew
    Call UpdateMshRows
    ReconnectTimer.Enabled = True
End Sub

Private Sub MenuSocketOpen_Click()
'Stop reconnecting while editing socket
    ReconnectTimer.Enabled = False
    Call Socketcfg.SocketOpen(SelIdx)
    Call UpdateMshRows
    ReconnectTimer.Enabled = True
End Sub

Private Sub MenuSocketDelete_Click()
'Stop reconnecting while editing socket
    ReconnectTimer.Enabled = False
    Call Socketcfg.SocketDelete(SelIdx)
    Call UpdateMshRows
    ReconnectTimer.Enabled = True
End Sub


Private Sub MenuSocketRecorder_Click()
'Should never be 0
    If SelIdx = 0 Then Exit Sub
'Toggle
    With Sockets(SelIdx)
'Can we enable create
        If .Recorder.Output Is Nothing Then
            If IsRecordable(SelIdx) Then
                MenuSocketRecorderCreate.Enabled = True
                MenuSocketRecorderRemove.Enabled = False
            End If
        Else
'or remove
            MenuSocketRecorderRemove.Enabled = True
            MenuSocketRecorderCreate.Enabled = False
        End If
    End With
End Sub

Private Sub MenuSocketRecorderCreate_Click()
'only created if it does not exist
    Sockets(SelIdx).Recorder.Enabled = True
    Call CreateRecorder(SelIdx)
End Sub

Private Sub MenuSocketRecorderRemove_Click()
    Call RemoveRecorder(SelIdx)
End Sub

Private Sub MenuViewComm_Click()
    Call DisplayComms
End Sub

Private Sub MenuViewConnections_Click()
    Call DisplaySockets
End Sub

Private Sub MenuViewFile_Click()
    Call DisplayFiles
End Sub

Private Sub MenuViewForwarding_Click()
    Call DisplayForwarding
End Sub

Private Sub MenuViewInoutData_Click()
Static SavedHeight As Single

'Set inout display window to the new setting menu
    MenuViewInoutData.Checked = Not MenuViewInoutData.Checked
    With txtTerm
        .Enabled = MenuViewInoutData.Checked
'first to avoid flicker
'We have to save the height so if we re-enable its restored to the same height
        If .Enabled = True Then
            .Height = SavedHeight
            .BackColor = vbWhite
        Else
            SavedHeight = .Height
            .Height = 0
            .BackColor = vbRed
        End If
    End With
    Call UpdateMshRows
End Sub

Private Sub MenuViewInoutGraph_Click()
    MenuViewInoutGraph.Checked = Not MenuViewInoutGraph.Checked
    MenuConfigureGraph.Enabled = MenuViewInoutGraph.Checked
'Dont try and config graph if myBook not set
    If MenuViewInoutGraph.Checked = True Then
        GraphTimer.Enabled = True
        If ExcelUpdateInterval > 0 Then
            Call CreateWorkbook
            GraphTimer.Enabled = True
        Else
            MsgBox "Graphs will not be created until the" & vbCrLf _
            & "Graph Update Interval is set above Zero minutes" & vbCrLf & vbCrLf _
            & "Use Configure > Graphs to set the Update Inteval" & vbCrLf, , "Enable Graphs"
        End If
    Else
        Call CloseExcel
    End If
End Sub

Private Sub MenuViewInoutRoutes_Click()
Static SavedHeight As Single
    MenuViewInoutRoutes.Checked = Not MenuViewInoutRoutes.Checked
    With mshRoutes
        .Enabled = MenuViewInoutRoutes.Checked
'first to avoid flicker
'We have to save the height so if we re-enable its restored to the same height
        If .Enabled = True Then
            .Height = SavedHeight
            .BackColor = vbWhite
        Else
            SavedHeight = .Height
            .Height = 0
            .BackColor = vbRed
        End If
    End With
    Call UpdateMshRows
    Call ArrangeControls
End Sub

Private Sub MenuViewInoutSockets_Click()
Static SavedHeight As Single
'Set inout display window to the new setting menu
    MenuViewInoutSockets.Checked = Not MenuViewInoutSockets.Checked
    With mshSockets
        .Enabled = MenuViewInoutSockets.Checked
'first to avoid flicker
'We have to save the height so if we re-enable its restored to the same height
        If .Enabled = True Then
            .Height = SavedHeight
            .BackColor = vbWhite
        Else
            SavedHeight = .Height
            .Height = 0
            .BackColor = vbRed
        End If
    End With
    Call UpdateMshRows
End Sub

Private Sub MenuViewLog_Click()
    Call CloseLog
    Shell "notepad " & LogFileName, vbNormalFocus
End Sub

Private Sub MenuViewRegistry_Click()
    frmRegistry.Show
    Call OutputAllKeys(ROUTERKEY)
End Sub

Private Sub MenuViewRoutes_Click()
    Call DisplayRoutes
End Sub

Private Sub MenuViewTTY_Click()
    Call DisplayTTYs
End Sub

Private Sub MenuViewWinsock_Click()
    Call DisplayWinsock
End Sub

Private Sub MenuFileExit_Click()
    Unload Me
End Sub

'This is event driven when data is received
'The Data MUST be buffered in this socket before being
'sent to be displayed or forwarded
Public Sub CommInput(thiscomm As clsComm, commdata As String)
    Dim cpos%                            '
    
'If Left$(commdata$, 1) = Chr$(10) Then
'Debug.Print "LF at left"
'End If
'Add on any previously received partial sentence
    commdata = thiscomm.PartSentence & commdata
    thiscomm.PartSentence = ""
    
    If commdata <> "" Then
'commdata$ always has a NULL appended by CommRead
' Substitute the CR with a CRLF pair, dump the LF
        Do Until Len(commdata$) = 0
            cpos% = InStr(commdata$, Chr$(13))
            If cpos% > 0 Then   'Complete sentence
'chrctrl added to try and see if any non ascii characters are being output
                Call CommRcv(Left$(commdata$, cpos% - 1) & vbCrLf, thiscomm.hIndex)
                commdata$ = Mid$(commdata$, cpos% + 1)
                cpos% = InStr(commdata$, Chr$(10))
                If cpos% > 0 Then
                    commdata$ = Mid$(commdata$, cpos% + 1)
                End If
             Else           'No CR
                cpos% = InStr(commdata$, Chr$(10))
'We probably ought to replace a LF Null with CRLF
                If cpos% > 0 Then   'But has LF, Keep LF + NULL
                    commdata$ = Mid$(commdata$, cpos% + 1)
                Else                'No CR or LF but will have last NULL
'save the partial sentence for this socket
                    thiscomm.PartSentence = thiscomm.PartSentence _
                    & Left$(commdata$, Len(commdata$) - 1)
                    commdata$ = ""
                End If
            End If
        Loop
    End If
End Sub

'When we receive the data we want to add the source into the call
'If we add it as a long (the socket) we can look up the naqme later
'If would be good to add the destinations as an array, but
'we cant if were going to put the received messages into a buffer.
'Source is the Comm Source Index (hIndex)
Public Sub CommRcv(commdata As String, Source As Long)
Dim bytestoforward As Long

'commdata$ has a NULL appended at end of last input
        If Right$(commdata$, 1) = Chr$(0) Then
            bytestoforward = Len(commdata$) - 1
        Else
            bytestoforward = Len(commdata$)
        End If
        If bytestoforward > 0 Then
            If Not Comms(Source) Is Nothing Then
                Call ForwardData(Left$(commdata$, bytestoforward), Comms(Source).sIndex)
            End If
        End If
'    Call ForwardData(commdata, Source)
End Sub

Private Sub MenuVCPCreate_Click()
Dim ret As Long
Dim ret1 As Long
Dim kb As String
Dim Idx As Long
Dim ExistingVCPs() As String
Dim NewVCPs() As String
Dim NewVCP As String
Dim i As Long
Dim j As Long

    Call CreateVCPDriver    'wont re-create if it already exists
    WriteLog ("Creating VCP Port")
    ExistingVCPs = GetVCPs
    Call DisableUAC
'Will ask for Re-boot if any com0com port is open by NmeaRouter
'--output  %Temp%\com0com.log
        kb = "cmd.exe /C ""setupc.exe --output  %Temp%\com0com.log install PortName=- PortName=COM#"""
        ret = ExecCmd(kb)
        Call AddFileToLog(Environ("Temp") & "\com0com.log")
        If ret <> 0 Then
            MsgBox "Create VCP Port Failed" & vbCrLf & ErrorDescriptionDLL
        End If
'You need to do the update after the install otherwise the port is disabled
'in the device manager
    WriteLog ("Updating VCP Ports")
        kb = "cmd.exe /C ""setupc.exe --output  %Temp%\com0com.log update"""
        ret1 = ExecCmd(kb)
        Call AddFileToLog(Environ("Temp") & "\com0com.log")
        If ret <> 0 Then
            MsgBox "Updating VCP Ports Failed" & vbCrLf & ErrorDescriptionDLL
        End If
        Call RestoreUAC
        
'Need to update (for stats timer, if VCP is already on this profile)
        For Idx = 1 To UBound(Sockets)
              Sockets(Idx).Comm.VCP = GetVCP(Sockets(Idx).Comm.Name)
        Next Idx
        
                
        If ret = 0 And ret1 = 0 Then
            NewVCPs = GetVCPs
            NewVCP = ""
            For i = 0 To UBound(NewVCPs)
                For j = 0 To UBound(ExistingVCPs)
                    If NewVCPs(i) = ExistingVCPs(j) Then Exit For
                Next j
                If j > UBound(ExistingVCPs) Then
                    NewVCP = NewVCP + NewVCPs(i) & " "
                End If
            Next i
            MsgBox "New VCP " & NewVCP & "created"
        End If
End Sub

Private Sub MenuVCPRemove_Click()
    Call VCPcfg.VCPRemove
End Sub

Private Sub mshRoutes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim kb As String
Dim Idx1 As Long
Dim Idx2 As Long
Dim Ridx As Long
Dim SaveStatsTimer As Boolean

#If False Then
Dim iRow As Long
Dim iColumn As Long
Dim iMaxColumns As Long
Dim iMaxRows As Long
Dim iTotal As Long
#End If

'Stop the stats timer updating the display (it changes the focus rectangle
'to update the colours)

SaveStatsTimer = StatsTimer.Enabled
StatsTimer.Enabled = False
With mshRoutes
    
#If False Then
    iMaxColumns = .Cols - 1
    iMaxRows = .Rows - 1
     
    For iColumn = 0 To iMaxColumns
        iTotal = iTotal + .ColWidth(iColumn)
        If iTotal > X Then ' found it
            .Col = iColumn
            Exit For
        End If
    Next
    iTotal = 0
    For iRow = 0 To iMaxRows
        iTotal = iTotal + .RowHeight(iRow)
        If iTotal > Y Then
            .Row = iRow
            Exit For
        End If
    Next
#End If
    .Row = .MouseRow    'V57
    .Col = .MouseCol
'Check its not the top row that has been selected
    .FocusRect = flexFocusNone ' (The selected cell changes to blue)
'must set defaults
    SelIdx1 = 0
    SelIdx2 = 0
    Ridx = 0
    If .Row > 0 Then
'Skip to next row if typemismatch because row are blank
        On Error Resume Next
        SelIdx1 = .TextMatrix(.Row, 4)
        SelIdx2 = .TextMatrix(.Row, 5)
        Ridx = .TextMatrix(.Row, 6)
        On Error GoTo 0
    End If
    
'OK to make the changes to Routes
    If Button = vbRightButton Then
            Select Case .TextMatrix(0, .Col)
            Case Is = "Route between"
'Requires SelIdx1 & SelIdx2 set to get default
                Call MenuRoute_Click
                PopupMenu frmRouter.MenuRoute()
            Case Is = "Enabled"
                If SelIdx1 > 0 And SelIdx2 > 0 And Ridx > 0 Then
                    Call ToggleRouteEnabled(SelIdx1, SelIdx2, Ridx)
                Else
                    For i = 1 To .Rows - 1
                        Idx1 = .TextMatrix(i, 4)
                        Idx2 = .TextMatrix(i, 5)
                        Ridx = .TextMatrix(i, 6)
                        Call ToggleRouteEnabled(Idx1, Idx2, Ridx)
                    Next i
                End If
            End Select
    End If
    
    .FocusRect = flexFocusLight ' (gets changed when the grid is updated)
    StatsTimer.Enabled = SaveStatsTimer
End With
End Sub

Private Sub ToggleRouteEnabled(Idx1 As Long, Idx2 As Long, Ridx As Long)
Dim kb As String
'Toggle enabled for this route
    If Idx1 > 0 And Idx2 > 0 And Ridx > 0 Then
        Sockets(Idx1).Routes(Ridx).Enabled _
        = Not Sockets(Idx1).Routes(Ridx).Enabled
'Reset the forwards
        If Sockets(Idx1).Routes(Ridx).Enabled = True Then
            Call Routecfg.CreateRouteForwards(Idx1, Idx2)
        Else
            Call Routecfg.RemoveRouteForwards(Idx1, Idx2)
        End If
    End If
End Sub

'If Sub MouseDown is defined, MouseUp is not triggered
'Private Sub mshSockets_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'With mshSockets
'MsgBox .MouseRow & " " & .MouseCol, , "Up"
'End With
'End Sub

Private Sub mshSockets_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim SaveStatsTimer As Boolean

#If False Then
Dim iRow As Long
Dim iColumn As Long
Dim iMaxColumns As Long
Dim iMaxRows As Long
Dim iTotal As Long
#End If

'Stop the stats timer updating the display (it changes the focus rectangle
'to update the colours)
SaveStatsTimer = StatsTimer.Enabled
StatsTimer.Enabled = False  'Must stop mouse moving
    
'Mouse col is where it is at that instant, if you move
'the mouse to (in VBE) the mouse col will differ
'to where what it was at the instant of the click event

With mshSockets
#If False Then
    iMaxColumns = .Cols - 1
    iMaxRows = .Rows - 1
     
    For iColumn = 0 To iMaxColumns
        iTotal = iTotal + .ColWidth(iColumn)
        If iTotal > X Then ' found it
            .Col = iColumn
            Exit For
        End If
    Next
    iTotal = 0
    For iRow = 0 To iMaxRows
        iTotal = iTotal + .RowHeight(iRow)
        If iTotal > Y Then
            .Row = iRow
            Exit For
        End If
    Next
#End If
    .Col = .MouseCol
    .Row = .MouseRow
    .FocusRect = flexFocusNone ' (The selected cell changes)
'Check its not the top row of socket that has been selected
    If IsNumeric(.TextMatrix(.Row, 9)) Then
        SelIdx = .TextMatrix(.Row, 9)
    Else
        SelIdx = 0      'no Socket (Could be top row)
    End If
    If Button = vbRightButton Then
            Select Case .TextMatrix(0, .Col)
            Case Is = "Connection"
'Requires SelIdx set to get default
                PopupMenu frmRouter.MenuSocket()
            Case Is = "Enabled"
                If SelIdx > 0 Then
                    Call ToggleSocketEnabled(SelIdx)
                Else
                    For i = 1 To UBound(Sockets)
                        Call ToggleSocketEnabled(i)
                    Next i
                End If
            Case Is = "Graph"
                If SelIdx > 0 Then
                    Call ToggleGraphEnabled(SelIdx)
                Else
                    For i = 1 To UBound(Sockets)
                        Call ToggleGraphEnabled(i)
                    Next i
                End If
            Case Is = "Sentences"
                If SelIdx > 0 Then      'Clear this socket only
                    Sockets(SelIdx).MsgCount = 0
                        Call ClearGraphMsgCount(SelIdx)
                Else                    'Clear all socket counts
                    For i = 1 To UBound(Sockets)
                        Sockets(i).MsgCount = 0
                        Call ClearGraphMsgCount(i)
                    Next i
                End If
            End Select
    End If
    
    .FocusRect = flexFocusLight ' (gets changed when the grid is updated)
End With

SelIdx = 0  'Prevent recorder being displayed on Configure > Connections
StatsTimer.Enabled = SaveStatsTimer
End Sub

'We dont actually need to stop the forwarding but if we dont
'the Output Queue fills and subsequent sentences are lost
'No damage is done, it just takes more processing time
'If we do remove the forwards then we must create them when
'the socket is re-opened
Private Sub ToggleSocketEnabled(Idx As Long)

    Sockets(Idx).Enabled = Not Sockets(Idx).Enabled
    If Sockets(Idx).Enabled = True And cmdStop.Enabled = True Then
        Call Routecfg.CreateSocketForwards(Idx)
        Call OpenHandler(Idx)
        If Sockets(Idx).Handler = 2 And Sockets(Idx).Direction = 1 Then
'Update the stats Now because as soon as we start reading the input
'file there will be a delay in showing the file is open

            Call StatsTimer_Timer
            Call RestartFileInput(CInt(Sockets(Idx).Hidx))

'Restart file input
        End If
'Check if TTY has been created (cant make visible when Modal form is displayd)
'This happens when Socketcfg is displayed
        Call MakeFormsVisisble
    Else
        Call Routecfg.RemoveSocketForwards(Idx)
        Call CloseHandler(Idx)
        
    End If
End Sub

Private Sub ToggleGraphEnabled(Idx As Long)

    Sockets(Idx).Graph = Not Sockets(Idx).Graph
    Call SyncGraphEnabled(Idx)
End Sub

#If V44 Then
Private Sub mshSockets_MouseMove_old(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With mshSockets
    Select Case .MouseRow
    Case Is = 0
        Select Case .MouseCol
        Case Is = 6
            .ToolTipText = "Right Click to toggle all Enable/Disable Connections"
        Case Is = 7
            .ToolTipText = "Right Click to Clear all Connection counters"
        Case Else
If jnasetup Then
            .ToolTipText = .MouseRow & ":" & .MouseCol
End If
        End Select
    Case Else
        Select Case .MouseCol
        Case Is = 1
            .ToolTipText = "Right Click to Configure Connection"
        Case Is = 6
            .ToolTipText = "Right Click to toggle Enable/Disable Connection"
        Case Is = 7
            .ToolTipText = "Right Click to Clear Connection counter"
        Case Else
If jnasetup Then
            .ToolTipText = .MouseRow & ":" & .MouseCol
End If
        End Select
    End Select
    End With
End Sub
#End If

'Poll for each socket in turn (55ms)
'It is probably more efficient to poll as use an event driven
'process if we are receiving at least
'20 sentences a minute, and if we are not the processor
'will not be loaded much anyway.
Private Sub PollTimer_Timer()
Dim Hidx As Integer
Dim kb As String

'If were not running do not reconnect
    If cmdStart.Enabled = True Then
        PollTimer.Enabled = False
        Exit Sub
    End If

'If comm has not yet been dimensioned we skip polling
    On Error GoTo Timer_error
    For Hidx = 1 To UBound(Comms)
        If Not Comms(Hidx) Is Nothing Then
'If Idx = 2 Then Stop
'        kb = kb & Idx & ":" & Sockets(Idx).State & " "
'Hidx may not yet have been set
            If Comms(Hidx).sIndex > 0 Then
                If Sockets(Comms(Hidx).sIndex).State > 0 Then
                    Comms(Hidx).Poll
                End If
            End If
        End If
    Next Hidx
Exit Sub

Timer_error:
'v45 changed from Sockets(Comms(Hidx).sIndex) to Hidx as Sockets() may not exist
    MsgBox "Poll Timer Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "On Comm (" & Hidx & ")", , "Poll Timer"
End Sub




Private Sub Winsock_ConnectionRequest(HidxListen As Integer, ByVal requestID As Long)
Dim IdxListen As Long
Dim IdxStream As Long
Dim HidxStream As Long
Dim Hidx As Long
Dim StreamCount As Long
Dim PortCount As Long
Dim IPCount As Long
Dim IPListen As String
Dim PortListen As Long
Dim ctrl As Winsock
Dim IPPort As String    '123.456.789.123:12345
Dim MaxPermitted As Long
Dim WinsockCount As Long

'Check we have not exceeded stream quota for this IP
    IdxListen = Winsocks(HidxListen).sIndex
'    IPListen = Sockets(IdxListen).Winsock.RemoteHostIP

'Get no of Streams, Ports & IP:Ports already in use
    With Winsock(HidxListen)
        If .State = sckListening Then
'Returns last 2 characters missing
            IPListen = .RemoteHostIP
            PortListen = .LocalPort
            For Each ctrl In Winsock
'Debug.Print ctrl.Index & ":" & ctrl.State
                If ctrl.SocketHandle > -1 Then  'allocared
                    WinsockCount = WinsockCount + 1
                    If ctrl.Protocol = sckTCPProtocol And ctrl.State <> sckListening Then
                        StreamCount = StreamCount + 1
                        If ctrl.LocalPort = PortListen Then
                            PortCount = PortCount + 1
                            If ctrl.RemoteHostIP = IPListen Then
                                IPCount = IPCount + 1
                            End If
                        End If
                    End If
                End If
If ctrl.Protocol = sckTCPProtocol Then
'    Debug.Print "SC:" & ctrl.RemoteHostIP & ":" & ctrl.LocalPort & "-" & ctrl.State
'Stop
End If
            Next ctrl
            
            IPPort = IPListen & ":" & PortListen
            If IPCount >= Sockets(IdxListen).Winsock.PermittedIPStreams Then
                MaxPermitted = Sockets(IdxListen).Winsock.PermittedIPStreams
                Call RejectConnection(IPPort, MaxPermitted, 1)
'                WriteLog "No free connections to Client " _
'                & IPPort & ", limit is " _
'                & Sockets(IdxListen).Winsock.PermittedIPStreams
                Exit Sub
            End If
        
            If PortCount >= Sockets(IdxListen).Winsock.PermittedStreams Then
                MaxPermitted = Sockets(IdxListen).Winsock.PermittedStreams
                Call RejectConnection(IPPort, MaxPermitted, 2)
'                WriteLog "No free Concurrent Clients " _
'                & IPPort & ", limit is " _
'                & Sockets(IdxListen).Winsock.PermittedStreams
                Exit Sub
            End If
        
            If StreamCount >= MAX_TCPSERVERSTREAMS Then
                MaxPermitted = MAX_TCPSERVERSTREAMS
                Call RejectConnection(IPPort, MaxPermitted, 3)
'                WriteLog "No free TCP Server Streams " _
'                & IPPort & ", limit is " _
'                & MAX_TCPSERVERSTREAMS
                Exit Sub
            End If
            
            If WinsockCount >= MAX_WINSOCKS Then
                MaxPermitted = MAX_WINSOCKS
                Call RejectConnection(IPPort, MaxPermitted, 4)
'                WriteLog "No free TCP Server Streams " _
'                & IPPort & ", limit is " _
'                & MAX_TCPSERVERSTREAMS
                Exit Sub
            End If

            If SocketCount >= MAX_SOCKETS Then
                MaxPermitted = MAX_SOCKETS
                Call RejectConnection(IPPort, MaxPermitted, 5)
                Exit Sub
            End If
        End If
    End With
            
'The first request ID is the handle of the Listening Winsock
'The Second request ID is the handle of the Winsock which we returned
'as the Accept the first time.
'Hence if the requested ID matches that of a Winsock we created first
'time it will connect

 'Call DisplayWinsock
'Stop so that it stream not try to reconnect while socket is connecting
'This will cause a second connection (with a closed first one)
        frmRouter.ReconnectTimer.Enabled = False
        
        IdxStream = CreateStreamSocket(IdxListen)
'Debug.print & Sockets(IdxListen).Winsock.PermittedStreams & " permitted streams"
        If IdxStream < 1 Then
'No free sockets or Streams
            frmRouter.ReconnectTimer.Enabled = True
'            WriteLog "Incoming TCP Connection Request to " & Sockets(IdxListen).DevName & " rejected - all " & Sockets(IdxListen).Winsock.PermittedStreams & " streams in use"
            Exit Sub
        End If

'Open the Stream Handler (must be in connection request
'otherwise a second handler is opened)
        Call OpenHandler(IdxStream)
' Call DisplayWinsock
'Then create the link from the Listener to the Stream
'because the Route to the stream will be via the Listener's Routes
        Sockets(IdxListen).Winsock.Streams(Sockets(IdxStream).Winsock.Sidx) _
        = IdxStream
'IdxStream = 5
        HidxStream = Sockets(IdxStream).Hidx
        
'Accept in incoming connection to a new winsock socket
'        If Winsock(HidxStream).SocketHandle = requestID Then
'Stop
'        Else
        Winsock(HidxStream).Accept requestID
'        End If


'DoEvents
'State should be connected (7)
        Sockets(IdxStream).State = Winsock(HidxStream).State
        If Sockets(IdxStream).State = sckConnected Then
'        Call CreateStreamForwards(IdxStream)
'try and use one routine
            Call Routecfg.CreateSocketForwards(IdxStream)
        End If
        Call UpdateMshRows
        frmRouter.ReconnectTimer.Enabled = True
Exit Sub

ConnectionRequest_error:
    On Error GoTo 0
    MsgBox "TCP ConnectionRequest Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "On " & Sockets(Winsocks(HidxListen).sIndex).DevName, , "ConnectionRequest"
End Sub

Private Sub RejectConnection(IPPort As String, MaxPermitted As Long, Reason As Long)
Dim Client As clsClient

On Error GoTo Key_NewClient
Set Client = Clients(IPPort)  'see if this ship is in collection
On Error GoTo 0             'if not create new ship in ships collection

Client_Update:
If IPPort <> "" Then Client.IPPort = IPPort
On Error Resume Next    'potential oveflow
With Client
    Select Case Reason
    Case Is = 1
        .IPRetryCount = .IPRetryCount + 1
        .MaxIPs = MaxPermitted
    Case Is = 2
        .PortRetryCount = .PortRetryCount + 1
        .MaxPorts = MaxPermitted
    Case Is = 3
        .StreamRetryCount = .StreamRetryCount + 1
        .MaxStreams = MaxPermitted
    Case Is = 4
        .WinsockRetryCount = .WinsockRetryCount + 1
        .MaxWinsocks = MaxPermitted
    Case Is = 5
        .SocketRetryCount = .SocketRetryCount + 1
        .MaxSockets = MaxPermitted
    End Select
End With
On Error GoTo 0
Set Client = Nothing
Exit Sub

Key_NewClient:                    'create new ship in ships collection
    On Error GoTo 0
'only set up vessel if we have a name to prevent large nos of mmsi's
'without names
    If IPPort = "" Then
        Set Client = Nothing
        Exit Sub
    End If
    Set Client = New clsClient
    Client.IPPort = IPPort
    Clients.Add Client, IPPort
'testing only Call SaveVessels
    Resume Client_Update
End Sub

Public Sub OutputRejectionLog()
Dim kb As String
Dim IPPort As String
Dim Client As clsClient
Dim Plural As String
    For Each Client In Clients
        With Client
            If .IPRetryCount > 0 Then
                If .IPRetryCount > 1 Then Plural = "s" Else Plural = ""
                kb = .IPRetryCount & " Client IP" & Plural & " rejected" _
                & " from " & Replace(.IPPort, ":", " to ") _
                & ", limit is " & .MaxIPs & " from same IP"
                WriteLog kb
            End If
            If .PortRetryCount > 0 Then
                If .PortRetryCount > 1 Then Plural = "s" Else Plural = ""
                kb = .PortRetryCount & " Client Port" & Plural & " rejected" _
                & " from " & Replace(.IPPort, ":", " to ") _
                & ", limit is " & .MaxPorts & " to same port"
                WriteLog kb
            End If
            If .StreamRetryCount > 0 Then
                If .StreamRetryCount > 1 Then Plural = "s" Else Plural = ""
                kb = .StreamRetryCount & " Client Stream" & Plural & " rejected" _
                & " from " & Replace(.IPPort, ":", " to ") _
                & ", limit is " & .MaxStreams & " Client streams"
                WriteLog kb
            End If
            If .WinsockRetryCount > 0 Then
                If .WinsockRetryCount > 1 Then Plural = "s" Else Plural = ""
                kb = .WinsockRetryCount & " TCP/IP Socket" & Plural & " rejected" _
                & " from " & Replace(.IPPort, ":", " to ") _
                & ", limit is " & .MaxWinsocks & " TCP/IP Sockets"
                WriteLog kb
            End If
            If .SocketRetryCount > 0 Then
                If .SocketRetryCount > 1 Then Plural = "s" Else Plural = ""
                kb = .SocketRetryCount & " Connection" & Plural & " rejected" _
                & " from " & Replace(.IPPort, ":", " to ") _
                & ", limit is " & .MaxSockets & " Connections"
                WriteLog kb
            End If
            Clients.Remove .IPPort
        End With
    Next Client
End Sub
'Returns Stream Socket or -1 if none available
Private Function CreateStreamSocket(IdxListen As Long) As Long
Dim IdxStream As Long
Dim Fidx As Long
Dim Sidx As Long 'Array Index of the Stream (in the Listening Socket)
Dim Idx As Long

'    For Idx = 1 To UBound(Sockets)
'        If IsTcpStream(Idx) Then
'           If Sockets(Idx).State <> sckConnected Then
'IdxStream = Idx
'Stop
'            End If
'        End If
'    Next Idx
'Get Socket to allocate to this connection request
    IdxStream = FreeSocket
    If IdxStream < 1 Then
        CreateStreamSocket = -1
        Exit Function
    End If
    
'Ensure we have not exceeded the maximum allowed streams
    Sidx = FreeStream(IdxListen)
    If Sidx < 1 Then
        CreateStreamSocket = -1
        Exit Function
    End If
'Copy the template socket - the "listening" port
    Sockets(IdxStream).DevName = Sockets(IdxListen).DevName & "_" & Sidx
    Sockets(IdxStream).Handler = Sockets(IdxListen).Handler
    Sockets(IdxStream).Direction = Sockets(IdxListen).Direction
    Sockets(IdxStream).Enabled = Sockets(IdxListen).Enabled
    Sockets(IdxStream).Graph = Sockets(IdxListen).Graph
    Sockets(IdxStream).Winsock.Protocol = Sockets(IdxListen).Winsock.Protocol
    Sockets(IdxStream).Winsock.Server = Sockets(IdxListen).Winsock.Server
'Use the same socket
    Sockets(IdxStream).Winsock.LocalPort = Sockets(IdxListen).Winsock.LocalPort  'Allocate new socket
'    Sockets(IdxStream).Winsock.LocalPort = 0  'Allocate new socket
'    Sockets(IdxStream).Winsock.LocalPort = Winsock(Sockets(IdxListen).Hidx).LocalPort  'Allocate new socket
    Sockets(IdxStream).Winsock.Oidx = IdxListen
'    Sockets(IdxListen).Winsock.Streams(Sidx) = IdxStream
'Not sure why I should need this in the Stream socket
'I set up the stream index
    Sockets(IdxStream).Winsock.Sidx = Sidx
'copy Output Format
    Sockets(IdxStream).OutputFormat = Sockets(IdxListen).OutputFormat
'copy IEC
    Sockets(IdxStream).IEC = Sockets(IdxListen).IEC
    
'    Call OpenHandler(IdxStream)
'OpenHandler moved (must be in connection request
'otherwise a second handler is opened)
    CreateStreamSocket = IdxStream
End Function

#If V44 Then
Public Sub CreateStreamForwards_old(IdxStream As Long)
Dim IdxListen As Long
'Dim Sidx As Long
Dim Ridx As Long

    IdxListen = Sockets(IdxStream).Winsock.Oidx
'There could be a route to
    If IdxListen > 0 Then
        For Ridx = 1 To UBound(Sockets(IdxListen).Routes)
            If Sockets(IdxListen).Routes(Ridx).Enabled = True Then
                Call Routecfg.CreateForwards(Sockets(IdxListen).Routes(Ridx).AndIdx, IdxStream)
            End If
        Next Ridx
    End If
End Sub
#End If

'Index is the Winsock Source Index (hIndex)
Private Sub Winsock_DataArrival(Hidx As Integer, ByVal bytesTotal As Long)
Dim DataRcv As String
Dim i As Long

    On Error GoTo DataArrival_err
'On Error GoTo 0  'debug
'removed v45 On Error GoTo 0     'debug
    If Winsock(Hidx).State <> sckConnected Then
'Call DisplayWinsock
    End If
    Winsock(Hidx).GetData DataRcv, vbString
'we need to keep this here as it is blank (on winsock) when the timer runs
'Sockets(Index).RemoteHostIP = Winsock(Index).RemoteHostIP
'If Index = 5 Then Stop

'Call frmDpyBox.DpyBox(DataRcv, 5, "DataArrival")
    
'test no NMEA sentence delimiter
'DataRcv = Replace(DataRcv, vbCrLf, "")
'test <LF> only
'DataRcv = Replace(DataRcv, vbCrLf, vbLf)
    
    If AcceptLForCRasNmeaTerminator = True Then
        i = InStr(1, DataRcv, vbCrLf)
        If i = 0 Then 'no CRLF
            i = InStr(1, DataRcv, vbLf)
            If i >= 1 Then
                DataRcv = Replace(DataRcv, vbLf, vbCrLf)
            End If
            If i = 0 Then 'no single LF
               i = InStr(1, DataRcv, vbCr)
                If i >= 1 Then
                    DataRcv = Replace(DataRcv, vbCr, vbCrLf)
                End If
            End If
        End If
'if neither add CRLF so user see what is happening (Nick McEvoy)
        If i = 0 Then   'not CRLF or LF or CR
            DataRcv = DataRcv & vbCrLf
        End If
    End If
    
    If bytesTotal > 0 Then
        Call ForwardData(DataRcv, Winsocks(Hidx).sIndex)
    End If
Exit Sub

DataArrival_err:
'    On Error GoTo 0 'debug
    Select Case err.Number
    Case Is = sckBadState
        StatusBar.Panels(1).Text = err.Description & " on " & Sockets(Winsocks(Hidx).sIndex).DevName
        ClearStatusBarTimer.Enabled = True
        'Believe this occurs when the buffer is full
        'Wrong protocol or connection state for the requeste
        'transaction or request
'Stop
    Case Is = sckMsgTooBig
        StatusBar.Panels(1).Text = err.Description & " on " & Sockets(Winsocks(Hidx).sIndex).DevName
        WriteLog Sockets(Winsocks(Hidx).sIndex).DevName & " UDP/TCP DataArrival Error " & Str(err.Number)
        WriteLog err.Description
        WriteLog "Received from " & Winsock(Hidx).RemoteHostIP
        ClearStatusBarTimer.Enabled = True
    Case Is = sckConnectionReset
        StatusBar.Panels(1).Text = err.Description & " on " & Sockets(Winsocks(Hidx).sIndex).DevName
        ClearStatusBarTimer.Enabled = True
        'Reset by remote side is OK
    Case Else
        MsgBox "UDP/TCP DataArrival Error " & Str(err.Number) & " " & err.Description & vbCrLf _
        & "on " & Sockets(Winsocks(Hidx).sIndex).DevName, , "DataArrival"
    End Select
End Sub

Public Function TermOutput(Data As String, Optional Source As Long) As Long
'Display Received Sentence
    If txtTerm.Enabled = True Then
        txtTerm.SelStart = Len(txtTerm.Text)
        If Source <> 0 Then
            txtTerm.SelText = CStr(Source) & "<" & Data
        Else
            txtTerm.SelText = Data
        End If
         If Len(txtTerm.Text) > 4096 Then
            txtTerm.Text = Right$(txtTerm.Text, 2048)
        End If
    End If
    TermOutput = SendMessageAsLong(txtTerm.hwnd, EM_GETLINECOUNT, 0, 0)
    End Function


#If V44 Then
Public Sub TermOutput_old(Data As String, Optional Source As Long)
'Display Received Sentence
    If txtTerm.Enabled = True Then
        txtTerm.SelStart = Len(txtTerm.Text)
        If Source <> 0 Then
            txtTerm.SelText = CStr(Source) & "<" & Data
        Else
            txtTerm.SelText = Data
        End If
         If Len(txtTerm.Text) > 4096 Then
            txtTerm.Text = Right$(txtTerm.Text, 2048)
        End If
    End If
End Sub
#End If

Private Sub CheckQueues()
Dim Index As Long
Dim Qleft As Long
    
    For Index = 1 To UBound(Sockets)
        With Sockets(Index)
            Qleft = .Qrear - .Qfront
            If Qleft < 0 Then Qleft = UBound(.Buffer) + Qleft
            If Qleft > 0 Then
'Debug.Print "CheckQueues " & Index & ":" & Qleft
                Call Deque(Index, 150)
            End If
        End With
    Next Index
End Sub

Private Sub DequeTimeout_Timer()
DequeTimeout.Enabled = False
End Sub

Public Sub DisplayForwarding(Optional Detail As Boolean)
Dim Idx As Long
Dim Fidx As Long
Dim kb As String
Dim Count As Long
Dim TotCount As Long
Dim Qleft As Long

    
    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
        For Fidx = 1 To UBound(.Forwards)
            If .Forwards(Fidx) > 0 Then
                kb = kb & Sockets(Idx).DevName
                If Detail = True Then
                    kb = kb & ", Idx(" & Idx & ")"
                End If
                kb = kb & " to " & Sockets(.Forwards(Fidx)).DevName
                If Detail = True Then
                    kb = kb & ", Fidx(" & Fidx & ")"
                End If
                kb = kb & vbCrLf
                With Sockets(.Forwards(Fidx))
                    Qleft = .Qrear - .Qfront
                    If Qleft < 0 Then Qleft = UBound(.Buffer) + Qleft
                    If Qleft > 0 Then
                        kb = kb & vbTab & Qleft & " Sentences in Output Queue" & vbCrLf
                    End If
                End With
                Count = Count + 1
            End If
        Next Fidx
        End With
        If Count = 0 Then
            If Detail = True Then
                kb = kb & Sockets(Idx).DevName
                kb = kb & ", Idx(" & Idx & ")"
                kb = kb & " to Nothing" & vbCrLf
                End If
        End If
    TotCount = TotCount + Count
    Count = 0
    Next Idx

    If TotCount = 0 Then
        kb = kb & "There is no Forwarding in use"
    End If
    MsgBox kb, , "Forwarding (" & TotCount & ")"
End Sub

Public Sub DisplayFiles()
Dim result As Boolean
Dim Idx As Long
Dim Hidx As Variant
Dim kb As String
Dim Count As Long
Dim i As Integer
Dim Fidx As Integer
Dim Size As Long

    For Hidx = 1 To UBound(Files)
        If Not Files(Hidx) Is Nothing Then
            Idx = Files(Hidx).sIndex
            If Sockets(Idx).State <> -1 Then
                kb = kb & Hidx & " = " & Files(Hidx).Name _
                & " [" & aEnabled(Sockets(Idx).Enabled) & "]" & vbCrLf
                kb = kb & vbTab & Files(Hidx).FileName & vbCrLf
                Select Case Sockets(Idx).Direction
                Case Is = 1     'Input
                    kb = kb & vbTab & "Open for " & aDirection(Sockets(Idx).Direction) & vbCrLf
                    kb = kb & vbTab & "Maximum Read Rate is " & aReadRate(Sockets(Idx).File.ReadRate) & vbCrLf
                Case Is = 2     'Output
                    If Sockets(Idx).File.RollOver = True Then
                        kb = kb & vbTab & "Open for " & aDirection(Sockets(Idx).Direction) & vbCrLf
                        kb = kb & vbTab & "File is being Rolled Over" & vbCrLf
                    Else
                        kb = kb & vbTab & "Open for " & aDirection(Sockets(Idx).Direction) & vbCrLf
                        kb = kb & vbTab & "File is not being Rolled Over" & vbCrLf
                    End If
                End Select
                If FileInputTimerExists(CInt(Hidx)) Then
                    kb = kb & vbTab & "Input Timer Enabled = " & FileInputTimer(Hidx).Enabled _
                    & " (" & FileInputTimer(Hidx).Interval & ")" & vbCrLf
                End If
                On Error Resume Next    'may not exist
                Size = FileLen(Files(Hidx).FileName)
                Select Case Size
                Case Is < 1000
                    kb = kb & vbTab & "Size is " & Size & " bytes" & vbCrLf
                Case Is < 1000000
                    kb = kb & vbTab & "Size is " & Int(Size / 1000) & " KB" & vbCrLf
                Case Else
                    kb = kb & vbTab & "Size is " & Int(Size / 1000000) & " MB" & vbCrLf
                End Select
                Count = Count + 1
            End If
        kb = kb & vbCrLf
        End If
    Next Hidx
    If Count = 0 Then
        kb = kb & "There are no File handlers allocated" & vbCrLf
    End If
    
'There should not be a handler if the socket is not open
    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).Handler = 2 & Sockets(Idx).State = 1 Then
            If Sockets(Idx).State = 0 And Sockets(Idx).Hidx = 0 Then
                kb = kb & "No " & aHandler(Sockets(Idx).Handler) & " Handler for " & Sockets(Idx).DevName & " [" & Sockets(Idx).File.SocketFileName & "]" & vbCrLf
            End If
        End If
    Next Idx
    MsgBox kb, , "Files"

End Sub

Public Sub DisplaySockets(Optional Detail As Boolean)
Dim Idx As Long
Dim Fidx As Long
Dim kb As String
Dim Count As Long
Dim i As Long
Dim iCount As Long
Dim myLayout As clsLayout

    Detail = False  'True TTYs sockets not in use
    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
            kb = kb & Idx & vbTab
            If .State = -1 And Detail = False Then
                kb = kb & .DevName & " (" & Idx & ") not in use" & ","
            Else
                Count = Count + 1
                kb = kb & .DevName
                kb = kb & "," & aState(.State)
                kb = kb & "," & aEnabled(.Enabled)
                kb = kb & "," & aDirection(.Direction)
                kb = kb & "," & aHandler(.Handler) & "("
                If .Hidx < 1 Then
                    kb = kb & "de-allocated"
                Else
                    kb = kb & "allocated"
                End If
                kb = kb & ")"
                iCount = 0
                For i = 1 To UBound(.Routes)
                    If .Routes(i).AndIdx > 0 Then
                        iCount = iCount + 1
                    End If
                Next i
                If iCount > 0 Then
                    If iCount = 1 Then
                        kb = kb & "," & iCount & " Route"
                    Else
                        kb = kb & "," & iCount & " Routes"
                    End If
                End If
                iCount = 0
                For i = 1 To UBound(.Forwards)
                    If .Forwards(i) <> 0 Then
                        iCount = iCount + 1
                    End If
                Next i
                If iCount > 0 Then
                    If iCount = 1 Then
                        kb = kb & "," & iCount & " Forward"
                    Else
                        kb = kb & "," & iCount & " Forwards"
                    End If
                End If
                If Not .Recorder.Output Is Nothing Then
                    kb = kb & ", Recorder"
                End If
                If SourceDuplicateFilter.DmzIdx = Idx Then
                    kb = kb & ", DMZ"
                End If
                If .QLost Then
                    kb = kb & "," & .QLost & " Lost"
                End If
                If .TryCount > 0 Then
                    If .TryCount = 1 Then
                        kb = kb & "," & .TryCount & " Try"
                    Else
                        kb = kb & "," & .TryCount & " Tries"
                    End If
                End If
            End If
            kb = kb & vbCrLf
        End With
    Next Idx
    If Count = 0 Then
        kb = "No Connections are in use"
    End If
    MsgBox kb, , "Connections"
End Sub

'All Routes are displayed, unless ReqIdx is given
'Detail TTYs the Indexes as well
Public Sub DisplayRoutes(Optional Detail As Boolean)
Dim Idx As Long
Dim Ridx As Long
Dim kb As String
Dim Count As Long
Dim SocketCount As Long
Dim RidxCount As Long

If jnasetup Then
    Detail = True
End If
    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
        For Ridx = 1 To UBound(.Routes)
            If .Routes(Ridx).AndIdx > 0 Then
                kb = kb & Sockets(Idx).DevName & " [" _
                & aDirection(Sockets(Idx).Direction) & "]"
                If Detail = True Then
                    kb = kb & ", Idx(" & Idx & ")"
                End If
                kb = kb & " and " & Sockets(.Routes(Ridx).AndIdx).DevName _
                & " [" & aDirection(Sockets(.Routes(Ridx).AndIdx).Direction) & "]"
                If Detail = True Then
                    kb = kb & ", Ridx(" & Ridx & ")"
                End If
                If .Routes(Ridx).Enabled = True Then
                    kb = kb & " [Enabled"
                Else
                    kb = kb & " [Disabled"
                End If
                If .Direction = Sockets(.Routes(Ridx).AndIdx).Direction _
                And .Direction <> 0 Then
                    kb = kb & "-Inactive]"
                Else
                    kb = kb & "]"
                End If
'                If Detail = True Then
                    kb = kb & ", For=" & .Routes(Ridx).ForwardCount & ", Rev=" & .Routes(Ridx).ReverseCount
'                End If
                kb = kb & vbCrLf
                RidxCount = RidxCount + 1
            End If
        Next Ridx
        End With
        If Detail = True Then
            If RidxCount = 0 Then
                kb = kb & Sockets(Idx).DevName & " [" _
                & aDirection(Sockets(Idx).Direction) & "]"
                If Detail = True Then
                    kb = kb & ", Idx(" & Idx & ")"
                End If
                kb = kb & " and Nothing" & vbCrLf
            End If
        End If
    Count = Count + RidxCount
    RidxCount = 0
    Next Idx
    If Count = 0 Then
        kb = kb & "There are no Routes in use"
    End If
'Route count is a check the no of routes are correct
    MsgBox kb, , "Routes (" & RouteCount & ")"
End Sub

Public Sub SetCaption()
    Caption = "NmeaRouter " & App.Major & "." & App.Minor & "." & App.Revision & " "
    If CurrentProfile <> "" Then
        Caption = Caption & " [" & CurrentProfile & "]"
#If test = True Then
    Caption = Caption & " Test mode"
#End If
    End If

End Sub


Public Sub FileInputTimer_Timer(Hidx As Integer)
    
'Each time it fires it calls ContinueFileInput which
'will read from 1 to 50 records. This will continue until
'the FileInputTimer is disabled
    
    Call Files(Hidx).ContinueFileInput
'If is disabled by pause and enabled by continue
'If stop is pressed the Handler is closed, which removes the timer
End Sub

'Contructs the Port List from the Ports Collection
'Done whenever connections or routes change - updates frmRouter display
'Or Graph View changed
Public Sub UpdateMshRows()
Const RowPadding = 15   'dont know why but it is
Dim Idx As Long
Dim Ridx As Long
Dim Row As Long
Dim Col As Long
Dim kb As String
Dim SaveStatsTimer As Boolean

'Call DisplayWindow("UpdateMshRows - Entry")
'Restore state of stats timer on exit
'Call ResizeCaption
    SaveStatsTimer = StatsTimer.Enabled
    StatsTimer.Enabled = False
    With mshSockets
.Redraw = False
'here to turn off red if a row has been deleted, because
'we are going to reconstruct the table
        .Rows = 2
'Clear the first row (it will be blank if no conns)
        For Col = 0 To .Cols - 1
            .TextMatrix(1, Col) = ""
            .Col = Col
            .CellBackColor = vbWhite    'Blank first row
        Next Col
        .Col = 0
        .ColSel = 0
kb = .Height
        For Idx = 1 To UBound(Sockets)
            If Sockets(Idx).State <> -1 Then
                Row = Row + 1
                If Row = .Rows Then .Rows = .Rows + 1
                    .TextMatrix(Row, 0) = Idx   'Port
                    .TextMatrix(Row, 1) = Sockets(Idx).DevName   'Comments
                    .TextMatrix(Row, 2) = aDirection(Sockets(Idx).Direction)   'Direction
                    '.TextMatrix(row, 3) = Sockets(Idx)   'Address
                    '.TextMatrix(row, 4) =   'Speed
                    .TextMatrix(Row, 5) = aHandler(Sockets(Idx).Handler)   'Protocol
                    If Sockets(Idx).Handler = 0 And Sockets(Idx).Hidx > 0 Then
                        .TextMatrix(Row, 5) = aProtocol(CInt(Sockets(Idx).Winsock.Protocol))
                    End If
                    .TextMatrix(Row, 6) = aEnabled(Sockets(Idx).Enabled)   'Enabled
                    .TextMatrix(Row, 7) = Sockets(Idx).MsgCount   'Count
                    .TextMatrix(Row, 8) = aEnabled(Sockets(Idx).Graph)
                    .TextMatrix(Row, 9) = Idx   'Socket
                    .TextMatrix(Row, 10) = Sockets(Idx).Hidx   'Socket
                    .TextMatrix(Row, 11) = aState(Sockets(Idx).State)   'State
                    .TextMatrix(Row, 12) = Sockets(Idx).errmsg   'Status

'Grey Baud if UDP/TCP
'When Col is changed ColSel is set to the same (so only 1 cell is selected)
'To select more than 1 col then sel colsel to left or right of col
                .Row = Row
            End If
        
        Next Idx
'If you set the height here the scoll bars disappear
'21/9/13        .Height = 4 * RowPadding + (.CellHeight + RowPadding) * (.Rows)
'set sort columns for sort key (all cols are actually sorted)
        .ColSel = 0 'from col
        .Col = 0    'to col
        .RowSel = 1 'first nonfixed row
        .Row = 1    'if same as above all rows are sorted
        .Sort = flexSortGenericAscending
'Do the form resizing here
'    Call frmPorts.Size
    Call SpeedTimer_Timer   'Uses matrix as the index
.Redraw = True
    End With

    Row = 0
    Col = 0
    With mshRoutes
.Redraw = False
'here to turn off red if a row has been deleted, because
'we are going to reconstruct the table
        .Rows = 2
'Clear the first row (it will be blank if no conns)
        For Col = 0 To .Cols - 1
            .TextMatrix(1, Col) = ""
            .Col = Col
            .CellBackColor = vbWhite    'Blank first row
        Next Col
        .Col = 3
        .ColSel = 0
kb = .Height
        For Idx = 1 To UBound(Sockets)
            For Ridx = 1 To UBound(Sockets(Idx).Routes)
                If Sockets(Idx).Routes(Ridx).AndIdx > 0 Then
                    Row = Row + 1
                    If Row = .Rows Then .Rows = .Rows + 1
                    .TextMatrix(Row, 0) = Sockets(Idx).DevName
'Arrow
                    .TextMatrix(Row, 2) = Sockets(Sockets(Idx).Routes(Ridx).AndIdx).DevName
                    .TextMatrix(Row, 3) = aEnabled(Sockets(Idx).Routes(Ridx).Enabled)
                    .TextMatrix(Row, 4) = Idx
                    .TextMatrix(Row, 5) = Sockets(Idx).Routes(Ridx).AndIdx
                    .TextMatrix(Row, 6) = Ridx
                    .Row = Row
                    If Sockets(Idx).Routes(Ridx).Enabled = True Then
                        .CellBackColor = vbGreen
                    End If
                End If
            Next Ridx
        Next Idx
    
'If you set the height here the scoll bars disappear
'    .Height = 4 * RowPadding + (.CellHeight + RowPadding) * (.Rows)
'set sort columns for sort key (all cols are actually sorted)
        .ColSel = 0 'from col
        .Col = 0    'to col
        .RowSel = 1 'first nonfixed row
        .Row = 1    'if same as above all rows are sorted
        .Sort = flexSortGenericAscending
'Do the form resizing here
'    Call frmPorts.Size
.Redraw = True
    End With

'Call ResizeCaption
'    Call ResizeControls
    Call ArrangeControls
'Call ResizeCaption
    StatsTimer.Enabled = SaveStatsTimer
'Call DisplayWindow("UpdateMshRows - Exit")
'.Print "UpdateMshRows"
End Sub

Public Function AllStreamCount() As Long
Dim Idx As Long
Dim i As Long
Dim Count As Long
    
    With mshSockets
        If .TextMatrix(1, 0) = "" Then
            Exit Function
        End If
'Must not be run when Flexgrids are being changed
        For i = 1 To .Rows - 1
'check Idx is numeric
            Idx = .TextMatrix(i, 9)
            If Idx > UBound(Sockets) Then   'Should not happen as MSH should have been cleared
                Exit Function
            End If
            Count = Count + StreamCount(Idx)
        Next i
    End With
    AllStreamCount = Count
End Function
'this fires every 1/2 second
Private Sub StatsTimer_Timer()
Dim i As Integer
Dim Idx As Long
Dim Hidx As Long
Dim Idx1 As Long
Dim Idx2 As Long
Dim Odx1 As Long
Dim Odx2 As Long
Dim Ridx As Long
Dim Fidx As Long
Dim IdxListen As Long
Dim Alternate As Boolean    'If Server toggle between local and remote
Dim strTemp As String
Dim kb As String
Dim k As Long
Dim SortRequired As Boolean

'Check weve at least one port in the flexgrid
    With mshSockets
        If .TextMatrix(1, 0) = "" Then
            Exit Sub
        End If
    
'Must not be run when Flexgrids are being changed

'Speed time appears to be disabled when profile is changed
'        If SpeedTimer.Enabled = False Then
'            SpeedTimer.Enabled = True
'        End If
        
        For i = 1 To .Rows - 1
'check Idx is numeric
            Idx = .TextMatrix(i, 9)
            If Idx > UBound(Sockets) Then   'Should not happen as MSH should have been cleared
                Exit Sub
            End If
            
'        .Refresh
'        .Redraw = False
            If Sockets(Idx).TryCount = 0 And Sockets(Idx).errmsg <> "" Then
                Sockets(Idx).errmsg = ""
            End If
            
            .Row = i
            .Col = 6
            If Sockets(Idx).Enabled = False Then
                .CellBackColor = vbWhite
                .TextMatrix(i, 6) = "Disabled"
            Else
                .CellBackColor = vbYellow
                .TextMatrix(i, 6) = "Enabled"
'Set up Socket(idx).state from Winsock
                Select Case Sockets(Idx).Handler
                Case Is = 0          'Winsock
                    Hidx = Sockets(Idx).Hidx
                    If Hidx > 0 Then
                        Sockets(Idx).State = Winsock(CInt(Hidx)).State
                            Sockets(Idx).Winsock.RemoteHostIP = Winsock(CInt(Hidx)).RemoteHostIP
                            Sockets(Idx).Winsock.RemotePort = Winsock(CInt(Hidx)).RemotePort
            Sockets(Idx).Winsock.LocalPort = Winsock(CInt(Hidx)).LocalPort
            Sockets(Idx).Winsock.LocalIP = Winsock(CInt(Hidx)).LocalIP

                    Else    'Must be closed
                        Sockets(Idx).State = 0
                    End If
                    Select Case Sockets(Idx).State
                    Case Is = sckListening
                        .CellBackColor = &H19D2FF   'Orange
                    Case Is = sckConnecting
                        .CellBackColor = &H19D2FF   'Orange
                    Case Is = sckClosing
                        .CellBackColor = vbRed
                    Case Else
                        .CellBackColor = vbYellow   'Connection is made
                    End Select
                 Case Else        'Non Winsock handler
                    Select Case Sockets(Idx).State
                    Case Is = sckClosed
                        .CellBackColor = vbYellow   'Orange
                    Case Is = sckOpen
                        .CellBackColor = vbYellow   'Non Winsock open socket
                    Case Else
                        .CellBackColor = vbRed
                    End Select
                End Select
            
'Display Green if any messages received on a not Closed port
                If CLng(.TextMatrix(i, 7)) <> Sockets(Idx).MsgCount Then
                    .CellBackColor = vbGreen   'Connection is made
                End If
'                .TextMatrix(i, 7) = CStr(Sockets(Idx).MsgCount)
            End If  'Enabled ports


'Enabled or Disabled from here

            Select Case Sockets(Idx).Handler
#If False Then
            Case Is = 0
'When acting as a server display the Local Address when listening
'and the remote address when sending
                If Sockets(Idx).State = sckListening Then
                    .TextMatrix(i, 3) = Sockets(Idx).Winsock.LocalIP _
                    & ":" & Sockets(Idx).Winsock.LocalPort
                Else    'Not listening
'Or display where we are sending it to ?
                    If Sockets(Idx).Direction = 1 Then  'input
                        If Sockets(Idx).Winsock.RemoteHostIP <> "" Then
'When it has been opened we know the IP
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHostIP
                        Else
'Before it has been opened
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHost
                        End If

                        .TextMatrix(i, 3) = .TextMatrix(i, 3) & ":" & Sockets(Idx).Winsock.RemotePort
                    Else                'Output
                        If Sockets(Idx).Winsock.RemoteHost <> "" Then
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHost & ":" & Sockets(Idx).Winsock.RemotePort
                        Else
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHostIP & ":" & Sockets(Idx).Winsock.RemotePort
                        End If
                    End If  'input or Output
                    If Sockets(Idx).Winsock.Server = 1 And Sockets(Idx).State = 0 Then
                        .TextMatrix(i, 3) = Sockets(Idx).Winsock.LocalIP & ":" & Sockets(Idx).Winsock.LocalPort
                    End If
                End If
                .TextMatrix(i, 5) = aProtocol(CInt(Sockets(Idx).Winsock.Protocol))
#End If
            Case Is = 0
'When acting as a server display the Local Address when listening
'and the remote address when sending
                If Sockets(Idx).State = sckListening Then
                    .TextMatrix(i, 3) = Sockets(Idx).Winsock.LocalIP
                Else    'Not listening
'Or display where we are sending it to ?
                    If Sockets(Idx).Direction = 1 Then  'input
                        If Sockets(Idx).Winsock.RemoteHostIP <> "" Then
'When it has been opened we know the IP
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHostIP
                        Else
'Before it has been opened
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHost
                        End If
                    Else                'Output
                        If Sockets(Idx).Winsock.RemoteHost <> "" Then
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHost
                        Else
                            .TextMatrix(i, 3) = Sockets(Idx).Winsock.RemoteHostIP
                        End If
                    End If  'input or Output
                    If Sockets(Idx).Winsock.Server = 1 And Sockets(Idx).State = 0 Then
                        .TextMatrix(i, 3) = Sockets(Idx).Winsock.LocalIP
                    End If
                End If
'                If Sockets(Idx).Winsock.Protocol = sckTCPProtocol Then
                    If Sockets(Idx).Winsock.Server = 1 Then
                        .TextMatrix(i, 3) = .TextMatrix(i, 3) & ":" & Sockets(Idx).Winsock.LocalPort
                    Else
                        .TextMatrix(i, 3) = .TextMatrix(i, 3) & ":" & Sockets(Idx).Winsock.RemotePort
                    End If
'                Else    'udp
'                    If Sockets(Idx).Winsock.Server = 1 Then
'                        .TextMatrix(i, 3) = .TextMatrix(i, 3) & ":" & Sockets(Idx).Winsock.LocalPort
'                    Else
'                        .TextMatrix(i, 3) = .TextMatrix(i, 3) & ":" & Sockets(Idx).Winsock.RemotePort
'                    End If
'                End If
                .TextMatrix(i, 5) = aProtocol(CInt(Sockets(Idx).Winsock.Protocol))
            
            Case Is = 1
                If Sockets(Idx).Comm.VCP <> "" Then
                    .TextMatrix(i, 3) = Sockets(Idx).Comm.Name & " <--> " & Sockets(Idx).Comm.VCP
                Else
                    .TextMatrix(i, 3) = Sockets(Idx).Comm.Name
                End If
            End Select  'Winsock ports

            .Col = 8
            If Sockets(Idx).Graph = True Then
                .CellBackColor = vbGreen
                .TextMatrix(i, 8) = "Enabled"
            Else
                .CellBackColor = vbYellow
                .TextMatrix(i, 8) = "Disabled"
            End If
            If Sockets(Idx).Enabled = False Then
                .CellBackColor = vbWhite
            End If

'Display status of all ports
            .TextMatrix(i, 2) = aDirection(Sockets(Idx).Direction)   'Direction
            .TextMatrix(i, 7) = CStr(Sockets(Idx).MsgCount)
            .TextMatrix(i, 9) = Idx
            .TextMatrix(i, 10) = Sockets(Idx).Hidx
            .TextMatrix(i, 11) = aState(Sockets(Idx).State)
            .TextMatrix(i, 12) = aState(Sockets(Idx).State) & " "
            If Sockets(Idx).errmsg <> "" Then
                .TextMatrix(i, 12) = Sockets(Idx).errmsg
'When a socket is closed the errmsg is cleared
'Not good enough as enough
'                Sockets(Idx).errmsg = ""
            End If
                
        Next i
'Set the selected range to 0 (otherwise its left Blue)
'by setting it to the current Row & COl
    .Row = 0
    .Col = 0
    .RowSel = .Row
    .ColSel = .Col
'    .Redraw = True
'    .Refresh
    End With

    With mshRoutes
'Check weve at least one route in the flexgrid
        If .TextMatrix(1, 0) = "" Then
            Exit Sub
        End If
                
'        .Refresh
'        .Redraw = False
        For i = 1 To .Rows - 1
            .Row = i

'The routes are always on the Socket that "Owns" the connection
'Normally this is the same as the connection
'but with a TCP server the Route is on the "Listening" socket
'This will should not be the case of TextMatrix because the route
'is always on the TCP Listener (which is the Owner)
'But the forward is on the stream socket, which is NOT the listener
'            Idx1 = cOidx(.TextMatrix(i, 4))
'            Idx2 = cOidx(.TextMatrix(i, 5))
'If i = 1 Then Stop
            Idx1 = .TextMatrix(i, 4)
            Idx2 = .TextMatrix(i, 5)
'MsgBox Routecfg.RouteExists(Idx1, Idx2)
            Ridx = .TextMatrix(i, 6)
            .TextMatrix(i, 3) = aEnabled(Sockets(Idx1).Routes(Ridx).Enabled)

'Set the arrows bold
            .Col = 1
            .CellFontBold = True
            
'UpdateMshRows now reverses arrows (then we can reverse arrows)
'Arrow
            If Sockets(Idx1).Routes(Ridx).ReverseCount = 0 _
            And Sockets(Idx1).Routes(Ridx).ForwardCount = 0 Then
 'No direction enabled
                .TextMatrix(i, 1) = ""
            Else
                If Sockets(Idx1).Routes(Ridx).ReverseCount > 0 _
                And Sockets(Idx1).Routes(Ridx).ForwardCount > 0 Then
'Both directions enabled
                    .TextMatrix(i, 1) = "<---->"
                Else
                    If Sockets(Idx1).Routes(Ridx).ForwardCount > 0 Then
'Forward direction enabled
                        .TextMatrix(i, 1) = "-->"
                    End If
                    If Sockets(Idx1).Routes(Ridx).ReverseCount > 0 Then
'If i = 1 Then
'kb = ""
'For k = 0 To 5
'    kb = kb & "," & k & "=" & .TextMatrix(i, k)
'Next k
'Stop
'End If
'Reverse direction enabled
                        If .TextMatrix(i, 0) = Sockets(.TextMatrix(i, 4)).DevName Then
'Name not yet reversed from --> to
                            strTemp = .TextMatrix(i, 0)
                            .TextMatrix(i, 0) = .TextMatrix(i, 2)
                            .TextMatrix(i, 2) = strTemp
                            SortRequired = True
                        End If
                        .TextMatrix(i, 1) = "-->"
                    End If
                End If
            End If
            
'            .TextMatrix(i, 1) = ""
'            If Sockets(Idx1).Routes(Ridx).ReverseCount > 0 Then
'                .TextMatrix(i, 1) = "<--"
'            End If
'            If Sockets(Idx1).Routes(Ridx).ForwardCount > 0 Then
'                .TextMatrix(i, 1) = .TextMatrix(i, 1) & "-->"
'            End If
'            If .TextMatrix(i, 1) = "<--" Then
'                strTemp = .TextMatrix(i, 0)
'                .TextMatrix(i, 0) = .TextMatrix(i, 2)
'                .TextMatrix(i, 2) = strTemp
'                .TextMatrix(i, 1) = "-->"
'Stop
'            End If

'Set the colour
            .Col = 3
            If Sockets(Idx1).Routes(Ridx).Enabled = False Then
                .CellBackColor = vbWhite
            Else
                If Sockets(Idx1).Enabled = True _
                And Sockets(Idx2).Enabled = True _
                And Sockets(Idx1).Direction = Sockets(Idx2).Direction _
                And Sockets(Idx1).Direction <> 0 Then
'Route is inactive as both directions are the same
                    .CellBackColor = vbRed
                Else
                    If .TextMatrix(i, 1) = "" Then
                        .CellBackColor = vbYellow
                    Else
                        .CellBackColor = vbGreen
                    End If
                End If
            End If
        
        Next i
    If SortRequired Then
        .ColSel = 0 'from col
        .Col = 0    'to col
        .RowSel = 1 'first nonfixed row
        .Row = 1    'if same as above all rows are sorted
        .Sort = flexSortGenericAscending
    End If
'Set the selected range to 0 (otherwise its left Blue)
'by setting it to the current Row & COl
    .Row = 0
    .Col = 0
    .RowSel = .Row
    .ColSel = .Col
'    .Redraw = True
'    .Refresh
    End With
    
'Debug only
'    If Not SourceDuplicateFilter Is Nothing Then
'        Call SourceDuplicateFilter.Status
'    End If
    
    Exit Sub
StatsTimer_err:         'Output Server
    Sockets(Idx).errmsg = err.Description
    MsgBox "StatsTimer Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "Socket " & Idx, , "Stats Timer"
End Sub

'fires ever 6 seconds
Private Sub SpeedTimer_Timer()
Dim Idx As Long
Dim i As Long
Static LastTick As Double   'long
Dim Tick As Double          'long
Dim Elapsed As Long
'Must not run while Flexgrids are being altered

'If sockets are blank must not try and update (subscript error)
    If mshSockets.TextMatrix(1, 0) = "" Then
        Exit Sub
    End If
'4,294,967,296 max
'63369609
'Debug.Print 4294967296# - 63457986
'Long increments to 2147483647 and then wraps to -2147483648.
'Mar17 https://msdn.microsoft.com/en-us/library/windows/desktop/ms724408%28v=vs.85%29.aspx
'says GetTickCount wraps to 0
'highest GetTickCount is 2147483647
'Lowest GetTickCount is 0 when first run
'Tick = LongToUnsigned(-2147483648#)    'Lowest long
'If Tick < LastTick Then LastTick = LastTick - 2147483648#
'Debug.Print (2 ^ 32 - 1)
    
    On Error GoTo Overflow  'trap overflow error v66
    Tick = LongToUnsigned(GetTickCount) '2147483647 to 0
'If LastTick=0 assume this is the first time it is called & elapsed time is 0
    If LastTick = 0 Then LastTick = Tick
'If Tick has rolled over since LastTick
    If Tick < LastTick Then LastTick = LastTick - 2147483648#   '(#=Double)

'LastTick = 4294967295#
    Elapsed = UnsignedToLong(Tick - LastTick)
    If Elapsed > 0 Then 'not first invalid first or possibly wraps after 49 days
        With mshSockets
            For i = 1 To .Rows - 1
                Idx = CInt(.TextMatrix(i, 9))
                If Idx > UBound(Sockets) Then   'Should not happen as MSH should have been cleared
                    Exit Sub
                End If
                If Sockets(Idx).Enabled = True And Sockets(Idx).Chrs <> 0 Then
'                    .TextMatrix(i, 4) = Format$(Sockets(Idx).Chrs * 8 / Elapsed, "##0.0")
                    .TextMatrix(i, 4) = Format$((Sockets(Idx).MsgCount - Sockets(Idx).LastMsgCount) * 60000 / Elapsed, "####0")
                    Sockets(Idx).LastMsgCount = Sockets(Idx).MsgCount
                    Sockets(Idx).Chrs = 0
                Else
                    .TextMatrix(i, 4) = ""
                End If
            Next i
        End With
    End If
    LastTick = Tick
'Debug.Print "Speed"
    Exit Sub

Overflow:
    LastTick = 0
    Resume Next
End Sub

Public Sub ReconnectTimer_Timer()
Dim Idx As Long
Dim Hidx As Long
Dim IdxListen As Long
Dim HidxListen As Long
Dim HandlerState As Long

'debug classes destroyed
'Call DisplayDebugSerial
'Dim Sidx As Long

'If were not running do not reconnect
'it will be restarted by CmdStart
    If cmdStart.Enabled = True Then
        ReconnectTimer.Enabled = False
        Exit Sub
    End If

'Removes any stream sockets which have been closed
'or if the Owner has been closed
    Call RemoveStreamSockets

'If the socket has been opened or handler closed
    For Idx = 1 To UBound(Sockets)
        
'With Winsock the handler status may not be open (it could be
'connecting) at the end of the OpenHandler routine.
'In this case the ReconnectTimer will spot that the Handler'
'is now Open although the Socket status will be Connecting.
'If this occurs, the Reconnect Timer must change the socket
'status and decr the TryCount.

        If Sockets(Idx).Enabled = True Then
'Clear error message from status display before re-trying
'Dont this upsets the retry counter
'Sockets(Idx).errmsg = ""
            Hidx = Sockets(Idx).Hidx
            Select Case Sockets(Idx).Handler
            Case Is = 0     'Winsock
                Select Case WinsockState(Idx)
                Case Is = sckClosed, sckError, sckClosing
                    Call OpenHandler(Idx)
'Stop
                Case Is = sckConnecting
'With TCP connection to AISHub, if the program stalls (eg with a break)afer
'the Open, and data is received, an error 40006 is generated
'The sockets is left in the connecting state, but will only start
'receiving input if closed and re-opened
                    If Sockets(Idx).Winsock.Server = 0 Then 'Client
'If Winsock(Sockets(Idx).Hidx).State <> sckconnecting Then
'Stop
'End If
'Call DisplayWinsock
                        Call CloseHandler(Idx)
                        Call OpenHandler(Idx)
                    End If
'This happens if there is a delay in a Listening Connection
'making the actual connection to the stream. I think Windows TCP stack
'goes into a WAIT state then times out removing the actual netstat
'socket
                Case Is = sckConnectionPending
                    If Sockets(Idx).Winsock.Server = 1 Then 'Server
                        Call CloseHandler(Idx)
                        Call OpenHandler(Idx)
                    End If
'Stop
                End Select
'Update the Socket state after trying to re-open handler
                Sockets(Idx).State = WinsockState(Idx)

'Serial check the Comm state as well as the Sockets state (Woodsons)
            Case Is = 1     'Serial
                If Hidx > 0 Then
                    If Comms(Hidx).State <> 1 Then
                        Call OpenHandler(Idx)
                    End If
                Else
                    If Sockets(Idx).State <> 1 Then
                        Call OpenHandler(Idx)
                    End If
                End If
'File dont try re-opening a file if in error
            Case Is = 2     'File

'Don't try reopening TTY window if user has closed it
            Case Is = 3     'TTY
                If Hidx > 0 Then
                    If TTYs(Hidx).TTY.Visible = False Then
                        TTYs(Hidx).TTY.Visible = True
                        Call RestoreWindow(TTYs(Hidx).TTY)
                    End If
                End If
            Case Else
'try reconnecting other handlers
                If Sockets(Idx).State <> 1 Then
                    Call OpenHandler(Idx)
                End If
            End Select
            
'End of all enabled sockets
        Else
'These sockets are disabled
'Stop
        End If
    Next Idx
        
'Remove any closed TCP Streams
    For Idx = 1 To UBound(Sockets)
        If IsTcpStream(Idx) Then
            Select Case WinsockState(Idx)
            Case Is = sckClosed
                Call RemoveSocket(Idx)
                Call UpdateMshRows
            End Select
        End If
                    
        If IsTcpClient(Idx) Then 'TCP/IP
            Select Case WinsockState(Idx)
            Case Is = sckConnected  ', sckConnecting

'Reset any emabled sockets that have not received any data for last
'5 reconnect cycles
'If running
                If cmdStop.Enabled = True Then
'WriteLog "ResetCount " & Sockets(Idx).ResetCount & " " & Sockets(Idx).DevName
                    Select Case Sockets(Idx).ResetCount
                    Case Is = 3, 8, 16, 32, 64
WriteLog "Resetting " & Sockets(Idx).DevName
                        Call Routecfg.RemoveSocketForwards(Idx)
                        Call CloseHandler(Idx)
                        Call Routecfg.CreateSocketForwards(Idx)
                        Call OpenHandler(Idx)
                    Case Is = 65
                        Sockets(Idx).ResetCount = 0
                    Case Else
                    End Select
                    If Sockets(Idx).Chrs = 0 Then
                        Sockets(Idx).ResetCount = Sockets(Idx).ResetCount + 1
                    Else
                        Sockets(Idx).ResetCount = 0
                    End If
                End If
            End Select
        End If
    Next Idx

'Output & Clear any Rejected connections
    Call OutputRejectionLog
    
'If at the end of the Reconnect
'when the Sockets(Idx).Status is OK (Open) reset the TryCount
    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).TryCount > 0 Then
'MsgBox WinsockState(Idx)
            Select Case Sockets(Idx).State
            Case Is = sckOpen, sckListening, sckConnected
                Call ResetTries(Idx)
            End Select
        End If
    Next Idx

    If frmRouter.GraphTimer.Enabled = True Then
        mshSockets.ColWidth(8) = 800
        If WorkbookExists = False Then
            Call CloseExcel
        End If
    End If

'Reset interval to 10 secs (on cmdStart we set it to 500 initially)
    ReconnectTimer.Interval = 10000 '10 secs
'update the display
    Call StatsTimer_Timer
End Sub

'Removes any stream sockets which have been closed
'or if the Owner has been closed
Public Function RemoveStreamSockets()
Dim Idx As Long
Dim IdxListen As Long

    For Idx = 1 To UBound(Sockets)
        If IsTcpStream(Idx) Then
'If socket is allocated update state (important as it may have changed)
            Select Case WinsockState(Idx)
            Case Is = sckClosed, sckError, sckClosing
                Call RemoveSocket(Idx)
                Call UpdateMshRows
            End Select
        End If
    Next Idx

'Remove any TCP Streams if the OwnerSocket is not Listening
'or making a connection
    For Idx = 1 To UBound(Sockets)
        If IsTcpStream(Idx) Then
            IdxListen = Sockets(Idx).Winsock.Oidx
            Select Case WinsockState(IdxListen)
            Case Is = sckListening, sckConnectionPending
            Case Else
'This the case where the Owner of the stream has been closed
                Call RemoveSocket(Idx)
                Call UpdateMshRows
            End Select
        End If
    Next Idx
End Function

Private Sub TryOpenHandler_old(Idx As Long)

    WriteLog "Trying to open " & Sockets(Idx).DevName
    Sockets(Idx).TryCount = Sockets(Idx).TryCount + 1
    Call OpenHandler(Idx)
End Sub

'To display correct info make sure state is set before calling
Public Sub ResetTries(Idx As Long)
Dim kb As String

'    Call DisplayTries("Reset Tries", Idx)
    With Sockets(Idx)
        If .TryCount > 0 Then
            kb = aState(.State) & " " & .DevName _
            & ", " & .TryCount
            If .TryCount = 1 Then
                kb = kb & " Open attempted"
            Else
                kb = kb & " Opens attempted"
            End If
            Sockets(Idx).TryCount = 0
            WriteLog kb
        End If
    End With
End Sub

'Used for debugging timers
Private Sub StatusBarTimer_Timer()
Dim kb As String
Dim ctrl As Control

    For Each ctrl In frmRouter
        If TypeOf ctrl Is Timer Then
            If ctrl.Enabled = True Then
                kb = kb & ctrl.Name
                On Error Resume Next
                kb = kb & ctrl.Index
                On Error GoTo 0
                kb = kb & "."
            End If
'            ctrl.Enabled = False
        
        End If
    Next ctrl
    frmRouter.StatusBar.Panels(1).Text = kb

End Sub

Public Sub ClearStatusBarTimer_Timer()
    StatusBar.Panels(1).Text = ""
    ClearStatusBarTimer.Enabled = False
'Debugging memory leak
'frmRouter.StatusBar.Panels(1).Text = Format$(GetWorkingSetSize \ 1024, "#,###,###") & " Kb"
'Call SourceDuplicateFilter.Status
End Sub

Private Sub GraphTimer_Timer()
'Interval is set as 60000 = 1 min
Dim Data As String
Dim Header As String
Dim Idx As Long

'Ensure we dont try updating if interval is 0
    If ExcelUpdateInterval = 0 Then
        GraphTimer.Enabled = False
        Exit Sub
    End If
    
    If MenuViewInoutGraph.Checked = False Then
        GraphTimer.Enabled = False
        Exit Sub
    End If
        
    ExcelLastUpdate = ExcelLastUpdate + 1
    If ExcelLastUpdate = ExcelUpdateInterval Then
        
        Call AddSheetRow
        ExcelLastUpdate = 0
    End If
End Sub


Private Sub Winsock_Error(Hidx As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Sockets(Winsocks(Hidx).sIndex).errmsg = Description
End Sub

'hwnd=WindowHandle, hShou=True or False to show in task bar
Public Sub ShowInTheTaskbar(hwnd As Long, bShow As Boolean)
Dim lStyle As Long
       
    ShowWindow hwnd, SW_HIDE
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle
    App.TaskVisible = bShow
    ShowWindow hwnd, SW_NORMAL
End Sub
     
'Not used
Private Function IsVisibleInTheTaskbar(hwnd As Long) As Boolean
Dim lStyle As Long
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If lStyle And WS_EX_APPWINDOW Then
        IsVisibleInTheTaskbar = True
    End If
End Function
     
Private Sub cmdToggle_Click()    'not used
Static bVisible As Boolean
       
    bVisible = Not IsVisibleInTheTaskbar(Me.hwnd)
    ShowInTheTaskbar Me.hwnd, bVisible
End Sub

' Arrange the controls on the form.
Public Sub ArrangeControls()
Dim hgt1 As Single  'Sockets
Dim hgt2 As Single  'Routes
Dim hgt3 As Single  'TextTerm
Dim wid1 As Single
Dim wid2 As Single
Dim StaticHeight As Single
Dim PaneCount As Long
Dim LastPane As String
Static LastHeight As Single
Static LastHgt1 As Single
Static LastHgt2 As Single
Static LastHgt3 As Single
Static LastWid1 As Single
    
    PercentageHorizontal = 1
'Debug.Print Time
' Don't bother if we're iconized.
    If WindowState = vbMinimized Then
'Remove from task bar before creating Systray
        If cmdSysTray = True Then
            Load frmSysTray
'Hide the minimised window otherwise the top bar only
'will be displayed on the desktop
            Me.Hide
        End If
        Exit Sub
    Else
'Not actioned if frmSystray is not loaded
        If cmdSysTray = True Then
            Unload frmSysTray
        End If
    End If
    
    StaticHeight = tlbRouter.Height + StatusBar.Height
'Debug.Print mshSockets.Height & ":" & IsHScrollVisible(mshSockets.hwnd)
    
'Tried limiting minimum form size when resizing but
'all methods cause problems (see Internet)
'So just exit the routine as even if we dont nothing will
'get changed
    If ScaleHeight - StaticHeight <= 0 Then Exit Sub

'Get the no of visible panes and the last one to use
'as the "balancing" pane
    mshSockets.Visible = MenuViewInoutSockets.Checked
    If mshSockets.Visible = True Then
        PaneCount = PaneCount + 1
        LastPane = "hgt1"
    End If
    mshRoutes.Visible = MenuViewInoutRoutes.Checked
    If mshRoutes.Visible = True Then
        PaneCount = PaneCount + 1
        LastPane = "hgt2"
    End If
    txtTerm.Visible = MenuViewInoutData.Checked
    If txtTerm.Visible = True Then
        PaneCount = PaneCount + 1
        LastPane = "hgt3"
    End If
    
'Split the visible pane sizes to approx equal sizes
    Select Case PaneCount
    Case Is = 1
        PercentageVertical = 1
    Case Is = 2
        PercentageVertical = 0.5
'This is the special case where only Sockets & Routes are visible
'We don't want the scroll bar on sockets to appear at 50% but when
'the height of both grids exceeds the available height (scale height)
'Therefore the Sockets want to be
        If LastPane = "hgt2" Then
            PercentageVertical = GridHeight(mshSockets) / (GridHeight(mshRoutes) + GridHeight(mshSockets))
        End If
   Case Is = 3
        PercentageVertical = 0.3
    Case Else
        PercentageVertical = 0
    End Select
         
'Calculate 1 2 & 3 BUT they will not add up to 100
'which is why we have to use the last one to keep the total
'size correct
    If mshSockets.Visible = True Then
        hgt1 = (ScaleHeight - StaticHeight) * PercentageVertical

'Dont make the pane smaller than it would be with one entry in
'the grid, this ensures any scroll bar will remain visible
        If hgt1 < hgt1Min Then hgt1 = hgt1Min
'If the grid is smaller than 30% reduce to remove blank surrounding
'Horizontal Scroll bars must be disabled - if visible they take more space
        If GridHeight(mshSockets) < hgt1 Then
'scroll bars not visible
            hgt1 = GridHeight(mshSockets)
        End If
    End If
    
    If mshRoutes.Visible = True Then
        hgt2 = (ScaleHeight - StaticHeight) * PercentageVertical
'        If hgt2 < 0 Then hgt2 = 0
        If hgt2 < hgt2Min Then hgt2 = hgt2Min
        If GridHeight(mshRoutes) < hgt2 Then
'scroll bars not visible
            hgt2 = GridHeight(mshRoutes)
        End If
    End If
    
    If txtTerm.Visible = True Then
        hgt3 = (ScaleHeight - StaticHeight) * PercentageVertical
'        If hgt3 < 0 Then hgt3 = 0
        If hgt3 < hgt3Min Then hgt3 = hgt3Min
    End If
    
'Set the Last Pane to the balance height
    Select Case LastPane
    Case Is = "hgt1"
        hgt1 = (ScaleHeight - StaticHeight) - hgt2 - hgt3
    Case Is = "hgt2"
        hgt2 = (ScaleHeight - StaticHeight) - hgt1 - hgt3
    Case Is = "hgt3"
        hgt3 = (ScaleHeight - StaticHeight) - hgt1 - hgt2
    Case Else
'no panes visible, in this case there will be a "blank"
'background shown - similar to closing all sheets in excel
'You cannot get rid of this becuase altering the form size
'would cause a resize loop
    End Select
    
'Recheck No -ve heights as they may have changed and
'a -ve height in a move is invalid
    If hgt1 < 0 Then hgt1 = 0
    If hgt2 < 0 Then hgt2 = 0
    If hgt3 < 0 Then hgt3 = 0
    wid1 = ScaleWidth
    
'StatusBar.Panels(1).Text = ScaleHeight - StaticHeight & "-" & PercentageVertical & "%:" & hgt1 & ":" & hgt2 & ":" & hgt3

'Stop flicker if nothing changed
    If LastHeight <> tlbRouter.Height Or _
    LastHgt1 <> hgt1 Or _
    LastHgt2 <> hgt2 Or _
    LastHgt3 <> hgt3 Or _
    LastWid1 <> wid1 Then
        mshSockets.Move 0, tlbRouter.Height, wid1, hgt1
        mshRoutes.Move 0, tlbRouter.Height + hgt1, wid1, hgt2
'Remaining height
        txtTerm.Move 0, tlbRouter.Height + hgt1 + hgt2, wid1, hgt3
        LastHeight = tlbRouter.Height
        LastHgt1 = hgt1
        LastHgt2 = hgt2
        LastHgt3 = hgt3
        LastWid1 = wid1
    End If
End Sub


Private Function GridHeight(Grid As MSHFlexGrid) As Single
Const RowPadding = 15   'dont know why but it is
'Const HScrollHeight = 280

    With Grid
        GridHeight = 4 * RowPadding + (.CellHeight + RowPadding) * (.Rows)
'Not used in this application
'        If IsHScrollVisible(Grid.Hwnd) = True Then
'            GridHeight = GridHeight + HScrollHeight
'        End If
End With
End Function

Public Sub VDOTimer_Timer()
Dim Idx As Long
Dim ReEnable As Boolean

    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).OutputFormat.OwnShipMmsi <> "" Then
            With Sockets(Idx).VDO
'If Ownship set up on more than 1 socket the
'sentence to be spoofed my not yet have been received
'on the other socket
                If .Destination Then
                    Call Queue(.SequenceNo, .Source, .UtcUnix, .Destination, .Data, "VDO=Spoof")
                    .LastVdoUpdate = .LastVdoUpdate + 1
                    If .LastVdoUpdate < 60 Then  '3 mins
                        ReEnable = True
                    End If
                End If
            End With
        End If
    Next Idx
    VDOTimer.Enabled = ReEnable
End Sub
