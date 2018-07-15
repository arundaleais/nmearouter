VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Commcfg 
   Caption         =   "Serial Port Configuration"
   ClientHeight    =   2625
   ClientLeft      =   1515
   ClientTop       =   1920
   ClientWidth     =   3720
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
   Icon            =   "Commcfg.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   3720
   Begin VB.Frame Frame1 
      Caption         =   "Data Direction"
      Height          =   885
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3495
      Begin VB.OptionButton optDirection 
         Caption         =   "Input,Output"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Output"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.ComboBox cboBaudRate 
      Height          =   315
      ItemData        =   "Commcfg.frx":058A
      Left            =   1320
      List            =   "Commcfg.frx":05AC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox cboCommName 
      Height          =   315
      ItemData        =   "Commcfg.frx":05F6
      Left            =   1320
      List            =   "Commcfg.frx":05F8
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Baud Rate"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Device"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Commcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurrentBaudRate As Long
Private CurrentDeviceName As String
Dim Hidx As Long
'Because this form is modal, the value cannot be changed once the form
'is displayed - it is cleared when the form is unloaded

Private Sub cboCommName_Click()
    StatusBar.Panels(1).Text = FriendlyName(cboCommName.Text)

End Sub

'Private CurrentIdx As Integer
'Private Cancel As Boolean

Private Sub cmdCancel_Click()
Dim kb As String
    kb = Sockets(CurrentSocket).DevName
    Set Comms(Hidx) = Nothing
    Unload Me
End Sub

' The port to use and configuration may have changed
Private Sub cmdOK_Click()
Dim ret As Long


'This is set on the form but allows text to be entered in text box
'It cannot be set at run time
'    cboCommName.Style = vbComboDropdown
'Have we a name
    If cboCommName.Text = "" Then
        MsgBox "DevName cannot be blank"
        Exit Sub
    End If
        
'extend Comm array if required
'If array not intialised create initial entry (=nothing)
'You need to stop the timer because if the Polling tries to read the
'Comm array while it is being redimensioned it causes a subscript error here !!!
'    Hidx = CLng(txtHidx)
    Sockets(CurrentSocket).Comm.Name = cboCommName.Text
    Sockets(CurrentSocket).Comm.BaudRate = cboBaudRate.Text
    Sockets(CurrentSocket).Direction = CurrentDirection
    Sockets(CurrentSocket).Comm.VCP = GetVCP(cboCommName.Text)
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
Dim arry() As String
'Dim NameValueKey As String
'Dim NameValueCount As Long
'Dim Names() As Variant  'must be for passing as array argument
'Dim Values() As Variant
Dim k As Long
Dim i As Long
Dim Ports() As String
Dim PortCount As Long

'    On Error GoTo Load_Error
'Debug.Print "CommLoad-Enter"
'If Comm array not yet initialised create a minimum of 1
'Otherwise UBound(comms) will fail
'Stop
'You need to stop the timer because if the Polling tries to read the
'Comm array while it is being redimensioned it causes a subscript error here !!!
    frmRouter.PollTimer.Enabled = False

'Display any current Comms values that are only
'on the handler (before we close it)
'If values are on Sockets then its OK
    If Sockets(CurrentSocket).Hidx > 0 Then
        Hidx = Sockets(CurrentSocket).Hidx
'Disable Sockets (must use close handler to set
'Sockets() up correctly
        Call CloseHandler(CurrentSocket)
    Else
        Hidx = FreeComm
'Set defaults
'WRONG if the handler is closed Hidx=0 but direction is already set
'on Sockets when it was loaded
'Direction is got from initial setup on frmComm
        If Sockets(CurrentSocket).Direction = -1 Then
            Sockets(CurrentSocket).Direction = CurrentDirection
        End If
    
    End If
    
    frmRouter.PollTimer.Enabled = True  'before Exit Sub
    
    If Hidx = -1 Then
        MsgBox "No free Serial Sockets", , "Commcfg.Load"
        Exit Sub
    End If
    
'Set Direction option to same as in Sockets()
    optDirection(Sockets(CurrentSocket).Direction).Value = True
    
'Set up Comm port List
'    Ports = GetAvailablePorts("")  'api method
    Ports = GetSerialPorts      'registry method
    On Error Resume Next
    PortCount = UBound(Ports) + 1
    On Error GoTo 0
    If PortCount = 0 Then
        MsgBox "There are no PC Serial Ports", , "Commcfg.Load"
        Exit Sub
    End If
    For k = 0 To UBound(Ports)
        cboCommName.AddItem Ports(k)
    Next k
    cboCommName.ListIndex = 0  'default
 
    
'Think its now always here as we are closing Comm(Hidx)
    If Comms(Hidx) Is Nothing Then
'Here when the Handler is first selected
        cboBaudRate.ListIndex = 1   '4800
'        cboCommName.Text = UCase(Sockets(CurrentSocket).DevName)
'Allow the Comm Port to be selected
        cboCommName.Enabled = True
    End If
'The handler will be nothing if the socket WAS disabled
'So we need to check if handler info is on Sockets()
    If Sockets(CurrentSocket).Handler = 1 Then  'Comms handler
        For i = 0 To cboBaudRate.ListCount - 1
            If cboBaudRate.List(i) = Sockets(CurrentSocket).Comm.BaudRate Then
                cboBaudRate.ListIndex = i
                Exit For
            End If
        Next i
        For i = 0 To cboCommName.ListCount - 1
            If cboCommName.List(i) = Sockets(CurrentSocket).Comm.Name Then
                cboCommName.ListIndex = i
'Dont allow user to change Comm name (if we have one)
'v31 allow                cboCommName.Enabled = False
                Exit For
            End If
        Next i
'Dont allow user to change Comm name (if we have one)
'We will have one if were are editing an existing handler
'and wont if its a new handler
'        If Sockets(CurrentSocket).Comm.Name <> "" Then
'            cboCommName.Enabled = False
'        End If
'1st desination only (at the moment)
'        txtForward = Comms(Sockets(CurrentSocket).Hidx).Destination(1)
    End If
    ' Set option button to current device
    StatusBar.Panels(1).AutoSize = sbrContents
    StatusBar.Panels(1).Text = FriendlyName(cboCommName.Text)
'Shown by Socketcfg (which is why were in the load routine)
'    Me.Show vbModal
'Debug.Print "CommLoad-Exit" & Me.Visible
    Exit Sub
Load_error:
'Stop
'    frmRouter.PollTimer.Enabled = False
'    Sleep 1000
'    Resume
    MsgBox err.Number & " " & err.Description
    End Sub

Private Function ShortDevName(DeviceName As String) As String
Dim arry() As String
        arry = Split(DeviceName, "\")
        ShortDevName = arry(UBound(arry))
End Function

'Ensures the Comm array includes this Idx (which is the same as the socket).
'If not extends the array to include this socket and returns true
'If we can't create this comms(Idx) returns false
'Comm is an array of Objects so set comms(Idx) = Object
'Must be used to set up the Object before any values can be added
Private Function SetComm(Idx As Long) As Boolean
Dim i As Long
'Dim OldUbound As Long
    If Idx <= MAX_SOCKETS Then
        If UBound(Comms) < Idx Then
'            OldUbound = UBound(comms)
            ReDim Preserve Comms(1 To Idx)
'            For i = OldUbound To UBound(comms)
'                comms(i).State = -1   'Not allocated
'            Next i
        End If
        SetComm = True
    End If
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmRouter.PollTimer.Enabled = True
End Sub

'Returns the Direction from optDirection
Private Function CurrentDirection() As Long
Dim i As Integer
    For i = 0 To optDirection.UBound
        If optDirection(i).Value = True Then
            CurrentDirection = i
            Exit For
        End If
    Next i
End Function

