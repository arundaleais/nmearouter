VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Socketcfg 
   Caption         =   "Socket Configuration"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   Icon            =   "Socketcfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Output Only Format"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
      Begin VB.TextBox txtOwnShipMMSI 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkPlainNmea 
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkIECEnabled 
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "OwnShip MMSI"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Type of Connection"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Only NMEA"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Add IEC 61162-1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkEnabled 
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox cboHandler 
      Height          =   315
      ItemData        =   "Socketcfg.frx":058A
      Left            =   1800
      List            =   "Socketcfg.frx":059D
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox cboDevName 
      Height          =   315
      ItemData        =   "Socketcfg.frx":05CC
      Left            =   1800
      List            =   "Socketcfg.frx":05CE
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3405
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Enabled"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Type of Connection"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Connection Name"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Socketcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'These form level variables re used to determine if the
'Values in Sockets() need reloading into the Display
Dim SelectedIdx As Long 'The one selected in the form


Private Sub cboDevName_Click()
Dim Cancel As Boolean
'Display the details (if existing) for this socket
'We have to detect the click as the validate is not actioned
'until the field is left.
'Also note that we need to detect any change in the Device
'in order to update the detail displayed from Sockets()
'Only validate when the form is loaded
    If Me.Visible Then
        Call cboDevName_Validate(Cancel)
        If Cancel = False Then
            If SelectedIdx = 0 Then
'For some reason SelectedIdx can be 0 here
                SelectedIdx = FreeSocket
            End If
            Call SocketToEditFields(SelectedIdx)
        End If
    End If
End Sub

Private Sub cboDevName_Validate(Cancel As Boolean)
Dim i As Integer
Dim arry() As String
Dim Idx As Long
    
    If cboDevName.Text = "" Then
        MsgBox "You must enter a Connection Name"
        Cancel = True
        Exit Sub
    End If

    arry = Split(Caption)
    Idx = DevNameToSocket(cboDevName.Text)    'returns 0 if not found
    Select Case arry(0)
    Case Is = "New"
'Check Device name not already in use
        If Idx > 0 Then
            MsgBox cboDevName.Text & " already exists"
            Cancel = True
            Exit Sub
        End If
    Case Is = "Open"
        If Idx = 0 Then
            MsgBox cboDevName.Text & " not found"
            Cancel = True
            Exit Sub
        End If
    Case Is = "Delete"
        If Idx = 0 Then
            MsgBox cboDevName.Text & " not found"
            Cancel = True
            Exit Sub
        End If
    Case Else
        MsgBox "Invalid Menu Option " & arry(0)
        Cancel = True
        Exit Sub
    End Select
    SelectedIdx = Idx
End Sub

Private Sub cboHandler_Validate(Cancel As Boolean)

'Forwarding is dependant on the handler
        Select Case cboHandler.ListIndex
        Case Is = -1   'Undetermined
            MsgBox "You must specify a Handler"
            Cancel = True
        Case Is = 0     'Winsock not cancelled
        Case Is = 1     'Serial Note not cancelled
        Case Is = 2     'File
'            MsgBox cboHandler.List(cboHandler.ListIndex) & " not yet avaialable"
'            Cancel = True
        Case Is = 3     'TTY Note not cancelled
        Case Is = 4     'LoopBack Note not cancelled
        Case Else
            MsgBox "Invalid Type of Connection"
            Cancel = True
        End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Cancel As Boolean
Dim Hidx As Long
Dim arry() As String
Dim OldDirection As Long
Dim NewDirection As Long
Dim ret As Integer
Dim msg As String

    
    On Error GoTo cmdOK_error
'we have to validate the fields
    Call ValidateControls
    Call cboDevName_Validate(Cancel)
    If Cancel = True Then Exit Sub
    Call cboHandler_Validate(Cancel)
    If Cancel = True Then Exit Sub
    arry = Split(Caption)
    Select Case arry(0)
    Case Is = "New"
        SelectedIdx = FreeSocket
        WriteLog "Socket " & FreeSocket & " allocated"
    Case Is = "Open"
        SelectedIdx = DevNameToSocket(cboDevName.Text)    'returns 0 if not found
    Case Is = "Delete"
        SelectedIdx = DevNameToSocket(cboDevName.Text)    'returns 0 if not found
        Call QueryRouteDelete(SelectedIdx, Cancel)
        If Cancel Then Exit Sub
        Call RemoveSocket(SelectedIdx)
        Unload Me   'Must be unloaded to reset Combo values when re-loaded
'Create a new excel workbook (if one is open)
            ExcelOpen = False
        Exit Sub
    Case Else
        MsgBox arry(0) & " not found"
        Unload Me   'Must be unloaded to reset Combo values when re-loaded
        Exit Sub
    End Select
    
'Code is now common for New and Open

'Get the details for this particular handler and open the handler
'We only set up the Forward Socket when the actual socket is set up
'Set all values in Sockets (Routes() are set up by validate)
    Sockets(SelectedIdx).DevName = cboDevName.Text
    Sockets(SelectedIdx).Handler = cboHandler.ListIndex
'13 Aug Cannot reset state (if paused file input state will be 1)
'and we need to know this to stop file name being changed
'    Sockets(SelectedIdx).State = 0  'Created but port is closed
    Sockets(SelectedIdx).errmsg = ""
    If chkIECEnabled.Value = vbChecked Then
        Sockets(SelectedIdx).IEC.Enabled = True
    Else
        Sockets(SelectedIdx).IEC.Enabled = False
    End If
    If chkPlainNmea.Value = vbChecked Then
        Sockets(SelectedIdx).OutputFormat.PlainNmea = True
    Else
        Sockets(SelectedIdx).OutputFormat.PlainNmea = False
    End If
    
    Sockets(SelectedIdx).OutputFormat.OwnShipMmsi = txtOwnShipMMSI.Text
    
    If chkEnabled.Value = vbUnchecked Then
        Sockets(SelectedIdx).Enabled = False
'Check if we have complete setup data
        If Socket_Validate(SelectedIdx) = True Then
'            If Sockets(SelectedIdx).State > 0 Then
'Close anyway as Winsock state may say closed when it isnt
                Call CloseHandler(SelectedIdx)
'            End If
'If the Socket is disabled
'We dont want to enter the Handler config screen
'but we must ensure the Handler is closed
'We can release all handler resources Handler(Hidx)
'Set Sockets(Idx).Hidx=0
'If it is re-enabled then OpenHandler(Idx) will re-create it
'v33 changed to enter config screen even if connection is closed
            Unload Me
            Exit Sub
        End If
    Else
        Sockets(SelectedIdx).Enabled = True
    End If
'From here only Enabled sockets
'or sockets requiring more data
'Pass socket no to handler
    CurrentSocket = SelectedIdx
    
'Keep Direction to detect if changed
    OldDirection = Sockets(SelectedIdx).Direction
    
'Stop subscript error in Timer as Sockets(IDX) is set up
    frmRouter.ReconnectTimer.Enabled = False

'Display the correct handler setup form
    Select Case Sockets(SelectedIdx).Handler
    Case Is = 0
'This is a modal form so we must close it to display the Handler Setup form
'As we continue after Commcfg is closed we only hide this form
        Me.Hide
        Winsockcfg.Show vbModal
'COMM HANDLER
    Case Is = 1     'COMM
        Me.Hide
        Commcfg.Show vbModal
'FILE HANDLER
    Case Is = 2     'File
        Me.Hide
        Filecfg.Show vbModal
'MsgBox "file not yet available", , "Socket Config"
    
    Case Is = 3     'TTY
        Me.Hide
        TTYcfg.Show vbModal
    Case Is = 4     'LoopBack
'This sets up the HandlerIndex and any handler details
'which would normally be done if there was a
'cfgLoopBack form (but is a Subroutine below)
        Call SpoofShowLoopBack
     Case Else
    End Select
    
    arry = Split(Me.Caption)
        
'The Socket and Handler is opened/changed
'Now using the settings that are in sockets()
    If Socket_Validate(SelectedIdx) = True Then
        If Sockets(SelectedIdx).Enabled = True Then
            Call OpenHandler(SelectedIdx)
        End If
    End If
    
'If new and weve failed to create a handler delete
'the socket
    Select Case arry(0)
    Case Is = "New"
        If Sockets(SelectedIdx).Hidx = -1 Then
            ret = MsgBox("Remove " & Sockets(SelectedIdx).DevName & " ? " & vbCrLf _
            & "(" & Sockets(SelectedIdx).errmsg & ")" & vbCrLf & vbCrLf _
            & "If OK and you later save this profile, the connection" & vbCrLf _
            & "will be permanently deleted" & vbCrLf & vbCrLf _
            & "If Cancel, the connection could be re-opened in this session" _
            , vbQuestion + vbOKCancel, Caption)
            If ret = vbOK Then
                Call ClearSocket(SelectedIdx)
            End If
        Else
'Create a new excel workbook (if one is open)
            ExcelOpen = False
        End If
'Handler has been cancelled, without being set-up
        If Sockets(SelectedIdx).Hidx = 0 Then
            Call ClearSocket(SelectedIdx)
        End If
'v62 close handler if new connection created when STOPed
        If frmRouter.cmdStart.Enabled = True Then
            Call CloseHandler(SelectedIdx)
        End If
    Case Is = "Open"
'If the direction has been changed - reset the forwarding on this socket
        If Sockets(SelectedIdx).Direction <> OldDirection Then
            Call Routecfg.RemoveSocketForwards(SelectedIdx)
            Call Routecfg.CreateSocketForwards(SelectedIdx)
        End If
'If Direction of DMZ changed to Input Remove DMZ
        If SelectedIdx = SourceDuplicateFilter.DmzIdx Then
            If Sockets(SelectedIdx).Direction = 1 Then
                MsgBox "The DMZ Connection (" & Sockets(SelectedIdx).DevName & ") has no Output" & vbCrLf _
                & "The DMZ will be changed to None" & vbCrLf & vbCrLf _
                & "To reconfigure the DMZ to another Connection" & vbCrLf _
                & "use Configure > Filters", vbInformation + vbOKOnly, "DMZ Connection"
                SourceDuplicateFilter.DmzIdx = 0
            End If
        End If
    End Select
    
'We must have a handler index at this point
'(if not will have exited above & socket will be deleted)

    Unload Me

'Check if any Routes on this socket are deactivated
'if so tell the user
    If Not IsTcpListener(SelectedIdx) Then
        Call Routecfg.InactiveSocketRoutes(SelectedIdx)
    End If
'After unloading this modal form, display the non modal form
'The Only one at present is TTY
    Call MakeFormsVisisble
    Exit Sub

cmdOK_error:
    On Error GoTo 0
    Select Case err.Number
    Case Else
    End Select
    If err.Number Then MsgBox err.Number & " " & err.Description
'This is a modal form so we must close it to display the Handler
    Unload Me
End Sub

Private Sub Form_Load()
Dim Idx As Long
Dim i As Long

'Must set the Caption before loading
    If CurrentSocket > 0 Then
'Load Current Socket Device Names
'The selection is done after the caption is set
'when we know New,Open or Delete
        For Idx = 1 To UBound(Sockets)
            If Sockets(Idx).State <> -1 Then
                cboDevName.AddItem Sockets(Idx).DevName
                If Idx = CurrentSocket Then
                    cboDevName.ListIndex = cboDevName.ListCount - 1
                    Call SocketToEditFields(Idx)
                End If
            End If
        Next Idx
    Else
        cboDevName.AddItem "Connection" & FreeSocket
        cboDevName.ListIndex = 0
        cboHandler.ListIndex = 1
        chkEnabled.Value = vbChecked
        chkIECEnabled.Value = vbUnchecked
        chkPlainNmea.Value = vbUnchecked
        txtOwnShipMMSI.Text = ""
    End If

End Sub

Private Sub SocketToEditFields(Idx As Long)
'Set field values from sockets()
Dim i As Integer
    For i = 0 To cboDevName.ListCount - 1
        If cboDevName.List(i) = Sockets(Idx).DevName Then
            cboDevName.ListIndex = i
            Exit For
        End If
    Next i
    For i = 0 To cboHandler.ListCount - 1
        If i = Sockets(Idx).Handler Then
            cboHandler.ListIndex = i
            Exit For
        End If
    Next i
    If Sockets(Idx).Enabled = True Then
        chkEnabled.Value = vbChecked
    Else
        chkEnabled.Value = vbUnchecked
    End If
    If Sockets(Idx).IEC.Enabled = True Then
        chkIECEnabled.Value = vbChecked
    Else
        chkIECEnabled.Value = vbUnchecked
    End If
    If Sockets(Idx).OutputFormat.PlainNmea = True Then
        chkPlainNmea.Value = vbChecked
    Else
        chkPlainNmea.Value = vbUnchecked
    End If
    txtOwnShipMMSI.Text = Sockets(Idx).OutputFormat.OwnShipMmsi
End Sub

'Called from frmRouter(Menu)
Public Sub SocketNew()
Dim Idx As Long

'    cboDevName.Clear
    Idx = FreeSocket
    If Idx > 0 Then
'Loading socketcfg will allocate a new socket
        CurrentSocket = 0
'Setting Caption loads the form
        Me.Caption = "New Connection"
        cboHandler.Enabled = True
        chkEnabled.Enabled = True
        chkIECEnabled.Enabled = True
        chkPlainNmea.Enabled = True
        txtOwnShipMMSI = ""
        Me.Show vbModal
    Else
        MsgBox "No available Connections", , "Socket New"
    End If
End Sub

'Called from frmRouter(Menu)
Public Sub SocketOpen(Optional ReqIdx As Long)
Dim i As Long
Dim Idx As Long

'Set the first socket in use as the current socket
    If ReqIdx = 0 Then
        For Idx = 1 To UBound(Sockets)
            If Sockets(Idx).State > -1 Then
                CurrentSocket = Idx
                Exit For
            End If
        Next Idx
    Else
        CurrentSocket = ReqIdx
    End If
    
    If CurrentSocket > 0 Then
'Setting Caption loads the form
        Me.Caption = "Open Connection"
'Hidx will be -1 if when first opened the handler could not be opened
'and the user had decided not to Remove the Socket
'because it they be open the socket later
'When they do have access to the handler (eg the port becomes available)
        If Sockets(CurrentSocket).Hidx = -1 Then
            cboHandler.Enabled = True
        Else
            cboHandler.Enabled = False
        End If
        chkEnabled.Enabled = True
        chkIECEnabled.Enabled = True
        chkPlainNmea.Enabled = True
        txtOwnShipMMSI.Enabled = True
        Me.Show vbModal
    Else
'The Open Option should not be displayed
        MsgBox "There are no Connections", , "Open Connection"
    End If
End Sub

'Called from frmRouter(Menu)
Public Sub SocketDelete(Optional ReqIdx As Long)
Dim i As Long
Dim Idx As Long
    
'Set the first socket in use as the current socket
    If ReqIdx = 0 Then
        For Idx = 1 To UBound(Sockets)
            If Sockets(Idx).State > -1 Then
                CurrentSocket = Idx
                Exit For
            End If
        Next Idx
    Else
        CurrentSocket = ReqIdx
    End If
    
    If CurrentSocket > 0 Then
'Setting Caption loads the form
        Me.Caption = "Delete Connection"
        cboHandler.Enabled = False
        chkEnabled.Enabled = False
        chkIECEnabled.Enabled = False
        chkPlainNmea.Enabled = False
        txtOwnShipMMSI.Enabled = False
        Me.Show vbModal
    Else
'The Open Option should not be displayed
        MsgBox "There are no Connections", , "Delete Connection"
    End If
        
End Sub
Private Sub GetSockets()
Dim Idx As Long
'Load Current Socket Device Names
    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).State <> -1 Then
            cboDevName.AddItem Sockets(Idx).DevName
        End If
    Next Idx
End Sub

'If there are Routes to this sockets Queries the Socket Delete
Public Sub QueryRouteDelete(ReqIdx As Long, Cancel As Boolean)
Dim Idx As Long
Dim Ridx As Long
Dim kb As String
Dim Count As Long
Dim SocketCount As Long
Dim RidxCount As Long
Dim ret As Integer
Dim Detail As Boolean   'Set True to display details

    For Idx = 1 To UBound(Sockets)
        RidxCount = 0
        With Sockets(Idx)
        For Ridx = 1 To UBound(.Routes)
            If ReqIdx = 0 Or Idx = ReqIdx Or .Routes(Ridx).AndIdx = ReqIdx Then
                If .Routes(Ridx).AndIdx > 0 Then
                    kb = kb & "Between " & Sockets(Idx).DevName & " [" _
                    & aDirection(Sockets(Idx).Direction) & "]"
                    If Detail = True Then
                        kb = kb & ", Idx(" & Idx & ")"
                    End If
                    kb = kb & " and " & Sockets(.Routes(Ridx).AndIdx).DevName _
                    & " [" & aDirection(Sockets(.Routes(Ridx).AndIdx).Direction) & "]"
                    kb = kb & vbCrLf
                    RidxCount = RidxCount + 1
                End If
            End If
        Next Ridx
        End With
    Count = Count + RidxCount
    Next Idx
    If Count = 0 Then
        Exit Sub
    End If
'Route count is a check the no of routes are correct
    If Count = 1 Then
        kb = "This Route will also be Deleted" & vbCrLf & vbCrLf & kb
    Else
        kb = "These Routes will also be Deleted" & vbCrLf & vbCrLf & kb
    End If
    ret = MsgBox(kb, vbOKCancel, "Delete Connection " & Sockets(ReqIdx).DevName)
    If ret = vbCancel Then
        Cancel = True
    End If
End Sub

'Creates a loopback socket without displaying a cfg window
Private Sub SpoofShowLoopBack()
Dim Hidx As Long

'Must have CurrentSocket set
    If Sockets(CurrentSocket).Hidx > 0 Then
        Hidx = Sockets(CurrentSocket).Hidx
'Display/Keep any current values that are only
'on the handler (before we close it)
'Disable Pollong Socket (if required)
'We must use close handler to set Sockets() up correctly
        Call CloseHandler(CurrentSocket)
    Else
'Ensure we can create the handler extending Handlers()
'if reqd
'Create a class module so we can call FreeLoopBack
        Hidx = FreeLoopBack
'Set defaults
'Direction is got from initial setup on frmHandler
        Sockets(CurrentSocket).Direction = 0
    End If
    
    If Hidx = -1 Then
        MsgBox "No free LoopBack Sockets", , "LoopBack.Load"
        Exit Sub
    End If
'Think its now always here as we are closing Comm(Hidx)
    If LoopBacks(Hidx) Is Nothing Then
'Here when the Handler is first selected
'So set up the Fields on the Form to defaults
'The handler will be nothing if the socket WAS disabled
'So we need to check if handler info is on Sockets()
    
'Set any Combo box choices on Form, with List index
'set to any values on sockets(CurrrentSocket).Handler

'Disable any ComboBoxes a user must not change

' Set option button to current device
    End If
'EXIT the LOAD
End Sub


Private Sub txtOwnShipMMSI_Validate(Cancel As Boolean)
    If txtOwnShipMMSI.Text <> "" Then
        If IsNumeric(txtOwnShipMMSI.Text) = False _
        Or Len(txtOwnShipMMSI) <> 9 Then
            MsgBox "MMSI must be numeric (9 digits)"
            Cancel = True
            Exit Sub
        End If
    End If
End Sub
