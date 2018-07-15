VERSION 5.00
Begin VB.Form Winsockcfg 
   Caption         =   "UDP/TCP"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   Icon            =   "Winsockcfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Connection"
      Height          =   1365
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   3495
      Begin VB.TextBox txtPermittedIPStreams 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtPermittedStreams 
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Server ""Listens"""
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Client ""Requests"""
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Max From Same IP"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Max Concurrent"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Direction"
      Height          =   885
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   3495
      Begin VB.OptionButton optDirection 
         Caption         =   "Output"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input,Output"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtLocalIP 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtLocalHostName 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Text            =   "0"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtRemoteHostIP 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "29421"
      Top             =   3840
      Width           =   855
   End
   Begin VB.ComboBox cboProtocol 
      Height          =   315
      ItemData        =   "Winsockcfg.frx":058A
      Left            =   1680
      List            =   "Winsockcfg.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host IP"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Local IP"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Local Host Name"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Port"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Local Port"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Protocol"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Winsockcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Note the Winsock(Index), Direction and Server when Form is loaded
'This is to allow the fields to be reset to the loaded state
Private Hidx As Integer
Private LastProtocol As Long
Private LastServer As Long

'Only when the form is first loaded (with a new edit) FieldsReset will be false

'Because this form is modal, the value cannot be changed once the form
'is displayed - it is cleared when the form is unloaded

Private Sub cboProtocol_Click()
    Call SetEditableFields(cboProtocol.ListIndex, CurrentServer, CurrentDirection)
End Sub

Private Sub cboProtocol_Validate(Cancel As Boolean)
    Call SetEditableFields(CurrentProtocol, CurrentServer, CurrentDirection)
End Sub

'Private CurrentIdx As Integer
'Private Cancel As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' The port to use and configuration may have changed
Private Sub CmdOk_Click()
Dim Cancel As Boolean
Dim arry() As String
   
'Create Winsock closes existing socket and allocate Hidx (if req)
'Either connection will validate
        Call optServer_Validate(0, Cancel)
        If Cancel = True Then Exit Sub
        Call txtLocalPort_Validate(Cancel)
        If Cancel = True Then Exit Sub
        Call txtRemoteHost_Validate(Cancel)
        If Cancel = True Then Exit Sub
        Call txtRemotePort_Validate(Cancel)
        If Cancel = True Then Exit Sub
        Call txtPermittedStreams_Validate(Cancel)
        If Cancel = True Then Exit Sub
        Call txtPermittedIPStreams_Validate(Cancel)
        If Cancel = True Then Exit Sub
        
'Load new settings into Sockets()
        With Sockets(CurrentSocket)
            .Winsock.Protocol = cboProtocol.ListIndex
            .Winsock.Server = CurrentServer
            .Direction = CurrentDirection
            .Winsock.LocalPort = txtLocalPort
            .Winsock.RemoteHost = txtRemoteHost
            .Winsock.RemotePort = txtRemotePort
            .Winsock.PermittedStreams = txtPermittedStreams
            .Winsock.PermittedIPStreams = txtPermittedIPStreams
        End With
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
'    On Error GoTo Load_Error
'With Winsock we use Winsock(0) to set the Fields if there is no Hidx
'Get Handler Index, if none allocate a handler
    If Sockets(CurrentSocket).Hidx > 0 Then
        Hidx = Sockets(CurrentSocket).Hidx
'Display any current Winsock values that are only
'on the handler (before we close it)
        txtLocalHostName = frmRouter.Winsock(Hidx).LocalHostName
        txtLocalIP = frmRouter.Winsock(Hidx).LocalIP
'Disable Winsock (must use close handler to set
'Sockets() up correctly
        Call CloseHandler(CurrentSocket)
    Else
        Hidx = FreeWinsock
'Set defaults
'Direction is got from initial setup on frmWinsock
        If Sockets(CurrentSocket).Direction = -1 Then
            Sockets(CurrentSocket).Direction = CurrentDirection
        End If
    End If
    
    If Hidx = -1 Then
        MsgBox "No free TCP/IP handlers", , "Winsockcfg.Load"
        Exit Sub
    End If
       
    With Sockets(CurrentSocket).Winsock
        txtLocalPort = .LocalPort
        txtRemoteHost = .RemoteHost
        txtRemoteHostIP = .RemoteHostIP 'Read Only
        txtRemotePort = .RemotePort
        If .Protocol <> -1 Then
            cboProtocol.ListIndex = .Protocol
        End If
        If .Server <> -1 Then
            optServer(.Server).value = True
        End If
        txtPermittedStreams = .PermittedStreams
        txtPermittedIPStreams = .PermittedIPStreams
        
    End With
    
'Set up initial fields on the form. On initial load use the defaults
'initiallially on the form for Server & Direction
'Set intial values on load
    
    Call SetEditableFields(Sockets(CurrentSocket).Winsock.Protocol, _
    Sockets(CurrentSocket).Winsock.Server, Sockets(CurrentSocket).Direction)
 'Set Direction option to same as in Sockets()
    optDirection(Sockets(CurrentSocket).Direction).value = True

    Exit Sub

Load_error:
    MsgBox err.Number & " " & err.Description, , "Winsockcfg.Load"
    End Sub

Private Sub optDirection_Click(Index As Integer)
'The NEW direction is NOT set
    Call SetEditableFields(CurrentProtocol, CurrentServer, CLng(Index))
End Sub

Private Sub optDirection_Validate(Index As Integer, Cancel As Boolean)
    Call SetEditableFields(CurrentProtocol, CurrentServer, CurrentDirection)
End Sub

Private Sub optServer_Click(Index As Integer)
'The NEW server is not yet set
    Call SetEditableFields(CurrentProtocol, CLng(Index), CurrentDirection)
End Sub

Private Sub optServer_Validate(Index As Integer, Cancel As Boolean)
    If cboProtocol.ListIndex = sckTCPProtocol Then
        If CurrentServer = -1 Then
            MsgBox "You must set a Connection", , "Validate Server"
            Cancel = True
        End If
    Call SetEditableFields(CurrentProtocol, CurrentServer, CurrentDirection)
    End If
End Sub

#If False Then
'complile error
Private Sub txtHidx_Click()
    currentIdx = CInt(txtHidx)
End Sub
#End If

Private Function ShortDevName(DeviceName As String) As String
Dim arry() As String
        arry = Split(DeviceName, "\")
        ShortDevName = arry(UBound(arry))
End Function

'The arguments are the currentValue (returned by CurrentProtocol etc)
'or if called by the click event the clicked value
Private Sub SetEditableFields(Protocol As Long, Server As Long, Direction As Long)

'Server =-1 when form loaded
    If Protocol <> LastProtocol Or Server = -1 Then
        Select Case Protocol
        Case Is = sckTCPProtocol
'Set Direction according to Server
'Set default as Protocol has been changed
            Select Case Direction
            Case Is = 1     'Input
                optServer(0) = True 'Client
            Case Is = 2     'Output
                optServer(1) = True 'Server
            End Select
        Case Is = sckUDPProtocol
'Set Server according to Direction
'Both is invalid = set default to input
            If optDirection(0) = True Then
                optDirection(1) = True
            End If
'v45 Digi definition UDP Client=0 (Outgoing Device=2)
'UDP Server=1 (Incoming Device=1, default)
            Select Case Direction   'v45 Digi definition UDP Client (Outgoing Device)
            Case Is = 1     'Input
                optServer(1) = True 'Server
            Case Is = 2     'Output
                optServer(0) = True 'Client
            End Select
        End Select
    End If
    
    If Server <> LastServer Then
        Select Case Protocol
        Case Is = sckTCPProtocol
            Select Case Server
                Case Is = 0 'Client
                    optDirection(1) = True
                Case Is = 1 'Server
                    optDirection(2) = True
'                    If txtPermittedStreams = "0" Then
'                        txtPermittedStreams = "1"
'                    End If
            End Select
        End Select
    End If

'Read Only Fields
    txtLocalHostName.Enabled = False
    txtLocalIP.Enabled = False

    
    Select Case Protocol
    Case Is = sckTCPProtocol
        
'If TCP allow user to change Server
optServer(0).Enabled = True
optServer(1).Enabled = True
optDirection(0).Enabled = True
optDirection(1).Enabled = True
optDirection(2).Enabled = True
        optDirection(0).Caption = "Input,Output"
        
        Select Case Server
 'TCP Client
        Case Is = 0
'Output Only
            txtRemoteHost.Enabled = True
            txtRemotePort.Enabled = True
            txtLocalPort.Enabled = False
            txtPermittedStreams.Enabled = False
            txtPermittedIPStreams.Enabled = False
'TCP Server
'Output Only, because we have to set a route to
'each Client stream using the Server Routing
        Case Is = 1
            txtLocalPort.Enabled = True
            txtRemotePort.Enabled = False
            txtRemoteHost.Enabled = False
            txtPermittedStreams.Enabled = True
            txtPermittedIPStreams.Enabled = True
        End Select
    
    Case Is = sckUDPProtocol

'If UDP allow user to change Direction not Server
optServer(0).Enabled = False
optServer(1).Enabled = False
optDirection(0).Enabled = False
optDirection(1).Enabled = True
optDirection(2).Enabled = True
        optDirection(0).Caption = "Invalid"
        txtPermittedStreams.Enabled = False
        txtPermittedIPStreams.Enabled = False
        
'If Protocol has changed, check Direction is valid
'If not set to default
        If CurrentProtocol <> LastProtocol Then

'Both is invalid = set default to input
            If optDirection(0) = True Then
                optDirection(1) = True
            End If
        End If

'If UDP set direction Server to direction
'v45 Digi definition UDP Client=0 (Outgoing Device=2)
'UDP Server=1 (Incoming Device=1, default)
        Select Case CurrentDirection
        Case Is = 1     'Input
            optServer(1) = True 'Server
        Case Is = 2     'Output
            optServer(0) = True 'Client
        End Select

'Bind listens on the port specified
        txtLocalPort.Enabled = False
        txtRemotePort.Enabled = False
        txtRemoteHost.Enabled = False
 'UDP Input
 'Only need to bind to the port to which the incoming data has been sent.
 'You can specify the local adaptor IP (when you BIND) if there are
 'multiple network adaptors. Bind opens the port.
        Select Case Direction
        Case Is = 1
            txtLocalPort.Enabled = True
'Output only
        Case Is = 2
            txtLocalPort.Enabled = False
            txtRemotePort.Enabled = True
            txtRemoteHost.Enabled = True
'Both
        Case Is = 0
            txtLocalPort.Enabled = True
            txtRemotePort.Enabled = True
            txtRemoteHost.Enabled = True
        End Select
    End Select  'UDP

'Set backgrounds
    With txtLocalHostName
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With
    With txtLocalIP
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With
    With txtLocalPort
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With
    With txtRemoteHost
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With
    With txtRemotePort
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With

    With txtPermittedStreams
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With

    With txtPermittedIPStreams
        If .Enabled = True Then
            .BackColor = vbWhite
        Else
            .BackColor = Me.BackColor
        End If
    End With

'Keep the current one to see if it has been changed
'Neat time SetEditableFields is called
    LastProtocol = CurrentProtocol
    LastServer = CurrentServer
End Sub

'The arguments are the currentValue (returned by CurrentProtocol etc)
'or if called by the click event the clicked value
Private Sub SetEditableFields_old(Protocol As Long, Server As Long, Direction As Long)

    txtLocalHostName.Enabled = False
    txtLocalHostName.BackColor = Me.BackColor
    txtLocalIP.Enabled = False
    txtLocalIP.BackColor = Me.BackColor
    txtLocalIP.Enabled = False
    txtLocalIP.BackColor = Me.BackColor
    Select Case Protocol
    Case Is = sckTCPProtocol
        optDirection(0).Caption = "Bi-directional"
        optServer(0).Enabled = True
        optServer(1).Enabled = True
        If optServer(0).value = optServer(1).value Then
            optServer(0).value = True
        End If
        Select Case Server
 'TCP Client
        Case Is = 0
'Output Only
            txtRemoteHost.Enabled = True
            txtRemoteHost.BackColor = vbWhite
            txtRemotePort.Enabled = True
            txtRemotePort.BackColor = vbWhite
            txtLocalPort.Enabled = False
            txtLocalPort.BackColor = Me.BackColor
'TCP Server
'Output Only, because we have to set a route to
'each Client stream using the Server Routing
        Case Is = 1
            optDirection(0).Enabled = False
            optDirection(1).Enabled = False
            optDirection(2) = True
            txtLocalPort.Enabled = True
            txtLocalPort.BackColor = vbWhite
            txtRemotePort.Enabled = False
            txtRemotePort.BackColor = Me.BackColor
            txtRemoteHost.Enabled = False
            txtRemoteHost.BackColor = Me.BackColor
        End Select
    
    Case Is = sckUDPProtocol
        optDirection(0).Caption = "Invalid"
        optServer(0).Enabled = False
        optServer(1).Enabled = False
        optServer(0).value = False
        optServer(1).value = False
'Bind listens on the port specified
        txtLocalPort.Enabled = False
        txtLocalPort.BackColor = vbWhite
        txtRemotePort.Enabled = False
        txtRemotePort.BackColor = Me.BackColor
        txtRemoteHost.Enabled = False
        txtRemoteHost.BackColor = Me.BackColor
 'UDP Input
 'Only need to bind to the port to which the incoming data has been sent.
 'You can specify the local adaptor IP (when you BIND) if there are
 'multiple network adaptors. Bind opens the port.
        Select Case Direction
        Case Is = 1
            txtLocalPort.Enabled = True
            txtLocalPort.BackColor = vbWhite
'Output only
        Case Is = 2
            txtLocalPort.Enabled = False
            txtLocalPort.BackColor = Me.BackColor
            txtRemotePort.Enabled = True
            txtRemotePort.BackColor = vbWhite
            txtRemoteHost.Enabled = True
            txtRemoteHost.BackColor = vbWhite
'Both
        Case Is = 0
            txtLocalPort.Enabled = True
            txtLocalPort.BackColor = vbWhite
            txtRemotePort.Enabled = True
            txtRemotePort.BackColor = vbWhite
            txtRemoteHost.Enabled = True
            txtRemoteHost.BackColor = vbWhite
        End Select
    End Select
End Sub

'Returns the currentProtocol from cboProtocol.ListIndex
'This is not really necessary as the click event returns the same
'but if makes the call to SetEditFields clearer
Private Function CurrentProtocol() As Long
    CurrentProtocol = cboProtocol.ListIndex
End Function
'Returns the Direction from optDirection
Private Function CurrentDirection() As Long
Dim i As Integer
    For i = 0 To optDirection.UBound
        If optDirection(i).value = True Then
            CurrentDirection = i
            Exit For
        End If
    Next i
End Function

'Returns the CurrentServer from optServer
'0=Client, 1=Server
Private Function CurrentServer() As Long
Dim i As Integer
'default has not been set
    CurrentServer = -1
    For i = 0 To optServer.UBound
        If optServer(i).value = True Then
            CurrentServer = i
        End If
    Next i
End Function

Private Sub txtLocalPort_Validate(Cancel As Boolean)
    If IsNumeric(txtLocalPort) = False Then txtLocalPort = "0"
    Select Case cboProtocol.ListIndex
    Case Is = sckTCPProtocol
        Select Case CurrentServer
        Case Is = 1     'Server
            If txtLocalPort = "0" Then
                MsgBox "Local Port must be set to the Port number" & vbCrLf & "the Client will contact this server"
                Cancel = True
            End If
        End Select
    Case Is = sckUDPProtocol
        Select Case CurrentDirection
        Case Is = 1     'input
            If txtLocalPort = "0" Then
                MsgBox "Local Port must be set to the Port number" & vbCrLf & "the Remote server is sending you data to"
                Cancel = True
            End If
        End Select
    End Select
End Sub

Private Sub txtPermittedStreams_Validate(Cancel As Boolean)
    Select Case cboProtocol.ListIndex
    Case Is = sckTCPProtocol
        Select Case CurrentServer
        Case Is = 1     'Server
            If IsNumeric(txtPermittedStreams) = False Then
                MsgBox "Concurrent Connections must be numeric"
                Cancel = True
            Else
                If CLng(txtPermittedStreams) > MAX_TCPSERVERSTREAMS Then
                    MsgBox "Maximum Concurrent Connections are limited to " & MAX_TCPSERVERSTREAMS _
                    & vbCrLf & "for each TCP Listener"
                    Cancel = True
                End If
                If CLng(txtPermittedStreams) < 1 Then
                    MsgBox "Concurrent Connections must be at least 1"
                    Cancel = True
                End If
            End If
        End Select
    End Select
End Sub

Private Sub txtPermittedIPStreams_Validate(Cancel As Boolean)
    Select Case cboProtocol.ListIndex
    Case Is = sckTCPProtocol
        Select Case CurrentServer
        Case Is = 1     'Server
            If IsNumeric(txtPermittedIPStreams) = False Then
                MsgBox "Concurrent Connections must be numeric"
                Cancel = True
            Else
                If CLng(txtPermittedIPStreams) > CLng(txtPermittedIPStreams) Then
                    MsgBox "Maximum Connections for each IP are limited to " & MAX_TCPSERVERSTREAMS _
                    & vbCrLf & "Maximum Concurrent Connections"
                    Cancel = True
                End If
                If CLng(txtPermittedIPStreams) < 1 Then
                    MsgBox "Maximum for each IP must be at least 1"
                    Cancel = True
                End If
            End If
        End Select
    End Select
End Sub

Private Sub txtRemoteHost_Validate(Cancel As Boolean)
    Select Case cboProtocol.ListIndex
    Case Is = sckTCPProtocol
        Select Case CurrentServer
        Case Is = 0     'Client
            If txtRemoteHost = "" Then
                MsgBox "Remote Host must be set to the" & vbCrLf & "IP address or Server Name of the" & vbCrLf & "Server you wish to connect to"
                Cancel = True
            End If
        End Select
    Case Is = sckUDPProtocol
        Select Case CurrentDirection
        Case Is = 2     'input
            If txtRemoteHost = "" Then
                MsgBox "Remote Host must be set to the" & vbCrLf & "IP address or Client Name of the" & vbCrLf & "Client you wish to send data to"
                Cancel = True
            End If
        End Select
    End Select
End Sub

Private Sub txtRemotePort_Validate(Cancel As Boolean)
    If IsNumeric(txtRemotePort) = False Then txtRemotePort = "0"
    Select Case cboProtocol.ListIndex
    Case Is = sckTCPProtocol
        Select Case CurrentServer
        Case Is = 0     'Client
            If txtRemotePort = "0" Then
                MsgBox "Remote Port must be set to the" & vbCrLf & "Port number of the" & vbCrLf & "Server you wish to connect to"
                Cancel = True
            End If
        End Select
    Case Is = sckUDPProtocol
        Select Case CurrentDirection
        Case Is = 2     'output
            If txtRemotePort = "0" Then
                MsgBox "Remote Port must be set to the" & vbCrLf & "Remote Port number of the" & vbCrLf & "Client you wish to send data to"
                Cancel = True
            End If
        End Select
    End Select
End Sub

