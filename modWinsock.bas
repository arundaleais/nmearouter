Attribute VB_Name = "modWinsock"
Option Explicit
Private Type WinsockDef
    sIndex As Long      'Sockets() index for this Handler
'    Pause As Boolean    'If true stops Winsock Timer trying to re-open a
                        'Connection while it is being edited
'    Direction As Long   '0 = Both
                        '1-Input
                        '2=Output
'    Server As Long      '-1=undefined
                        '0=Client
                        '1=Server
End Type

Public Winsocks() As WinsockDef

'ONLY must be called by CloseHandler
'Closes Winsock, Empties the buffer does not erase the buffer
Public Sub CloseWinsock(Hidx As Integer)
Dim Idx As Long
'Dim Sidx As Long

    On Error GoTo Winsock_err
    
'If a Multi-Stream TCP Connection
'We CANNOT Remove the Socket here
'as the program will loop round CloseHAndlet>CloseWindock>RemoveSocket>CloseHandler

'#If False Then
'Stop
'If a TCP stream RemoveSocketForwards will have
'set Sidx on the OwnerServer Socket to 0
'            Idx = Winsocks(Hidx).sIndex
'If TCP Server, remove Streams as well
'    If Sockets(Idx).Winsock.Protocol = 0 _
'    And Sockets(Idx).Winsock.Server = 1 Then
'        For Sidx = 1 To UBound(Sockets(Idx).Winsock.Streams)
'            If Sockets(Idx).Winsock.Streams(Sidx) > 0 Then
'                Call RemoveSocket(Sockets(Idx).Winsock.Streams(Sidx))
'            End If
'        Next Sidx
'    End If
'#End If

    If WinsockExists(Hidx) = True Then
        frmRouter.Winsock(Hidx).Close
        WriteLog "TCP/IP handler " & Hidx & ", Local Port " & frmRouter.Winsock(Hidx).LocalPort & " closed"
'        Idx = Winsocks(Hidx).sIndex
        Unload frmRouter.Winsock(Hidx)
'Once unloaded the remote IP:Port must be cleared from frmrouter(sockets)
'        Sockets(Idx).Winsock.RemoteHostIP = ""
'        Sockets(Idx).Winsock.RemotePort = 0
        End If
    Call frmRouter.UpdateMshRows

'Stop
'    If IsTcpListener(Winsocks(Hidx).sIndex) Then
'        Call frmRouter.RemoveStreamSockets(Winsocks(Hidx).sIndex)
'    End If
    
Exit Sub
Winsock_err:
    On Error GoTo 0
    Sockets(Winsocks(Hidx).sIndex).errmsg = err.Description
    MsgBox "CloseWinsock Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "Socket " & Winsocks(Hidx).sIndex, , "Close Winsock"
End Sub

Sub WinsockOutput(Hidx As Long, DataSnd As String)
Dim i As Long
Static Big As Long

'the Winsock Control element may not have had time to be created
'by the time it starts
'jna debug to find the error line
'    On Error GoTo SendData_err
    If Hidx > 0 Then
'the Port & Socket may have been closed by the user while there'
'were unsent Sentences in the buffer
        frmRouter.Winsock(Hidx).SendData DataSnd

'Test Big sentences
'        If Big < 10 Then
'            DataSnd = Left$(DataSnd, Len(DataSnd) - 2)
'            Big = Big + 1
'        Else
'            Big = 0
'        End If
'Test split sentences
'        For i = 1 To Len(DataSnd)
'            frmRouter.Winsock(Hidx).SendData Mid$(DataSnd, i, 1) & vbCr
'        Next i
    End If
Exit Sub
SendData_err:
    MsgBox "Send Data Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "Winsock Index " & Hidx, , "SendData"
End Sub

'Check to see if Winsock(index) exists, there seems to be no other way to check
'other than trying to access the index
Public Function WinsockExists(Hidx As Integer) As Boolean
    On Error GoTo NoWinsock
    If frmRouter.Winsock(Hidx).Index Then
        WinsockExists = True
    End If
NoWinsock:
End Function

Public Sub DisplayWinsock()
Dim result As Boolean
Dim ctrl As Winsock
Dim kb As String
Dim Count As Long
Dim Idx As Long
Dim Hidx As Long
Dim i As Long
Dim IdxStream As Long

'When winsock is closed winsock(Hidx) does not exist
    For Each ctrl In frmRouter.Winsock
'There is always a zero index that is not used
        If ctrl.Index > 0 Then
            For Idx = 1 To UBound(Sockets)
                If Sockets(Idx).Handler = 0 And Sockets(Idx).Hidx = ctrl.Index Then
                    kb = kb & Sockets(Idx).DevName & vbCrLf
                    Exit For
                End If
            Next Idx
            kb = kb & vbTab & "Index=" & ctrl.Index & vbCrLf
            kb = kb & vbTab & "Local IP=" & ctrl.LocalIP & vbCrLf
            kb = kb & vbTab & "Local Port=" & ctrl.LocalPort & vbCrLf
            kb = kb & vbTab & "Protocol=" & aProtocol(ctrl.Protocol) & vbCrLf
            kb = kb & vbTab & "Remote Host=" & ctrl.RemoteHost & vbCrLf
            kb = kb & vbTab & "Remote Host IP=" & ctrl.RemoteHostIP & vbCrLf
            kb = kb & vbTab & "Remote Port=" & ctrl.RemotePort & vbCrLf
            kb = kb & vbTab & "State=" & aState(ctrl.State) & vbCrLf
If StreamCount(Idx) > 0 Then
            kb = kb & vbTab & "Client Servers=" & StreamCount(Idx) & vbCrLf

            For i = 1 To UBound(Sockets(Idx).Winsock.Streams)
                IdxStream = Sockets(Idx).Winsock.Streams(i)
                If IdxStream > 0 Then
                    kb = kb & vbTab & vbTab & Sockets(IdxStream).DevName & vbCrLf
                End If
            Next i

End If
            Count = Count + 1
        End If
    Next ctrl
    If Count = 0 Then
        kb = kb & "There are no TCP/IP Sockets in use"
    End If
    MsgBox kb, , "TCP/IP Sockets"
End Sub

'Creates winsock from info in Sockets
'if Sockets(Idx).Hidx = 0, then new Hidx is allocated
'State is set
'Returns Hidx if successful
Public Function CreateWinsock(Idx As Long) As Long
    
    WriteLog "Creating TCP/IP Socket for " & Sockets(Idx).DevName
    
    If Sockets(Idx).Hidx <= 0 Then
        Sockets(Idx).Hidx = FreeWinsock
        If Sockets(Idx).Hidx = -1 Then
            WriteLog "No free TCP/IP sockets, limit is " & MAX_WINSOCKS
            Exit Function
        End If
    End If

'Create the Winsock(index) control if required
    If WinsockExists(CInt(Sockets(Idx).Hidx)) = False Then
        WriteLog aHandler(Sockets(Idx).Handler) & " Handler " _
& Sockets(Idx).Hidx & " allocated to " & Sockets(Idx).DevName
            Load frmRouter.Winsock(Sockets(Idx).Hidx)
    Else
        WriteLog "Using " & aHandler(Sockets(Idx).Handler) _
        & " Handler " & Sockets(Idx).Hidx
'You must unload & reload winsock otherwise if the state is
'connection, it will never re-connect (or cause an invalid state error)
        Unload frmRouter.Winsock(Sockets(Idx).Hidx)
        Load frmRouter.Winsock(Sockets(Idx).Hidx)
    End If
    
    With frmRouter.Winsock(Sockets(Idx).Hidx)
'Close the port if open (will be if Sockets(Idx).Hidx <> 0)
        If .State <> sckClosed Then
            .Close
        End If
            
'Set up Winsock Handler to values in Sockets()
        Select Case Sockets(Idx).Winsock.Protocol
        Case Is = sckUDPProtocol
               .Protocol = sckUDPProtocol
            Select Case Sockets(Idx).Direction
'Input Only=1
            Case Is = 1
'Create a local port that is the same as the remote port that
'is sending the data
                On Error GoTo UDP_Bind_error
                .Bind Sockets(Idx).Winsock.LocalPort  'Input Bind to remote port.
                On Error GoTo 0
'If we are sending data and we do not have a local port
'create a random local port
'Output Only
            Case Is = 2
                On Error GoTo UDP_Bind_error
                .Bind 0         'Get a random free local port
                On Error GoTo 0
'If both Input and Output are specified we are setting up a
'peer-to-peer connection
'The lines below could set this up automatically when we receive data
'but could contradict setting up a route to another host
'Input & Output
            Case Is = 0
'see if we are receiving data from a remote port
'If so this will set up .RemoteHost and .RemotePort in the Winsock control
                DoEvents
            
            End Select

WriteLog "Opening TCP/IP socket " & .LocalPort
'If Output is to 127.0.0.1 (as opposed to the actual'
'interface IP eg 192.168.6.54 it causes (sometimes)
'looping when it tries to connect
'If it IS changed and the network interface is changed you have to reset the IP
'address to 127.0.0.1 (jeffrey.van.gils@rws.nl )
            If Sockets(Idx).Winsock.RemoteHost = "127.0.0.1" Then
'                Sockets(Idx).Winsock.RemoteHost = .LocalIP
            End If
'These only required if Sending data
            If Sockets(Idx).Direction = 2 Then
                .RemoteHost = Sockets(Idx).Winsock.RemoteHost   'Output Only
                .RemotePort = Sockets(Idx).Winsock.RemotePort   'Output Only
            End If
        Case Is = sckTCPProtocol
           .Protocol = sckTCPProtocol
            Select Case Sockets(Idx).Winsock.Server
            Case Is = 0  'Client
                .RemoteHost = Sockets(Idx).Winsock.RemoteHost   'Output Only
                .RemotePort = Sockets(Idx).Winsock.RemotePort   'Output Only
                On Error GoTo TCP_Connect_error
                .Connect .RemoteHost, .RemotePort
                On Error GoTo 0
            Case Is = 1  'Server
                .LocalPort = Sockets(Idx).Winsock.LocalPort
'Call DisplayWinsock
'Oidx must not be -1 to create a Client stream
                If IsTcpListener(Idx) Then
'                If Sockets(Idx).Winsock.Oidx = -1 Then
                    On Error GoTo TCP_Listen_error
                    .Listen
                    On Error GoTo 0
                Else
'Must be a stream
'Stop
'.Bind 0
                End If
            Case Else
                MsgBox "Invalid Direction"
            End Select
        End Select      'Protocol
        End With
'This is required to get the Idx when the data is received
'If 0 data is discarded
'Possible Subscipt error - because sindex = -1 if we cant allocate handler
    Winsocks(Sockets(Idx).Hidx).sIndex = Idx

'no longer reqd (use sockets)
'        Winsocks(Sockets(Idx).Hidx).Direction = Direction
'        Winsocks(Sockets(Idx).Hidx).Server = Server
'to here
    
'Can be called from LoadProfile or from Winsockcfg
    If FormExists("Winsockcfg") = True Then
        Unload Winsockcfg
    End If
    
'Copy Winsock state to Sockets()
    Sockets(Idx).State = frmRouter.Winsock(Sockets(Idx).Hidx).State
'causes significant delay    Sockets(Idx).Winsock.LocalIP = frmRouter.Winsock(Sockets(Idx).Hidx).LocalIP
    Exit Function

UDP_Bind_error:
WriteLog "Open failed with error " & err.Number & " " & err.Description
    Sockets(Idx).State = frmRouter.Winsock(Sockets(Idx).Hidx).State
    Sockets(Idx).errmsg = err.Description
    If err.Number = 10048 Then  'Address in use
        Sockets(Idx).State = 12
'Call DisplayWinsock
    End If
'Only display the message if not triggered by the timer
'Now it is displayed in the status column
'Otherwise it keeps coming up every 10 secs
'    If frmRouter.ReconnectTimer.Enabled = False Then
'        MsgBox "CreateWinsock Error " & Str(Err.Number) & " " & Err.Description & vbCrLf _
'        & "UDP Bind Error, Port " & frmRouter.Winsock(Sockets(Idx).Hidx).LocalIP & ":" & frmRouter.Winsock(Sockets(Idx).Hidx).LocalPort, , "Open Socket"
'    Else
'Check if the abose is now required
'Stop
'    End If
    Exit Function
TCP_Connect_error:        'Input Client
    Select Case err.Number
    Case Is = sckAddressInUse
'dont display as it will be shown in the display status
    Case Else
        WriteLog "Open failed with error " & err.Number & " " & err.Description
        Sockets(Idx).State = frmRouter.Winsock(Sockets(Idx).Hidx).State
        Sockets(CurrentSocket).errmsg = err.Description
        If frmRouter.ReconnectTimer.Enabled = False Then
            MsgBox "TCP Connect Error " & Str(err.Number) & " " & err.Description & vbCrLf _
            & "Can't open " _
            & "connection " & frmRouter.Winsock(Sockets(Idx).Hidx).RemoteHostIP _
            & ":" & frmRouter.Winsock(Sockets(Idx).Hidx).RemotePort & vbCrLf
       Else
'Check if above is required
'Stop
        End If
    End Select
    Exit Function
TCP_Listen_error:         'Output Server

WriteLog "Open failed with error " & err.Number & " " & err.Description
    Sockets(Idx).State = frmRouter.Winsock(Sockets(Idx).Hidx).State
    Sockets(Idx).errmsg = err.Description
    Select Case err.Number
    Case Is = sckAddressInUse
        Exit Function
    End Select
    MsgBox "TCP Connect Error " & Str(err.Number) & " " & err.Description & vbCrLf _
'    & "Can't open " & aProtocol(myPort.Protocol) _
'    & " connection " & myPort.Address _
'    & ":" & myPort.Port & vbCrLf
End Function


Public Function FreeWinsock() As Integer
Dim i As Integer

'If Winsock(Hidx)=Nothing then handler is free to use (but array does exist)
    For i = 1 To UBound(Winsocks)
        If WinsockExists(i) = False Then
            Exit For
            Exit Function
        End If
    Next i
    
    FreeWinsock = i
    If FreeWinsock > MAX_WINSOCKS Then
'no free Winsocks
        WriteLog "No free TCP/IP handlers, limit is " & MAX_WINSOCKS
        FreeWinsock = -1
    Else
        If FreeWinsock > UBound(Winsocks) Then
'We can still allocate more Winsocks
'On error fot this array is temporarily blocked
            On Error Resume Next
            ReDim Preserve Winsocks(1 To FreeWinsock)
            On Error GoTo 0
        End If
    End If
End Function

Public Function WinsockCount() As Long
Dim i As Integer
Dim Count As Long
    For i = 1 To UBound(Winsocks)
        If WinsockExists(i) = True Then Count = Count + 1
    Next i
    WinsockCount = i - 1    '(0) is not used
End Function
'Returns the CURRENT Winsock state for a socket
'or Closed if no winsock
Public Function WinsockState(Idx As Long) As Long
        On Error GoTo NoWinsock
        Select Case Sockets(Idx).Handler
        Case Is = 0     'winsock
            WinsockState = frmRouter.Winsock(Sockets(Idx).Hidx).State
        End Select
Exit Function

NoWinsock:
    WinsockState = sckClosed
End Function

