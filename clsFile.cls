VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

'Public FullName As String
'Public Hidx As Long    'hindex
'Public Idx As Long     'sindex

Private Type ADDINFO    'additional info
'    Index As Long
'    hName As String
'    ErrMsg As String
    State As Long      '0 = closed, 1=Opem (Only polled when state is 1)
    Hidx As Long    'Comm Array Index(So we know the Socket when data rcvd)
    Idx As Long      'Sockets index (So we know the Socket when data rcvd)
    ch As Integer
    hName As String 'User defined handler name eg File1
    Direction As Long   'Need to know when we open the file
    RollOverDate As String
    FileName As String  'Actual file name wre using (including rollover)
    Multiplier As Long  'No of times Timer needs to fire
    TimerCount  As Long     'No of times timer has fired
End Type

' Private members
Private Info As ADDINFO

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Public Property Get RollOverDate() As String
    RollOverDate = Info.RollOverDate
End Property

Public Property Let RollOverDate(vNewValue As String)
    Info.RollOverDate = vNewValue
End Property

Public Property Get FileName() As String
    FileName = Info.FileName
End Property

Public Property Let FileName(vNewValue As String)
    Info.FileName = vNewValue
End Property

Public Property Get State() As Long
    State = Info.State
End Property

Public Property Let State(vNewValue As Long)
    Info.State = vNewValue
End Property

Public Property Get hIndex() As Long
    hIndex = Info.Hidx
End Property

Public Property Let hIndex(vNewValue As Long)
    Info.Hidx = vNewValue
End Property

Public Property Get sIndex() As Long
    sIndex = Info.Idx
End Property

Public Property Let sIndex(vNewValue As Long)
    Info.Idx = vNewValue
End Property

Public Property Get Name() As String
    Name = Info.hName
End Property

Public Property Let Name(vNewValue As String)
    Info.hName = vNewValue
End Property

Public Property Get ch() As Integer
    ch = Info.ch
End Property

Public Property Let ch(vNewValue As Integer)
    Info.ch = vNewValue
End Property

Public Property Let sDirection(vNewValue As Integer)
    Info.Direction = vNewValue
End Property

Public Property Get sDirection() As Integer
    sDirection = Info.Direction
End Property

Public Property Let Multiplier(vNewValue As Long)
    Info.Multiplier = vNewValue
End Property

Public Property Get Multiplier() As Long
    Multiplier = Info.Multiplier
End Property

Public Property Let TimerCount(vNewValue As Long)
    Info.TimerCount = vNewValue
End Property

Public Property Get TimerCount() As Long
    TimerCount = Info.TimerCount
End Property

' The main function for opening a File
Public Function OpenFile() As Long
Dim kb As String

'A File can be Paused. If so the channel must not be re-opend as we
'must retain context to continue reading records where we left off

    FileName = Sockets(sIndex).File.SocketFileName
    WriteLog "Opening " & NameFromFullPath(FileName)
    On Error GoTo OpenFile_error
    If State = 0 Then   'Closed
        State = 11
    End If
    
'Not already open, get new channel
'If Paused, channel will not be closed
    If ch = 0 Then
        ch = FreeFile
    Else
        WriteLog NameFromFullPath(FileName) & " already open"
    End If

    With frmRouter.FileInputTimer(CInt(hIndex))
        Select Case Sockets(sIndex).File.ReadRate
            Case Is = 0     'unlimited
                .Interval = 1000    'Fire the timer every second
'Causes an exit from the ContinueFileInput after 10,000 records
'have been read. This means if there is more than one file open
'for input, the next file will start getting read as well
                Multiplier = 10000   'Read
            Case Is = 1
                .Interval = 60000   '1 msgs/min
                Multiplier = 1
            Case Is = 2
                .Interval = 6000    '10 msgs/min
                Multiplier = 1
            Case Is = 3
                .Interval = 1200    '50 msgs/min
                Multiplier = 1
            Case Is = 4
                .Interval = 600      '100
                Multiplier = 1
            Case Is = 5
                .Interval = 600      '500
                Multiplier = 5
            Case Is = 6
                .Interval = 600      '1000
                Multiplier = 10
            Case Is = 7
                .Interval = 600      '5000
                Multiplier = 50
            End Select
        End With
    
    Select Case Sockets(sIndex).Direction
    Case Is = 1
'The file can already be open, if editing while file input is paused.
'If the File name has been changed in Filecfg, the file should
'have been closed
        If State <> 1 Then
            Open FileName For Input As #ch
        End If
        State = 1
    Case Is = 2
'MsgBox FileName & " for " & aDirection(Sockets(sIndex).Direction), , "clsFile.OpenFile"
        If Sockets(sIndex).File.RollOver = True Then
'Put the date were using into the handler
            Call NewFileName
        Else
            RollOverDate = ""
        End If
        Open FileName For Append As #ch
        State = 1
    Case Else
        State = 9
        err.Raise -1, "OpenFile", "No File Direction"
    End Select
'        sDirection = Direction
    
'Triggerd by Start button as stop may not yet be enabled in cmdStart
    If IsFileHandlerOpenForInput And Not frmRouter.cmdStart.Enabled Then
        frmRouter.cmdPause.Visible = True
        frmRouter.cmdPause.Enabled = True
'You cannot start file input when the file is opened as the output
'handler may no yet have been opened.
'Stop
    Else
        frmRouter.cmdPause.Visible = False
        frmRouter.cmdPause.Enabled = False
'Set caption for next time any file is enabled for input
        frmRouter.cmdPause.Caption = "Pause"
    End If
    Exit Function

OpenFile_error:
'MsgBox Err.Number & " " & Err.Description
'The error number is passed back to Commcfg & reported by it
    OpenFile = err.Number   'return error
    Sockets(sIndex).State = 9
    Sockets(sIndex).errmsg = err.Description
    Set Files(Sockets(sIndex).Hidx) = Nothing   'Remove the Comm Socket if closed
    Sockets(sIndex).Hidx = -1
WriteLog "Error " & err.Number & " opening file - " & err.Description
End Function    'will clear err

Public Function CloseFile()
Dim kb As String
'close the Channel
'PauseOneFile
'Stop
    If FileInputTimerExists(CInt(hIndex)) Then
        If Sockets(sIndex).Direction = 1 Then
kb = frmRouter.FileInputTimer(hIndex).Enabled
'Possibility the Files does not exist but the timer does
'Can happen if the File is closed with an error
            If Not Files(hIndex) Is Nothing Then
                Call Files(hIndex).PauseFileInput
            End If
'            frmRouter.FileInputTimer(CInt(hIndex)).Enabled = False
        End If
        Unload frmRouter.FileInputTimer(hIndex)
    End If
    
    Set Files(CInt(hIndex)) = Nothing

'If called by terminate the frmRouter will not be visible
    If frmRouter.Visible = True Then
        If IsFileHandlerOpenForInput = False Then
            frmRouter.cmdPause.Visible = False
            frmRouter.cmdPause.Enabled = False
        End If
    Else
'Stop    'Want to see if this happens (dont think it does)
    End If
    
    State = 0
    Close ch
    ch = 0
End Function

Public Function FileOutput(Data As String)
    If ch > 0 Then 'rollover (if date changed ot null)
        If RollOverDate <> "" And SysDate <> RollOverDate Then
            Close ch
            ch = 0
        End If
    End If
    
    If ch = 0 Then
        Call OpenFile
    End If
    
    If ch Then
'bad file name
        On Error Resume Next
        Print #ch, Data
    End If
End Function

'Continues (or Starts) the file input for Files(Hidx)
'Returns False at EOF
Public Function ContinueFileInput() As Boolean
Dim kb As String
Dim Count As Long

If jnasetup Then
    WriteLog "Repeat-Continue " & Sockets(sIndex).DevName
End If

'Possibility of socket not created but file trying to input
'when profile is changed
    
'The channel should be open
    If ch = 0 Then
        WriteLog "Channel not open in ContinueFileInput"
        Exit Function
    End If
 
    
'Call frmRouter.DisplayForwarding
'If the socket is not yet open keep trying every one sec

'Start the FileInputTimer and continue looping until
'the timer timeout

With frmRouter.FileInputTimer(CInt(hIndex))
    .Enabled = True
    Do Until EOF(ch)
        Line Input #ch, kb
        Count = Count + 1
        Call ForwardData(kb & vbCrLf, sIndex)
'You must DoEvents otherwise the Comm Poll Event is not triggered
        If Count Mod 50 = 0 Then
            DoEvents
        End If
'One event could be to close the handler
'If so files(hindex) will have been removed
        If Not Files(hIndex) Is Nothing Then
        Else
'No more input (File is Closed)
'Timer should have also been removed
            Exit Function
            End If
'File Input Could also have been paused
        If .Enabled = False Then
'Timer will have been disabled
            Exit Function   'Paused
        End If
    
'Exit and wait until Timer call ContinueFileInput again
'Will be > if input rate has been changed to a lower speed
        If Count >= Files(hIndex).Multiplier Then
            Exit Function
        End If
    Loop
'EOF reached
    Call Routecfg.RemoveSocketForwards(Files(hIndex).sIndex)
    Call CloseHandler(Files(hIndex).sIndex)
'This will also stop the timer
End With
End Function

'Stops the file input for Files(Hidx)
'All you need to do is to stop the timer
Public Function PauseFileInput()

If jnasetup Then
    WriteLog "Pause " & Sockets(sIndex).DevName
End If

    If Not Files(hIndex) Is Nothing Then
        frmRouter.FileInputTimer(CInt(hIndex)).Enabled = False
    End If
End Function

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsFile; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
    If ch <> 0 Then
         Call CloseFile
    End If
    WriteLog Info.hName & " [" & NameFromFullPath(FileName) & "] closed"
End Sub

Public Function NewFileName()
'Input: File name including path and optionally Rollover
'Output: Full File name with Rollover as current date
Dim FullName As String
Dim NameAndExt As String
    
    FullName = Sockets(sIndex).File.SocketFileName
    NameAndExt = NameFromFullPath(FullName, , True)
'True=remove any rollover date
    FileName = ExtendFullName(FullName, "_" & SysDate)
    RollOverDate = SysDate
End Function


