Attribute VB_Name = "modCreateClasses"
Option Explicit

'The Functions to Create a Class cannot reside in the Class itself
'because at this point the Class does not exist

Public Function CreateLoopBack(Idx As Long)
Dim Hidx As Long
Dim Ridx As Long

WriteLog "Creating LoopBack Socket " & Sockets(Idx).DevName
    If Sockets(Idx).Hidx <= 0 Then
        Sockets(Idx).Hidx = FreeLoopBack
        If Sockets(Idx).Hidx = -1 Then
            MsgBox "Cant create LoopBack handler"
            Exit Function
        End If
    End If
'For the LoopBack were not using a set-up form so we initialise the form
'This is done in cmd_OK if using setup form
'Add the Object(s) to the array LoopBacks() SetLoopBackhas re-dimensioned
'the array (if it was too small)
        WriteLog aHandler(Sockets(Idx).Handler) & " Handler " _
& Sockets(Idx).Hidx & " allocated to " & Sockets(Idx).DevName
'This is here to make the App Log consistent with other handlers
        WriteLog "Opening " & Sockets(Idx).DevName
        Set LoopBacks(Sockets(Idx).Hidx) = New clsLoopBack
        LoopBacks(Sockets(Idx).Hidx).sIndex = Idx
        LoopBacks(Sockets(Idx).Hidx).hIndex = Sockets(Idx).Hidx
        LoopBacks(Sockets(Idx).Hidx).Name = Sockets(Idx).DevName
'        Ridx = Routecfg.CreateRoute(Idx, Idx)
'        If Ridx < 1 Then
'            Exit Function
'        End If
'Create Route does not Add the forwards because we need to
'establish first if the route is enabled
'        Sockets(Idx).Routes(Ridx).Enabled = True
'Now we can create any forwarding required
        LoopBacks(Sockets(Idx).Hidx).State = 1
End Function

Public Function FreeLoopBack() As Long
Dim i As Long

    For i = 1 To UBound(LoopBacks)
        If Not LoopBacks(i) Is Nothing Then
            If LoopBacks(i).State = -1 Then
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
    
    FreeLoopBack = i
    If FreeLoopBack > MAX_LOOPBACKS Then
'no free LoopBacks
        WriteLog "No free loopback handlers, limit is " & MAX_LOOPBACKS
        FreeLoopBack = -1
    Else
        If FreeLoopBack > UBound(LoopBacks) Then
'We can still allocate more LoopBacks
            ReDim Preserve LoopBacks(1 To UBound(LoopBacks) + 1)
        End If
    End If
End Function

'If Hidx is 0 then a new handler is created
'Else set the existing one
'Puts the Hidx into Sockets(Idx).Hidx if successful
'Else -1
Public Sub CreateFile(Idx As Long)
Dim ret As Long
    
WriteLog "Creating File Socket " & Sockets(Idx).DevName
    On Error GoTo CreateFile_error
'On Error GoTo 0  'debug
    If Sockets(Idx).Hidx <= 0 Then
'When first opened HIDX = 0
        Sockets(Idx).Hidx = FreeFileHidx
        If Sockets(Idx).Hidx = -1 Then
            MsgBox "Cant create File handler"
            Exit Sub
        End If
    End If
    
'Create the FileInputTimer(index) control if required
    If FileTimerExists(CInt(Sockets(Idx).Hidx)) = False Then
        WriteLog "File Timer " _
& Sockets(Idx).Hidx & " allocated to " & Sockets(Idx).DevName
        Load frmRouter.FileInputTimer(Sockets(Idx).Hidx)
    Else
        WriteLog "Using File Timer " & Sockets(Idx).Hidx
    End If
    
    With Sockets(Idx)
'If .Hidx = 0 Then Stop
'Create the Files(index) control if required
        If UBound(Files) < .Hidx Then ReDim Preserve Files(.Hidx)
    
'Are we opening an existing File
        
        If Not Files(.Hidx) Is Nothing Then
WriteLog "Using " & aHandler(Sockets(Idx).Handler) & " Handler " _
& Sockets(Idx).Hidx
            ret = Files(.Hidx).OpenFile()
        Else
WriteLog aHandler(Sockets(Idx).Handler) & " Handler " _
& Sockets(Idx).Hidx & " allocated to " & Sockets(Idx).File.SocketFileName
'If .Hidx = 0 Then Stop
            Set Files(.Hidx) = New clsFile
            Files(.Hidx).hIndex = .Hidx
            Files(.Hidx).sIndex = Idx
            Files(.Hidx).Name = .DevName    'USer defined
'            ret = Files(.Hidx).OpenFile(.File.SocketFileName _
'            , .Direction, frmRouter)
            ret = Files(.Hidx).OpenFile()
        End If
        If ret <> 0 Then
MsgBox "Can't open " & .File.SocketFileName, , "Create File Error"
            Sockets(Idx).State = 9
        Else

        End If
    End With
    If FormExists("Filecfg") = True Then    'False on load from registry
        Unload Filecfg
    End If
    Exit Sub

CreateFile_error:
    Sockets(Idx).State = 9  'Error
'Stop
'This will clear the error     On Error GoTo 0
    Select Case err.Number
    Case Else
        MsgBox err.Number & " " & err.Description, , "Create File Error"
    End Select
'Dont unload the form so that user has to cancel or enter a valid port
End Sub

Public Function FreeFileHidx() As Long
Dim i As Long

'If Files() is nothing the use this Hidx
'Otherwise if State is Closed of less we can also use this index
    For i = 1 To UBound(Files)
        If Not Files(i) Is Nothing Then
'            If Files(i).State = -1 Then
            If Files(i).State <= 0 Then
                FreeFileHidx = i
                Exit Function
            End If
        Else            'then Files(i)=nothing
            FreeFileHidx = i
            Exit Function
        End If
    Next i
    If UBound(Files) = MAX_FILES Then
'no free Files
'no free Winsocks
        WriteLog "No free File handlers, limit is " & MAX_FILES
        FreeFileHidx = -1
    Else
'We can still allocate more sockets
        ReDim Preserve Files(1 To UBound(Files) + 1)
'Cant set File state here as the Object will be nothing
'            Files(UBound(Files)).State = -1   'Not allocated
        FreeFileHidx = UBound(Files)
    End If
End Function

Public Function FileInputTimerExists(Hidx As Integer) As Boolean
    On Error GoTo NoTimer
    If frmRouter.FileInputTimer(Hidx).Index Then
        FileInputTimerExists = True
    End If
NoTimer:
End Function

'Check to see if Files(index) exists, there seems to be no other way to check
'other than trying to access the index
Public Function FileTimerExists(Hidx As Long) As Boolean
'see if next line always works
If Not Files(Hidx) Is Nothing Then
    On Error GoTo NoFile
    If Files(Hidx).hIndex > 0 Then
        FileTimerExists = True
    End If
End If
NoFile:
End Function

Public Function IsFileHandlerOpenForInput() As Boolean
Dim Idx As Long
Dim Hidx As Variant

    For Hidx = 1 To UBound(Files)
        If Not Files(Hidx) Is Nothing Then
            Idx = Files(Hidx).sIndex
            If Sockets(Idx).State <> -1 Then
                Select Case Sockets(Idx).Direction
                Case Is = 1     'Input
                    IsFileHandlerOpenForInput = True
                End Select
            End If
        End If
    Next Hidx
End Function

'Returns true is there is any file input to continue with
'The file must be open
Public Function ContinueAllFileInput() As Boolean
Dim Idx As Long
Dim Hidx As Variant

    For Hidx = 1 To UBound(Files)
        If Not Files(Hidx) Is Nothing Then
            Idx = Files(Hidx).sIndex
            If Sockets(Idx).State = 1 Then
                Select Case Sockets(Idx).Direction
                Case Is = 1     'Input
'Pause may be pressed before a second input file has started
                    If frmRouter.cmdPause.Caption = "Pause" Then
                        frmRouter.FileInputTimer(Hidx).Enabled = True
'Call Files(Hidx).ContinueFileInput
                        ContinueAllFileInput = True
                    End If
                End Select
            End If
        End If
    Next Hidx
End Function

Public Function RestartFileInput(Hidx As Integer) As Boolean
Dim Idx As Long

    If Hidx > 0 Then    'Handler must have been created
        If Not Files(Hidx) Is Nothing Then
            Idx = Files(Hidx).sIndex
            If Sockets(Idx).State = 1 Then
                Select Case Sockets(Idx).Direction
                Case Is = 1     'Input
                    Call Files(Hidx).ContinueFileInput
                    RestartFileInput = True
                End Select
            End If
        End If
    End If
End Function

'Returns true is there is any file input to pause
'The file must be open
Public Function PauseAllFileInput() As Boolean
Dim Idx As Long
Dim Hidx As Variant

    For Hidx = 1 To UBound(Files)
        If Not Files(Hidx) Is Nothing Then
            Idx = Files(Hidx).sIndex
            If Sockets(Idx).State = 1 Then
                Select Case Sockets(Idx).Direction
                Case Is = 1     'Input
                    Call Files(Hidx).PauseFileInput
                    PauseAllFileInput = True
                End Select
            End If
        End If
    Next Hidx
End Function


