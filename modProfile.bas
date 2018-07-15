Attribute VB_Name = "modProfile"
Option Explicit

Public LoadingProfile As Boolean

'Cancel indicates an error in loading
'Loads a Profile from Registry into memory
'Needs to call LoadSocket to set up individual sockets
'as this is required if a disabled socket is re-enabled
Public Sub LoadProfile(ProfileName As String, Cancel As Boolean)
Dim Key As String
Dim SubKeys() As Variant  'Must be variant
Dim SubKeyCount As Long
'Dim NameValueCount As Long
'Dim Names() As Variant  'Must be variant
'Dim Values() As Variant  'Must be variant
Dim i As Long
Dim k As Long
Dim j As Long
Dim Idx As Long
Dim SocketKey As String
Dim msg As String
Dim Hidx As Long
Dim RouteKey As String
Dim f As Form
Dim DevName1 As String
Dim DevName2 As String
Dim FilterKey As String
Dim GraphKey As String
Dim GraphName As String
Dim Idx1 As Long
Dim Idx2 As Long
Dim Ridx As Long
Dim kb As String

Dim OutputKey As String
Dim SubOutputKeyCount As Long
Dim SubOutputKeys() As Variant
Dim SubOutputKeyNo As Long
Dim OutputFormatKey As String

WriteLog "Loading Profile " & ProfileName
        CurrentProfile = ProfileName
        Cancel = False
'Check if Profile exists in registry, If so read new profile
        Key = ROUTERKEY & "\Profiles\" & ProfileName & "\Sockets"

'Set up the Socket and Handler Details
        SubKeyCount = ReadKeys(HKEY_CURRENT_USER, Key, SubKeys)
        If SubKeyCount > 0 Then
            For k = 0 To SubKeyCount - 1
WriteLog "Setting up Connection " & SubKeys(k)
                Idx = FreeSocket
                If Idx < 1 Then
                    msg = "Cannot load profile " & ProfileName & vbCrLf & "(no free connections)"
                    GoTo Load_error
                End If
                WriteLog "Socket " & FreeSocket & " allocated"
                SocketKey = Key & "\" & SubKeys(k)
                Sockets(Idx).DevName = SubKeys(k)
                Sockets(Idx).Handler = _
                QueryValue(HKEY_CURRENT_USER, SocketKey, "Handler")
                Sockets(Idx).Direction = _
                CLng(QueryValue(HKEY_CURRENT_USER, SocketKey, "Direction"))
                Sockets(Idx).Enabled = BooleanA(QueryValue(HKEY_CURRENT_USER, SocketKey, "Enabled"))
                Sockets(Idx).Graph = BooleanA(QueryValue(HKEY_CURRENT_USER, SocketKey, "Graph"))
                
                Select Case Sockets(Idx).Handler
                Case Is = 0 'UDP/TCP
'Set up handler (must get a new handler index)
'Load Registry settings into Sockets().Winsock
        Sockets(Idx).Winsock.Protocol = CLng(QueryValue(HKEY_CURRENT_USER, SocketKey, "Protocol"))
        Sockets(Idx).Winsock.Server = CLng(QueryValue(HKEY_CURRENT_USER, SocketKey, "Server"))
'v45 Digi definition UDP Client=0 (Outgoing Device=2)
'UDP Server=1 (Incoming Device=1, default)
If Sockets(Idx).Winsock.Protocol = sckUDPProtocol Then
    If Sockets(Idx).Direction = 1 Then '0=both, 1=Input, 2=Output
        Sockets(Idx).Winsock.Server = 1
    Else
        Sockets(Idx).Winsock.Server = 0
    End If
End If
        Sockets(Idx).Winsock.LocalPort = QueryValue(HKEY_CURRENT_USER, SocketKey, "LocalPort")
        Sockets(Idx).Winsock.RemoteHost = QueryValue(HKEY_CURRENT_USER, SocketKey, "RemoteHost")
        Sockets(Idx).Winsock.RemotePort = QueryValue(HKEY_CURRENT_USER, SocketKey, "RemotePort")
        Sockets(Idx).Winsock.PermittedStreams = QueryValue(HKEY_CURRENT_USER, SocketKey, "PermittedStreams")
        Sockets(Idx).Winsock.PermittedIPStreams = QueryValue(HKEY_CURRENT_USER, SocketKey, "PermittedIPStreams")
'Will be blank=0 if loading existing profile
        If Sockets(Idx).Winsock.PermittedIPStreams < 1 Then
            Sockets(Idx).Winsock.PermittedIPStreams = 1
        End If
                
                Case Is = 1     'COMM
'Set up handler
    Sockets(Idx).Comm.Name = QueryValue(HKEY_CURRENT_USER, SocketKey, "CommName")
    Sockets(Idx).Comm.BaudRate = QueryValue(HKEY_CURRENT_USER, SocketKey, "BaudRate")
    Sockets(Idx).Comm.VCP = GetVCP(Sockets(Idx).Comm.Name)

                Case Is = 2     'File
'MsgBox "File handler not yet available", , "Load Profile"
    Sockets(Idx).File.SocketFileName = QueryValue(HKEY_CURRENT_USER, SocketKey, "SocketFileName")
    Sockets(Idx).File.ReadRate = QueryValue(HKEY_CURRENT_USER, SocketKey, "ReadRate")
    Sockets(Idx).File.RollOver = BooleanA(QueryValue(HKEY_CURRENT_USER, SocketKey, "RollOver"))
                Case Is = 3     'TTY
                Case Is = 4     'LoopBacks
                Case Else
MsgBox "Handler " & Sockets(Idx).Handler & " not available", , "Load Profile"
'release the socket
                End Select
                
'If we do not have all the details disable the socket
               If Socket_Validate(Idx) = True Then
'Must set to other than -1 otherwise as new IDX will not be allocated
'for next socket
                    Sockets(Idx).State = 3     'connection pending
                Else
                    Sockets(Idx).Enabled = False
                End If
                
            OutputKey = SocketKey & "\OutputFormat\IEC"
Sockets(Idx).IEC.Enabled = BooleanA(QueryValue(HKEY_CURRENT_USER, OutputKey, "Enabled"))
            OutputKey = SocketKey & "\OutputFormat\PlainNmea"
Sockets(Idx).OutputFormat.PlainNmea = BooleanA(QueryValue(HKEY_CURRENT_USER, OutputKey, "Enabled"))
            OutputKey = SocketKey & "\OutputFormat\OwnShip"
Sockets(Idx).OutputFormat.OwnShipMmsi = QueryValue(HKEY_CURRENT_USER, OutputKey, "MMSI")
'This is how to get the IEC Options detail (Time,Source etc)
'- not yet implemented
            SubOutputKeyCount = ReadKeys(HKEY_CURRENT_USER, OutputKey, SubOutputKeys)
            If SubOutputKeyCount > 0 Then
                For SubOutputKeyNo = 0 To SubOutputKeyCount - 1
                Next SubOutputKeyNo
            End If
'End if the IEC detail

            Next k  'Next Socket to set up
        End If  'End of Sockets
        
'Routes
'Need to use the same procedure for setting up Routes as fro Sockets
'We should look up the Idx allocated above and
'then validate the route (see above)
'then Open the Route

        Key = ROUTERKEY & "\Profiles\" & ProfileName & "\Routes"
'        NameValueCount = ReadNameValues(HKEY_CURRENT_USER, Key, Names, Values)
'        If NameValueCount > 0 Then
'        End If
'Set up the Route Details
        SubKeyCount = ReadKeys(HKEY_CURRENT_USER, Key, SubKeys)
        If SubKeyCount > 0 Then
            For k = 0 To SubKeyCount - 1
                RouteKey = Key & "\" & SubKeys(k)
DevName1 = QueryValue(HKEY_CURRENT_USER, RouteKey, "DevName1")
DevName2 = QueryValue(HKEY_CURRENT_USER, RouteKey, "DevName2")
                Idx1 = DevNameToSocket(DevName1)
                Idx2 = DevNameToSocket(DevName2)
                If Routecfg.ValidateNewRoute(Idx1, Idx2) = True Then
'Idx1 is returned as the lower
'Hidx applies to Idx1 NOT Idx2
                    Ridx = Routecfg.CreateRoute(Idx1, Idx2)
                    If Ridx > 0 Then
Sockets(Idx1).Routes(Ridx).Enabled = BooleanA(QueryValue(HKEY_CURRENT_USER, RouteKey, "Enabled"))
                    End If

'Forwards are created when cmdStart is called
                Else
MsgBox "Cannot create " & SubKeys(k) & vbCrLf _
& "beteen " & DevName1 & " and " & DevName2
                End If
'Get any OutputFormat

            Next k
        End If
'End of Routes

'Filters
'Always Create the class to hold the registry values
'Destroyed when the profile is closed
        Set SourceDuplicateFilter = New clsDuplicateFilter
       
       Key = ROUTERKEY & "\Profiles\" & ProfileName & "\Filters"
        DevName1 = QueryValue(HKEY_CURRENT_USER, Key, "DmzDeviceName")
        SourceDuplicateFilter.DmzIdx = DevNameToSocket(DevName1)
        SubKeyCount = ReadKeys(HKEY_CURRENT_USER, Key, SubKeys)
        If SubKeyCount > 0 Then
            For k = 0 To SubKeyCount - 1
                FilterKey = Key & "\" & SubKeys(k)
                SourceDuplicateFilter.Enabled = BooleanA(QueryValue(HKEY_CURRENT_USER, FilterKey, "RemoveSourceDuplicates"))
                SourceDuplicateFilter.OnlyVdm = BooleanA(QueryValue(HKEY_CURRENT_USER, FilterKey, "RemoveSourceNotAivdm"))
                SourceDuplicateFilter.RejectMmsi = QueryValue(HKEY_CURRENT_USER, FilterKey, "RemoveSourceNotMmsi")
                SourceDuplicateFilter.RejectPayloadErrors = QueryValue(HKEY_CURRENT_USER, FilterKey, "RejectPayloadErrors")
            Next k
        End If
'Graphs
        Key = ROUTERKEY & "\Profiles\" & ProfileName & "\Graphs"
        SubKeyCount = ReadKeys(HKEY_CURRENT_USER, Key, SubKeys)
        If SubKeyCount > 0 Then
            For k = 0 To SubKeyCount - 1
                GraphKey = Key & "\" & SubKeys(k)
ExcelUpdateInterval = QueryValue(HKEY_CURRENT_USER, GraphKey, "ExcelUpdateInterval")
ExcelRange = QueryValue(HKEY_CURRENT_USER, GraphKey, "ExcelRange")
ExcelUTC = BooleanA(QueryValue(HKEY_CURRENT_USER, GraphKey, "ExcelUTC"))
            Next k
        Else
'This is the default when an existing profile is loaded
'since Graphs were coded
            ExcelUpdateInterval = 1
            ExcelRange = 100
            ExcelUTC = True
        End If
        
'Now set the Current Profile
'can be 0 when a profile is deletedIf Len(Trim$(CurrentProfile)) = 0 Then Stop 'v59test
    Call SetKeyValue(HKEY_CURRENT_USER, ROUTERKEY, "CurrentProfile", CurrentProfile, REG_SZ)
'End of Open
'create a new reg file containing the current profile. It will be checked on exit
'to see if there are any changes
    Call SaveProfile(CurrentProfile, CurrentProfile & ".txt")
    
'Update the display & set the window size of frmRouter to Controls
'    Call frmRouter.UpdateMshRows
    
'Start processing input data
    frmRouter.cmdStart_Click

'This MUST be done after the FlexGrids have been added into frmRouter
'Otherwise the resizing will be incorrect
'AND it must be after the TTY's have been made visible
    
'RestoreWindow sets windows sizes according to the value
'currently in the registry
'Call DisplayWindow("modProfile.LoadProfile - before restore")
    kb = ""
    For Each f In Forms
'start minimised in form properties
        kb = kb & RestoreWindow(f)

'do not save frmRouterMinimized
'        If f.Name = "frmRouter" Then
'            frmRouter.WindowState = vbNormal
'        End If
    Next f
'Call DisplayWindow("modProfile.LoadProfile - after restore")
    
'Moved from frmRouter.Load
    If frmRouter.Visible = False Then   'v59a Cant show if Modal form visible (Open Profile)
        frmRouter.Show  'Causes the first resize
    End If
'Probably should be in load profile
'Show any TTY forms that have been created when the profile was loaded
    Call MakeFormsVisisble
    frmRouter.SetCaption
    frmRouter.MenuRouteNew.Enabled = IsNewRoute
    
    Call frmRouter.UpdateMshRows

'Status bar timer is only used when jnasetup is true
'To display the status uf all timers
If jnasetup Then
    frmRouter.StatusBarTimer = True
End If


'End of moved from frmRouter.Load

WriteLog "Profile Loaded [" & ProfileName & "]"
    
    
'Set Sockets(2).Recorder = New clsRecorder
'Sockets(2).Recorder.ParentIdx = 2
'Sockets(2).Recorder.Show
'Set Sockets(6).Recorder = New frmRecorder
    Exit Sub

Load_error:
    On Error GoTo 0
    If err.Number Then
        msg = "Error number " & err.Number & ", " & err.Description & msg & vbCrLf
    End If
    MsgBox msg, , "Load Profile"
    Cancel = True
End Sub

'Saves changed profile
'Close all sockets
Public Sub CloseProfile(ProfileName As String)
Dim Idx As Long
Dim Hidx As Long
Dim f As Form
Dim OldFileName As String
Dim NewFileName As String
Dim ret As Integer
Dim i As Long
Dim ctrl As Control
Dim Pctrl As Control
Dim kb As String
Dim SaveGraph As Integer
Dim Col As Long

WriteLog "Closing Profile " & ProfileName
'Stop the all event timers
'kb = ""
    For Each ctrl In frmRouter
        If TypeOf ctrl Is Timer Then
'kb = kb & ctrl.Name & vbCrLf
            ctrl.Enabled = False
        End If
    Next ctrl

'MsgBox kb
'Stop
    ret = 0

    If ProfileName <> "" Then
        'Save Profile to a new file
        Call SaveProfile(ProfileName, ProfileName & ".new")
        OldFileName = TempPath & "Router" & ProfileName & ".txt"
'The Old file will not exist if a New (or SaveAs) file
        If FileExists(OldFileName) = True Then
            NewFileName = TempPath & "Router" & ProfileName & ".new"
            If FileCompare(OldFileName, NewFileName) = False Then
                ret = MsgBox("Profile " & ProfileName & " has changed " _
                & vbCrLf & "Save Changed Profile ? ", vbYesNo, "Close Profile")
            Else
                ret = vbNo
            End If
        Select Case ret
        Case Is = vbYes
'Save to registry
            Call SaveProfile(ProfileName)
        Case Is = vbNo
        End Select
        End If
    
'Always save the window positions when a profile is closed
    Call SaveWindows
    End If
    
'Clear the display MSH grids first
    Call frmRouter.ResetDisplay 'v59a
#If False Then  'v59a
    With frmRouter.mshSockets
        Do Until .Rows = .FixedRows + 1
            .RemoveItem (.Rows - 1)
        Loop
        
        .Row = .FixedRows
        .Col = .FixedCols
        Do Until .Col = .Cols - 1
        .TextMatrix(.Row, .Col) = .Col
        .CellBackColor = vbWhite
        If .Col = .Cols - 1 Then Exit Do
            .Col = .Col + 1
        Loop
'.TextMatrix(1, 11) = "11"
    End With
    
    With frmRouter.mshRoutes
        Do Until .Rows = .FixedRows + 1
            .RemoveItem (.Rows - 1)
        Loop
        .Row = .FixedRows
        .Col = .FixedCols
        Do Until .Col = .Cols - 1
        .TextMatrix(.Row, .Col) = ""
        .CellBackColor = vbWhite
        If .Col = .Cols - 1 Then Exit Do
            .Col = .Col + 1
        Loop

    End With
#End If

'Close the Handlers
    If SocketCount > 0 Then
        For Idx = 1 To UBound(Sockets)
            Call CloseHandler(Idx)  'added v26, removed code below
            Call RemoveRecorder(Idx)
            Call ClearSocket(Idx)
        Next Idx
    End If
'Clear Clients collection
    Call frmRouter.OutputRejectionLog
    
    On Error Resume Next
    ReDim Sockets(1 To 1)
    Call ClearSocket(1)

    ReDim TTYs(1 To 1)  'must initialise first element
    ReDim Comms(1 To 1)
    ReDim Winsocks(1 To 1)
    
'Remove Winsocks
    For Each ctrl In frmRouter.Winsock
kb = kb & "Index=" & ctrl.Index & vbCrLf
        If ctrl.Index > 0 Then
        Unload frmRouter.Winsock(ctrl.Index)
        End If
    Next ctrl

'Unload all forms EXCEPT frmRouter & Profilecfg as we may be opening a new profile (Profilecfg will
'be open and holds the new profile name)
'If closing the program then CloseProfile will have been called by frmRouter.Form_QueryUnload
'Profilecfg cannot be open as it is a MDI form.
    For Each f In Forms
        Select Case f.Name
        Case Is = "frmRouter", "Profilecfg"
        Case Else
            Unload f
        End Select
    Next f

'Ask before we clear frmRouter (looks naff otherwise)
    If ExcelOpen = True And WorkbookExists = True Then
        frmRouter.Visible = False
        SaveGraph = MsgBox("Would you like to save your Graph ? ", vbYesNo)
        Call CloseExcel(SaveGraph)    'vbNo=7
    End If

'Destroy the existing filter - create when profile loaded
    Set SourceDuplicateFilter = Nothing
    
    If InputLogCh > 0 Then
        Close #InputLogCh
        InputLogCh = 0
    End If
WriteLog "Profile " & ProfileName & " closed"
    If StartupLogFileCh <> 0 Then
        Close #StartupLogFileCh
    End If
    Call DisplayDebugSerial
End Sub

Public Sub DeleteProfile(ProfileName As String)
Dim Key As String
Dim SubKeys()
Dim SubKeyCount As Long
Dim i As Long

    Key = ROUTERKEY & "\Profiles\" & ProfileName
    Call DeleteKeys(Key)

End Sub

'If ProfileName is blank All Profiles are Saved to the file
'otherwise only the Named Profile is saved (enabling changes to be checked)
'If a Reg File Name is passed, the ProfileName is saved
'to a file, Otherwise the Profile is Saved to the Registry
Public Sub SaveProfile(ProfileName As String, Optional ByVal RegFileName As String)
Dim Idx As Long
Dim Ridx As Long
Dim ProfileKey As String
Dim SocketKey As String
Dim HandlerKey As String
Dim OutputKey As String
Dim OutputFormatKey As String
Dim RouteKey As String
Dim RouteCount As Long
Dim FilterKey As String
Dim GraphKey As String
Dim Idx2 As Long
Dim RegFileCh As Integer
Dim kb As String
Dim kbLog As String
Dim DevName As String
    
    If ProfileName <> "" Then
'modProfile.LoadProfile (AisDecoder,AisDecoder.txt)
'frmRouter.Form_Load
'modRouter.ResetRouter
        
'frmRouter.MenuProfileOpen_Click
'Profilecfg.ProfileOpen
'Profilecfg.CmdOk_Click
'modProfile.CloseProfile (AisDecoder,AisDecoder)    - Clears RegFileName
        
'frmRouter.MenuProfileOpen_Click    'New with profile loaded
'Profilecfg.ProfileOpen
'Profilecfg.CmdOk_Click
'modProfile.CloseProfile (AisDecoder,AisDecoder.new)
'modRouter.ResetRouter
'frmRouter.Form_Load
'modProfile.LoadProfile (Broos,Broos.txt)

        kbLog = "Saving Profile " & ProfileName
    Else
        kbLog = "Saving all Profiles"
    End If
        
    If RegFileName = "" Then
'modProfile.CloseProfile (AisDecoder,"")
WriteLog kbLog & " to registry"
        Call DeleteProfile(ProfileName)
    Else
'modProfile.LoadProfile (AisDecoder,AisDecoder.txt)
'modProfile.CloseProfile (AisDecoder,AisDecoder.new)
'modProfile.LoadProfile (Broos,Broos.txt)
        RegFileCh = FreeFile
        kb = TempPath & "Router" & RegFileName
'Dont report save to file WriteLog kbLog & " to " & kb
        Open kb For Output As #RegFileCh
'Cant save file name as files will always differ
'        Print #RegFileCh, kb
    End If
    
    If ProfileName = "" Then    'Output all the registry
        ProfileKey = ROUTERKEY
    Else
        ProfileKey = ROUTERKEY & "\Profiles\" & ProfileName
    End If
        
    Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, ProfileKey)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, ROUTERKEY, "CurrentProfile", CurrentProfile, REG_SZ)
    If ActivationKey <> "" Then
        Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, ROUTERKEY, "ActivationKey", ActivationKey, REG_SZ)
    End If
'When the profile is first loaded, the current profile is saved to
'C:\Documents and Settings\jna\Local Settings\Temp\RouterProfile1.txt
    For Idx = 1 To UBound(Sockets)
'Dont save a TCP server Stream - it is transient and
'only created when a remote client tries to connect
        If Sockets(Idx).State <> -1 _
        And Sockets(Idx).Winsock.Oidx <= 0 Then
            If SocketKey = "" Then
                SocketKey = ProfileKey & "\Sockets"
                Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, SocketKey)
            End If
            SocketKey = ProfileKey & "\Sockets\" & Sockets(Idx).DevName
            Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, SocketKey)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "Handler", Sockets(Idx).Handler, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "Direction", Sockets(Idx).Direction, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "Enabled", Sockets(Idx).Enabled, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "Graph", Sockets(Idx).Graph, REG_SZ)
'Note Hidx must not be used because if socket not enabled then
'Hidx will be 0 (but we still want to save the handler settings)
            Select Case Sockets(Idx).Handler
            Case Is = 0            '0 = Winsock
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "Server", Sockets(Idx).Winsock.Server, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "Protocol", Sockets(Idx).Winsock.Protocol, REG_SZ)
                Select Case Sockets(Idx).Winsock.Protocol
                Case Is = sckTCPProtocol
                    Select Case Sockets(Idx).Winsock.Server
                    Case Is = 0 'client
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RemoteHost", Sockets(Idx).Winsock.RemoteHost, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RemotePort", Sockets(Idx).Winsock.RemotePort, REG_SZ)
                    Case Is = 1 'server
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "LocalPort", Sockets(Idx).Winsock.LocalPort, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "PermittedStreams", Sockets(Idx).Winsock.PermittedStreams, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "PermittedIPStreams", Sockets(Idx).Winsock.PermittedIPStreams, REG_SZ)
                    End Select
                Case Is = sckUDPProtocol
                    Select Case Sockets(Idx).Direction
                    Case Is = 0 'both
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "LocalPort", Sockets(Idx).Winsock.LocalPort, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RemoteHost", Sockets(Idx).Winsock.RemoteHost, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RemotePort", Sockets(Idx).Winsock.RemotePort, REG_SZ)
                    Case Is = 1 'Input
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "LocalPort", Sockets(Idx).Winsock.LocalPort, REG_SZ)
                    Case Is = 2 'Output
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RemoteHost", Sockets(Idx).Winsock.RemoteHost, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RemotePort", Sockets(Idx).Winsock.RemotePort, REG_SZ)
                    End Select
                End Select
            Case Is = 1            '1 = COMM
'Remove \\.\
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "CommName", Sockets(Idx).Comm.Name, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "BaudRate", Sockets(Idx).Comm.BaudRate, REG_SZ)
            Case Is = 2            '2 = File
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "SocketFileName", Sockets(Idx).File.SocketFileName, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "ReadRate", Sockets(Idx).File.ReadRate, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, SocketKey, "RollOver", Sockets(Idx).File.RollOver, REG_SZ)
            Case Is = 3            '3 = TTY
'Doesnt use a separate config
            Case Is = 4             'LoopBack
            Case Else
                MsgBox "Handler " & Sockets(Idx).Handler & " not in use", , "Save Profile"
            End Select
            
Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, SocketKey & "\OutputFormat")
OutputFormatKey = SocketKey & "\OutputFormat\IEC"
Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, OutputFormatKey)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, OutputFormatKey, "Enabled", Sockets(Idx).IEC.Enabled, REG_SZ)
OutputFormatKey = SocketKey & "\OutputFormat\PlainNmea"
Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, OutputFormatKey)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, OutputFormatKey, "Enabled", Sockets(Idx).OutputFormat.PlainNmea, REG_SZ)
OutputFormatKey = SocketKey & "\OutputFormat\OwnShip"
Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, OutputFormatKey)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, OutputFormatKey, "MMSI", Sockets(Idx).OutputFormat.OwnShipMmsi, REG_SZ)

'Now the Routes
            For Ridx = 1 To UBound(Sockets(Idx).Routes)
                Idx2 = Sockets(Idx).Routes(Ridx).AndIdx
                If Idx2 > 0 Then
                    If RouteKey = "" Then
                        RouteKey = ProfileKey & "\Routes"
                        Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, RouteKey)
                        End If
                    RouteCount = RouteCount + 1
                    RouteKey = ProfileKey & "\Routes\Route" & RouteCount
                    Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, RouteKey)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, RouteKey, "DevName1", Sockets(Idx).DevName, REG_SZ)
Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, RouteKey, "DevName2", Sockets(Idx2).DevName, REG_SZ)
'If Sockets(Idx).Routes(Ridx).Enabled = True Then
'    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, RouteKey, "Enabled", "True", REG_SZ)
'Else
'    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, RouteKey, "Enabled", "False", REG_SZ)
'End If
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, RouteKey, "Enabled", Sockets(Idx).Routes(Ridx).Enabled, REG_SZ)
                End If
            Next Ridx
        End If
    Next Idx
    
'Now the Filter (all)
    FilterKey = ProfileKey & "\Filters"
    Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, FilterKey)
    Idx = SourceDuplicateFilter.DmzIdx
    If Idx > 0 Then
        DevName = Sockets(SourceDuplicateFilter.DmzIdx).DevName
    Else
        DevName = ""
    End If
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, FilterKey, "DmzDeviceName", DevName, REG_SZ)
    FilterKey = ProfileKey & "\Filters\SourceDuplicates"
    Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, FilterKey)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, FilterKey, "RemoveSourceDuplicates", SourceDuplicateFilter.Enabled, REG_SZ)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, FilterKey, "RemoveSourceNotAivdm", SourceDuplicateFilter.OnlyVdm, REG_SZ)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, FilterKey, "RemoveSourceNotMmsi", SourceDuplicateFilter.RejectMmsi, REG_SZ)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, FilterKey, "RejectPayloadErrors", SourceDuplicateFilter.RejectPayloadErrors, REG_SZ)

'Now the Graph Key
    GraphKey = ProfileKey & "\Graphs"
    Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, GraphKey)
    GraphKey = ProfileKey & "\Graphs\" & "Graph1"
    Call CreateProfileKey(RegFileCh, HKEY_CURRENT_USER, GraphKey)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, GraphKey, "ExcelUpdateInterval", ExcelUpdateInterval, REG_SZ)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, GraphKey, "ExcelRange", ExcelRange, REG_SZ)
    Call CreateProfileKeyValue(RegFileCh, HKEY_CURRENT_USER, GraphKey, "ExcelUTC", ExcelUTC, REG_SZ)
    If RegFileCh <> 0 Then
        Close RegFileCh
        RegFileCh = 0
    End If

End Sub

Private Sub CreateProfileKey(RegFileCh As Integer, lPredefinedKey As Long, sNewKeyName As String)
    If RegFileCh = 0 Then
        Call CreateNewKey(lPredefinedKey, sNewKeyName)
    Else
        Print #RegFileCh, "[HKEY_CURRENT_USER\" & sNewKeyName & "]"
    End If
End Sub

Public Sub CreateProfileKeyValue(RegFileCh As Integer, lPredefinedKey As Long, sKeyName As String, _
sValueName As String, vValueSetting As Variant, lValueType As Long)
    If RegFileCh = 0 Then
        Call SetKeyValue(lPredefinedKey, sKeyName, sValueName, vValueSetting, lValueType)
    Else
        Print #RegFileCh, """" & sValueName & """" & "=" & """" & vValueSetting & """"
    End If
End Sub

'fname is optional to allow frmRouter to be saved when it is minimized
Public Sub SaveWindows(Optional fName As String)
'Save all the windows - not saved to the registry file
'because we do not want to prompt for changes to windowing
Dim arry() As String
Dim f As Form
Dim Position(6) As String
Dim ProfileKey As String
Dim WindowKey As String
Dim SocketKey As String
Dim KeyValue As String
Dim kb As String
    ProfileKey = ROUTERKEY & "\Profiles\" & CurrentProfile
    For Each f In Forms
        If (fName = "" Or f.Name = fName) And f.Name <> "frmSysTray" Then
            arry = Split(f.Caption, " ")
            Select Case f.Name
            Case Is = "frmRouter"
                WindowKey = ProfileKey
                Call CreateProfileKeyValue(0, HKEY_CURRENT_USER, WindowKey, "ViewData", frmRouter.MenuViewInoutData.Checked, REG_SZ)
                Call CreateProfileKeyValue(0, HKEY_CURRENT_USER, WindowKey, "ViewSockets", frmRouter.MenuViewInoutSockets.Checked, REG_SZ)
                Call CreateProfileKeyValue(0, HKEY_CURRENT_USER, WindowKey, "ViewRoutes", frmRouter.MenuViewInoutRoutes.Checked, REG_SZ)
                Call CreateProfileKeyValue(0, HKEY_CURRENT_USER, WindowKey, "ViewGraph", frmRouter.MenuViewInoutGraph.Checked, REG_SZ)
'YOU cant do this as it will fail when the profile is loaded
'        Case Is = "frmTTY"
'Check to see if this socket has been set up in the registry
'It may not if the TTY has just been created but the profile
'not saved yet
'            WindowKey = ProfileKey & "\Sockets\" & arry(0)
'            KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "Window")
'            If KeyValue = "" Then
'                Call CreateProfileKey(0, HKEY_CURRENT_USER, WindowKey)
'            End If
'Stop

            Case Else
                WindowKey = ProfileKey & "\Sockets\" & arry(0)
            End Select

            If f.WindowState = vbMinimized Then
                f.Show  'must show to change to vbnormal
                f.WindowState = vbNormal    'cant normalise when hidden
                f.Hide  'hiding while normal retains dimensions
            End If

'will be -48000 if we cant get dimensions
'This should not happen but is a backstop
            If f.Left >= 0 Then
            
                Position(0) = f.Left
                Position(1) = f.Top
                Position(2) = f.Width
                Position(3) = f.Height
                Position(4) = f.WindowState
                Position(5) = f.Visible
                Call CreateProfileKeyValue(0, HKEY_CURRENT_USER, WindowKey, "Window", Join(Position, ","), REG_SZ)
                kb = kb & WindowKey & "\Window = " & Join(Position, ",") & vbCrLf
            End If
        End If
    Next f

End Sub

Public Function RestoreWindow(Window As Form) As String
Dim Position() As String
Dim ProfileKey As String
Dim WindowKey As String
Dim KeyValue As String
Dim kb As String
Dim arry() As String
Dim CurrState As Integer
Dim ReqVisible As Boolean

    If Window.Name = "frmSysTray" Then Exit Function
    
    ProfileKey = ROUTERKEY & "\Profiles\" & CurrentProfile
    Select Case Window.Name
    Case Is = "frmRouter"
'        ResizeOk = True
        WindowKey = ProfileKey
        KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "ViewData")
'If not in registry then dont set
        If KeyValue <> "" Then
            frmRouter.MenuViewInoutData.Checked = BooleanA(KeyValue)
        End If
        KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "ViewSockets")
        If KeyValue <> "" Then
            frmRouter.MenuViewInoutSockets.Checked = BooleanA(KeyValue)
        End If
        KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "ViewRoutes")
        If KeyValue <> "" Then
            frmRouter.MenuViewInoutRoutes.Checked = BooleanA(KeyValue)
        End If
        KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "ViewGraph")
        If KeyValue <> "" Then
            frmRouter.MenuViewInoutGraph.Checked = BooleanA(KeyValue)
'This has to be done here otherwise it will not be started
'when the profile is loaded
            If frmRouter.MenuViewInoutGraph.Checked = True Then
                If ExcelOpen = False Then
                    Call CreateWorkbook
                End If
                frmRouter.GraphTimer.Enabled = True
            End If
       End If
'DONT move/resize frmRouter - done in frmRouter Load
        Exit Function
    Case Else
        arry = Split(Window.Caption, " ")
        WindowKey = ProfileKey & "\Sockets\" & arry(0)
    End Select
    KeyValue = QueryValue(HKEY_CURRENT_USER, WindowKey, "Window")
'Call DisplayWindow("modProfile.RestoreWindow")
    kb = kb & WindowKey & "\Window = " & KeyValue & vbCrLf
    Position = Split(KeyValue, ",")

'Ensure we have the values in the registry
    If UBound(Position) = 6 Then
'Causes a from resize event as we are only repositioning
'may want to suppress the resize event
        Window.Move CSng(Position(0)), CSng(Position(1)), CSng(Position(2)), CSng(Position(3))
        If Position(4) <> 0 Then Window.WindowState = Position(4)
'Call DisplayWindow("modProfile.RestoreWindow")
    End If
'Added

'Call DisplayWindow("modProfile.RestoreWindow")
'Return what weve done for debugging
    RestoreWindow = kb
End Function
