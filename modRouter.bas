Attribute VB_Name = "modRouter"

Option Explicit
'Copyright@ 2009-17 Neal Arundale

'Public Const V44 = False    'false (dont complile)
'Public Const v48 = True    'If declared here it is private
'to modRouter. For all modules it must ve decalred in the
'Project make box

'The user only sets up a Route from one Socket to
'another Socket, from this we construct what is in effect a
'forwarding table (Forewards()) for both Sockets.
'This is because each route can be in one or both directions or
'bi-directional.
'For programming convenience we also keep a record of the
'Route the User has set up. This makes it much easier to
'check if a route already exists, as well as constructing the
'forwarding table from the registry.
'Note although the program uses indexes into the Routes and Forwarding
'arrays, this is transparent to the user.
'When the user data is saved to the registry, these indexes must
'not be saved as they may (and probabaly will) be different
'when the registry details are recovered.
'Note also it is important the Route is only save once so the route
'is always saved in the lower socket index of the two sockets
'specifying the array. This makes it simple to see if a Route exists
'and to check the number of Routes in use.
'The Input to a socket is considered to be the Data arriving
'at the Router from an external source. Output is considered to be
'Data that is sent to the outside world. This can be confusing because
'Data that is to be displayed in a Window or to a File, when Forwarded
'by the Router to the Handler isw actually an Input to the Handler
'Which the handler will Output to the Oustide World.

'In effect we have World > Handler > Router > Handler > World
'World(Output) > Handler(Input)
'Handler(Output) > Routeing (Input) [Receive, Read Data]
'Routing (Output) > Handler (Input) [Send, Write Data]
'Handler (Output) > World (Input)   [Sent Data]
'This is confusing as the Input to the World is normally called
'Output (from the Router).

'===========================================
'http://support.microsoft.com/kb/176085
'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public nid As NOTIFYICONDATA

'Public ResizeOk As Boolean 'Supress when loading or in sysTray
'===========================================

Public Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

'Public Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
'Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'Private Const OFFSET_4 = 4294967296#

Public Const MAX_PROFILES = 20
Public Const MAX_FILES = 10
Public Const MAX_TTYS = 10
'Public Const MAX_COMMS = 10    'moved to modComm
'Public MAX_SOCKETS As Long '= 10       'moved to modAIS
Public MAX_WINSOCKS As Long  '= 10
Public MAX_TCPSERVERSTREAMS As Long '= 10  'On any socket
Public MAX_ROUTES As Long '= 10    'Total on ALL Sockets
Public MAX_SOCKETROUTES As Long '= 10    'Total on any ONE Socket
Public MAX_SOCKETFORWARDS As Long '= 10  'On any ONE socket - excludes TCP Stream Clients
Public Const MAX_LOOPBACKS = 10     'Loopbacks
Public Const DEFAULT_HANDLER_INDEX = 0
Public Const Qmax = 100
'Public Const MAX_COMM_OUTPUT_BUFFER_SIZE = 50000 (now defined in modComm)
Public Const ROUTERKEY = "Software\Arundale\NmeaRouter"

Public Const REGISTER_SCRIPT = "http://arundale.com/docs/ais/register.php"
Public Const MY_EMAIL_ADDRESS = "neal@arundale.com"
'Variables set when NmeaRouter is started
Public NmeaRouterIcon As String
Public ModuleCode As Long
Public ActivationKey As String
Public SupressNmeaTerminatorMsg As Boolean
Public AcceptLForCRasNmeaTerminator As Boolean

Public DownloadURL As String
Public MyNewVersion As New clsNewVersion

'Variables set when NmeaRouter is started
'and reset when new Profile is loaded
Public CurrentProfile As String

Public Declare Function ShellExecute _
                            Lib "SHELL32.DLL" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNORMAL = 1  ' aka SW_NORMAL
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'Purpose     :  Return the error message associated with LastDLLError
'Inputs      :  [lLastDLLError]               The error number of the last DLL error (from Err.LastDllError)
'Outputs     :  Returns the error message associated with the DLL error number

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public InCount As Double
Public OutCount As Double

'The private types containg the data we want to load into the
'handler from the Registry or Save to the Registry
'These are the same as the Fields we ask for on the
'Handlercfg windows.
Private Type WinsockDef 'Must be Private as there is a winsockdef in modWinsock
    Server As Long      '0=Client
                        '1=Server
                        '-1=undefined
    Protocol As Long    '0=TCP
                        '1=UDP
    RemoteHost As String
    RemotePort As String
    LocalPort As String
    LocalIP As String
    RemoteHostIP As String      'Where the last message came from
    Streams() As Long   'TCP Server Streams this socket owns
    PermittedStreams As Long    'Max Permitted by user
    PermittedIPStreams As Long
    Oidx As Long     'TCP Server Idx Owning this stream
    Sidx As Long    'Stream Index for this Stream
End Type

Private Type CommDef    'must be private
    BaudRate As String
    Name As String      'Short name COM6 (to see if weve changed it)
    VCP As String       'Associated VCP if any
    AutoBaud As Boolean
End Type

Private Type FileDef
    SocketFileName As String  'Full Name including path
    RollOver As Boolean
    ReadRate As Long    'Msg/minute when reading a file
End Type

'This is used to split the data held in the buffers
Public Type AdrDef 'AddressedDataRecord
    SequenceNo As Currency  '64 bit signed integer in 1/10,000's
    Source As Long  'Socket (Idx1)
    UtcUnix As Long    'Unix Time (Secs since 1 Jan 1970)
    Data As String      'Individual Sentences (No terminator)
    Info As String      'Information eg Reason for Reject filter
End Type

'Routes are between Idx and Sockets(Idx).Routes(Ridx).AndIdx
'Was                        Sockets(Idx).Routes(Ridx)
'Idx must be the lower Idx, AndIdx is the higher Idx
'SortIdx(Idx1,Idx2) ensures the order is correct

Private Type RecorderDef
    Enabled As Boolean      'Recoder is required
    Output As frmRecorder 'Only created Output
End Type

Private Type VdoDef
    SequenceNo As Currency
    Source As Long
    UtcUnix As Long
    Destination As Long
    Data As String
    LastVdoUpdate As Long
End Type

Public Type OutputFormatDef
    PlainNmea As Boolean
    OwnShipMmsi As String
End Type

Public Type IecFormatDef
    Enabled As Boolean
    Time As Boolean
    Destination As Boolean
    Information As Boolean
    Source As Boolean
    Counter As Boolean
    Group As Boolean
End Type

Public Type Routedef
    AndIdx As Long
    Enabled As Boolean
    ForwardCount As Long
    ReverseCount As Long
End Type

Public Type SocketDef
    Fragment As String  'Partial senetence received on this socket
                        'CFLF not received
    Buffer(1 To Qmax) As AdrDef  'Full when 1 spare slot
    Qrear As Long   'added to rear
    Qfront As Long  'Points to 1 AHEAD of next one to be removed
    QLost As Long
    DevName As String   'Determined by the User
                        'PORT_number
                        'COMMnumber
                        'Name
                        'TEXTBOX_number
    Handler As Long     'The Type of handler
'These are the same is the selected cboHandler.ListIndex
                        '-1 = not determined
                        '0 = Winsock
                        '1 = COMM
                        '2 = File
                        '3 = TTY
                        '4 = Loopback
    Hidx As Long         'Handler Array index comms(Hidx), Winsock(Hidx)
    State As Integer    'The State returned by the Handler
                        '-1 =  Handler Object (inc DCB) Index = Nothing
                        '0 = Handler Object Index exists (but Port Closed)
                        '11 = Handler Opening (or trying to) Open Port
                        '1  = Port Open
                        '18 = Handler Closing (or trying to) Close Port
                        '21 = Serial Data Loss
                        '22 = Serial Pending Output buffer not empty
    errmsg As String    'Normally the error reported by the handler
    Enabled As Boolean  'Set by the user to suspend this socket
    MsgCount As Double
    LastMsgCount As Long    'used by speed timer
    Graph As Boolean    'Set by user to enable/disable Graph series
    LostMsgCount As Double
    TryCount As Long
    Chrs As Long        'Since last Speed time
    ResetCount As Long      'No of times no chrs received on each reconnect
    Routes() As Routedef      'Route index
                        'Only created if Idx less than Ridx, as its
                        'included for one Idx irrespective of Direction
    Forwards() As Long    'Socket number of destination
                        'Allocated sequentially
    Direction As Long   '0=both Input,Output
                        '1=Input Only
                        '2=Output Only
    OutputFormat As OutputFormatDef
    IEC As IecFormatDef
    Winsock As WinsockDef
    Comm As CommDef
    File As FileDef
    Recorder As RecorderDef
    VDO As VdoDef
'    AisFilter As clsAisFilter
End Type

Public Sockets() As SocketDef   'Array of data

Public CurrentSocket As Long
'Public Dpys() As DisplayDef
Public TTYs() As clsTTY     'Array of Objects
Public LoopBacks() As clsLoopBack
Public Files() As clsFile
Public Stopped As Boolean    '
'Public TempPath As String   'moved to modGeneral (used for log files)
'Public LogFileName As String   'moved to modGeneral (used for log files)
'Public LogFileCh As Integer   'moved to modGeneral (used for log files)
'Always create duplicate filter (we need to set initial values from profile)
Public SourceDuplicateFilter As clsDuplicateFilter 'applies to all sockets
'not used Public SourceNmeaFilter As clsNmeaFilter
Public DupeLogCh As Long
Public InputLogName As String   'test only
Public InputLogCh As Long   'test only
Public SentenceSequenceNo As Currency
Public jnasetup

Private StartUpCommand As String
Private cmdOptions() As String
Private cmdProfile As String
'Public to override the registry setting but only when the
'profile is first loaded on startup - set to false when used
Public cmdSysTray As Boolean
Public cmdDontQueryUnload As Boolean

Sub Main()
Dim Cancel As Boolean
Dim i As Long
Dim j As Long
'testing only Load frmDebugSerial

    On Error GoTo Main_Error

#If setactcode Then
    jnasetup = True
#End If
    
'INNO places the icon file
'[Icons]
'Name: "{group}\NmeaRouter"; Filename: "{app}\NmeaRouter.exe"; IconFilename:"{app}\router.ico"
'Load before any form are loaded to display the Icon
    NmeaRouterIcon = App.Path & "\router.ico"
    
'NOTE At Startup frmSysTray is loaded first
'Testing routine only
'    Call EnumAllKeys(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum")
    
    TempPath = LongFileName(Environ("TEMP") & "\")
    
    InputLogCh = -1 'open new file (used for debugging only)
    
    LogFileCh = -1  'force open for output
    WriteLog "NmeaRouter " & App.Major & "." & App.Minor & "." & App.Revision & " Started"
    If Is64bit Then
        WriteLog "Version = " & GetVersion1() & " (64 bit)"
    Else
        WriteLog "Version = " & GetVersion1() & " (32 bit)"
    End If
    
    Set MyNewVersion = New clsNewVersion
    DownloadURL = "arundale.com/docs/ais/"
    Call WriteLog("Checking " & DownloadURL & " for later version")
    
    Call MyNewVersion.CheckNewVersion(DownloadURL, False) '2nd arg displays debug info
    
    If MyNewVersion.Downloaded = True Then
'Must exit calling program otherwise install will fail because this exe is referenced
        Call WriteLog("Downloaded " & MyNewVersion.DownloadedURL)
        If MyNewVersion.SetupExecuted = True Then
            Set MyNewVersion = Nothing
            End 'Exit program
        Else
            Call WriteLog("SetupExecuted = " & MyNewVersion.SetupExecuted)
        End If
    End If
    Set MyNewVersion = Nothing
        
    WriteLog "User Default LocaleID = " & GetUserDefaultLCID()
    WriteLog "User has Administrator Rights = " & CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
            
'Check any command line options
    If Command$ <> "" Then
        Call WriteLog("Startup Cammand is " & Command$)
    End If
    StartUpCommand = LCase$(Command$)
'stop   'to test startup command
'    StartUpCommand = "/systray/profile=""Tcp Client"""
'    StartUpCommand = "systray/profile=""Tcp Client"""
'    StartUpCommand = "systray"
    If StartUpCommand <> "" Then
        If InStr(1, StartUpCommand, "/") = 0 Then
            ReDim cmdOptions(0)
            cmdOptions(0) = StartUpCommand
        Else
            cmdOptions = Split(StartUpCommand, "/")
        End If
        For i = 0 To UBound(cmdOptions)
            j = InStr(1, cmdOptions(i), "=")
            If j = 0 Then j = Len(cmdOptions(i))
            Select Case Trim$(Left$(cmdOptions(i), j))
            Case Is = "jna"
                jnasetup = True
            Case Is = "profile="
'cmdProfile is active only while loading the initial profile
                cmdProfile = Mid$(cmdOptions(i), j + 1)
'remove " from file name (if any)
                cmdProfile = Replace(cmdProfile, """", "")
            Case Is = "systray"
'cmdSystray is active for the entire session, and can only be set as a command
'line option, therefore it is NEVER saved to the registry
                cmdSysTray = True
            Case Is = "dontqueryexit"
                cmdDontQueryUnload = True
            End Select
        Next i
    End If
    
'Get any variables that must be set when the module is first loaded
'ie they are NmeaRouter variables
    If cmdProfile <> "" Then
        CurrentProfile = cmdProfile
    Else
        CurrentProfile = QueryValue(HKEY_CURRENT_USER, ROUTERKEY, "CurrentProfile")
    End If
    
    SupressNmeaTerminatorMsg = BooleanA(QueryValue(HKEY_CURRENT_USER, ROUTERKEY, "SupressNmeaTerminatorMsg"))
    AcceptLForCRasNmeaTerminator = BooleanA(QueryValue(HKEY_CURRENT_USER, ROUTERKEY, "AcceptLForCRasNmeaTerminator"))

'Get the ModuleCode from the ActivationKey - if any
    ActivationKey = QueryValue(HKEY_CURRENT_USER, ROUTERKEY, "ActivationKey")
    ModuleCode = CheckKey(ActivationKey)
'If activation key is invalid clear Key so frmRegister Key is blanked out
'otherwise Register Now is not displayed (Contact me is)
    If ModuleCode = -1 Then
        ActivationKey = ""
    End If
    Call WriteLog("Module Code = " & ModuleCode)
'If cant get Drive Serial No then module code = -1
'If No Activation Key Module Key = 0
    Select Case ModuleCode
    Case Is = 2
        MAX_SOCKETS = 81   'including TCP Client Streams
        MAX_WINSOCKS = 81   'including TCP Client Streams
        MAX_TCPSERVERSTREAMS = 80  'On any socket
        MAX_ROUTES = 80   'Total on ALL Sockets
        MAX_SOCKETROUTES = 80   'Total on any ONE Socket
        MAX_SOCKETFORWARDS = 80 'On any ONE socket - excludes TCP Stream Clients
    Case Else
        MAX_SOCKETS = 10
        MAX_WINSOCKS = 10
        MAX_TCPSERVERSTREAMS = 10  'On any socket
        MAX_ROUTES = 10   'Total on ALL Sockets
        MAX_SOCKETROUTES = 10   'Total on any ONE Socket
        MAX_SOCKETFORWARDS = 10 'On any ONE socket - excludes TCP Stream Clients
    End Select
    Load frmRouter  'v59
    Call ResetRouter
'v59a     Call LoadProfile(CurrentProfile, Cancel)    'v59
    
Main_Exit:
Exit Sub

Main_Error:
    Select Case err.Number
    Case Else
        MsgBox "Error " & err.Number & " - " & err.Description, vbCritical, "Main_Error"
    End Select
    Resume Next 'Only in sub Main (otherwise application does not terminate)
'    Resume Main_Exit
 End Sub

'Reset everything except any command line arguments
Sub ResetRouter()
Dim Cancel As Boolean

'If not loaded then not actioned
'Otherwise will query to save changes
'v59    Unload frmRouter
'    Set SourceDuplicateFilter = Nothing
'Reset any Global Variables held in modRouter
'If you don't initialise the first element you cant't tell
'if it exists and you get a subscript error when you try
'and check it.
    ReDim Sockets(1 To 1)
    Call ClearSocket(1) 'Set the initial values on sockets(1)
    ReDim Files(1 To 1)
    ReDim TTYs(1 To 1)  'must initialise first element
    ReDim LoopBacks(1 To 1)  'must initialise first element
    ReDim Comms(1 To 1)
    ReDim Winsocks(1 To 1)
'v59    Load frmRouter
    Call frmRouter.ResetDisplay
    Call LoadProfile(CurrentProfile, Cancel)    'v59a
End Sub

'Called when Socketcfg is loaded and by validate_txtDevNeme
'Returns the socket associated with a device name
'If this device name is not set up in Sockets returns 0
Public Function DevNameToSocket(DevName As String) As Long
Dim i As Long

    For i = 1 To UBound(Sockets)
        If Sockets(i).State <> -1 Then
'Note first free socket
            If UCase$(Sockets(i).DevName) = UCase$(DevName) Then
'Yes, use this socket
                DevNameToSocket = i
                Exit Function
            End If
        End If
    Next i
End Function

'Gets first available socket
'Returns -1 if none available
'Extends Sockets array up to MAX_SOCKETS
'Clears Sockets(Idx) if exists but State=-1
Public Function FreeSocket() As Long
Dim i As Long

'Test limits
#If False Then
If UBound(Sockets) < 10 Then
    ReDim Sockets(1 To 10)
    For i = 1 To 10
        ClearSocket (i)
'Force no free sockets
'        Sockets(i).State = 0
    Next i
End If
#End If
    
'Try & allocate a released socket (if any)
    For i = 1 To UBound(Sockets)
'DOESNT work for sockets
'        If Not Sockets(i) Is Nothing Then
            If Sockets(i).State = -1 Then
                Exit For
            End If
'        Else
'            Exit For
'        End If
'Debug.Print i & ":" & Sockets(i).State & "-" & Sockets(i).TryCount
    Next i

'If no released sockets, i will be the next one available
    FreeSocket = i

    If FreeSocket > MAX_SOCKETS Then    'no free sockets
        FreeSocket = -1
        WriteLog "No free sockets, limit is " & MAX_SOCKETS
    Else
'still at least 1 free socket
        If FreeSocket > UBound(Sockets) Then
'Need to create a new array element
'Possible error array is tmapoarily locked
            On Error GoTo Locked
            ReDim Preserve Sockets(1 To FreeSocket)
            On Error GoTo 0
         End If
'Reset the initial values on the socket
'DONT for COMMS
        Call ClearSocket(FreeSocket)
    End If
    Exit Function
Locked:
    FreeSocket = -1
End Function

Public Sub ClearSocket(Idx As Long)
Dim i As Long
            
            With Sockets(Idx)
                If .TryCount > 0 Then
                    Call frmRouter.ResetTries(Idx)
                End If
                Call RemoveRecorder(Idx)
                .Fragment = ""
                For i = 1 To Qmax
                   .Buffer(i).Data = ""
                   .Buffer(i).Source = 0
                    .Buffer(i).UtcUnix = 0
                Next i
                'Full when 1 spare slot
                .Qrear = 0
                .Qfront = 0
                .QLost = 0
                .DevName = "Connection " & Idx
                .Handler = DEFAULT_HANDLER_INDEX
                .Hidx = 0
                .State = -1
                .errmsg = ""
                .Enabled = False
                .MsgCount = 0
                .Graph = False
                .LostMsgCount = 0
                .Chrs = 0
                .ResetCount = 0
                ReDim .Routes(1 To 1)
                ReDim .Forwards(1 To 1)
                .Direction = -1      'undefined
                .Winsock.Server = -1   '-1=undefined
                .Winsock.Protocol = sckUDPProtocol
                .Winsock.RemoteHost = ""
                .Winsock.RemotePort = ""
                .Winsock.LocalPort = ""
                .Winsock.RemoteHostIP = ""
                ReDim .Winsock.Streams(1 To 1)
                .Winsock.Oidx = -1
                .Winsock.Sidx = -1
                .Winsock.PermittedStreams = 1
                .Winsock.PermittedIPStreams = 1
                .Comm.BaudRate = ""
                .Comm.Name = ""
                .Comm.VCP = ""
                .Comm.AutoBaud = False
                .File.SocketFileName = ""
                .File.ReadRate = 0
                .File.RollOver = False
                .Recorder.Enabled = False
                .VDO.SequenceNo = 0
                .VDO.Source = 0
                .VDO.UtcUnix = 0
                .VDO.Destination = 0
                .VDO.Data = ""
                .VDO.LastVdoUpdate = 0
                If Not .Recorder.Output Is Nothing Then
                    Unload .Recorder.Output
                End If
            End With
End Sub

'Returns the Handler Index for a given socket (Socket(Idx).Handler)
'from the Handler must be set
'This is used to set up Hidx on Sockets following creation by the
'Handler
Public Function GetHidx(Idx As Long) As Long
Dim i As Long
    GetHidx = -1    'Default cant find
    Select Case Sockets(Idx).Handler
    Case Is = 0     'Winsock
        For i = 1 To UBound(Winsocks)
            If Winsocks(i).sIndex = Idx Then
                GetHidx = i
                Exit For
            End If
        Next i
    Case Is = 1     'Serial
        For i = 1 To UBound(Comms)
            If i <= UBound(Comms) Then
                If Not Comms(i) Is Nothing Then
                    If Comms(i).sIndex = Idx Then
                        GetHidx = i
                        Exit For
                    End If
                End If
            End If
        Next i
    Case Is = 2     'File
MsgBox "File not yet implemented"
    Case Is = 3     'TTY
        For i = 1 To UBound(TTYs)
            If TTYs(i).sIndex = Idx Then
                GetHidx = i
                Exit For
            End If
        Next i
    Case Is = 4     'loopbacks
        For i = 1 To UBound(LoopBacks)
            If LoopBacks(i).sIndex = Idx Then
                GetHidx = i
                Exit For
            End If
        Next i
        
    Case Else
        MsgBox "Not available in GetHidx"
    End Select
End Function


'Data can by multiple sentences, but each sentence will be a
'complete sentence terminated with <CRLF>
'Should not contain a trailing fragment
'Looks up the Forwarding.
'Constructs the Adr record
'Queues each sentence to each Destination socket buffer
Public Function ForwardData(Data As String, Source As Long)
Dim bytestoforward&
Dim Fidx As Long
Dim i As Long
Dim j As Long
Dim UtcUnix As Long
Dim Destination As Long
'Dim Sidx As Long    'TCP Server Stream Index
Dim kb As String
Dim myWinsock As Winsock
'Dim myobj As Object
Dim myWinsock1 As WinsockDef
Dim DataLen As Long
Dim SdfOffset As Long   'SourceDuplicateFilter -1 = not a duplicate
Dim RejectData As Boolean 'True if this sentence is rejected
Dim DataOutput As String    'Individual senetences - no CRLF
Dim OutputInfo As String    'USed for NMEA Comment Block
Dim kbT As String   'Term Output Only
Dim Aterm As String

'Call WriteInputLog(Data, Source) (Debug only)
On Error GoTo Forward_Error
'On Error GoTo 0 'debug
    
'   Covnvert System Time into UtcUnix Time
'Multiple sentences are output in one hit <crlf> terminated
'    Call frmRouter.TermOutput(Data, Source)

'We MUST check Source is valid because althouh a Socket may have been opened
'the Source may not yet have been setup
    If Source < 1 Or Source > UBound(Sockets) Then
        Exit Function
    End If

    DataLen = Len(Data)
'v59    If DataLen = 0 = 0 Then 'Should not happen
    If DataLen = 0 Then 'v59 Should not happen
        Exit Function
    End If
'Split DataRcv into individual lines and add to received list
    Sockets(Source).Fragment = Sockets(Source).Fragment & Data
'    Sockets(Source).Chrs = Sockets(Source).Chrs + Len(Data)
    Sockets(Source).Chrs = Sockets(Source).Chrs + 1 'sentences
    j = InStr(1, Sockets(Source).Fragment, vbCrLf)
    
    If j = 0 Then   'NO CRLF
        If SupressNmeaTerminatorMsg = False Then
            If Len(Data) >= 1 Then
                Aterm = "<" & AscB(Mid$(Data, Len(Data), 1)) & ">"
            End If
            If Len(Data) >= 2 Then
                Aterm = Aterm & "<" & AscB(Mid$(Data, Len(Data) - 1, 1)) & ">"
            End If
            If Aterm <> "" Then Aterm = " + " & Aterm
            Call frmDpyBox.DpyBox(Len(Data) - 2 & Aterm & " bytes received without NMEA <CR><LF> sentence terminator" & vbCrLf & Data & " is rejected" & vbCrLf, 5, "NMEA Delimiter Error")
        End If
'Stop
    End If
    

'Fragments do not enter the DO loop
    Do Until j = 0
'Multiple sentences are output singly <crlf> terminated
'This is the first place individual complete sentences
'are found
'SequenceNo 64-bit currency
'Reset for this sentence
        RejectData = False
        OutputInfo = ""
        SentenceSequenceNo = SentenceSequenceNo + 1
        UtcUnix = UnixNow

'Terminated debug
' If Left$(Sockets(Source).Fragment, 1) <> "!" Then
' kb = Asc(Left$(Data, 1))
' End If
        
'<CR><LF> has now been removed at this point
        DataOutput = Left$(Sockets(Source).Fragment, j - 1)
'        Call frmRouter.TermOutput(DataOutput & vbCrLf, Source)

'Check we do not have a 0 length sentence or not termination crlf
'kb = Len(Sockets(Source).Fragment)
'Count the incoming messages even if not connected to an output
        Sockets(Source).MsgCount = Sockets(Source).MsgCount + 1   'sent
        If SourceDuplicateFilter.Enabled = True Then
            SdfOffset = SourceDuplicateFilter.IsDuplicate(SentenceSequenceNo, Source, DataOutput)
            If SdfOffset <> -1 Then
                RejectData = True
                OutputInfo = OutputInfo & "|Src Duplicate Offset=" & SdfOffset
            End If
        End If
        If SourceDuplicateFilter.OnlyVdm = True Then
'Care don not overwrite RejectData if already true
            If IsAivdm(DataOutput) = False Then
                RejectData = True
                OutputInfo = OutputInfo & "|Src Not __VDM"
    Sockets(Source).MsgCount = Sockets(Source).MsgCount - 1   'Marcus
            End If
        End If
        If SourceDuplicateFilter.RejectMmsi <> "" Then
            kb = DecodeMMSI(DataOutput)
            If kb = SourceDuplicateFilter.RejectMmsi Then
                RejectData = True
                OutputInfo = OutputInfo & "|Src RejectMMSI_" & kb
    Sockets(Source).MsgCount = Sockets(Source).MsgCount - 1   'Marcus
            End If
        End If
'UFO's
        If SourceDuplicateFilter.RejectPayloadErrors = True Then
            If IsPayloadError(DataOutput) <> "" Then
                RejectData = True
            End If
        End If

'Display before the reject is sent to DMZ
        If frmRouter.MenuViewInoutData.Checked = True Then
            kbT = DataOutput & " [Src:" & Sockets(Source).DevName
            If OutputInfo <> "" Then
                kbT = kbT & OutputInfo
            End If
            kbT = kbT & "]" & vbCrLf
            Call frmRouter.TermOutput(kbT)
        End If
        
        
        If RejectData = False Then
'Non rejected sentence queued to all forwarded sockets
                Call ForwardSentence(SentenceSequenceNo, Source, UtcUnix, DataOutput)
        Else
'Rejected Complete sentence queued to the DMZ
'If set to a socket
            If SourceDuplicateFilter.DmzIdx > 0 Then
                Call Queue(SentenceSequenceNo, Source, UtcUnix, SourceDuplicateFilter.DmzIdx, DataOutput, OutputInfo)
            Else
' - otherwise it is silently dropped here
            End If
        End If
        
'frmNotify.lblMessage = "This is a notification of a Split Sentence"
'This requires sorting then moving
        If j + 1 >= Len(Sockets(Source).Fragment) Then
            Sockets(Source).Fragment = ""
        Else
'MsgBox Sockets(Source).Fragment
'Can be nul string "" not sure why make j+1 above
'kb = Asc(Data) '(lf only)
            Sockets(Source).Fragment = Right$(Sockets(Source).Fragment, Len(Sockets(Source).Fragment) - j - 1)
        End If
        j = InStr(1, Sockets(Source).Fragment, vbCrLf)
    Loop    'end of all sentences

'v59 If j = 0 Then Stop

Exit Function
Forward_Error:
    Select Case err.Number
    Case Is = 9   'Ocassional subscript error ??
    Case Else
        MsgBox err.Number & " " & err.Description, , "modRouter.ForwardData"
    End Select
End Function

'Queues complete sentences to the output queues
'by scanning the forwarding for each route
'Sentences rejected by the input filter destined for
'the DMZ are queued by ForwardData
Private Function ForwardSentence(SequenceNo As Currency, Source As Long, UtcUnix As Long, Sentence As String, Optional Info As String)
Dim Fidx As Long
Dim Destination As Long
    For Fidx = 1 To UBound(Sockets(Source).Forwards)
        Destination = Sockets(Source).Forwards(Fidx)
'If 0 then no socket to forward to
        If UBound(Sockets) >= Destination And Destination > 0 Then
'Sockets(socket) is not an object so cannot be nothing

'MUST NOT check if the destination socket is open
'As when cmdStart is Clicked if may not be & if not we would
'miss sentences received until the distination has been opened
            Call Queue(SequenceNo, Source, UtcUnix, Destination, Sentence, Info)
        End If
    Next Fidx   'next Forward to Output Queue

End Function

'Was in frmRouter
'see Page & Wilson p80
'This will deal with multiple line in one DataRcv,but Forward
'outputs separate (un-terminated) lines to Queue
'This queue is for the Destination
'Sequence no is allocated here, unless input spoofed from VDO
Public Sub Queue(SequenceNo As Currency, Source As Long, UtcUnix As Long, Destination As Long, Data As String, Info As String)
Dim Hidx As Long

'Removed to try and find where the error that causes port open error occurs
    On Error GoTo Queue_err
'On Error GoTo 0 'debug
    With Sockets(Destination)
        If .Handler = 0 Then    'Winsock
            Hidx = Sockets(Destination).Hidx
            .Winsock.RemoteHostIP = frmRouter.Winsock(Hidx).RemoteHostIP   'Where it came from
'when a TCP Client terminated the connection it does not
'always get set to closed. You have to force it to listen.
'NEW This does not appear to be correct as
'Listen causes an error
'            If Winsock(Hidx).State = sckClosing Then
'                Winsock(Hidx).Listen
'            End If
        End If
'must be at least 1 spare slot in the Send List
        If .Qrear = Qmax Then
            .Qrear = LBound(.Buffer)
        Else
            .Qrear = .Qrear + 1
        End If
'frmRouter.StatusBar.Panels(1).Text = .QRear - .QFront
        If .Qrear = .Qfront Then 'Q full
'I believe it starts the Q again & you loose Ubound(.buffer) sentences
            .QLost = .QLost + 1
'frmRouter.StatusBar.Panels(1).Text = .QLost
'I think you could reset .Qfront by 1 to only loose the olIdx sentence
            If .Qfront = Qmax Then
                .Qfront = LBound(.Buffer)
            Else
                .Qfront = .Qfront + 1
            End If
        Else
'jna                .Buffer(.QRear) = SentenceRcv
            .Buffer(.Qrear).SequenceNo = SequenceNo
            .Buffer(.Qrear).Data = Data
            .Buffer(.Qrear).Source = Source
            .Buffer(.Qrear).UtcUnix = UtcUnix
            .Buffer(.Qrear).Info = Info

'Test number messages rather than put actual message into buffer
'                .Buffer(.QRear) = .MsgCount
'Index is the Source
'Debug.Print "In " & Index & "->" & Destination & ":" & .Buffer(.QRear)
            End If
        End With
'Sentence has been added to the Output Buffer
        If frmRouter.DequeTimeout.Enabled = False Then
            Call Deque(Destination, 100)
'Debug.Print "Deque " & Destination
        Else
'Debug.Print "No Deque " & Destination
        End If
Exit Sub
Queue_err:
    MsgBox "Queue Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "Socket " & Destination, , "Queue Error"
End Sub

'Was in frmRouter
'The Deque Timer sets the minimum interval before a Deque is attempted
'after the last data has been queued. This is to prevent a fast
'queue rate Dequeing every time data is received (in which case
'there would be no point in queing in the first place).
'The next data on received for output will trigger the Deque
'Destination is the index of the Dstination Queue we wish to output
'The ADR should only contain Data destined for this destination
'Note if we are using the Addressed data (Adr) as the passed
'Argument to the Output Handler, the Handler Routine MUST not
'be in a Form Module as the DataType cannot be passed
'It must be in a Public Module or a CLass Module
Public Sub Deque(Destination As Long, Timeout As Long)
Dim Qleft
Dim SendErr As Boolean
Dim kb As String
Dim Hidx As Long

'Removed to try and find where the error that causes port open error occurs
    On Error GoTo Deque_err
'On Error GoTo 0 'debug
    Hidx = Sockets(Destination).Hidx
'There is a possibility data may be received for a Destination
'before the Destination socket has been set up
    If Hidx < 1 Then
'Do this by disabling forwarding (Not quick enough)
'Stop
        Exit Sub
    End If
'Deal with any issues regarding deque for different handlers
    Select Case Sockets(Destination).Handler
    Case Is = 0     'Winsock
        With frmRouter.Winsock(Hidx)
'check if winsock is connected
            Select Case .Protocol
            Case Is = sckTCPProtocol '0
'If outputting and not connected, state will be listening
                Select Case .State
'this will happen if the Winsock has been connected
'but has been closed by the client
                Case Is = sckClosed
'Now controlled by Reconnect  .Listen
                    Exit Sub
'If the client is completely connected then don't Deque
                Case Is <> sckConnected
kb = .State
                    Exit Sub
                End Select
            Case Is = sckUDPProtocol
'Ensure UDP has completed opening for output
                If .State <> sckOpen Then
'                Exit Sub
                End If
            End Select
        End With
    Case Is = 1     'Serial
        If Comms(Hidx).PendingOutputLen >= MAX_COMM_OUTPUT_BUFFER_SIZE Then

'            Sockets(Destination).State = 21 'Serial data loss
        End If
    End Select  'Handler
'Now Dequeue
'Debug.Print "Dequeue " & Destination
    With Sockets(Destination)
        frmRouter.DequeTimeout.Interval = Timeout
'So Loopback actions dequeue
        If Timeout > 0 Then
            frmRouter.DequeTimeout.Enabled = True
        End If
        Do Until .Qrear = .Qfront   'queue is not empty
            If .Qfront = Qmax Then
                .Qfront = LBound(.Buffer)
            Else
                .Qfront = .Qfront + 1
            End If
            
'Serial data loss
            If Sockets(Destination).State = 21 Then
                Sockets(Destination).LostMsgCount = Sockets(Destination).LostMsgCount + Len(.Buffer(.Qfront))
            Else
'Send Data for Formatting & Output
                Call OutputFormatter( _
                .Buffer(.Qfront).SequenceNo, _
                .Buffer(.Qfront).Source, _
                .Buffer(.Qfront).UtcUnix, _
                Destination, _
                .Buffer(.Qfront).Data, _
                .Buffer(.Qfront).Info)
'                Sockets(Destination).MsgCount = Sockets(Destination).MsgCount + Len(.Buffer(.QFront))
                Sockets(Destination).MsgCount = Sockets(Destination).MsgCount + 1
            End If
            
'            SendErr = SendData(Destination, .Buffer(.QFront))
'            Sockets(Destination).Chrs = Sockets(Destination).Chrs + Len(.Buffer(.QFront)) + 2
'Debug.Print "Out " & Destination & SendErr & ":" & .Buffer(.QFront)
'            DoEvents
        
            If frmRouter.DequeTimeout.Enabled = False Then
                Qleft = .Qrear - .Qfront
                If Qleft < 0 Then
                    Qleft = UBound(.Buffer) + Qleft
                End If
'Debug.Print "Timeout " & .QFront & ":" & .QRear & "=" & Qleft & "," & .QLost
'            Exit Sub
            End If
        Loop
    frmRouter.DequeTimeout.Enabled = False
    End With
Exit Sub
Deque_err:
    WriteLog "Deque Error " & Str(err.Number) & " " & err.Description _
    & " On " & Sockets(Destination).DevName
End Sub

'Filter/Formatter then send data to Destination Handler
'If A Recording Source will be 0, The Source may change between Recording
'& playback so we must not use it
Public Function OutputFormatter(SequenceNo As Currency, Source As Long, UtcUnix As Long, Destination As Long, Data As String, Optional Info As String)
Dim Hidx As Long
Dim CommentBlock As String
Dim Separator As String
Dim DataOutput As String
Dim kb As String
Dim i As Long
Dim j As Long
Dim kbT As String   'Term Output Only

    Hidx = Sockets(Destination).Hidx
    
'Ensure handler not been closed
    If Hidx < 1 Then
        Exit Function
    End If
    
    If Sockets(Destination).OutputFormat.PlainNmea = True Then
        i = InStr(1, Data, "$")
        If i = 0 Then
            i = InStr(1, Data, "!")
        End If
        If i = 0 Then
            Exit Function   'Nmea Sentence not found
        End If
        j = InStr(i, Data, "*")
        If j = 0 Or Len(Data) - j < 2 Then
            Exit Function   'Nmea Sentence not found
        End If
        Data = Mid$(Data, i, j + 2 - i + 1)
    End If
    
    If Not Sockets(Destination).Recorder.Output Is Nothing Then
'if recording
        If Sockets(Destination).Recorder.Output.shpRec.Visible Then
'Not a playback sentence
            If Source <> 0 Then
                Call Sockets(Destination).Recorder.Output.RecordData(SequenceNo, UtcUnix, Data)
            End If
        End If
        If Sockets(Destination).Recorder.Output.cmdStop.Enabled Then
            If Source = 0 Then
'data is recorded accept recorded data output
            Else
'Discard live data output as were playing back
                Exit Function
            End If
            
        End If
    End If
    
    If Sockets(Destination).OutputFormat.OwnShipMmsi <> "" Then
        kb = DecodeMMSI(Data)
        If kb = Sockets(Destination).OutputFormat.OwnShipMmsi _
        And clsSentence.NmeaSentenceType = "!AIVDM" Then
'Not recorded
            If Source <> 0 Then
                Data = VDMtoVDO(Data)
                Select Case clsSentence.AisMsgType
                Case Is = 1, 2, 3, 9, 18, 19, 27
                    With Sockets(Destination).VDO
                        .SequenceNo = 0
                        .Source = Source
                        .UtcUnix = UtcUnix
                        .Destination = Destination
                        .Data = Data
                        .LastVdoUpdate = 0
                    End With
                    Call frmRouter.VDOTimer_Timer
                End Select
            End If
        End If
    End If
                
'We cannot add comment block if playing back - we dont record source
'as it may have changed since recording
    If Sockets(Destination).IEC.Enabled = True And Source <> 0 Then
        CommentBlock = Separator & "c:" & UtcUnix
        Separator = ","
        CommentBlock = CommentBlock & Separator & "d:" & Sockets(Destination).DevName
        CommentBlock = CommentBlock & Separator & "s:" & Sockets(Source).DevName
        If Info <> "" Then
            CommentBlock = CommentBlock _
            & Separator & "i:" & Info
        End If
        CommentBlock = CommentBlock & "*[CRC]"
        CommentBlock = Replace(CommentBlock, "[CRC]", NmeaCrcChk(CommentBlock))
        DataOutput = "\" & CommentBlock & "\" & Data
    Else
        DataOutput = Data
    End If

'Data is nmea only - no comment block
    If frmRouter.MenuViewInoutData.Checked = True Then
        kbT = DataOutput & " [Dst:" & Sockets(Destination).DevName
        If Info <> "" Then
            kbT = kbT & "," & Info
        End If
        kbT = kbT & "]" & vbCrLf
'    Call frmRouter.TermOutput(Destination & ">" & Data & vbCrLf)
        Call frmRouter.TermOutput(kbT)
    End If
'    Sockets(Destination).Chrs = Sockets(Destination).Chrs + Len(DataOutput)
    Sockets(Destination).Chrs = Sockets(Destination).Chrs + 1
            
    Select Case Sockets(Destination).Handler
    Case Is = 0
        Call modWinsock.WinsockOutput(Hidx, DataOutput & vbCrLf)
    Case Is = 1
        Comms(Hidx).CommOutput DataOutput & vbCrLf
'            frmRouter.StatusBar.Panels(2).Text = Sockets(Destination).MsgCount _
'& ":" & Sockets(Destination).LostMsgCount & ":" & Len(.Buffer(.QFront))
    Case Is = 2
        If SysDate <> Files(Hidx).RollOverDate Then
kb = "Sysdate=" & SysDate & vbCrLf & "File date = " & Files(Hidx).RollOverDate & vbCrLf
        End If
        Files(Hidx).FileOutput DataOutput
    Case Is = 3
        Call TTYs(Hidx).TTYOutput(DataOutput & vbCrLf)
    Case Is = 4
'Disallow loopback to recorder !
        If Source <> 0 Then
            Call LoopBacks(Hidx).LoopBackOutput(DataOutput & vbCrLf, Source, Destination)
        End If
    Case Else
        MsgBox "Handler " & Sockets(Destination).Handler & " not found", , "Deque"
    End Select
End Function

'Gets first available Route for this Socket
'Returns -1 if none available
'Extends Routes array up to MAX_SOCKETROUTES
Public Function FreeRoute(Idx As Long) As Long
Dim Ridx As Long
    
    If RouteCount = MAX_ROUTES Then
        FreeRoute = -1
        Exit Function
    End If
    
    For Ridx = 1 To UBound(Sockets(Idx).Routes)
        If Sockets(Idx).Routes(Ridx).AndIdx <= 0 Then
            FreeRoute = Ridx
            Exit Function
        End If
    Next Ridx
    FreeRoute = Ridx
    
    If FreeRoute > UBound(Sockets(Idx).Routes) Then
'Extend the Route array for this socket
        ReDim Preserve Sockets(Idx).Routes(1 To FreeRoute)
    End If
End Function

'Returns No of Routes set up (even if disabled)
Public Function RouteCount() As Long
Dim Count As Long
Dim Idx As Long
Dim i As Long

    If SocketCount > 0 Then
    For Idx = 1 To UBound(Sockets)
'        If Sockets(Idx).State <> -1 Then
            For i = 1 To UBound(Sockets(Idx).Routes)
                If Sockets(Idx).Routes(i).AndIdx > 0 Then
                    Count = Count + 1
                End If
            Next i
'        End If
    Next Idx
    RouteCount = Count
    End If
End Function

'Gets first available Forward for this Socket
'Returns -1 if none available
'Extends Forwards array up to MAX_SOCKETFORWARDS
Public Function FreeForward(Source As Long) As Long
Dim i As Long
    For i = 1 To UBound(Sockets(Source).Forwards)
        If Sockets(Source).Forwards(i) = 0 Then
            FreeForward = i
            Exit Function
        End If
    Next i
'V44 removed all restriction of forwards, because it is not reall relevant to the user
'   If UBound(Sockets(Source).Forwards) = MAX_SOCKETFORWARDS Then
'no free forwards
'        FreeForward = -1
'    Else
'We can still allocate more forwards
        ReDim Preserve Sockets(Source).Forwards(1 To UBound(Sockets(Source).Forwards) + 1)
        FreeForward = UBound(Sockets(Source).Forwards)
'    End If
End Function

'Gets first available Stream Index for this Socket
'Returns -1 if none available
'Extends Streams array up to PermittedStreams for this Listening
'Socket
Public Function FreeStream(IdxListen As Long) As Long
Dim Count As Long
Dim i As Long
Dim IPCount As Long

    Count = StreamCount(IdxListen)
'Make the minimum streams 1, if first time
'otherwise confuses setup
    If Count = 0 And Sockets(IdxListen).Winsock.PermittedStreams < 1 Then
        Sockets(IdxListen).Winsock.PermittedStreams = 1
        Sockets(IdxListen).Winsock.PermittedIPStreams = 1
    End If

'Get first free stream
    For i = 1 To UBound(Sockets(IdxListen).Winsock.Streams)
'        If Sockets(IdxListen).Winsock.Streams(i) > 0 Then
'            Count = Count + 1
'        Else
        If Sockets(IdxListen).Winsock.Streams(i) <= 0 Then
            If FreeStream = 0 Then
                FreeStream = i
            End If
        End If
    Next i
        
'If no blank slots then the next one must be i
'which will be 1 past Ubound(streams)
    If FreeStream = 0 Then
        FreeStream = i
    End If
    
    If Count >= MAX_TCPSERVERSTREAMS Then
        FreeStream = -1
    Else
        If Count >= Sockets(IdxListen).Winsock.PermittedStreams Then
            WriteLog "No free Concurrent Clients limit is " & Sockets(IdxListen).Winsock.PermittedStreams
            FreeStream = -1
'        Else
'            IpCount = StreamIPCount(IdxListen)
'IpCount does not include the current connect were trying to make
'Debug.Print "FS:" & IpCount
'            If IpCount >= Sockets(IdxListen).Winsock.PermittedIPStreams Then
'                WriteLog "No free connections to same Client, limit is " & Sockets(IdxListen).Winsock.PermittedIPStreams
'                FreeStream = -1
'            End If
        End If
    End If

    If FreeStream <= 0 Then
        Exit Function
    End If
    
    If FreeStream > UBound(Sockets(IdxListen).Winsock.Streams) Then
'Extend the Route array for this socket
        ReDim Preserve Sockets(IdxListen).Winsock.Streams(1 To FreeStream)
    End If

End Function

Public Function StreamIPCount(IdxListen As Long) As Long
Dim Count As Long
Dim Sidx As Long
Dim HidxListen As Long
Dim HidxStream As Long
Dim IPListen As String
Dim PortListen As Long
Dim RemoteIPStream As String
Dim i As Long
Dim Hidx As Long
Dim ctrl As Winsock

'Call DisplayWinsock
    HidxListen = (Sockets(IdxListen).Hidx)
'frmrouter.winsock(hidxlisten).remotehostip returns a short IP address !!!
    IPListen = Sockets(IdxListen).Winsock.RemoteHostIP
    PortListen = Sockets(IdxListen).Winsock.LocalPort
    With frmRouter.Winsock(HidxListen)
        For Each ctrl In frmRouter.Winsock
            If ctrl.RemoteHostIP = IPListen _
            And ctrl.LocalPort = PortListen _
            And ctrl.State <> sckListening Then
                Count = Count + 1
            End If
If ctrl.Protocol = sckTCPProtocol Then
    Debug.Print "SC:" & ctrl.RemoteHostIP & ":" & ctrl.LocalPort & "-" & ctrl.State
'Stop
End If
        Next ctrl
'        For Hidx = 0 To UBound(frmRouter.Winsock)
'        Next Hidx
'    For i = 1 To UBound(Sockets(IdxListen).Winsock.Streams)
'        Sidx = Sockets(IdxListen).Winsock.Streams(i)
'        If Sidx > 0 Then
'            HidxStream = Sockets(Sidx).Hidx
'            RemoteIPListen = frmRouter.Winsock(HidxListen).RemoteHostIP
'            RemoteIPStream = frmRouter.Winsock(HidxStream).RemoteHostIP
'            If RemoteIPStream = RemoteIPListen Then
'                Count = Count + 1
'            End If
'        End If
'    Next i
    End With
    StreamIPCount = Count

End Function

Public Function StreamCount(IdxListen As Long) As Long
Dim Count As Long
Dim i As Long

    For i = 1 To UBound(Sockets(IdxListen).Winsock.Streams)
        If Sockets(IdxListen).Winsock.Streams(i) > 0 Then
            Count = Count + 1
        End If
    Next i
    StreamCount = Count
End Function


Public Function SocketCount() As Long
Dim i As Long
Dim Count As Long
    On Error GoTo noSockets
    For i = 1 To UBound(Sockets)
        If Sockets(i).State <> -1 Then
           Count = Count + 1
        End If
    Next i
noSockets:
    SocketCount = Count
End Function

Public Function TTYCount() As Long
Dim Count As Long
Dim f As Form

    For Each f In Forms
        If f.Name = "frmTTY" Then
            Count = Count + 1
        End If
    Next f
TTYCount = Count
End Function

'Removes any linked Routes and Forwards (even if enabled)
Public Function RemoveSocket(ReqIdx As Long)
Dim Idx1 As Long
Dim Idx2 As Long
Dim Ridx As Long
Dim Fidx As Long

'If the socket is not actually used then it should be cleared here
    WriteLog "Removing Socket " & ReqIdx
    Sockets(ReqIdx).Enabled = False
    
    Call Routecfg.RemoveSocketForwards(ReqIdx)
    
'Remove all Routes to/from this Socket
    For Idx1 = 1 To UBound(Sockets)
        If Idx1 = ReqIdx Then
        Else
        With Sockets(Idx1)
            For Ridx = 1 To UBound(.Routes)
                Idx2 = .Routes(Ridx).AndIdx
                If Idx2 > 0 Then
                    If Idx1 = ReqIdx Or Idx2 = ReqIdx Then
'Route and Forwards are removed even if enabled
                        Call Routecfg.RemoveRoute(Idx1, Idx2)
                    End If
                End If
            Next Ridx
        End With
        End If
    Next Idx1
    
    Call CloseHandler(ReqIdx)
    Call ClearSocket(ReqIdx)
End Function

'Allows user to select record on menu
Public Function IsRecordable(Idx As Long) As Boolean
'    If Sockets(Idx).Enabled Then
    
                        '0=both Input,Output
                        '1=Input Only
                        '2=Output Only
        Select Case Sockets(Idx).Direction
            Case Is = 0, 2
                If IsTcpListener(Idx) = False Then
'                    If Sockets(Idx).Recorder.Output Is Nothing Then
                        IsRecordable = True
'                     End If
                End If
        End Select
'    End If

End Function

'Only creates if enabled (enabled must be in sockets)
'Cannot be in Recorder because it has not been created at this point
Public Function CreateRecorder(Idx As Long)
    If Sockets(Idx).Recorder.Enabled Then
        If Sockets(Idx).Recorder.Output Is Nothing Then
            Set Sockets(Idx).Recorder.Output = New frmRecorder
        End If
        Sockets(Idx).Recorder.Output.ParentIdx = Idx
        Sockets(Idx).Recorder.Output.Caption = "Recorder [" & Sockets(Idx).DevName & "]"
        Sockets(Idx).Recorder.Output.Visible = True
    End If
End Function

Public Function RemoveRecorder(Optional Idx As Long)
    On Error Resume Next    'Socket may not exist
    If Idx = 0 Then
        For Idx = 1 To UBound(Sockets)
            If Not Sockets(Idx).Recorder.Output Is Nothing Then
                Unload Sockets(Idx).Recorder.Output
                Set Sockets(Idx).Recorder.Output = Nothing
            End If
        Next Idx
    Else
        If Not Sockets(Idx).Recorder.Output Is Nothing Then
            Unload Sockets(Idx).Recorder.Output
            Set Sockets(Idx).Recorder.Output = Nothing
        End If
    End If
End Function

'Returns True if ANY New Route can be set up
'Called By MenuRoute_click
Public Function IsNewRoute() As Boolean
Dim Idx1 As Long
Dim Idx2 As Long
Dim NewRouteCount As Long

'Must have minimum of 2 sockets, if not you cannot select
'Configuration > Route
    
    If RouteCount = MAX_ROUTES Then
        Exit Function
    End If
    
    If SocketCount > 0 Then
        For Idx1 = 1 To UBound(Sockets)
            For Idx2 = Idx1 + 1 To UBound(Sockets)
'If valid (Silent=do not display error message)
                If Routecfg.ValidateNewRoute(Idx1, Idx2) = True Then
                    IsNewRoute = True
                    Exit Function
                End If
            Next Idx2
        Next Idx1
    End If
End Function

Public Function IsTcpListener(Idx As Long) As Boolean
    With Sockets(Idx)
        If .Handler = 0 _
        And .Winsock.Protocol = sckTCPProtocol _
        And .Winsock.Server = 1 _
        And .Winsock.Oidx < 1 Then
            IsTcpListener = True
        End If
    End With
End Function

Public Function IsTcpStream(Idx As Long) As Boolean
    With Sockets(Idx)
        If .Handler = 0 _
        And .Winsock.Protocol = sckTCPProtocol _
        And .Winsock.Server = 1 _
        And .Winsock.Oidx >= 1 Then
            IsTcpStream = True
        End If
    End With
End Function


Public Function IsTcpClient(Idx As Long) As Boolean
    With Sockets(Idx)
        If .Handler = 0 _
        And .Winsock.Protocol = sckTCPProtocol _
        And .Winsock.Server = 0 Then
            IsTcpClient = True
        End If
    End With
End Function

Public Function IsTcpServer(Idx As Long) As Boolean
    With Sockets(Idx)
        If .Handler = 0 _
        And .Winsock.Protocol = sckTCPProtocol _
        And .Winsock.Server = 1 Then
            IsTcpServer = True
        End If
    End With
End Function

'This returns the Owner of the Route used for creating the forwards
Public Function cOidx(Idx As Long) As Long
'It is always the Socket, unless it is a TCP Stream
    cOidx = Idx
'If it is a TcpStream then it is the owner of the stream
    If Idx > 0 Then
        With Sockets(Idx)
            If .Handler = 0 _
            And .Winsock.Protocol = sckTCPProtocol _
            And .Winsock.Server = 1 _
            And .Winsock.Oidx > 0 Then
                cOidx = .Winsock.Oidx
            End If
        End With
    End If
'This Idx is then used to create the forwarding
End Function

'Returns True if there is ANY route that can be deleted
Public Function IsDeleteRoute() As Boolean

End Function
Public Function LongFileName(ByVal short_name As String) As _
    String
Dim pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        pos = 3
    Else
        result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If pos = 0 Then
'            long_name = Dir$(short_name, vbNormal + _
'                vbHidden + vbSystem + vbDirectory)
'a blank name (above returns a ".")
            long_name = ""
        Else
            long_name = Dir$(Left$(short_name, pos - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    LongFileName = result
End Function

'return true if they are the same
Public Function FileCompare(File1 As String, File2 As String) As Boolean
Dim Len1 As Long
Dim Ch1 As Long
Dim b1() As Byte
Dim Len2 As Long
Dim fNum As Long
Dim b2() As Byte
Dim i As Long
Dim kb1 As String
Dim kb2 As String
Dim j As Long
Dim LastLf As Long

    On Error GoTo Err_NoFile
    Len1 = FileLen(File1)
    Len2 = FileLen(File2)
    On Error GoTo 0
'    If Len1 <> Len2 Then   'speed up exit
'        Exit Function
'    End If
    fNum = FreeFile
    ReDim b1(1 To Len1)
    Open File1 For Binary As #fNum
    Get #fNum, 1, b1
    Close fNum
    fNum = FreeFile
    ReDim b2(1 To Len2)
    Open File2 For Binary As #fNum
    Get #fNum, 1, b2
    Close fNum
    For i = 1 To Len1
        If b1(i) = 10 Then LastLf = i
        If b1(i) <> b2(i) Then
 'display first changed line in status bar File1 is the new file
            i = LastLf + 1
            Do Until i > Len1
                If b1(i) = 13 Then Exit Do  'end of line
                kb1 = kb1 & Chr$(b1(i))
                i = i + 1
            Loop
'            NmeaRcv.StatusBar1.Panels(1) = kb1
            i = LastLf + 1
            Do Until i > Len2
                If b2(i) = 13 Then Exit Do  'end of line
                kb2 = kb2 & Chr$(b2(i))
                i = i + 1
            Loop
'MsgBox File2 & vbCrLf & kb2 & vbCrLf & File1 & vbCrLf & kb1
            Exit Function
        End If
    Next i
    FileCompare = True
    Exit Function
Err_NoFile:

End Function

'Creates the Handler from details in Sockets(Idx)
'If successful updates state of Sockets(Hidx) with Handler(Hidx)
Public Function CreateHandler_old(Idx As Long) As String
Dim msg As String

    CurrentSocket = Idx
    
'Stop any forwarding on this socket because the Socket handler will
'be closed before -t is re-opened
'    If Sockets(Idx).Hidx > 0 Then
'Also stop any handler timers (Comms,File input) first
'        msg = msg & Routecfg.SocketForwards(Idx, "Remove")
'    End If
    
    
    If Sockets(Idx).Enabled = False Then
        Sockets(Idx).State = 0
        WriteLog "Socket " & Sockets(Idx).DevName & " is Disabled"
        Exit Function
    End If
'Clear any previous err message
    Sockets(Idx).errmsg = ""
    
'    Call OpenHandler(Idx)

#If True Then
        
    Sockets(Idx).State = 11 'Trying to open handler
        Select Case Sockets(Idx).Handler
'Sockets().Hidx is set by the handler (-1) if cant open
        Case Is = 0            '0 = Winsock
            Call CreateWinsock(CInt(Idx))
            If Sockets(Idx).Hidx > 0 Then
                Sockets(Idx).State = frmRouter.Winsock(Sockets(Idx).Hidx).State
            Else
                Sockets(Idx).State = 0  'Closed
            End If
        Case Is = 1           '1 = COMM
'Hidx needs to be -1 to create a new Comm or Open and existing one
            Call CreateComm(Idx)
            If Sockets(Idx).Hidx > 0 Then
                Sockets(Idx).State = Comms(Sockets(Idx).Hidx).State
            Else
                Sockets(Idx).State = 0  'Closed
            End If
        Case Is = 2            '2 = File
'            MsgBox "File handler not yet implemented", , "Close Profile"
'Create an object so that we can call CreateFile
            Call CreateFile(Idx)
            If Sockets(Idx).Hidx > 0 Then
                Sockets(Idx).State = Files(Sockets(Idx).Hidx).State
            Else
'If closed in error leave the state as an error
'                Sockets(Idx).State = 0  'Closed
            End If
        Case Is = 3            '3 = TTY
           Call CreateTTY(Idx)
'Hidx is left as 11 if trying to open TTYs
            If Sockets(Idx).Hidx > 0 Then
                Sockets(Idx).State = TTYs(Sockets(Idx).Hidx).State
            End If
        Case Is = 4            '4 = LoopBack
           Call CreateLoopBack(Idx)
            If Sockets(Idx).Hidx > 0 Then
                Sockets(Idx).State = LoopBacks(Sockets(Idx).Hidx).State
            End If
        Case Else
            MsgBox "Handler " & Sockets(Idx).Handler & " not found", , "Close Profile"
        End Select
    
#End If

End Function
Function UTCDateNow() As Long
Dim Systime As SYSTEMTIME
    Call GetSystemTime(Systime)
    With Systime
        UTCDateNow = CLng(.wYear) * 10000 + CLng(.wMonth) * 100 + CLng(.wDay)
    End With
End Function

Function UTCTimeNow() As Long
Dim Systime As SYSTEMTIME
    Call GetSystemTime(Systime)
    With Systime
        UTCTimeNow = .wHour * 100 + .wMinute
    End With
End Function

'V59
Public Function SysTimeToUnix(Systime As SYSTEMTIME) As Long
Dim DinY()
Dim Dayno As Long
Dim ret As Long
Dim kb As String

    With Systime
        SysTimeToUnix = DateDiff("s", DateSerial(1970, 1, 1), DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond))
    End With

End Function

Function UnixNow() As Long
Dim Systime As SYSTEMTIME

    Call GetSystemTime(Systime)
    UnixNow = SysTimeToUnix(Systime)
    
#If False Then  'v59
Dim DinY()
Dim Dayno As Long
Dim ret As Long
Dim kb As String
    If Dayno = 0 Then
        DinY = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
    End If
    
    With Systime
'.wYear = 2000
'.wMonth = 5
'.wDay = 10
'.wHour = 0
'.wMinute = 0
'.wSecond = 0
        Dayno = 365 * (.wYear - 1970) + Int(((.wYear - 1970) + 2) / 4) + DinY(.wMonth - 1) + .wDay - 1
        If (.wYear = (Int(.wYear / 4) * 4) _
        And .wYear And .wMonth > 1) _
        Then Dayno = Dayno + 1
'2000 is NOT a leap year
        If .wYear >= 2000 And .wMonth > 2 _
        Then Dayno = Dayno - 1
        UnixNow = CLng(Dayno) * 86400 + CLng(.wHour) * 3600 + CLng(.wMinute) * 60 + CLng(.wSecond)
    End With
#End If
End Function

Public Function BooleanA(kb As String) As Boolean
'kb = "Vero"
    If kb = CStr(True) Then BooleanA = True
'initial files distributed are in english
'and may not have been used/converted
'check if works in German
'If kb = "Wahr" Then booleana = True
'If kb = "Vero" Then booleana = True
    If IsNumeric(kb) Then
        If CBool(kb) = True Then BooleanA = True
    End If
End Function

'Public Function BooleanA(Text As String) As Boolean
'    If Text = "True" Then BooleanA = True
'End Function

Public Function BooleanChecked(Checked As Integer) As Boolean
    If Checked = vbChecked Then
        BooleanChecked = True
    End If
End Function

Public Function SysDate() As String
Dim Systime As SYSTEMTIME
    
    Call GetSystemTime(Systime)
    SysDate = Systime.wYear & Format$(Systime.wMonth, "00") & Format$(Systime.wDay, "00")
End Function


Public Function HttpSpawn(Url As String)
Dim r As Long
Dim Command As String

If Environ("windir") <> "" Then
    r = ShellExecute(0, "open", Url, 0, 0, 1)
Else
'try for linux compatibility
    Command = "winebrowser " & Url & " ""%1"""

    Shell (Command)
End If
End Function
'Only used for debugging input
Public Function WriteInputLog(kb As String, Optional Idx As Long)
    If InputLogCh = -1 Then  'First time its opened
        InputLogCh = FreeFile
        InputLogName = TempPath & "NmeaRouter_Input.log"
        Open InputLogName For Output As #InputLogCh
    End If
    If InputLogCh = 0 Then   'Re-open in append
        InputLogCh = FreeFile
        Open InputLogName For Append As #InputLogCh
    End If
    
    Print #InputLogCh, Sockets(Idx).DevName & vbTab & kb
End Function

Public Function AddFileToLog(FileName As String)
Dim ch As Long
Dim nextline As String

    ch = FreeFile
    On Error GoTo nofil
    Open FileName For Input As #ch
    Do Until EOF(ch)
        Line Input #ch, nextline
        Call WriteLog(nextline)
    Loop
    Close #ch
'com0com log adds incremantally to the con0com log file
'so we need to delete it when we transfer the details to my event log
    Kill FileName
nofil:
End Function

Public Function CloseLog(Optional Flush As Boolean)
Dim Idx As Long

    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
            If .TryCount > 0 Then
                WriteLog .TryCount & " Tries Opening " & .DevName
            End If
        End With
    Next Idx

    Close #LogFileCh
    If Flush Then
        Open LogFileName For Append As #LogFileCh
    Else
        LogFileCh = 0
    End If
End Function

Public Function ErrorDescriptionDLL(Optional ByVal lLastDLLError As Long) As String
    Dim sBuff As String * 256
    Dim lCount As Long
    Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
    Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
    Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

    If lLastDLLError = 0 Then
        'Use Err object to get dll error number
        lLastDLLError = err.LastDllError
    End If

    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        ErrorDescriptionDLL = Left$(sBuff, lCount - 2)    'Remove line feeds
    End If
    
End Function

Public Sub DisplayTries(Caller As String, Optional ReqIdx As Long)
Dim kb As String
Dim Idx As Long

    For Idx = 1 To UBound(Sockets)
        If ReqIdx = 0 Or Idx = ReqIdx Or Sockets(Idx).TryCount <> 0 Then
            kb = kb & Sockets(Idx).DevName & " (" & aState(Sockets(Idx).State) & "), Tries=" & Sockets(Idx).TryCount & vbCrLf
        End If
    Next Idx
    MsgBox kb, , Caller & " " & ReqIdx
End Sub


Public Sub DisplayWindow(Caller As String)
Dim kb As String
    
'    If frmRouter.Visible = True Then
'        If ResizeOk Then
'            kb = "In GUI"
'        Else
'            kb = "Restore to GUI"
'        End If
'    Else
'        If ResizeOk Then
'            kb = "Send to sysTray"
'        Else
'            kb = "In sysTray"
'When in sysTray all dimensions are incorrect
'        End If
'    End If
'Exit Sub
MsgBox frmRouter.Name & ", L=" & frmRouter.Left & ", T=" & frmRouter.Top & ", W=" & frmRouter.Width & ", H=" & frmRouter.Height & ", S=" & frmRouter.WindowState & ", V=" & frmRouter.Visible & vbCrLf _
& "ScaleHeight=" & frmRouter.ScaleHeight & ", ScaleWidth=" & frmRouter.ScaleWidth & vbCrLf _
& "Visible=" & frmRouter.Visible & vbCrLf _
& "State=" & frmRouter.WindowState & vbCrLf _
& kb & vbCrLf _
, , Caller

End Sub

Public Function AllForwardsCount() As Long
Dim Idx As Long
Dim Fidx As Long
Dim Count As Long
    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
            For Fidx = 1 To UBound(.Forwards)
                If .Forwards(Fidx) > 0 Then
                    Count = Count + 1
                End If
            Next Fidx
        End With
    Next Idx
    AllForwardsCount = Count
End Function

Public Sub IncrForward(Source As Long, Destination As Long, Reverse As Boolean)
Dim cOIdx1 As Long
Dim cOidx2 As Long
Dim Ridx As Long
    
'1. Get the Route Index used by this Forward
'Convert the Idx to the Owner because the Route is set up on the
'Owner if a TCP stream socket. If not the Owner will be the same
'as the Passed Idx (Source or Destination)
    cOIdx1 = cOidx(Source)
    cOidx2 = cOidx(Destination)
    Ridx = Routecfg.RouteExists(cOIdx1, cOidx2) 'Passed by ref
'The route should always exist - but check anyway
    If Ridx > 0 Then
        If Reverse = False Then
            Sockets(cOIdx1).Routes(Ridx).ForwardCount = _
            Sockets(cOIdx1).Routes(Ridx).ForwardCount + 1
        Else
            Sockets(cOIdx1).Routes(Ridx).ReverseCount = _
            Sockets(cOIdx1).Routes(Ridx).ReverseCount + 1
        End If
    End If
End Sub

Public Sub DecrForward(Source As Long, Destination As Long, Optional Reverse As Boolean)
Dim cOIdx1 As Long
Dim cOidx2 As Long
Dim Ridx As Long
    
    cOIdx1 = cOidx(Source)
    cOidx2 = cOidx(Destination)
    Ridx = Routecfg.RouteExists(cOIdx1, cOidx2) 'Passed by ref
    If Ridx > 0 Then
        With Sockets(cOIdx1).Routes(Ridx)
            If Reverse = False Then
'note can be called by Remove SocketForwards when there are NO forwards
                If .ForwardCount > 0 Then .ForwardCount = .ForwardCount - 1
            Else
                If .ReverseCount > 0 Then .ReverseCount = .ReverseCount - 1
            End If
        End With
    End If

End Sub

Public Sub MakeFormsVisisble()
Dim i As Long

    For i = 0 To Forms.Count - 1
        Select Case Forms(i).Name
        Case Is = "frmTTY", "frmRecorder"
            Forms(i).Visible = True
        End Select
    Next i
End Sub

'Check if an AisSentence Payload bits and fill bits OK
'Returns False if not AIS sentence or Payload appears OK for the Message type
Private Function PayloadError(Data As String) As Boolean
'Split sentence into words (Check right no of words)
'If AIS sentence
'Select case "1st Payload character"
'If Fixed length, check Payload Bits
Stop
End Function

'Check if Output sentence is !**VDM, used by OnlyVDM filter
Private Function IsAivdm(Data As String) As Boolean
Dim j As Long
    On Error Resume Next
'instr can return null
    j = InStr(1, Data, "!")
    If j > 0 Then
        If InStr(j + 3, Data, "VDM,") > 0 Then
            IsAivdm = True
        End If
    End If
    On Error GoTo 0
'If IsAivdm = False Then Stop
End Function

#If False Then
'Socket dependant filter
'Cannot be in Recorder because it has not been created at this point
Public Function CreateFilterDuplicate(Idx As Long)
        If Sockets(Idx).DuplicateFilter Is Nothing Then
            Set Sockets(Idx).DuplicateFilter = New clsDuplicateFilter
        End If
        Sockets(Idx).DuplicateFilter.ParentIdx = Idx
End Function

Public Function RemoveFilterDuplicate(Optional Idx As Long)
    On Error Resume Next    'Socket may not exist
    If Idx = 0 Then
        For Idx = 1 To UBound(Sockets)
            If Not Sockets(Idx).DuplicateFilter Is Nothing Then
                Unload Sockets(Idx).DuplicateFilter
                Set Sockets(Idx).DuplicateFilter = Nothing
            End If
        Next Idx
    Else
        If Not Sockets(Idx).Recorder.Output Is Nothing Then
            Unload Sockets(Idx).Recorder.Output
            Set Sockets(Idx).Recorder.Output = Nothing
        End If
    End If
End Function

#End If


'Closes a handler
'Clears all associated memory
'If Idx =0 Closes all handlers
Public Sub CloseHandler(ReqIdx As Long)
Dim Idx As Long
Dim Hidx As Long
Dim kb As String
'MsgBox "CloseHandler " & ReqIdx

    For Idx = 1 To UBound(Sockets)
        If Idx = ReqIdx Or ReqIdx = 0 Then
            Sockets(Idx).errmsg = ""
            Hidx = Sockets(Idx).Hidx
            If Sockets(Idx).State <> -1 Then
                If Hidx > 0 Then
'Handler has been allocated

'Dont remove the Forwards - but dont re-create when OpenHandler is called
'                    kb = Routecfg.SocketForwards(ReqIdx, "Remove")
'If jnasetup = True Then
'    MsgBox kb, , "CloseHandler-Remove (" & Idx & ")"
'End If
                    Select Case Sockets(Idx).Handler
                    Case Is = 0            '0 = Winsock
                        Call CloseWinsock(CInt(Sockets(Idx).Hidx))
                    Case Is = 1           '1 = COMM
                        Set Comms(Sockets(Idx).Hidx) = Nothing
                    Case Is = 2            '2 = File
                        Call Files(Sockets(Idx).Hidx).CloseFile
'Added v26 (same as when profile is closed)
                        Set Files(Sockets(Idx).Hidx) = Nothing
                    Case Is = 3            '3 = TTY
                        Call TTYs(Sockets(Idx).Hidx).TTYClose
                    Case Is = 4            '4 = LoopBack
                        Set LoopBacks(Sockets(Idx).Hidx) = Nothing
                    Case Else
                        MsgBox "Handler " & Sockets(Idx).Handler & " not found", , "Close Profile"
                    End Select
                                        
                    WriteLog "Closed " & Sockets(Idx).DevName _
                    & " " & aHandler(Sockets(Idx).Handler) _
                    & " handler"
'Remove link to handler
                    Sockets(Idx).Hidx = 0
                End If
                    
'set state before ResetTries
                Sockets(Idx).State = sckClosed
                If Sockets(Idx).TryCount > 0 Then
                    Call frmRouter.ResetTries(Idx)
                End If
            End If
'Reset Clears all values from Sockets() so values
'must be recreated (as New) or re-loaded from Registry
            If ReqIdx = 0 Then
                Call ClearSocket(Idx)
            End If
        End If

    Next Idx
Exit Sub

Reset_error:
    Select Case err.Number
    Case Is = 10
        Exit Sub
    Case Else
MsgBox "Error " & err.Number & " - " & err.Description, , "CloseHandler"
    End Select
'Stop
End Sub

'Called by cmdStart, ReconnectTimer, ToggleSocketEnabled
' and ConnectionRequest
'Opens the handler for a socket if it is enabled
Public Sub OpenHandler(Idx As Long)
Dim kb As String
    If Sockets(Idx).Enabled = False Then
        Sockets(Idx).State = 0
        Exit Sub
    End If

'When a call is made to Open a handler Incr the TryCount
    Sockets(Idx).TryCount = Sockets(Idx).TryCount + 1

'Only log if not a re-open by
    Select Case Sockets(Idx).TryCount
    Case Is = 1
        WriteLog "Opening Handler for " & Sockets(Idx).DevName _
        & ", Current state is " & aState(Sockets(Idx).State)
    Case Is = 2
        WriteLog "Trying to open " & Sockets(Idx).DevName _
        & ", Current state is " & aState(Sockets(Idx).State)
    Case Else       'Supress if Try
'        WriteLog "to open " & Sockets(Idx).DevName _
'        & ", Current state is " _
'        & aState(Sockets(Idx).State), Idx
    End Select


'    Call DisplayTries("OpenHandler - Enter", Idx)
'State must be set to 11 to prevent File reading first record until timer is started
    Sockets(Idx).State = 11 'Trying to open handler
    
    Select Case Sockets(Idx).Handler
'Sockets().Hidx is set by the handler (-1) if cant open
    Case Is = 0            '0 = Winsock
        Call CreateWinsock(CInt(Idx))
        If Sockets(Idx).Hidx > 0 Then
            Sockets(Idx).State = frmRouter.Winsock(Sockets(Idx).Hidx).State
            If Sockets(Idx).State = sckListening Then
'                Sockets(Idx).Winsock.Sidx = 0
            End If
        Else
            Sockets(Idx).State = 0  'Closed
        End If
    Case Is = 1           '1 = COMM
'Hidx needs to be -1 to create a new Comm or Open and existing one
        Call CreateComm(Idx)
        If Sockets(Idx).Hidx > 0 Then
            Sockets(Idx).State = Comms(Sockets(Idx).Hidx).State
'        Else
'            Sockets(Idx).State = 0  'Closed
        End If
    Case Is = 2            '2 = File
'            MsgBox "File handler not yet implemented", , "Close Profile"
        Call CreateFile(Idx)
        If Sockets(Idx).Hidx > 0 Then
            Sockets(Idx).State = Files(Sockets(Idx).Hidx).State
        Else
'Sockets(idx).state should be set in create file
'            Sockets(Idx).State = 0  'Closed
        End If
    Case Is = 3            '3 = TTY
        Call CreateTTY(Idx)
'Hidx is left as 11 if trying to open TTYs
        If Sockets(Idx).Hidx > 0 Then
            Sockets(Idx).State = TTYs(Sockets(Idx).Hidx).State
        End If
    Case Is = 4            '4 = LoopBack
        Call CreateLoopBack(Idx)
        If Sockets(Idx).Hidx > 0 Then
            Sockets(Idx).State = LoopBacks(Sockets(Idx).Hidx).State
        End If
    Case Else
        MsgBox "Handler " & Sockets(Idx).Handler & " not found", , "Close Profile"
    End Select
    
    WriteLog Sockets(Idx).DevName & " " _
    & aHandler(Sockets(Idx).Handler) _
    & " is " & aState(Sockets(Idx).State)

'If at the end of the Open Handler call the when the Sockets(Idx).Status
'is 1 (Open) decr the Try Count
'MsgBox WinsockState(Idx)
    Select Case Sockets(Idx).State
    Case Is = sckOpen, sckListening, sckConnected
        Sockets(Idx).TryCount = Sockets(Idx).TryCount - 1
        If Sockets(Idx).TryCount > 0 Then
            Call frmRouter.ResetTries(Idx)
        End If
    End Select
    
'end of trying to open handler
End Sub

'Loads a socket from Registry into memory
'If Idx = 0 Loads all sockets from registry
Public Sub LoadSocket(Idx As Long)
MsgBox "LoadSocket-needs writing"
End Sub

'Check if we have complete setup data
'That is required to create the handler
Public Function Socket_Validate(Idx As Long) As Boolean
        Select Case Sockets(Idx).Handler
        Case Is = 0            '0 = Winsock
            Select Case Sockets(Idx).Winsock.Protocol
            Case Is = sckTCPProtocol
                Select Case Sockets(Idx).Winsock.Server
                Case Is = 0      '0=Client
                    If Sockets(Idx).Winsock.RemotePort = 0 Then Exit Function
                Case Is = 1        '1=Server
                    If Sockets(Idx).Winsock.LocalPort = 0 Then Exit Function
                Case Else
                    Exit Function
                End Select
            Case Is = sckUDPProtocol
                Select Case Sockets(Idx).Direction
                Case Is = 0 'both
                    If Sockets(Idx).Winsock.LocalPort = "" Then Exit Function
                    If Sockets(Idx).Winsock.RemotePort = "" Then Exit Function
                Case Is = 1 'input
                    If Sockets(Idx).Winsock.LocalPort = "" Then Exit Function
                Case Is = 2 'Output
                    If Sockets(Idx).Winsock.RemotePort = "" Then Exit Function
                Case Else
                    Exit Function
                End Select
            End Select
        Case Is = 1           '1 = COMM
            If Sockets(Idx).Comm.Name = "" Then Exit Function
        Case Is = 2            '2 = File
            If Sockets(Idx).File.SocketFileName = "" Then Exit Function
        Case Is = 3            '3 = TTY
'Nothing to set up (but check no modal forms before enabling TTY)
        Case Is = 4            '4 = LoopBack
'Nothing to set up
        Case Else
        End Select
    Socket_Validate = True
End Function




