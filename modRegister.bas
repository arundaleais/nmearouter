Attribute VB_Name = "modRegister"
Option Explicit

Public Const Encoding = "0123456789ABCDEF" 'GHJKLMNPQRSTUVWXYZ"
Private Const SCODE_KEY = 219048765
Private Const ACODE_KEY = 740030213
'
#If False Then
'---------------------------------------------------------------------------
' Used to get the MAC address.
'---------------------------------------------------------------------------
'
Private Const NCBNAMSZ As Long = 16
Private Const NCBENUM As Long = &H37
Private Const NCBRESET As Long = &H32
Private Const NCBASTAT As Long = &H33
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4

Private Type NET_CONTROL_BLOCK  'NCB
    ncb_command    As Byte
    ncb_retcode    As Byte
    ncb_lsn        As Byte
    ncb_num        As Byte
    ncb_buffer     As Long
    ncb_length     As Integer
    ncb_callname   As String * NCBNAMSZ
    ncb_name       As String * NCBNAMSZ
    ncb_rto        As Byte
    ncb_sto        As Byte
    ncb_post       As Long
    ncb_lana_num   As Byte
    ncb_cmd_cplt   As Byte
    ncb_reserve(9) As Byte 'Reserved, must be 0
    ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major          As Byte
    reserved0          As Byte
    adapter_type       As Byte
    rev_minor          As Byte
    duration           As Integer
    frmr_recv          As Integer
    frmr_xmit          As Integer
    iframe_recv_err    As Integer
    xmit_aborts        As Integer
    xmit_success       As Long
    recv_success       As Long
    iframe_xmit_err    As Integer
    recv_buff_unavail  As Integer
    t1_timeouts        As Integer
    ti_timeouts        As Integer
    Reserved1          As Long
    free_ncbs          As Integer
    max_cfg_ncbs       As Integer
    max_ncbs           As Integer
    xmit_buf_unavail   As Integer
    max_dgram_size     As Integer
    pending_sess       As Integer
    max_cfg_sess       As Integer
    max_sess           As Integer
    max_sess_pkt_size  As Integer
    name_count         As Integer
End Type

Private Type NAME_BUFFER
    name_(0 To NCBNAMSZ - 1) As Byte
    name_num                 As Byte
    name_flags               As Byte
End Type

Private Type ASTAT
    adapt             As ADAPTER_STATUS
    NameBuff(0 To 29) As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "netapi32" _
        (pncb As NET_CONTROL_BLOCK) As Byte

Private Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" (hpvDest As Any, ByVal _
        hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        lpMem As Any) As Long
'===================================
#End If

'Functions required for sysinfo
Private Declare Function GetLocaleInfo Lib "kernel32" _
Alias "GetLocaleInfoA" _
(ByVal Locale As Long, _
ByVal LCType As Long, _
ByVal lpLCData As String, _
ByVal cchData As Long) As Long
 
Private Const LOCALE_SDECIMAL = &HE
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
'http://vbcity.com/forums/p/99944/422558.aspx
Private Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'===================================
'These are for detecting 64 bit processor
Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
    
Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As Long
    
Private Declare Function GetCurrentProcess Lib "kernel32" _
    () As Long

Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long
'=====================================
Public Function CheckKey(ByVal ActivationKey As String) As Long
Dim X As Long
Dim Y As Long
Dim Key() As String
    If ActivationKey = "" Then
        CheckKey = -1    'Not registered
        Exit Function
    End If
    Key = Split(ActivationKey, "-")
    X = GetSerialNo()
    If X <= 0 Then      'cant get drive serial no
        CheckKey = -2
        Exit Function
    End If
    Y = X + 53 / X + 113 * (X / 4)
    If UBound(Key) > 0 And Y = ConvertBaseTo10(Key(0)) Then
        CheckKey = ConvertBaseTo10(Key(1)) / 479 '479 is prime no
    Else
        CheckKey = -1   'failed
    End If

End Function


Public Function ConvertBaseTo10(OldBase As String) As Long
Dim i As Long
Dim sum As Long
Dim Chrno As Long
Dim Chr As String
    For i = 1 To Len(OldBase)
        Chr = Mid$(OldBase, i, 1)
        Chrno = InStr(1, Encoding, Chr) - 1
        If Chrno < 0 Then   'Invalid character
            ConvertBaseTo10 = -1
            Exit Function
        End If
        sum = sum * Len(Encoding) + CInt(Chrno)
    Next i
ConvertBaseTo10 = sum
End Function


Public Function GetSerialNo() As Long
Dim HexString As String
Dim i As Long
Dim SerialNo As Long
Dim Chr As String
    
'Returns 0 if cant get the drive serial no
'HexString will be ""
    HexString = GetDriveSerialNo    '8 characters
#If False Then
    HexString = GetMacAddress       '12 characters
#End If
'Convert HEX into a serial number approx up to 10000k
    If Len(HexString) = 8 Then
        HexString = Right$(HexString, 2) & HexString
    End If
    For i = 1 To Len(HexString)
        Chr = Mid$(HexString, i, 1)
        If IsNumeric(Chr) Then
            SerialNo = SerialNo + CInt(Chr) * 2 ^ i
        Else
            SerialNo = SerialNo + (Asc(Chr) - 54) * 2 ^ i
        End If
    Next i
       
'    SerialNo = 22516 'test only
    
'To Generate an Activation key if JNA has entered a serial number
'supplied by the user
    If IsNumeric(frmRegister.txtSerialNo.Text) Then
        GetSerialNo = frmRegister.txtSerialNo.Text
    Else
        GetSerialNo = SerialNo
    End If
'    ActivationKey = GenerateKey(14030)
'    ActivationKey = frmRegister.GenerateKey(SerialNo)
'    NewBaseKey = Convert10ToBase(ActivationKey)
'    kb = ConvertBaseTo10(NewBaseKey)
'Debug.Print ActivationKey & ":" & NewBaseKey & ":" & kb
End Function

'Returns a hex string
Private Function GetDriveSerialNo() As String   'Hex
Dim fso As New FileSystemObject
Dim objDrive
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set objDrive = fso.GetDrive(Environ$("HOMEDRIVE"))
GetDriveSerialNo = Hex(objDrive.SerialNumber)
End Function

#If False Then
Private Function GetMacAddress() As String
    Dim l As Long
    Dim lngError As Long
    Dim lngSize As Long
    Dim pAdapt As Long
    Dim pAddrStr As Long
    Dim pASTAT As Long
    Dim strTemp As String
    Dim strAddress As String
    Dim strMACAddress As String
    Dim AST As ASTAT
    Dim NCB As NET_CONTROL_BLOCK
'
    '---------------------------------------------------------------------------
    ' Get the network interface card's MAC address.
    '----------------------------------------------------------------------------
    '
    On Error GoTo errorhandler
    GetMacAddress = ""
    strMACAddress = ""

    '
    ' Try to get MAC address from NetBios. Requires NetBios installed.
    '
    ' Supported on 95, 98, ME, NT, 2K, XP
    '
    ' Results Connected Disconnected
    ' ------- --------- ------------
    '   XP       OK         Fail (Fail after reboot)
    '   NT       OK         OK   (OK after reboot)
    '   98       OK         OK   (OK after reboot)
    '   95       OK         OK   (OK after reboot)
    '
    NCB.ncb_command = NCBRESET
    Call Netbios(NCB)

    NCB.ncb_callname = "*               "
    NCB.ncb_command = NCBASTAT
    NCB.ncb_lana_num = 0
    NCB.ncb_length = Len(AST)

    pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or _
                       HEAP_ZERO_MEMORY, NCB.ncb_length)
    If pASTAT = 0 Then GoTo errorhandler

    NCB.ncb_buffer = pASTAT
    Call Netbios(NCB)

    Call CopyMemory(AST, NCB.ncb_buffer, Len(AST))

'Last 5 hex characters, otherwise key is too big
    strMACAddress = Right$("00" & Hex(AST.adapt.adapter_address(0)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(1)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(2)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(3)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(4)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(5)), 2)

    Call HeapFree(GetProcessHeap(), 0, pASTAT)

    GetMacAddress = strMACAddress
    
    
    GoTo NormalExit

errorhandler:
    Call MsgBox(err.Description, vbCritical, "Error")

NormalExit:
    End Function
#End If

Public Function sysinfo() As String
Dim kb As String
Dim i As Long
Dim ch As Long
Dim nextline As String

kb = ""
    kb = vbCrLf & "Environment Settings" & vbCrLf
'    i = 1
'    Do
'        kb = kb & vbTab & Environ(i) & vbCrLf
'        i = i + 1
'    Loop Until Environ(i) = ""
    
'There is no environment variable to get the ALL USERS\Application Data location
'internationally Application data is re-named si the english must NOT be used

'    kb = kb & vbTab & "GetSpecialFolder CSIDL_COMMON_APPDATA = " & GetSpecialFolderA(CSIDL_COMMON_APPDATA) & vbCrLf
    kb = kb & vbTab & "Windows Version = " & GetVersion1() & vbCrLf
    kb = kb & vbTab & "User Default LocaleID = " & GetUserDefaultLCID() & vbCrLf
    kb = kb & vbTab & "User has Administrator Rights = " & CBool(IsNTAdmin(ByVal 0&, ByVal 0&)) & vbCrLf
'now the log file
'Debug.Print Len(kb)

Dim arProfiles() As Variant
    kb = kb & "PROFILES " & ReadKeys(HKEY_CURRENT_USER, ROUTERKEY & "\" & "Profiles", arProfiles) & "(" & MAX_PROFILES & ")" & vbCrLf
    kb = kb & "FILES " & UBound(Files) & "(" & MAX_FILES & ")" & vbCrLf
    kb = kb & "TTYS " & FreeTTY & "(" & MAX_TTYS & ")" & vbCrLf
    kb = kb & "COMMS " & FreeComm & "(" & MAX_COMMS & ")" & vbCrLf
    kb = kb & "SOCKETS " & SocketCount & "(" & MAX_SOCKETS & ")" & vbCrLf
    kb = kb & "WINSOCKS " & WinsockCount & "(" & MAX_WINSOCKS & ")" & vbCrLf
    kb = kb & "All-TCPSERVERSTREAMS " & frmRouter.AllStreamCount & "(" & MAX_TCPSERVERSTREAMS & ")" & vbCrLf
    kb = kb & "ROUTES " & RouteCount & "(" & MAX_ROUTES & ")" & vbCrLf
    kb = kb & "SOCKETROUTES " & " socket dependant " & "(" & MAX_SOCKETROUTES & ")" & vbCrLf
    kb = kb & "LOOPBACKS " & FreeLoopBack & "(" & MAX_LOOPBACKS & ")" & vbCrLf
    kb = kb & "All-SOCKETFORWARDS " & AllForwardsCount & "(" & MAX_SOCKETFORWARDS & ")" & vbCrLf
        

#If False Then
    Call CloseLog
    kb = kb & "Print of: " & LogFileName & vbCrLf
    ch = FreeFile
    On Error GoTo nofil
    Open LogFileName For Input As #ch
    Do Until EOF(ch)
        Line Input #ch, nextline
        kb = kb & vbTab & nextline & vbCrLf
    Loop
    Close #ch
    kb = kb & "Finished print of: " & LogFileName & vbCrLf
'Debug.Print Len(kb)
'    sysinfo = Replace(kb, vbCrLf, "\n")
    sysinfo = kb
'Debug.Print Len(sysinfo)
    Exit Function
nofil:
    On Error GoTo 0
    Close #ch
    kb = kb & "File not found " & LogFileName & vbCrLf
    sysinfo = Replace(kb, vbCrLf, "<br>")
#End If
    sysinfo = kb
'    sysinfo = Replace(kb, vbCrLf, "<br>")
End Function

Public Function GetVersion1() As String
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    With osinfo
        Select Case .dwPlatformId
        Case 1
            Select Case .dwMinorVersion
            Case 0
                GetVersion1 = "Windows 95"
            Case 10
                GetVersion1 = "Windows 98"
            Case 90
                GetVersion1 = "Windows Millenium"
            Case Else
                GetVersion1 = "Unknown"
            End Select
        Case 2
            Select Case .dwMajorVersion
            Case 3
                GetVersion1 = "Windows NT 3.51"
            Case 4
                    GetVersion1 = "Windows NT 4.0"
            Case 5
                Select Case .dwMinorVersion
                Case 0
                    GetVersion1 = "Windows 2000"
                Case 1
                    GetVersion1 = "Windows XP"
                Case 2
                    GetVersion1 = "Windows Server 2003"
                Case Else
                    GetVersion1 = "Unknown"
                End Select
            Case 6
                Select Case .dwMinorVersion
                Case 0
                    GetVersion1 = "Windows Vista"
                Case 1
                    GetVersion1 = "Windows 7"
                Case 2
                    GetVersion1 = "Windows 8"
                Case Else
                    GetVersion1 = "Unknown"
                End Select
            Case Else
                GetVersion1 = "Unknown"
            End Select
        Case Else
            GetVersion1 = "Unknown"
        End Select
    End With
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=9043
Public Function Is64bit() As Boolean
    Dim handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function
