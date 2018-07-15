Attribute VB_Name = "mPSAPI"
'http://www.vbaccelerator.com/home/VB/Tips/Getting_Process_Information_Using_PSAPI/VB_PSAPI_Demonstration.asp
Option Explicit

'typedef struct _MODULEINFO {    LPVOID lpBaseOfDll;    DWORD SizeOfImage;
'    LPVOID EntryPoint;} MODULEINFO, *LPMODULEINFO;
Type MODULEINFO
   lpBaseOfDLL As Long
   SizeOfImage As Long
   EntryPoint As Long
End Type
'typedef struct _PROCESS_MEMORY_COUNTERS {
'    DWORD cb;
'    DWORD PageFaultCount;
'    DWORD PeakWorkingSetSize;
'    DWORD WorkingSetSize;
'    DWORD QuotaPeakPagedPoolUsage;
'    DWORD QuotaPagedPoolUsage;
'    DWORD QuotaPeakNonPagedPoolUsage;
'    DWORD QuotaNonPagedPoolUsage;
'    DWORD PagefileUsage;
'    DWORD PeakPagefileUsage;
'} PROCESS_MEMORY_COUNTERS;
'typedef PROCESS_MEMORY_COUNTERS *PPROCESS_MEMORY_COUNTERS;

Type PROCESS_MEMORY_COUNTERS
   cb As Long
   PageFaultCount As Long
   PeakWorkingSetSize As Long
   WorkingSetsize As Long
   QuotaPeakPagedPoolUsage As Long
   QuotaPagedPoolUsage As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage As Long
   PagefileUsage As Long
   PeakPagefileUsage As Long
End Type

'typedef struct _PSAPI_WS_WATCH_INFORMATION {
'    LPVOID FaultingPc;
'    LPVOID FaultingVa;
'} PSAPI_WS_WATCH_INFORMATION, *PPSAPI_WS_WATCH_INFORMATION;

'Type PSAPI_WS_WATCH_INFORMATION
'   FaultingPc As Long
'   FaultingVa As Long
'End Type

'BOOL EmptyWorkingSet(  HANDLE hProcess  // identifies the process);
'Private Declare Function EmptyWorkingSet Lib "PSAPI.DLL" ( _
'   ByVal hProcess As Long _
') As Long

'BOOL EnumDeviceDrivers(
'  LPVOID *lpImageBase,  // array to receive the load addresses
'  DWORD cb,             // size of the array
'  LPDWORD lpcbNeeded    // receives the number of bytes returned);
'Private Declare Function EnumDeviceDrivers Lib "PSAPI.DLL" ( _
'   lpImageBase() As Long, _
'   ByVal cb As Long, _
'   lpcbNeeded As Long _
') As Long

'BOOL EnumProcesses(
'  DWORD * lpidProcess,  // array to receive the process identifiers
'  DWORD cb,             // size of the array
'  DWORD * cbNeeded      // receives the number of bytes returned);
'Private Declare Function EnumProcesses Lib "PSAPI.DLL" ( _
'   lpidProcess As Long, _
   ByVal cb As Long, _
   cbNeeded As Long _
) As Long

'BOOL EnumProcessModules(  HANDLE hProcess,      // handle to the process
'  HMODULE * lphModule,  // array to receive the module handles
'  DWORD cb,             // size of the array
'  LPDWORD lpcbNeeded    // receives the number of bytes returned);
'Public Declare Function EnumProcessModules Lib "PSAPI.DLL" _
'   (ByVal hProcess As Long, _
'   lphModule As Long, _
'   ByVal cb As Long, _
'   lpcbNeeded As Long _
') As Long

'DWORD GetDeviceDriverBaseName(
'  LPVOID ImageBase,  // the load address of the driver
'  LPTSTR lpBaseName, // receives the base name of the driver
'  DWORD nSize        // size of the buffer);
'Public Declare Function GetDeviceDriverBaseName Lib "PSAPI.DLL" Alias "GetDeviceDriverBaseNameA" _
'   (ByVal ImageBase As Long, _
'   ByVal lpBaseName As String, _
'   ByVal nSize As Long _
') As Long

'DWORD GetDeviceDriverFileName(
'  LPVOID ImageBase,  // the load address of the driver
'  LPTSTR lpFilename, // buffer that receives the path
'  DWORD nSize        // size of the buffer);
'Public Declare Function GetDeviceDriverFileName Lib "PSAPI.DLL" Alias "GetDeviceDriverFileNameA" _
'   (ByVal ImageBase As Long, _
'   ByVal lpFileName As String, _
'   ByVal nSize As Long _
') As Long

'DWORD GetMappedFileName(  HANDLE hProcess,    // handle to the process
'  LPVOID lpv,         // the address to verify
'  LPTSTR lpFilename,  // buffer that receives the filename
'  DWORD nSize         // size of the buffer);
'Public Declare Function GetMappedFileName Lib "PSAPI.DLL" Alias "GetMappedFileNameA" _
'   (ByVal hProcess As Long, _
'   ByVal lpv As Long, _
'   ByVal lpFileName As String, _
'   ByVal nSize As Long _
') As Long

'DWORD GetModuleBaseName(  HANDLE hProcess,    // handle to the process
'  HMODULE hModule,    // handle to the module
'  LPTSTR lpBaseName,  // buffer that receives the base name
'  DWORD nSize         // size of the buffer);
'Public Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" _
'   (ByVal hProcess As Long, _
'   ByVal hModule As Long, _
'   ByVal lpFileName As String, _
'   ByVal nSize As Long _
') As Long

'DWORD GetModuleFileNameEx(  HANDLE hProcess,    // handle to the process
'  HMODULE hModule,    // handle to the module
'  LPTSTR lpFilename,  // buffer that receives the path
'  DWORD nSize         // size of the buffer);
'Public Declare Function GetModuleFileNameEx Lib "PSAPI.DLL" Alias "GetModuleFileNameExA" _
'   (ByVal hProcess As Long, _
'   ByVal hModule As Long, _
'   ByVal lpFileName As String, _
'   ByVal nSize As Long _
') As Long

'BOOL GetModuleInformation(  HANDLE hProcess,         // handle to the process
'  HMODULE hModule,         // handle to the module
'  LPMODULEINFO lpmodinfo,  // structure that receives information
'  DWORD cb                 // size of the structure);
'Public Declare Function GetModuleInformation Lib "PSAPI.DLL" _
'   (ByVal hProcess As Long, _
'   ByVal hModule As Long, _
'   lpmodinfo As MODULEINFO, _
'   ByVal cb As Long _
') As Long

'BOOL GetProcessMemoryInfo(  HANDLE Process,  // handle to the process
'  PPROCESS_MEMORY_COUNTERS ppsmemCounters,
'                   // structure that receives information
'  DWORD cb         // size of the structure);
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" _
   (ByVal hProcess As Long, _
   ppsmemCounters As PROCESS_MEMORY_COUNTERS, _
   ByVal cb As Long _
) As Long

'BOOL GetWsChanges(  HANDLE hProcess,  // handle to the process
'  PPSAPI_WS_WATCH_INFORMATION lpWatchInfo,
'                    // structure that receives information
'  DWORD cb          // size of the structure);
'Public Declare Function GetWsChanges Lib "PSAPI.DLL" _
'   (ByVal hProcess As Long, _
'   lpWatchInfo As PSAPI_WS_WATCH_INFORMATION, _
'   ByVal cb As Long _
') As Long

'BOOL InitializeProcessForWsWatch(  HANDLE hProcess  // handle to the process);
'Public Declare Function InitializeProcessForWsWatch Lib "PSAPI.DLL" _
'   (ByVal hProcess As Long _
') As Long
'End If

'BOOL QueryWorkingSet(  HANDLE hProcess,  // handle to the process
'  PVOID pv,         // buffer that receives the information
'  DWORD cb          // size of the buffer);
Private Declare Function QueryWorkingSet Lib "PSAPI.DLL" _
   (ByVal hProcess As Long, _
   pv As Long, _
   ByVal cb As Long _
) As Long

'Moved from the form
Private Const MAX_PATH = 260
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_QUERY_INFORMATION = &H400
'private const PROCESS_ALL_ACCESS        =(STANDARD_RIGHTS_REQUIRED or SYNCHRONIZE | \
'                                   0xFFF)
'Added by jna
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function GetWorkingSetSize() As Long
Dim lProcessID As Long
Dim hProcess As Long
Dim tPMC As PROCESS_MEMORY_COUNTERS

    lProcessID = GetCurrentProcessId

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                   PROCESS_VM_READ, _
                                   0, lProcessID)
    If (GetProcessMemoryInfo(hProcess, tPMC, Len(tPMC)) <> 0) Then
'This is the memory usage
'Debug.Print Format$(tPMC.WorkingSetsize \ 1024, "#,###,###") & "KB"
        GetWorkingSetSize = tPMC.WorkingSetsize
        End If
   CloseHandle hProcess
End Function
