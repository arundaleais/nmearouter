VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Public IPPort As String       '123.456.789.123:12345
Public IPRetryCount As Long
Public MaxIPs As Long
Public PortRetryCount As Long
Public MaxPorts As Long
Public StreamRetryCount As Long        'Failures since last successful
Public MaxStreams As Long
Public WinsockRetryCount As Long        'Failures since last successful
Public MaxWinsocks As Long
Public SocketRetryCount As Long        'Failures since last successful
Public MaxSockets As Long

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsClient; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)

End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
