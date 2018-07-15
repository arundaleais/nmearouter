VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoopBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Public Name As String
Public State As Long
Private Hidx As Long    'LoopBack index
Private Idx As Long     'Socket Index

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Public Property Let sIndex(vNewValue As Long)
    Idx = vNewValue
End Property
Public Property Get sIndex() As Long
    sIndex = Idx
End Property

Public Property Let hIndex(vNewValue As Long)
    Hidx = vNewValue
End Property
Public Property Get hIndex() As Long
    hIndex = Hidx
End Property

Public Function LoopBackOutput(Data As String, Source As Long, Destination As Long)
'If other than a LoopBack this Handler would just Output the data
'We have to force the deque of the original Data (IF a Loopback)
'As it would not normally be sent until the next data is received from the
'same source, because the deque timer will stop it.
    Call modRouter.Deque(Source, 100)
'Here we should process any data (for example if this was an output filter)
    Call ForwardData(Data, Destination)
End Function



Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsLoopBack; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
   

End Sub

Private Sub Class_Terminate()
    ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
   WriteLog Name & " closed"
End Sub
