VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNmeaFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

'Parent is only set of this Filter is owned by a socket
Private Parent As Long  'Idx of parent (0 if no parent)
Private vEnabled As Boolean 'Only used if Owned by sockets
Private vDmzIdx As Long
Private RejectCount As Double
Private MsgCount As Double

Public Property Let ParentIdx(vNewValue As Long)
    Parent = vNewValue
End Property
Public Property Get ParentIdx() As Long
    ParentIdx = Parent
End Property

Public Property Let Enabled(vNewValue As Boolean)
    vEnabled = vNewValue
End Property
Public Property Get Enabled() As Boolean
    Enabled = vEnabled
End Property

Public Property Let DmzIdx(vNewValue As Long)
    vDmzIdx = vNewValue
End Property
Public Property Get DmzIdx() As Long
    DmzIdx = vDmzIdx
End Property

Public Function IsRejected(Data As String) As Boolean
    If vEnabled = False Then Exit Function
    On Error GoTo IsRejected_err
    MsgCount = MsgCount + 1
    If False Then
        RejectCount = RejectCount + 1
        IsRejected = True
    End If
    Exit Function

IsRejected_err:
    MsgBox "AisFilter Error " & Str(err.Number) & " " & err.Description _
    & vbCrLf, , "Ais Filter Error"
End Function

Public Function Reset()
    MsgCount = 0
    RejectCount = 0
End Function

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsNmeaFilter; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
    Call Reset
End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub


