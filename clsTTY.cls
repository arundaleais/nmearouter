VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Public TTY As frmTTY    'TTYs(CurrentSocket).TTY
Public Name As String
Public State As Long
Private Hidx As Long    'TTY index
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

Public Function TTYOutput(Data As String)
    
    With TTYs(Hidx).TTY.TermText
        .SelStart = Len(.Text)
        .SelText = Data
         If Len(.Text) > 4096 Then
            .Text = Right$(.Text, 2048)
        End If
'    Call ForwardData(commdata, Source)
    End With
End Function

Public Function TTYClose()
Dim i As Long

'Find this form
    For i = 0 To Forms.Count - 1
        If Forms(i).Caption = Name Then
            Exit For
        End If
    Next i
'Remove references to form before unloading
    Set TTYs(Sockets(Idx).Hidx).TTY = Nothing
    Set TTYs(Sockets(Idx).Hidx) = Nothing
    Unload Forms(i)
End Function

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsTTY; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
    WriteLog Name & " closed"
End Sub
