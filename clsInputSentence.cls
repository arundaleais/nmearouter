VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInputSentence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Storage for the debug ID.
Private mlngDebugID As Long

'This contains the decoded data for
'   1.  Input Filter
'   2.  Output CRC,NMEA and View
'these must be defined public in the class module (if not using properies)
Public FullSentence As String   'Includes Comment Block
Public NmeaPrefix As String     'Part of sentence before ! or $ if no comment block
'Public CommentBlock As String   '/.../ at begining of sentence IEC61162-1
'Public CBerrmsg As String   'Comment Block error
Public NmeaSentence As String
Public NmeaSentenceType As String  'nmea word 0 sentence type
Public NmeaCrc As String
Public CRCerrmsg As String
Public AisPayload As String         'nmea word 6
Public AisRcvTime As String      'nmea word 7 created time
Public AivdmComments As String
Public Aivdmerr As Boolean
Public AisMsgPart1Missing As Boolean    'means we cannot get AisMsgType
Public AisMsgPartsComplete As Boolean   'this is the last part
                                        'and all prev received ok
Public AisPayloadComments As String     'payload too short
Public IsAisSentence As Boolean   'True if AIS !aaVDO or !aaVDM

Public AisMsgType As String     'Number 1- 27
Public SentencePart As String   '0-9 nmea word (2)
Public AisMsgFromMmsi As String
Public AisMsgToMmsi As String
Public AisMsgDac As String
Public AisMsgFi As String
Public AisMsgFiId As String

Public VesselName As String

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsInputSentence; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub




