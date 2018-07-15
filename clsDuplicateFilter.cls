VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDuplicateFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Storage for the debug ID.
Private mlngDebugID As Long

Private Const MAX_SECTION_NO = (2 ^ 11) - 1 '127
Private Const MIN_BUFFER_SIZE = 20
Private MIN_TTL As Long   '6 mins
Private MAX_TTL As Long

Private Type BufDef
    SequenceNo As Currency
    UnixTime As Long
    Data As String
End Type

Private Type SectionDef
    Buffer() As BufDef  'base 1
    Qfront As Long  'Last record inserted into the queue
    Qrear As Long    'End of queue (Oldest record)
    Qmax As Long    'Number of slots in the queue
    Qoldest As Long 'reqd for debug Oldest record in queue
    OldestTime As Long  'of earliest message in this section
    Qinuse As Long  'slots in use
    Qlast As Long   'Last insert
End Type

'Status Bar Variables
Private NextSBTime As Long  'Unix time for next display

Private TargLatency As Long
'Parent is only set of this Filter is owned by a socket
Private Parent As Long  'Idx of parent (0 if no parent)
Private vEnabled As Boolean 'Only used if Owned by sockets
Private vDmzIdx As Long
Private vOnlyVdm As Boolean 'True = Only pass VDM
Private vRejectMmsi As String  'Blank= none
Private vRejectPayloadErrors As Boolean  'True= reject if an error
Private Section(0 To MAX_SECTION_NO) As SectionDef
Private DupCount As Double
Private MsgCount As Double
Private sTtl As Long
Private arry() As Long  'used to resize .buffer

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

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
Public Property Let OnlyVdm(vNewValue As Boolean)
    vOnlyVdm = vNewValue
End Property
Public Property Get OnlyVdm() As Boolean
    OnlyVdm = vOnlyVdm
End Property
Public Property Let RejectMmsi(vNewValue As String)
    vRejectMmsi = vNewValue
End Property
Public Property Get RejectMmsi() As String
    RejectMmsi = vRejectMmsi
End Property
Public Property Let RejectPayloadErrors(vNewValue As Boolean)
    vRejectPayloadErrors = vNewValue
End Property
Public Property Get RejectPayloadErrors() As Boolean
    RejectPayloadErrors = vRejectPayloadErrors
End Property

Public Property Let DmzIdx(vNewValue As Long)
    vDmzIdx = vNewValue
End Property
Public Property Get DmzIdx() As Long
    DmzIdx = vDmzIdx
End Property

'Changed to byref to try & stop memory leak
Public Function IsDuplicate(ByRef SequenceNo As Currency, ByRef Source As Long, ByRef Data As String) As Long
Dim j As Long
Dim Qdiff As Long
Dim kb As String
Static LastPartDiff(1 To 9) As Long
Dim NoOfParts As Long
Dim PartNo As Long
Dim SectionNo As Long
Dim CurrTime As Long    'Unix Time
Dim i As Long
Dim Ttl As Long
Dim Crc As Long
Dim Age As Long
Dim ClearBuffer As BufDef
'First time IsDuplicate is called - any Section
    If MsgCount = 0 Then
        MsgCount = 1
'        If DupeLogCh = 0 Then
'            DupeLogCh = FreeFile
'            Open TempPath & "Dupe.log" For Output As #DupeLogCh
'        End If
        Call Status 'Initial
Debug.Print "-----"
    End If

'IsDuplicate = -1    '3
    
    IsDuplicate = -1    'assume not a duplicate
If True Then   '9
    If vEnabled = False Then Exit Function

'Exit Function       '4

kb = Right$(Data, 2)
'Val returns 0-127 if 2 characters
'Uppercase letter are converted to lowercase
'If any letter is not a valid Hex Letter the 0 is returned
    SectionNo = Val("&H" & Right$(Data, 2))
    Crc = CRCchk(Mid$(Data, 25, 4)) '0 to 15
    SectionNo = SectionNo + 128 * Crc   '0 to 2047
'SectionNo = Crc    'test 0-15
    If SectionNo > MAX_SECTION_NO Then
        SectionNo = MAX_SECTION_NO
    End If
'Debug distribution of messages
'arry(Crc) = arry(Crc) + 1
With Section(SectionNo)
'    On Error GoTo IsDuplicate_err
'On Error GoTo 0 'debug only
'First time for this buffer
    
        If .Qrear = 1 Then   'Last in buffer
        End If
        
        MsgCount = MsgCount + 1
'debug        CurrTime = UnixNow - 450 'Hold as variable so I can check it easily
        CurrTime = UnixNow 'Hold as variable so I can check it easily
'First check if we already have this sentence in the buffer
'Check for duplicates - most recent message first
'If not first time for this section
    
If True Then   '6
    If .Qinuse > 0 Then
'at least 1 sentence in the buffer
        j = .Qlast
        Do
'If the first entry in the buffer matches the incoming sentence
'the difference is 1 slot
'If we have not found a duplicate when no time
'or older than Max_TTL then we will not find one in
'this slot
            Age = CurrTime - .Buffer(j).UnixTime
            If Age > MAX_TTL Then
'No duplicate found for this data
                .Buffer(j) = ClearBuffer
'                Exit Do
            End If
'If true Then   '5 gros if this code is not used
            If .Buffer(j).Data = Data Then
Qdiff = SequenceNo - .Buffer(j).SequenceNo
                IsDuplicate = Qdiff
'                On Error Resume Next    'skip
                If IsNumeric(Mid$(Data, 8, 1)) = False Then
                    NoOfParts = 0
                Else
                    If IsNumeric(Mid$(Data, 10, 1)) = False Then
                        NoOfParts = 0
                    Else
                        NoOfParts = CLng(Mid$(Data, 8, 1))
                    End If
                End If
                If NoOfParts > 1 Then
'Multi-part message
                    PartNo = CLng(Mid$(Data, 10, 1))
                    LastPartDiff(PartNo) = Qdiff
                    If PartNo > 1 Then
                        If LastPartDiff(PartNo) <> LastPartDiff(1) Then
                            IsDuplicate = -1
                        Else
                            IsDuplicate = Qdiff
                        End If
                    End If
                End If
'                On Error GoTo IsDuplicate_err
                If IsDuplicate > 0 Then
                    DupCount = DupCount + 1
kb = Qdiff & vbTab & Source & vbTab & Data
If DupeLogCh <> 0 Then Print #DupeLogCh, kb
                    Exit Do
                End If
            End If
            If j = .Qoldest Then
                Exit Do
            End If
            j = j - 1
            If j < LBound(.Buffer) Then
                j = UBound(.Buffer) '.Qmax
            End If  'duplicate
        Loop
    End If  'not first time for this section
End If
'IsDuplicate = -1    '1-grows
'Exit Function       '2 still grows slowly
'We have a message to put in the buffer
'FIRST redimension the buffer if required

'Consider extending buffer
'   If .Qrear = .Qmax Then
'        .OldestTime = .Buffer(.Qfront).UnixTime
'Must be set when sentence is placed in buffer
'at the end of the routine
''        .Qoldest = .Qfront  'reqd to debug
'        Ttl = CurrTime - .OldestTime
'        If sTtl = 0 Then sTtl = Ttl 'initial setting
'        sTtl = sTtl + 0.03 * (Ttl - sTtl)
'        If Ttl < MIN_TTL Then
'            Call ExtendBuffer(SectionNo, CurrTime)
'        End If
'   End If

'Consider Contracting Buffer
'Weve got a time in the last slot
'    If .Buffer(.Qmax).UnixTime > 0 Then
'        If CurrTime - .Buffer(.Qmax).UnixTime > MAX_TTL Then
'            If .Qmax > MIN_BUFFER_SIZE Then
'                Call ContractBuffer(SectionNo, CurrTime)
'            End If
'        End If
'    End If

If True Then   '7
'Check if queue is empty
    If .Qinuse = 0 Or .Qoldest = 0 Then
'Debug.Print "Empty"
    Else
        .OldestTime = CurrTime - .Buffer(.Qoldest).UnixTime
'consider extending buffer
        Select Case .OldestTime
        Case Is < MIN_TTL
'No point if buffer is not full
            If .Qinuse = UBound(.Buffer) Then
'                Call ExtendBuffer(SectionNo)
                Call ResizeBuffer(SectionNo, CurrTime)
'Stop
            End If
        Case Is > MAX_TTL
'consider contracting buffer
            If .OldestTime > MAX_TTL Then
                Call ResizeBuffer(SectionNo, CurrTime)
            End If
        End Select
    End If
End If  '7
If True Then   '8
'Queue Next message (first message goes in slot 2
    If .Qlast = UBound(.Buffer) Then '.Qmax Then
        .Qlast = LBound(.Buffer)
    Else
        .Qlast = .Qlast + 1
    End If
            
'empty slot
    If .Buffer(.Qlast).UnixTime = 0 Then
        .Qinuse = .Qinuse + 1
    End If
'This is where the Oldest Sentence is overwritten
'if the buffer is full
    .Buffer(.Qlast).Data = Data
    .Buffer(.Qlast).SequenceNo = SequenceNo
    .Buffer(.Qlast).UnixTime = CurrTime

    .Qoldest = .Qlast
'scan forwards until the first slot in use
'which will be the oldest
'we need to do this because the buffer may have been extended
        Do
            .Qoldest = .Qoldest + 1
            If .Qoldest > UBound(.Buffer) Then
                .Qoldest = LBound(.Buffer)
            End If
        Loop Until .Buffer(.Qoldest).UnixTime <> 0
    
'Debug.Print .Qlast & " Oldest = " & .Qoldest & ":" & CurrTime - .Buffer(.Qoldest).UnixTime & " secs"

'Call DisplayBuffer(SectionNo, CurrTime, "End IsDuplicate")

'most you can get on a spreadsheet If DupCount > 60000 Then Stop
End If  '8
End With    'Section(SectionNo)
If UnixNow >= NextSBTime Then Call Status
End If  '9
    Exit Function

IsDuplicate_err:
    MsgBox "IsDuplicate Error " & Str(err.Number) & " " & err.Description & vbCrLf _
    & "Counter " & Section(SectionNo).Qfront, , "IsDuplicate Error"
End Function

Private Function ResizeBuffer(SectionNo As Long, CurrTime As Long)
Dim Factor As Single
Dim SizeAdj As Long
Dim OldElem As Long
Dim kb As String
Dim arry() As BufDef
Dim NewQmax As Long
Dim NewElem As Long
Dim OldestAge As Long
Dim Age As Long

If False Then
'Exit Function
End If
    With Section(SectionNo)
'Debug.Print "Section=" & SectionNo
'Debug.Print ".Oldest=" & .Qoldest & ", OldestTime=" & .OldestTime
'Call DisplayBuffer(SectionNo, Time, "ExtendBuffer")
        
'should not happen
        If .Qoldest = 0 Then Exit Function

        For OldElem = 1 To UBound(.Buffer)
            If .Buffer(OldElem).UnixTime Then
                Age = CurrTime - .Buffer(OldElem).UnixTime
                If Age > OldestAge Then
                    OldestAge = Age
                End If
            End If
        Next OldElem
        If .OldestTime <> OldestAge Then
'Call DisplayBuffer(SectionNo, CurrTime, "Age error")
'            Stop
            Exit Function
        End If
'Div/0
        If .OldestTime > 0 Then
            Factor = MAX_TTL / .OldestTime
'50% max increase in size
        Else
'add 50% if time is very fast
            Factor = 1.5
        End If
'as a minimum add 10%
        Select Case Factor
        Case Is > 1.5
If False Then
            Factor = 1.5
End If
        Case Is < 0.5
            Factor = 0.5
        Case Is > 1
If False Then
            Factor = 1.1
End If
        Case Is < 0.8
            Factor = 0.9
        Case Else   'leave
'Debug.Print "Factor=" & Factor
            Factor = 1
        End Select
        If Factor = 1 Then Exit Function
        SizeAdj = UBound(.Buffer) * Factor - UBound(.Buffer)
'50% of 20 Is 10
        If SizeAdj < 9 And SizeAdj > -9 Then
            SizeAdj = 0
        End If
        NewQmax = UBound(.Buffer) + SizeAdj
'check actual limits
        If NewQmax < MIN_BUFFER_SIZE Then
            NewQmax = MIN_BUFFER_SIZE
        End If

If NewQmax = UBound(.Buffer) Then
kb = "Leaving " & UBound(.Buffer) & " as " & NewQmax & " Old=" & .OldestTime
End If
'no change in buffer size
        If NewQmax = UBound(.Buffer) Then
            Exit Function
        End If
'create temporary array to new size
        ReDim arry(1 To NewQmax)
'load temp array from existing - oldest first
        OldElem = .Qoldest
        Do
            If .Buffer(OldElem).UnixTime <> 0 _
                And CurrTime - .Buffer(OldElem).UnixTime <= MAX_TTL Then
                NewElem = NewElem + 1
                arry(NewElem) = .Buffer(OldElem)
            End If
            If OldElem = UBound(.Buffer) Then
                OldElem = LBound(.Buffer)
            Else
                OldElem = OldElem + 1
            End If
'All elements in the old buffer have been transferred
'(ie buffer is being extenmded)
            If OldElem = .Qoldest Then
                Exit Do
            End If
'Possibility of there still being more recent
'sentences left in the old buffer
'In which case do not contract the buffer at
'this time (it will be done later)
            If NewElem = UBound(arry) Then
                Exit Function
            End If
        Loop
            
kb = kb & .OldestTime & " "
        .Qoldest = 1
        .Qinuse = NewElem
        .Qlast = NewElem
        .OldestTime = CurrTime - arry(.Qoldest).UnixTime
If NewQmax > UBound(.Buffer) Then
kb = kb & "Incr " & UBound(.Buffer) & " to " & NewQmax & " Old=" & .OldestTime
End If
If NewQmax < UBound(.Buffer) Then
kb = kb & "Decr " & UBound(.Buffer) & " to " & NewQmax & " Old=" & .OldestTime
End If
Debug.Print kb
kb = kb & " " & SectionNo & " " & UBound(arry) - UBound(.Buffer)
        Erase .Buffer
        .Buffer = arry
        Erase arry
    kb = kb & ", " & Format$(GetWorkingSetSize \ 1024, "#,###,###") & " Kb"
'Debug.Print kb
    End With
End Function

Private Sub DisplayBuffer(SectionNo As Long, Time As Long, Title As String)
Dim kb As String
Dim i As Long
Dim Ttl As Long

'Exit Sub
    With Section(SectionNo)
        kb = ""
        For i = 1 To UBound(.Buffer)
            kb = kb & .Buffer(i).SequenceNo & ","
        Next i
        kb = kb & vbCrLf & "Qfront=" & .Qfront & ", Qrear=" & .Qrear & ", Qmax=" & .Qmax
        kb = kb & vbCrLf & "OldestTime=" & .OldestTime
        kb = kb & " (" & .Qoldest & ")"
        kb = kb & ", Current Time=" & Time
        Ttl = Time - .OldestTime
        kb = kb & vbCrLf & "TTL=" & Ttl
        If Ttl < MIN_TTL Then
            kb = kb & vbCrLf & "TTL < MIN_TTL (" & MIN_TTL & ")"
        End If
        If Ttl > MAX_TTL Then
            kb = kb & vbCrLf & "TTL > MAX_TTL (" & MAX_TTL & ")"
            kb = kb & vbCrLf
            For i = 1 To UBound(.Buffer)
                If .Buffer(i).UnixTime Then
                    kb = kb & Time - .Buffer(i).UnixTime
                End If
                kb = kb & ","
            Next i
        End If
    End With
MsgBox kb, , Title

End Sub

Public Function Status(Optional SectionNo As Variant)
Dim UsedBufferSize As Long
Dim TotBufferSize As Long
Dim kb As String
Dim i As Long
Dim RejectionRate As Single
Dim j As Long
Dim CurrTime As Long

    CurrTime = UnixNow 'Hold as variable so I can check it easily
    kb = ""
    For i = 0 To UBound(Section)
        With Section(i)
'Buffer is empty
            On Error Resume Next
            TotBufferSize = TotBufferSize + UBound(.Buffer)
'    If .OldestTime > 0 And (CurrTime - .OldestTime) > MAX_TTL * 1.2 Then
'        Call ContractBuffer(i, CurrTime)
'    Stop
'    End If
            
            For j = 1 To UBound(.Buffer)
                If .Buffer(j).SequenceNo > 0 Then
                    UsedBufferSize = UsedBufferSize + 1
                End If
            Next j
            On Error GoTo 0
        End With
'arry(i) = UBound(Section(i).Buffer)
'        kb = kb & arry(i)
'If Not IsMissing(SectionNo) Then
'            If i = SectionNo Then kb = kb & "*"
'        End If
        kb = kb & ","
    Next i
'    If Not IsMissing(SectionNo) Then
'MsgBox kb
'    end if
    kb = ""
    If MsgCount Then
        If MsgCount Then
            RejectionRate = DupCount / MsgCount
        End If
        If Not IsMissing(SectionNo) Then
            kb = Format$(SectionNo, "0000") & ": "
        End If
        kb = kb & "Duplicate Rejection rate is " _
        & Format$(RejectionRate, "#0.00%")
'    BufferBytes = UsedBufferSize * 200
        kb = kb & ", Checking " & UsedBufferSize & " sentences"
'    kb = kb & Format$(CSng(BufferBytes / 1000000), "##0.0") & " Mbytes"
        If Not IsMissing(SectionNo) Then
            kb = kb & ":" & UBound(Section(SectionNo).Buffer)
        End If
        If Not IsMissing(SectionNo) Then
        kb = kb & ", TTL is " & Format$(sTtl, "###0.00")
            kb = kb & " (" _
            & Format$(UnixNow _
            - Section(SectionNo).OldestTime, "###0.00") & ")"
        kb = kb & " Secs"
        End If
    
        If kb <> "" Then kb = kb & ", "
    End If
'    kb = kb & Format$(GetWorkingSetSize \ 1024, "#,###,###") & " Kb"
    kb = kb & " (" & TotBufferSize & ")"
    frmRouter.StatusBar.Panels(1).Text = kb
'clear after 30 secs
    frmRouter.ClearStatusBarTimer.Enabled = True
    NextSBTime = UnixNow + 10  '30 seconds
End Function

Public Function Reset()
Dim SectionNo As Long

    MsgCount = 0
    DupCount = 0
    MIN_TTL = 6 * 60
    MAX_TTL = MIN_TTL * 1.2
    For SectionNo = 0 To UBound(Section)
        With Section(SectionNo)
            Erase .Buffer
            .Qmax = MIN_BUFFER_SIZE      ' 1 sec at 600 msgs/min (default)
            ReDim .Buffer(1 To .Qmax)
            .Qfront = 0 '.Qmax    'Fill from end (redim when full loses oldest msgs)
            .Qrear = 1  'was 0
            .OldestTime = 0
            .Qoldest = 0    'was 0
            .Qinuse = 0
            .Qlast = 0
        End With
    Next SectionNo

End Function

Private Function CRCchk(Data As String) As Long
Dim CheckSum As Byte
Dim Chr As String
Dim HexChecksum As String
Dim b() As Byte
Dim i As Long

b = StrConv(Data, vbFromUnicode)
If UBound(b) <= 0 Then Exit Function '0 or less characters
CheckSum = b(1) 'set the first byte to be checked
For i = 1 + 1 To UBound(b) 'Excluces !,$,\ and * in last word
    CheckSum = CheckSum Xor b(i)
Next
CRCchk = CheckSum And 15 'bottom 4 bits
End Function

Private Sub Class_Initialize()
   
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsDuplicateFilter; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)

'set initial values
Debug.Print "Initialise " & True & "/" & False
    Call Reset
End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
Debug.Print "Terminate"
End Sub
