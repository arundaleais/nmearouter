Attribute VB_Name = "modMMSI"
Option Explicit
    
Public NmeaWords() As String    'cant keep dynamic array in class module
Public clsSentence As New clsInputSentence   'create

Public Function VDMtoVDO(Sentence As String) As String
Dim i As Long   'Position of *
Dim VDO As String
Dim Crc As String
    i = InStr(1, Sentence, "*")
    If i > 0 Then
        Crc = Mid$(Sentence, i, 3)
        VDO = Replace(Sentence, "!AIVDM", "!AIVDO")
        VDO = Replace(VDO, Crc, "*[CRC]")
        VDO = Replace(VDO, "[CRC]", NmeaCrcChk(VDO))
        VDMtoVDO = VDO
    Else
        VDMtoVDO = Sentence
    End If
End Function


Public Function DecodeMMSI(InputSentence As String) As String
Dim From As Long    'to start position of bits
Dim Payload8Bits As Long
Dim CBFrom As Long  'Start of comment block
Dim CBTo As Long   'end of Comment block = 0 (excl delimeter) if not found
Dim NmeaFrom As Long   'Start of NMEA (incl $ or !)part of sentence = 0 if not found

    If InputSentence = "" Then Exit Function
'Same as last sentence decoded
    If clsSentence.FullSentence = InputSentence Then Exit Function
    
    Set clsSentence = Nothing
'    Set clsCB = Nothing
    Erase NmeaWords
'    Erase CbWords
    clsSentence.FullSentence = InputSentence
'Split the FullSentence into Comments and NMEA
'find start of nmea (IEC spec - ! or $ is always start of NMEA
    NmeaFrom = InStr(1, clsSentence.FullSentence, "!")
    If NmeaFrom = 0 Then
        NmeaFrom = InStr(1, clsSentence.FullSentence, "$")
    End If
    If NmeaFrom <> 0 Then   'start of NMEA found (! or $ always terminates CB)
'Extract NMEA sentence
        clsSentence.NmeaSentence = Mid$(clsSentence.FullSentence, NmeaFrom)
    End If
    
#If False Then
'Comments - sentence must start with \
    CBFrom = InStr(1, clsSentence.FullSentence, "\")
    If CBFrom > 0 Then
'find closing \
        CBTo = InStr(CBFrom + 1, clsSentence.FullSentence, "\")
        If CBTo > CBFrom Then   'no closing \ separator
'CBTo must exist so minimum must be at least 1
            clsCB.Block = Mid$(clsSentence.FullSentence, CBFrom, CBTo - CBFrom + 1)
            Call DecodeCommentBlock
        Else
'No comment block
            CBFrom = 0
            CBTo = 0
        End If
    End If

'Have we a NmeaPrefix (not a comment with only one \ separator)
'results in a prefix
    If NmeaFrom > CBTo + 1 Then
        clsSentence.NmeaPrefix = Mid$(clsSentence.FullSentence, CBTo + 1, NmeaFrom - 1)
    End If
#End If
'check we don't just have a comment block
    If clsSentence.NmeaSentence = "" Then Exit Function
'Note NmeaCRC check excludes first character ! or $
    clsSentence.CRCerrmsg = NmeaCrcChk(clsSentence.NmeaSentence)
    
'created into Public Variable becuase you cant have a dynamic array in a class module
    NmeaWords = Split(clsSentence.NmeaSentence, ",")  '350k/min
    clsSentence.NmeaSentenceType = NmeaWords(0)

'if CRC error all details are suspect (but display first word)
    If clsSentence.CRCerrmsg <> "" Then Exit Function
    If UBound(NmeaWords) >= 7 Then
'        ChkDate = CDbl(NmeaWords(7))
        If IsDate(NmeaWords(UBound(NmeaWords))) Then _
        clsSentence.AisRcvTime = NmeaWords(UBound(NmeaWords))
    Else
'        clsSentence.AisRcvTime = NowUtc()  'locale format
    End If
    
'compatible with !aaVDM or !aaVDO ais sentences
    If Left$(clsSentence.NmeaSentenceType, 1) = "!" _
    And (Right$(clsSentence.NmeaSentenceType, 3) = "VDM" Or _
    Right$(clsSentence.NmeaSentenceType, 3) = "VDO") Then
         clsSentence.IsAisSentence = True
    End If

'    Select Case clssentence.NmeaSentenceType
'    Case Is = "!AIVDM", "!AIVDO", "!BSVDM", "!BSVDO"
    If clsSentence.IsAisSentence = True Then
        If UBound(NmeaWords) < 6 Then
            clsSentence.AivdmComments = "AIS NMEA sentence incomplete"
            Exit Function
        End If
        If NmeaWords(2) = "1" And Len(NmeaWords(5)) < 2 Then
            clsSentence.AivdmComments = "Payload incomplete"
            Exit Function
        End If
'        clsSentence.NmeaCrc = SplitCrc(NmeaWords(6)) 'remove *crc
'extract information available if incomplete
        clsSentence.SentencePart = NmeaWords(2)
        clsSentence.Aivdmerr = Aivdm(clsSentence.NmeaSentence, clsSentence.AivdmComments) 'all bits
'Returns error if part 1 of multipart
'        If clsSentence.Aivdmerr = True Then
'            Exit Function
'        End If
'get AisMsgtype if part 1
        Payload8Bits = (PayloadByteArraySize + 1) * 8
        If clsSentence.AisMsgPart1Missing = False Then
            clsSentence.AisMsgType = pLong(1, 6)  '230k/min
        End If
'get other details when weve got a message type and enough bits
'message type will be missing if we have not received part 1
        If clsSentence.AisMsgType <> "" Then
            If Payload8Bits >= 38 Then      '39
                clsSentence.AisMsgFromMmsi = Format$(pLong(9, 30), "000000000")
            End If
'set flag to indicate ok to full decode of this sentence payload is ok
            If clsSentence.Aivdmerr = True Then
                clsSentence.AisMsgPartsComplete = False
            Else
                clsSentence.AisMsgPartsComplete = True
            End If  'complete ais message
        End If  'got AisMsgType
    End If 'AIS of not ais sentence
    DecodeMMSI = clsSentence.AisMsgFromMmsi
End Function

'If not AIS sentence then Error is False (Only AIS sentences can return an error)
'Error if IsPayloadError() <> ""
Public Function IsPayloadError(InputSentence As String) As String
Dim From As Long    'to start position of bits
Dim Payload8Bits As Long
Dim CBFrom As Long  'Start of comment block
Dim CBTo As Long   'end of Comment block = 0 (excl delimeter) if not found
Dim NmeaFrom As Long   'Start of NMEA (incl $ or !)part of sentence = 0 if not found
Dim kb As String

    If InputSentence = "" Then Exit Function
'find start of nmea (IEC spec - ! or $ is always start of NMEA
    NmeaFrom = InStr(1, InputSentence, "!")
    If NmeaFrom = 0 Then Exit Function  'Not AIS
'Comment Block is discarded
    kb = Mid$(InputSentence, NmeaFrom)
'check we don't just have a comment block
    If kb = "" Then Exit Function
'created into Public Variable becuase you cant have a dynamic array in a class module
    NmeaWords = Split(kb, ",")  '350k/min
    If UBound(NmeaWords) < 6 Then
        IsPayloadError = "AIS NMEA sentence incomplete"
        Exit Function
    End If
    
    If NmeaWords(2) = "1" And Len(NmeaWords(5)) < 2 Then
        IsPayloadError = "Payload incomplete"
        Exit Function
    End If
'The first character of the 1st Part
    If NmeaWords(2) = "1" Then
    Select Case Left$(NmeaWords(5), 1)  'Message type
    Case Is = "1", "2", "3", "4", "9", "B"  '4=Base,9=SAR,B=msg18
        If Len(NmeaWords(5)) <> 28 Then
            IsPayloadError = "Payload length incorrect"
            Exit Function
        End If
    End Select
    End If
    
End Function


