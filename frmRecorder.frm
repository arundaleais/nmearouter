VERSION 5.00
Begin VB.Form frmRecorder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recorder"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecorder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer CurrentTime 
      Interval        =   1000
      Left            =   3360
      Top             =   0
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4455
      Begin VB.Label Label5 
         Caption         =   "Record"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblRecCounter 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Playback"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    0"
         Height          =   300
         Left            =   3720
         TabIndex        =   12
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblPlayCounter 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblTime 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   10
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Time UTC"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer ReadTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         MaskColor       =   &H00000000&
         TabIndex        =   6
         ToolTipText     =   "Stop"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         MaskColor       =   &H00000000&
         TabIndex        =   5
         ToolTipText     =   "Fast Forward"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdRewind 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         MaskColor       =   &H00000000&
         TabIndex        =   4
         ToolTipText     =   "Rewind"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "ll"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaskColor       =   &H00000000&
         TabIndex        =   3
         ToolTipText     =   "Pause"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "< >"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         MaskColor       =   &H00000000&
         TabIndex        =   1
         ToolTipText     =   "Play"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdRec 
         Caption         =   "Rec"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   2
         ToolTipText     =   "Record"
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape shpPlay 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpForward 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpRewind 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpPause 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpRec 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   360
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Auto-Reverse Playback"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TimedNmeadef
    SequenceNo As Currency
    UtcUnix As Long    'Unix Time (Secs since 1 Jan 1970)
    NmeaData As String * 80
End Type

Private WriteFileName As String 'Set when form created, never changes
Private WriteFileCh As Integer
Private ReadFileCh As Integer
Private Parent As Long  'Idx
Private PlayCounter As Long
Private RecCounter As Long
Private Speed As Long
Private WriteRec As TimedNmeadef
Private ReadRec As TimedNmeadef
Private NxtReadRecNo As Long

Public Sub RecordData(SequenceNo As Currency, UtcUnix As Long, Data As String)
    If shpRec.Visible = True Then
        WriteRec.SequenceNo = SequenceNo
        WriteRec.UtcUnix = UtcUnix
        WriteRec.NmeaData = Data
        Put #WriteFileCh, Int(LOF(WriteFileCh) / Len(WriteRec)) + 1, WriteRec
        lblRecCounter = Format$(Int(LOF(WriteFileCh) / Len(WriteRec)), "000000")
        If lblTime.BackColor = vbRed Then
            lblTime = UnixTimeToDate(WriteRec.UtcUnix)
        End If
Debug.Print WriteRec.UtcUnix
    End If
End Sub

Private Sub cmdForward_Click()
    Select Case Speed
    Case Is > 0
        Speed = Speed * 2
    Case Is < -1
        Speed = Speed / 2
    Case Else   '-1 or 0
        Speed = 1
    End Select
    Call shpDirection
'Turn off at stop,Pause,EOF
End Sub

Private Sub cmdRewind_Click()
    Select Case Speed
    Case Is > 1
        Speed = Speed / 2
    Case Is < 0
        Speed = Speed * 2
    Case Else   '1 or 0
        Speed = -1
    End Select
    
    Call shpDirection
'Turn off at stop,Pause,EOF
End Sub

Private Sub cmdPause_Click()
'Toggles on off
    shpPause.Visible = Not shpPause.Visible
End Sub

Private Sub cmdPlay_Click()
Dim kb As String

'Dont restart without stopping first
    
    Call shpDirection
    If shpPlay.Visible = True Then
        Call Reverse
    End If
    shpPlay.Visible = True
    Call OpenFiles
    cmdStop.Enabled = True
    Call NextBlock
    Call TimeColor
End Sub

Private Sub cmdRec_Click()
'Toggle
    shpRec.Visible = Not shpRec.Visible
    If shpRec.Visible = True Then
        Call OpenFiles
    End If
    Call TimeColor
End Sub

Private Sub cmdStop_Click()
'Turns off Pause,FF,Rew,Play
'Must just do this and let the readloop close the channel
'otherwise the there is a possibility of EOF(ReadFileCh)
'returning and error
        
        cmdStop.Enabled = False
End Sub

Private Sub CurrentTime_Timer()
    CurrentTime.Enabled = True
    lblTime = UnixTimeToDate(UnixNow)
End Sub

Private Sub Form_Load()
'When Play first clicked will be changed to 1
    Speed = 1
    Call shpDirection
End Sub

Public Property Let ParentIdx(vNewValue As Long)
    Parent = vNewValue
'The Write file never changes name
    WriteFileName = TempPath & "Recorder_" & Sockets(Parent).DevName & ".rec"
End Property
Public Property Get ParentIdx() As Long
    ParentIdx = Parent
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Note if a TCP stream socket, the recorder will be closed when
'the sockets closed OR if frmRouter STOP is pressed
'When STOP is pressed other recorders will not be closed
'Otherwise the Recorder will only be closed when the sockets is removed
    If WriteFileCh > 0 Then
        Close WriteFileCh
        WriteFileCh = 0
    End If
    If ReadFileCh > 0 Then
        Close ReadFileCh
        ReadFileCh = 0
    End If
'We must call a Module to destroy the Object if (X) clicked
    Call RemoveRecorder(Parent)
End Sub

'Get next block of Speed records
Private Function NextBlock() As Boolean
Dim kb As String
Dim Count As Long

'Channel can be closed while NextBlock is waiting
    If ReadFileCh > 0 Then
        If NxtReadRecNo = 0 Then NxtReadRecNo = 1
        
        Get #ReadFileCh, NxtReadRecNo, ReadRec
        Do Until EOF(ReadFileCh) Or NxtReadRecNo = 0 _
            Or shpPause.Visible = True Or cmdStop.Enabled = False
            lblPlayCounter = Format$(NxtReadRecNo, "000000")
            lblTime = UnixTimeToDate(ReadRec.UtcUnix)
            NxtReadRecNo = NxtReadRecNo + (1 * Sgn(Speed))
            Count = Count + 1
'Possibility output socket is not open or been closed (would cause error)
            If Sockets(Parent).State = 1 Then
                Call OutputFormatter(ReadRec.SequenceNo, 0, ReadRec.UtcUnix, Parent, Trim$(ReadRec.NmeaData))
            End If
'Allow events at least every 100 cycles
If Count Mod 100 = 0 Then DoEvents
            If Count = Abs(Speed) Then Exit Do
'check output socket not been closed - possibly nmeaRouter stop pressed
            If Sockets(Parent).State <> 1 Then Exit Do
'Could be still in the count loop
            If NxtReadRecNo > 0 Then
                Get #ReadFileCh, NxtReadRecNo, ReadRec
            End If
        Loop
        If EOF(ReadFileCh) Or NxtReadRecNo = 0 Then
'You need to back off the next record by 1 only when EOF going forward
'Because when the counter is 0 next record is changed to 1 in nxtBlock
            If Sgn(Speed) = 1 Then
                NxtReadRecNo = NxtReadRecNo - 1
            End If
'update the counter to either 0 or record last successfully read
            lblPlayCounter = Format$(NxtReadRecNo, "000000")
            Call Reverse
        Else
'check output socket not been closed
            If Sockets(Parent).State = 1 Then
                ReadTimer.Enabled = True
            End If
'Nmea Router stop pressed, must not close ReadCh until out of loop
'otherwise EOF when channel closed causes error
            If cmdStop.Enabled = False Then
                Call ClosePlay
            End If
        End If
    Else
'only close when out of the read loop
        Call ClosePlay
    End If
    Call TimeColor
End Function

Private Sub ClosePlay()
        ReadTimer.Enabled = False
        NxtReadRecNo = 0
        lblPlayCounter = Format$(NxtReadRecNo, "000000")
        If Speed < 0 Then Speed = -Speed
        lblSpeed = Format$(Speed, "####")
        lblTime = ""
        lblTime.BackColor = vbWhite
        shpPlay.Visible = False
        shpPause.Visible = False
        cmdStop.Enabled = False
End Sub

Private Sub ReadTimer_Timer()
    
    Call NextBlock
End Sub

'This is the direction it will go when started
Private Sub shpDirection()
    shpForward.Visible = False
    shpRewind.Visible = False
    If Speed > 0 Then shpForward.Visible = True
    If Speed < 0 Then shpRewind.Visible = True
    lblSpeed = Format$(Speed, "####")
End Sub

Private Sub Reverse()
    Speed = -Speed
    Call shpDirection
Debug.Print NxtReadRecNo
End Sub

'Returns False if eof
Private Function FileRead(RecNo As Long) As Boolean
    Get #ReadFileCh, RecNo, ReadRec
    FileRead = True
    Exit Function
Error_Read:
End Function


Private Function UnixTimeToDate(ByVal Timestamp As Long) As String

          Dim intDays As Integer, intHours As Integer, intMins As Integer, intSecs As Integer
          intDays = Timestamp \ 86400
          intHours = (Timestamp Mod 86400) \ 3600
          intMins = (Timestamp Mod 3600) \ 60
          intSecs = Timestamp Mod 60
          UnixTimeToDate = DateSerial(1970, 1, intDays + 1) + TimeSerial(intHours, intMins, intSecs)
'            UnixTimeToDate = Format$(UnixTimeToDate, DateTimeOutputFormat)
      
End Function

Private Sub TimeColor()
    lblTime.BackColor = vbWhite     'If not recording or playback
    If shpPlay.Visible = True Then  'Takes precedence over recording
        lblTime.BackColor = vbGreen
    Else
        If shpRec.Visible = True Then
            lblTime.BackColor = vbRed
        End If
    End If
End Sub

Private Sub OpenFiles()
Dim reply As Long
'dont open twice
    If WriteFileCh = 0 Then
        If FileExists(WriteFileName) Then
            reply = MsgBox("If Yes, the existing recording " _
            & NameFromFullPath(WriteFileName) & vbCrLf _
            & "will be added to when recording and also " & vbCrLf _
            & "used for playback" & vbCrLf _
            & "If No, a new recording will be started" & vbCrLf _
            , vbYesNo + vbQuestion, "Existing Recording")
            If reply = vbNo Then
                Kill WriteFileName
            End If
        End If
        WriteFileCh = FreeFile
        Open WriteFileName For Random As #WriteFileCh Len = Len(WriteRec)
        lblRecCounter = Format$(Int(LOF(WriteFileCh) / Len(WriteRec)), "000000")
    End If
'Dont open again
    If ReadFileCh = 0 Then
        ReadFileCh = FreeFile
        Open WriteFileName For Random As #ReadFileCh Len = Len(ReadRec)
    End If
End Sub
