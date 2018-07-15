VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form Filecfg 
   Caption         =   "File Configuration"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3225
   Icon            =   "Filecfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboReadRate 
      Height          =   315
      ItemData        =   "Filecfg.frx":058A
      Left            =   1440
      List            =   "Filecfg.frx":05A9
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkRollOver 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".log"
      DialogTitle     =   "Open for Input"
      FileName        =   "Router"
      InitDir         =   "%APPDATA%\Arundale\Router"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Direction"
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optDirection 
         Caption         =   "Output"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input,Output"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Roll Over"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sentences/min"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Filecfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancel As Boolean

Dim Hidx As Long


Private Sub cmdCancel_Click()
    Cancel = False
    Set Files(Hidx) = Nothing
    Unload Me
End Sub


Private Sub CmdOk_Click()
    
'Get the Required file name and validate it
    On Error GoTo Cancelled
    With dlgFile
        Select Case CurrentDirection
            Case Is = 1     'input
            .DialogTitle = "Open for Input"
            .Flags = cdlOFNPathMustExist
            .Flags = cdlOFNFileMustExist
            .Flags = .Flags + cdlOFNHideReadOnly
            .DefaultExt = ".log"
        Case Is = 2     'output
            .DialogTitle = "Open for Output"
            .Flags = cdlOFNOverwritePrompt
            .Flags = .Flags + cdlOFNPathMustExist
            .Flags = .Flags + cdlOFNExplorer
            .Flags = .Flags + cdlOFNHideReadOnly
            .Flags = .Flags + cdlOFNNoReadOnlyReturn
            .DefaultExt = ".log"
        End Select
'If file is open for input do not allow file name to be changed
'(File must be closed) - Same as Form load
        If Not (Sockets(CurrentSocket).State > 0 _
        And CurrentDirection = 1) Then
            .ShowOpen
'If not cancelled set up in Sockets()
            Sockets(CurrentSocket).File.SocketFileName = .FileName
        End If
        Sockets(CurrentSocket).Direction = CurrentDirection
       If chkRollOver = vbChecked Then
           Sockets(CurrentSocket).File.RollOver = True
        Else
           Sockets(CurrentSocket).File.RollOver = False
        End If
        Sockets(CurrentSocket).File.ReadRate = cboReadRate.ListIndex
        Unload Me   'Return to frmRouter
    Exit Sub
    
Cancelled:
MsgBox .FileName, , "Cancelled"
    End With
End Sub

Private Sub Form_Load()

    If Sockets(CurrentSocket).Hidx > 0 Then
        Hidx = Sockets(CurrentSocket).Hidx
'Disable Comm
'        Call TTYs(Hidx).CloseTTY
    Else
        Hidx = FreeFileHidx
'Set defaults
'Direction is got from initial setup on Filecfg
'if this is a new handler and not an existing one thats been closed
        If Sockets(CurrentSocket).Direction = -1 Then
            Sockets(CurrentSocket).Direction = CurrentDirection
        End If
    End If
    
    If Hidx = -1 Then
        MsgBox "No free File Sockets", , "Filecfg.Load"
        Exit Sub
    End If
    
    Caption = "File Configuration [" & Sockets(CurrentSocket).DevName & "]"
'Set Direction option to same as in Sockets()
    optDirection(Sockets(CurrentSocket).Direction).Value = Sockets(CurrentSocket).Direction
    cboReadRate.ListIndex = Sockets(CurrentSocket).File.ReadRate
    
'If file is open for input do not allow Direction to be changed
'(File must be closed)
        If Sockets(CurrentSocket).State > 0 _
        And CurrentDirection = 1 Then
            optDirection(1).Enabled = False
            optDirection(2).Enabled = False
        Else
            optDirection(1).Enabled = True
            optDirection(2).Enabled = True
        End If

'Enable/disable options
    Call CurrentDirection
    If Sockets(CurrentSocket).File.RollOver = True Then
        chkRollOver = vbChecked
    Else
        chkRollOver = vbUnchecked
    End If
  '  cboReadRate.ListIndex = Sockets(CurrentSocket).File.ReadRate
    
    
    dlgFile.Filter = "All Files (*.*)|*.*|" _
    & "Log Files (*.log)|*.log"
    dlgFile.FilterIndex = 2
End Sub
'Returns the Direction from optDirection
Private Function CurrentDirection() As Long
Dim i As Integer
    For i = 0 To optDirection.UBound
        If optDirection(i).Value = True Then
            CurrentDirection = i
            Exit For
        End If
    Next i
    Select Case CurrentDirection
        Case Is = 1
            chkRollOver.Enabled = False
            cboReadRate.Enabled = True
        Case Is = 2
            chkRollOver.Enabled = True
            cboReadRate.Enabled = False
    End Select
End Function

Private Sub optDirection_Click(Index As Integer)
'    If Me.Visible Then
    Call CurrentDirection
'    End If
End Sub



