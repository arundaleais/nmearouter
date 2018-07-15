VERSION 5.00
Begin VB.Form Filtercfg 
   Caption         =   "Filter"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   Icon            =   "Filtercfg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRejectPayloadErrors 
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtNotMmsi 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox chkSourceNotAivdm 
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.ComboBox cboDevName 
      Height          =   315
      ItemData        =   "Filtercfg.frx":058A
      Left            =   1560
      List            =   "Filtercfg.frx":058C
      TabIndex        =   3
      Text            =   "Select an Output Connect or None"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox chkSourceDuplicates 
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Remove Ais Payload Error Sentences from all Sources"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Remove All Source Sentences from this MMSI"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Remove Non-!**VDM Sentences from all Sources"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Remove duplicated Sentences from all Sources"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "DMZ Connection"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "This is the Connection (if any) to which all rejected sentences will be Output"
      Top             =   150
      Width           =   1455
   End
End
Attribute VB_Name = "Filtercfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancel As Boolean

Private Sub cmdCancel_Click()
    Cancel = True
    Unload Me
End Sub


Private Sub CmdOk_Click()
    Cancel = False
    Call txtNotMmsi_Validate(Cancel)
    If Cancel = True Then Exit Sub
        
    If Cancel = False Then
        If chkSourceDuplicates.Value <> 0 Then
            SourceDuplicateFilter.Enabled = True
        Else
            SourceDuplicateFilter.Enabled = False
            Call SourceDuplicateFilter.Reset
        End If
        If chkSourceNotAivdm.Value <> 0 Then
            SourceDuplicateFilter.OnlyVdm = True
        Else
            SourceDuplicateFilter.OnlyVdm = False
        End If
        SourceDuplicateFilter.RejectMmsi = txtNotMmsi.Text
        If chkRejectPayloadErrors.Value <> 0 Then
            SourceDuplicateFilter.RejectPayloadErrors = True
        Else
            SourceDuplicateFilter.RejectPayloadErrors = False
        End If
        
        SourceDuplicateFilter.DmzIdx = _
        DevNameToSocket(cboDevName)
    End If
    
    Unload Me   'Return to frmRouter
Exit Sub
    
Cancelled:
End Sub

Private Sub Form_Load()
Dim Idx As Long

    cboDevName.AddItem "None"
    cboDevName.ListIndex = 0    'None
    For Idx = 1 To UBound(Sockets)
        If Sockets(Idx).State <> -1 And Sockets(Idx).Direction <> 1 Then
            cboDevName.AddItem Sockets(Idx).DevName
            If SourceDuplicateFilter.DmzIdx = Idx Then
                cboDevName.ListIndex = cboDevName.ListCount - 1
            End If
        End If
    Next Idx
    If SourceDuplicateFilter.Enabled = False Then
        chkSourceDuplicates.Value = 0
    Else
        chkSourceDuplicates.Value = 1
    End If
    If SourceDuplicateFilter.OnlyVdm = False Then
        chkSourceNotAivdm.Value = 0
    Else
        chkSourceNotAivdm.Value = 1
    End If
    txtNotMmsi.Text = SourceDuplicateFilter.RejectMmsi
    If SourceDuplicateFilter.RejectPayloadErrors = False Then
        chkRejectPayloadErrors.Value = 0
    Else
        chkRejectPayloadErrors.Value = 1
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    If UnloadMode = vbFormControlMenu Then    'clicked <X>
'(X) clicked - cancel unload
        MsgBox "Cancelled", , "Filter"
    End If
        
End Sub


Private Sub txtNotMmsi_Validate(Cancel As Boolean)
    If txtNotMmsi.Text <> "" Then
        If IsNumeric(txtNotMmsi.Text) = False _
        Or Len(txtNotMmsi) <> 9 Then
            MsgBox "MMSI must be numeric (9 digits)"
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

