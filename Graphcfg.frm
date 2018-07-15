VERSION 5.00
Begin VB.Form Graphcfg 
   Caption         =   "Graph Configuration"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3495
   Icon            =   "Graphcfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUTC 
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtRange 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtPeriod 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "5"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "UTC Time"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum Range"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Update Interval (Mins)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Graphcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveGraphTimer As Boolean

Public Sub GraphOpen()
Dim kb As String

    txtPeriod = ExcelUpdateInterval
    txtRange = ExcelRange
    If ExcelUTC = True Then
        chkUTC = vbChecked
    Else
        chkUTC = vbUnchecked
    End If
    If SetMyChart = False Then
        chkUTC.Enabled = False
    Else
        chkUTC.Enabled = True
    End If
    Me.Show vbModal
kb = frmRouter.GraphTimer.Enabled
    If ExcelUpdateInterval > 0 _
    And frmRouter.MenuViewInoutGraph.Checked = True Then
        If ExcelOpen = False Then
            Call CreateWorkbook
        End If
        frmRouter.GraphTimer.Enabled = True
    End If

End Sub

Private Sub chkUTC_click()
    If chkUTC = vbGrayed Then chkUTC = vbUnchecked
End Sub

Private Sub Form_Load()
    SaveGraphTimer = frmRouter.GraphTimer.Enabled
    frmRouter.GraphTimer.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRouter.GraphTimer.Enabled = SaveGraphTimer
End Sub

Private Sub txtPeriod_Validate(Cancel As Boolean)
    If IsNumeric(txtPeriod) = False Then
        txtPeriod = "0"
        MsgBox "Update Interval must be numeric"
        Cancel = True
    Else
        Select Case CLng(txtPeriod)
        Case Is < 0
            MsgBox "Update Interval cannot be less than 0"
            Cancel = True
        End Select
    End If
End Sub

Private Sub txtRange_Validate(Cancel As Boolean)
    If IsNumeric(txtRange) = False Then
        txtRange = "0"
        MsgBox "Maximum Range must be numeric"
        Cancel = True
    Else
        Select Case CLng(txtRange)
        Case Is < 2
            MsgBox "Maximum Range cannot be less than 2"
            Cancel = True
        End Select
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim Cancel As Boolean
Dim arry() As String
   
'Create Winsock closes existing socket and allocate Hidx (if req)
'Either connection will validate
        Call txtPeriod_Validate(Cancel)
        If Cancel = True Then Exit Sub
        Call txtRange_Validate(Cancel)
        If Cancel = True Then Exit Sub
'Load new settings into Excel
        ExcelUpdateInterval = CLng(txtPeriod.Text)
        ExcelRange = CLng(txtRange.Text)
        If chkUTC = vbChecked Then
            ExcelUTC = True
        Else
            ExcelUTC = False
        End If
    Unload Me
End Sub

