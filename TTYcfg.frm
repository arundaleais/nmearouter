VERSION 5.00
Begin VB.Form TTYcfg 
   Caption         =   "Display Configuration"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "TTYcfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Direction"
      Height          =   885
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton optDirection 
         Caption         =   "Output"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Input,Output"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "TTYcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Hidx As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Sockets(CurrentSocket).Hidx > 0 Then
        Hidx = Sockets(CurrentSocket).Hidx
'Disable Comm
'        Call TTYs(Hidx).CloseTTY
    Else
        Hidx = FreeTTY
'Set defaults
'Direction is got from initial setup on frmTTY
        If Sockets(CurrentSocket).Direction = -1 Then
            Sockets(CurrentSocket).Direction = CurrentDirection
        End If
    End If
    
    If Hidx = -1 Then
        MsgBox "No free TTY Sockets", , "TTYcfg.Load"
        Exit Sub
    End If
    
'Set Direction option to same as in Sockets()
    optDirection(Sockets(CurrentSocket).Direction).value = True
    
End Sub

Private Sub CmdOk_Click()
Dim Cancel As Boolean
   
'Validate any changed
        
'Load new settings into Sockets()
        With Sockets(CurrentSocket)
            .Direction = CurrentDirection
        End With
    Unload Me
End Sub

'Returns the Direction from optDirection
Private Function CurrentDirection() As Long
Dim i As Integer
    For i = 0 To optDirection.UBound
        If optDirection(i).value = True Then
            CurrentDirection = i
            Exit For
        End If
    Next i
End Function

