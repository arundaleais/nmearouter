VERSION 5.00
Begin VB.Form frmTTY 
   Caption         =   "Display"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TermText 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmTTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Idx As Long

'If X clicked as close the handler
'Note Re-connect timer will not try and re-open
    If UnloadMode = vbFormControlMenu Then
        Idx = DevNameToSocket(Caption)
        Call CloseHandler(Idx)
    End If
End Sub

Private Sub Form_Resize()
    TermText.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

'This is how you get input to a form
Private Sub txtTerm_KeyPress_old(KeyAscii As Integer)
    If Not (Comms(0) Is Nothing) Then
        Comms(0).CommOutput (Chr$(KeyAscii))
    End If
    KeyAscii = 0
End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
'Debug.Print Chr$(KeyAscii) & DevNameToSocket(Me.Caption)
'The keysrokes require keeping in an input buffer
'That is held for this Display(Socket) until a CRLF is detected
'Then it should be treated as an Input & therefore forwarded
'As required to any relevant sockets.
End Sub

