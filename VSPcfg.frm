VERSION 5.00
Begin VB.Form VCPcfg 
   Caption         =   "Delete Virtual Com Port"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3795
   Icon            =   "VSPcfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox cboVCPName 
      Height          =   315
      ItemData        =   "VSPcfg.frx":058A
      Left            =   240
      List            =   "VSPcfg.frx":058C
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "VCPcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cancel As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim ret As Long
Dim Words() As String
Dim PortNo As Long
Dim kb As String
Dim VCPs() As String
Dim i As Long

'ensure we genuinely have port 0
    PortNo = -1
    If cboVCPName.Text <> "" Then
        Words = Split(cboVCPName.Text, " ")

        If IsNumeric(Mid(Words(0), 5)) Then
            PortNo = Mid(Words(0), 5)
        End If
    
        ret = MsgBox("Are you sure you wish to remove " & cboVCPName.Text, vbOKCancel)
    End If
    If ret <> vbOK Then
        MsgBox "Cancelled"
    Else
WriteLog "Removing VCP " & cboVCPName.Text
        Call DisableUAC
        kb = "cmd.exe /C ""setupc --output %Temp%\com0com.log remove " & PortNo & """"
        ret = ExecCmd(kb)
        Call AddFileToLog(Environ("Temp") & "\com0com.log")
        Call RestoreUAC
'Check to see if it has been removed
        VCPs = GetVCPs
        For i = 0 To UBound(VCPs)
            If VCPs(i) = Words(2) Then Exit For
        Next i
        If UBound(VCPs) = -1 Or i > UBound(VCPs) Then
            MsgBox cboVCPName.Text & " has been removed"
        Else
            MsgBox "Failed to remove " & cboVCPName.Text
        End If
        Unload Me
    End If

End Sub

Private Sub Form_Load()
Dim Key As String
Dim KeyCount As Long
Dim Keys() As Variant
Dim k As Long
Dim Subkey As String
Dim DeviceKey As String
Dim PortName As String
Dim PortA As String
Dim PortB As String
Dim PortNo As Long
Dim Words() As String

    PortNo = -1
    Key = "SYSTEM\CurrentControlSet\Enum\com0com\port"
'    If FriendlyName <> "" Then
'    FriendlyName = QueryValue(HKEY_LOCAL_MACHINE, SubKey, "FriendlyName")
    KeyCount = ReadKeys(HKEY_LOCAL_MACHINE, Key, Keys)
    If KeyCount > 0 Then
        For k = 0 To KeyCount - 1
            If Left$(Keys(k), 4) = "CNCA" Then
                PortA = Keys(k)
                DeviceKey = Key & "\" & Keys(k) & "\Device Parameters"
                DeviceKey = Replace(DeviceKey, "CNCA", "CNCB")
                PortB = QueryValue(HKEY_LOCAL_MACHINE, DeviceKey, "PortName")
                If IsNumeric(Mid(PortA, 5)) And PortB <> "" Then
                    PortNo = Mid(PortA, 5)
                cboVCPName.AddItem PortA & " <--> " & PortB
                End If
            End If
        Next k
    End If
    If cboVCPName.ListCount > 0 Then
        cboVCPName.ListIndex = 0  'default
    End If
End Sub

Public Sub VCPRemove()
Dim i As Long

    If cboVCPName.ListCount > 0 Then
    
        VCPcfg.Show vbModal
        Unload Me
    Else
        MsgBox "There as no VCP's to remove"
    End If
End Sub


