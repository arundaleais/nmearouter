VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3435
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSerialNo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   "SerialNo"
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Register Later"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtActivationCode 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register Now"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Registration is currently being tested"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Activation Code"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cancel As Boolean

Private SerialNo As String
Dim ActivationCode As String
'===================================

Private Sub cmdLater_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
Dim MailTo As String
Dim Subject As String
Dim Body As String
Dim ret As Long

    Select Case cmdRegister.Caption
    Case Is = "Register Now", "Contact Me"
        MailTo = MY_EMAIL_ADDRESS
        Subject = "NmeaRouter " & cmdRegister.Caption & " " & SerialNo
        Body = sysinfo
        ret = SendMail(MailTo, Subject, Body)
    Case Is = "Activate Now"
'Check Activation Code also when program loads
        If txtActivationCode.Text <> "" Then
            txtActivationCode.Text = UCase(txtActivationCode.Text)
'-1 = fail
            ModuleCode = CheckKey(txtActivationCode.Text)
            If ModuleCode > -1 Then
                ActivationKey = UCase(txtActivationCode.Text)
                MsgBox "Activation Code (" & ModuleCode & ") verified" & vbCrLf _
                & "You must Exit & Re-start NmeaRouter" & vbCrLf _
                & "to enable the Activation"
            Else
                MsgBox "Activation Failed"
            End If
        End If
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
Dim strMACAddress As String
Dim b() As Byte
Dim kb As String
    
'Uncomment 2 lines below to get a user the Activation Key
    
    txtSerialNo.Visible = jnasetup
    txtSerialNo.Enabled = jnasetup
    
    SerialNo = GetSerialNo()
    
    Label2.Caption = SerialNo
'Check registry for activation key
'If not registered key -1, -2 = cant get serial no
    If CheckKey(ActivationKey) < 0 Then
        cmdRegister.Caption = "Register Now"
    Else
        txtActivationCode = ActivationKey
        txtActivationCode.Enabled = jnasetup
        cmdRegister.Caption = "Contact Me"
    End If
        
End Sub
Public Function SendMail(MailTo As String, Subject As String, Body As String) As Long
Dim ret As Long
Dim i As Long

Body = Replace(Body, " ", "%20")
Body = Replace(Body, "  ", "%09")
Body = Replace(Body, vbCrLf, "%0D%0A")
'Debug.Print Len(Body)

    ret = ShellExecute(0, "open", REGISTER_SCRIPT _
    & "?mailto= " & MailTo _
    & "&subject= " & Subject _
    & "&body= " & Body _
    , vbNullString, vbNullString, SW_SHOW)
    If ret < 32 Then MsgBox "Send Email Failed"
    SendMail = ret
End Function

'http://www.codeproject.com/Articles/28550/Protecting-Your-Software-Using-Simple-Serial-Numbe
'f(x) = X2 + 53 / x + 113 * (x / 4)
Public Function GenerateKey(ByVal serial As Long) As String
Dim X As Long
    X = serial
    X = X + 53 / X + 113 * (X / 4)
GenerateKey = Convert10ToBase(X)
End Function


'http://www.freevbcode.com/ShowCode.asp?ID=6604
'Converts a number to any base ("0123456789ABCDEF") is hex
Public Function Convert10ToBase(ByVal d As Double) As String
    Dim S As String, tmp As Double, i As Integer, lastI As Integer
    Dim BaseSize As Integer
    BaseSize = Len(Encoding)
    i = 1   'if D = 0 then return 0
    Do While Val(d) <> 0
        tmp = d
        i = 0
        Do While tmp >= BaseSize
            i = i + 1
            tmp = tmp / BaseSize
        Loop
        If i <> lastI - 1 And lastI <> 0 Then S = S & String(lastI - i - 1, Left(Encoding, 1)) 'get the zero digits inside the number
        tmp = Int(tmp) 'truncate decimals
        S = S + Mid(Encoding, tmp + 1, 1)
        d = d - tmp * (BaseSize ^ i)
        lastI = i
    Loop
    S = S & String(i, Left(Encoding, 1)) 'get the zero digits at the end of the number
    Convert10ToBase = S
End Function

#If False Then
Public Function ConvertDecToBaseN_notused(ByVal dValue As Double, _
              Optional ByVal bybase As Byte = 16) As String
     
    Const BASENUMBERS As String = "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
     
    Dim sResult As String
    Dim dRemainder As Double
     
    On Error GoTo errorhandler
     
    sResult = ""
    If (bybase > 2) And (bybase < 37) Then
      dValue = Abs(dValue)
      Do
        dRemainder = dValue - (bybase * Int((dValue / bybase)))
        sResult = Mid$(BASENUMBERS, dRemainder + 1, 1) & sResult & ","
        dValue = Int(dValue / bybase)
      Loop While (dValue > 0)
    End If
    ConvertDecToBaseN = sResult
    Exit Function
     
errorhandler:
     
      err.Raise err.Number, "ConvertDecTobaseN", err.Description
       
    End Function

Function inverse_notused(i As Long) As Long
Dim sum As Long
Dim m As Long

    sum = 1
    m = 1 - i
    If i Mod (2) = 0 Then
        inverse = 0
        Exit Function
    End If
    
    Do
        sum = Int(sum) + m
'Debug.Print sum & ":" & m
    Loop Until m = m \ (1 - i)
    inverse = sum
End Function

Function encrypt_notused(n As Long, Key As Long) As Long
    encrypt = (n * Key) Mod 10
End Function

Function decrypt_notused(n As Long, Key As Long) As Long
    decrypt = n * inverse(Key)
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=4398
Public Function RC4_not_used(ByVal Expression As String, ByVal Password As String) As String
On Error Resume Next
Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
If Len(Password) = 0 Then
    Exit Function
End If
If Len(Expression) = 0 Then
    Exit Function
End If
If Len(Password) > 256 Then
    Key() = StrConv(Left$(Password, 256), vbFromUnicode)
Else
    Key() = StrConv(Password, vbFromUnicode)
End If
For X = 0 To 255
    RB(X) = X
Next X
X = 0
Y = 0
Z = 0
For X = 0 To 255
    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
    Temp = RB(X)
    RB(X) = RB(Y)
    RB(Y) = Temp
Next X
X = 0
Y = 0
Z = 0
ByteArray() = StrConv(Expression, vbFromUnicode)
For X = 0 To Len(Expression)
    Y = (Y + 1) Mod 256
    Z = (Z + RB(Y)) Mod 256
    Temp = RB(Y)
    RB(Y) = RB(Z)
    RB(Z) = Temp
    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
Next X
RC4 = StrConv(ByteArray, vbUnicode)
End Function


Public Function HexToBytes_notused(HexString As String) As Byte()
Dim i As Long
Dim p As Long
Dim b() As Byte
    ReDim b(Len(HexString) / 2 - 1)
    For i = 1 To Len(HexString) Step 2 'our fixed size string input.
        b(p) = CByte("&H" & Mid$(HexString, i, 2)) 'convert each pair of digits to a byte, store in the output array
        p = p + 1
    Next i
    HexToBytes = b
End Function

#End If

Private Sub txtActivationCode_Change()
    If Replace(txtActivationCode.Text, " ", "") = "" Then
        cmdRegister.Caption = "Register Now"
        cmdCancel.Caption = "Register Later"
    Else
        cmdRegister.Caption = "Activate Now"
        cmdCancel.Caption = "Cancel"
    End If
End Sub

'This generates an activation key and module code (separated by -)
'Only used when jna enters a user's serial no
Private Sub txtSerialNo_Validate(Cancel As Boolean)
    SerialNo = txtSerialNo    'enter the users serial no
    ActivationKey = GenerateKey(SerialNo) & "-" & Convert10ToBase(2 * 479) '479 is prime no
    txtActivationCode = ActivationKey
End Sub
