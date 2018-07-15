VERSION 5.00
Begin VB.Form Profilecfg 
   Caption         =   "Profile"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   Icon            =   "Profilecfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox cboProfile 
      Height          =   315
      ItemData        =   "Profilecfg.frx":058A
      Left            =   1440
      List            =   "Profilecfg.frx":058C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Profilecfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Existing Profiles are always created when form is loaded
Dim Profiles() As Variant
Dim ProfileCount As Long
Dim Cancel As Boolean

Private Sub cboProfile_Validate(Cancel As Boolean)
Dim arry() As String
Dim Ret As Integer

'Dont allow blank profile or none selected
    
    If cboProfile.Text = "" Then
        MsgBox "Please enter a profile"
        Cancel = True
        Exit Sub
    End If
    
    arry = Split(Me.Caption)
    Select Case arry(0)
    Case Is = "New"
'Check Profile is New
        If ProfileExists(cboProfile.Text) = True Then
            MsgBox cboProfile.Text & " already exists"
            Cancel = True
        End If
    Case Is = "Open"
'Check if Profile exists
        If ProfileExists(cboProfile.Text) = False Then
            MsgBox cboProfile.Text & " does not exist", , Me.Caption
            Cancel = True
            Exit Sub
        End If
    Case Is = "Save"
        If CurrentProfile <> cboProfile.Text Then
            Ret = MsgBox(CurrentProfile & " differs from " & cboProfile.Text, vbOKCancel, "Save Profile")
            If Ret = vbCancel Then
                Exit Sub
            End If
        End If
    Case Is = "SaveAs"
        If ProfileExists(cboProfile.Text) = True Then
            Ret = MsgBox(cboProfile.Text & " already exists and will be replaced", vbOKCancel, "Save Profile As")
            If Ret = vbCancel Then
                Exit Sub
            End If
        End If
    Case Is = "Delete"
        If ProfileExists(cboProfile.Text) = False Then
            MsgBox cboProfile.Text & " does not exist", vbExclamation, "Delete Profile"
            Cancel = True
            Exit Sub
        End If
        If cboProfile.Text = CurrentProfile Then
            Ret = MsgBox(cboProfile.Text & " is the Current Profile in use" & vbCrLf & "and will be deleted", vbOKCancel, "Delete Profile")
            If Ret = vbCancel Then
                Exit Sub
            End If
        End If
    Case Else
        MsgBox "Invalid Option " & arry(0)
    End Select
        
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Idx As Long
Dim arry() As String

    If cboProfile.Text = "" Then
        MsgBox "Please enter a profile"
        Exit Sub
    End If
    
    If Cancel = False Then
        arry = Split(Me.Caption)
        Select Case arry(0)
        Case Is = "New"
            If cboProfile.Text <> CurrentProfile Then
'We need to keep the Current Profile as well in case we want to save it
'                frmRouter.NewProfile = cboProfile.Text
                Call CloseProfile(CurrentProfile)
                CurrentProfile = cboProfile.Text
                Call ResetRouter
            End If
        Case Is = "Open"
            If cboProfile.Text <> CurrentProfile Then
'We need to keep the Current Profile as well in case we want to save it
'                frmRouter.NewProfile = cboProfile.Text
                Call CloseProfile(CurrentProfile)
                CurrentProfile = cboProfile.Text
                Call ResetRouter
            End If
        Case Is = "Save"
'Remove the "Old" file so that you are not prompted if the file has changed
            On Error Resume Next
            Kill TempPath & "Router" & cboProfile.Text & ".txt"
            On Error GoTo 0
            Call SaveProfile(cboProfile.Text)
    
        Case Is = "SaveAs"
            Call SaveProfile(cboProfile.Text)
            CurrentProfile = cboProfile.Text
            
        Case Is = "Delete"
            If cboProfile.Text = CurrentProfile Then
                Call CloseProfile(CurrentProfile)
            End If
            Call DeleteProfile(cboProfile.Text)
'            frmRouter.NewProfile = ""
            CurrentProfile = ""
            Call ResetRouter
'Delete profile and all subkeys in registry
'Ask for new profile
        Case Else
            MsgBox "Invalid Option " & arry(0)
        End Select
    
        Call frmRouter.SetCaption
    End If
    Unload Me

'will be invisible if profile is being changed and graph in use
    frmRouter.Visible = True    'Causes a resize
'If in systray it will still be no visible
'Move in front of graph
    If frmRouter.Visible = True Then
        frmRouter.SetFocus      'error if not visible
    End If
    End Sub
    
'Load List of Profiles in cboProfiles
'Set List Index to the Current profile
Private Sub Form_Load()
    Call GetProfiles(ProfileCount, Profiles)
End Sub

'Load List of Profiles in cboProfiles
'Set List Index to the Current profile
Public Sub GetProfiles(ProfileCount As Long, arProfiles() As Variant)
Dim ProfilesKey As String
Dim i As Long
    
    ProfilesKey = ROUTERKEY & "\" & "Profiles"
    ProfileCount = ReadKeys(HKEY_CURRENT_USER, ProfilesKey, arProfiles)
End Sub

Private Function ProfileExists(Profile As String) As Boolean
Dim i As Long
    If ProfileCount > 0 Then
        For i = 0 To UBound(Profiles)
            If StrComp(Profiles(i), Profile, vbTextCompare) = 0 Then
                ProfileExists = True
                Exit Function
            End If
        Next i
    End If
End Function
Public Sub ProfileNew()
Dim i As Long

    Me.Caption = "New Profile"
'    Call GetProfiles(ProfileCount, Profiles)
    Do
        i = i + 1
    Loop Until ProfileExists("Profile" & i) = False
    cboProfile.AddItem "Profile" & i
    cboProfile.ListIndex = 0
    Profilecfg.Show vbModal
    Unload Me
End Sub

Public Sub ProfileOpen()
Dim i As Long
    Me.Caption = "Open Profile"
'    Call GetProfiles(ProfileCount, Profiles)
    If ProfileCount > 0 Then
        For i = 0 To UBound(Profiles)
            cboProfile.AddItem Profiles(i)
        Next i
    End If
'Set listindex to current profile
    For i = 0 To cboProfile.ListCount - 1
        If cboProfile.List(i) = CurrentProfile Then
            cboProfile.ListIndex = i
        End If
    Next i
'If no current profile set to first profile
    If cboProfile.ListIndex = -1 And cboProfile.ListCount > 0 Then
        cboProfile.ListIndex = 0
    End If
    Profilecfg.Show vbModal
    Unload Me
End Sub

Public Sub ProfileSave()
'    Call GetProfiles(ProfileCount, Profiles)
    Me.Caption = "Save Profile [" & CurrentProfile & "]"
    cboProfile.AddItem CurrentProfile
    cboProfile.ListIndex = 0
    cboProfile.Enabled = False
    Profilecfg.Show vbModal
    Unload Me
End Sub

Public Sub ProfileSaveAs()
'    Call GetProfiles(ProfileCount, Profiles)
    cboProfile.AddItem "Profile" & ProfileCount + 1
    cboProfile.ListIndex = 0
    Me.Caption = "SaveAs Profile"
    Profilecfg.Show vbModal
    Unload Me
End Sub

Public Sub ProfileDelete()
Dim i As Long

'v59a    Me.Caption = "Delete Profile [" & CurrentProfile & "]"
    Me.Caption = "Delete Profile"    'v59a
    If ProfileCount > 0 Then
        For i = 0 To UBound(Profiles)
            cboProfile.AddItem Profiles(i)
        Next i
'Set listindex to last profile
    cboProfile.ListIndex = ProfileCount - 1
    End If
'Set listindex to current profile
'    For i = 0 To cboProfile.ListCount - 1
'        If cboProfile.List(i) = CurrentProfile Then
'            cboProfile.ListIndex = i
'        End If
'    Next i
    Profilecfg.Show vbModal
    Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    frmRouter.PollTimer.Enabled = True
    Call frmRouter.SetCaption
End Sub
