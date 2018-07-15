VERSION 5.00
Begin VB.Form Routecfg 
   Caption         =   "Route Configuration"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   Icon            =   "Routecfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnabled 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox cboDevName 
      Height          =   315
      Index           =   1
      ItemData        =   "Routecfg.frx":058A
      Left            =   1440
      List            =   "Routecfg.frx":058C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox cboDevName 
      Height          =   315
      Index           =   0
      ItemData        =   "Routecfg.frx":058E
      Left            =   1440
      List            =   "Routecfg.frx":0590
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "And"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Between"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Routecfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancel As Boolean
Dim SelectedDevName(2) As String
Dim LastMsg As String
Dim Resetting As Boolean

Private Sub cboDevName_Click(Index As Integer)
'Because cbodevname().listindex= causes a click event
'and this is used in RouteToEditFields to set the other list
'default in the text box, we must not recall reset list
'otherwise we get a stack overflow
    If Me.Visible And Resetting = False Then
'May need to call Validate
        Call RouteToEditFields(OtherIndex(Index))
    End If
End Sub

Private Sub cmdCancel_Click()
Cancel = False
Unload Me
End Sub

'Called from frmRouter(Menu) -
Public Sub RouteNew()
Dim ret As Boolean

    Me.Caption = "New Route"
'The form is now loaded but has no default route set
'Get the first free route available
    ret = RouteToEditFields(1)
    If ret = True Then
        Me.Show vbModal
    Else
        MsgBox "There are no Routes available", , "RouteNew"
    End If
End Sub

'Called from frmRouter(Menu) -
Public Sub RouteOpen(Idx1 As Long, Idx2 As Long)
Dim ret As Boolean

    Me.Caption = "Open Route"
'The form is now loaded but has no valid routes set
'Get the first free route available
    ret = RouteToEditFields(1, Idx1, Idx2)
    If ret = True Then
        
        Me.Show vbModal
    Else
        MsgBox "There are no Routes available", , "RouteNew"
    End If
End Sub

'Called from frmRouter(Menu) only enabled if there are routes
Public Sub RouteDelete(Idx1 As Long, Idx2 As Long)
Dim ret As Boolean

    Me.Caption = "Delete Route"
'The form is now loaded but has no default route set
'Get the first free route available
    ret = RouteToEditFields(1, Idx1, Idx2)
    If ret = True Then
        Me.Show vbModal
    Else
        MsgBox "There are no Routes available", , "RouteNew"
    End If
End Sub

'Index is the cboDevname(Index).List we are constructing given the
'Other Index (ie the Index we have already set)
'Returns false if no valid routes
Public Function RouteToEditFields(Index As Integer, Optional ReqIdx1 As Long, Optional ReqIdx2 As Long) As Boolean

Dim Idx1 As Long
Dim Idx2 As Long
Dim SavedDevName2 As String
Dim kb As String
Dim arry() As String
Dim i As Long
Dim ValidRouteCount As Long
Dim List1() As Boolean   'True if Valid as Idx1
Dim List2() As Boolean  'True if Valid as Idx2
Dim DevName1 As String
Dim DevName2 As String
Dim Silent As Boolean

'Dont allow a recursive call if/when click event triggered within this routine
    If Resetting = True Then Exit Function
'Stop setting ListIndex calling RouteToEditFields
    Resetting = True
'Silent=do not display any error message
    Silent = True
    arry = Split(Me.Caption)
'Testing Validation
'Call AllToList
'All Combinations to list

'   If both lists are empty re-create both lists
    If cboDevName(0).ListCount = 0 And cboDevName(1).ListCount = 0 Then
    
        ReDim List1(1 To UBound(Sockets))
        ReDim List2(1 To UBound(Sockets))
        For Idx1 = 1 To UBound(Sockets)
            For Idx2 = Idx1 + 1 To UBound(Sockets)
            Select Case arry(0)
            Case Is = "New"
                If ValidateNewRoute(Idx1, Idx2) = True Then
                    List1(Idx1) = True
                    List2(Idx2) = True
                End If
            Case Is = "Open", "Delete"
'Idx1,Idx2 nust not be sorted
                If RouteExists(Idx1, Idx2) <> -1 Then
                    List1(Idx1) = True
                    List2(Idx2) = True
                End If
            End Select
            Next Idx2
        Next Idx1
'You have to determine which indexes are valid before the
'list is loaded to avoid duplicating entries in the list
'Now Add the IDX into cboDevname Lists using List1 & List2
        For Idx1 = 1 To UBound(Sockets)
            If List1(Idx1) = True Then
                cboDevName(0).AddItem Sockets(Idx1).DevName
                If Idx1 = ReqIdx1 Then
                    cboDevName(0).ListIndex = cboDevName(0).ListCount - 1
                End If
            End If
        Next Idx1
        For Idx2 = 1 To UBound(Sockets)
            If List2(Idx2) = True Then
                cboDevName(1).AddItem Sockets(Idx2).DevName
                If Idx2 = ReqIdx2 Then
                    cboDevName(1).ListIndex = cboDevName(1).ListCount - 1
                End If
            End If
        Next Idx2
'Note Idx1 will be the lower index
        If cboDevName(0).ListIndex = -1 And cboDevName(0).ListCount > 0 Then
            cboDevName(0).ListIndex = 0
        End If
        
        If cboDevName(1).ListIndex = -1 And cboDevName(1).ListCount > 0 Then
            cboDevName(1).ListIndex = 0
        End If
        If cboDevName(0).ListCount > 0 And cboDevName(1).ListCount > 0 Then
            RouteToEditFields = True    'we have a valid list
        Else
'We do not have any Routes (to Add or Delete)
'            Exit Function dont exit without setting resetting to false
        End If
    End If
    DevName1 = cboDevName(0).List(cboDevName(0).ListIndex)
    DevName2 = cboDevName(1).List(cboDevName(1).ListIndex)
    Idx1 = DevNameToSocket(DevName1)
    Idx2 = DevNameToSocket(DevName2)
    If RouteEnabled(Idx1, Idx2) = True Then
        chkEnabled = vbChecked
    Else
        chkEnabled = vbUnchecked
    End If
    Resetting = False
End Function
 
'Checks if a given route may be added
'We have to keep any error msg from Forward as only 1 error is OK
'If Silent any error message is NOT displayed
Public Function ValidateNewRoute(Idx1 As Long, Idx2 As Long) As Boolean
Dim ReqForwardCount As Long
Dim ForwardErrMsg(2) As String    'From ValidateAddForward
Dim msg As String

'Check both Indexes are valid
    If Idx1 < 1 Or Idx1 > UBound(Sockets) Then
        msg = "Socket(" & Idx1 & ") is invalid" & vbCrLf
        GoTo ValidateNewRoute_Error
    End If
    If Idx2 < 1 Or Idx2 > UBound(Sockets) Then
        msg = "Socket(" & Idx2 & ") is invalid" & vbCrLf
        GoTo ValidateNewRoute_Error
    End If

'Check both sockets are in use
    If Sockets(Idx1).State = -1 Then
        msg = msg & Sockets(Idx1).DevName & " is not Open" & vbCrLf
        GoTo ValidateNewRoute_Error
    End If
    If Sockets(Idx2).State = -1 Then
        msg = msg & Sockets(Idx2).DevName & " is not Open" & vbCrLf
        ValidateNewRoute = False
    End If
        
'Both Sockets cant be the same
'They can be if a loopback
'    If Idx1 = Idx2 Then
'        Msg = Msg & "You cannot Route to the same Connection (" _
'        & Sockets(Idx1).DevName & ")" & vbCrLf
'        GoTo ValidateNewRoute_Error
'    End If
    
'RouteExists checks both Directions
'Idx1,Idx2 doesnt matter if sorted
    If RouteExists(Idx1, Idx2) <> -1 Then
            msg = msg & "Route " & Sockets(Idx1).DevName _
            & " to " & Sockets(Idx2).DevName & " already exists" & vbCrLf
        GoTo ValidateNewRoute_Error
        End If
    
    ValidateNewRoute = True
Exit Function

'You cannot display errors in this routine
'because it validate routes when deciding whther they need
'including in the list af valid routes
ValidateNewRoute_Error:
'        MsgBox Msg, , "Validate New Route"
        LastMsg = msg
End Function

Private Sub cmdOK_Click()
Dim DevName1 As String
Dim DevName2 As String
Dim Idx1 As Long
Dim Idx2 As Long
Dim Ridx As Long
Dim Enabled As Boolean
Dim arry() As String
Dim Silent As Boolean
Dim kb As String

Silent = False
'We need the Devname as AddRoute must use Davname as the Argument
'because we need to be ale to load the routes from the registry
    DevName1 = cboDevName(0).List(cboDevName(0).ListIndex)
    DevName2 = cboDevName(1).List(cboDevName(1).ListIndex)
    Idx1 = DevNameToSocket(DevName1)
    Idx2 = DevNameToSocket(DevName2)
    Ridx = RouteExists(Idx1, Idx2)
'    If chkEnabled = chkEnabled Then
'        Enabled = True
'    End If
    arry = Split(Me.Caption)
    Select Case arry(0)
    Case Is = "New"
'Idx1,Idx2 must be sorted because Ridx pertains to lower Idx
        Ridx = CreateRoute(Idx1, Idx2)
'Will be -1 if both sockets are input
        If Ridx < 1 Then
            Exit Sub
        End If
'Create Route does not Add the forwards because we need to
'establish first if the route is enabled
        Sockets(Idx1).Routes(Ridx).Enabled = BooleanChecked(chkEnabled)
'If new Route no forwards can exist
        If Sockets(Idx1).Routes(Ridx).Enabled = True Then
            Call CreateRouteForwards(Idx1, Idx2)
        End If
   
    Case Is = "Open"
'Ridx = -1 if route is inactive becauae both sockets are input
        If Ridx < 1 Then
            Exit Sub
        End If
'If we have changed the enabled we must either create or remove
'the forwards
        If Sockets(Idx1).Routes(Ridx).Enabled <> BooleanChecked(chkEnabled) Then
            Sockets(Idx1).Routes(Ridx).Enabled = BooleanChecked(chkEnabled)
            If Sockets(Idx1).Routes(Ridx).Enabled = True Then
                Call CreateRouteForwards(Idx1, Idx2)
            Else
                Call RemoveRouteForwards(Idx1, Idx2)
            End If
        End If
    
    Case Is = "Delete"
'Ridx = -1 if route is inactive becauae both sockets are input
        If Ridx < 1 Then
            Exit Sub
        End If
        If Sockets(Idx1).Routes(Ridx).Enabled = True Then
            Call RemoveRouteForwards(Idx1, Idx2)
        End If
        Call RemoveRoute(Idx1, Idx2)    'Also Removes forwards
        If Cancel = True Then
            MsgBox "Cannot Remove Route between " & DevName1 _
            & " and " & DevName2, , Me.Caption
            Exit Sub
        End If
    Case Else
        MsgBox "Invalid in CmdOK"
'        Stop
        Exit Sub
    End Select
    
    kb = InactiveRoute(Idx1, Idx2)
    Unload Me
    If kb <> "" Then
        kb = "This Route is inactive because both Connections are " & aDirection(Sockets(Idx1).Direction) & vbCrLf & kb
        MsgBox kb, , "Inactive Route"
    End If
End Sub

'If there are Routes to this sockets Queries the Socket Delete
Public Sub InactiveSocketRoutes(ReqIdx As Long)
Dim Idx As Long
Dim Ridx As Long
Dim kb As String
Dim Count As Long
Dim SocketCount As Long
Dim RidxCount As Long
Dim ret As Integer
Dim Detail As Boolean   'Set True to display details
Dim msg As String

    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
        For Ridx = 1 To UBound(.Routes)
            If ReqIdx = 0 Or Idx = ReqIdx Or .Routes(Ridx).AndIdx = ReqIdx Then
                If .Routes(Ridx).AndIdx > 0 Then
                    msg = InactiveRoute(Idx, .Routes(Ridx).AndIdx)
                    If msg <> "" Then
                        RidxCount = RidxCount + 1
                        kb = kb & msg
                    End If
                End If
            End If
        Next Ridx
        End With
    Next Idx
    If RidxCount > 0 Then
        If RidxCount > 1 Then
           kb = "These Routes are inactive because both Connections are Input or Output only" & vbCrLf & kb
        Else
           kb = "This Route is inactive because both Connections are Input or Output only" & vbCrLf & kb
        End If
        MsgBox kb, , "Inactive Routes"
    End If
End Sub

Public Function InactiveRoute(Idx1 As Long, Idx2 As Long) As String
Dim kb As String

'Report to user if No Active route unless direction is changed
    If Sockets(Idx1).Enabled = True And Sockets(Idx2).Enabled = True Then
        If Sockets(Idx1).Direction = Sockets(Idx2).Direction Then
            If Sockets(Idx1).Direction <> 0 Then
                kb = kb & "Between " & Sockets(Idx1).DevName & " [" _
                & aDirection(Sockets(Idx1).Direction) & "]"
                kb = kb & " and " & Sockets(Idx2).DevName _
                & " [" & aDirection(Sockets(Idx2).Direction) & "]"
                kb = kb & vbCrLf
            End If
        End If
    End If
    InactiveRoute = kb
End Function

'Will delete the Route
'even if enabled
Public Sub RemoveRoute(Idx1 As Long, Idx2 As Long)
Dim Ridx As Integer
Dim Fidx As Long

    Call SortIdx(Idx1, Idx2)
    If RouteExists(Idx1, Idx2) = -1 Then
        Exit Sub
    End If
        
    For Ridx = 1 To UBound(Sockets(Idx1).Routes)
        If Sockets(Idx1).Routes(Ridx).AndIdx = Idx2 Then
            WriteLog "Removed Route between " & Sockets(Idx1).DevName _
            & " and " & Sockets(Idx2).DevName
            Sockets(Idx1).Routes(Ridx).AndIdx = -1
            Exit For    'This route removed
        End If
    Next Ridx
'Check if another route can be set up
'If not disable New on popup menu
'Done here because the Click event noes not always do it
    frmRouter.MenuRouteNew.Enabled = IsNewRoute


End Sub

'Returns Ridx or -1 if it doesnt exist
Public Function RouteExists(Idx1 As Long, Idx2 As Long) As Long
Dim Ridx As Long

    Call SortIdx(Idx1, Idx2)
'Must have valid Indexes
    If Idx1 < 1 Or Idx2 < 1 Then
        RouteExists = -1
        Exit Function
    End If
    
'If Routes().AndIdx is 0, -1 is still returned
    For Ridx = 1 To UBound(Sockets(Idx1).Routes)
        If Sockets(Idx1).Routes(Ridx).AndIdx = Idx2 Then
            RouteExists = Ridx
            Exit Function    'This destination exists
        End If
    Next Ridx
    RouteExists = -1
'Note Idx1 and Idx2 are not swapped on the return to the calling
'Subroutine call is (ByVal)
End Function

'Returns if a route is enabled or disabled
Public Function RouteEnabled(Idx1 As Long, Idx2 As Long) As Boolean
Dim Ridx As Long

    Call SortIdx(Idx1, Idx2)
'Must have valid Indexes
    If Idx1 < 1 Then
        RouteEnabled = False
        Exit Function
    End If
    
    For Ridx = 1 To UBound(Sockets(Idx1).Routes)
        If Sockets(Idx1).Routes(Ridx).AndIdx = Idx2 Then
            RouteEnabled = Sockets(Idx1).Routes(Ridx).Enabled
            Exit For    'Got the route
        End If
    Next Ridx
'Note Idx1 and Idx2 are not switched on the return to the calling
'Subroutine call is (ByVal)
End Function

Private Function ForwardExists(ByVal Source As Long, ByVal Destination As Long) As Boolean
Dim Ridx As Long

    For Ridx = 1 To UBound(Sockets(Source).Forwards)
        If Sockets(Source).Forwards(Ridx) = Destination Then
            ForwardExists = True
            Exit For    'This destination exists
        End If
    Next Ridx
End Function

Private Function OtherIndex(ThisIndex As Integer) As Integer
'OptionsListIndex is the one we are moving to
    If ThisIndex = 0 Then
        OtherIndex = 1
    Else
        OtherIndex = 0
    End If
End Function

Private Sub Form_Load()

'Wants moving to cmdOK (Dont need to stop timer yet)
    frmRouter.PollTimer.Enabled = False
    
'At this point Caption is not set so you cannot
'use arry() to determine what routine has called
'Routecfg
        
'Not required
'    Call RouteToEditFields(1)
'We assume it is a New Route
'This will be reset to select a route after the form
'is loaded
    
End Sub

'This will add all possible combinations to the list
'without duplications. USED FOR TESTING VALIDATION
Private Sub AllToList()
Dim Idx1 As Long
Dim Idx2 As Long

    For Idx1 = 1 To UBound(Sockets)
        cboDevName(0).AddItem Sockets(Idx1).DevName
'Idx2 is the last Idx2 added to the list (do not duplicate)
        If Idx1 > Idx2 Then
            cboDevName(1).AddItem Sockets(Idx1).DevName
            Idx2 = Idx1
        End If
    Next Idx1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmRouter.PollTimer.Enabled = True

End Sub


Public Sub SortIdx(ByRef Idx1 As Long, ByRef Idx2 As Long)
Dim Temp As Long
    If Idx1 > Idx2 Then
        Temp = Idx1
        Idx1 = Idx2
        Idx2 = Temp
    End If
End Sub

'called by Routecfg OK "New" and LoadProfile
'Idx1 must be as the lower because The returned Hidx applies to
'Idx1 & not Idx2
'Returns -1 if it cant be created
Public Function CreateRoute(ByRef Idx1 As Long, ByRef Idx2 As Long) As Long
Dim Ridx As Long

WriteLog "Creating Route between " & Sockets(Idx1).DevName & " and " & Sockets(Idx2).DevName
'Get lowest index first
    Call SortIdx(Idx1, Idx2)
    If ValidateNewRoute(Idx1, Idx2) = False Then
        CreateRoute = -1
        Exit Function
    End If
    Ridx = FreeRoute(Idx1)
'Call frmRouter.DisplayRoutes
    If Ridx > 0 Then
        Sockets(Idx1).Routes(Ridx).AndIdx = Idx2
        CreateRoute = Ridx
    Else
        Call frmDpyBox.DpyBox("No Free Routes", 10, "CreateRoute")
    End If
'Check if another route can be set up
'If not disable New on popup menu
'Done here because the Click event does not always do it
    frmRouter.MenuRouteNew.Enabled = IsNewRoute
'The calling routine should then set Enabled or Disabled
'THEN call RouteForwards to set up the forwards
End Function

'Creates Forwards for any Enabled Route using this socket
'if enabled
'Called by cmdStart,ToggleSocketEnabled and Socketcfg.cmdOK
'CmdStart cycles through all enabled sockets
Public Function CreateSocketForwards(ReqIdx As Long)
Dim Ridx As Long
Dim Idx1 As Long
Dim Idx2 As Long
Dim Fidx As Long
Dim ReqOidx As Long

'If a Stream socket the Route will be on the Owner socket
'Use the Owner index to find the route
    If IsTcpStream(ReqIdx) Then
        ReqOidx = Sockets(ReqIdx).Winsock.Oidx
    Else
        ReqOidx = ReqIdx
    End If
    
'Routes are always TCP Listeners
    For Idx1 = 1 To UBound(Sockets)
        For Ridx = 1 To UBound(Sockets(Idx1).Routes)
            Idx2 = Sockets(Idx1).Routes(Ridx).AndIdx
            If Idx2 > 0 And (Idx1 = ReqOidx Or Idx2 = ReqOidx) Then
'Found Route to/from this socket
'The Route containing this socket must be enabled
                If Sockets(Idx1).Routes(Ridx).Enabled = True Then
'Both Sockets must be enabled
                    If Sockets(Idx2).Enabled = True Then
'Weve got a route with this Idx
'Both will be Owner Indexes (TCP Listers - never Streams)
                        Call CreateRouteForwards(Idx1, Idx2)
                    End If
                End If  'Valid Route
            End If  'Route could be 0 or not this socket
        Next Ridx
    Next Idx1
End Function

'Removes any forwards to this socket
'Called if socket is removed or disabled
Public Function RemoveSocketForwards(ReqIdx As Long)
Dim Fidx As Long
Dim Idx As Long
Dim Sidx As Long

    For Idx = 1 To UBound(Sockets)
        For Fidx = 1 To UBound(Sockets(Idx).Forwards)
            If Idx = ReqIdx Or Sockets(Idx).Forwards(Fidx) = ReqIdx Then
                If Sockets(Idx).Forwards(Fidx) > 0 Then
                    WriteLog "Removed Forward " & Sockets(Idx).DevName _
                    & " to " & Sockets(Sockets(Idx).Forwards(Fidx)).DevName
'Removes forwards in both directions
                    Call DecrForward(Idx, Sockets(Idx).Forwards(Fidx), False)
                    Call DecrForward(Idx, Sockets(Idx).Forwards(Fidx), True)
                    Sockets(Idx).Forwards(Fidx) = 0
                End If
            End If
        Next Fidx
    Next Idx
    
'Scan All the TCP Server streams and remove any that will
'be forwarded to the socket to which we are removing forwards
    For Idx = 1 To UBound(Sockets)
        For Sidx = 1 To UBound(Sockets(Idx).Winsock.Streams)
'Remove any TCP Server streams TO the Req Socket socket
            If Sockets(Idx).Winsock.Streams(Sidx) = ReqIdx Then
                    WriteLog "Removed TCP Server Forward " & Sockets(Idx).DevName _
                    & " to " & Sockets(Sockets(Idx).Winsock.Streams(Sidx)).DevName
                    Sockets(Idx).Winsock.Streams(Sidx) = 0
            End If
        Next Sidx
    Next Idx

End Function


'The Route must exist - but can be enabled or disabled
'Called if route is deleted or disabled
Public Sub RemoveRouteForwards(Idx1 As Long, Idx2 As Long)
Dim Ridx As Integer
Dim Fidx As Long

    Call SortIdx(Idx1, Idx2)
    If RouteExists(Idx1, Idx2) = -1 Then
        Exit Sub
    End If
    
    For Ridx = 1 To UBound(Sockets(Idx1).Routes)
        If Sockets(Idx1).Routes(Ridx).AndIdx = Idx2 Then
'Found the route between Idx1 and Idx2
'Have we any Forwards from Idx1 to Idx2
            For Fidx = 1 To UBound(Sockets(Idx1).Forwards)
                If Sockets(Idx1).Forwards(Fidx) = Idx2 Then
                    WriteLog "Removed Forward " & Sockets(Idx1).DevName _
                    & " to " & Sockets(Sockets(Idx1).Forwards(Fidx)).DevName
                    Sockets(Idx1).Forwards(Fidx) = 0
Call DecrForward(Idx1, Idx2, False)
                End If
            Next Fidx
'Have we any (reverse) Forwards from Idx2 to Idx1
            For Fidx = 1 To UBound(Sockets(Idx2).Forwards)
                If Sockets(Idx2).Forwards(Fidx) = Idx1 Then
                    WriteLog "Removed Forward " & Sockets(Idx2).DevName _
                    & " to " & Sockets(Sockets(Idx2).Forwards(Fidx)).DevName
                    Sockets(Idx2).Forwards(Fidx) = 0
Call DecrForward(Idx1, Idx2, True)
                End If
            Next Fidx
        End If
    Next Ridx

End Sub

'Creates the forwards from an Enabled Route
'The Route must exist, both sockets and the Route have been enabled
'Called by ToggleRouteEnabled,Routecfg.cmsOK
'and CreateSocketForwards
Public Sub CreateRouteForwards(Idx1 As Long, Idx2 As Long)
Dim Ridx As Integer
'Dim Fidx As Long
Dim errmsg As String
    
'Idx1 and Idx2 will both be Owner Indexes (Listener if TCP)
    Call SortIdx(Idx1, Idx2)
    
'If either socket is a TCP Stream do not check the routes
'as there wont be one
    If IsTcpStream(Idx1) Or IsTcpStream(Idx2) Then
'Idx1 & Idx2 should never be stream sockets
        errmsg = "Route between " & Sockets(Idx1).DevName _
        & " and " & Sockets(Idx2).DevName & " not permitted"
        GoTo CreateRouteForwards_error
    Else
        If RouteExists(Idx1, Idx2) = -1 Then
            errmsg = "Route between " & Sockets(Idx1).DevName _
            & " and " & Sockets(Idx2).DevName & " does not exist"
            GoTo CreateRouteForwards_error
        End If
    
'This code is to allow debugging if it is incorrectly called
'Check to see if Route and both Sockets enabled (they should be)
        For Ridx = 1 To UBound(Sockets(Idx1).Routes)
            If Sockets(Idx1).Routes(Ridx).AndIdx = Idx2 Then
                If Sockets(Idx1).Enabled = False Then
                    errmsg = Sockets(Idx1).DevName & " - "
                End If
                If Sockets(Idx2).Enabled = False Then
                    errmsg = errmsg & Sockets(Idx2).DevName & " - "
                End If
'The Route containing this socket must be enabled
                If Sockets(Idx1).Routes(Ridx).Enabled = False Then
                    errmsg = "Route between " & Sockets(Idx1).DevName _
                    & " and " & Sockets(Idx2).DevName & " - "
                End If
                If errmsg <> "" Then
                    errmsg = errmsg & "is disabled"
                    GoTo CreateRouteForwards_error
                Exit For
                End If
            End If
        Next Ridx
        End If

'Found the route between Idx1 and Idx2
    Call CreateStreamForwards(Idx1, Idx2)
Exit Sub

CreateRouteForwards_error:
    WriteLog "No Forwarding between " & Sockets(Idx1).DevName _
    & " and " & Sockets(Idx2).DevName & " because " & errmsg
End Sub

'Only called by CreateRouteForwards
'Input is the Route forward
'If either or both the Sockets are a TCP Listener
'Creates breaks out the forwards to the relevant streams
'if any
'If either is a TCP stream (when called) does not
'create the forward (should not really happen as the
'SocketForward should be blocked for a Stream
Public Sub CreateStreamForwards(Idx1 As Long, Idx2 As Long)
Dim Sidx1 As Long
Dim Sidx2 As Long
Dim ForwardIdx1 As Long
Dim ForwardIdx2 As Long

'Both Idx1 & Idx2 may be TCP Listeners
'So the forwards must be created using any Streams
'and not the Listener
    For Sidx1 = 1 To UBound(Sockets(Idx1).Winsock.Streams)
        For Sidx2 = 1 To UBound(Sockets(Idx2).Winsock.Streams)
            If IsTcpListener(Idx1) Then
                ForwardIdx1 = Sockets(Idx1).Winsock.Streams(Sidx1)
            Else
                ForwardIdx1 = Idx1
            End If
            If IsTcpListener(Idx2) Then
                ForwardIdx2 = Sockets(Idx2).Winsock.Streams(Sidx2)
            Else
                ForwardIdx2 = Idx2
            End If
'ForwardIdx will be 0 if there are no streans set up for
'a listener socket
            If ForwardIdx1 > 0 And ForwardIdx2 > 0 Then
                Call CreateForwards(ForwardIdx1, ForwardIdx2)
            End If
        Next Sidx2
    Next Sidx1
End Sub

'Creates forwards dependant only on the directios of the sockets
'Called by CreateRouteForwards,CreateStreamForwards
Public Sub CreateForwards(Idx1 As Long, Idx2 As Long)
'Dim Ridx As Integer
Dim Fidx As Long
Dim errmsg As String

    If Sockets(Idx1).Direction = Sockets(Idx2).Direction Then
        If Sockets(Idx1).Direction = 0 Then
            If ForwardExists(Idx1, Idx2) = False Then
                Fidx = FreeForward(Idx1)
                Sockets(Idx1).Forwards(Fidx) = Idx2
                Call IncrForward(Idx1, Idx2, False)
WriteLog "Added Forward " & Sockets(Idx1).DevName & " to " & Sockets(Idx2).DevName
            End If
            If ForwardExists(Idx2, Idx1) = False Then
                Fidx = FreeForward(Idx2)
                Sockets(Idx2).Forwards(Fidx) = Idx1
                Call IncrForward(Idx1, Idx2, True)
WriteLog "Added Forward " & Sockets(Idx2).DevName & " to " & Sockets(Idx1).DevName
            End If
        Else
'If Source & Destination both input or Both Output no route
        End If
    Else    'Directions differ
'v46 changed to <=
        If Sockets(Idx1).Direction <= 1 Then  'Idx1 to Idx2
            If ForwardExists(Idx1, Idx2) = False Then
                Fidx = FreeForward(Idx1)
                Sockets(Idx1).Forwards(Fidx) = Idx2
                Call IncrForward(Idx1, Idx2, False)
WriteLog "Added Forward " & Sockets(Idx1).DevName & " to " & Sockets(Idx2).DevName
            End If
        Else                    'Idx2 to Idx1
            If ForwardExists(Idx2, Idx1) = False Then
                Fidx = FreeForward(Idx2)
                Sockets(Idx2).Forwards(Fidx) = Idx1
                Call IncrForward(Idx1, Idx2, True)
WriteLog "Added Forward " & Sockets(Idx2).DevName & " to " & Sockets(Idx1).DevName
            End If
        End If
    End If
Exit Sub

CreateForwards_error:
    WriteLog "No Forwarding between " & Sockets(Idx1).DevName _
    & " and " & Sockets(Idx2).DevName & " because " & errmsg
End Sub


