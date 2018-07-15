Attribute VB_Name = "modTty"
Option Explicit

Public Sub CreateTTY(Idx As Long)
Dim Hidx As Long
   
WriteLog "Creating TTY Socket " & Sockets(Idx).DevName
    If Sockets(Idx).Hidx <= 0 Then
        Sockets(Idx).Hidx = FreeTTY
        If Sockets(Idx).Hidx = -1 Then
            MsgBox "Cant create TTY handler"
            Exit Sub
        End If
    End If
'For the TTY were not using a set-up form so we initialise the form
'This is done in cmd_OK if using setup form
'Add the Object(s) to the array TTYs() SetTTY has re-dimensioned
'the array (if it was too small)
        WriteLog aHandler(Sockets(Idx).Handler) & " Handler " _
& Sockets(Idx).Hidx & " allocated to " & Sockets(Idx).DevName
'This is here to make the App Log consistent with other handlers
        WriteLog "Opening " & Sockets(Idx).DevName
        Set TTYs(Sockets(Idx).Hidx) = New clsTTY
        Set TTYs(Sockets(Idx).Hidx).TTY = New frmTTY
        TTYs(Sockets(Idx).Hidx).sIndex = Idx
        TTYs(Sockets(Idx).Hidx).hIndex = Sockets(Idx).Hidx
        TTYs(Sockets(Idx).Hidx).TTY.Caption = Sockets(Idx).DevName
        TTYs(Sockets(Idx).Hidx).Name = Sockets(Idx).DevName
'We cannot display the TTY Form until the Modal form calling
'it has been closed
'This is done by MakeFormsVisisble
        TTYs(Sockets(Idx).Hidx).State = 1
End Sub


'Display TTY's (Config Menu)
Public Sub DisplayTTYs(Optional Caption As String)
Dim result As Boolean
Dim Idx As Long
Dim Hidx As Variant
Dim kb As String
Dim Count As Long
Dim i As Integer
Dim Fidx As Integer

    For Hidx = 1 To UBound(TTYs)
        If Not TTYs(Hidx) Is Nothing Then
            If Sockets(TTYs(Hidx).hIndex).State <> -1 Then
                kb = kb & Hidx & " = " & TTYs(Hidx).Name
                Fidx = -1
                If FormExists("frmTTY") = True Then
                    For i = 0 To Forms.Count - 1
                        If Forms(i).Caption = TTYs(Hidx).Name Then
                            Fidx = i
                            Exit For
                        End If
                    Next i
                End If
                If Fidx > -1 Then
                    If Forms(Fidx).Visible = True Then
                        kb = kb & " is Visible" & vbCrLf
                    Else
                        kb = kb & " is not Visible" & vbCrLf
                    End If
                Else
                    kb = kb & " is not Created" & vbCrLf
                End If
                Count = Count + 1
            End If
        End If
    Next Hidx
    If Count = 0 Then
        kb = kb & "There are no TTY handlers allocated" & vbCrLf
    End If
    If Caption = "" Then
        Caption = "TTY Ports"
    End If
    MsgBox kb, , Caption
End Sub

'Same as FreeComm
Public Function FreeTTY() As Long
Dim i As Long

    For i = 1 To UBound(TTYs)
        If Not TTYs(i) Is Nothing Then
            If TTYs(i).State = -1 Then
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
    
    FreeTTY = i
    If FreeTTY > MAX_TTYS Then
'no free TTYs
        WriteLog "No free display handlers, limit is " & MAX_TTYS
        FreeTTY = -1
    Else
        If FreeTTY > UBound(TTYs) Then
'We can still allocate more TTYs
            ReDim Preserve TTYs(1 To FreeTTY)
        End If
'The one just created
    End If
End Function

