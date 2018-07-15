VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If V44 Then
'Loaded into Collection Layouts the order thet are displayed
Public Name As String   'Key
Public FixHeight As Boolean 'True if Internal and Fixed
Public FixWidth As Boolean
'Public External As Boolean  'True external to the scale (Top of Form)
Public ctrl As Object
#End If
