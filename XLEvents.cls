VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XLEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'http://msdn.microsoft.com/en-us/library/office/aa140934%28v=office.10%29.aspx
Option Explicit

Private WithEvents xlApp As Excel.Application
Attribute xlApp.VB_VarHelpID = -1

Private Sub Class_Initialize()
   Set xlApp = Application
End Sub

Public Function NewWb() As Workbook
    Set NewWb = xlApp.Workbooks.Add
End Function
