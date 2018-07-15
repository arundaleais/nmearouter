VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

'This is how to trap Excel events in a VB program
Public WithEvents myWorkbookClass As Workbook
Attribute myWorkbookClass.VB_VarHelpID = -1

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsWorkBook; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub

Private Sub myWorkbookClass_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim A As Integer
    A = MsgBox("Do you really want to save the workbook?", vbYesNo)
    If A = vbNo Then Cancel = True

End Sub

'From ExcelVBA Object Model Help
'Using Events with the Application Object
'Before you can use events with the Application object,
'you must create a new class module and declare an object
'of type Application with events. For example, assume that a
'new class module is created and called EventClassModule.
'The new class module contains the following code.

'Public WithEvents App As ApplicationAfter the new object has been
'declared with events, it appears in the Object drop-down list box
'in the class module, and you can write event procedures for the
'new object. (When you select the new object in the Object box, the
'valid events for that object are listed in the Procedure drop-down
'list box.)

'Before the procedures will run, however, you must connect the
'declared object in the class module with the Application object.
'You can do this with the following code from any module.

'Dim X As New EventClassModule

'Sub InitializeApp()
'    Set X.App = Application
'End Sub
'After you run the InitializeApp procedure, the App object in the
'class module points to the Microsoft Excel Application object,
'and the event procedures in the class module will run when the
'events occur.

