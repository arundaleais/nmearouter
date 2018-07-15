Attribute VB_Name = "modExcelChart"
'see http://support.microsoft.com/kb/242243
'see http://support.microsoft.com/kb/142387
Option Explicit

Dim myExcelWorkbookModule As New clsWorkbook

Dim RowCount As Long    'Excludes header row
Dim ColCount As Long    'include time Column
Dim LastMsgCount() As Variant   'Cumulative
Dim LastColCount As Long
Dim ChartHeight As Single
Dim ChartWidth As Single
Dim myColours(1 To 6) As Long

Private Type defEmbedded
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
Dim Embedded As defEmbedded

Public ExcelLastUpdate As Long  'Minutes since
Public ExcelUpdateInterval As Long  'Minutes between updates
Public ExcelRange As Long       'Mas no of updates displayed
Public ExcelUTC As Boolean
Public ExcelOpen As Boolean
Const MsoFalse = 0
Const msoScaleFromTopLeft = 0
Const msoScaleFromMiddle = 1
Const msoScaleFromBottomRight = 2

Private myExcel As Excel.Application 'Time need to check if exists
Private myBook As Workbook      ' Excel workbook (must be declared as an object)
Private mySheet As Worksheet     ' Excel Worksheet
'Dim myShape As Shape        'Cant do this use mysheet.Shapes("Chart 1")
'Dim MyChartSheet As Worksheet
Private myChart As Chart     ' Excel Chart
'Dim mySeriesCollection As SeriesCollection

'This is my Excel Event class XLEvents
Private p_evtEvents As XLEvents

Public Function CreateWorkbook() As Object
Dim Col As Long      ' Index variable for the Column (Series)
Dim arry() As String
Dim i As Long
Dim ColourNo As Long
Dim Idx As Long
Dim kb As String
Dim Ver As String
Dim myLegendEntry As LegendEntry
Dim myLegendKey As LegendKey
Dim myObj As Object 'testing code

    ReDim arHeaderData(1 To 1)
    ReDim LastMsgCount(1 To 1)
    RowCount = 0
    ColCount = 1
    LastColCount = 0
    ExcelLastUpdate = 0     'No of times Timer has been called
    myColours(1) = vbRed
    myColours(2) = vbBlue
    myColours(3) = vbGreen
    myColours(4) = vbMagenta
    myColours(5) = vbCyan
    myColours(6) = vbYellow
    
    WriteLog "Creating Excel Application, update interval " & ExcelUpdateInterval & " mins"
    
#If test Then
    frmRouter.GraphTimer.Interval = 10000
    ExcelRange = 3
#Else
    On Error GoTo Create_Error
#End If

'Create Excel
    Set myExcel = New Excel.Application
    Select Case myExcel.Version
        Case "2.0"
            Ver = "1987 Excel 2.0 for Windows"
        Case "3.0"
            Ver = "1990 Excel 3"
        Case "4.0"
            Ver = "1992 Excel 4"
        Case "5.0"
            Ver = "1993 Excel 5.0 (Office 4.2 & 4.3)"
        Case "7.0"
            Ver = "1995 Excel 7.0 (Office 95)"
        Case "8.0"
            Ver = "1997 Excel 8.0 (Office 97)"
        Case "9.0"
            Ver = "1999 Excel 9.0 (Office 2000)"
        Case "10.0"
            Ver = "2001 Excel 10.0 (Office XP)"
        Case "11.0"
            Ver = "2003 Excel 11.0 (Office 2003)"
        Case "12.0"
            Ver = "2007 Excel 12.0 (Office 2007)"
        Case "13.0"
            Ver = "2010 Excel 13.0 (Office 2010)"
        Case "14.0"
            Ver = "2013 Excel 14.0 (Office 2013)"
        Case Else
            Ver = Ver & " Unknown version"
        End Select
    WriteLog "Excel Application " & Ver & " Created"

'Create a workbook(1) & Window
    Set myBook = myExcel.Workbooks.Add
    WriteLog "Excel Workbook " & myBook.Name & " Created"

'Exit Function
'Make Excel visible
    myExcel.Visible = True
    Set mySheet = myBook.Worksheets.Item(1)
    myBook.Windows(1).DisplayHeadings = False
    mySheet.Name = "Data"
    WriteLog "Excel Worksheet " & mySheet.Name & " Created"
'This creates only a chart as a separate sheet
'We have to do this first then change the location

    Call CreateChart

'True = When program exits leave excel window open
'False = When program exits close excel
    myExcel.UserControl = False
 '   mySheet.Application.UserControl = False
    ExcelOpen = True
    RowCount = -1   '-1 sets the initial values
    Call SetChartTitle("Initialised")
    Call AddSheetRow
'Rowcount is now 0
WriteLog "Excel Application Initialised"

    Exit Function

Create_Error:
    On Error GoTo 0
    MsgBox "Create Error " & Str(err.Number) & " " & err.Description & vbCrLf, , "Create Workbook"
End Function

Public Function AddSheetRow()
Dim Col As Long
Dim Row As Long
'Dim arHeaderData() As Variant
Dim arRowData() As Variant 'This = Last by col
Dim ThisMsgCount() As Long   'This=last by Socket
Dim kb As String
Dim i As Long
Dim Idx As Long
Dim ColourNo As Long
Dim NewColAdded As Boolean
Dim myLegendEntry As LegendEntry
Dim myLegendKey As LegendKey
Dim SumCol As Long
Dim MySeries As Series
Dim SeriesNo As Long
Dim ColName As String
Dim myRange As Range


'This is required, must be a minimum of 1 col to have set title
'    If ExcelOpen = False Or ColCount = 0 Then
    If ExcelOpen = False Then
        Exit Function
    End If
        
    If WorkbookExists = False Then
        ExcelOpen = False
        Exit Function
    End If
    
#If test = False Then  'v48
    On Error GoTo ChartFormat_Error
#End If
    Set mySheet = myBook.Worksheets("data")
'Scan the top row of the sheet to find the device name
'and create the total new messages for each column
'into arRowData
    ReDim ThisMsgCount(1 To UBound(Sockets))
    ReDim arRowData(1 To ColCount)
    If ExcelUTC = True Then
        arRowData(1) = Format$(UTCTimeNow, "#0:00")
    Else
        arRowData(1) = Format$(Time, "hh:nn")
    End If
    If UBound(LastMsgCount) < UBound(Sockets) Then
        ReDim Preserve LastMsgCount(1 To UBound(Sockets))
    End If
    
'The first time its call set the inital message count
    If RowCount = -1 Then
        For Idx = 1 To UBound(Sockets)
            LastMsgCount(Idx) = Sockets(Idx).MsgCount
        Next Idx
'The next row to be output, if we have data will be 1
'If no data it will remain as 0
        RowCount = 0
    End If
    
    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
'Accumulate by socket first because a socket could have closed and
'reopened on with a different device since the last msg count
'When Initialiasing ALL non TCPCLient sockets are set up
'otherwise an addition pass has to be made to set the Col header
            If .State > 0 _
            And .MsgCount > LastMsgCount(Idx) _
            And Not IsTcpListener(Idx) Then
'Excludes TcpListeners (always 0 count)
'With TCP Clients mutiple sockets can be accumulated in 1 col
                ThisMsgCount(Idx) = .MsgCount - LastMsgCount(Idx)
                
                Col = 2
                Do Until mySheet.Cells(1, Col) = ""
                    If mySheet.Cells(1, Col) = GetColName(Idx) Then
                        Exit Do
                    End If
                    Col = Col + 1
                Loop
'Insert a new column - col name not found
                If Col > ColCount Then
                     mySheet.Cells(1, Col) = GetColName(Idx)
                    ColCount = Col
                End If
                If ColCount > UBound(arRowData) Then
                    ReDim Preserve arRowData(1 To ColCount)
                End If
'If data (since last called) on the socket
                If ThisMsgCount(Idx) > 0 Then
                    arRowData(Col) = arRowData(Col) + ThisMsgCount(Idx)
                End If
            LastMsgCount(Idx) = .MsgCount
            End If
        End With    'sockets
    Next Idx
            
'Exit if no data cols
    If ColCount <= 1 Then
        Call SetChartTitle("Awaiting Data")
Debug.Print RowCount & ":" & ColCount
        Exit Function
    End If
    
'We have data to add so incr the row count
    RowCount = RowCount + 1
    
'Debug.Print RowCount & ":" & ColCount
        
'Stop Graph flickering
#If test Then
    myExcel.ScreenUpdating = True
#Else
    myExcel.ScreenUpdating = False
#End If

'Rowcount does not include header line
'Rowcount is the current row we are adding
'XXXXXXXXXXXX Delete Row
'Save the Size of the Graph
    If mySheet.Shapes.Count = 1 Then
        Embedded.Left = mySheet.Shapes(1).Left
        Embedded.Width = mySheet.Shapes(1).Width
        Embedded.Top = mySheet.Shapes(1).Top
        Embedded.Height = mySheet.Shapes(1).Height
    End If
    
    For Row = RowCount To ExcelRange + 1 Step -1
        mySheet.Rows(2).Delete (xlUp)
        RowCount = RowCount - 1
    Next Row

'Restore the size of the Graph
    If mySheet.Shapes.Count = 1 Then
        mySheet.Shapes(1).Left = Embedded.Left
        mySheet.Shapes(1).Width = Embedded.Width
        mySheet.Shapes(1).Top = Embedded.Top
        mySheet.Shapes(1).Height = Embedded.Height
    End If
                                
                
'Add the new data into next line after AddSheetRow called
'once (to set initial Counts)
    If RowCount > 0 Then
        mySheet.Range("A1").Offset(RowCount, 0).Resize(1, ColCount).Value = arRowData
    End If
'Remove any cols total nil from the right
'XXXXXXXXXXXXX Delete Zero Balance Column
'Save the Size of the Graph
    If mySheet.Shapes.Count = 1 Then
        Embedded.Left = mySheet.Shapes(1).Left
        Embedded.Width = mySheet.Shapes(1).Width
        Embedded.Top = mySheet.Shapes(1).Top
        Embedded.Height = mySheet.Shapes(1).Height
    End If
    
    For Col = ColCount To 2 Step -1
        If RowCount > 0 And Col > 1 Then
'added v48
'Does not fix Worksheet function error
'This was caused because the Range must be full qualified
'ie .cells not cells (cells refers to the active sheet)
            With mySheet
                Set myRange = .Range(.Cells(2, Col), .Cells(ColCount, Col))
            End With
            SumCol = WorksheetFunction.sum(myRange)
            If SumCol = 0 Then
'Must use Columns not range
                mySheet.Columns(Col).Delete (xlLeft)
                ColCount = ColCount - 1
            End If
        End If
    Next Col
    
'Restore the Size of the Graph
    If mySheet.Shapes.Count = 1 Then
        mySheet.Shapes(1).Left = Embedded.Left
        mySheet.Shapes(1).Width = Embedded.Width
        mySheet.Shapes(1).Top = Embedded.Top
        mySheet.Shapes(1).Height = Embedded.Height
    End If
    Call UpdateChart
    myExcel.ScreenUpdating = True
    Exit Function

NmySheet_Error:
    myExcel.ScreenUpdating = True
    MsgBox "Excel Sheet has been deleted "
    Call CloseExcel(vbNo)
    Exit Function
    
NmyChart_Error:
    myExcel.ScreenUpdating = True
    MsgBox "Excel Chart has been deleted"
    Call CloseExcel(vbNo)
    Exit Function
    
ChartFormat_Error:
    myExcel.ScreenUpdating = True
    MsgBox "Excel Chart Format is in error " & err.Number & " " & err.Description
    Call CloseExcel(vbNo)
    Exit Function

End Function

'Only called once when chart is first created
Private Function CreateChart()
Dim txtTextBox As String
        
    On Error GoTo Create_Error
    If Not myChart Is Nothing Then
'Stop 'MyChart should have been set to nothing when closed
        Set myChart = Nothing
        Exit Function
    End If
    
'Creates a New Sheet Chart 1 in Book
    Set myChart = myBook.Charts.Add(, mySheet)
'HasLegend is created and true when Chart is created
'BUT you cannot access it
'Call DisplayHas 'HasLegend
'Must create with an embedded sheet
    Call AwaitingData
'Initially set up either as embedded or separate sheet
    myChart.Location Where:=xlLocationAsNewSheet, Name:="Graph"
#If test Then
    myChart.Location Where:=xlLocationAsObject, Name:="Data"
#End If
'This is required to change the mychart object if its
'Location has been changed above to embedded
    If SetMyChart = True Then
        myChart.ChartType = xlLineMarkers
    End If
Call UpdateChart    'Test update with no data
    Exit Function
'Error probably caused by user changing chart whilst
'It was being created
Create_Error:
    MsgBox "Error " & err.Description & vbCrLf & "Failed to Create Chart", , "Create Chart"
End Function
        
Private Function UpdateChart()
Dim MySeries As Series
Dim Col As Long
Dim kb As String
Dim NewTitle As Boolean
Dim NewAxis As Boolean
Dim NewLegend As Boolean
Static TitleFormatted As Boolean
Static LegendFormatted As Boolean
Static AxisFormatted As Boolean
Dim MyShape As Shape
Dim MyTextBox As Textbox
Dim SeriesCount As Long

#If test = False Then
    On Error GoTo Update_Error
#End If

'Cant plot if sheet has been deleted
    If mySheet Is Nothing Then Exit Function
'Check chart exists and set correct location in myChart
    If SetMyChart = False Then
        Exit Function
    End If
'Have we data   Have we data
'to plot        Plotted
'no             no        update time, exit function
'yes            no        format chart area
'no             yes       remove format
'yes            yes       plot data
        
'Because the text box has been created, this now works
'with embedded and separate graph
'    myChart.Shapes(1).TextFrame.Characters.Text = CStr(UTCDateNow) & " UTC"
'No data
    If ColCount <= 1 Then
        Call AwaitingData
'Delete chart SeriesCollections left because no columns
        Do Until myChart.SeriesCollection.Count = 0
            Set MySeries = myChart.SeriesCollection(1)
            MySeries.Delete 'deletes the series of the Graph
            Set MySeries = Nothing
        Loop

 'Call DisplayHas    'HasLegend
   Else
'Remove awaiting data
        If myChart.Shapes.Count > 0 Then
#If test = False Then
            On Error Resume Next
#End If
            myChart.Shapes(1).Delete
#If test = False Then
            On Error GoTo Update_Error
#End If
        End If
        myChart.SetSourceData Source:=mySheet.Range("A1").Resize(RowCount + 1, ColCount), PlotBy:=xlColumns
'Got to reset my chart to pick up the legend entries !!!!
'Remove any series you dont want to display on the graph
        Call SetMyChart
        For Each MySeries In myChart.SeriesCollection
            If MySeries.Name <> "" Then
                If IsSeriesEnabled(MySeries.Name) = False Then
                    MySeries.Delete 'deletes the series of the Graph
                Else
                    SeriesCount = SeriesCount + 1
                End If
            End If
        Next
        If SeriesCount = 0 Then
            Call AwaitingData
        End If
#If False Then
        For Col = 2 To ColCount
'Remove any series you dont want to display on the graph
            If mySheet.Cells(1, Col).Text = "Homexx" Then
                Set MySeries = myChart.SeriesCollection(mySheet.Cells(1, Col).Text)
'Sets all the values to 0, without removing the legend key
'                    MySeries.Values = "{0}"
                    MySeries.Delete 'deletes the series of the Graph
                Set MySeries = Nothing
            End If
        Next Col
#End If
    End If
    
'With Title you have to set HasTitle, then format it
       If TitleExists = True Then
            If TitleFormatted = False Then
'Format title
                myChart.ChartTitle.Characters.Text = "NmeaRouter [" & CurrentProfile & "]"
                TitleFormatted = True
            End If
        Else
            TitleFormatted = False
        End If
    
    If LegendExists = True Then
        If LegendFormatted = False Then
'Format Legend
            LegendFormatted = True
        End If
        Call UpdateLegend
    Else
        LegendFormatted = False
    End If
    
    If AxisExists = True Then
        With myChart
            If AxisFormatted = False Then
'Format Axis
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Sentences"
                With .Axes(xlValue)
                End With
                .Axes(xlCategory, xlPrimary).HasTitle = True
                If ExcelUTC = True Then
                    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text _
                    = CStr(UTCDateNow) & " UTC"
                Else
                    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text _
                    = CStr(Date) & " Local"
                End If
'must be done after axis
                .PlotArea.Interior.ColorIndex = xlNone
                AxisFormatted = True
            End If
        End With
    Else
        AxisFormatted = False
    End If

    myChart.Deselect
    kb = kb & "TitleFormatted=" & TitleFormatted & vbCrLf
    kb = kb & "LegendFormatted=" & LegendFormatted & vbCrLf
    kb = kb & "AxisFormatted=" & AxisFormatted & vbCrLf
'MsgBox kb
    Exit Function
'Error probably caused by user changing chart whilst
'It was being created
Update_Error:
    MsgBox "Error " & err.Description & vbCrLf & "Failed to Create Chart", , "Update Chart"
End Function

'If you want to start a new Workbook just set ExcelOpen=false
Public Function CloseExcel(Optional SaveBook As Integer)

    On Error GoTo Close_Error
    frmRouter.mshSockets.ColWidth(8) = 0

'need to check at the start of each sub in case
'worksheet is closed (x) by user
    If WorkbookExists = False Then
'Play safe & ensure rest are set to nothing
        Set myChart = Nothing
        Set mySheet = Nothing
        Set myBook = Nothing
        frmRouter.GraphTimer.Enabled = False
        frmRouter.MenuViewInoutGraph.Checked = False
        Exit Function
        End If
WriteLog "Closing Excel"
    frmRouter.GraphTimer.Enabled = False
    frmRouter.MenuViewInoutGraph.Checked = False
    myExcel.ScreenUpdating = True
'User may have closed excel
    On Error Resume Next
'    myBook.Close False  'Close & dont save
    myExcel.Visible = False 'otherwise a blank frmRouter is displayed
'    myBook.SaveAs FileName:=myBook.Name
    If SaveBook = vbNo Then    ' dont save workbook
        myBook.Close False
WriteLog "Workbook not saved"
    Else
        myBook.Save    'save any other windows open
WriteLog "Workbook saved"
    End If
'    myBook.Close True
    Set myChart = Nothing
    Set mySheet = Nothing
    Set myBook = Nothing
'    myExcel.Workbooks.Close
    myExcel.Quit        'Shut excel down
    Set myExcel = Nothing
'You can call this to give a message
'    myExcelWorkbookModule.myWorkbookClass.Close
'    Set myExcelWorkbookModule = Nothing
   ExcelOpen = False
WriteLog "Excel Terminated"
    Exit Function
'Error probably caused by user changing chart whilst
'It was being created
Close_Error:
    MsgBox "Error " & err.Description & vbCrLf & "Failed to Create Chart", , "Close Excel"
End Function

Private Function GetColName(Idx As Long) As String
    On Error GoTo NoIdx
    If IsTcpStream(Idx) Then
        GetColName = Sockets(Idx).Winsock.RemoteHostIP
    Else
        GetColName = Sockets(Idx).DevName
    End If
    Exit Function
NoIdx:
    GetColName = ""
End Function
Public Function ClearGraphMsgCount(Idx As Long)
Dim i As Long
Dim Col As Long
'need to check at the start of each sub in case
'worksheet is closed (x) by user
    If WorkbookExists = False Then Exit Function
    For Col = 2 To ColCount
        If mySheet.Cells(1, Col) = GetColName(Idx) Then
            LastMsgCount(Idx) = 0
        End If
    Next Col
End Function

'Data range must have been set in arData
'AND Range must have been updated in Sheet
Private Sub UpdateLegend()
Dim i As Long
Dim ColourNo As Long
Dim myLegendEntry As LegendEntry
Dim myLegendKey As LegendKey
Dim kb As String
Dim Count As Long

'need to check at the start of each sub in case
'worksheet is closed (x) by user
    If WorkbookExists = False Then Exit Sub
    If SetMyChart = False Then Exit Sub
    With myChart
        If .HasLegend Then
            For Each myLegendEntry In .Legend.LegendEntries
                Set myLegendKey = myLegendEntry.LegendKey
                With myLegendKey
Count = Count + 1   'just checking legendentries.count
'Use my colours
                    ColourNo = ColourNo + 1
                    If ColourNo > UBound(myColours) Then ColourNo = 1
                        .Border.Color = myColours(ColourNo)
                        .Border.Weight = xlThick
                        .MarkerBackgroundColor = myColours(ColourNo)
                        .MarkerForegroundColor = myColours(ColourNo)
'Remove Legend Key Markers
                        If RowCount > 1 Then
                            .MarkerBackgroundColor = myColours(ColourNo)
                            .MarkerForegroundColor = myColours(ColourNo)
                            .MarkerStyle = xlSquare
                            .Smooth = False
                            .MarkerSize = 2
                            .Shadow = False
                           .MarkerStyle = xlNone
                        End If
                End With
            Next
'You must refresh the chart to see the changes
            .Refresh
        End If
    End With
End Sub

Public Function WorkbookExists() As Boolean
'On error traps a subscript error which occurs if
'the workbook has been closed with (X)
    On Error GoTo Notfound
    If Not myExcel.Windows.Item(1) Is Nothing Then
        WorkbookExists = True
    End If
    Exit Function
Notfound:
    WorkbookExists = False
End Function

'Trying to get the name causes an error if no data points
'returns ""
Private Function GetSeriesName(MySeries As Series) As String
Dim SeriesNo As Long
    
    On Error GoTo NoName
    If MySeries.Name <> "" Then
        GetSeriesName = MySeries.Name
        Exit Function
    End If
NoName:
End Function

Private Function GetSeriesNo(SeriesName As String) As Long
Dim MySeries As Series
    If Not MySeries Is Nothing Then
        For Each MySeries In myChart.SeriesCollection
            If MySeries.Name = SeriesName Then
'                GetSeriesNo = MySeries
            End If
        Next MySeries
    End If
End Function


Private Function MySeriesCount() As Long
Dim MySeries As Series
Dim Count As Long
        For Each MySeries In myChart.SeriesCollection
            If MySeries.Name <> "" Then
                Count = Count + 1
            End If
        Next MySeries
    MySeriesCount = Count
End Function


Public Function SetMyChart() As Boolean
    
'Trap mybook not set when Graphcfg first loaded (Marcus)
    If myBook Is Nothing Then
        Exit Function
    End If
    On Error GoTo TryEmbedded
    If myBook.Charts.Count > 0 Then
        Set myChart = myBook.Charts(1)
        SetMyChart = True
        Exit Function
    End If
TryEmbedded:
    On Error GoTo Nochart
    If mySheet.ChartObjects.Count > 0 Then
        Set myChart = mySheet.ChartObjects(1).Chart
        On Error GoTo 0
        SetMyChart = True
        Exit Function
    End If
Nochart:
End Function

Private Function SetChartTitle(ChartState As String)
    With myChart
    On Error GoTo Chart_Error
'        .HasTitle = True
        Select Case ChartState
        Case Is = "Initialising"
            .ChartTitle.Characters.Text = frmRouter.Caption & " [" & ChartState & "]"
        Case Else
            .ChartTitle.Characters.Text = "NmeaRouter [" & ChartState & "]"
        End Select
    End With
Chart_Error:
End Function

Private Function ChartTime() As String
    With myChart
        If ExcelUTC = True Then
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time UTC"
        Else
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time Local"
        End If
    End With
End Function

Private Sub DisplayHas()
Dim kb As String

With myChart
    kb = "HasTitle=" & .HasTitle & vbCrLf _
& "HasLegend=" & .HasLegend & vbCrLf
    On Error Resume Next
    kb = kb & "HasAxis=" & .HasAxis(xlValue, xlPrimary) & vbCrLf
    MsgBox kb
End With

End Sub

'The HasAxis variable is only created when the Axis is
'created so we have to check like this to avoid an error
Private Function AxisExists() As Boolean
    On Error GoTo NoAxis
    AxisExists = myChart.HasAxis(xlValue, xlPrimary)
    If AxisExists = False Then Exit Function
    AxisExists = myChart.HasAxis(xlValue)
'    myChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text
    Exit Function
NoAxis:
End Function

Private Function LegendExists() As Boolean
    If myChart.HasLegend = True Then
        On Error GoTo NoLegend
        If myChart.Legend.LegendEntries.Count > 0 Then
            LegendExists = True
        End If
    End If
NoLegend:
End Function

Private Function TitleExists() As Boolean
Dim kb As String
    On Error GoTo NoTitle
    If myChart.HasTitle = False Then myChart.HasTitle = True
    If myChart.HasTitle = True Then
        kb = myChart.ChartTitle.Characters.Text
        TitleExists = True
    End If
NoTitle:
End Function

Private Function AwaitingData()
Dim txtTextBox As String
    If ExcelUTC = True Then
        txtTextBox = Format$((UTCTimeNow), "0000") & " UTC" & vbLf & "[Awaiting Data]"
    Else
        txtTextBox = CStr(Date) & " Local" & vbLf & "[Awaiting Data]"
    End If
    If myChart.Shapes.Count = 0 Then
        With myChart.TextBoxes.Add(200, 25, 50, 14)
            .Name = "AwaitingData"
            .AutoSize = True
            .Text = txtTextBox
        End With
    Else
        myChart.Shapes(1).TextFrame.Characters.Text = txtTextBox
    End If
End Function

Public Function SyncGraphEnabled(ReqIdx As Long)
Dim Idx As Long
Dim ReqColName As String

'Sychronise all enabled graphs
'If TCP Server (with streams) is changed synchronise
'any streams owned by this server to server
'If TCP Stream Syncronise to ColName
'Anything else Syncronise to ColName
    
    ReqColName = GetColName(ReqIdx)
    
    For Idx = 1 To UBound(Sockets)
'Is Server Selected has any streams sync streams to server
        If IsTcpServer(ReqIdx) Then
            If Sockets(Idx).Winsock.Oidx = ReqIdx Then
                Sockets(Idx).Graph = Sockets(Sockets(Idx).Winsock.Oidx).Graph
            End If
        End If
'If anthing else synce any names to selection
        If GetColName(Idx) = GetColName(ReqIdx) Then
            Sockets(Idx).Graph = Sockets(ReqIdx).Graph
        End If
    Next Idx
    Call UpdateChart
End Function

Private Function IsSeriesEnabled(SeriesName As String) As Boolean
Dim Idx As Long

    For Idx = 1 To UBound(Sockets)
        If GetColName(Idx) = SeriesName Then
            If Sockets(Idx).Graph = True Then
                IsSeriesEnabled = True
                Exit Function
            End If
        End If
    Next Idx
End Function

