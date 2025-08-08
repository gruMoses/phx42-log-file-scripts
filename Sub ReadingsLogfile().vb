'/**
' * Sensor Readings Logfile Processing Module
' *
' * This module processes sensor data from CSV files, applies formatting,
' * identifies anomalies, and performs data analysis.
' *
' * @author Kevin Moses
' * @version 2.0
' */

' Constants for column positions
Private Const LPH2_COLUMN As Integer = 10          ' Column J (lph2)
Private Const SOLENOID_COLUMN As Integer = 19      ' Column S (solenoid)
Private Const FLAMEOUT_COLUMN As Integer = 13      ' Column M (iTemp - internal temp)
Private Const IS_IGNITED_COLUMN As Integer = 24    ' Column X (is ignited)
Private Const COMPARISON_COLUMN As Integer = 27     ' Column AA (comparison column)

' Constants for thresholds
Private Const MIN_OPERATING_TEMP As Double = 100     ' Minimum temperature to consider as operating
Private Const STEADY_STATE_SAMPLES As Integer = 5    ' Minimum samples to establish normal operating range
Private Const STEADY_STATE_THRESHOLD As Double = 0.005
Private Const BLIP_THRESHOLD As Double = 0.05
Private Const STEADY_STATE_MAX As Double = 1.3
Private Const VACUUM_GREEN_THRESHOLD As Double = -0.6
Private Const VACUUM_RED_THRESHOLD As Double = -1#
Private Const LPH2_COMPARISON_THRESHOLD As Double = 0.01  ' Threshold for LPH2 comparison (10% difference)

' Color variables (initialized at runtime)
Private COLOR_LIGHT_GREEN As Long
Private COLOR_LIGHT_RED As Long

'/**
' * Initialize color variables
' */
Private Sub InitializeColors()
    COLOR_LIGHT_GREEN = RGB(144, 238, 144)
    COLOR_LIGHT_RED = RGB(255, 182, 193)

End Sub

Sub CreatePressurePowerChart()
    On Error GoTo ErrorHandler
    g_Step = "init"
    
    Dim wbPath As String
    wbPath = GetActiveWorkbookPath()
    EnableChartDebugLogging wbPath
    g_Step = "enter"
    Dim dataWs As Worksheet
    Set dataWs = ActiveSheet
    Call LogDebug("ActiveSheet name=" & dataWs.Name)
    Dim dataWb As Workbook
    Set dataWb = dataWs.Parent
    Call LogDebug("Workbook path=" & dataWb.Path)
    
    ' Determine last row and last column
    Dim lastRow As Long, lastCol As Long
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    Call StepTag("lastrow=" & lastRow)
    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
    If lastRow < 3 Then
        MsgBox "Not enough data rows to chart.", vbExclamation
        Exit Sub
    End If
    
    ' Find columns by header names (row 1)
    Dim colTime As Long
    colTime = 1  ' Column A is time in this workbook
    
    Dim colSPress As Long, colCPress As Long, colSPPL As Long, colCPPL As Long
    colSPress = FindHeaderColumn(dataWs, "sPress")
    colCPress = FindHeaderColumn(dataWs, "cPress")
    colSPPL = FindHeaderColumn(dataWs, "sPPL")
    colCPPL = FindHeaderColumn(dataWs, "cPPL")
    
    ' If not found on the active sheet, auto-detect a worksheet with required headers
    If (colSPress = 0 And colCPress = 0) Or (colSPPL = 0 And colCPPL = 0) Then
        Dim wsCandidate As Worksheet
        Dim tSPress As Long, tCPress As Long, tSPPL As Long, tCPPL As Long
        For Each wsCandidate In dataWb.Worksheets
            ' Skip temporary chart sheet name if it exists later
            tSPress = FindHeaderColumn(wsCandidate, "sPress")
            tCPress = FindHeaderColumn(wsCandidate, "cPress")
            tSPPL = FindHeaderColumn(wsCandidate, "sPPL")
            tCPPL = FindHeaderColumn(wsCandidate, "cPPL")
            If (tSPress > 0 Or tCPress > 0) And (tSPPL > 0 Or tCPPL > 0) Then
                Set dataWs = wsCandidate
                colSPress = tSPress
                colCPress = tCPress
                colSPPL = tSPPL
                colCPPL = tCPPL
                Exit For
            End If
        Next wsCandidate
    End If
    
    If colSPress = 0 And colCPress = 0 Then
        MsgBox "Could not find pressure columns (sPress/cPress) in the header row.", vbExclamation
        Exit Sub
    End If
    Call StepTag("headers initial: sPress=" & colSPress & ", cPress=" & colCPress & ", sPPL=" & colSPPL & ", cPPL=" & colCPPL)
    If colSPPL = 0 And colCPPL = 0 Then
        MsgBox "Could not find pump power level columns (sPPL/cPPL) in the header row.", vbExclamation
        Exit Sub
    End If
    
    ' Create or clear chart sheet
    Dim chartWs As Worksheet
    On Error Resume Next
    Set chartWs = dataWb.Worksheets("PressurePowerChart")
    On Error GoTo ErrorHandler
    If chartWs Is Nothing Then
        Set chartWs = dataWb.Worksheets.Add(After:=dataWb.Sheets(dataWb.Sheets.Count))
        chartWs.Name = "PressurePowerChart"
    Else
        chartWs.Cells.Clear
        Dim shp As Shape
        For Each shp In chartWs.Shapes
            shp.Delete
        Next shp
    End If
    
    Call StepTag("using data sheet: " & dataWs.Name & " lastRow pre-recalc=" & lastRow)
    ' UI labels and default zoom settings
    ' Recalculate lastRow in case dataWs changed during auto-detect
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    Call StepTag("lastrow=" & lastRow)
    chartWs.Range("A1").Value = "Zoom start (points)"
    chartWs.Range("A2").Value = "Window size (points)"
    chartWs.Range("B1").Value = 0
    chartWs.Range("B2").Value = Application.WorksheetFunction.Min(1000, lastRow - 1)
    chartWs.Range("A4").Value = "Tip: Use the scroll bars to pan/zoom."
    chartWs.Range("A5").Value = "Primary axis: Pressures (sPress, cPress). Secondary axis: Pump power (sPPL, cPPL)."
    
    ' Build ranges directly from the data (no workbook names)
    Dim xRange As Range
    Dim rSPress As Range, rCPress As Range, rSPPL As Range, rCPPL As Range
    Set xRange = dataWs.Range(dataWs.Cells(2, colTime), dataWs.Cells(lastRow, colTime))
    If colSPress > 0 Then Set rSPress = dataWs.Range(dataWs.Cells(2, colSPress), dataWs.Cells(lastRow, colSPress))
    If colCPress > 0 Then Set rCPress = dataWs.Range(dataWs.Cells(2, colCPress), dataWs.Cells(lastRow, colCPress))
    If colSPPL > 0 Then Set rSPPL = dataWs.Range(dataWs.Cells(2, colSPPL), dataWs.Cells(lastRow, colSPPL))
    If colCPPL > 0 Then Set rCPPL = dataWs.Range(dataWs.Cells(2, colCPPL), dataWs.Cells(lastRow, colCPPL))
' Create the chart
    Dim co As ChartObject
    ' Use a fixed, visible size so the chart renders clearly on a cleared sheet
    Set co = chartWs.ChartObjects.Add(Left:=20, Top:=80, Width:=1200, Height:=500)
    co.Name = "PressurePowerChartObject"
    Call StepTag("chart created")
    
    With co.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = True
        .ChartTitle.Text = "Low Pressure Sensors and Pump Power Levels"
        .Legend.Position = xlLegendPositionBottom
        
        ' Add pressure series on primary axis
        If colSPress > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "sPress"
                .XValues = xRange
                .Values = rSPress
                .AxisGroup = xlPrimary
                .Format.Line.ForeColor.RGB = RGB(33, 150, 243) ' blue
                .Format.Line.Weight = 2.25
            End With
        End If
        If colCPress > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "cPress"
                .XValues = xRange
                .Values = rCPress
                .AxisGroup = xlPrimary
                .Format.Line.ForeColor.RGB = RGB(76, 175, 80) ' green
                .Format.Line.Weight = 2.25
            End With
        End If
        
        ' Add pump power level series on secondary axis
        If colSPPL > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "sPPL"
                .XValues = xRange
                .Values = rSPPL
                .AxisGroup = xlSecondary
                .Format.Line.ForeColor.RGB = RGB(255, 152, 0) ' orange
                .Format.Line.Weight = 2.25
            End With
        End If
        If colCPPL > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "cPPL"
                .XValues = xRange
                .Values = rCPPL
                .AxisGroup = xlSecondary
                .Format.Line.ForeColor.RGB = RGB(233, 30, 99) ' pink/red
                .Format.Line.Weight = 2.25
            End With
        End If
        
        ' Axes formatting
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Time"
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Pressure"
        End With
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "Pump Power Level"
        End With
    End With
    
    ' Scrollbars removed for macOS stability
On Error GoTo ErrorHandler
    
    chartWs.Activate
    Call LogDebug("SUCCESS")
    
    Exit Sub
    
ErrorHandler:
    Dim errNum As Long, errDesc As String, errStep As String
    errNum = Err.Number: errDesc = Err.Description: errStep = g_Step
    On Error Resume Next
    Dim logPath As String
    Dim ff As Integer
    logPath = GetActiveWorkbookPath()
    If Right$(logPath, 1) <> Application.PathSeparator Then logPath = logPath & Application.PathSeparator
    logPath = logPath & "phx42_chart_debug.log"
    ff = FreeFile
    Open logPath For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ERROR: step=" & errStep & " code=" & CStr(errNum) & " - " & errDesc
    Close #ff
    On Error GoTo 0
    MsgBox "Error in CreatePressurePowerChart: " & Err.Description, vbExclamation
End Sub

' Finds column index by exact header match (case-insensitive). Returns 0 if not found.
Private Function FindHeaderColumn(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long, c As Long
    Dim target As String, h As String
    target = LCase(Trim(headerText))
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        ' Exact match first
        If h = target Then
            FindHeaderColumn = c
            Exit Function
        End If
        ' Accept common synonyms from raw headers
        Select Case target
            Case "spress": If h = "sample pressure" Then FindHeaderColumn = c: Exit Function
            Case "cpress": If h = "combustion pressure" Then FindHeaderColumn = c: Exit Function
            Case "sppl":   If h = "sample ppl" Then FindHeaderColumn = c: Exit Function
            Case "cppl":   If h = "combustion ppl" Then FindHeaderColumn = c: Exit Function
        End Select
    Next c
    FindHeaderColumn = 0
End Function

' Adds a workbook-level name, replacing it if it exists
Private Sub AddOrReplaceName(ByVal nameText As String, ByVal refersToFormula As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersToFormula
End Sub

'/**
' * Makes the header row bold
' */
Sub BoldHeaderRow()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Find the last column with data
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Make the entire header row bold
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Font.Bold = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in BoldHeaderRow: " & Err.Description, vbExclamation
End Sub

'/**
' * Renames headers according to specified mapping
' */
Sub RenameHeaders()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Get the header row
    Dim headerRow As Range
    Set headerRow = ws.Rows(1)
    
    ' Find the last column with data
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim col As Long
    Dim headerText As String
    Dim lowerHeader As String
    
    ' Loop through each column in the header row
    For col = 1 To lastCol
        headerText = Trim(CStr(headerRow.Cells(1, col).value))
        lowerHeader = LCase(headerText)
        
        ' Check and rename headers using direct string comparisons
        Select Case lowerHeader
            Case "pa offset"
                headerRow.Cells(1, col).value = "Ofs"
            Case "sample pressure"
                headerRow.Cells(1, col).value = "sPress"
            Case "sample ppl"
                headerRow.Cells(1, col).value = "sPPL"
            Case "combustion pressure"
                headerRow.Cells(1, col).value = "cPress"
            Case "combustion ppl"
                headerRow.Cells(1, col).value = "cPPL"
            Case "internal temp."
                headerRow.Cells(1, col).value = "iTemp"
            Case "external temp."
                headerRow.Cells(1, col).value = "eTemp"
            Case "case temp."
                headerRow.Cells(1, col).value = "cTemp"
            Case "needle valve"
                headerRow.Cells(1, col).value = "MOV"
        End Select
    Next col
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in RenameHeaders: " & Err.Description, vbExclamation
End Sub

'/**
' * Formats numeric cells to show appropriate decimal places based on actual data
' */
Sub FormatDecimalPlaces()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Find the last row and column with data
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim col As Long
    Dim row As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim maxDecimals As Integer
    Dim currentDecimals As Integer
    Dim dotPos As Integer
    Dim hasDecimals As Boolean
    
    ' Process each column individually
    For col = 1 To lastCol
        ' Skip column A (time column) to preserve time formatting
        If col = 1 Then
            GoTo NextColumn
        End If
        
        maxDecimals = 0
        hasDecimals = False
        
        ' Analyze each cell in the column to find maximum decimal places
        For row = 2 To lastRow
            If Not IsEmpty(ws.Cells(row, col)) And IsNumeric(ws.Cells(row, col)) Then
                cellValue = ws.Cells(row, col).value
                cellText = CStr(cellValue)
                
                ' Check if the value has decimal places
                If cellValue <> Int(cellValue) Then
                    hasDecimals = True
                    
                    ' Find the decimal point
                    dotPos = InStr(cellText, ".")
                    If dotPos > 0 Then
                        ' Count decimal places
                        currentDecimals = Len(cellText) - dotPos
                        If currentDecimals > maxDecimals Then
                            maxDecimals = currentDecimals
                        End If
                    End If
                End If
            End If
        Next row
        
        ' Apply formatting based on the maximum decimal places found
        If hasDecimals Then
            Select Case maxDecimals
                Case 1
                    ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).NumberFormat = "0.0"
                Case 2
                    ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).NumberFormat = "0.00"
                Case 3
                    ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).NumberFormat = "0.000"
                Case 4
                    ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).NumberFormat = "0.000"
                Case Else
                    ' For more than 4 decimal places, use 3 decimal format
                    ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).NumberFormat = "0.000"
            End Select
        Else
            ' Apply general number format for columns without decimals
            ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).NumberFormat = "0"
        End If
        
NextColumn:
    Next col
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in FormatDecimalPlaces: " & Err.Description, vbExclamation
End Sub

'/**
' * Deletes rows based on specific conditions (N/A, NA, blank cells, blank sample pressure, "sample pressure" text)
' * Optimized for performance using arrays
' */
Sub DeleteRowsBasedOnConditions()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    Dim j As Integer
    Dim deleteRow As Boolean
    Dim cellValue As Variant
    Dim cellText As String
    
    ' Loop through rows from bottom to top to avoid issues with shifting rows
    For i = lastRow To 2 Step -1
        ' Show progress every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Cleaning data: Processing row " & i & " of " & lastRow & "..."
            DoEvents
        End If
        
        deleteRow = False
        
        ' Check column F (6th column) for blank/null sample pressure
        If IsEmpty(ws.Cells(i, 6)) Or IsNull(ws.Cells(i, 6)) Or _
           Trim(CStr(ws.Cells(i, 6).value)) = "" Or _
           ws.Cells(i, 6).value = "" Then
            deleteRow = True
        End If
        
        ' Check all columns in the row for "N/A" or "NA" or completely empty rows
        If Not deleteRow Then
            Dim hasData As Boolean
            hasData = False
            
            ' Find the last column with data in the first row to determine range
            Dim lastCol As Long
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            For j = 1 To lastCol
                cellValue = ws.Cells(i, j).value
                
                ' Check if cell has data
                If Not IsEmpty(cellValue) Then
                    hasData = True
                    
                    ' Convert to string and check for N/A variations
                    cellText = Trim(CStr(cellValue))
                    If UCase(cellText) = "N/A" Or UCase(cellText) = "NA" Or _
                       UCase(cellText) = "N/A " Or UCase(cellText) = "NA " Or _
                       UCase(cellText) = " N/A" Or UCase(cellText) = " NA" Then
                        deleteRow = True
                        Exit For
                    End If
                    
                    ' Check for "sample pressure" text (case insensitive)
                    If UCase(cellText) = "SAMPLE PRESSURE" Then
                        deleteRow = True
                        Exit For
                    End If
                End If
            Next j
            
            ' If no data found in the entire row, mark for deletion
            If Not hasData Then
                deleteRow = True
            End If
        End If
        
        ' Delete the row if any condition is met
        If deleteRow Then
            ws.Rows(i).Delete
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in DeleteRowsBasedOnConditions: " & Err.Description & vbNewLine & _
           "Row: " & i & ", Column: " & j, vbExclamation
End Sub

'/**
' * Auto-sizes all columns to fit their content
' */
Sub AutoSizeAllColumns()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Find the last column with data
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Auto-size all columns from 1 to lastCol
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).EntireColumn.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in AutoSizeAllColumns: " & Err.Description, vbExclamation
End Sub

'/**
' * Main entry point for processing sensor readings logfile
' * Performs data cleaning, formatting, and analysis
' */
Sub ReadingsLogfile()
    On Error GoTo ErrorHandler
    
    ' Store reference to the data workbook (not Personal workbook)
    Dim dataWorkbook As Workbook
    Set dataWorkbook = ActiveWorkbook
    
    ' Initialize color variables
    Call InitializeColors
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Show initial status
    Application.StatusBar = "Starting data processing..."
    DoEvents
    
    ' Create backup of original data
    Application.StatusBar = "Creating backup of original data..."
    DoEvents
    Call CreateBackupSheet(dataWorkbook)
    
    ' Set up window layout
    Application.StatusBar = "Setting up worksheet layout..."
    DoEvents
    Range("A1").Activate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ' Clear existing formatting
    Application.StatusBar = "Clearing existing formatting..."
    DoEvents
    Cells.FormatConditions.Delete
    
    ' Clean data first
    Application.StatusBar = "Cleaning data (removing invalid rows)..."
    DoEvents
    Call DeleteRowsBasedOnConditions
    
    ' Rename headers
    Application.StatusBar = "Renaming headers..."
    DoEvents
    Call RenameHeaders
    
    ' Bold the header row
    Application.StatusBar = "Bolding header row..."
    DoEvents
    Call BoldHeaderRow
    
    ' Format datetime in column A
    Application.StatusBar = "Formatting datetime column..."
    DoEvents
    Call FormatDateTimeColumn
    
    ' Format decimal places
    Application.StatusBar = "Formatting decimal places..."
    DoEvents
    Call FormatDecimalPlaces
    
    ' Apply data processing
    Application.StatusBar = "Applying voltage formatting..."
    DoEvents
    Call ColorRowsVoltage
    
    Application.StatusBar = "Applying vacuum formatting..."
    DoEvents
    Call ColorRowsVacuum
    
    ' Identify and highlight flameout events
    Application.StatusBar = "Identifying flameout events..."
    DoEvents
    Call IdentifyFlameouts
    
    ' Add watermark with serial number
    Application.StatusBar = "Adding serial number..."
    DoEvents
    Call AddSerialNumberToColumnB
    
    ' Process ignition states
    Application.StatusBar = "Processing ignition states..."
    DoEvents
    Call ProcessIgnitionStates
    
    ' Flag LPH2 changes
    Application.StatusBar = "Flagging LPH2 changes..."
    DoEvents
    Call FlagLPH2Changes
    
    ' Auto-size all columns
    Application.StatusBar = "Auto-sizing columns..."
    DoEvents
    Call AutoSizeAllColumns
    
    ' Find firmware logs
    Application.StatusBar = "Finding firmware logs..."
    DoEvents
    Call FindFirmwareLogs(dataWorkbook)
    
    ' Parse firmware log contents
    Application.StatusBar = "Parsing firmware log contents..."
    DoEvents
    
    Call ParseFirmwareLogContents(dataWorkbook)
    

    
    ' Save as Excel file
    Application.StatusBar = "Saving processed file..."
    DoEvents
    Call SaveAsExcelFile
    
    ' Show completion message
    Application.StatusBar = "Processing complete!"
    Application.Wait Now + timeValue("00:00:02")  ' Show completion for 2 seconds
    Application.StatusBar = False  ' Clear status bar
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error in ReadingsLogfile: " & Err.Description, vbExclamation
End Sub

'/**
' * Creates a backup of the original data in a new sheet named "RAW"
' */
Sub CreateBackupSheet(targetWorkbook As Workbook)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Check if RAW sheet already exists and delete it
    Dim rawSheet As Worksheet
    On Error Resume Next
    Set rawSheet = targetWorkbook.Worksheets("RAW")
    On Error GoTo ErrorHandler
    
    If Not rawSheet Is Nothing Then
        Application.DisplayAlerts = False
        rawSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new RAW sheet in the data workbook (not Personal workbook)
    Dim newSheet As Worksheet
    Set newSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    newSheet.Name = "RAW"
    
    ' Copy all data from active sheet to RAW sheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 0 And lastCol > 0 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy
        newSheet.Range("A1").PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    End If
    
    ' Activate the original sheet for processing
    ws.Activate
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Error in CreateBackupSheet: " & Err.Description, vbExclamation
End Sub

'/**
' * Applies conditional formatting to voltage column
' */
Sub ColorRowsVoltage()
    On Error GoTo ErrorHandler
    
    With Range("V:V").FormatConditions.AddAboveAverage
        .AboveBelow = xlAboveStdDev
        With .Font
            .Bold = True
            .ColorIndex = 3
        End With
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ColorRowsVoltage: " & Err.Description, vbExclamation
End Sub

'/**
' * Identifies and highlights flameout events using is ignited status and temperature drops
' */
Sub IdentifyFlameouts()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long, k As Long
    Dim value As Double
    Dim previous As Double
    Dim inFlameout As Boolean
    Dim peak As Double
    Dim intensity As Double
    Dim red As Integer, green As Integer, blue As Integer
    Dim isIgnited As Variant
    Dim wasIgnited As Boolean
    Dim flameoutStartRow As Long
    Dim steadyStateTemp As Double
    Dim tempDropping As Boolean
    
    inFlameout = False
    previous = 0
    peak = 0
    wasIgnited = False
    flameoutStartRow = 0
    steadyStateTemp = 0
    tempDropping = False
    
    For i = 2 To lastRow
        ' Show progress every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Identifying flameouts: Processing row " & i & " of " & lastRow & "..."
            DoEvents
        End If
        
        ' Check if we have valid data in the required columns and they exist
        If i <= lastRow And _
           FLAMEOUT_COLUMN <= ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column And _
           SOLENOID_COLUMN <= ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column And _
           IS_IGNITED_COLUMN <= ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column And _
           IsNumeric(ws.Cells(i, FLAMEOUT_COLUMN).value) And _
           Not IsEmpty(ws.Cells(i, SOLENOID_COLUMN).value) And _
           Not IsEmpty(ws.Cells(i, IS_IGNITED_COLUMN).value) And _
           ws.Cells(i, FLAMEOUT_COLUMN).value <> "" And _
           ws.Cells(i, SOLENOID_COLUMN).value <> "" And _
           ws.Cells(i, IS_IGNITED_COLUMN).value <> "" Then
            
            value = ws.Cells(i, FLAMEOUT_COLUMN).value
            Dim solenoid As Variant
            solenoid = ws.Cells(i, SOLENOID_COLUMN).value
            isIgnited = ws.Cells(i, IS_IGNITED_COLUMN).value
            
            If solenoid = 0 Then
                ' Reset and clear highlight if solenoid is off
                inFlameout = False
                ws.Cells(i, FLAMEOUT_COLUMN).Interior.ColorIndex = xlNone
                ws.Cells(i, FLAMEOUT_COLUMN).Font.Color = RGB(0, 0, 0)
                wasIgnited = False
                tempDropping = False
            Else
                ' Check for flameout using is ignited status
                If wasIgnited And (isIgnited = False Or isIgnited = "FALSE") Then
                    ' Flameout detected: was ignited, now not ignited
                    inFlameout = True
                    
                    ' Find where temperature first started dropping from steady state
                    Dim startRow As Long
                    startRow = i
                    Dim lookBack As Integer
                    lookBack = 0
                    
                    ' Find steady state temperature (average of last 5 readings before the flameout detection)
                    steadyStateTemp = 0
                    Dim steadyStateCount As Integer
                    steadyStateCount = 0
                    
                    For lookBack = 1 To 5
                        If startRow - lookBack >= 2 And _
                           IsNumeric(ws.Cells(startRow - lookBack, FLAMEOUT_COLUMN).value) Then
                            steadyStateTemp = steadyStateTemp + ws.Cells(startRow - lookBack, FLAMEOUT_COLUMN).value
                            steadyStateCount = steadyStateCount + 1
                        End If
                    Next lookBack
                    
                    If steadyStateCount > 0 Then
                        steadyStateTemp = steadyStateTemp / steadyStateCount
                    End If
                    
                    ' Look back to find where temperature first started dropping
                    Do While startRow > 2 And lookBack < 50  ' Look back up to 50 rows
                        If startRow <= lastRow And startRow > 1 And _
                           IsNumeric(ws.Cells(startRow, FLAMEOUT_COLUMN).value) And _
                           IsNumeric(ws.Cells(startRow - 1, FLAMEOUT_COLUMN).value) Then
                            
                            Dim currentTemp As Double
                            Dim previousTemp As Double
                            currentTemp = ws.Cells(startRow, FLAMEOUT_COLUMN).value
                            previousTemp = ws.Cells(startRow - 1, FLAMEOUT_COLUMN).value
                            
                            ' Check if this is where temperature started dropping
                            If currentTemp < previousTemp Then  ' Any temperature drop, no matter how small
                                startRow = startRow - 1
                                lookBack = lookBack + 1
                            Else
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    ' Peak temperature is at startRow
                    peak = ws.Cells(startRow, FLAMEOUT_COLUMN).value
                    flameoutStartRow = startRow + 1
                    
                    ' Highlight from where temperature first started dropping to current row
                    For k = flameoutStartRow To i
                        If IsNumeric(ws.Cells(k, FLAMEOUT_COLUMN).value) Then
                            Dim valueK As Double
                            valueK = ws.Cells(k, FLAMEOUT_COLUMN).value
                            intensity = (peak - valueK) / (peak - 40)  ' Assume minimum temp of 40°C
                            If intensity > 1 Then intensity = 1
                            If intensity < 0 Then intensity = 0
                            
                            ' Interpolate from light red to dark red
                            red = 255 - intensity * (255 - 139)
                            green = 160 - intensity * 160
                            blue = 122 - intensity * 122
                            
                            ws.Cells(k, FLAMEOUT_COLUMN).Interior.Color = RGB(red, green, blue)
                            
                            ' Contrasting text color
                            If intensity > 0.5 Then
                                ws.Cells(k, FLAMEOUT_COLUMN).Font.Color = RGB(255, 255, 255)  ' White for darker shades
                            Else
                                ws.Cells(k, FLAMEOUT_COLUMN).Font.Color = RGB(0, 0, 0)  ' Black for lighter shades
                            End If
                        End If
                    Next k
                End If
                
                ' Update ignition status
                wasIgnited = (isIgnited = True Or isIgnited = "TRUE")
                
                ' End flameout if returning to ignited status
                If wasIgnited And inFlameout Then
                    inFlameout = False
                    ws.Cells(i, FLAMEOUT_COLUMN).Interior.ColorIndex = xlNone
                    ws.Cells(i, FLAMEOUT_COLUMN).Font.Color = RGB(0, 0, 0)
                    tempDropping = False
                End If
                
                ' Continue highlighting if still in flameout
                If inFlameout Then
                    ' Highlight current row if not already done
                    intensity = (peak - value) / (peak - 40)
                    If intensity > 1 Then intensity = 1
                    If intensity < 0 Then intensity = 0
                    
                    red = 255 - intensity * (255 - 139)
                    green = 160 - intensity * 160
                    blue = 122 - intensity * 122
                    
                    ws.Cells(i, FLAMEOUT_COLUMN).Interior.Color = RGB(red, green, blue)
                    
                    If intensity > 0.5 Then
                        ws.Cells(i, FLAMEOUT_COLUMN).Font.Color = RGB(255, 255, 255)
                    Else
                        ws.Cells(i, FLAMEOUT_COLUMN).Font.Color = RGB(0, 0, 0)
                    End If
                Else
                    ws.Cells(i, FLAMEOUT_COLUMN).Interior.ColorIndex = xlNone
                    ws.Cells(i, FLAMEOUT_COLUMN).Font.Color = RGB(0, 0, 0)
                End If
            End If
            
            previous = value
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in IdentifyFlameouts: " & Err.Description, vbExclamation
End Sub

'/**
' * Applies conditional formatting to vacuum column
' */
Sub ColorRowsVacuum()
    On Error GoTo ErrorHandler
    
    Dim rg As Range
    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
    
    Set rg = Range("L2", Range("L2").End(xlDown))
    
    ' Clear any existing conditional formatting
    rg.FormatConditions.Delete
    
    ' Define the rules for each conditional format
    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, VACUUM_GREEN_THRESHOLD)
    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, VACUUM_RED_THRESHOLD)
    Set cond3 = rg.FormatConditions.Add(xlCellValue, xlLess, VACUUM_GREEN_THRESHOLD)
    
    ' Define the format applied for each conditional format
    With cond1
        .Interior.Color = vbGreen
    End With
    
    With cond2
        .Interior.Color = vbRed
    End With
    
    With cond3
        .Interior.Color = vbYellow
        .Font.Color = vbRed
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ColorRowsVacuum: " & Err.Description, vbExclamation
End Sub

'/**
' * Saves the processed workbook as an Excel file
' */
Sub SaveAsExcelFile()
    On Error GoTo ErrorHandler
    
    Dim originalName As String
    Dim newName As String
    Dim filePath As String
    Dim fullPath As String
    Dim counter As Integer
    
    ' Get the original file name and path
    originalName = ActiveWorkbook.Name
    ' Save to current folder instead of workbook path
    filePath = CurDir()
    
    ' Check if we're in the Excel sandbox (macOS issue)
    If InStr(filePath, "Library/Containers/com.microsoft.Excel") > 0 Then
        ' Use the original workbook path instead
        filePath = ActiveWorkbook.Path
    End If
    
    ' Ensure proper path separator for macOS
    If Right(filePath, 1) <> "/" And Right(filePath, 1) <> "\" Then
        filePath = filePath & "/"
    End If
    
    ' Create new name by removing .csv extension if present and adding .xlsx
    If LCase(Right(originalName, 4)) = ".csv" Then
        newName = Left(originalName, Len(originalName) - 4) & "_processed.xlsx"
    Else
        newName = Left(originalName, InStrRev(originalName, ".") - 1) & "_processed.xlsx"
    End If
    
    ' If no extension found, just add _processed.xlsx
    If InStrRev(originalName, ".") = 0 Then
        newName = originalName & "_processed.xlsx"
    End If
    
    ' Check if file already exists and create unique name
    counter = 1
    fullPath = filePath & newName
    
    Do While Dir(fullPath) <> ""
        newName = Left(newName, InStrRev(newName, ".") - 1) & "_" & counter & ".xlsx"
        fullPath = filePath & newName
        counter = counter + 1
        
        ' Prevent infinite loop
        If counter > 100 Then
            MsgBox "Could not create unique filename. Please close any open files and try again.", vbExclamation
            Exit Sub
        End If
    Loop
    
    ' Prompt user for permission to save file
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Do you want to save the processed file to:" & vbNewLine & _
                         fullPath & vbNewLine & vbNewLine & _
                         "Click 'Yes' to save or 'No' to cancel.", _
                         vbYesNo + vbQuestion, "Save Processed File")
    
    If userResponse = vbYes Then
        ' Try to save the workbook to current folder only
        Application.DisplayAlerts = False
        
        On Error Resume Next
        ActiveWorkbook.SaveAs fullPath, xlOpenXMLWorkbook
        If Err.Number = 0 Then
            Application.DisplayAlerts = True
            MsgBox "File saved successfully to:" & vbNewLine & fullPath, vbInformation, "Save Complete"
            Exit Sub
        End If
        On Error GoTo ErrorHandler
        
        ' If saving fails, show error message
        Application.DisplayAlerts = True
        MsgBox "Could not save file to current folder: " & fullPath & vbNewLine & _
               "Error: " & Err.Description, vbExclamation
    Else
        ' User cancelled the save operation
        MsgBox "File save operation cancelled by user.", vbInformation, "Save Cancelled"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Error in SaveAsExcelFile: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical
End Sub

'/**
' * Adds serial number to column B for every row
' * Replaces column B data with "phx42-XXXX" format
' */
Sub AddSerialNumberToColumnB()
    On Error GoTo ErrorHandler
    
    Dim serialNumber As String
    serialNumber = ExtractSerialNumber()
    
    If serialNumber <> "" Then
        ' Get the last row with data
        Dim lastRow As Long
        lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).row
        
        ' Fill column B with serial number for all rows
        Dim serialText As String
        serialText = "phx42-" & serialNumber
        
        ' Apply to all rows from 1 to lastRow
        ActiveSheet.Range("B1:B" & lastRow).value = serialText
        
        ' No popup message - silent operation
    Else
        ' No popup message - silent operation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in AddSerialNumberToColumnB: " & Err.Description, vbExclamation
End Sub

'/**
' * Extracts serial number from filename
' * Looks for pattern phx42-XXXX where XXXX is the serial number
' */
Function ExtractSerialNumber() As String
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    Dim pattern As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim serialNumber As String
    
    fileName = ActiveWorkbook.Name
    
    ' Look for phx42- pattern
    pattern = "phx42-"
    startPos = InStr(1, LCase(fileName), pattern)
    
    If startPos > 0 Then
        ' Find the end of the serial number (non-alphanumeric character or end of string)
        startPos = startPos + Len(pattern)
        endPos = startPos
        
        ' Find the end of the serial number
        Do While endPos <= Len(fileName)
            Dim char As String
            char = Mid(fileName, endPos, 1)
            
            ' Check if character is alphanumeric or dash
            If Not (char Like "[A-Za-z0-9-]") Then
                Exit Do
            End If
            
            endPos = endPos + 1
        Loop
        
        ' Extract the serial number
        If endPos > startPos Then
            serialNumber = Mid(fileName, startPos, endPos - startPos)
        End If
    End If
    
    ExtractSerialNumber = serialNumber
    Exit Function
    
ErrorHandler:
    ExtractSerialNumber = ""
End Function

'/**
' * Formats datetime in column A to show only time without UTC offset and milliseconds
' */
Sub FormatDateTimeColumn()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    Dim cellValue As Variant
    Dim timeOnly As String
    Dim spacePos As Integer
    Dim dashPos As Integer
    Dim dotPos As Integer
    Dim timeValue As Date
    
    ' Process each row from 2 to lastRow (skip header)
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, 1)) Then
            cellValue = ws.Cells(i, 1).value
            
            ' Check if it's already a decimal number (Excel time format)
            If IsNumeric(cellValue) Then
                ' This is likely a datetime that Excel converted to decimal
                ' Convert decimal datetime to proper time format
                Dim decimalDateTime As Double
                decimalDateTime = CDbl(cellValue)
                
                ' Extract just the time portion from the decimal datetime
                ' Excel stores dates as whole numbers and times as fractions
                Dim timeFraction As Double
                timeFraction = decimalDateTime - Int(decimalDateTime)
                
                ' Convert time fraction to hours, minutes, seconds
                Dim totalSeconds As Long
                Dim hours As Integer
                Dim minutes As Integer
                Dim seconds As Integer
                
                totalSeconds = CLng(timeFraction * 24 * 3600)
                hours = totalSeconds \ 3600
                minutes = (totalSeconds Mod 3600) \ 60
                seconds = totalSeconds Mod 60
                
                ' Create time string
                timeOnly = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(seconds, "00")
                
                ' Convert to proper time value
                If IsDate("1/1/1900 " & timeOnly) Then
                    timeValue = CDate("1/1/1900 " & timeOnly)
                    ws.Cells(i, 1).value = timeValue
                    ws.Cells(i, 1).NumberFormat = "hh:mm:ss"
                End If
            Else
                ' Original datetime string format
                Dim cellString As String
                cellString = CStr(cellValue)
                
                ' Find the space between date and time
                spacePos = InStr(cellString, " ")
                If spacePos > 0 Then
                    ' Find the dash before UTC offset
                    dashPos = InStr(spacePos, cellString, "-")
                    If dashPos > 0 Then
                        ' Extract time portion (between space and dash)
                        timeOnly = Mid(cellString, spacePos + 1, dashPos - spacePos - 1)
                        
                        ' Remove milliseconds (everything after the dot)
                        dotPos = InStr(timeOnly, ".")
                        If dotPos > 0 Then
                            timeOnly = Left(timeOnly, dotPos - 1)
                        End If
                        
                        ' Convert to proper time value and apply formatting
                        If IsDate("1/1/1900 " & timeOnly) Then
                            timeValue = CDate("1/1/1900 " & timeOnly)
                            ws.Cells(i, 1).value = timeValue
                            ws.Cells(i, 1).NumberFormat = "hh:mm:ss"
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in FormatDateTimeColumn: " & Err.Description, vbExclamation
End Sub

'/**
' * Processes ignition state changes and "Attempting to ignite" messages
' * Updates column B with appropriate text and color based on ignition state transitions
' */
Sub ProcessIgnitionStates()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' Constants for column positions
    Const MESSAGE_COLUMN As Integer = 30       ' Column AD (message)
    Const COLUMN_B As Integer = 2              ' Column B
    
    ' Color variables
    Dim COLOR_LIGHT_YELLOW As Long
    Dim COLOR_GREEN As Long
    Dim COLOR_RED As Long
    
    ' Initialize colors
    COLOR_LIGHT_YELLOW = RGB(255, 255, 224)  ' Light yellow for Attempt
    COLOR_GREEN = RGB(0, 255, 0)             ' Green for Ignited
    COLOR_RED = RGB(255, 0, 0)               ' Red for Flameout
    
    Dim i As Long
    Dim isIgnited As Variant
    Dim previousIgnited As Variant
    Dim messageText As String
    Dim serialNumber As String
    Dim lastCol As Long
    
    ' Get the serial number that was already added to column B
    serialNumber = ws.Cells(2, COLUMN_B).value
    
    ' Get the last column with data (calculate once before loop)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Initialize previous ignition state
    previousIgnited = False
    
    For i = 2 To lastRow
        ' Show progress every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing ignition states: Processing row " & i & " of " & lastRow & "..."
            DoEvents
        End If
        
        ' Check if we have valid data in the required columns
        If IS_IGNITED_COLUMN <= lastCol And _
           MESSAGE_COLUMN <= lastCol And _
           Not IsEmpty(ws.Cells(i, IS_IGNITED_COLUMN).value) And _
           Not IsEmpty(ws.Cells(i, MESSAGE_COLUMN).value) Then
            
            isIgnited = ws.Cells(i, IS_IGNITED_COLUMN).value
            messageText = Trim(CStr(ws.Cells(i, MESSAGE_COLUMN).value))
            
            ' Check for "Attempting to ignite" message first (highest priority)
            If UCase(messageText) = "ATTEMPTING TO IGNITE" Then
                ws.Cells(i, COLUMN_B).value = "Attempt"
                ws.Cells(i, COLUMN_B).Interior.Color = COLOR_LIGHT_YELLOW
            Else
                ' Check for ignition state changes
                If i > 2 Then  ' Need at least 2 rows to compare
                    ' Check if ignition went from true to false (flameout)
                    If (previousIgnited = True Or previousIgnited = "TRUE") And _
                       (isIgnited = False Or isIgnited = "FALSE") Then
                        ws.Cells(i, COLUMN_B).value = "Flameout"
                        ws.Cells(i, COLUMN_B).Interior.Color = COLOR_RED
                    ' Check if ignition went from false to true (ignited)
                    ElseIf (previousIgnited = False Or previousIgnited = "FALSE") And _
                           (isIgnited = True Or isIgnited = "TRUE") Then
                        ws.Cells(i, COLUMN_B).value = "Ignited"
                        ws.Cells(i, COLUMN_B).Interior.Color = COLOR_GREEN
                    Else
                        ' Keep the serial number if no state change
                        ws.Cells(i, COLUMN_B).value = serialNumber
                        ws.Cells(i, COLUMN_B).Interior.ColorIndex = xlNone
                    End If
                Else
                    ' First row - keep serial number
                    ws.Cells(i, COLUMN_B).value = serialNumber
                    ws.Cells(i, COLUMN_B).Interior.ColorIndex = xlNone
                End If
            End If
            
            ' Update previous ignition state for next iteration
            previousIgnited = isIgnited
        Else
            ' If missing data, keep the serial number
            ws.Cells(i, COLUMN_B).value = serialNumber
            ws.Cells(i, COLUMN_B).Interior.ColorIndex = xlNone
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in ProcessIgnitionStates: " & Err.Description, vbExclamation
End Sub

'/**
' * Finds firmware logs in the same folder as the current file
' * Filters by date from current filename and lists them in a new sheet
' */
Sub FindFirmwareLogs(targetWorkbook As Workbook)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Get current working directory and filename
    Dim currentPath As String
    Dim currentFileName As String
    Dim targetDate As String
    Dim serialNumber As String
    
    ' On macOS, use the original workbook path as fallback if CurDir() returns sandbox path
    currentPath = CurDir()
    currentFileName = ActiveWorkbook.Name
    
    ' Check if we're in the Excel sandbox (macOS issue)
    If InStr(currentPath, "Library/Containers/com.microsoft.Excel") > 0 Then
        ' Use the original workbook path instead
        currentPath = ActiveWorkbook.Path
    End If
    
    ' Ensure proper path separator for macOS
    If Right(currentPath, 1) <> "/" And Right(currentPath, 1) <> "\" Then
        currentPath = currentPath & "/"
    End If
    
    ' Extract date and serial number from current filename
    targetDate = ExtractDateFromFileName(currentFileName)
    serialNumber = ExtractSerialNumber()
    
    ' If no date found, use a default name
    If targetDate = "" Then
        targetDate = "UnknownDate"
    End If
    
    ' Create new worksheet for firmware logs
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "FirmwareLogs_" & targetDate
    
    ' Check if sheet already exists and delete it
    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    ' Set up headers
    ws.Cells(1, 1).value = "Filename"
    ws.Cells(1, 2).value = "Date"
    ws.Cells(1, 3).value = "Serial Number"
    ws.Cells(1, 4).value = "Log Type"
    ws.Cells(1, 5).value = "File Size (KB)"
    ws.Cells(1, 6).value = "Full Path"
    
    ' Format headers
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(200, 200, 200)
    
    ' Search for firmware logs
    Dim filePattern As String
    Dim fileName As String
    Dim filePath As String
    Dim row As Long
    Dim fileCount As Long
    
    row = 2
    fileCount = 0
    
    ' Search for files with FirmwareLog in the name
    fileName = Dir(currentPath & "*FirmwareLog*")
    
    Do While fileName <> ""
        ' Check if file matches our date and serial number pattern
        If IsFirmwareLogFile(fileName, targetDate, serialNumber) Then
            filePath = currentPath & fileName
            
            ' Extract components from filename
            Dim fileDate As String
            Dim fileSN As String
            Dim logType As String
            
            ParseFirmwareLogFileName fileName, fileDate, fileSN, logType
            
            ' Get file size
            Dim fileSize As Long
            fileSize = FileLen(filePath)
            
            ' Add to worksheet
            ws.Cells(row, 1).value = fileName
            ws.Cells(row, 2).value = fileDate
            ws.Cells(row, 3).value = fileSN
            ws.Cells(row, 4).value = logType
            ws.Cells(row, 5).value = Round(fileSize / 1024, 2)  ' Convert to KB
            ws.Cells(row, 6).value = filePath
            
            row = row + 1
            fileCount = fileCount + 1
        End If
        
        fileName = Dir()
    Loop
    
    ' Auto-fit columns
    ws.Columns("A:F").AutoFit
    
    ' Add summary information
    ws.Cells(row + 1, 1).value = "Summary:"
    ws.Cells(row + 1, 1).Font.Bold = True
    ws.Cells(row + 2, 1).value = "Date Filter:"
    ws.Cells(row + 2, 2).value = targetDate
    ws.Cells(row + 3, 1).value = "Serial Number:"
    ws.Cells(row + 3, 2).value = serialNumber
    ws.Cells(row + 4, 1).value = "Files Found:"
    ws.Cells(row + 4, 2).value = fileCount
    
    ' Activate the new sheet
    ws.Activate
    
    Application.ScreenUpdating = True
    

    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error in FindFirmwareLogs: " & Err.Description, vbExclamation
End Sub

'/**
' * Parses the contents of firmware log files and creates a new sheet with the data
' * Extracts date, time, and message from each log entry
' * Filters out entries older than 7 days
' */
Sub ParseFirmwareLogContents(targetWorkbook As Workbook)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Get current working directory and filename
    Dim currentPath As String
    Dim currentFileName As String
    Dim targetDate As String
    Dim serialNumber As String
    
    currentPath = CurDir()
    currentFileName = ActiveWorkbook.Name
    
    ' Check if we're in the Excel sandbox (macOS issue)
    If InStr(currentPath, "Library/Containers/com.microsoft.Excel") > 0 Then
        ' Use the original workbook path instead
        currentPath = ActiveWorkbook.Path
    End If
    
    ' Ensure proper path separator for macOS
    If Right(currentPath, 1) <> "/" And Right(currentPath, 1) <> "\" Then
        currentPath = currentPath & "/"
    End If
    
    ' Extract date and serial number from current filename
    targetDate = ExtractDateFromFileName(currentFileName)
    serialNumber = ExtractSerialNumber()
    
    ' If no date found, use a default name
    If targetDate = "" Then
        targetDate = "UnknownDate"
    End If
    
    ' Create new worksheet for firmware log contents
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "FirmwareLogContents_" & targetDate
    
    ' Check if sheet already exists and delete it
    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    ' Set up headers
    ws.Cells(1, 1).value = "Date"
    ws.Cells(1, 2).value = "Time"
    ws.Cells(1, 3).value = "Message"
    ws.Cells(1, 4).value = "Source File"
    
    ' Format headers
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Interior.Color = RGB(200, 200, 200)
    
    ' Calculate cutoff date (date from current filename)
    Dim cutoffDate As Date
    If targetDate <> "UnknownDate" Then
        ' Convert YYYYMMDD format to Date
        cutoffDate = DateSerial(CInt(Left(targetDate, 4)), CInt(Mid(targetDate, 5, 2)), CInt(Right(targetDate, 2)))
    Else
        ' Fallback to current date if no date found in filename
        cutoffDate = Date
    End If
    
    ' Search for firmware log files
    Dim fileName As String
    Dim filePath As String
    Dim row As Long
    Dim totalEntries As Long
    Dim filesProcessed As Long
    
    row = 2
    totalEntries = 0
    filesProcessed = 0
    
    ' Use a dynamic approach that finds files with spaces in names
    Dim fileCount As Long
    fileCount = 0
    

    
    ' Search for files with different patterns to handle spaces
    Dim searchPatterns(1 To 3) As String
    searchPatterns(1) = "*FirmwareLog.log"
    searchPatterns(2) = "*FirmwareLog *.log"  ' Pattern for files with space and number
    searchPatterns(3) = "*FirmwareLog*.log"   ' Fallback pattern
    
    Dim i As Integer
    Dim foundFiles As String
    foundFiles = ""
    
    For i = 1 To 3
        fileName = Dir(currentPath & searchPatterns(i))
        Do While fileName <> ""
            ' Check if we already found this file (avoid duplicates)
            If InStr(foundFiles, fileName) = 0 Then
                foundFiles = foundFiles & fileName & "|"
                fileCount = fileCount + 1
                
                ' Check if file matches our criteria
                If IsFirmwareLogFile(fileName, targetDate, serialNumber) Then
                    filePath = currentPath & fileName
                    filesProcessed = filesProcessed + 1
                    
                    ' Parse the log file
                    Call ParseSingleFirmwareLog(filePath, ws, row, cutoffDate, totalEntries)
                End If
            End If
            fileName = Dir()
        Loop
    Next i
    
    ' Auto-fit columns
    ws.Columns("A:D").AutoFit
    
    ' Sort the data by date (column A) and then time (column B) if we have data
    If row > 2 Then
        ' Define the range to sort (from row 2 to the last data row, columns A to D)
        Dim sortRange As Range
        Set sortRange = ws.Range("A2:D" & (row - 1))
        
        ' Sort by date (A) then time (B) in ascending order
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("A2"), Order:=xlAscending
            .SortFields.Add Key:=ws.Range("B2"), Order:=xlAscending
            .SetRange sortRange
            .Header = xlNo
            .Apply
        End With
    End If
    
    ' Add summary information
    ws.Cells(row + 1, 1).value = "Summary:"
    ws.Cells(row + 1, 1).Font.Bold = True
    ws.Cells(row + 2, 1).value = "Date Filter:"
    ws.Cells(row + 2, 2).value = targetDate
    ws.Cells(row + 3, 1).value = "Serial Number:"
    ws.Cells(row + 3, 2).value = serialNumber
    ws.Cells(row + 4, 1).value = "Total Entries:"
    ws.Cells(row + 4, 2).value = totalEntries
    ws.Cells(row + 5, 1).value = "Files Processed:"
    ws.Cells(row + 5, 2).value = filesProcessed
    ws.Cells(row + 6, 1).value = "Cutoff Date:"
    ws.Cells(row + 6, 2).value = cutoffDate
    
    ' Activate the new sheet
    ws.Activate
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error in ParseFirmwareLogContents: " & Err.Description, vbExclamation
End Sub

'/**
' * Extracts date from filename in format YYYYMMDD_SN_LogType
' */
Function ExtractDateFromFileName(fileName As String) As String
    On Error GoTo ErrorHandler
    
    Dim dateStr As String
    Dim underscorePos As Integer
    
    ' Find first underscore
    underscorePos = InStr(fileName, "_")
    
    If underscorePos > 0 Then
        ' Extract the part before first underscore
        dateStr = Left(fileName, underscorePos - 1)
        
        ' Check if it's a valid date format (8 digits)
        If Len(dateStr) = 8 And IsNumeric(dateStr) Then
            ExtractDateFromFileName = dateStr
            Exit Function
        End If
    End If
    
    ExtractDateFromFileName = ""
    Exit Function
    
ErrorHandler:
    ExtractDateFromFileName = ""
End Function

'/**
' * Checks if a file is a firmware log file matching the target date and serial number
' */
Function IsFirmwareLogFile(fileName As String, targetDate As String, serialNumber As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if filename contains "FirmwareLog"
    If InStr(1, LCase(fileName), "firmwarelog") = 0 Then
        IsFirmwareLogFile = False
        Exit Function
    End If
    
    ' Check if file has .log extension
    If LCase(Right(fileName, 4)) <> ".log" Then
        IsFirmwareLogFile = False
        Exit Function
    End If
    
    ' Check if filename starts with target date (skip if UnknownDate)
    ' The firmware log files have format: YYYYMMDD_phx42-XXXX_FirmwareLog.log
    ' So we need to check if the filename starts with the target date followed by underscore
    If targetDate <> "UnknownDate" Then
        Dim expectedPrefix As String
        expectedPrefix = targetDate & "_"
        If Left(fileName, Len(expectedPrefix)) <> expectedPrefix Then
            IsFirmwareLogFile = False
            Exit Function
        End If
    End If
    
    ' Check if filename contains the serial number (if we have one)
    If serialNumber <> "" Then
        If InStr(1, fileName, serialNumber) = 0 Then
            IsFirmwareLogFile = False
            Exit Function
        End If
    End If
    
    IsFirmwareLogFile = True
    Exit Function
    
ErrorHandler:
    IsFirmwareLogFile = False
End Function

'/**
' * Parses firmware log filename to extract date, serial number, and log type
' */
Sub ParseFirmwareLogFileName(fileName As String, ByRef fileDate As String, ByRef fileSN As String, ByRef logType As String)
    On Error GoTo ErrorHandler
    
    Dim underscorePos1 As Integer
    Dim underscorePos2 As Integer
    Dim dotPos As Integer
    
    ' Find positions of underscores
    underscorePos1 = InStr(fileName, "_")
    underscorePos2 = InStr(underscorePos1 + 1, fileName, "_")
    
    If underscorePos1 > 0 Then
        ' Extract date (first 8 characters)
        fileDate = Left(fileName, 8)
        
        If underscorePos2 > 0 Then
            ' Extract serial number (between first and second underscore)
            fileSN = Mid(fileName, underscorePos1 + 1, underscorePos2 - underscorePos1 - 1)
            
            ' Extract log type (between second underscore and dot)
            dotPos = InStrRev(fileName, ".")
            If dotPos > underscorePos2 Then
                logType = Mid(fileName, underscorePos2 + 1, dotPos - underscorePos2 - 1)
            Else
                logType = Mid(fileName, underscorePos2 + 1)
            End If
        Else
            ' Only one underscore found
            dotPos = InStrRev(fileName, ".")
            If dotPos > underscorePos1 Then
                fileSN = Mid(fileName, underscorePos1 + 1, dotPos - underscorePos1 - 1)
                logType = "Unknown"
            Else
                fileSN = Mid(fileName, underscorePos1 + 1)
                logType = "Unknown"
            End If
        End If
    Else
        fileDate = "Unknown"
        fileSN = "Unknown"
        logType = "Unknown"
    End If
    
    Exit Sub
    
ErrorHandler:
    fileDate = "Error"
    fileSN = "Error"
    logType = "Error"
End Sub

'/**
' * Parses a single firmware log file and adds entries to the worksheet
' * Extracts date, time, and message from each line
' * Filters out entries older than the cutoff date
' */
Sub ParseSingleFirmwareLog(filePath As String, ws As Worksheet, ByRef row As Long, cutoffDate As Date, ByRef totalEntries As Long)
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim lineText As String
    Dim logDate As Date
    Dim logTime As String
    Dim message As String
    Dim fileName As String
    
    ' Extract filename for source column
    fileName = Dir(filePath)
    
    ' Open the file for reading
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Debug: Show file opened successfully
    Dim linesRead As Long
    linesRead = 0
    
        ' Read each line
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        linesRead = linesRead + 1
        
        ' Skip empty lines
        If Trim(lineText) <> "" Then
            ' Try to parse the log entry
            If ParseLogEntry(lineText, logDate, logTime, message) Then
                ' Check if entry is from current date or within 7 days before
                If logDate >= cutoffDate - 7 Then
                    ' Add to worksheet
                    ws.Cells(row, 1).value = logDate
                    ws.Cells(row, 2).value = logTime
                    ws.Cells(row, 3).value = message
                    ws.Cells(row, 4).value = fileName
                    
                    ' Format date column
                    ws.Cells(row, 1).NumberFormat = "mm/dd/yyyy"
                    
                    row = row + 1
                    totalEntries = totalEntries + 1
                End If
            End If
        End If
    Loop
    
    ' Close the file
    Close fileNum
    
    ' Finished processing file
    
    Exit Sub
    
ErrorHandler:
    ' Close file if it's open
    If fileNum > 0 Then
        Close fileNum
    End If
    ' Continue processing other files
End Sub

'/**
' * Parses a single log entry line to extract date, time, and message
' * Returns True if parsing was successful
' */
Function ParseLogEntry(lineText As String, ByRef logDate As Date, ByRef logTime As String, ByRef message As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim underscorePos As Integer
    Dim spacePos As Integer
    Dim dateTimeStr As String
    Dim timeStr As String
    
    ' Look for the actual log format: YYYY/MM/DD_HH:MM:SS message
    ' Find underscore (separates date from time)
    underscorePos = InStr(lineText, "_")
    If underscorePos = 0 Then
        ParseLogEntry = False
        Exit Function
    End If
    
    ' Extract date part (before underscore)
    dateTimeStr = Left(lineText, underscorePos - 1)
    
    ' Find space (separates time from message)
    spacePos = InStr(underscorePos + 1, lineText, " ")
    If spacePos = 0 Then
        ParseLogEntry = False
        Exit Function
    End If
    
    ' Extract time part (between underscore and space)
    timeStr = Mid(lineText, underscorePos + 1, spacePos - underscorePos - 1)
    
    ' Extract message (everything after the space)
    message = Mid(lineText, spacePos + 1)
    
    ' Try to parse the date and time
    If IsDate(dateTimeStr) And IsDate("1/1/1900 " & timeStr) Then
        logDate = CDate(dateTimeStr)
        logTime = timeStr
        ParseLogEntry = True
        Exit Function
    End If
    
    ' If that didn't work, try parsing the full datetime string
    If IsDate(dateTimeStr & " " & timeStr) Then
        Dim fullDateTime As Date
        fullDateTime = CDate(dateTimeStr & " " & timeStr)
        logDate = DateValue(fullDateTime)
        logTime = timeValue(fullDateTime)
        ParseLogEntry = True
        Exit Function
    End If
    
    ParseLogEntry = False
    Exit Function
    
ErrorHandler:
    ParseLogEntry = False
End Function

'/**
' * Flags changes in LPH2 by comparing column J to column AA when column S equals 1
' * Changes text color of cells in column J when values are not very close to each other
' */
Sub FlagLPH2Changes()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    Dim solenoidValue As Variant
    Dim lph2Value As Variant
    Dim comparisonValue As Variant
    Dim difference As Double
    Dim percentDifference As Double
    
    ' Color for flagged cells (red text)
    Dim FLAG_COLOR As Long
    FLAG_COLOR = RGB(255, 0, 0)  ' Red text
    
    For i = 2 To lastRow
        ' Show progress every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Flagging LPH2 changes: Processing row " & i & " of " & lastRow & "..."
            DoEvents
        End If
        
        ' Check if we have valid data in the required columns
        If SOLENOID_COLUMN <= ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column And _
           LPH2_COLUMN <= ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column And _
           COMPARISON_COLUMN <= ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column And _
           Not IsEmpty(ws.Cells(i, SOLENOID_COLUMN).value) And _
           Not IsEmpty(ws.Cells(i, LPH2_COLUMN).value) And _
           Not IsEmpty(ws.Cells(i, COMPARISON_COLUMN).value) And _
           IsNumeric(ws.Cells(i, SOLENOID_COLUMN).value) And _
           IsNumeric(ws.Cells(i, LPH2_COLUMN).value) And _
           IsNumeric(ws.Cells(i, COMPARISON_COLUMN).value) Then
            
            solenoidValue = ws.Cells(i, SOLENOID_COLUMN).value
            lph2Value = ws.Cells(i, LPH2_COLUMN).value
            comparisonValue = ws.Cells(i, COMPARISON_COLUMN).value
            
            ' Check if solenoid is on (column S = 1)
            If solenoidValue = 1 Then
                ' Calculate the difference between column J and column AA
                difference = Abs(lph2Value - comparisonValue)
                
                ' Calculate percentage difference based on the larger value
                If comparisonValue <> 0 Then
                    percentDifference = (difference / Abs(comparisonValue)) * 100
                ElseIf lph2Value <> 0 Then
                    percentDifference = (difference / Abs(lph2Value)) * 100
                Else
                    percentDifference = 0
                End If
                
                ' Flag the cell if the difference is significant (above threshold)
                If percentDifference > LPH2_COMPARISON_THRESHOLD * 100 Then
                    ws.Cells(i, LPH2_COLUMN).Font.Color = FLAG_COLOR
                Else
                    ' Reset to default color if difference is within threshold
                    ws.Cells(i, LPH2_COLUMN).Font.Color = RGB(0, 0, 0)  ' Black
                End If
            Else
                ' Reset to default color when solenoid is off
                ws.Cells(i, LPH2_COLUMN).Font.Color = RGB(0, 0, 0)  ' Black
            End If
        Else
            ' Reset to default color if data is missing or invalid
            ws.Cells(i, LPH2_COLUMN).Font.Color = RGB(0, 0, 0)  ' Black
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in FlagLPH2Changes: " & Err.Description, vbExclamation
End Sub




' Adds a workbook-level name in the specified workbook, replacing it if it exists
Private Sub AddOrReplaceNameInWb(ByVal wb As Workbook, ByVal nameText As String, ByVal refersToFormula As String)
    On Error Resume Next
    wb.Names(nameText).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nameText, RefersTo:=refersToFormula
End Sub
