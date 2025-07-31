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

' Constants for thresholds
Private Const MIN_OPERATING_TEMP As Double = 100     ' Minimum temperature to consider as operating
Private Const STEADY_STATE_SAMPLES As Integer = 5    ' Minimum samples to establish normal operating range
Private Const STEADY_STATE_THRESHOLD As Double = 0.005
Private Const BLIP_THRESHOLD As Double = 0.05
Private Const STEADY_STATE_MAX As Double = 1.3
Private Const VACUUM_GREEN_THRESHOLD As Double = -0.6
Private Const VACUUM_RED_THRESHOLD As Double = -1.0

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
        headerText = Trim(CStr(headerRow.Cells(1, col).Value))
        lowerHeader = LCase(headerText)
        
        ' Check and rename headers using direct string comparisons
        Select Case lowerHeader
            Case "pa offset"
                headerRow.Cells(1, col).Value = "Ofs"
            Case "sample pressure"
                headerRow.Cells(1, col).Value = "sPress"
            Case "sample ppl"
                headerRow.Cells(1, col).Value = "sPPL"
            Case "combustion pressure"
                headerRow.Cells(1, col).Value = "cPress"
            Case "combustion ppl"
                headerRow.Cells(1, col).Value = "cPPL"
            Case "internal temp."
                headerRow.Cells(1, col).Value = "iTemp"
            Case "external temp."
                headerRow.Cells(1, col).Value = "eTemp"
            Case "case temp."
                headerRow.Cells(1, col).Value = "cTemp"
            Case "needle valve"
                headerRow.Cells(1, col).Value = "MOV"
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
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
                cellValue = ws.Cells(row, col).Value
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
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
           Trim(CStr(ws.Cells(i, 6).Value)) = "" Or _
           ws.Cells(i, 6).Value = "" Then
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
                cellValue = ws.Cells(i, j).Value
                
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
    Call CreateBackupSheet
    
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
    
    ' Auto-size all columns
    Application.StatusBar = "Auto-sizing columns..."
    DoEvents
    Call AutoSizeAllColumns
    
    ' Save as Excel file
    Application.StatusBar = "Saving processed file..."
    DoEvents
    Call SaveAsExcelFile
    
    ' Show completion message
    Application.StatusBar = "Processing complete!"
    Application.Wait Now + TimeValue("00:00:02")  ' Show completion for 2 seconds
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
Sub CreateBackupSheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Check if RAW sheet already exists and delete it
    Dim rawSheet As Worksheet
    On Error Resume Next
    Set rawSheet = ThisWorkbook.Worksheets("RAW")
    On Error GoTo ErrorHandler
    
    If Not rawSheet Is Nothing Then
        Application.DisplayAlerts = False
        rawSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new RAW sheet
    Dim newSheet As Worksheet
    Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newSheet.Name = "RAW"
    
    ' Copy all data from active sheet to RAW sheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
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
           IsNumeric(ws.Cells(i, FLAMEOUT_COLUMN).Value) And _
           Not IsEmpty(ws.Cells(i, SOLENOID_COLUMN).Value) And _
           Not IsEmpty(ws.Cells(i, IS_IGNITED_COLUMN).Value) And _
           ws.Cells(i, FLAMEOUT_COLUMN).Value <> "" And _
           ws.Cells(i, SOLENOID_COLUMN).Value <> "" And _
           ws.Cells(i, IS_IGNITED_COLUMN).Value <> "" Then
            
            value = ws.Cells(i, FLAMEOUT_COLUMN).Value
            Dim solenoid As Variant
            solenoid = ws.Cells(i, SOLENOID_COLUMN).Value
            isIgnited = ws.Cells(i, IS_IGNITED_COLUMN).Value
            
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
                           IsNumeric(ws.Cells(startRow - lookBack, FLAMEOUT_COLUMN).Value) Then
                            steadyStateTemp = steadyStateTemp + ws.Cells(startRow - lookBack, FLAMEOUT_COLUMN).Value
                            steadyStateCount = steadyStateCount + 1
                        End If
                    Next lookBack
                    
                    If steadyStateCount > 0 Then
                        steadyStateTemp = steadyStateTemp / steadyStateCount
                    End If
                    
                    ' Look back to find where temperature first started dropping
                    Do While startRow > 2 And lookBack < 50  ' Look back up to 50 rows
                        If startRow <= lastRow And startRow > 1 And _
                           IsNumeric(ws.Cells(startRow, FLAMEOUT_COLUMN).Value) And _
                           IsNumeric(ws.Cells(startRow - 1, FLAMEOUT_COLUMN).Value) Then
                            
                            Dim currentTemp As Double
                            Dim previousTemp As Double
                            currentTemp = ws.Cells(startRow, FLAMEOUT_COLUMN).Value
                            previousTemp = ws.Cells(startRow - 1, FLAMEOUT_COLUMN).Value
                            
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
                    peak = ws.Cells(startRow, FLAMEOUT_COLUMN).Value
                    flameoutStartRow = startRow + 1
                    
                    ' Highlight from where temperature first started dropping to current row
                    For k = flameoutStartRow To i
                        If IsNumeric(ws.Cells(k, FLAMEOUT_COLUMN).Value) Then
                            Dim valueK As Double
                            valueK = ws.Cells(k, FLAMEOUT_COLUMN).Value
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
    filePath = ActiveWorkbook.Path & "\"
    
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
    
    ' Try to save the workbook
    Application.DisplayAlerts = False
    
    ' First try to save as new file
    On Error Resume Next
    ActiveWorkbook.SaveAs fullPath, xlOpenXMLWorkbook
    If Err.Number = 0 Then
        Application.DisplayAlerts = True
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' If that fails, try saving to desktop
    Dim desktopPath As String
    desktopPath = Environ("USERPROFILE") & "\Desktop\" & newName
    
    On Error Resume Next
    ActiveWorkbook.SaveAs desktopPath, xlOpenXMLWorkbook
    If Err.Number = 0 Then
        Application.DisplayAlerts = True
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' If that also fails, try saving to Downloads folder
    Dim downloadsPath As String
    downloadsPath = Environ("USERPROFILE") & "\Downloads\" & newName
    
    On Error Resume Next
    ActiveWorkbook.SaveAs downloadsPath, xlOpenXMLWorkbook
    If Err.Number = 0 Then
        Application.DisplayAlerts = True
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' If all else fails, just save the current workbook
    Application.DisplayAlerts = True
    MsgBox "Could not save as new file. Saving current workbook instead.", vbExclamation
    ActiveWorkbook.Save
    
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
        lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        
        ' Fill column B with serial number for all rows
        Dim serialText As String
        serialText = "phx42-" & serialNumber
        
        ' Apply to all rows from 1 to lastRow
        ActiveSheet.Range("B1:B" & lastRow).Value = serialText
        
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
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
            cellValue = ws.Cells(i, 1).Value
            
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
                    ws.Cells(i, 1).Value = timeValue
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
                            ws.Cells(i, 1).Value = timeValue
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
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
    serialNumber = ws.Cells(2, COLUMN_B).Value
    
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
           Not IsEmpty(ws.Cells(i, IS_IGNITED_COLUMN).Value) And _
           Not IsEmpty(ws.Cells(i, MESSAGE_COLUMN).Value) Then
            
            isIgnited = ws.Cells(i, IS_IGNITED_COLUMN).Value
            messageText = Trim(CStr(ws.Cells(i, MESSAGE_COLUMN).Value))
            
            ' Check for "Attempting to ignite" message first (highest priority)
            If UCase(messageText) = "ATTEMPTING TO IGNITE" Then
                ws.Cells(i, COLUMN_B).Value = "Attempt"
                ws.Cells(i, COLUMN_B).Interior.Color = COLOR_LIGHT_YELLOW
            Else
                ' Check for ignition state changes
                If i > 2 Then  ' Need at least 2 rows to compare
                    ' Check if ignition went from true to false (flameout)
                    If (previousIgnited = True Or previousIgnited = "TRUE") And _
                       (isIgnited = False Or isIgnited = "FALSE") Then
                        ws.Cells(i, COLUMN_B).Value = "Flameout"
                        ws.Cells(i, COLUMN_B).Interior.Color = COLOR_RED
                    ' Check if ignition went from false to true (ignited)
                    ElseIf (previousIgnited = False Or previousIgnited = "FALSE") And _
                           (isIgnited = True Or isIgnited = "TRUE") Then
                        ws.Cells(i, COLUMN_B).Value = "Ignited"
                        ws.Cells(i, COLUMN_B).Interior.Color = COLOR_GREEN
                    Else
                        ' Keep the serial number if no state change
                        ws.Cells(i, COLUMN_B).Value = serialNumber
                        ws.Cells(i, COLUMN_B).Interior.ColorIndex = xlNone
                    End If
                Else
                    ' First row - keep serial number
                    ws.Cells(i, COLUMN_B).Value = serialNumber
                    ws.Cells(i, COLUMN_B).Interior.ColorIndex = xlNone
                End If
            End If
            
            ' Update previous ignition state for next iteration
            previousIgnited = isIgnited
        Else
            ' If missing data, keep the serial number
            ws.Cells(i, COLUMN_B).Value = serialNumber
            ws.Cells(i, COLUMN_B).Interior.ColorIndex = xlNone
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in ProcessIgnitionStates: " & Err.Description, vbExclamation
End Sub