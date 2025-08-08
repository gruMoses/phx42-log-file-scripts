Attribute VB_Name = "modChartLogging"
Option Explicit

' Logging globals
Private g_LoggingEnabled As Boolean
Private g_LogFilePath As String
Public g_Step As String

Public Sub EnableChartDebugLogging(Optional ByVal logFolder As String)
    On Error Resume Next
    g_LoggingEnabled = True
    Dim sep As String: sep = Application.PathSeparator
    Dim folder As String
    If logFolder <> "" Then
        folder = logFolder
    ElseIf Not ActiveWorkbook Is Nothing And ActiveWorkbook.Path <> "" Then
        folder = ActiveWorkbook.Path
    Else
        folder = CurDir()
    End If
    If Right$(folder, 1) <> sep Then folder = folder & sep
    g_LogFilePath = folder & "phx42_chart_debug.log"

    Dim ff As Integer: ff = FreeFile
    Open g_LogFilePath For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss"), "--- logging started ---"
    Close #ff
End Sub

Public Sub LogDebug(ByVal message As String)
    On Error Resume Next
    If Not g_LoggingEnabled Then Exit Sub
    If g_LogFilePath = "" Then Exit Sub
    Dim ff As Integer: ff = FreeFile
    Open g_LogFilePath For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss"), message
    Close #ff
End Sub

Public Sub DisableChartDebugLogging()
    g_LoggingEnabled = False
End Sub

Public Function GetActiveWorkbookPath() As String
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        If ActiveWorkbook.Path <> "" Then
            GetActiveWorkbookPath = ActiveWorkbook.Path
        Else
            GetActiveWorkbookPath = CurDir()
        End If
    Else
        GetActiveWorkbookPath = CurDir()
    End If
End Function

Public Sub StepTag(ByVal tag As String)
    g_Step = tag
    On Error Resume Next
    LogDebug "STEP: " & tag
End Sub
