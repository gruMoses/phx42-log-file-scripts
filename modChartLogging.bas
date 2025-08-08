Attribute VB_Name = "modChartLogging"
Option Explicit

' Logging globals
Private g_LoggingEnabled As Boolean
Private g_LogFilePath As String
Public g_Step As String

Public Sub EnableChartDebugLogging(Optional ByVal logFolder As String)
    g_LoggingEnabled = False
    g_LogFilePath = ""
End Sub

Public Sub LogDebug(ByVal message As String)
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
