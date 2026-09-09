Attribute VB_Name = "modDispatcher"
Option Explicit

Private Const REGISTRY_START_ROW As Long = 3

Public Sub RunSelectedTransform()
    Dim controlSheet As Worksheet

    Set controlSheet = ThisWorkbook.Worksheets(RUNNER_SHEET_NAME)
    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Dispatching"
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, "Dispatcher started"
    WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Application.StatusBar = "Dispatching selected transform..."
    DoEvents

    RunConfiguredTransform ThisWorkbook, controlSheet
End Sub

Public Sub RunConfiguredTransform(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim transformLabel As String
    Dim macroName As String
    Dim ranMacro As Boolean
    Dim lastErrNumber As Long
    Dim lastErrDescription As String
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousCalculation As XlCalculation

    transformLabel = ReadControlValue(controlSheet, CONTROL_LABEL_TRANSFORM)
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, "Dispatching transform: " & transformLabel

    macroName = ResolveConfiguredMacroName(runnerWorkbook, transformLabel)
    If LenB(macroName) = 0 Then
        WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Failed"
        WriteControlValue controlSheet, RESULT_LABEL_ISSUES, "No configured macro found for transform: " & transformLabel
        HighlightRunResult controlSheet, "Failed", 0
        Err.Raise vbObjectError + 1702, "modDispatcher.RunConfiguredTransform", _
            "No configured macro found for transform: " & transformLabel
    End If

    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, "Dispatching transform: " & transformLabel & " | macro: " & macroName

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ranMacro = TryRunTransformMacro(runnerWorkbook.Name & "!" & macroName, runnerWorkbook, controlSheet, lastErrNumber, lastErrDescription)
    If Not ranMacro Then
        ranMacro = TryRunTransformMacro(macroName, runnerWorkbook, controlSheet, lastErrNumber, lastErrDescription)
    End If

    Application.Calculation = previousCalculation
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating

    If Not ranMacro Then
        WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Failed"
        WriteControlValue controlSheet, RESULT_LABEL_ISSUES, _
            "Unable to run configured macro: " & macroName & _
            IIf(LenB(lastErrDescription) = 0, vbNullString, " | " & lastErrDescription)
        WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
        HighlightRunResult controlSheet, "Failed", 0
        Err.Raise vbObjectError + 1703, "modDispatcher.RunConfiguredTransform", _
            "Unable to run configured macro: " & macroName & _
            IIf(LenB(lastErrDescription) = 0, vbNullString, " | VBA error " & CStr(lastErrNumber) & ": " & lastErrDescription)
    End If

    HighlightRunResult controlSheet, _
        ReadControlValue(controlSheet, RESULT_LABEL_RUN_STATUS), _
        CLng(Val(ReadControlValue(controlSheet, RESULT_LABEL_ROWS_SKIPPED)))
End Sub

Private Function ResolveConfiguredMacroName(ByVal runnerWorkbook As Workbook, ByVal transformLabel As String) As String
    Dim registrySheet As Worksheet
    Dim lastRegistryRow As Long
    Dim rowIndex As Long
    Dim candidateLabel As String

    Set registrySheet = GetRegistrySheet(runnerWorkbook)
    lastRegistryRow = registrySheet.Cells(registrySheet.Rows.Count, REGISTRY_TRANSFORM_LABEL_COLUMN).End(xlUp).Row
    If lastRegistryRow < REGISTRY_START_ROW Then
        Exit Function
    End If

    For rowIndex = REGISTRY_START_ROW To lastRegistryRow
        candidateLabel = Trim$(CStr(registrySheet.Cells(rowIndex, REGISTRY_TRANSFORM_LABEL_COLUMN).Value))
        If StrComp(candidateLabel, transformLabel, vbTextCompare) = 0 Then
            ResolveConfiguredMacroName = Trim$(CStr(registrySheet.Cells(rowIndex, REGISTRY_MACRO_NAME_COLUMN).Value))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function TryRunTransformMacro( _
    ByVal macroName As String, _
    ByVal runnerWorkbook As Workbook, _
    ByVal controlSheet As Worksheet, _
    ByRef capturedErrNumber As Long, _
    ByRef capturedErrDescription As String _
) As Boolean
    On Error Resume Next
    Application.Run macroName, runnerWorkbook, controlSheet
    TryRunTransformMacro = (Err.Number = 0)
    capturedErrNumber = Err.Number
    capturedErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
End Function
