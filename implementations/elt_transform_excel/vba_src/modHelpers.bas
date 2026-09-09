Attribute VB_Name = "modHelpers"
Option Explicit

Public Const CONTROL_LABEL_TRANSFORM As String = "Transform"
Public Const CONTROL_LABEL_SOURCE_TYPE As String = "Source Type"
Public Const CONTROL_LABEL_SOURCE_DATA As String = "Source Data"
Public Const CONTROL_LABEL_TARGET_EIB As String = "Target EIB"
Public Const CONTROL_LABEL_OUTPUT_FILE_PATH As String = "Output File Path"
Public Const CONTROL_LABEL_MAPPING_SPREADSHEET As String = "Mapping Spreadsheet"
Public Const CONTROL_LABEL_MAPPING_SHEET As String = "Mapping Sheet"
Public Const CONTROL_LABEL_CUSTOMER_SPREADSHEET As String = "Customer Spreadsheet"
Public Const CONTROL_LABEL_SUPPLIER_SPREADSHEET As String = "Supplier Spreadsheet"
Public Const CONTROL_LABEL_WORKDAY_REFERENCE_IDS As String = "Workday Reference IDs"
Public Const CONTROL_LABEL_EIB_POPULATION_TYPE As String = "EIB Population Type"
Public Const CONTROL_LABEL_SOURCE_WORKSHEET_NAME As String = "Source Worksheet Name"
Public Const CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION As String = "Source Worksheet Resolution"
Public Const CONTROL_LABEL_SOURCE_HEADER_ROW As String = "Source Header Row"
Public Const CONTROL_LABEL_SOURCE_DATA_START_ROW As String = "Source Data Start Row"
Public Const CONTROL_LABEL_REQUIRED_HEADERS As String = "Required Headers Pipe Delimited"

Public Const RESULT_LABEL_RUN_STATUS As String = "Run Status"
Public Const RESULT_LABEL_ROWS_PROCESSED As String = "Rows Processed"
Public Const RESULT_LABEL_ROWS_SKIPPED As String = "Rows Skipped"
Public Const RESULT_LABEL_ISSUES As String = "Issues"
Public Const RESULT_LABEL_OUTPUT_WORKBOOK As String = "Output Workbook"
Public Const RESULT_LABEL_SELECTION_KEY As String = "Selection Key"
Public Const RESULT_LABEL_LAST_RUN_AT As String = "Last Run At"

Public Const MAPPING_LOOKUP_SHEET_NAME As String = "mapping"
Public Const REGISTRY_SHEET_NAME As String = "registry"
Public Const RUNNER_SHEET_NAME As String = "Runner"

Public Const REGISTRY_TRANSFORM_LABEL_COLUMN As Long = 1
Public Const REGISTRY_PROFILE_KEY_COLUMN As Long = 2
Public Const REGISTRY_SOURCE_TYPE_COLUMN As Long = 3
Public Const REGISTRY_MACRO_NAME_COLUMN As Long = 13
Public Const REGISTRY_START_ROW As Long = 3

' Reject-log path for the in-flight run, resolved by BeginRun to sit beside the run's output EIB
' (see BuildRejectLogPath). Empty until BeginRun sets it; RejectLogPath falls back to the Runner
' folder when empty so rejects are never silently lost.
Private mRejectLogPath As String

' Live runtime lookup against the Runner workbook's "mapping" sheet (entity/source/target
' columns) -- reference data such as company-entity names that can change independently of
' the transform's column-mapping shape, unlike the column-to-column rules baked into each
' transform module. Returns vbNullString if the sheet is missing or no row matches; callers
' decide their own fallback.
Public Function LookupMappingValue(ByVal runnerWorkbook As Workbook, ByVal entityName As String, ByVal sourceValue As String) As String
    Dim mappingSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    On Error Resume Next
    Set mappingSheet = runnerWorkbook.Worksheets(MAPPING_LOOKUP_SHEET_NAME)
    On Error GoTo 0
    If mappingSheet Is Nothing Then
        LookupMappingValue = vbNullString
        Exit Function
    End If

    lastRow = mappingSheet.Cells(mappingSheet.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(mappingSheet.Cells(rowIndex, 1).Value)), entityName, vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(mappingSheet.Cells(rowIndex, 2).Value)), sourceValue, vbTextCompare) = 0 Then
            LookupMappingValue = Trim$(CStr(mappingSheet.Cells(rowIndex, 3).Value))
            Exit Function
        End If
    Next rowIndex

    LookupMappingValue = vbNullString
End Function

Public Function CloneTemplateWorkbook(ByVal templatePath As String, Optional ByVal outputPath As String = vbNullString) As Workbook
    Dim resolvedOutputPath As String

    EnsureFilePath templatePath, CONTROL_LABEL_TARGET_EIB

    resolvedOutputPath = outputPath
    If LenB(resolvedOutputPath) = 0 Then
        resolvedOutputPath = BuildSiblingFilePath(templatePath, "_output")
    End If

    If FileExists(resolvedOutputPath) Then
        Kill resolvedOutputPath
    End If

    FileCopy templatePath, resolvedOutputPath
    Set CloneTemplateWorkbook = Workbooks.Open(Filename:=resolvedOutputPath, ReadOnly:=False)
End Function

' Clear previously-published data rows from a cloned Target EIB sheet before writing a fresh run.
' UsedRange (not a column-A scan) reflects the sheet's true populated extent: the raw template can
' carry sample placeholder text in scattered columns (e.g. KE/KF) well below where column A's own
' content ends. lastColumnLetters is the rightmost column the transform writes, so the clear spans
' every output column even when the template's own UsedRange is narrower.
Public Sub ClearPublishedRows(ByVal targetSheet As Worksheet, ByVal targetDataStartRow As Long, ByVal lastColumnLetters As String)
    Dim lastRow As Long
    Dim lastColumn As Long

    lastRow = targetSheet.UsedRange.Row + targetSheet.UsedRange.Rows.Count - 1
    If lastRow < targetDataStartRow Then
        lastRow = targetDataStartRow
    End If

    lastColumn = Application.Max(targetSheet.UsedRange.Columns.Count, ColumnIndex(lastColumnLetters))
    targetSheet.Range(targetSheet.Cells(targetDataStartRow, 1), targetSheet.Cells(lastRow, lastColumn)).ClearContents
End Sub

Public Function GetRegistrySheet(ByVal runnerWorkbook As Workbook) As Worksheet
    On Error Resume Next
    Set GetRegistrySheet = runnerWorkbook.Worksheets(REGISTRY_SHEET_NAME)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError + 1609, "modHelpers.GetRegistrySheet", _
            "Runner workbook is missing the '" & REGISTRY_SHEET_NAME & "' sheet. " & _
            "Rebuild the Runner with: uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project " & _
            "python elt_transform_excel/scripts/build_runner.py"
    End If
    On Error GoTo 0
End Function

Public Function ResolveSelectedTransformKey(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet) As String
    Dim transformLabel As String
    Dim registrySheet As Worksheet

    transformLabel = ReadControlValue(controlSheet, CONTROL_LABEL_TRANSFORM)
    If LenB(transformLabel) = 0 Then
        ResolveSelectedTransformKey = ComposeSelectionKey( _
            ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_TYPE), _
            ReadControlValue(controlSheet, CONTROL_LABEL_EIB_POPULATION_TYPE) _
        )
        Exit Function
    End If

    Set registrySheet = GetRegistrySheet(runnerWorkbook)
    ResolveSelectedTransformKey = ResolveTransformRegistryValue( _
        registrySheet, _
        transformLabel, _
        REGISTRY_PROFILE_KEY_COLUMN _
    )
    If LenB(ResolveSelectedTransformKey) = 0 Then
        Err.Raise vbObjectError + 1608, "modHelpers.ResolveSelectedTransformKey", _
            "No transform profile key found for selected transform: " & transformLabel
    End If
End Function

Public Function ComposeSelectionKey(ByVal sourceType As String, ByVal eibPopulationType As String) As String
    ComposeSelectionKey = Slugify(sourceType) & "__" & Slugify(eibPopulationType)
End Function

Public Function ResolveTransformRegistryValue( _
    ByVal registrySheet As Worksheet, _
    ByVal transformLabel As String, _
    ByVal valueColumn As Long _
) As String
    Dim lastRegistryRow As Long
    Dim rowIndex As Long
    Dim candidateLabel As String

    lastRegistryRow = registrySheet.Cells(registrySheet.Rows.Count, REGISTRY_TRANSFORM_LABEL_COLUMN).End(xlUp).Row
    If lastRegistryRow < REGISTRY_START_ROW Then
        Exit Function
    End If

    For rowIndex = REGISTRY_START_ROW To lastRegistryRow
        candidateLabel = Trim$(CStr(registrySheet.Cells(rowIndex, REGISTRY_TRANSFORM_LABEL_COLUMN).Value))
        If StrComp(candidateLabel, transformLabel, vbTextCompare) = 0 Then
            ResolveTransformRegistryValue = Trim$(CStr(registrySheet.Cells(rowIndex, valueColumn).Value))
            Exit Function
        End If
    Next rowIndex
End Function

' Plain Workbooks.Open on a .csv hands column typing to Excel's own CSV importer, which
' guesses each column's type from the active regional date/number settings -- not from any
' source_format declared in the mapping workbook. A value like "16.05.26" is invalid as
' m.d.yy (day 16 can't be a month), and Excel's importer responds by silently coercing the
' cell to a Date with serial 0 rather than leaving it as text, which downstream renders as
' the Excel epoch (1899-12-30). Opening CSVs via OpenText with every column forced to Text
' guarantees the raw string survives untouched for the transform's own manual parsing
' (NormalizeSourceDate, ResolveAmount, etc.) regardless of the host machine's locale.
Public Function OpenWorkbookReadOnly(ByVal filePath As String) As Workbook
    EnsureFilePath filePath, "Workbook Path"

    If LCase$(FileExtension(filePath)) = "csv" Then
        Set OpenWorkbookReadOnly = OpenCsvAsText(filePath)
    Else
        Set OpenWorkbookReadOnly = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    End If
End Function

Private Function OpenCsvAsText(ByVal filePath As String) As Workbook
    Const MAX_CSV_COLUMNS As Long = 100
    Dim fieldInfo(1 To MAX_CSV_COLUMNS) As Variant
    Dim columnIndex As Long

    For columnIndex = 1 To MAX_CSV_COLUMNS
        fieldInfo(columnIndex) = Array(columnIndex, xlTextFormat)
    Next columnIndex

    Workbooks.OpenText Filename:=filePath, DataType:=xlDelimited, Comma:=True, FieldInfo:=fieldInfo
    Set OpenCsvAsText = ActiveWorkbook
End Function

Public Function ReadControlValue(ByVal controlSheet As Worksheet, ByVal labelText As String) As String
    Dim labelCell As Range

    Set labelCell = FindLabelCell(controlSheet, labelText)
    If labelCell Is Nothing Then
        Err.Raise vbObjectError + 1601, "modHelpers.ReadControlValue", "Control label not found: " & labelText
    End If

    ReadControlValue = Trim$(CStr(labelCell.Offset(0, 1).Value))
End Function

' Parse the CONTROL_LABEL_REQUIRED_HEADERS value as written by build_runner.py
' into a 1D VBA String array suitable for direct assignment to requiredHeaders.
' Separator between entries is double-pipe "||" because each individual required
' header entry may itself contain single-pipe "|" as an alias list (e.g.
' "Total Claim ()|Total Claim|Total Claim (GBP)").  Single pipe inside an entry
' is LEFT INTACT so SplitHeaderAliases / FindColumnIndexInRow can resolve the
' alias list at runtime.  Empty entries are skipped; leading/trailing spaces per
' entry are trimmed.
Public Function ParseRequiredHeadersFromControl(ByVal controlText As String) As Variant
    Dim s As String
    s = Trim$(controlText)
    If LenB(s) = 0 Then
        Dim outEmpty(0 To 0) As String
        outEmpty(0) = vbNullString
        ParseRequiredHeadersFromControl = outEmpty
        Exit Function
    End If
    Dim parts() As String
    parts = Split(s, "||")
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1
    Dim result() As String
    ReDim result(0 To n - 1) As String
    Dim j As Long
    j = 0
    Dim i As Long
    Dim entry As String
    For i = LBound(parts) To UBound(parts)
        entry = Trim$(parts(i))
        If LenB(entry) > 0 Then
            result(j) = entry
            j = j + 1
        End If
    Next i
    If j = 0 Then
        Dim outEmpty2(0 To 0) As String
        outEmpty2(0) = vbNullString
        ParseRequiredHeadersFromControl = outEmpty2
        Exit Function
    End If
    If j - 1 < UBound(result) Then
        ReDim Preserve result(0 To j - 1) As String
    End If
    ParseRequiredHeadersFromControl = result
End Function

' Convenience wrapper for a per-transform module: reads CONTROL_LABEL_REQUIRED_HEADERS
' from the Config sheet, parses it via ParseRequiredHeadersFromControl, validates a
' minimum entry count, and raises a descriptive error that tells the maintainer exactly
' what to run (build_runner.py from repo) when the Config sheet is stale / blank /
' insufficiently populated.  minExpectedCount is the number of headers the transform
' expects (used only for the error message so a typo in the YAML is caught on first
' run against the source sheet rather than after an hour of debugging "why column X
' unresolved").  moduleNameForErr is the bas module name (e.g. "modApStationerySupplierInvoice")
' for a precise error source.
'
' Parameter typed Object (rather than Worksheet) and ByRef (the VBA default for object
' refs) because early-bound Worksheet typed as ByVal on a public cross-module function
' can trigger "Type mismatch" compile errors on Mac Excel VBA after the modules are
' re-imported via VBComponents.Import (the Excel type library is rebound mid-import).
' The caller always passes a Worksheet; inside, we TypeOf-check for a clear error if
' someone accidentally passes the wrong thing (e.g. a Workbook).
Public Function ResolveRequiredHeadersFromConfig( _
    ByRef controlSheet As Object, _
    ByVal moduleNameForErr As String, _
    ByVal minExpectedCount As Long _
) As Variant
    Dim raw As String
    Dim arr As Variant
    If (controlSheet Is Nothing) Then
        Err.Raise vbObjectError + 2010, "modHelpers.ResolveRequiredHeadersFromConfig", _
            "controlSheet was Nothing.  Caller must pass the Config sheet Worksheet."
    End If
    If Not (TypeOf controlSheet Is Worksheet) Then
        Err.Raise vbObjectError + 2010, "modHelpers.ResolveRequiredHeadersFromConfig", _
            "controlSheet was not a Worksheet (TypeOf=" & TypeName(controlSheet) & _
                ").  Caller must pass the Config sheet Worksheet object."
    End If
    raw = ReadControlValue(controlSheet, CONTROL_LABEL_REQUIRED_HEADERS)
    arr = ParseRequiredHeadersFromControl(raw)
    Dim ok As Boolean
    ok = True
    If (Not IsArray(arr)) Then
        ok = False
    Else
        Dim lb As Long, ub As Long
        lb = LBound(arr)
        ub = UBound(arr)
        If (ub - lb + 1) < minExpectedCount Then ok = False
        If LenB(Trim$(CStr(arr(lb)))) = 0 Then ok = False
    End If
    If Not ok Then
        Err.Raise vbObjectError + 2011, moduleNameForErr, _
            "Required headers list was empty or shorter than the expected minimum " _
                & "(" & CStr(minExpectedCount) & ") after parsing Config '" _
                & CONTROL_LABEL_REQUIRED_HEADERS & "' = " & raw _
                & ".  The source of truth is config/transforms/*.yaml field " _
                & "source_layout.required_headers; re-run build_runner.py from the repo " _
                & "to re-write the Config sheet."
    End If
    ResolveRequiredHeadersFromConfig = arr
End Function

' Colour the run-result cells so a run that dropped rows is impossible to miss on the control
' sheet next to the Run button: GREEN when every row wrote, AMBER when rows were rejected
' (mapping/lookup misses -- see the <output>.rejects.csv beside the output EIB), RED when
' the run failed outright. Best-effort and cosmetic only: never raises.
Public Sub HighlightRunResult(ByVal controlSheet As Worksheet, ByVal runStatus As String, ByVal rowsSkipped As Long)
    Dim statusCell As Range
    Dim skippedCell As Range
    Dim issuesCell As Range

    On Error GoTo CleanUp
    Set statusCell = ResultValueCell(controlSheet, RESULT_LABEL_RUN_STATUS)
    Set skippedCell = ResultValueCell(controlSheet, RESULT_LABEL_ROWS_SKIPPED)
    Set issuesCell = ResultValueCell(controlSheet, RESULT_LABEL_ISSUES)

    If StrComp(runStatus, "Failed", vbTextCompare) = 0 Then
        SetCellFill statusCell, RGB(255, 150, 150), True
        SetCellFill issuesCell, RGB(255, 150, 150), False
        SetCellFill skippedCell, -1, False
    ElseIf rowsSkipped > 0 Then
        SetCellFill statusCell, RGB(255, 226, 130), True
        SetCellFill skippedCell, RGB(255, 226, 130), True
        SetCellFill issuesCell, RGB(255, 226, 130), False
    Else
        SetCellFill statusCell, RGB(198, 239, 206), False
        SetCellFill skippedCell, -1, False
        SetCellFill issuesCell, -1, False
    End If
    Exit Sub

CleanUp:
End Sub

Private Function ResultValueCell(ByVal controlSheet As Worksheet, ByVal labelText As String) As Range
    Dim labelCell As Range

    Set labelCell = FindLabelCell(controlSheet, labelText)
    If labelCell Is Nothing Then Exit Function
    Set ResultValueCell = labelCell.Offset(0, 1)
End Function

' Apply (or clear, when fillColor = -1) a background fill to a result cell. Bold marks the
' cells that most need attention (status, reject count). Never raises.
Private Sub SetCellFill(ByVal targetCell As Range, ByVal fillColor As Long, ByVal makeBold As Boolean)
    If targetCell Is Nothing Then Exit Sub
    On Error Resume Next
    If fillColor = -1 Then
        targetCell.Interior.ColorIndex = xlNone
    Else
        targetCell.Interior.Color = fillColor
    End If
    targetCell.Font.Bold = makeBold
    On Error GoTo 0
End Sub

' UsedRange can be inflated far past the real data when a source workbook carries stray
' formatting or previously-cleared cells (observed: an ~800-row "New CC Updated" sheet whose
' UsedRange spanned the full 1,048,576-row height). Reading UsedRange.Value directly then
' allocates a multi-million-cell array and forces every caller to loop empty rows to UBound.
' Bound the read to the last cell that actually contains a value. The returned array is 1-based
' from A1; callers detect the header row by content (FindHeaderRow), so the origin is stable.
Public Function ReadWorksheetTable(ByVal sourceSheet As Worksheet) As Variant
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim cellValues As Variant

    lastRow = LastDataRow(sourceSheet)
    lastColumn = LastDataColumn(sourceSheet)
    If lastRow = 0 Or lastColumn = 0 Then
        ReadWorksheetTable = BuildSingleCellTable(vbNullString)
        Exit Function
    End If

    Set dataRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastColumn))
    cellValues = dataRange.Value
    ReadWorksheetTable = CoerceRangeValueToTable(cellValues, dataRange.Rows.Count, dataRange.Columns.Count)
End Function

' Excel's Find (starting after A1, searching backwards) returns the last populated cell via the
' sheet's internal index rather than scanning cells, so it stays fast even on a bloated UsedRange.
' LookIn:=xlFormulas with What:="*" matches any non-empty cell, including purely numeric ones.
Private Function LastDataRow(ByVal sourceSheet As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = sourceSheet.Cells.Find( _
        What:="*", After:=sourceSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If foundCell Is Nothing Then
        LastDataRow = 0
    Else
        LastDataRow = foundCell.Row
    End If
End Function

Private Function LastDataColumn(ByVal sourceSheet As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = sourceSheet.Cells.Find( _
        What:="*", After:=sourceSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    If foundCell Is Nothing Then
        LastDataColumn = 0
    Else
        LastDataColumn = foundCell.Column
    End If
End Function

Public Sub WriteControlValue(ByVal controlSheet As Worksheet, ByVal labelText As String, ByVal value As String)
    Dim labelCell As Range

    Set labelCell = FindOrCreateLabelCell(controlSheet, labelText)
    labelCell.Offset(0, 1).Value = value
End Sub

Public Function ResolveSourceWorksheet(ByVal sourceWorkbook As Workbook, ByVal sourceSheetName As String) As Worksheet
    Dim currentSheet As Worksheet
    Dim existingVisibleList As String
    Dim existingAllList As String
    Dim sh As Worksheet

    If LenB(sourceSheetName) > 0 Then
        On Error Resume Next
        Set ResolveSourceWorksheet = sourceWorkbook.Worksheets(sourceSheetName)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            existingVisibleList = vbNullString
            existingAllList = vbNullString
            For Each sh In sourceWorkbook.Worksheets
                If LenB(existingAllList) = 0 Then
                    existingAllList = "'" & sh.Name & "'"
                Else
                    existingAllList = existingAllList & ", '" & sh.Name & "'"
                End If
                If sh.Visible = xlSheetVisible Then
                    If LenB(existingVisibleList) = 0 Then
                        existingVisibleList = "'" & sh.Name & "'"
                    Else
                        existingVisibleList = existingVisibleList & ", '" & sh.Name & "'"
                    End If
                End If
            Next sh
            If LenB(existingVisibleList) = 0 Then existingVisibleList = "(none - all sheets are hidden)"
            Err.Raise vbObjectError + 1605, "modHelpers.ResolveSourceWorksheet", _
                "Expected worksheet '" & sourceSheetName & "' was not found in workbook '" & sourceWorkbook.Name & "'." & vbCrLf & _
                "  Visible sheets in that workbook: " & existingVisibleList & vbCrLf & _
                "  All sheets in that workbook    : " & existingAllList & vbCrLf & _
                "  Hint: check that you selected the correct source file for this transform profile, or run `validate_input.py --transform <profile_key>` from the CLI for a full 12-point schema check report."
        End If
        On Error GoTo 0
        Exit Function
    End If

    For Each currentSheet In sourceWorkbook.Worksheets
        If currentSheet.Visible = xlSheetVisible Then
            Set ResolveSourceWorksheet = currentSheet
            Exit Function
        End If
    Next currentSheet

    Err.Raise vbObjectError + 1604, "modHelpers.ResolveSourceWorksheet", "No visible worksheet found in " & sourceWorkbook.Name
End Function

Public Function JoinCollectionValues(ByVal values As Collection, ByVal delimiter As String) As String
    Dim result As String
    Dim itemValue As Variant

    For Each itemValue In values
        If LenB(result) = 0 Then
            result = CStr(itemValue)
        Else
            result = result & delimiter & CStr(itemValue)
        End If
    Next itemValue

    JoinCollectionValues = result
End Function

Public Sub EnsureFilePath(ByVal filePath As String, ByVal labelName As String)
    If LenB(Trim$(filePath)) = 0 Then
        Err.Raise vbObjectError + 1602, "modHelpers", labelName & " is blank."
    End If

    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1603, "modHelpers", labelName & " does not exist: " & filePath
    End If
End Sub

Public Sub EnsureNonBlank(ByVal value As String, ByVal labelName As String)
    If LenB(Trim$(value)) = 0 Then
        Err.Raise vbObjectError + 1607, "modHelpers", labelName & " is blank."
    End If
End Sub

' --- Run status, progress, and external run log (shared across all transforms) ---

' Make the run visibly alive before the single-threaded work blocks the UI thread. Excel
' cannot stay interactive during a macro, but writing "Running" and flushing one repaint,
' plus a live status bar during the loop, tells a business user it is working rather than
' hung. Every Application state set here MUST be restored by EndRun on all exit paths.
Public Sub BeginRun(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet, ByVal transformKey As String)
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.StatusBar = transformKey & ": preparing..."
    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Running"
    WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WriteRunStatusFile runnerWorkbook, transformKey & ": preparing... (" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ")"
    mRejectLogPath = BuildRejectLogPath(controlSheet)
    ResetRejectLog runnerWorkbook, transformKey
    ' Flush the "Running" cell to screen before the heavy work starts.
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
End Sub

' Called periodically inside the row loop. Updates the status bar (for the workbook itself)
' and overwrites RunnerStatus.txt (an out-of-band channel a user can tail/refresh while the
' grid is frozen). Both are best-effort; neither blocks the run.
Public Sub UpdateRunProgress(ByVal runnerWorkbook As Workbook, ByVal transformKey As String, ByVal currentRow As Long, ByVal lastRow As Long, ByVal startTime As Date)
    Dim message As String
    message = transformKey & ": row " & Format$(currentRow, "#,##0") & " of " & Format$(lastRow, "#,##0") & _
        " (" & Format$((Now - startTime) * 86400#, "0") & "s elapsed)"
    Application.StatusBar = message & "..."
    WriteRunStatusFile runnerWorkbook, message
    DoEvents
End Sub

Public Sub EndRun()
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
End Sub

' Append one row per run to RunnerLog.csv beside the Runner workbook. Native VBA file I/O
' (no FileSystemObject); the file survives rebuilds and the user can delete it. Best-effort:
' returns an issue string instead of raising, so a log-write failure never breaks the run.
Public Function AppendRunLog( _
    ByVal runnerWorkbook As Workbook, _
    ByVal transformKey As String, _
    ByVal status As String, _
    ByVal rowsProcessed As Long, _
    ByVal rowsRejected As Long, _
    ByVal startTime As Date, _
    ByVal endTime As Date, _
    ByVal outputPath As String, _
    ByVal issues As String _
) As String
    Dim logPath As String
    Dim fileNumber As Integer
    Dim needHeader As Boolean

    On Error GoTo Failed

    If LenB(runnerWorkbook.Path) = 0 Then
        AppendRunLog = "Runner workbook is unsaved; RunnerLog.csv not written."
        Exit Function
    End If

    logPath = runnerWorkbook.Path & Application.PathSeparator & "RunnerLog.csv"
    needHeader = (LenB(Dir$(logPath)) = 0)

    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
    If needHeader Then
        Print #fileNumber, "Run Start,Run End,Duration (s),Transform,Status,Rows Processed,Rows Rejected,Output Workbook,Issues"
    End If
    Print #fileNumber, _
        CsvField(Format$(startTime, "yyyy-mm-dd hh:nn:ss")) & "," & _
        CsvField(Format$(endTime, "yyyy-mm-dd hh:nn:ss")) & "," & _
        CsvField(Format$((endTime - startTime) * 86400#, "0.0")) & "," & _
        CsvField(transformKey) & "," & _
        CsvField(status) & "," & _
        CsvField(CStr(rowsProcessed)) & "," & _
        CsvField(CStr(rowsRejected)) & "," & _
        CsvField(outputPath) & "," & _
        CsvField(issues)
    Close #fileNumber
    WriteRunStatusFile runnerWorkbook, transformKey & ": " & status & " - " & CStr(rowsProcessed) & _
        " processed, " & CStr(rowsRejected) & " rejected, " & Format$((endTime - startTime) * 86400#, "0.0") & "s"
    AppendRunLog = vbNullString
    Exit Function

Failed:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
    AppendRunLog = "RunnerLog.csv write failed: " & Err.Description
End Function

Private Function CsvField(ByVal value As String) As String
    CsvField = """" & Replace(value, """", """""") & """"
End Function

' Overwrite RunnerStatus.txt with the latest run state, beside the Runner workbook. Overwrite
' (not append) so a refresh always shows the current line. Best-effort: never raises, so a
' locked/held file (e.g. open in Excel) just skips this tick without affecting the run.
Private Sub WriteRunStatusFile(ByVal runnerWorkbook As Workbook, ByVal message As String)
    Dim statusPath As String
    Dim fileNumber As Integer

    On Error GoTo CleanUp
    If LenB(runnerWorkbook.Path) = 0 Then Exit Sub

    statusPath = runnerWorkbook.Path & Application.PathSeparator & "RunnerStatus.txt"
    fileNumber = FreeFile
    Open statusPath For Output As #fileNumber
    Print #fileNumber, message
    Close #fileNumber
    Exit Sub

CleanUp:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
End Sub

' Append one skipped source row to the run's rejects file (<output>.rejects.csv beside the output
' EIB; see BuildRejectLogPath), so finance can see exactly which rows were dropped and why.
' Best-effort (never raises).
' sourceRow is the actual worksheet row number in the source sheet.
' sourceValue is the source cell value the reject turned on (the unmapped code for a failed
' lookup, the offending field for a missing-field reject) -- kept in its own column so the team
' that maintains those codes can pivot the file on it. Empty when no single value is to blame.
' details stays a free-text handle for recognising the row (memo, description, ...).
Public Sub AppendRejectRow(ByVal runnerWorkbook As Workbook, ByVal transformKey As String, ByVal sourceRow As Long, ByVal reason As String, ByVal sourceValue As String, ByVal details As String)
    Dim rejectPath As String
    Dim fileNumber As Integer

    On Error GoTo CleanUp
    If LenB(runnerWorkbook.Path) = 0 Then Exit Sub

    rejectPath = RejectLogPath(runnerWorkbook, transformKey)
    fileNumber = FreeFile
    Open rejectPath For Append As #fileNumber
    Print #fileNumber, CsvField(CStr(sourceRow)) & "," & CsvField(reason) & "," & CsvField(sourceValue) & "," & CsvField(details)
    Close #fileNumber
    Exit Sub

CleanUp:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
End Sub

' Truncate + header the per-transform rejects file at the start of each run (called by BeginRun),
' so it always reflects only the latest run rather than accumulating. Best-effort.
Private Sub ResetRejectLog(ByVal runnerWorkbook As Workbook, ByVal transformKey As String)
    Dim rejectPath As String
    Dim fileNumber As Integer

    On Error GoTo CleanUp
    If LenB(runnerWorkbook.Path) = 0 Then Exit Sub

    rejectPath = RejectLogPath(runnerWorkbook, transformKey)
    fileNumber = FreeFile
    Open rejectPath For Output As #fileNumber
    Print #fileNumber, "Source Row,Reason,Source Value,Details"
    Close #fileNumber
    Exit Sub

CleanUp:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
End Sub

Private Function RejectLogPath(ByVal runnerWorkbook As Workbook, ByVal transformKey As String) As String
    If LenB(mRejectLogPath) > 0 Then
        RejectLogPath = mRejectLogPath
    Else
        ' Fallback: BeginRun could not resolve the output location -- keep the legacy behaviour of
        ' writing beside the Runner workbook so rejects are never silently lost.
        RejectLogPath = runnerWorkbook.Path & Application.PathSeparator & "RunnerRejects_" & transformKey & ".csv"
    End If
End Function

' Resolve where the per-run rejects file lives: beside the EIB the run produces, named to pair
' with it (AP_Travel_output.xlsx -> AP_Travel_output.rejects.csv), so a user finds the generated
' EIB and its dropped rows in the same folder. Mirrors CloneTemplateWorkbook's output resolution
' (Output File Path, else a "_output" sibling of the Target EIB template) so the pairing holds in
' both cases. Returns "" on any control-read failure so RejectLogPath falls back to the Runner
' folder -- best-effort, never aborts a run.
Private Function BuildRejectLogPath(ByVal controlSheet As Worksheet) As String
    Dim outPath As String

    On Error GoTo Fallback
    outPath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    If LenB(outPath) = 0 Then
        outPath = BuildSiblingFilePath(ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB), "_output")
    End If
    If LenB(outPath) = 0 Then Exit Function
    BuildRejectLogPath = JoinPath(ParentFolder(outPath), FileNameWithoutExtension(outPath) & ".rejects.csv")
    Exit Function

Fallback:
    BuildRejectLogPath = vbNullString
End Function

Private Function BuildSiblingFilePath(ByVal filePath As String, ByVal suffix As String) As String
    Dim parentFolderPath As String
    Dim fileName As String
    Dim extensionName As String

    parentFolderPath = ParentFolder(filePath)
    fileName = FileNameWithoutExtension(filePath)
    extensionName = FileExtension(filePath)

    If LenB(extensionName) = 0 Then
        BuildSiblingFilePath = JoinPath(parentFolderPath, fileName & suffix)
    Else
        BuildSiblingFilePath = JoinPath(parentFolderPath, fileName & suffix & "." & extensionName)
    End If
End Function

Private Function BuildSingleCellTable(ByVal cellValue As Variant) As Variant
    Dim result(1 To 1, 1 To 1) As Variant

    result(1, 1) = cellValue
    BuildSingleCellTable = result
End Function

Private Function CoerceRangeValueToTable(ByVal rangeValue As Variant, ByVal rowCount As Long, ByVal columnCount As Long) As Variant
    Dim result() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    If IsArray(rangeValue) Then
        CoerceRangeValueToTable = rangeValue
        Exit Function
    End If

    ReDim result(1 To rowCount, 1 To columnCount)
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            result(rowIndex, columnIndex) = rangeValue
        Next columnIndex
    Next rowIndex

    CoerceRangeValueToTable = result
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (LenB(Dir$(filePath)) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FileExtension(ByVal filePath As String) As String
    Dim dotPosition As Long
    Dim fileName As String

    fileName = BaseName(filePath)
    dotPosition = InStrRev(fileName, ".")
    If dotPosition = 0 Then
        FileExtension = vbNullString
    Else
        FileExtension = Mid$(fileName, dotPosition + 1)
    End If
End Function

Private Function FileNameWithoutExtension(ByVal filePath As String) As String
    Dim fileName As String
    Dim dotPosition As Long

    fileName = BaseName(filePath)
    dotPosition = InStrRev(fileName, ".")
    If dotPosition = 0 Then
        FileNameWithoutExtension = fileName
    Else
        FileNameWithoutExtension = Left$(fileName, dotPosition - 1)
    End If
End Function

Private Function BaseName(ByVal filePath As String) As String
    Dim slashPosition As Long
    Dim colonPosition As Long
    Dim splitPosition As Long

    slashPosition = InStrRev(filePath, "/")
    colonPosition = InStrRev(filePath, ":")
    splitPosition = slashPosition
    If colonPosition > splitPosition Then
        splitPosition = colonPosition
    End If

    If splitPosition = 0 Then
        BaseName = filePath
    Else
        BaseName = Mid$(filePath, splitPosition + 1)
    End If
End Function

Private Function FindLabelCell(ByVal targetSheet As Worksheet, ByVal labelText As String) As Range
    Dim searchRange As Range
    Dim currentCell As Range

    Set searchRange = targetSheet.UsedRange.Columns(1)
    For Each currentCell In searchRange.Cells
        If Trim$(CStr(currentCell.Value)) = labelText Then
            Set FindLabelCell = currentCell
            Exit Function
        End If
    Next currentCell
End Function

Private Function FindOrCreateLabelCell(ByVal targetSheet As Worksheet, ByVal labelText As String) As Range
    Dim labelCell As Range
    Dim nextRow As Long

    Set labelCell = FindLabelCell(targetSheet, labelText)
    If Not labelCell Is Nothing Then
        Set FindOrCreateLabelCell = labelCell
        Exit Function
    End If

    nextRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    If LenB(Trim$(CStr(targetSheet.Cells(nextRow, 1).Value))) > 0 Then
        nextRow = nextRow + 1
    End If

    Set labelCell = targetSheet.Cells(nextRow, 1)
    labelCell.Value = labelText
    Set FindOrCreateLabelCell = labelCell
End Function

Private Function JoinPath(ByVal parentFolder As String, ByVal childName As String) As String
    Dim separatorText As String

    separatorText = DetectPathSeparator(parentFolder)
    If Right$(parentFolder, 1) = separatorText Then
        JoinPath = parentFolder & childName
    Else
        JoinPath = parentFolder & separatorText & childName
    End If
End Function

Private Function DetectPathSeparator(ByVal filePath As String) As String
    If InStr(filePath, "/") > 0 Then
        DetectPathSeparator = "/"
    ElseIf InStr(filePath, ":") > 0 Then
        DetectPathSeparator = ":"
    Else
        DetectPathSeparator = Application.PathSeparator
    End If
End Function

Private Function ParentFolder(ByVal filePath As String) As String
    Dim slashPosition As Long
    Dim colonPosition As Long
    Dim splitPosition As Long

    slashPosition = InStrRev(filePath, "/")
    colonPosition = InStrRev(filePath, ":")
    splitPosition = slashPosition
    If colonPosition > splitPosition Then
        splitPosition = colonPosition
    End If

    If splitPosition = 0 Then
        ParentFolder = ResolveBaseFolder()
    Else
        ParentFolder = Left$(filePath, splitPosition - 1)
    End If
End Function

Private Function ResolveBaseFolder() As String
    If LenB(ThisWorkbook.Path) > 0 Then
        ResolveBaseFolder = ThisWorkbook.Path
    Else
        ResolveBaseFolder = Environ$("HOME")
    End If
End Function

Private Function Slugify(ByVal value As String) As String
    Dim i As Long
    Dim currentCharacter As String
    Dim result As String

    value = LCase$(Trim$(value))
    For i = 1 To Len(value)
        currentCharacter = Mid$(value, i, 1)
        If currentCharacter Like "[a-z0-9]" Then
            result = result & currentCharacter
        ElseIf LenB(result) = 0 Or Right$(result, 1) <> "_" Then
            result = result & "_"
        End If
    Next i

    Do While Right$(result, 1) = "_"
        result = Left$(result, Len(result) - 1)
        If LenB(result) = 0 Then
            Exit Do
        End If
    Loop

    Slugify = result
End Function
