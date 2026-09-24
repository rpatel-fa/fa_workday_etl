Attribute VB_Name = "modApCfaDistribAmendment"
Option Explicit

' Source of truth: config/transform_specs/ap_cfa_distributions_amendments.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

' NEW 2026-09-24. CFA Distributions CREDIT NOTES -> Submit Supplier Invoice Adjustment EIB
' (Submit_Supplier_Invoice_Adjustment_v46.0). Reads the SAME "Distributions - <Month>.xlsm" /
' "Finance Master analysis" tab as the invoices transform (modApCfaDistribSupplierInvoice) but
' processes ONLY rows where col U "Invoice type" <> "Invoice" (credit notes; negative Original
' Amount at the bottom of col V). Column mapping mirrors the GP-ingest precedent
' workday_supplier_credit_note.sql. Amendment lines carry NO worktag coding (Spend Category DH
' left blank) and NO tax; there is NO Mapping MASTER lookup. Each credit note is its own
' adjustment document (per-row Spreadsheet Key sequence; NOT invoice-line grouped, because the
' source Invoice Number label repeats across the distinct per-supplier credit notes).

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice Adjustm"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const CLEAR_END_COL As String = "FB"

Private Const OUTPUT_ADD_ONLY As String = "Y"
Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"
Private Const OUTPUT_LINE_ROW_ID As String = "1"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const CONSTANT_COMPANY_CODE As String = "FA"
Private Const INVOICE_TYPE_INVOICE As String = "INVOICE"

Private mCompanyId As String

Public Sub RunApCfaDistribAmendment(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As Variant
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unresolvedSupplierRows As Long
    Dim invoiceRows As Long
    Dim rowIndex As Long
    Dim docNoColumn As Long                  ' A  DocNo / row key (headerless; = CFA_col - 1); trace
    Dim cfaNameColumn As Long                ' B  County FA name (anchor)
    Dim creditorIdColumn As Long             ' D  Creditor ID (trace)
    Dim workdayIdColumn As Long              ' T  Workday ID (supplier passthrough -> EIB S)
    Dim invoiceTypeColumn As Long            ' U  Invoice type (filter: keep <> "Invoice")
    Dim originalAmountColumn As Long         ' V  Original Amount (-> EIB AP + EQ; negative)
    Dim invoiceNumberColumn As Long          ' Z  Invoice Number (-> EIB K)
    Dim invoiceDescriptionColumn As Long     ' AA Invoice Description (-> EIB DA)
    Dim invoiceDateColumn As Long            ' AD Invoice Date (-> EIB Y adjustment date)
    Dim companyId As String
    Dim supplierId As String
    Dim issueParts As Collection
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim outputPath As String
    Dim headerRowIndex As Long
    Dim dataStartRowIndex As Long
    Dim requiredHeaders As Variant
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
    Dim worksheetNameStr As String
    Dim worksheetResolutionStr As String
    Dim sheetToResolve As String
    Dim sourceHeaderRow As Long
    Dim sourceDataStartRow As Long

    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)

    Set issueParts = New Collection

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If StrComp(worksheetResolutionStr, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetNameStr
    End If

    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH

    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApCfaDistribAmendment", 8)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    Set sourceSheet = ResolveSourceWorksheet(sourceWorkbook, sheetToResolve)
    sourceTable = ReadWorksheetTable(sourceSheet)
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    ' Schema validation phase (Tier 2 VBA defence-in-depth).
    Dim schemaMsg As String
    Dim schemaHeaderRow As Long
    Dim schemaDiag As String
    schemaMsg = ValidateSourceWorkbookSchema( _
        sheetToResolve, _
        sourceTable, _
        sourceHeaderRow, _
        sourceDataStartRow, _
        requiredHeaders, _
        , _
        schemaHeaderRow, _
        schemaDiag _
    )
    If LenB(schemaMsg) > 0 Then
        Err.Raise vbObjectError + 2110, "modApCfaDistribAmendment.RunApCfaDistribAmendment", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    cfaNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "CFA")
    docNoColumn = cfaNameColumn - 1
    creditorIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Creditor ID")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday ID")
    invoiceTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice type")
    originalAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Original Amount")
    invoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Number")
    invoiceDescriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Description|Invoice Desctiption")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Date")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, CLEAR_END_COL
    targetRow = TARGET_DATA_START_ROW

    companyId = ResolveMappedCompany(runnerWorkbook, CONSTANT_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, _
            Array(docNoColumn, creditorIdColumn, invoiceTypeColumn, invoiceNumberColumn, workdayIdColumn), _
            Array(originalAmountColumn), _
            Array(docNoColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(docNoColumn, creditorIdColumn, invoiceTypeColumn, invoiceNumberColumn, workdayIdColumn), _
                Array(originalAmountColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 8 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsInvoiceType(sourceTable, rowIndex, invoiceTypeColumn) Then
            '  -- Invoice rows: handled by ap_cfa_distributions. Skip silently here.
            invoiceRows = invoiceRows + 1
            consecutiveEmptyRows = 0
        ElseIf RowIsSkippable(sourceTable, rowIndex, workdayIdColumn, originalAmountColumn, _
            invoiceNumberColumn, invoiceDescriptionColumn, invoiceDateColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required credit-note field (Workday-ID/T / Invoice-Number/Z / " _
                    & "Invoice-Description/AA / Invoice-Date/AD blank, or Original-Amount/V = 0)", _
                Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
        Else
            supplierId = Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
            If LenB(supplierId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unresolvedSupplierRows = unresolvedSupplierRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Supplier unresolved: Workday ID (col T) was BLANK on a credit-note row. " _
                        & "Creditor ID / col D (for traceability): " _
                        & Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
            Else
                If WriteAmendmentRow(targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                    supplierId, originalAmountColumn, invoiceNumberColumn, _
                    invoiceDescriptionColumn, invoiceDateColumn) Then
                    rowsProcessed = rowsProcessed + 1
                    targetRow = targetRow + 1
                Else
                    rowsSkipped = rowsSkipped + 1
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) _
            & " credit-note row(s) with a missing required field (Workday ID/T, Invoice Number/Z, " _
            & "Invoice Description/AA, Invoice Date/AD, or Original Amount/V = 0)."
    End If
    If invoiceRows > 0 Then
        issueParts.Add CStr(invoiceRows) _
            & " row(s) with Invoice type = 'Invoice' were skipped here -- they are handled by " _
            & "ap_cfa_distributions (Submit Supplier Invoice EIB)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " credit-note row(s) whose col T ""Workday ID"" was BLANK."
    End If

    outputPath = outputWorkbook.FullName
    outputWorkbook.Save
    outputWorkbook.Close SaveChanges:=False
    Set outputWorkbook = Nothing

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Completed", rowsProcessed, rowsSkipped, runStartTime, runEndTime, outputPath, JoinCollectionValues(issueParts, " | "))
    If LenB(logIssue) > 0 Then issueParts.Add logIssue
    EndRun

    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Completed"
    WriteControlValue controlSheet, RESULT_LABEL_ROWS_PROCESSED, CStr(rowsProcessed)
    WriteControlValue controlSheet, RESULT_LABEL_ROWS_SKIPPED, CStr(rowsSkipped)
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, JoinCollectionValues(issueParts, " | ")
    WriteControlValue controlSheet, RESULT_LABEL_OUTPUT_WORKBOOK, outputPath
    WriteControlValue controlSheet, RESULT_LABEL_SELECTION_KEY, selectionKey
    WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    runnerWorkbook.Save
    Exit Sub

HandleFailure:
    failureNumber = Err.Number
    failureMessage = Err.Description
    On Error Resume Next
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    If Not outputWorkbook Is Nothing Then outputWorkbook.Close SaveChanges:=False
    On Error GoTo 0

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Failed", rowsProcessed, rowsSkipped, runStartTime, runEndTime, "", failureMessage)
    If LenB(logIssue) > 0 Then failureMessage = failureMessage & " | " & logIssue
    EndRun

    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Failed"
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, failureMessage
    WriteControlValue controlSheet, RESULT_LABEL_SELECTION_KEY, selectionKey
    WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error Resume Next
    runnerWorkbook.Save
    On Error GoTo 0
    Err.Raise failureNumber, "modApCfaDistribAmendment.RunApCfaDistribAmendment", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one Supplier Invoice Adjustment (credit-note) row.
'     B  Spreadsheet Key (per-row sequence) ; C Add Only (Y) ; E Auto Complete (Y) ; I Submit (Y)
'     K  Invoice Number (col Z)
'     P  Company (constant FA -> THE FA GROUP - FA)
'     R  Currency (GBP)
'     S  Supplier (col T Workday ID passthrough)
'     X  Adjustment Reason (BLANK)
'     Y  Adjustment Date (col AD via NormalizeSourceDate)
'     AP Control Total Amount (col V; as-is, negative)
'     CV Line Row ID (1)
'     DA Item Description (col AA)
'     DH Spend Category (BLANK)
'     EQ Extended Amount (col V; as-is, negative)
'     FB Memo (BLANK)
'  --------------------------------------------------------------
Private Function WriteAmendmentRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal supplierId As String, _
    ByVal originalAmountColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal invoiceDescriptionColumn As Long, _
    ByVal invoiceDateColumn As Long _
) As Boolean
    Dim amount As Double
    Dim invoiceNumberValue As String
    Dim descriptionValue As String
    Dim adjustmentDateValue As String

    invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = Trim$(CStr(sourceTable(sourceRow, invoiceDescriptionColumn)))
    amount = ResolveAmount(sourceTable(sourceRow, originalAmountColumn))
    adjustmentDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = sequenceNumber
    targetSheet.Cells(targetRow, ColumnIndex("C")).Value = OUTPUT_ADD_ONLY
    targetSheet.Cells(targetRow, ColumnIndex("E")).Value = OUTPUT_AUTO_COMPLETE
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("P")).Value = mCompanyId
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = OUTPUT_DEFAULT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex("S")).Value = supplierId
    targetSheet.Cells(targetRow, ColumnIndex("X")).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex("Y")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("Y")).Value = adjustmentDateValue
    targetSheet.Cells(targetRow, ColumnIndex("AP")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("CV")).Value = OUTPUT_LINE_ROW_ID
    targetSheet.Cells(targetRow, ColumnIndex("DA")).Value = descriptionValue
    targetSheet.Cells(targetRow, ColumnIndex("DH")).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex("EQ")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("FB")).Value = vbNullString
    WriteAmendmentRow = True
End Function

'  --------------------------------------------------------------
'  RowIsInvoiceType -- TRUE when col U "Invoice type" = "Invoice" (routed to the invoices
'  transform). Blank Invoice type is NOT treated as invoice here (it falls through to the
'  missing-field gate so genuine gaps are still reported).
'  --------------------------------------------------------------
Private Function RowIsInvoiceType( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal invoiceTypeColumn As Long _
) As Boolean
    Dim v As String
    v = Trim$(CStr(tableData(rowIndex, invoiceTypeColumn)))
    If LenB(v) = 0 Then
        RowIsInvoiceType = False
    Else
        RowIsInvoiceType = (StrComp(UCase$(v), INVOICE_TYPE_INVOICE, vbBinaryCompare) = 0)
    End If
End Function

'  --------------------------------------------------------------
'  RowIsSkippable -- per-row missing-field gate (matches skip_rule in the spec YAML).
'  Workday ID/T, Invoice Number/Z, Invoice Description/AA, Invoice Date/AD blank,
'  OR Original Amount/V = 0.
'  --------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal workdayIdColumn As Long, _
    ByVal originalAmountColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal invoiceDescriptionColumn As Long, _
    ByVal invoiceDateColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, workdayIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDescriptionColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, originalAmountColumn)) = 0#
End Function
