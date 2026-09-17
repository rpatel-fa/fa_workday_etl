Attribute VB_Name = "modArManualInvoiceCustInvoice"
Option Explicit

' Source of truth: config/transform_specs/ar_manual_invoice.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner enforces spec<->code column agreement

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const OUTPUT_ADD_ONLY As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_COMPANY As String = "THE FA GROUP - WNSL"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
' Was "Standard". Changed to "EBW" per requester instruction (2026-09-17) -- see
' config/transform_specs/ar_manual_invoice.yaml eib_house_defaults.invoice_type.
Private Const OUTPUT_INVOICE_TYPE As String = "EBW"
Private Const OUTPUT_LINE_QUANTITY As Double = 1#
Private Const OUTPUT_LINE_ITEM_DESCRIPTION As String = "EBW Membership Season 26/27"
' Line Tax Applicability (FA) / Line Tax Code (FB) intentionally left blank on every row
' per requester instruction (2026-09-15) -- see config/transform_specs/ar_manual_invoice.yaml
' fields.line_tax_applicability / fields.line_tax_code.
Private Const OUTPUT_COST_CENTRE As String = "CC04001"

'  ============================================================================
'  Target EIB column letters -- Submit_Customer_Invoice_v46.1.xlsx (Submit Customer Invoice)
'  1:1 mirror of ar_manual_invoice.yaml transform_constants.customer_invoice_eib_columns.
'  ============================================================================
Private Const TGT_SPREADSHEET_KEY As String = "B"
Private Const TGT_ADD_ONLY As String = "C"
Private Const TGT_CUSTOMER_INVOICE_REF_ID As String = "D"
Private Const TGT_SUBMIT As String = "H"
Private Const TGT_LOCKED_IN_WORKDAY As String = "I"
Private Const TGT_COMPANY As String = "J"
Private Const TGT_CURRENCY As String = "AI"
Private Const TGT_CUSTOMER As String = "AJ"
Private Const TGT_INVOICE_DATE As String = "BR"
Private Const TGT_PAYMENT_STATUS As String = "CK"
Private Const TGT_DOCUMENT_STATUS As String = "CL"
Private Const TGT_AMOUNT_DUE As String = "CM"
Private Const TGT_CUSTOMER_INVOICE_TYPE As String = "CO"
Private Const TGT_HEADER_MEMO As String = "DB"
Private Const TGT_ROW_ID_LINE As String = "DS"
Private Const TGT_LINE_REFERENCE As String = "DU"
Private Const TGT_LINE_ORDER As String = "DV"
Private Const TGT_REVENUE_CATEGORY As String = "EM"
Private Const TGT_LINE_ITEM_DESCRIPTION As String = "EN"
Private Const TGT_LINE_TAX_APPLICABILITY As String = "FA"
Private Const TGT_LINE_TAX_CODE As String = "FB"
Private Const TGT_LINE_QUANTITY As String = "FH"
Private Const TGT_LINE_UNIT_COST As String = "FL"
Private Const TGT_LINE_EXTENDED_AMOUNT As String = "FO"
Private Const TGT_COST_CENTRE As String = "FX"

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

Public Sub RunArManualInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceTable As Variant
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim createdOnColumn As Long, licenceRefColumn As Long, fanColumn As Long
    Dim invoiceAmountColumn As Long, workdayIdColumn As Long
    Dim issues As Collection
    Dim outputPath As String
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
    Dim requiredHeaders As Variant
    Dim resolvedSourceSheet As Worksheet
    Dim headerRowIndex As Long
    Dim dataStartRowIndex As Long
    Dim schemaMsg As String
    Dim schemaHeaderRow As Long
    Dim schemaDiag As String
    Dim worksheetNameStr As String
    Dim worksheetResolutionStr As String
    Dim sheetToResolve As String
    Dim sourceHeaderRow As Long
    Dim sourceDataStartRow As Long

    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)

    Set issues = New Collection

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If StrComp(worksheetResolutionStr, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetNameStr
    End If

    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modArManualInvoiceCustInvoice", 6)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    Set resolvedSourceSheet = ResolveSourceWorksheet(sourceWorkbook, sheetToResolve)
    sourceTable = ReadWorksheetTable(resolvedSourceSheet)
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing
    Set resolvedSourceSheet = Nothing

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
        Err.Raise vbObjectError + 1911, "modArManualInvoiceCustInvoice.RunArManualInvoiceCustomerInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issues.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    createdOnColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Created On")
    licenceRefColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Licence Ref No")
    fanColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "FAN#")
    invoiceAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Amount (Net)")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday ID")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "MV"
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, createdOnColumn, licenceRefColumn, invoiceAmountColumn, workdayIdColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Created On / Licence Ref No / Invoice Amount (Net) / Workday ID)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, licenceRefColumn)))
        Else
            rowsProcessed = rowsProcessed + 1
            WriteOutputRow runnerWorkbook, targetSheet, targetRow, sourceTable, rowIndex, _
                createdOnColumn, licenceRefColumn, fanColumn, invoiceAmountColumn, workdayIdColumn
            targetRow = targetRow + 1
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing required fields."
    End If

    outputPath = outputWorkbook.FullName
    outputWorkbook.Save
    outputWorkbook.Close SaveChanges:=False
    Set outputWorkbook = Nothing

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Completed", rowsProcessed, rowsSkipped, runStartTime, runEndTime, outputPath, JoinCollectionValues(issues, " | "))
    If LenB(logIssue) > 0 Then issues.Add logIssue
    EndRun

    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Completed"
    WriteControlValue controlSheet, RESULT_LABEL_ROWS_PROCESSED, CStr(rowsProcessed)
    WriteControlValue controlSheet, RESULT_LABEL_ROWS_SKIPPED, CStr(rowsSkipped)
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, JoinCollectionValues(issues, " | ")
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
    Err.Raise failureNumber, "modArManualInvoiceCustInvoice.RunArManualInvoiceCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputRow( _
    ByVal runnerWorkbook As Workbook, _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal createdOnColumn As Long, ByVal licenceRefColumn As Long, ByVal fanColumn As Long, _
    ByVal invoiceAmountColumn As Long, ByVal workdayIdColumn As Long _
)
    Dim invoiceDate As String
    Dim invoiceAmount As Double
    Dim licenceRefValue As String
    Dim fanValue As String
    Dim headerMemoValue As String
    Dim workdayCustomerId As String
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long

    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, createdOnColumn))
    invoiceAmount = ResolveAmount(sourceTable(sourceRow, invoiceAmountColumn))
    licenceRefValue = Trim$(CStr(sourceTable(sourceRow, licenceRefColumn)))
    fanValue = Trim$(CStr(sourceTable(sourceRow, fanColumn)))
    headerMemoValue = licenceRefValue & "_" & fanValue
    workdayCustomerId = Trim$(CStr(sourceTable(sourceRow, workdayIdColumn)))

    ' No source invoice-number column exists: every row is its own single-line invoice
    ' (AssignInvoiceLineGroup's documented defensive behaviour for an empty-string key).
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, vbNullString, mGrpNextKey, lineOrderValue)

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = OUTPUT_ADD_ONLY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REF_ID)).Value = licenceRefValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LOCKED_IN_WORKDAY)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = OUTPUT_COMPANY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = workdayCustomerId

    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = invoiceDate

    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = invoiceAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_MEMO)).Value = headerMemoValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(lineOrderValue)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_REFERENCE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = CStr(lineOrderValue)

    ' Revenue Category (EM) intentionally left blank -- no accounting-string driver
    ' in this source (see config/transform_specs/ar_manual_invoice.yaml accounting_string).
    targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = ""

    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = OUTPUT_LINE_ITEM_DESCRIPTION
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = OUTPUT_LINE_QUANTITY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = invoiceAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).Value = invoiceAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COST_CENTRE)).Value = OUTPUT_COST_CENTRE
End Sub

Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal createdOnColumn As Long, _
    ByVal licenceRefColumn As Long, _
    ByVal invoiceAmountColumn As Long, _
    ByVal workdayIdColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, createdOnColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, licenceRefColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, invoiceAmountColumn)) = 0# Or _
        LenB(Trim$(CStr(tableData(rowIndex, workdayIdColumn)))) = 0
End Function
