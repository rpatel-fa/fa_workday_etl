Attribute VB_Name = "modArDisciplineCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const OUTPUT_ADD_ONLY As String = "Y"
' The fixed prefix of the "d" (Customer Invoice Reference ID) row's output_format in
' config/transform_specs/ar_discipline_invoice.yaml.
Private Const REFERENCE_ID_PREFIX As String = "cfa_discipline_"
' The target_format of the fl row in config/transform_specs/ar_discipline_invoice.yaml.
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"
' The "j" (Company) row in config/transform_specs/ar_discipline_invoice.yaml's mapping --
' the fallback used only if the source sheet's first-data-row Company cell (col M) is blank.
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"

'  --- Profile-local Tax Triple overrides (hybrid scope; shadows modFinanceCommon.bas shared defaults) ---
'  These values come from ar_discipline_invoice.yaml transform_constants.tax_triple_per_row and
'  mirror ar_integration_invoice (modArIntegrationCustInvoice.bas); keep them profile-local so a
'  change to the shared modFinanceCommon defaults does not alter this transform. Set 2026-10-06.
Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Sale_of_Goods_and_Services"
Private Const TAX_APPLICABILITY_ZERO As String = "GBR_Outside_of_Scope"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_CODE_ZERO As String = "GBR_Zero_Rated"

' Source column A holds the invoice date (dd.mm.yy text); source column F is ALSO headered
' "DATE" (a Date-typed invoice-month cell, unused), so a header-text lookup would collide.
' Read A positionally via this named constant instead of a header search. See mapping.br.
Private Const SRC_COL_INVOICE_DATE As String = "A"

' ============================================================================
' Target EIB column letters -- Submit_Customer_Invoice_v46.1.xlsx (Submit Customer Invoice).
' 1:1 mirror of config/transform_specs/ar_discipline_invoice.yaml mapping: section.
' DB (Memo) is intentionally NOT declared -- left blank, no source column (see mapping.db).
' I (Locked in Workday) is also left blank by requester instruction (see mapping.i).
' ============================================================================
Private Const TGT_SPREADSHEET_KEY As String = "B"
Private Const TGT_ADD_ONLY As String = "C"
Private Const TGT_CUSTOMER_INVOICE_REFERENCE_ID As String = "D"
Private Const TGT_CUSTOMER_INVOICE_ID As String = "G"
Private Const TGT_SUBMIT As String = "H"
Private Const TGT_COMPANY As String = "J"
Private Const TGT_CURRENCY As String = "AI"
Private Const TGT_CUSTOMER As String = "AJ"
Private Const TGT_INVOICE_NUMBER As String = "BO"
Private Const TGT_INVOICE_DATE As String = "BR"
Private Const TGT_PAYMENT_STATUS As String = "CK"
Private Const TGT_DOCUMENT_STATUS As String = "CL"
Private Const TGT_AMOUNT_DUE As String = "CM"
Private Const TGT_CUSTOMER_INVOICE_TYPE As String = "CO"
Private Const TGT_ROW_ID_LINE As String = "DS"
Private Const TGT_LINE_ORDER As String = "DV"
Private Const TGT_LINE_TAX_APPLICABILITY As String = "FA"
Private Const TGT_LINE_TAX_CODE As String = "FB"
Private Const TGT_LINE_TRANSACTION_DATE As String = "FP"
Private Const TGT_LINE_QUANTITY As String = "FH"
Private Const TGT_LINE_MEMO As String = "FV"
Private Const TGT_REVENUE_CATEGORY As String = "EM"
Private Const TGT_LINE_ITEM_DESCRIPTION As String = "EN"
Private Const TGT_LINE_UNIT_COST As String = "FL"
Private Const TGT_LINE_EXTENDED_AMOUNT As String = "FO"
Private Const TGT_COST_CENTRE As String = "FX"
Private Const TGT_REPORTING_TYPE As String = "FY"
Private Const TARGET_LAST_COLUMN_LETTERS As String = TGT_REPORTING_TYPE

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Direct transform: load the Discipline (fines/payments) invoice extract (one row per
' customer per month, "Rolling" sheet) and write straight into the cloned target
' EIB. Company, Cost Centre, Reporting Type, and Revenue Category are read ONCE from the
' sheet's first data row (source columns M/N/O/P, populated on that row only) and
' broadcast to every output row. The Customer ID is read directly from source column B
' ("WORKDAY ID") -- no reference-extract lookup. EIB Customer Invoice Reference ID (D) is
' the generated cfa_discipline_{WORKDAY ID without hyphens}_{yyyyMMdd}_{Spreadsheet Key}
' value; Invoice Number (BO) and Memo (DB) are left blank; Line Tax Applicability (FA) and
' Line Tax Code (FB) are set per row from the source NET/VAT split (same rule as
' ar_integration_invoice). Line grouping (Spreadsheet
' Key / Line Order) is keyed directly by source `WORKDAY ID`, which is already unique per row
' on this sheet. The field mapping mirrors the mapping: section in
' config/transform_specs/ar_discipline_invoice.yaml, the same pattern AR Invoice uses.
Public Sub RunArDisciplineInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceTable As Variant
    Dim workdayIdColumn As Long, invoiceDateColumn As Long, netColumn As Long, descriptionColumn As Long, qtyColumn As Long
    Dim vatColumn As Long, grossColumn As Long
    Dim companyColumn As Long, costCentreColumn As Long, reportingTypeColumn As Long, revenueCategoryColumn As Long
    Dim companyBroadcast As String, costCentreBroadcast As String
    Dim reportingTypeBroadcast As String, revenueCategoryBroadcast As String
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim issues As Collection
    Dim outputPath As String
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
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

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    sourceTable = ReadWorksheetTable(ResolveSourceWorksheet(sourceWorkbook, sheetToResolve))
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    workdayIdColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "WORKDAY ID")
    netColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "NET")
    vatColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "VAT")
    ' GROSS is optional, as in ar_integration_invoice: absent -> Net + VAT. See mapping.fa.
    grossColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "GROSS")
    descriptionColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "DESCRIPTION")
    qtyColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "QTY")
    companyColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Company")
    costCentreColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Cost Centre")
    reportingTypeColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Reporting Type")
    revenueCategoryColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Revenue Category")
    ' LANDMINE: the source sheet has TWO columns headered "Date" (col A, the invoice date
    ' as dd.mm.yy text) and "DATE" (col F, an unused invoice-month cell) -- NormalizeToken
    ' lower-cases both to "date", so a header-text lookup is ambiguous. Read col A
    ' positionally instead. See mapping.br in the transform spec.
    invoiceDateColumn = ColumnIndex(SRC_COL_INVOICE_DATE)

    ' Company / Cost Centre / Reporting Type / Revenue Category are populated on the
    ' sheet's first data row ONLY (every later row is blank) -- read once here and
    ' broadcast the same value to every output row (mapping.j / .fx / .fy / .em).
    companyBroadcast = Trim$(CStr(sourceTable(sourceDataStartRow, companyColumn)))
    If LenB(companyBroadcast) = 0 Then companyBroadcast = COMPANY_DEFAULT_FALLBACK
    costCentreBroadcast = Trim$(CStr(sourceTable(sourceDataStartRow, costCentreColumn)))
    reportingTypeBroadcast = Trim$(CStr(sourceTable(sourceDataStartRow, reportingTypeColumn)))
    revenueCategoryBroadcast = Trim$(CStr(sourceTable(sourceDataStartRow, revenueCategoryColumn)))

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, TARGET_LAST_COLUMN_LETTERS
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, workdayIdColumn, invoiceDateColumn, netColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (WORKDAY ID / Invoice Date / NET) or trailer/totals row", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
        Else
            rowsProcessed = rowsProcessed + 1
            WriteOutputRow targetSheet, targetRow, sourceTable, rowIndex, _
                workdayIdColumn, invoiceDateColumn, netColumn, descriptionColumn, qtyColumn, _
                vatColumn, grossColumn, _
                companyBroadcast, costCentreBroadcast, reportingTypeBroadcast, revenueCategoryBroadcast
            targetRow = targetRow + 1
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing WORKDAY ID, Invoice Date, or NET (includes the sheet's trailer/totals row)."
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
    Err.Raise failureNumber, "modArDisciplineCustInvoice.RunArDisciplineInvoiceCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal workdayIdColumn As Long, ByVal invoiceDateColumn As Long, ByVal netColumn As Long, ByVal descriptionColumn As Long, ByVal qtyColumn As Long, _
    ByVal vatColumn As Long, ByVal grossColumn As Long, _
    ByVal companyBroadcast As String, ByVal costCentreBroadcast As String, _
    ByVal reportingTypeBroadcast As String, ByVal revenueCategoryBroadcast As String _
)
    Dim workdayId As String
    Dim invoiceDate As String
    Dim referenceId As String
    Dim netAmount As Double
    Dim vatAmount As Double
    Dim grossAmount As Double
    Dim vatRatio As Double
    Dim taxApplicability As String
    Dim taxCode As String
    Dim qtyValue As Double
    Dim descriptionValue As String
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long

    workdayId = Trim$(CStr(sourceTable(sourceRow, workdayIdColumn)))
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    netAmount = ResolveAmount(sourceTable(sourceRow, netColumn))
    vatAmount = ResolveAmount(sourceTable(sourceRow, vatColumn))
    If grossColumn > 0 Then
        grossAmount = ResolveAmount(sourceTable(sourceRow, grossColumn))
    Else
        grossAmount = netAmount + vatAmount
    End If
    qtyValue = ResolveAmount(sourceTable(sourceRow, qtyColumn))
    descriptionValue = Trim$(CStr(sourceTable(sourceRow, descriptionColumn)))
    ' Grouping keyed directly by WORKDAY ID, which is already unique per row on this sheet.
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, workdayId, mGrpNextKey, lineOrderValue)

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = OUTPUT_ADD_ONLY
    ' Customer Invoice ID (G) = cfa_discipline_{WORKDAY ID without hyphens}_{yyyyMMdd of BR}_{Spreadsheet Key B}
    ' (e.g. C-000190 + 2026-09-18 + key 1 -> cfa_discipline_C000190_20260918_1).
    ' Customer Invoice Reference ID (D) is intentionally left blank. See mapping.g / mapping.d / mapping.bo.
    referenceId = REFERENCE_ID_PREFIX & Replace(workdayId, "-", vbNullString) _
        & "_" & Replace(invoiceDate, "-", vbNullString) _
        & "_" & CStr(spreadsheetKeyValue)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).Value = referenceId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REFERENCE_ID)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_NUMBER)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    ' I (Locked in Workday) intentionally left blank -- see mapping.i.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = companyBroadcast
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = workdayId
    ' Force Text format before assigning: the target template's BR cell is formatted
    ' General, so assigning a "yyyy-mm-dd" string straight to .Value lets Excel
    ' auto-detect it as a date and silently re-render it in the system's default
    ' short-date format -- overriding the ar_discipline_invoice spec mapping's
    ' target_format = YYYY-MM-DD.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = invoiceDate
    ' Transaction Date (FP) is also source column A, the same value as Invoice Date (BR).
    ' See mapping.fp.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    ' Amount Due (CM) intentionally left blank on every row (2026-10-06), same as
    ' ar_integration_invoice. See mapping.cm.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    ' CY (Include Customer Worktags) intentionally left blank -- see mapping.cy.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(lineOrderValue)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = CStr(lineOrderValue)
    ' Line Tax Applicability (FA) / Line Tax Code (FB): per-row tax triple from the source
    ' NET/VAT ratio, mirroring ar_integration_invoice (2026-10-06). See mapping.fa / mapping.fb.
    If netAmount > 0# Then
        vatRatio = vatAmount / netAmount
    Else
        vatRatio = 0#
    End If

    If (netAmount > 0# And vatAmount = 0#) Or (netAmount = 0# And grossAmount = 0#) Then
        taxApplicability = TAX_APPLICABILITY_ZERO
        taxCode = TAX_CODE_ZERO
    ElseIf Abs(vatRatio - 0.2) < 0.01 Then
        taxApplicability = TAX_APPLICABILITY_STANDARD
        taxCode = TAX_CODE_STANDARD
    Else
        taxApplicability = TAX_APPLICABILITY_STANDARD
        taxCode = TAX_CODE_STANDARD
    End If
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = taxApplicability
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = taxCode
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = qtyValue
    If LenB(revenueCategoryBroadcast) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = revenueCategoryBroadcast
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = descriptionValue
    ' Line Memo (FV) mirrors Line Item Description (EN) -- same convention as ar_gbe_invoice.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_MEMO)).Value = descriptionValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = netAmount
    ' Extended Amount (FO) intentionally left blank on every row (2026-10-06), same as
    ' ar_integration_invoice. See mapping.fo.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).Value = vbNullString
    If LenB(costCentreBroadcast) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_COST_CENTRE)).Value = costCentreBroadcast
    If LenB(reportingTypeBroadcast) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_REPORTING_TYPE)).Value = reportingTypeBroadcast
    ' DB (Memo) intentionally left blank -- no source narrative column mapped. See mapping.db.
End Sub

Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal workdayIdColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal netColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, workdayIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, netColumn)) = 0#
End Function
