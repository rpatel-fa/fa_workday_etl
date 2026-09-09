Attribute VB_Name = "modArMultiLineCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).
'  Historical hard-coded constants retained for documentation audit only; NOT used at runtime:
'  Private Const SOURCE_WORKSHEET_NAME As String = vbNullString
'  Private Const SOURCE_HEADER_ROW As Long = 1
'  Private Const SOURCE_DATA_START_ROW As Long = 2

Private Const SOURCE_COMPANY_CODE As String = "FA"

Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"

Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"
Private Const WBO_REVENUE_CATEGORY As String = "Revenue Category"
Private Const WBO_COST_CENTER As String = "Cost Center"
Private Const WBO_EVENT As String = "Events - Custom Org"
Private Const WBO_CAMP As String = "Camps - Custom Org"
Private Const WBO_FA_LEARNING As String = "FA Learning - Custom Org"
Private Const WBO_COMMERCIAL_PARTNER As String = "Commercial Partners - Custom Worktag"
Private Const WBO_GRANT As String = "Grant"
Private Const WBO_PROGRAM As String = "Program - Delivered Org"
Private Const WBO_PROGRAMME As String = "Program - Delivered Org"
Private Const WBO_PROJECT As String = "Projects (Opex)"
Private Const WBO_SALES_ITEM As String = "Sales Item"

'  ============================================================================
'  Target EIB column letters -- Submit_Customer_Invoice_v46.1.xlsx (Submit Customer Invoice)
'  These constants are the 1:1 mirror of ar_multi_line_invoice.yaml customer_invoice_eib_columns
'  If any YAML letter changes, the hygiene check will fail and you must update this block.
'  ============================================================================
Private Const TGT_CUSTOMER_INVOICE_REF_ID As String = "D"
Private Const TGT_CUSTOMER_INVOICE_ID As String = "G"
Private Const TGT_INVOICE_NUMBER As String = "BO"
Private Const TGT_SPREADSHEET_KEY As String = "B"
Private Const TGT_ADD_ONLY As String = "C"
Private Const TGT_HEADER_COMMENT As String = "E"
Private Const TGT_SUBMIT As String = "H"
Private Const TGT_LOCKED_IN_WORKDAY As String = "I"
Private Const TGT_COMPANY As String = "J"
Private Const TGT_ROW_ID_HEADER_1 As String = "K"
Private Const TGT_ROW_ID_HEADER_2 As String = "L"
Private Const TGT_ROW_ID_HEADER_3 As String = "M"
Private Const TGT_CURRENCY As String = "AI"
Private Const TGT_CUSTOMER As String = "AJ"
Private Const TGT_SOLD_TO_CUSTOMER As String = "AK"
Private Const TGT_BILLABLE_PROJECT_HEADER As String = "BN"
Private Const TGT_INVOICE_DATE As String = "BR"
Private Const TGT_PAYMENT_STATUS As String = "CK"
Private Const TGT_DOCUMENT_STATUS As String = "CL"
Private Const TGT_AMOUNT_DUE As String = "CM"
Private Const TGT_CUSTOMER_INVOICE_TYPE As String = "CO"
Private Const TGT_HEADER_TAX_CODE As String = "CP"
Private Const TGT_CUSTOMER_PO_NUMBER As String = "CQ"
Private Const TGT_INCLUDE_CUSTOMER_WORKTAGS As String = "CY"
Private Const TGT_HEADER_WORKTAGS_PLUS As String = "CZ"
Private Const TGT_HEADER_MEMO As String = "DB"
Private Const TGT_ROW_ID_LINE As String = "DS"
Private Const TGT_LINE_REFERENCE As String = "DU"
Private Const TGT_LINE_ORDER As String = "DV"
Private Const TGT_SALES_ITEM As String = "EJ"
Private Const TGT_REVENUE_CATEGORY As String = "EM"
Private Const TGT_LINE_ITEM_DESCRIPTION As String = "EN"
Private Const TGT_LINE_TAX_APPLICABILITY As String = "FA"
Private Const TGT_LINE_TAX_CODE As String = "FB"
Private Const TGT_LINE_QUANTITY As String = "FH"
Private Const TGT_LINE_UNIT_OF_MEASURE As String = "FI"
Private Const TGT_LINE_UNIT_COST As String = "FL"
Private Const TGT_LINE_EXTENDED_AMOUNT As String = "FO"
Private Const TGT_LINE_TRANSACTION_DATE As String = "FP"
Private Const TGT_LINE_FROM_DATE As String = "FQ"
Private Const TGT_LINE_TO_DATE As String = "FR"
Private Const TGT_DEFERRED_REVENUE As String = "FS"
Private Const TGT_LINE_MEMO As String = "FV"
Private Const TGT_BILLABLE_PROJECT_LINE As String = "GF"
Private Const TGT_COST_CENTRE As String = "FX"
Private Const TGT_REPORTING_TYPE As String = "FY"
Private Const TGT_CAMP As String = "FZ"
Private Const TGT_EVENT As String = "GA"
Private Const TGT_COMMERCIAL_PARTNER As String = "GB"
Private Const TGT_FA_LEARNING_COURSE As String = "GC"
Private Const TGT_GRANT As String = "GD"
Private Const TGT_PROGRAMME As String = "GE"
Private Const TGT_PROJECT_LINE_WORKTAG As String = "GF"

Public Sub RunArMultiLineInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim customerSpreadsheetPath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim selectionKey As String
    Dim customerWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim mappingWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim customerLookup As Collection
    Dim sourceTable As Variant
    Dim mappingTable As Variant
    Dim revenueCategoryLookup As Collection
    Dim costCenterLookup As Collection
    Dim eventLookup As Collection
    Dim campLookup As Collection
    Dim faLearningLookup As Collection
    Dim commercialPartnerLookup As Collection
    Dim grantLookup As Collection
    Dim programLookup As Collection
    Dim projectLookup As Collection
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim lineIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim totalGross As Double
    Dim invoiceDateColumn As Long, customerIdColumn As Long, lineDescColumn As Long
    Dim lineDateColumn As Long, qtyColumn As Long
    Dim netColumn As Long, vatColumn As Long, grossColumn As Long
    Dim descBottomColumn As Long, costCodeColumn As Long, invoiceNumberColumn As Long
    Dim headerCustomerId As String
    Dim workdayCustomerId As String
    Dim headerInvoiceDate As String
    Dim headerDescBottom As String
    Dim headerCompanyId As String
    Dim headerInvoiceNumber As String
    Dim issues As Collection
    Dim outputPath As String
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
    Dim requiredHeaders As Variant
    Dim resolvedSourceSheet As Worksheet
    Dim resolvedSourceSheetName As String
    Dim headerRowIndex As Long
    Dim dataStartRowIndex As Long
    Dim schemaMsg As String
    Dim schemaHeaderRow As Long
    Dim schemaDiag As String
    Dim firstDataRowIndex As Long
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
    customerSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_CUSTOMER_SPREADSHEET)
    mappingSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_MAPPING_SPREADSHEET)
    mappingSheetName = ReadControlValue(controlSheet, CONTROL_LABEL_MAPPING_SHEET)
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If StrComp(worksheetResolutionStr, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetNameStr
    End If
    EnsureNonBlank mappingSpreadsheetPath, CONTROL_LABEL_MAPPING_SPREADSHEET
    If LenB(mappingSheetName) = 0 Then mappingSheetName = DEFAULT_MAPPING_SHEET_NAME

    Set customerWorkbook = OpenWorkbookReadOnly(customerSpreadsheetPath)
    Set customerLookup = BuildReferenceLookupFromWorkbook(customerWorkbook, "Customer")
    customerWorkbook.Close SaveChanges:=False
    Set customerWorkbook = Nothing

    Set mappingWorkbook = OpenWorkbookReadOnly(mappingSpreadsheetPath)
    mappingTable = ReadWorksheetTable(ResolveSourceWorksheet(mappingWorkbook, mappingSheetName))
    mappingWorkbook.Close SaveChanges:=False
    Set mappingWorkbook = Nothing

    Set revenueCategoryLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_REVENUE_CATEGORY)
    Set costCenterLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_COST_CENTER)
    Set eventLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_EVENT)
    Set campLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_CAMP)
    Set faLearningLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_FA_LEARNING)
    Set commercialPartnerLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_COMMERCIAL_PARTNER)
    Set grantLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_GRANT)
    Set programLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_PROGRAM)
    Set projectLookup = BuildBusinessObjectHierarchyLookupMulti(mappingTable, WBO_PROJECT, WBO_PROJECT_CAPEX)

    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modArMultiLineCustInvoice", 10)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    Set resolvedSourceSheet = ResolveSourceWorksheet(sourceWorkbook, sheetToResolve)
    resolvedSourceSheetName = resolvedSourceSheet.Name
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
        Err.Raise vbObjectError + 1910, "modArMultiLineCustInvoice.RunArMultiLineInvoiceCustomerInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issues.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "DATE OF INVOICE")
    customerIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "CUSTOMER ID")
    lineDescColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "INVOICE LINE DESCRIPTION")
    lineDateColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Date")
    qtyColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "QUANTITY")
    netColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "NET")
    vatColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "VAT")
    grossColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "GROSS")
    descBottomColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "DESCRIPTION LINE BOTTOM OF INVOICE")
    costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "COST CODE")
    invoiceNumberColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "INVOICE NUMBER")

    firstDataRowIndex = dataStartRowIndex

    headerCustomerId = Trim$(CStr(sourceTable(firstDataRowIndex, customerIdColumn)))
    If LenB(headerCustomerId) = 0 Then
        Err.Raise vbObjectError + 2201, "modArMultiLineCustInvoice.RunArMultiLineInvoiceCustomerInvoice", _
            "First data row's CUSTOMER ID is blank -- cannot build a multi-line invoice without a customer."
    End If
    workdayCustomerId = LookupCollectionValue(customerLookup, headerCustomerId)
    If LenB(workdayCustomerId) = 0 Then
        Err.Raise vbObjectError + 2202, "modArMultiLineCustInvoice.RunArMultiLineInvoiceCustomerInvoice", _
            "First data row's CUSTOMER ID '" & headerCustomerId & "' is not found in the customer extract." & _
            " A multi-line invoice run needs a single resolved customer for the whole invoice header."
    End If

    headerInvoiceDate = NormalizeSourceDate(sourceTable(firstDataRowIndex, invoiceDateColumn))
    If LenB(headerInvoiceDate) = 0 Then
        Err.Raise vbObjectError + 2203, "modArMultiLineCustInvoice.RunArMultiLineInvoiceCustomerInvoice", _
            "First data row's DATE OF INVOICE is blank -- cannot build a multi-line invoice header."
    End If
    If descBottomColumn > 0 Then
        headerDescBottom = CleanMemoToSingleLine(CStr(sourceTable(firstDataRowIndex, descBottomColumn)))
    Else
        headerDescBottom = vbNullString
    End If
    If invoiceNumberColumn > 0 Then
        headerInvoiceNumber = Trim$(CStr(sourceTable(firstDataRowIndex, invoiceNumberColumn)))
    Else
        headerInvoiceNumber = vbNullString
    End If
    headerCompanyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, COMPANY_DEFAULT_FALLBACK)

    totalGross = 0#
    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If Not LineIsSkippable(sourceTable, rowIndex, grossColumn, costCodeColumn) Then
            totalGross = totalGross + ResolveAmount(sourceTable(rowIndex, grossColumn))
        End If
    Next rowIndex

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "MV"
    targetRow = TARGET_DATA_START_ROW
    lineIndex = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If LineIsSkippable(sourceTable, rowIndex, grossColumn, costCodeColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Skipping invoice line: missing GROSS / COST CODE or zero GROSS amount.", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, lineDescColumn)))
        Else
            lineIndex = lineIndex + 1
            rowsProcessed = rowsProcessed + 1
            WriteOutputLine runnerWorkbook, targetSheet, targetRow, 1, lineIndex, _
                sourceTable, rowIndex, _
                lineDescColumn, lineDateColumn, qtyColumn, netColumn, vatColumn, grossColumn, costCodeColumn, _
                headerCompanyId, workdayCustomerId, headerInvoiceDate, headerInvoiceNumber, _
                headerDescBottom, totalGross, _
                revenueCategoryLookup, costCenterLookup, _
                eventLookup, campLookup, faLearningLookup, commercialPartnerLookup, _
                grantLookup, programLookup, projectLookup
            targetRow = targetRow + 1
        End If
    Next rowIndex

    If lineIndex = 0 Then
        Err.Raise vbObjectError + 2204, "modArMultiLineCustInvoice.RunArMultiLineInvoiceCustomerInvoice", _
            "Zero invoice lines survived filtering -- the multi-line invoice would contain no line items. Aborting run."
    End If

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source line(s) with missing GROSS / COST CODE."
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
    If Not customerWorkbook Is Nothing Then customerWorkbook.Close SaveChanges:=False
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    If Not mappingWorkbook Is Nothing Then mappingWorkbook.Close SaveChanges:=False
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
    Err.Raise failureNumber, "modArMultiLineCustInvoice.RunArMultiLineInvoiceCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputLine( _
    ByVal runnerWorkbook As Workbook, ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal invoiceIndex As Long, ByVal lineNumber As Long, ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, ByVal lineDescColumn As Long, ByVal lineDateColumn As Long, ByVal qtyColumn As Long, ByVal netColumn As Long, ByVal vatColumn As Long, ByVal grossColumn As Long, ByVal costCodeColumn As Long, _
    ByVal headerCompanyId As String, ByVal headerCustomerId As String, ByVal headerInvoiceDate As String, ByVal headerInvoiceNumber As String, ByVal headerDescBottom As String, _
    ByVal headerTotalGross As Double, ByVal revenueCategoryLookup As Collection, ByVal costCenterLookup As Collection, ByVal eventLookup As Collection, ByVal campLookup As Collection, _
    ByVal faLearningLookup As Collection, ByVal commercialPartnerLookup As Collection, ByVal grantLookup As Collection, ByVal programLookup As Collection, ByVal projectLookup As Collection _
)
    Dim lineDate As String
    Dim grossAmount As Double
    Dim netAmount As Double
    Dim qtyValue As Double
    Dim costCodeValue As String
    Dim revenueCategory As String
    Dim costCenterRef As String
    Dim lineDescValue As String
    Dim vatRate As Long
    Dim lineTaxApplicability As String
    Dim lineTaxCode As String
    Dim reportingType As String
    Dim eventRef As String
    Dim campRef As String
    Dim faLearningRef As String
    Dim commercialPartnerRef As String
    Dim grantRef As String
    Dim programRef As String
    Dim projectRef As String
    Dim worktagIndex As Long
    Dim worktagColumns As Variant
    Dim worktagValue As Variant

    grossAmount = ResolveAmount(sourceTable(sourceRow, grossColumn))
    If qtyColumn > 0 Then qtyValue = ResolveAmount(sourceTable(sourceRow, qtyColumn)) Else qtyValue = 1#
    If netColumn > 0 Then netAmount = ResolveAmount(sourceTable(sourceRow, netColumn)) Else netAmount = grossAmount
    lineDate = headerInvoiceDate
    lineDescValue = CleanMemoToSingleLine(CStr(sourceTable(sourceRow, lineDescColumn)))
    costCodeValue = Trim$(CStr(sourceTable(sourceRow, costCodeColumn)))

    If vatColumn > 0 Then vatRate = DeriveVatPercentage(sourceTable(sourceRow, vatColumn), netAmount) Else vatRate = 0
    lineTaxApplicability = TaxValueForRate(vatRate, TAX_APPLICABILITY_ZERO, TAX_APPLICABILITY_STANDARD)
    lineTaxCode = TaxValueForRate(vatRate, TAX_CODE_ZERO, TAX_CODE_STANDARD)

    If LenB(costCodeValue) > 0 Then
        revenueCategory = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(revenueCategoryLookup, costCodeValue)
        costCenterRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, costCodeValue)
        reportingType = ResolveReportingTypeFromFinalCode(costCodeValue)
        eventRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(eventLookup, costCodeValue)
        campRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(campLookup, costCodeValue)
        faLearningRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(faLearningLookup, costCodeValue)
        commercialPartnerRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(commercialPartnerLookup, costCodeValue)
        grantRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(grantLookup, costCodeValue)
        programRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(programLookup, costCodeValue)
        projectRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(projectLookup, costCodeValue)
    End If

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = invoiceIndex
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_COMMENT)).Value = OUTPUT_AUTO_COMPLETE
    If LenB(headerInvoiceNumber) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REF_ID)).Value = headerInvoiceNumber
        targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).Value = headerInvoiceNumber
        targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_NUMBER)).Value = headerInvoiceNumber
    End If
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LOCKED_IN_WORKDAY)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = headerCompanyId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = headerCustomerId

    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = headerInvoiceDate

    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = headerTotalGross
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_MEMO)).Value = headerDescBottom
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(lineNumber)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_REFERENCE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = CStr(lineNumber)

    If LenB(revenueCategory) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = revenueCategory

    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = lineDescValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = lineTaxApplicability
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = lineTaxCode
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = qtyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).Value = grossAmount

    If LenB(lineDate) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).NumberFormat = "@"
        targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).Value = lineDate
    End If

    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_MEMO)).Value = lineDescValue

    If LenB(projectRef) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_BILLABLE_PROJECT_HEADER)).Value = projectRef
        targetSheet.Cells(targetRow, ColumnIndex(TGT_BILLABLE_PROJECT_LINE)).Value = projectRef
    End If

    ' 9-slot worktags in the canonical Submit_Customer_Invoice_v46.1 columns FX FY FZ GA GB GC GD GE GF
    ' order of worktags matches the AR profile YAML worktags_plus_order declaration:
    ' 0: sales_item (EJ - multi-line source has no explicit sales_item column; written blank if unresolved),
    ' 1: reporting_type (FY), 2: event (GA), 3: camp (FZ), 4: fa_learning (GC), 5: commercial_partner (GB),
    ' 6: grant (GD), 7: programme (GE), 8: cost_centre (FX)
    ' Note: sales_item intentionally not written here (multi-line source has no SALES ITEM column);
    '       populate the remaining 8 worktag ref slots.
    worktagColumns = Array( _
        TGT_REPORTING_TYPE, _
        TGT_EVENT, _
        TGT_CAMP, _
        TGT_FA_LEARNING_COURSE, _
        TGT_COMMERCIAL_PARTNER, _
        TGT_GRANT, _
        TGT_PROGRAMME, _
        TGT_COST_CENTRE _
    )
    worktagValue = Empty
    For worktagIndex = LBound(worktagColumns) To UBound(worktagColumns)
        Select Case worktagIndex
            Case 0: worktagValue = reportingType
            Case 1: worktagValue = eventRef
            Case 2: worktagValue = campRef
            Case 3: worktagValue = faLearningRef
            Case 4: worktagValue = commercialPartnerRef
            Case 5: worktagValue = grantRef
            Case 6: worktagValue = programRef
            Case 7: worktagValue = costCenterRef
        End Select
        If LenB(CStr(worktagValue)) > 0 Then
            targetSheet.Cells(targetRow, ColumnIndex(CStr(worktagColumns(worktagIndex)))).Value = CStr(worktagValue)
        End If
    Next worktagIndex

    ' Write project_line worktag if resolved (column GF = TGT_PROJECT_LINE_WORKTAG)
    If LenB(projectRef) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_PROJECT_LINE_WORKTAG)).Value = projectRef
    End If
End Sub

Private Function CleanMemoToSingleLine(ByVal rawText As String) As String
    Dim temp As String
    temp = Replace$(rawText, vbCrLf, " ")
    temp = Replace$(temp, vbCr, " ")
    temp = Replace$(temp, vbLf, " ")
    temp = Replace$(temp, vbTab, " ")
    Do While InStr(temp, "  ") > 0
        temp = Replace$(temp, "  ", " ")
    Loop
    CleanMemoToSingleLine = Trim$(temp)
End Function

Private Function DeriveVatPercentage(ByVal vatValue As Variant, ByVal netAmountValue As Variant) As Long
    Dim vatAmount As Double
    Dim netAmount As Double

    vatAmount = ResolveAmount(vatValue)
    netAmount = ResolveAmount(netAmountValue)

    If netAmount <= 0# Then
        DeriveVatPercentage = 0
        Exit Function
    End If

    DeriveVatPercentage = CLng(Int((vatAmount / netAmount) * 100# + 0.5))
End Function

Private Function LineIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal grossColumn As Long, _
    ByVal costCodeColumn As Long _
) As Boolean
    LineIsSkippable = _
        ResolveAmount(tableData(rowIndex, grossColumn)) = 0# Or _
        LenB(Trim$(CStr(tableData(rowIndex, costCodeColumn)))) = 0
End Function
