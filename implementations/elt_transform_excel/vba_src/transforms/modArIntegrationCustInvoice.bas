Attribute VB_Name = "modArIntegrationCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - NFC"
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"

'  Customer Invoice ID (G) prefix -- see ar_integration_invoice.yaml
'  transform_constants and transform_specs mapping.g. (D is left blank.)
Private Const REFERENCE_ID_PREFIX As String = "integration_"

Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

'  ============================================================================
'  Target EIB column letters -- Submit_Customer_Invoice_v46.1.xlsx (Submit Customer Invoice)
'  These constants are the 1:1 mirror of ar_integration_invoice.yaml customer_invoice_eib_columns.
'  If any YAML letter changes, the hygiene check will fail and you must update this block.
'  ============================================================================
Private Const TGT_CUSTOMER_INVOICE_REF_ID As String = "D"
Private Const TGT_CUSTOMER_INVOICE_ID As String = "G"
Private Const TGT_INVOICE_NUMBER As String = "BO"
Private Const TGT_SPREADSHEET_KEY As String = "B"
Private Const TGT_ADD_ONLY As String = "C"
Private Const TGT_HEADER_COMMENT As String = "E"
Private Const TGT_SUBMIT As String = "H"
Private Const TGT_COMPANY As String = "J"
Private Const TGT_ROW_ID_HEADER_1 As String = "K"
Private Const TGT_ROW_ID_HEADER_2 As String = "L"
Private Const TGT_ROW_ID_HEADER_3 As String = "M"
Private Const TGT_CURRENCY As String = "AI"
Private Const TGT_CUSTOMER As String = "AJ"
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

'  --- Code lists tab return value packing order inside courseHierarchyLookup pipe-delimited string ---
'  Indexes are 1-based after Split(pipeDelimited, "|").  Keep in lock-step with
'  ar_integration_invoice.yaml transform_constants.source_lookups.course_hierarchy.return_columns.
Private Const COURSE_PACK_COMPANY As Long = 1
Private Const COURSE_PACK_COST_CENTRE As Long = 2
Private Const COURSE_PACK_REPORTING_TYPE As Long = 3
Private Const COURSE_PACK_FA_LEARNING As Long = 4
Private Const COURSE_PACK_CAMP As Long = 5
Private Const COURSE_PACK_COMMERCIAL_PARTNER As Long = 6
Private Const COURSE_PACK_EVENT As Long = 7
Private Const COURSE_PACK_GRANT As Long = 8
Private Const COURSE_PACK_PROGRAMME As Long = 9
Private Const COURSE_PACK_REVENUE_CATEGORY As Long = 10
Private Const COURSE_PACK_SLOT_COUNT As Long = 10

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

Public Sub RunArIntegrationCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim invoiceGroupCounters As New Collection
    Dim integrationTable As Variant
    Dim codesTable As Variant
    Dim courseHierarchyLookup As Collection
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unresolvedCustomerRows As Long
    Dim unresolvedCourseRows As Long
    Dim invoiceDateColumn As Long
    Dim customerIdColumn As Long
    Dim workdayIdColumn As Long
    Dim orgNameColumn As Long
    Dim courseNameColumn As Long
    Dim descriptionColumn As Long
    Dim qtyColumn As Long
    Dim netColumn As Long
    Dim vatColumn As Long
    Dim grossColumn As Long
    Dim poNumberColumn As Long
    Dim emailColumn As Long
    Dim sourceCustomerId As String
    Dim workdayCustomerId As String
    Dim issues As Collection
    Dim unresolvedCounts As Collection
    Dim unresolvedKeys As Collection
    Dim outputPath As String
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
    Dim requiredHeaders As Variant
    Dim resolvedIntegrationSheet As Worksheet
    Dim resolvedCodesSheet As Worksheet
    Dim resolvedIntegrationSheetName As String
    Dim headerRowIndex As Long
    Dim dataStartRowIndex As Long
    Dim schemaMsg As String
    Dim schemaHeaderRow As Long
    Dim schemaDiag As String
    Dim worksheetNameStr As String
    Dim worksheetResolutionStr As String
    Dim sourceHeaderRow As Long
    Dim sourceDataStartRow As Long
    Dim worktagsHeaderRowIndex As Long
    Dim costCentreColumn As Long
    Dim reportingTypeColumn As Long
    Dim revenueCategoryColumn As Long
    Dim faLearningColumn As Long
    Dim codesHeaderRowIndex As Long
    Dim codesDataStartRowIndex As Long
    Dim codesCourseColumn As Long
    Dim codesCompanyColumn As Long
    Dim codesCostCentreColumn As Long
    Dim codesReportingTypeColumn As Long
    Dim codesFaLearningColumn As Long
    Dim codesCampColumn As Long
    Dim codesCommercialPartnerColumn As Long
    Dim codesEventColumn As Long
    Dim codesGrantColumn As Long
    Dim codesProgramColumn As Long
    Dim codesRevenueCatColumn As Long
    Dim poNumberValueLcl As String

    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)

    Set issues = New Collection
    Set unresolvedCounts = New Collection
    Set unresolvedKeys = New Collection

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH
    mappingSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_MAPPING_SPREADSHEET)
    mappingSheetName = ReadControlValue(controlSheet, CONTROL_LABEL_MAPPING_SHEET)
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If LenB(mappingSheetName) = 0 Then mappingSheetName = DEFAULT_MAPPING_SHEET_NAME

    ' 2026-09-22 amendment: Customer (AJ) no longer resolves via the Customer Spreadsheet
    ' extract (Customer ID -> Reference ID lookup) -- Finance now provides the resolved
    ' Workday customer reference directly per row in the new "WORKDAY ID" source column.
    ' See workdayIdColumn below.

    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modArIntegrationCustInvoice", 4)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)

    Set resolvedIntegrationSheet = ResolveSourceWorksheet(sourceWorkbook, "Integration")
    resolvedIntegrationSheetName = resolvedIntegrationSheet.Name
    integrationTable = ReadWorksheetTable(resolvedIntegrationSheet)

    Set resolvedCodesSheet = ResolveSourceWorksheet(sourceWorkbook, "Code lists")
    codesTable = ReadWorksheetTable(resolvedCodesSheet)

    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing
    Set resolvedIntegrationSheet = Nothing
    Set resolvedCodesSheet = Nothing

    schemaMsg = ValidateSourceWorkbookSchema( _
        "Integration", _
        integrationTable, _
        sourceHeaderRow, _
        sourceDataStartRow, _
        requiredHeaders, _
        , _
        schemaHeaderRow, _
        schemaDiag _
    )
    If LenB(schemaMsg) > 0 Then
        Err.Raise vbObjectError + 1920, "modArIntegrationCustInvoice.RunArIntegrationCustomerInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issues.Add "[Schema Integration] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(integrationTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow integrationTable, headerRowIndex, requiredHeaders

    invoiceDateColumn = FindColumnIndexInRow(integrationTable, headerRowIndex, "Date")
    customerIdColumn = FindColumnIndexInRow(integrationTable, headerRowIndex, "Customer ID")
    workdayIdColumn = FindColumnIndexInRow(integrationTable, headerRowIndex, "WORKDAY ID")
    orgNameColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Organisation Name")
    courseNameColumn = FindColumnIndexInRow(integrationTable, headerRowIndex, "Course Name")
    descriptionColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Description")
    If descriptionColumn = 0 Then descriptionColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Course Code")
    qtyColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Quantity")

    netColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Net " & ChrW(163))
    If netColumn = 0 Then netColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Net PS")
    If netColumn = 0 Then netColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Net")

    vatColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Vat " & ChrW(163))
    If vatColumn = 0 Then vatColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "VAT " & ChrW(163))
    If vatColumn = 0 Then vatColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Vat PS")
    If vatColumn = 0 Then vatColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "VAT PS")
    If vatColumn = 0 Then vatColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Vat")
    If vatColumn = 0 Then vatColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "VAT")

    grossColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Gross Price " & ChrW(163))
    If grossColumn = 0 Then grossColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Gross Price PS")
    If grossColumn = 0 Then grossColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Gross Price")

    poNumberColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "PO Number")
    emailColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Email Address- can these be copied in")
    If emailColumn = 0 Then emailColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Email Address")

    ' The "Combined 30-34" Integration tab has a two-row header: row 1 carries the field
    ' names plus the "Required Fields"/"Optional Fields" group labels; row 2 carries the
    ' real sub-headers for the fields grouped under those labels (Cost Centre, Reporting
    ' Type, Revenue Category, ...). 2026-09-22: Cost Centre / Reporting Type / Revenue
    ' Category are read directly off that row 2, per row, instead of via the Code lists tab.
    worktagsHeaderRowIndex = headerRowIndex + 1
    costCentreColumn = FindOptionalColumnIndexInRow(integrationTable, worktagsHeaderRowIndex, "Cost Centre")
    reportingTypeColumn = FindOptionalColumnIndexInRow(integrationTable, worktagsHeaderRowIndex, "Reporting Type")
    revenueCategoryColumn = FindOptionalColumnIndexInRow(integrationTable, worktagsHeaderRowIndex, "Revenue Category")
    faLearningColumn = FindOptionalColumnIndexInRow(integrationTable, worktagsHeaderRowIndex, "FA Learning Course")

    ' NOTE (2026-09-22): the "Combined 30-34" source workbook's Code lists tab is a plain
    ' 3-column dropdown reference list (Cost Centres / Reporting Type / Revenue Category)
    ' with NO "Course" key column, so this Course-Name-keyed hierarchy lookup can no
    ' longer resolve against it. Every lookup below is made OPTIONAL (was required) so
    ' a missing "Course" header degrades to "no course hierarchy available" (Company
    ' falls back to COMPANY_DEFAULT_FALLBACK) instead of a hard crash. Cost Centre /
    ' Reporting Type / Revenue Category / FA Learning Course no longer depend on this
    ' lookup at all -- see costCentreColumn / reportingTypeColumn / revenueCategoryColumn
    ' / faLearningColumn above, read directly off the Integration sheet's own row 2.
    ' Camp, Event, Commercial Partner, Grant, and Programme are intentionally always
    ' blank per 2026-09-22 Finance instruction (see WriteOutputRow). Company still
    ' relies on this Code lists lookup and so stays on its fallback. Project and Sales Item
    ' have no source column and stay blank.
    codesHeaderRowIndex = 1
    codesDataStartRowIndex = 2
    codesCourseColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Course")
    codesCompanyColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Company")
    codesCostCentreColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Cost centre")
    codesReportingTypeColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Reporting Type")
    codesFaLearningColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "FA Learning courses")
    codesCampColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Camp")
    codesCommercialPartnerColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Commerical Partners")
    codesEventColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Events")
    codesGrantColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Grant")
    codesProgramColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Program")
    codesRevenueCatColumn = FindOptionalColumnIndexInRow(codesTable, codesHeaderRowIndex, "Revenue Cat")

    Set courseHierarchyLookup = BuildCourseHierarchyLookup( _
        codesTable, codesDataStartRowIndex, codesCourseColumn, _
        codesCompanyColumn, codesCostCentreColumn, codesReportingTypeColumn, _
        codesFaLearningColumn, codesCampColumn, codesCommercialPartnerColumn, _
        codesEventColumn, codesGrantColumn, codesProgramColumn, codesRevenueCatColumn _
    )

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "MV"
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    For rowIndex = dataStartRowIndex To UBound(integrationTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(integrationTable, 1), runStartTime
        If IntegrationRowIsSkippable(integrationTable, rowIndex, invoiceDateColumn, customerIdColumn, netColumn, vatColumn, grossColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Missing required field (Date / Customer ID / fully-zero Net+Vat+Gross)", _
                    vbNullString, _
                    SafeCellStr(integrationTable(rowIndex, customerIdColumn))
            Else
                sourceCustomerId = SafeCellStr(integrationTable(rowIndex, customerIdColumn))
                ' 2026-09-22 amendment: Customer (AJ) is resolved from the source's own
                ' "WORKDAY ID" column, not via a Customer Spreadsheet extract lookup.
                workdayCustomerId = SafeCellStr(integrationTable(rowIndex, workdayIdColumn))
            If LenB(workdayCustomerId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unresolvedCustomerRows = unresolvedCustomerRows + 1
                IncrementCount unresolvedCounts, unresolvedKeys, sourceCustomerId
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "WORKDAY ID not provided on source row", _
                    sourceCustomerId, _
                    SafeCellStr(integrationTable(rowIndex, courseNameColumn))
            Else
                rowsProcessed = rowsProcessed + 1
                If poNumberColumn > 0 Then
                    poNumberValueLcl = SafeCellStr(integrationTable(rowIndex, poNumberColumn))
                Else
                    poNumberValueLcl = vbNullString
                End If
                WriteOutputRow runnerWorkbook, targetSheet, targetRow, rowsProcessed, integrationTable, rowIndex, _
                    invoiceDateColumn, orgNameColumn, courseNameColumn, descriptionColumn, _
                    qtyColumn, netColumn, vatColumn, grossColumn, poNumberColumn, emailColumn, _
                    costCentreColumn, reportingTypeColumn, revenueCategoryColumn, faLearningColumn, _
                    workdayCustomerId, courseHierarchyLookup, unresolvedCourseRows, _
                    unresolvedCounts, unresolvedKeys, invoiceGroupCounters, poNumberValueLcl, sourceCustomerId
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing required fields."
    End If
    If unresolvedCustomerRows > 0 Then
        issues.Add BuildRejectedLookupIssue(unresolvedCounts, unresolvedKeys, "WORKDAY ID")
    End If
    If unresolvedCourseRows > 0 Then
        issues.Add "Course Name did not resolve in Code lists tab for " & CStr(unresolvedCourseRows) & _
                  " row(s) - those rows were written with blank worktags (silent_blank rule)."
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
    Err.Raise failureNumber, "modArIntegrationCustInvoice.RunArIntegrationCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputRow( _
    ByVal runnerWorkbook As Workbook, _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal invoiceDateColumn As Long, ByVal orgNameColumn As Long, ByVal courseNameColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal qtyColumn As Long, ByVal netColumn As Long, ByVal vatColumn As Long, _
    ByVal grossColumn As Long, ByVal poNumberColumn As Long, ByVal emailColumn As Long, _
    ByVal costCentreColumn As Long, ByVal reportingTypeColumn As Long, ByVal revenueCategoryColumn As Long, _
    ByVal faLearningColumn As Long, _
    ByVal outputCustomerId As String, _
    ByVal courseHierarchyLookup As Collection, _
    ByRef unresolvedCourseRows As Long, _
    ByVal unresolvedCounts As Collection, _
    ByVal unresolvedKeys As Collection, _
    ByVal invoiceGroupCounters As Collection, _
    ByVal poNumberValue As String, _
    ByVal sourceCustomerId As String _
)
    Dim companyId As String
    Dim invoiceDate As String
    Dim orderDate As String
    Dim grossAmount As Double
    Dim qtyValue As Double
    Dim courseNameValue As String
    Dim descriptionValue As String
    Dim emailValue As String
    Dim orgNameValue As String
    Dim coursePacked As String
    Dim courseParts As Variant
    Dim revenueCategory As String
    Dim costCenterRef As String
    Dim reportingType As String
    Dim campRef As String
    Dim eventRef As String
    Dim commercialPartnerRef As String
    Dim faLearningRef As String
    Dim grantRef As String
    Dim programRef As String
    Dim projectRef As String
    Dim headerMemo As String
    Dim lineDescription As String
    Dim lineMemo As String
    Dim invoiceRefId As String
    Dim referenceId As String
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long
    Dim worktagIndex As Long
    Dim worktagColumns As Variant
    Dim worktagValue As Variant
    Dim foundCourse As Boolean

    invoiceDate = NormalizeSourceDate(SafeCellStr(sourceTable(sourceRow, invoiceDateColumn)))
    ' Transaction Date FP mirrors Invoice Date BR (both = source "Date"); see transform_specs target: fp.
    orderDate = invoiceDate

    ' Quantity (FH) = source "Quantity" (col H), capped at 1 per 2026-09-30 instruction:
    ' Unit Cost (FL) carries the full row Gross, so FH x FL must equal the row total.
    ' Blank, text or <= 0 defaults to 1.
    If qtyColumn > 0 Then qtyValue = SafeCellDbl(sourceTable(sourceRow, qtyColumn)) Else qtyValue = 1#
    If qtyValue <= 0# Then qtyValue = 1#
    If qtyValue > 1# Then qtyValue = 1#

    ' Amount Due (CM), Unit Cost (FL) and Extended Amount (FO) = source Gross Price
    ' (col K) per 2026-09-30 instruction. Line tax (FA/FB) is blank, so Workday adds no
    ' VAT on top. See transform_specs targets cm/fl/fo.
    If grossColumn > 0 Then grossAmount = SafeCellDbl(sourceTable(sourceRow, grossColumn)) Else grossAmount = 0#
    ' Line Tax Applicability (FA) / Line Tax Code (FB) are left blank on every row, so no
    ' Net/VAT tax triple is derived here (see transform_specs tax_rules).

    courseNameValue = SafeCellStr(sourceTable(sourceRow, courseNameColumn))
    If descriptionColumn > 0 Then descriptionValue = SafeCellStr(sourceTable(sourceRow, descriptionColumn)) Else descriptionValue = vbNullString
    If emailColumn > 0 Then emailValue = SafeCellStr(sourceTable(sourceRow, emailColumn)) Else emailValue = vbNullString
    If orgNameColumn > 0 Then orgNameValue = CleanMemoToSingleLine(SafeCellStr(sourceTable(sourceRow, orgNameColumn))) Else orgNameValue = vbNullString
    If poNumberColumn > 0 Then poNumberValue = SafeCellStr(sourceTable(sourceRow, poNumberColumn)) Else poNumberValue = vbNullString

    foundCourse = False
    If LenB(courseNameValue) > 0 Then
        coursePacked = LookupCollectionValue(courseHierarchyLookup, courseNameValue)
        If LenB(coursePacked) > 0 Then
            foundCourse = True
        End If
    End If

    If foundCourse Then
        courseParts = Split(coursePacked, "|", COURSE_PACK_SLOT_COUNT)
        If UBound(courseParts) >= (COURSE_PACK_SLOT_COUNT - 1) Then
            companyId = Trim$(CStr(courseParts(COURSE_PACK_COMPANY - 1)))
        End If
    Else
        unresolvedCourseRows = unresolvedCourseRows + 1
        If LenB(courseNameValue) > 0 Then
            IncrementCount unresolvedCounts, unresolvedKeys, courseNameValue
        End If
    End If

    ' 2026-09-22: Camp / Event / Commercial Partner / Grant / Programme are
    ' intentionally left blank -- no longer read from the Code lists tab course
    ' lookup (per Finance instruction; Code lists tab no longer supports it).
    campRef = vbNullString
    commercialPartnerRef = vbNullString
    eventRef = vbNullString
    grantRef = vbNullString
    programRef = vbNullString

    ' Cost Centre / Reporting Type / Revenue Category / FA Learning Course no longer
    ' come from the Code lists tab course lookup above -- read directly off this row's own
    ' Integration-sheet columns (see worktagsHeaderRowIndex resolution in the caller).
    If costCentreColumn > 0 Then costCenterRef = SafeCellStr(sourceTable(sourceRow, costCentreColumn)) Else costCenterRef = vbNullString
    ' Reporting Type (FY): drop the leading "<number>: " prefix from col P (e.g.
    ' "1: FA Operations" -> "FA Operations"). See derived_worktags.reporting_type.
    If reportingTypeColumn > 0 Then reportingType = StripLeadingNumberPrefix(SafeCellStr(sourceTable(sourceRow, reportingTypeColumn))) Else reportingType = vbNullString
    If revenueCategoryColumn > 0 Then revenueCategory = SafeCellStr(sourceTable(sourceRow, revenueCategoryColumn)) Else revenueCategory = vbNullString
    If faLearningColumn > 0 Then faLearningRef = SafeCellStr(sourceTable(sourceRow, faLearningColumn)) Else faLearningRef = vbNullString

    If LenB(companyId) = 0 Then companyId = COMPANY_DEFAULT_FALLBACK
    projectRef = vbNullString

    headerMemo = ConcatFieldsWithSep(" | ", orgNameValue, courseNameValue, descriptionValue)
    If LenB(descriptionValue) > 0 Then
        lineDescription = courseNameValue & " - " & descriptionValue
    Else
        lineDescription = courseNameValue
    End If
    lineMemo = emailValue

    Dim invoiceGroupKey As String
    Dim nextSeq As Long
    If LenB(poNumberValue) > 0 Then
        invoiceGroupKey = "PO::" & poNumberValue
    Else
        invoiceGroupKey = "CUST::" & sourceCustomerId & "::" & invoiceDate
    End If
    If CollectionKeyExists(invoiceGroupCounters, invoiceGroupKey) Then
        nextSeq = CollectionGetItemLong(invoiceGroupCounters, invoiceGroupKey)
    Else
        nextSeq = 1
    End If
    CollectionSetItemLong invoiceGroupCounters, invoiceGroupKey, nextSeq + 1

    If LenB(poNumberValue) > 0 Then
        invoiceRefId = poNumberValue & "-" & CStr(nextSeq)
    Else
        If LenB(sourceCustomerId) > 0 Then
            invoiceRefId = sourceCustomerId & "-NOPO-" & CStr(nextSeq)
        Else
            invoiceRefId = "AR-INT-" & CStr(sequenceNumber)
        End If
    End If
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceRefId, mGrpNextKey, lineOrderValue)

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = "Y"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    ' Locked in Workday (I) is left blank per 2026-09-22 amendment (matches ar_discipline_invoice).
    ' Customer Invoice ID (G) = integration_{AJ without hyphens}_{BR as YYYYMMDD}_{B}
    ' (e.g. C-000190 + 2026-09-18 + key 1 -> integration_C000190_20260918_1).
    ' AJ keeps the original Customer ID. Customer Invoice Reference ID (D) is left
    ' blank: it looks up an EXISTING Workday invoice, so a new ID there fails the load.
    ' Invoice Number (BO) is left blank; the invoice_identifier (invoiceRefId) is used
    ' only as the line-group key above. See mapping.d / mapping.g / mapping.bo.
    referenceId = REFERENCE_ID_PREFIX & Replace(outputCustomerId, "-", vbNullString) & _
        "_" & Replace(invoiceDate, "-", vbNullString) & "_" & CStr(spreadsheetKeyValue)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REF_ID)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).Value = referenceId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_NUMBER)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = companyId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = outputCustomerId
    ' Sold To Customer (AK) is left blank per 2026-09-22 instruction (Workday marks it Optional).

    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = invoiceDate

    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    If LenB(poNumberValue) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_PO_NUMBER)).Value = poNumberValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INCLUDE_CUSTOMER_WORKTAGS)).Value = "Y"
    If LenB(headerMemo) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_MEMO)).Value = headerMemo

    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(lineOrderValue)
    ' Customer Invoice Line Reference ID (DU) is left blank per 2026-09-22 instruction.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = CStr(lineOrderValue)

    If LenB(revenueCategory) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = revenueCategory
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = lineDescription
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = qtyValue
    ' Unit of Measure (FI) is left blank per 2026-09-22 instruction.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).Value = grossAmount

    If LenB(orderDate) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).NumberFormat = "@"
        targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).Value = orderDate
    End If

    If LenB(lineMemo) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_MEMO)).Value = lineMemo

    If LenB(projectRef) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_BILLABLE_PROJECT_HEADER)).Value = projectRef
        targetSheet.Cells(targetRow, ColumnIndex(TGT_BILLABLE_PROJECT_LINE)).Value = projectRef
    End If

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

    If LenB(projectRef) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_PROJECT_LINE_WORKTAG)).Value = projectRef
    End If
End Sub

Private Function BuildCourseHierarchyLookup( _
    ByVal codesTable As Variant, _
    ByVal codesDataStartRowIndex As Long, _
    ByVal codesCourseColumn As Long, _
    ByVal codesCompanyColumn As Long, _
    ByVal codesCostCentreColumn As Long, _
    ByVal codesReportingTypeColumn As Long, _
    ByVal codesFaLearningColumn As Long, _
    ByVal codesCampColumn As Long, _
    ByVal codesCommercialPartnerColumn As Long, _
    ByVal codesEventColumn As Long, _
    ByVal codesGrantColumn As Long, _
    ByVal codesProgramColumn As Long, _
    ByVal codesRevenueCatColumn As Long _
) As Collection
    Dim lookup As Collection
    Dim rowIndex As Long
    Dim courseKey As String
    Dim packed As String

    Set lookup = New Collection

    If codesCourseColumn = 0 Then
        ' Code lists tab has no "Course" header (2026-09-22: "Combined 30-34" workbook's Code lists
        ' tab is a plain dropdown reference list, not a Course-keyed hierarchy) -- return an
        ' empty lookup so every row falls through to the unresolved-course / silent_blank path.
        Set BuildCourseHierarchyLookup = lookup
        Exit Function
    End If

    For rowIndex = codesDataStartRowIndex To UBound(codesTable, 1)
        courseKey = SafeCellStr(codesTable(rowIndex, codesCourseColumn))
        If LenB(courseKey) = 0 Then GoTo NextCodesRow

        packed = _
            SafeCellByCol(codesTable, rowIndex, codesCompanyColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesCostCentreColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesReportingTypeColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesFaLearningColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesCampColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesCommercialPartnerColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesEventColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesGrantColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesProgramColumn) & "|" & _
            SafeCellByCol(codesTable, rowIndex, codesRevenueCatColumn)

        On Error Resume Next
        lookup.Add Item:=packed, Key:=LCase$(courseKey)
        On Error GoTo 0
NextCodesRow:
    Next rowIndex

    Set BuildCourseHierarchyLookup = lookup
End Function

Private Function IntegrationRowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal customerIdColumn As Long, _
    ByVal netColumn As Long, _
    ByVal vatColumn As Long, _
    ByVal grossColumn As Long _
) As Boolean
    Dim netAmount As Double
    Dim vatAmount As Double
    Dim grossAmount As Double

    IntegrationRowIsSkippable = False

    If LenB(SafeCellStr(tableData(rowIndex, invoiceDateColumn))) = 0 Then
        IntegrationRowIsSkippable = True
        Exit Function
    End If

    If LenB(SafeCellStr(tableData(rowIndex, customerIdColumn))) = 0 Then
        IntegrationRowIsSkippable = True
        Exit Function
    End If

    If netColumn > 0 Then netAmount = SafeCellDbl(tableData(rowIndex, netColumn)) Else netAmount = 0#
    If vatColumn > 0 Then vatAmount = SafeCellDbl(tableData(rowIndex, vatColumn)) Else vatAmount = 0#
    If grossColumn > 0 Then grossAmount = SafeCellDbl(tableData(rowIndex, grossColumn)) Else grossAmount = 0#

    If netAmount = 0# And vatAmount = 0# And grossAmount = 0# Then
        IntegrationRowIsSkippable = True
        Exit Function
    End If
End Function

Private Function CleanMemoToSingleLine(ByVal rawText As String) As String
    Dim temp As String
    temp = Replace$(rawText, vbCrLf, " ")
    temp = Replace$(temp, vbCr, " ")
    temp = Replace$(temp, vbLf, " ")
    temp = Replace$(temp, vbTab, " ")
    temp = Replace$(temp, "  ", " ")
    Do While InStr(temp, "  ") > 0
        temp = Replace$(temp, "  ", " ")
    Loop
    CleanMemoToSingleLine = Trim$(temp)
End Function

Private Function ConcatFieldsWithSep(ByVal separator As String, ParamArray fields() As Variant) As String
    Dim i As Long
    Dim result As String
    Dim piece As String

    result = vbNullString
    For i = LBound(fields) To UBound(fields)
        piece = Trim$(CStr(fields(i)))
        If LenB(piece) > 0 Then
            If LenB(result) = 0 Then
                result = piece
            Else
                result = result & separator & piece
            End If
        End If
    Next i
    ConcatFieldsWithSep = result
End Function

' Keyed-Collection primitives (CollectionKeyExists / *SetItemLong / *GetItemLong /
' *SetItemDouble / *GetItemDouble) are shared in modHelpers -- no per-module copy (2026-09-14).


' Guards a table read against columnIndex = 0 (header not found -- FindOptionalColumnIndexInRow
' returns 0), which would otherwise be an invalid subscript into the 1-based table array.
Private Function SafeCellByCol(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If columnIndex = 0 Then
        SafeCellByCol = vbNullString
    Else
        SafeCellByCol = SafeCellStr(tableData(rowIndex, columnIndex))
    End If
End Function

Private Function StripLeadingNumberPrefix(ByVal rawText As String) As String
    ' "<number>: <text>" -> "<text>". Left unchanged when there is no ":" or the part
    ' before the first ":" is not all digits.
    Dim colonPos As Long
    Dim prefixText As String

    StripLeadingNumberPrefix = Trim$(rawText)
    colonPos = InStr(1, StripLeadingNumberPrefix, ":")
    If colonPos <= 1 Then Exit Function
    prefixText = Trim$(Left$(StripLeadingNumberPrefix, colonPos - 1))
    If LenB(prefixText) = 0 Then Exit Function
    If prefixText Like "*[!0-9]*" Then Exit Function
    StripLeadingNumberPrefix = Trim$(Mid$(StripLeadingNumberPrefix, colonPos + 1))
End Function

Private Function SafeCellStr(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        SafeCellStr = vbNullString
    ElseIf IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeCellStr = vbNullString
    ElseIf VarType(cellValue) = vbDate Then
            SafeCellStr = CStr(cellValue)
    Else
        SafeCellStr = Trim$(CStr(cellValue))
    End If
End Function

Private Function SafeCellDbl(ByVal cellValue As Variant) As Double
    If IsError(cellValue) Then
        SafeCellDbl = 0#
    ElseIf IsEmpty(cellValue) Or IsNull(cellValue) Then
        SafeCellDbl = 0#
    ElseIf IsNumeric(cellValue) Then
        SafeCellDbl = CDbl(cellValue)
    Else
        Dim s As String
        s = Trim$(CStr(cellValue))
        If LenB(s) = 0 Then
            SafeCellDbl = 0#
        Else
            s = Replace$(s, ",", vbNullString)
            s = Replace$(s, ChrW(163), vbNullString)
            s = Replace$(s, " ", vbNullString)
            If LenB(s) = 0 Then
                SafeCellDbl = 0#
            ElseIf IsNumeric(s) Then
                SafeCellDbl = CDbl(s)
            Else
                SafeCellDbl = 0#
            End If
        End If
    End If
End Function
