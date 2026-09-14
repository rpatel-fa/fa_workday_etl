Attribute VB_Name = "modArIntegrationCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - NFC"

'  --- Profile-local Tax Triple overrides (hybrid scope; shadows modFinanceCommon.bas shared defaults) ---
'  These values come from ar_integration_invoice.yaml transform_constants.tax_triple_per_row and MUST
'  NOT be replaced with the shared Public Const TAX_APPLICABILITY_ZERO / TAX_CODE_ZERO from
'  modFinanceCommon.  Shared constants use legacy GBR_Outside_of_Scope / GBR_Zero_Rated which fail
'  Workday validation.  AR Integration uses the domestic-specific 200_VAT / 000_No_VAT ref-IDs.
Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Sale_of_Goods_and_Services"
Private Const TAX_APPLICABILITY_ZERO As String = "GBR_Domestic_Outside_the_Scope"
Private Const TAX_CODE_STANDARD As String = "GBR_200_VAT"
Private Const TAX_CODE_ZERO As String = "GBR_000_No_VAT"

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

'  --- Codes-tab return value packing order inside courseHierarchyLookup pipe-delimited string ---
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

Public Sub RunArIntegrationCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim customerSpreadsheetPath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim selectionKey As String
    Dim customerWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim customerLookup As Collection
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
    Dim orgNameColumn As Long
    Dim courseNameColumn As Long
    Dim courseCodeColumn As Long
    Dim orderDateColumn As Long
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
    customerSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_CUSTOMER_SPREADSHEET)
    mappingSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_MAPPING_SPREADSHEET)
    mappingSheetName = ReadControlValue(controlSheet, CONTROL_LABEL_MAPPING_SHEET)
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If LenB(mappingSheetName) = 0 Then mappingSheetName = DEFAULT_MAPPING_SHEET_NAME

    Set customerWorkbook = OpenWorkbookReadOnly(customerSpreadsheetPath)
    Set customerLookup = BuildReferenceLookupFromWorkbook(customerWorkbook, "Customer")
    customerWorkbook.Close SaveChanges:=False
    Set customerWorkbook = Nothing

    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modArIntegrationCustInvoice", 4)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)

    Set resolvedIntegrationSheet = ResolveSourceWorksheet(sourceWorkbook, "Integration")
    resolvedIntegrationSheetName = resolvedIntegrationSheet.Name
    integrationTable = ReadWorksheetTable(resolvedIntegrationSheet)

    Set resolvedCodesSheet = ResolveSourceWorksheet(sourceWorkbook, "Codes")
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
    orgNameColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Organisation Name")
    courseNameColumn = FindColumnIndexInRow(integrationTable, headerRowIndex, "Course Name")
    courseCodeColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Course Code")
    orderDateColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Order Date")
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
    emailColumn = FindOptionalColumnIndexInRow(integrationTable, headerRowIndex, "Email Address")

    codesHeaderRowIndex = 1
    codesDataStartRowIndex = 2
    codesCourseColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Course")
    codesCompanyColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Company")
    codesCostCentreColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Cost centre")
    codesReportingTypeColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Reporting Type")
    codesFaLearningColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "FA Learning courses")
    codesCampColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Camp")
    codesCommercialPartnerColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Commerical Partners")
    codesEventColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Events")
    codesGrantColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Grant")
    codesProgramColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Program")
    codesRevenueCatColumn = FindColumnIndexInRow(codesTable, codesHeaderRowIndex, "Revenue Cat")

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
                workdayCustomerId = LookupCollectionValue(customerLookup, sourceCustomerId)
            If LenB(workdayCustomerId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unresolvedCustomerRows = unresolvedCustomerRows + 1
                IncrementCount unresolvedCounts, unresolvedKeys, sourceCustomerId
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "CUSTOMER ID not found in customer extract", _
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
                    invoiceDateColumn, orgNameColumn, courseNameColumn, courseCodeColumn, orderDateColumn, _
                    qtyColumn, netColumn, vatColumn, grossColumn, poNumberColumn, _
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
        issues.Add BuildRejectedLookupIssue(unresolvedCounts, unresolvedKeys, "CUSTOMER ID")
    End If
    If unresolvedCourseRows > 0 Then
        issues.Add "Course Name did not resolve in Codes tab for " & CStr(unresolvedCourseRows) & _
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
    If Not customerWorkbook Is Nothing Then customerWorkbook.Close SaveChanges:=False
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
    ByVal courseCodeColumn As Long, ByVal orderDateColumn As Long, _
    ByVal qtyColumn As Long, ByVal netColumn As Long, ByVal vatColumn As Long, _
    ByVal grossColumn As Long, ByVal poNumberColumn As Long, _
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
    Dim netAmount As Double
    Dim qtyValue As Double
    Dim vatAmount As Double
    Dim vatRatio As Double
    Dim courseNameValue As String
    Dim courseCodeValue As String
    Dim orgNameValue As String
    Dim taxApplicability As String
    Dim taxCode As String
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
    Dim worktagIndex As Long
    Dim worktagColumns As Variant
    Dim worktagValue As Variant
    Dim foundCourse As Boolean

    invoiceDate = NormalizeSourceDate(SafeCellStr(sourceTable(sourceRow, invoiceDateColumn)))
    If orderDateColumn > 0 Then orderDate = NormalizeSourceDate(SafeCellStr(sourceTable(sourceRow, orderDateColumn))) Else orderDate = vbNullString

    If qtyColumn > 0 Then qtyValue = SafeCellDbl(sourceTable(sourceRow, qtyColumn)) Else qtyValue = 1#
    If qtyValue <= 0# Then qtyValue = 1#

    If netColumn > 0 Then netAmount = SafeCellDbl(sourceTable(sourceRow, netColumn)) Else netAmount = 0#
    If vatColumn > 0 Then vatAmount = SafeCellDbl(sourceTable(sourceRow, vatColumn)) Else vatAmount = 0#
    If grossColumn > 0 Then
        grossAmount = SafeCellDbl(sourceTable(sourceRow, grossColumn))
    Else
        grossAmount = netAmount + vatAmount
    End If

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

    courseNameValue = SafeCellStr(sourceTable(sourceRow, courseNameColumn))
    If courseCodeColumn > 0 Then courseCodeValue = SafeCellStr(sourceTable(sourceRow, courseCodeColumn)) Else courseCodeValue = vbNullString
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
            costCenterRef = Trim$(CStr(courseParts(COURSE_PACK_COST_CENTRE - 1)))
            reportingType = Trim$(CStr(courseParts(COURSE_PACK_REPORTING_TYPE - 1)))
            faLearningRef = Trim$(CStr(courseParts(COURSE_PACK_FA_LEARNING - 1)))
            campRef = Trim$(CStr(courseParts(COURSE_PACK_CAMP - 1)))
            commercialPartnerRef = Trim$(CStr(courseParts(COURSE_PACK_COMMERCIAL_PARTNER - 1)))
            eventRef = Trim$(CStr(courseParts(COURSE_PACK_EVENT - 1)))
            grantRef = Trim$(CStr(courseParts(COURSE_PACK_GRANT - 1)))
            programRef = Trim$(CStr(courseParts(COURSE_PACK_PROGRAMME - 1)))
            revenueCategory = Trim$(CStr(courseParts(COURSE_PACK_REVENUE_CATEGORY - 1)))
        End If
    Else
        unresolvedCourseRows = unresolvedCourseRows + 1
        If LenB(courseNameValue) > 0 Then
            IncrementCount unresolvedCounts, unresolvedKeys, courseNameValue
        End If
    End If

    If LenB(companyId) = 0 Then companyId = COMPANY_DEFAULT_FALLBACK
    projectRef = vbNullString

    headerMemo = ConcatFieldsWithSep(" | ", orgNameValue, courseNameValue, courseCodeValue)
    If LenB(courseCodeValue) > 0 Then
        lineDescription = courseNameValue & " - " & courseCodeValue
    Else
        lineDescription = courseNameValue
    End If
    lineMemo = courseCodeValue

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

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = sequenceNumber
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = "Y"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LOCKED_IN_WORKDAY)).Value = OUTPUT_AUTO_COMPLETE
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REF_ID)).Value = invoiceRefId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).Value = invoiceRefId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_NUMBER)).Value = invoiceRefId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = companyId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = outputCustomerId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SOLD_TO_CUSTOMER)).Value = outputCustomerId

    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = invoiceDate

    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    If LenB(poNumberValue) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_PO_NUMBER)).Value = poNumberValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INCLUDE_CUSTOMER_WORKTAGS)).Value = "Y"
    If LenB(headerMemo) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_MEMO)).Value = headerMemo

    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(1)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_REFERENCE)).Value = "LINE-" & CStr(sequenceNumber) & "-1"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = "1"

    If LenB(revenueCategory) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = revenueCategory
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = lineDescription
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = taxApplicability
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = taxCode
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = qtyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_OF_MEASURE)).Value = "Each"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = netAmount
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

    For rowIndex = codesDataStartRowIndex To UBound(codesTable, 1)
        courseKey = SafeCellStr(codesTable(rowIndex, codesCourseColumn))
        If LenB(courseKey) = 0 Then GoTo NextCodesRow

        packed = _
            SafeCellStr(codesTable(rowIndex, codesCompanyColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesCostCentreColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesReportingTypeColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesFaLearningColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesCampColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesCommercialPartnerColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesEventColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesGrantColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesProgramColumn)) & "|" & _
            SafeCellStr(codesTable(rowIndex, codesRevenueCatColumn))

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
