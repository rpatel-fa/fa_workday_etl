Attribute VB_Name = "modArGbeInvoiceCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  DECIDED 2026-09-17: the source no longer carries separate Player/Reference columns; both
'  are embedded in one free-text cell ("Invoice Line Description", e.g.
'  "Romain Jordan Faivre-FOOT01/22/S-2024-1669"). Verified against every row of the 2026-09-17
'  AR file (435 Mens + 149 Womens, 100% match, incl. hyphenated surnames like
'  "Peyraud-Magnin"): the literal "-FOOT" token always marks the Player/Reference boundary and
'  never occurs more than once. Mirrors the inline split-on-anchor pattern already used for
'  ap_travel_places' SupplierID (InStr + Left$/Mid$, no shared helper - see
'  modApTravelPlacesSupInv.bas).
Private Const PLAYER_REFERENCE_SPLIT_ANCHOR As String = "-FOOT"

'  Customer Invoice ID (G) = gbe_{mens|womens}_{AJ without hyphens}_{BR as YYYYMMDD}_{B}
'  (DECIDED 2026-09-28, see mapping.g). D is intentionally left blank because Workday uses
'  it as an EXISTING-invoice lookup and rejects a new ID there. The mens/womens token is
'  chosen from the Runner variant profile_key; any other selection key fails the run rather
'  than guessing.
Private Const REFERENCE_ID_PREFIX As String = "gbe_"
Private Const VARIANT_KEY_MENS As String = "ar_gbe_invoice_mens"
Private Const VARIANT_KEY_WOMENS As String = "ar_gbe_invoice_womens"
Private Const VARIANT_TOKEN_MENS As String = "mens"
Private Const VARIANT_TOKEN_WOMENS As String = "womens"

'  --- Profile-local Tax Triple overrides (hybrid scope; shadows modFinanceCommon.bas shared defaults) ---
'  These values come from ar_gbe_invoice.yaml transform_constants.tax_triple_per_row and
'  mirror ar_integration_invoice (modArIntegrationCustInvoice.bas); keep them profile-local so a
'  change to the shared modFinanceCommon defaults does not alter this transform. Set 2026-10-07.
Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Sale_of_Goods_and_Services"
Private Const TAX_APPLICABILITY_ZERO As String = "GBR_Outside_of_Scope"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_CODE_ZERO As String = "GBR_Zero_Rated"

'  -- source layout: header=1, data=2; worksheet is named per Runner variant --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution /
'  Header Row / Data Start Row). ar_gbe_invoice ships TWO variant selections that share this
'  ONE macro and differ only by Source Worksheet Name ("Mens GBE" / "Womens GBE") + paths.

Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"

Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"
'  GBE has NO per-row COST CODE column, so the Mapping-MASTER worktag lookups below never
'  resolve (costCodeColumn is 0 -> the WriteOutputRow guards skip every lookup) for event /
'  camp / FA Learning / commercial partner / grant / programme / project. Cost Centre (FX),
'  Reporting Type (FY) and Revenue Category (EM) are handled separately as BATCH CONSTANTS
'  (see ResolveBatchConstantColumnValue) - see ar_gbe_invoice.yaml transform_constants.batch_constant_columns.
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
'  1:1 mirror of ar_gbe_invoice.yaml customer_invoice_eib_columns. If any YAML letter changes,
'  the hygiene check fails and this block must be updated.
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

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long
'  Header values (CQ / DB) captured from the FIRST row of each invoice group, keyed by
'  Customer ID (AJ) + "|" + Invoice Date (BR), so every line of a grouped invoice carries one
'  consistent header.
Private mGrpHdrPo As Collection
Private mGrpHdrMemo As Collection

Public Sub RunArGbeInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim mappingWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceTable As Variant
    Dim mappingTable As Variant
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
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unresolvedCustomerRows As Long
    Dim dateColumn As Long, clubColumn As Long, invoiceLineDescColumn As Long, typeColumn As Long
    Dim grossColumn As Long, netColumn As Long
    Dim salesItemColumn As Long, qtyColumn As Long, costCodeColumn As Long
    Dim customerIdColumn As Long, companyColumn As Long
    Dim costCentreColumn As Long, reportingTypeColumn As Long, revenueCategoryColumn As Long
    Dim poNumberColumn As Long
    Dim clubName As String
    Dim workdayCustomerId As String
    Dim playerValue As String
    Dim referenceValue As String
    Dim variantToken As String
    Dim missingReferenceRows As Long
    Dim gbeCompanyValue As String
    Dim gbeCostCentreValue As String
    Dim gbeReportingTypeValue As String
    Dim gbeRevenueCategoryValue As String
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
    Dim resolvedSourceSheet As Worksheet
    Dim resolvedSourceSheetName As String
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
    Set unresolvedCounts = New Collection
    Set unresolvedKeys = New Collection

    On Error GoTo HandleFailure
    variantToken = ResolveGbeVariantToken(selectionKey)

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
    If StrComp(worksheetResolutionStr, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetNameStr
    End If
    EnsureNonBlank mappingSpreadsheetPath, CONTROL_LABEL_MAPPING_SPREADSHEET
    If LenB(mappingSheetName) = 0 Then mappingSheetName = DEFAULT_MAPPING_SHEET_NAME

    Set mappingWorkbook = OpenWorkbookReadOnly(mappingSpreadsheetPath)
    mappingTable = ReadWorksheetTable(ResolveSourceWorksheet(mappingWorkbook, mappingSheetName))
    mappingWorkbook.Close SaveChanges:=False
    Set mappingWorkbook = Nothing

    '  Cost Centre (FX), Reporting Type (FY) and Revenue Category (EM) no longer come from
    '  this Mapping-MASTER hierarchy lookup (see ResolveBatchConstantColumnValue below); the
    '  remaining 6 worktag slots stay wired to it for archetype parity, though it never
    '  resolves for GBE (no per-row COST CODE column).
    Set eventLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_EVENT)
    Set campLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_CAMP)
    Set faLearningLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_FA_LEARNING)
    Set commercialPartnerLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_COMMERCIAL_PARTNER)
    Set grantLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_GRANT)
    Set programLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_PROGRAM)
    Set projectLookup = BuildBusinessObjectHierarchyLookupMulti(mappingTable, WBO_PROJECT, WBO_PROJECT_CAPEX)

    '  Source of truth: config/transforms/ar_gbe_invoice.yaml source_layout.required_headers,
    '  written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modArGbeInvoiceCustInvoice", 11)

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
        Err.Raise vbObjectError + 1920, "modArGbeInvoiceCustInvoice.RunArGbeInvoiceCustomerInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issues.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    dateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Date")
    clubColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Club|Customer Name")
    '  DECIDED 2026-09-17: the source no longer has separate Player/Reference columns; both are
    '  embedded in one free-text cell - split at runtime, see SplitPlayerAndReference below.
    invoiceLineDescColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Line Description")
    typeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice line part 2|Invoice Line Description 2")
    grossColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Gross")
    netColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net")
    '  DECIDED 2026-09-17: AR now delivers a ready-made Workday Customer ID per row (col A;
    '  header "WORKDAY ID" on the Mens tab / "Customer ID" on the Womens tab) - direct
    '  passthrough to AJ, see the customer-id read in the row loop below.
    customerIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "WORKDAY ID|Customer ID")
    '  DECIDED 2026-09-17: Company / Cost Centre / Reporting Type / Revenue Category are BATCH
    '  CONSTANTS - only the first data row carries a value; ResolveBatchConstantColumnValue below
    '  reads it once and WriteOutputRow applies it to every row.
    companyColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Company")
    costCentreColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Cost Centre")
    reportingTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Reporting Type")
    revenueCategoryColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Revenue Category")
    '  Optional columns: GBE carries no COST CODE / SALES ITEM / QUANTITY (kept optional -> 0 ->
    '  inert). The VAT column is intentionally NOT read: its header is ABSENT on the Womens tab, so
    '  the VAT amount for Line Tax Applicability (FA) / Line Tax Code (FB) is derived as
    '  Gross - Net (present on both tabs) in WriteOutputRow instead.
    salesItemColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "SALES ITEM CODE")
    qtyColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "QUANTITY")
    costCodeColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "COST CODE")
    '  Customer PO Number (CQ) = per-row direct passthrough of source col I "PO" (2026-10-22).
    '  Optional (not in required_headers): a missing column or blank cell leaves CQ empty.
    poNumberColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "PO")

    gbeCompanyValue = ResolveBatchConstantColumnValue(sourceTable, dataStartRowIndex, companyColumn)
    If LenB(gbeCompanyValue) = 0 Then gbeCompanyValue = COMPANY_DEFAULT_FALLBACK
    gbeCostCentreValue = ResolveBatchConstantColumnValue(sourceTable, dataStartRowIndex, costCentreColumn)
    gbeReportingTypeValue = ResolveBatchConstantColumnValue(sourceTable, dataStartRowIndex, reportingTypeColumn)
    gbeRevenueCategoryValue = ResolveBatchConstantColumnValue(sourceTable, dataStartRowIndex, revenueCategoryColumn)

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "MV"
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0
    Set mGrpHdrPo = New Collection
    Set mGrpHdrMemo = New Collection

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, dateColumn, clubColumn, grossColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Date / Club / Gross)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, clubColumn)))
        Else
            clubName = Trim$(CStr(sourceTable(rowIndex, clubColumn)))
            SplitPlayerAndReference CStr(sourceTable(rowIndex, invoiceLineDescColumn)), playerValue, referenceValue
            '  DECIDED 2026-09-17: direct passthrough of the AR-supplied Workday Customer ID
            '  (col A). No lookup, no name matching - a blank ID rejects the row.
            workdayCustomerId = Trim$(CStr(sourceTable(rowIndex, customerIdColumn)))
            If LenB(referenceValue) = 0 Then
                '  Reference drives the header Memo (DB) and the Player/Type line description,
                '  so a row whose Invoice Line Description does not contain the "-FOOT" anchor
                '  is rejected rather than written with a guessed split.
                rowsSkipped = rowsSkipped + 1
                missingReferenceRows = missingReferenceRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Reference not found in Invoice Line Description (no '" & PLAYER_REFERENCE_SPLIT_ANCHOR & "' token)", _
                    clubName, _
                    Trim$(CStr(sourceTable(rowIndex, invoiceLineDescColumn)))
            ElseIf LenB(workdayCustomerId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unresolvedCustomerRows = unresolvedCustomerRows + 1
                IncrementCount unresolvedCounts, unresolvedKeys, clubName
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Workday Customer ID missing in source column A", _
                    clubName, _
                    referenceValue
            Else
                rowsProcessed = rowsProcessed + 1
                WriteOutputRow runnerWorkbook, targetSheet, targetRow, rowsProcessed, sourceTable, rowIndex, _
                    dateColumn, playerValue, referenceValue, typeColumn, salesItemColumn, qtyColumn, _
                    netColumn, grossColumn, costCodeColumn, poNumberColumn, _
                    workdayCustomerId, gbeCompanyValue, gbeCostCentreValue, _
                    gbeReportingTypeValue, gbeRevenueCategoryValue, _
                    eventLookup, campLookup, faLearningLookup, commercialPartnerLookup, _
                    grantLookup, programLookup, projectLookup, variantToken
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Date / Club / Gross."
    End If
    If missingReferenceRows > 0 Then
        issues.Add "Rejected " & CStr(missingReferenceRows) & " row(s) whose Invoice Line Description has no '" & PLAYER_REFERENCE_SPLIT_ANCHOR & "' token."
    End If
    If unresolvedKeys.Count > 0 Then
        issues.Add BuildRejectedLookupIssue(unresolvedCounts, unresolvedKeys, "Club")
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
    Err.Raise failureNumber, "modArGbeInvoiceCustInvoice.RunArGbeInvoiceCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputRow( _
    ByVal runnerWorkbook As Workbook, ByVal targetSheet As Worksheet, ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, ByVal sourceTable As Variant, ByVal sourceRow As Long, _
    ByVal dateColumn As Long, ByVal playerValue As String, ByVal referenceValue As String, ByVal typeColumn As Long, _
    ByVal salesItemColumn As Long, ByVal qtyColumn As Long, ByVal netColumn As Long, _
    ByVal grossColumn As Long, ByVal costCodeColumn As Long, ByVal poNumberColumn As Long, _
    ByVal outputCustomerId As String, ByVal companyValue As String, ByVal costCentreValue As String, _
    ByVal reportingTypeValue As String, ByVal revenueCategoryValue As String, _
    ByVal eventLookup As Collection, ByVal campLookup As Collection, ByVal faLearningLookup As Collection, _
    ByVal commercialPartnerLookup As Collection, ByVal grantLookup As Collection, _
    ByVal programLookup As Collection, ByVal projectLookup As Collection, _
    ByVal variantToken As String _
)
    Dim companyId As String
    Dim invoiceDate As String
    Dim grossAmount As Double
    Dim netAmount As Double
    Dim vatAmount As Double
    Dim vatRatio As Double
    Dim taxApplicability As String
    Dim taxCode As String
    Dim qtyValue As Double
    Dim costCodeValue As String
    Dim revenueCategory As String
    Dim costCenterRef As String
    Dim salesItemValue As String
    Dim lineDescValue As String
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
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long
    Dim referenceId As String
    Dim poNumberValue As String
    Dim invoiceGroupKey As String
    Dim headerInvoiceDate As String
    Dim headerPoNumber As String
    Dim headerMemo As String

    companyId = companyValue
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, dateColumn))
    grossAmount = ResolveAmount(sourceTable(sourceRow, grossColumn))

    If qtyColumn > 0 Then qtyValue = ResolveAmount(sourceTable(sourceRow, qtyColumn)) Else qtyValue = 1#
    If netColumn > 0 Then netAmount = ResolveAmount(sourceTable(sourceRow, netColumn)) Else netAmount = grossAmount

    '  Line Tax Applicability (FA) / Line Tax Code (FB): per-row tax triple from the Net/VAT
    '  ratio, same rule and values as ar_integration_invoice (restored 2026-10-07; see
    '  transform_specs tax_rules). VAT = Gross - Net because the Womens tab header omits the
    '  VAT label, so a VAT column read cannot be relied on.
    vatAmount = grossAmount - netAmount
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

    '  Invoice-line grouping keyed by Customer ID (AJ) + Invoice Date (BR): every row for the
    '  same customer on the same date shares one Spreadsheet Key (B), and Row ID (DS) / Line
    '  Order (DV) run 1..N within it. BR (and so G) is uniform within a group by construction;
    '  CQ / DB come from the group's FIRST row so the grouped invoice has one consistent
    '  header. See transform_specs mapping.b / .ds / .dv.
    invoiceGroupKey = outputCustomerId & "|" & invoiceDate
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceGroupKey, mGrpNextKey, lineOrderValue)
    If poNumberColumn > 0 Then poNumberValue = Trim$(CStr(sourceTable(sourceRow, poNumberColumn))) Else poNumberValue = vbNullString
    If Not CollectionKeyExists(mGrpHdrPo, invoiceGroupKey) Then
        mGrpHdrPo.Add poNumberValue, invoiceGroupKey
        mGrpHdrMemo.Add referenceValue, invoiceGroupKey
    End If
    headerInvoiceDate = invoiceDate
    headerPoNumber = CStr(mGrpHdrPo(invoiceGroupKey))
    headerMemo = CStr(mGrpHdrMemo(invoiceGroupKey))
    lineDescValue = BuildLineDescription(playerValue, CStr(sourceTable(sourceRow, typeColumn)))
    If salesItemColumn > 0 Then salesItemValue = CStr(sourceTable(sourceRow, salesItemColumn)) Else salesItemValue = vbNullString

    '  DECIDED 2026-09-17: Revenue Category (EM) and Cost Centre (FX) are BATCH CONSTANTS
    '  passed in from the caller (read once from source cols N/L row 2) - no longer derived
    '  via the Mapping-MASTER cost-code hierarchy lookup.
    revenueCategory = revenueCategoryValue
    costCenterRef = costCentreValue

    '  GBE still has no per-row COST CODE column, so costCodeValue is empty and the remaining
    '  Mapping-MASTER worktag lookups below stay guard-skipped (event/camp/FA Learning/
    '  commercial partner/grant/programme/project stay blank; Finance codes pre-upload).
    If costCodeColumn > 0 Then costCodeValue = Trim$(CStr(sourceTable(sourceRow, costCodeColumn))) Else costCodeValue = vbNullString
    '  ResolveReportingTypeFromFinalCode() is always a no-op here (costCodeValue is always
    '  blank - GBE has no per-row COST CODE); kept as a formal call for archetype parity with
    '  the derived_worktags.reporting_type contract. Reporting Type (FY) actually written is
    '  the BATCH CONSTANT from source col M row 2 (DECIDED 2026-09-17), applied below.
    reportingType = ResolveReportingTypeFromFinalCode(costCodeValue)
    reportingType = reportingTypeValue
    If LenB(costCodeValue) > 0 Then
        eventRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(eventLookup, costCodeValue)
        campRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(campLookup, costCodeValue)
        faLearningRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(faLearningLookup, costCodeValue)
        commercialPartnerRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(commercialPartnerLookup, costCodeValue)
        grantRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(grantLookup, costCodeValue)
        programRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(programLookup, costCodeValue)
        projectRef = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(projectLookup, costCodeValue)
    End If

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_COMMENT)).Value = OUTPUT_AUTO_COMPLETE
    '  Customer Invoice ID (G) = gbe_{mens|womens}_{AJ without hyphens}_{BR as YYYYMMDD}_{B}
    '  (e.g. C-000102 + 2026-03-17 + key 1 -> gbe_mens_C000102_20260317_1).
    '  AJ keeps the original Customer ID. D is intentionally left blank; see mapping.g / mapping.d.
    referenceId = REFERENCE_ID_PREFIX & variantToken & "_" & Replace(outputCustomerId, "-", vbNullString) & _
        "_" & Replace(headerInvoiceDate, "-", vbNullString) & "_" & CStr(spreadsheetKeyValue)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).Value = referenceId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REF_ID)).Value = vbNullString
    '  Invoice Number (BO) intentionally left blank on every row per requester instruction
    '  (2026-09-28). See mapping.bo.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_NUMBER)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LOCKED_IN_WORKDAY)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = companyId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = outputCustomerId

    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = headerInvoiceDate

    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    '  Amount Due (CM) left blank on every row (2026-10-07), as ar_integration_invoice /
    '  ar_discipline_invoice. Was source col F (NET). See mapping.cm.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    '  Customer PO Number (CQ) = source col I "PO" verbatim, written as text so a numeric PO
    '  keeps any leading zeros; blank when the cell or column is absent. Taken from the
    '  customer group's first row. See mapping.cq.
    If LenB(headerPoNumber) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_PO_NUMBER)).NumberFormat = "@"
        targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_PO_NUMBER)).Value = headerPoNumber
    End If
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_MEMO)).Value = headerMemo
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(lineOrderValue)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_REFERENCE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = CStr(lineOrderValue)

    If LenB(salesItemValue) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_SALES_ITEM)).Value = salesItemValue
    If LenB(revenueCategory) > 0 Then targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = revenueCategory

    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = lineDescValue
    '  Line Tax Applicability (FA) / Line Tax Code (FB) per row, mirroring ar_integration_invoice
    '  (restored 2026-10-07). See mapping.fa / mapping.fb.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = taxApplicability
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = taxCode
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = qtyValue
    '  Unit Cost (FL) = source col F (NET), not GROSS (DECIDED 2026-09-17): Workday adds VAT
    '  from the line Tax Code (FB). Extended Amount (FO) left blank on every row (2026-10-07),
    '  as ar_integration_invoice / ar_discipline_invoice. See mapping.fl / mapping.fo.
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = netAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).Value = vbNullString

    If LenB(invoiceDate) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).NumberFormat = "@"
        targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).Value = invoiceDate
    End If

    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_MEMO)).Value = lineDescValue

    If LenB(projectRef) > 0 Then
        targetSheet.Cells(targetRow, ColumnIndex(TGT_BILLABLE_PROJECT_HEADER)).Value = projectRef
        targetSheet.Cells(targetRow, ColumnIndex(TGT_BILLABLE_PROJECT_LINE)).Value = projectRef
    End If

    '  9-slot worktags in the canonical Submit_Customer_Invoice_v46.1 columns FX FY FZ GA GB GC GD GE GF.
    '  Order matches ar_gbe_invoice.yaml worktags_plus_order. Reporting Type (FY) and Cost Centre
    '  (FX) now carry the batch-constant values passed in; the other 6 slots stay blank for GBE
    '  (no per-row COST CODE).
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

'  mens / womens token for Customer Invoice Reference ID (D), from the Runner variant
'  profile_key. Fails the run on any other key rather than guessing.
Private Function ResolveGbeVariantToken(ByVal selectionKey As String) As String
    Select Case LCase$(Trim$(selectionKey))
        Case VARIANT_KEY_MENS
            ResolveGbeVariantToken = VARIANT_TOKEN_MENS
        Case VARIANT_KEY_WOMENS
            ResolveGbeVariantToken = VARIANT_TOKEN_WOMENS
        Case Else
            Err.Raise vbObjectError + 1701, "modArGbeInvoiceCustInvoice.ResolveGbeVariantToken", _
                "Cannot tell Mens from Womens for selection key '" & selectionKey & "'. Expected " & _
                VARIANT_KEY_MENS & " or " & VARIANT_KEY_WOMENS & "."
    End Select
End Function

'  Line Item Description (EN) + line Memo (FV) = Player + " - " + Type
'  (e.g. "Boubacar Kamara - Extension"). Falls back to whichever part is present.
Private Function BuildLineDescription( _
    ByVal rawPlayerValue As String, _
    ByVal rawTypeValue As String _
) As String
    Dim playerValue As String
    Dim typeValue As String

    playerValue = CleanMemoToSingleLine(rawPlayerValue)
    typeValue = CleanMemoToSingleLine(rawTypeValue)

    If LenB(playerValue) > 0 And LenB(typeValue) > 0 Then
        BuildLineDescription = playerValue & " - " & typeValue
    ElseIf LenB(playerValue) > 0 Then
        BuildLineDescription = playerValue
    Else
        BuildLineDescription = typeValue
    End If
End Function

'  DECIDED 2026-09-17: split the combined "Invoice Line Description" cell on the literal
'  "-FOOT" anchor token - everything before it is the Player name, "FOOT..." onward is the
'  Reference. Verified against every row of the 2026-09-17 AR file (see
'  PLAYER_REFERENCE_SPLIT_ANCHOR declaration above for the verification note). outReference is
'  "" when the anchor is absent, which the caller treats as a reject (Reference is load-bearing
'  for DB and the line description - never guessed).
Private Sub SplitPlayerAndReference( _
    ByVal rawDescription As String, _
    ByRef outPlayer As String, _
    ByRef outReference As String _
)
    Dim cleaned As String
    Dim anchorPos As Long

    cleaned = CleanMemoToSingleLine(rawDescription)
    anchorPos = InStr(cleaned, PLAYER_REFERENCE_SPLIT_ANCHOR)

    If anchorPos > 0 Then
        outPlayer = Trim$(Left$(cleaned, anchorPos - 1))
        outReference = Trim$(Mid$(cleaned, anchorPos + 1))
    Else
        outPlayer = cleaned
        outReference = vbNullString
    End If
End Sub

'  Collapse a free-text source cell (Invoice Line Description / Type) to a single line for
'  EN/FV. Also strips the literal "_x000D_" token that some GBE rows carry as a baked-in
'  mis-encoded carriage-return marker (observed in the 2026-09-17 AR file, e.g. player names
'  and one "ESC_x000D_\n(New)" Type value) - a text-hygiene fix, not a business-logic change.
Private Function CleanMemoToSingleLine(ByVal rawText As String) As String
    Dim temp As String
    temp = Replace$(rawText, "_x000D_", " ")
    temp = Replace$(temp, vbCrLf, " ")
    temp = Replace$(temp, vbCr, " ")
    temp = Replace$(temp, vbLf, " ")
    temp = Replace$(temp, vbTab, " ")
    temp = Replace$(temp, "  ", " ")
    Do While InStr(temp, "  ") > 0
        temp = Replace$(temp, "  ", " ")
    Loop
    CleanMemoToSingleLine = Trim$(temp)
End Function

Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal dateColumn As Long, _
    ByVal clubColumn As Long, _
    ByVal grossColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, dateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, clubColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, grossColumn)) = 0#
End Function

'  DECIDED 2026-09-17: Company / Cost Centre / Reporting Type / Revenue Category arrive as
'  BATCH CONSTANTS in the GBE source file - only the first data row carries a value, every
'  other row in that column is blank. Read the first non-blank cell in the column (normally
'  row 2) once per run and the caller applies it to every output row. Returns "" when the
'  column is absent (columnIndex <= 0) or every data row is blank.
Private Function ResolveBatchConstantColumnValue( _
    ByVal sourceTable As Variant, _
    ByVal dataStartRowIndex As Long, _
    ByVal columnIndex As Long _
) As String
    Dim rowIndex As Long
    Dim candidate As String

    If columnIndex <= 0 Then
        ResolveBatchConstantColumnValue = vbNullString
        Exit Function
    End If

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        candidate = Trim$(CStr(sourceTable(rowIndex, columnIndex)))
        If LenB(candidate) > 0 Then
            ResolveBatchConstantColumnValue = candidate
            Exit Function
        End If
    Next rowIndex

    ResolveBatchConstantColumnValue = vbNullString
End Function
