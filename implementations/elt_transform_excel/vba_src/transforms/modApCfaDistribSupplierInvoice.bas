Attribute VB_Name = "modApCfaDistribSupplierInvoice"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_cfa_distributions.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

' REWORKED 2026-09-24 for the September-2026 "Finance Master analysis" tab. INVOICES ONLY:
' processes rows where col U "Invoice type" = "Invoice"; credit-note rows are filtered out here
' and handled by modApCfaDistribAmendment (Supplier Invoice Adjustment EIB). Invoice Number (Z),
' Invoice Description (AA), Invoice Date (AD) and Original Amount (V) are read directly; the CFAD
' batch-token / DocNo invoice-number derivation is retired. Worktags/Reporting Type still resolve
' from Distribution/Cost Code (col G) via Mapping MASTER; tax stays conditional on Tax amount (I).

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"

Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"

'  --- Profile-local Tax Triple override (shadows modFinanceCommon shared standard-rated) ---
'  AP Supplier-Invoice standard-rated lines use the domestic PURCHASE ref-IDs (AP confirmed
'  2026-09-14). Zero-rated + tax point/option defaults still resolve from shared modFinanceCommon.
Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Purchase_Of_Goods_And_Services"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_RECOVERABILITY_STANDARD As String = "Fully_Recoverable"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

' Company is a constant "FA" (CFA always = THE FA GROUP). Only invoice rows are processed here.
Private Const CONSTANT_COMPANY_CODE As String = "FA"
Private Const INVOICE_TYPE_INVOICE As String = "INVOICE"

Private mCompanyId As String
Private mSpendCategoryLookup As Collection
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

Public Sub RunApCfaDistributionsSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim mappingWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As Variant
    Dim mappingTable As Variant
    Dim costCenterLookup As Collection
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unmappedCostCenterRows As Long
    Dim unmappedSpendCategoryRows As Long
    Dim unresolvedSupplierRows As Long
    Dim optionalWorktagBlankRows As Long
    Dim nonInvoiceRows As Long
    Dim rowIndex As Long
    '  -- Finance Master analysis column indices (header row 5) --
    Dim docNoColumn As Long                  ' A  DocNo / row key (headerless; = CFA_col - 1); memo trace
    Dim cfaNameColumn As Long                ' B  County FA name (anchor)
    Dim creditorIdColumn As Long             ' D  Creditor ID (trace + skip gate only)
    Dim distributionRefCodeColumn As Long    ' G  Distribution/Cost Code (accounting string)
    Dim taxAmountColumn As Long              ' I  Tax amount (conditional switch)
    Dim workdayIdColumn As Long              ' T  Workday ID (supplier passthrough)
    Dim invoiceTypeColumn As Long            ' U  Invoice type ("Invoice" here; else routed to amendments)
    Dim originalAmountColumn As Long         ' V  Original Amount (net line; positive for invoices)
    Dim invoiceNumberColumn As Long          ' Z  Invoice Number (pre-computed -> EIB K + BG)
    Dim invoiceDescriptionColumn As Long     ' AA Invoice Description (pre-computed -> EIB EP)
    Dim invoiceDateColumn As Long            ' AD Invoice Date (pre-computed -> EIB AF/AG)
    Dim companyId As String
    Dim supplierId As String
    Dim costCenterId As String
    Dim spendCategoryId As String
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
    Dim sourceCompanyCode As String

    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)

    Set issueParts = New Collection

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
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

    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH
    EnsureNonBlank mappingSpreadsheetPath, CONTROL_LABEL_MAPPING_SPREADSHEET
    If LenB(mappingSheetName) = 0 Then mappingSheetName = DEFAULT_MAPPING_SHEET_NAME

    ' Supplier resolution is DIRECT passthrough from col T "Workday ID" (no Supplier Spreadsheet).

    '  Source of truth: config/transforms/ap_cfa_distributions.yaml source_layout.required_headers.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApCfaDistribSupplierInvoice", 10)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    Set sourceSheet = ResolveSourceWorksheet(sourceWorkbook, sheetToResolve)
    sourceTable = ReadWorksheetTable(sourceSheet)
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    Set mappingWorkbook = OpenWorkbookReadOnly(mappingSpreadsheetPath)
    mappingTable = ReadWorksheetTable(ResolveSourceWorksheet(mappingWorkbook, mappingSheetName))
    mappingWorkbook.Close SaveChanges:=False
    Set mappingWorkbook = Nothing

    Set costCenterLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_COST_CENTER)
    Set mSpendCategoryLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_SPEND_CATEGORY)
    Set mEventLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_EVENT)
    Set mCampLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_CAMP)
    Set mFaLearningLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_FA_LEARNING)
    Set mCommercialPartnerLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_COMMERCIAL_PARTNER)
    Set mGrantLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_GRANT)
    Set mProgramLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_PROGRAM)
    Set mProjectLookup = BuildBusinessObjectHierarchyLookupMulti(mappingTable, WBO_PROJECT, WBO_PROJECT_CAPEX)

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
        Err.Raise vbObjectError + 1910, "modApCfaDistribSupplierInvoice.RunApCfaDistributionsSupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    ' Column A = DocNo / row key has no explicit header; DocNo = CFA_col - 1.
    cfaNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "CFA")
    docNoColumn = cfaNameColumn - 1
    creditorIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Creditor ID")
    distributionRefCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Distribution/Cost Code")
    taxAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax amount")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday ID")
    invoiceTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice type")
    originalAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Original Amount")
    invoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Number")
    invoiceDescriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Description|Invoice Desctiption")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Date")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    sourceCompanyCode = CONSTANT_COMPANY_CODE
    companyId = ResolveMappedCompany(runnerWorkbook, sourceCompanyCode, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        '  -- Blank / trailer row gate (formula scaffold rows below the data). --
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
        ElseIf RowIsNonInvoiceType(sourceTable, rowIndex, invoiceTypeColumn) Then
            '  -- Credit notes / non-invoice rows: routed to ap_cfa_distributions_amendments.
            '     Skip silently here (not a reject) so the two 1-input->1-output transforms
            '     partition the same source by Invoice type.
            nonInvoiceRows = nonInvoiceRows + 1
            consecutiveEmptyRows = 0
        ElseIf RowIsSkippable(sourceTable, rowIndex, docNoColumn, creditorIdColumn, _
            distributionRefCodeColumn, originalAmountColumn, workdayIdColumn, _
            invoiceNumberColumn, invoiceDescriptionColumn, invoiceDateColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required invoice field (DocNo/A / Creditor ID/D / Distribution-Cost-Code/G / " _
                    & "Workday-ID/T / Invoice-Number/Z / Invoice-Description/AA / Invoice-Date/AD blank, " _
                    & "or Original-Amount/V = 0)", _
                Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
        Else
            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, sourceTable(rowIndex, distributionRefCodeColumn))
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Cost centre unresolved: Distribution/Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                    Trim$(CStr(sourceTable(rowIndex, distributionRefCodeColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, invoiceDescriptionColumn)))
            Else
                spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, sourceTable(rowIndex, distributionRefCodeColumn))
                If LenB(spendCategoryId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Spend Category unresolved: Distribution/Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", Core vs Prism = Core, lowest hierarchy match)", _
                        Trim$(CStr(sourceTable(rowIndex, distributionRefCodeColumn))), _
                        Trim$(CStr(sourceTable(rowIndex, invoiceDescriptionColumn)))
                Else
                    supplierId = Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
                    If LenB(supplierId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unresolvedSupplierRows = unresolvedSupplierRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                            "Supplier unresolved: Workday ID (Finance Master analysis col T) was BLANK. " _
                                & "AP curate the Workday Supplier Reference (S-XXXXXX) directly in col T; " _
                                & "Creditor ID / col D (for traceability): " _
                                & Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
                    Else
                        If WriteOutputRow( _
                            targetSheet, targetRow, sourceTable, rowIndex, _
                            docNoColumn, distributionRefCodeColumn, originalAmountColumn, taxAmountColumn, _
                            invoiceNumberColumn, invoiceDescriptionColumn, invoiceDateColumn, _
                            costCenterId, spendCategoryId, supplierId, optionalWorktagBlankRows) _
                        Then
                            rowsProcessed = rowsProcessed + 1
                            targetRow = targetRow + 1
                        Else
                            rowsSkipped = rowsSkipped + 1
                        End If
                    End If
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) _
            & " invoice row(s) with a missing required field (DocNo/A, Creditor ID/D, Distribution/Cost Code/G," _
            & " Workday ID/T, Invoice Number/Z, Invoice Description/AA, Invoice Date/AD, or Original Amount/V = 0)."
    End If
    If nonInvoiceRows > 0 Then
        issueParts.Add CStr(nonInvoiceRows) _
            & " row(s) with Invoice type <> 'Invoice' (credit notes) were skipped here -- they are" _
            & " handled by ap_cfa_distributions_amendments (Supplier Invoice Adjustment EIB)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " invoice row(s) whose col T ""Workday ID"" was BLANK. AP must populate it on every row."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) _
            & " invoice row(s) whose Distribution/Cost Code did not resolve to a Workday cost centre" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) _
            & " invoice row(s) whose Distribution/Cost Code did not resolve to a Workday spend category" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) _
            & " invoice row(s) had 1 or more blank optional worktags" _
            & " (Event/Camp/FA Learning/Commercial Partner/Grant/Program/Project)" _
            & " -- Distribution/Cost Code did not fully map to all 7 worktag hierarchies."
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
    Err.Raise failureNumber, "modApCfaDistribSupplierInvoice.RunApCfaDistributionsSupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row (Submit_Supplier_Invoice_v46.2).
'     B  Spreadsheet Key (AssignInvoiceLineGroup on Invoice Number)
'     I  Submit (Y)
'     K  Invoice Number (col Z, pre-computed)
'     R  Company (constant FA -> THE FA GROUP - FA)
'     T  Currency (GBP)
'     U  Supplier (col T Workday ID passthrough)
'     AF Invoice Date / AG Received Date (col AD via NormalizeSourceDate)
'     BG Suppliers Invoice Number (1:1 mirror of K)
'     EK Row ID / EM Line Order (1..N within Invoice Number group)
'     EP Item Description (col AA)
'     EW Spend Category (col G lowest-hierarchy match; gated upstream)
'     FE FF FH FK FL tax conditional_on_tax_amount (col I > 0 -> Standard, else Zero)
'     GE Unit Cost / GF Extended Amount (col V Original Amount)
'     GQ Memo (DocNo colA + " " + Distribution/Cost Code colG)
'     GR Reporting Type (col G second segment first char)
'     GT Cost Centre (col G lowest-hierarchy match; gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (from col G)
'  --------------------------------------------------------------
Private Function WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal docNoColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal originalAmountColumn As Long, _
    ByVal taxAmountColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal invoiceDescriptionColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal costCenterId As String, _
    ByVal spendCategoryId As String, _
    ByVal supplierId As String, _
    ByRef optionalWorktagBlankRows As Long _
) As Boolean
    Dim amount As Double
    Dim taxAmount As Double
    Dim distributionRefCodeValue As Variant
    Dim invoiceNumberValue As String
    Dim descriptionValue As String
    Dim invoiceDateValue As String
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
    Dim anyWorktagBlank As Boolean
    Dim taxIsStandard As Boolean
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long

    invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = Trim$(CStr(sourceTable(sourceRow, invoiceDescriptionColumn)))
    amount = ResolveAmount(sourceTable(sourceRow, originalAmountColumn))
    taxAmount = ResolveAmount(sourceTable(sourceRow, taxAmountColumn))
    distributionRefCodeValue = sourceTable(sourceRow, distributionRefCodeColumn)
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))

    ' Tax code conditional switch: Tax amount > 0 -> Standard-rated, else Zero-rated.
    taxIsStandard = (taxAmount > 0#)

    eventId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mEventLookup, distributionRefCodeValue)
    campId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCampLookup, distributionRefCodeValue)
    faLearningId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mFaLearningLookup, distributionRefCodeValue)
    commercialPartnerId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCommercialPartnerLookup, distributionRefCodeValue)
    grantId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mGrantLookup, distributionRefCodeValue)
    programId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProgramLookup, distributionRefCodeValue)
    projectId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProjectLookup, distributionRefCodeValue)

    anyWorktagBlank = (LenB(eventId) = 0) Or (LenB(campId) = 0) Or (LenB(faLearningId) = 0) _
        Or (LenB(commercialPartnerId) = 0) Or (LenB(grantId) = 0) Or (LenB(programId) = 0) _
        Or (LenB(projectId) = 0)
    If anyWorktagBlank Then
        optionalWorktagBlankRows = optionalWorktagBlankRows + 1
    End If

    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceNumberValue, mGrpNextKey, lineOrderValue)
    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = mCompanyId
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = OUTPUT_DEFAULT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = invoiceDateValue
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    If taxIsStandard Then
        targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = TAX_APPLICABILITY_STANDARD
        targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = TAX_CODE_STANDARD
        targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = TAX_RECOVERABILITY_STANDARD
    Else
        targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = TAX_APPLICABILITY_ZERO
        targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = TAX_CODE_ZERO
        targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = TAX_RECOVERABILITY_ZERO
    End If
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = Trim$(CStr(sourceTable(sourceRow, docNoColumn))) & " " & CStr(distributionRefCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(distributionRefCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = costCenterId
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = eventId
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = campId
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = faLearningId
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = commercialPartnerId
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = grantId
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = programId
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = projectId
    WriteOutputRow = True
End Function

'  --------------------------------------------------------------
'  RowIsNonInvoiceType -- TRUE when col U "Invoice type" is present and NOT "Invoice"
'  (e.g. "Credit Note"). Such rows are handled by ap_cfa_distributions_amendments.
'  A blank Invoice type is NOT treated as non-invoice here (it falls through to the
'  missing-field gate below so genuine data gaps are still reported).
'  --------------------------------------------------------------
Private Function RowIsNonInvoiceType( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal invoiceTypeColumn As Long _
) As Boolean
    Dim v As String
    v = Trim$(CStr(tableData(rowIndex, invoiceTypeColumn)))
    If LenB(v) = 0 Then
        RowIsNonInvoiceType = False
    Else
        RowIsNonInvoiceType = (StrComp(UCase$(v), INVOICE_TYPE_INVOICE, vbBinaryCompare) <> 0)
    End If
End Function

'  --------------------------------------------------------------
'  RowIsSkippable -- per-row missing-field gate (matches skip_rule in the spec YAML).
'  DocNo/A, Creditor ID/D, Distribution/Cost Code/G, Workday ID/T, Invoice Number/Z,
'  Invoice Description/AA, Invoice Date/AD blank, OR Original Amount/V = 0.
'  --------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal docNoColumn As Long, _
    ByVal creditorIdColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal originalAmountColumn As Long, _
    ByVal workdayIdColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal invoiceDescriptionColumn As Long, _
    ByVal invoiceDateColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, docNoColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, creditorIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, distributionRefCodeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, workdayIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDescriptionColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, originalAmountColumn)) = 0#
End Function
