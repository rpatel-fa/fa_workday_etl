Attribute VB_Name = "modApCfaClaimsSupplierInvoice"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_cfa_claims.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: "1.0 CFA Claims - Finance Approv" tab, header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).
'  Historical hard-coded constants retained for documentation audit only; NOT used at runtime:
'  Private Const SOURCE_WORKSHEET_NAME As String = "1.0 CFA Claims - Finance Approv"
'  Private Const SOURCE_HEADER_ROW As Long = 1
'  Private Const SOURCE_DATA_START_ROW As Long = 2

'  -- REMOVED per-period hard-coded constants / deprecated supplier-lookup paths (Rules 0/1/2):
'  Private Const SOURCE_COMPANY_CODE As String = "FA"
'  Private Const BATCH_ID_TOKEN As String = "CFAC0726"
'  mSupplierLookup / BuildReferenceLookupFromWorkbook("Supplier") / ResolveSupplierWithCompanyPrecedence
'     ALL DEPRECATED 2026-08-24 for CFA Claims.  Claims col L ("Workday Supplier ID") is the
'     AP-team-curated Workday Supplier Reference ID, used DIRECTLY (no County-FA-GP-ID lookup).
'  Company code is a constant "FA".  Claims tab B = County FA NAME, not legal-entity code.
'  batchIdToken is DERIVED from first-valid-row col H (Processed Date) -> "CFAC" + MMYY.

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"

'  transform_constants.uniform_tax (CFA Claims carry NO tax columns; ALL rows = zero-rated branch).
'  These are synced 1:1 to the uniform_tax.applicability/code/recoverability triple in
'  transforms/ap_cfa_claims.yaml -- build_runner.py VBA hygiene step enforces the match.
Private Const OUTPUT_TAX_APPLICABILITY As String = "GBR_Outside_of_Scope"
Private Const OUTPUT_TAX_CODE As String = "GBR_Zero_Rated"
Private Const OUTPUT_TAX_RECOVERABILITY As String = "Nonrecoverable_Allocate"

Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

' NEW 2026-08-24: CFA Claims Finance-Approv tab layout constants.
' Kept as Private Consts (not Config-driven) because they are structural change-of-source,
' not month-on-month tuning. Mirror the claims_constants block in the profile YAML.
Private Const CONSTANT_COMPANY_CODE As String = "FA"
Private Const BATCH_ID_TOKEN_PREFIX As String = "CFAC"

Private mCompanyId As String
Private mSpendCategoryLookup As Collection
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

Public Sub RunApCfaClaimsSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim supplierSpreadsheetPath As String
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
    Dim rowIndex As Long
    Dim dataRowCounter As Long
    '  -- Claims column indices (cols A..L; header row 1, ALL 12 headers present, no pipe aliases needed) --
    '  A = Sheet Name / Claim Category (e.g. "0.6 AdHoc Payments" -- repeats across group; used for AU suffix)
    '  B = County FA                      (claimant entity; informational, NOT legal-entity company)
    '  C = County FA GP ID                (e.g. BED013 -- RETAINED for skip-row gate + traceability, NOT resolution)
    '  D = Total Claim ()                (PurchasesAmount net)
    '  E = Budget Code                    (accounting string: dash-separated OR no-separator form -> normalize)
    '  F = Head of Approval               (informational comment; carried into Memo)
    '  G = Processed By                   (informational; carried into Memo)
    '  H = Processed Date                 (date; per-row InvoiceDate + ReceiptDate; bootstrap = batch token date)
    '  I = Payment Run                    ("August" -- informational)
    '  J = Budget Code Naming             (item description DI label; human readable)
    '  K = Supplier Name                  (informational only, not used for resolution)
    '  L = Workday Supplier ID            (AP-team-curated S-XXXXXX -- DIRECT passthrough to EIB col O)
    Dim sheetNameColumn As Long                 ' A
    Dim countyFaColumn As Long                  ' B
    Dim gpIdColumn As Long                      ' C (retained for skip-gate + traceability only)
    Dim purchasesAmountColumn As Long           ' D (Total Claim)
    Dim budgetCodeColumn As Long                ' E (accounting string -> dash-normalize before mapping lookup)
    Dim headOfApprovalColumn As Long            ' F (informational; memo part #1)
    Dim processedByColumn As Long               ' G (informational; memo part #2)
    Dim processedDateColumn As Long             ' H (InvoiceDate + ReceiptDate; bootstrap row = batch token)
    Dim paymentRunColumn As Long                ' I (informational)
    Dim budgetCodeNamingColumn As Long          ' J (Item Description DI label)
    Dim supplierNameColumn As Long              ' K (informational)
    Dim workdayIdColumn As Long                 ' L (DIRECT PASSTHROUGH -- AP-team Workday Supplier ref)
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
    Dim batchIdToken As String
    Dim mixedBatchRows As Long
    Dim workdayIdBlankRows As Long
    Dim invoiceGroupCounters As New Collection   ' key="Naming_PaymentRun_YYYY", item=next counter Long (1-based)
    ' Per-Rule 8 (2026-08-25 rework): InvoiceNumber = "{J}_{I}_{YYYY(H)}_{per-group counter}".
    ' Counters reset to 1 per unique (Budget Code Naming J + Payment Run I + Year-of-Processed-Date H) group.
    ' Using a plain VBA Collection (NOT Scripting.Dictionary).  Scripting.Dictionary requires the
    ' Scripting Runtime ActiveX library, which is NOT available on Mac Excel (it throws the
    ' run-time 429 "ActiveX component can't create object" the user saw).  Collection is built-in
    ' pure-VBA and works identically on Mac + Windows Excel.

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
    supplierSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_SUPPLIER_SPREADSHEET)
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

    ' NOTE (2026-08-24): supplierSpreadsheetPath is STILL READ from the Runner controls for
    ' backwards compatibility with the Runner Shell (which always emits it), but the Supplier
    ' Spreadsheet is NO LONGER OPENED and NO lookup is performed against it.  Claims col L
    ' ("Workday Supplier ID") = the AP-team-curated Workday Supplier Reference used DIRECTLY.
    ' The lines below intentionally removed:
    '   Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    '   Set mSupplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    '   supplierWorkbook.Close SaveChanges:=False

    '  Source of truth: config/transforms/ap_cfa_claims.yaml source_layout.required_headers.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApCfaClaimsSupplierInvoice", 12)

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
        Err.Raise vbObjectError + 2010, "modApCfaClaimsSupplierInvoice.RunApCfaClaimsSupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    ' Claims: all 12 columns have explicit headers in row 1.  No headerless DocNo columns (synthetic DocNo below).
    sheetNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Sheet Name")
    countyFaColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "County FA")
    gpIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "County FA GP ID")
    purchasesAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Total Claim ()|Total Claim|Total Claim (|Total Claim (GBP)")
    budgetCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Budget Code")
    headOfApprovalColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Head of Approval")
    processedByColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Processed By")
    processedDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Processed Date")
    paymentRunColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Payment Run")
    budgetCodeNamingColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Budget Code Naming")
    supplierNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Supplier Name")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday Supplier ID")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    '  -- Read BATCH-LEVEL constants / derived tokens from Claims + CONSTANTS --
    '  sourceCompanyCode = CONSTANT "FA".
    '  batchIdToken = "CFAC" + MMYY derived from FIRST VALID ROW's col H (Processed Date).
    '  NOTE on uniformity (mixedBatchRows): CFA Claims legitimately combine TWO payment-run
    '  cycles (e.g. 2026-07-27 + 2026-08-01) inside one month's file.  We count mismatches and
    '  REPORT them as a batch-level warning rather than hard-failing.
    sourceCompanyCode = CONSTANT_COMPANY_CODE
    batchIdToken = vbNullString
    workdayIdBlankRows = 0
    dataRowCounter = 0
    Dim bootstrapRow As Long
    For bootstrapRow = dataStartRowIndex To UBound(sourceTable, 1)
        If RowIsEmptyOrTrailer(sourceTable, bootstrapRow, _
            Array(sheetNameColumn, gpIdColumn, budgetCodeColumn, budgetCodeNamingColumn, workdayIdColumn), _
            Array(purchasesAmountColumn), _
            Array(sheetNameColumn)) Then
            GoTo nextBootstrapRowClaims
        End If
        '  -- Silent pre-flight 1: TRAILING ADMIN / SUM footer rows (e.g. Excel SUM appended below
        '     last claim line at row 273 in the Aug 2026 file).  ALL FOUR of County-FA/B,
        '     Budget-Code/E, Head-of-Approval/F, Workday-Supplier-ID/L are blank, amount D
        '     is the only non-blank (SUM).  Silently skip -- do NOT count as a reject
        '     row (the user asked explicitly: "ignore rows where cols B+E+F+L not populated"
        '     because these are Finance-appended summary lines, not real claims).
        If RowIsSilentTrailerAdminClaims(sourceTable, bootstrapRow, _
            countyFaColumn, budgetCodeColumn, headOfApprovalColumn, workdayIdColumn) Then
            GoTo nextBootstrapRowClaims
        End If
        If RowIsSkippable(sourceTable, bootstrapRow, _
            sheetNameColumn, countyFaColumn, gpIdColumn, _
            processedDateColumn, budgetCodeColumn, headOfApprovalColumn, _
            budgetCodeNamingColumn, purchasesAmountColumn, workdayIdColumn) Then
            GoTo nextBootstrapRowClaims
        End If
        ' First valid row -> derive batchIdToken from col H (Processed Date)
        Dim batchDateValue As Variant
        batchDateValue = sourceTable(bootstrapRow, processedDateColumn)
        batchIdToken = DeriveBatchIdTokenFromDate(batchDateValue)
        Exit For
nextBootstrapRowClaims:
    Next bootstrapRow

    If LenB(batchIdToken) = 0 Then
        Err.Raise vbObjectError + 2012, "modApCfaClaimsSupplierInvoice.RunApCfaClaimsSupplierInvoice", _
            "Claims col H (Processed Date = date) was BLANK or unparseable on every non-skippable " _
            & "data row. Cannot build batchIdToken (InvoiceNumber prefix). Expected a date cell " _
            & "e.g. 2026-07-27  derives CFAC0726. The Claims Processed Date column is the " _
            & "authoritative source for both InvoiceDate/ReceiptDate AND the invoice-number prefix."
    End If

    mixedBatchRows = 0
    ' invoiceGroupCounters is auto-instantiated as New Collection at first use (Dim'd As New Collection
    ' above).  No CreateObject call -- no Scripting.Runtime / ActiveX components required for Mac compatibility.

    companyId = ResolveMappedCompany(runnerWorkbook, sourceCompanyCode, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        '  -- Blank / trailer row gate (summation rows below the data). --
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, _
            Array(sheetNameColumn, gpIdColumn, budgetCodeColumn, budgetCodeNamingColumn, workdayIdColumn), _
            Array(purchasesAmountColumn), _
            Array(sheetNameColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(sheetNameColumn, gpIdColumn, budgetCodeColumn, budgetCodeNamingColumn, workdayIdColumn), _
                Array(purchasesAmountColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 8 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        '  -- Silent pre-flight: trailing admin / Excel SUM footer rows.  Disposition declared in
        '     transform_specs/ap_cfa_claims.yaml (skip_silently_not_counted_as_reject) -- do NOT
        '     increment rowsSkipped / missingFieldRows, do NOT AppendRejectRow, just advance.
        ElseIf RowIsSilentTrailerAdminClaims(sourceTable, rowIndex, _
            countyFaColumn, budgetCodeColumn, headOfApprovalColumn, workdayIdColumn) Then
            ' Intentionally blank: skip this row silently.
            consecutiveEmptyRows = 0
        ElseIf RowIsSkippable(sourceTable, rowIndex, _
            sheetNameColumn, countyFaColumn, gpIdColumn, _
            processedDateColumn, budgetCodeColumn, headOfApprovalColumn, _
            budgetCodeNamingColumn, purchasesAmountColumn, workdayIdColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required Claims field (Sheet Name/A blank / County-FA-GP-ID/C blank / " _
                    & "Processed-Date/H blank / Budget-Code/E blank / " _
                    & "Budget-Code-Naming/J blank / Total-Claim/D <=0 / Workday-Supplier-ID/L blank)", _
                Trim$(CStr(sourceTable(rowIndex, gpIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
        Else
            dataRowCounter = dataRowCounter + 1
            '  -- Uniformity check: non-skipped rows may DIFFER on ProcessedDate because CFA Claims
            '     combine two payment cycles in one month's file (Aug 2026 exemplar: 2026-07-27 +
            '     2026-08-01).  mismatches -> warning counter, not hard-fail.
            Dim rowBatchToken As String
            Dim rowBatchDate As Variant
            rowBatchDate = sourceTable(rowIndex, processedDateColumn)
            rowBatchToken = DeriveBatchIdTokenFromDate(rowBatchDate)
            If LenB(rowBatchToken) > 0 And rowBatchToken <> batchIdToken Then
                mixedBatchRows = mixedBatchRows + 1
            End If

            ' Budget Code (E) normalization: source provides the same accounting string in TWO
            ' forms within a single file: (a) dash-separated "1224-407003-..." and (b) no-separator
            ' "1224407003...".  Normalize both to dash-form BEFORE running the lowest-hierarchy
            ' mapping lookup (Mapping MASTER v37 stores OLD CODE with dashes).
            Dim normalizedBudgetCode As String
            normalizedBudgetCode = NormalizeCostCodeSeparators(sourceTable(rowIndex, budgetCodeColumn))

            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, normalizedBudgetCode)
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Cost centre unresolved: Budget Code (normalized) did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                    Trim$(CStr(sourceTable(rowIndex, budgetCodeColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, budgetCodeNamingColumn)))
            Else
                spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, normalizedBudgetCode)
                If LenB(spendCategoryId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Spend Category unresolved: Budget Code (normalized) did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", Core vs Prism = Core, lowest hierarchy match)", _
                        Trim$(CStr(sourceTable(rowIndex, budgetCodeColumn))), _
                        Trim$(CStr(sourceTable(rowIndex, budgetCodeNamingColumn)))
                Else
                    '  -- Supplier = WORKDAY SUPPLIER ID PASSTHROUGH (Rule 0). NO Supplier Spreadsheet lookup. --
                    supplierId = Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
                    If LenB(supplierId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unresolvedSupplierRows = unresolvedSupplierRows + 1
                        workdayIdBlankRows = workdayIdBlankRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                            "Supplier unresolved: Workday Supplier ID (Claims col L) was BLANK. " _
                                & "As of 2026-08-24 the AP team curates the Workday Supplier Reference " _
                                & "(S-XXXXXX) directly in Claims col L ""Workday Supplier ID""; " _
                                & "the County-FA-GP-ID / Creditor-ID Supplier Spreadsheet lookup is retired. " _
                                & "County FA GP ID / col C (for traceability): " _
                                & Trim$(CStr(sourceTable(rowIndex, gpIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, gpIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, supplierNameColumn)))
                    Else
                        If WriteOutputRow( _
                            targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                            dataRowCounter, sheetNameColumn, gpIdColumn, _
                            budgetCodeNamingColumn, paymentRunColumn, processedDateColumn, processedDateColumn, _
                            normalizedBudgetCode, purchasesAmountColumn, _
                            costCenterId, spendCategoryId, supplierId, optionalWorktagBlankRows, _
                            sourceCompanyCode, invoiceGroupCounters, _
                            headOfApprovalColumn, processedByColumn, supplierNameColumn) _
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
            & " source row(s) with missing Sheet Name/A, County FA GP ID/C, Processed Date/H blank, " _
            & "Budget Code/E blank, Budget Code Naming/J blank, Total Claim/D <=0, or Workday Supplier ID/L blank."
    End If
    If mixedBatchRows > 0 Then
        issueParts.Add "Batch-token uniformity warning: " & CStr(mixedBatchRows) _
            & " non-skipped source row(s) had a Processed-Date/H whose derived token DIFFERED from " _
            & "first valid row (expected same payment run, but CFA Claims files routinely combine " _
            & "TWO payment cycles (e.g. 2026-07-27 + 2026-08-01) in a single month file -- not a hard " _
            & "error.  Bootstrap token = " & batchIdToken & ")."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " source row(s) whose Claims col L ""Workday Supplier ID"" was BLANK. " _
            & "Supplier Spreadsheet / County-FA-GP-ID lookup was RETIRED 2026-08-24; " _
            & "AP team must populate ""Workday Supplier ID"" on every row."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) _
            & " source row(s) whose Budget Code (normalized) did not resolve to a Workday cost centre" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) _
            & " source row(s) whose Budget Code (normalized) did not resolve to a Workday spend category" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) _
            & " source row(s) had 1 or more blank optional worktags" _
            & " (Event/Camp/FA Learning/Commercial Partner/Grant/Program/Project)" _
            & " -- Budget Code (normalized) did not fully map all 7 worktag hierarchies."
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
    Err.Raise failureNumber, "modApCfaClaimsSupplierInvoice.RunApCfaClaimsSupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  DeriveBatchIdTokenFromDate -- turn a Claims col H date-cell into "CFAC" + MMYY.
'  INPUT: any date value acceptable to CDate / IsDate (Excel serial-date Double, Date,
'         or date-string "2026-07-27" already coerced to text).
'  OUTPUT: e.g. date = 2026-07-27  "CFAC0726".
'          Empty / unparseable  vbNullString (caller raises error).
'  --------------------------------------------------------------
Private Function DeriveBatchIdTokenFromDate(ByVal rawDateValue As Variant) As String
    DeriveBatchIdTokenFromDate = vbNullString
    If IsEmpty(rawDateValue) Then Exit Function
    If IsError(rawDateValue) Then Exit Function
    If LenB(CStr(rawDateValue)) = 0 Then Exit Function

    Dim parsed As Date
    On Error Resume Next
    parsed = CDate(rawDateValue)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim mm As String
    Dim yy As String
    mm = Right$("0" & CStr(Month(parsed)), 2)
    yy = Right$(CStr(Year(parsed)), 2)
    DeriveBatchIdTokenFromDate = BATCH_ID_TOKEN_PREFIX & mm & yy
End Function

'  --------------------------------------------------------------
'  Write one EIB output row.
'  INPUTS:
'     dataRowCounter = 1..N running index of NON-SKIPPED data rows (used to be the
'        SuppliersInvoiceNumber suffix; now BG is a mirror of K, but retained for trace counters)
'     sheetNameColumn / gpIdColumn / budgetCodeNamingColumn / paymentRunColumn /
'     invoiceDateColumn (= processedDateColumn, H) /
'     receiptDateColumn (= processedDateColumn, H -- same because no separate ReceiptDate in Claims) /
'     normalizedBudgetCode (already dash-normalized via NormalizeCostCodeSeparators) /
'     purchasesAmountColumn
'        -- 1-based column indices into sourceTable (Claims), plus normalized budget code string.
'     sourceCompanyCode = constant "FA"
'     invoiceGroupCounters = VBA Collection (Mac-safe: NO Scripting.Runtime / ActiveX required) keyed
'        by "{BudgetCodeNaming}_{PaymentRun}_{YYYY(ProcessedDate)}" storing next available
'        per-group counter (1-based, serialized as Variant/String Long).  See Rule 8 below and the
'        CollectionKeyExists / CollectionSetItemLong / CollectionGetItemLong module-level helpers.
'  OUTPUT COLUMN MAPPING (matches mapping: in config/transform_specs/ap_cfa_claims.yaml -- Submit_Supplier_Invoice_v46.2):
'     B  seq                (sequenceNumber)
'     I  submit             (Y)
'     K  InvoiceNumber      (Rule 8, CHANGED 2026-08-25: "{J Naming}_{I PaymentRun}_{YYYY(H ProcessedDate)}_{per-group counter}"
'                              e.g. "DWF_PartnerAugust2026_August_2026_1",
'                                   "CFA_COACHDEVWORKFUND_August_2026_1"
'                              Per-group counter RESETS to 1 per unique (J,I,YYYY(H)) group.)
'     R  Company            (mCompanyId resolved from constant "FA"  THE FA GROUP - FA)
'     T  Currency           (GBP default)
'     U  Supplier           (supplierId = Claims col L passthrough, NO lookup)
'     AF InvoiceDate        (Rule 5 -- Claims H per-row via NormalizeSourceDate)
'     AG ReceiptDate        (Rule 6 -- Invoice Received Date, Claims H per-row via NormalizeSourceDate; =InvoiceDate)
'     BG SuppliersInvoiceNumber / voucher  (Rule 7 CHANGED 2026-08-25 after Finance v2 witness --
'                              1:1 COPY of Col K InvoiceNumber.  WriteOutputRow writes the SAME
'                              `invoiceNumberValue` string instance to both ColumnIndex("K") AND
'                              ColumnIndex("BG") to guarantee bitwise equality per non-skipped data
'                              row.  CFA Workday validates cross-column consistency and rejects
'                              diverged pairs; v1 had a legacy "{GP ID C}-{SheetName A}-{dataRowCounter}"
'                              composite that Finance hand-edited 271 cells in v2 to equal K.)
'     EK line number 1      (line Row ID)
'     EP ItemDescription    (Rule 4 -- UPPER(TRIM(Claims J "Budget Code Naming")))
'     EW Spend Category     (normalized Budget Code low-hierarchy match; gated upstream)
'     FE FF FH FK FL  tax ALWAYS OUTSIDE / ZERO / NONRECOVERABLE (Claims source has NO Tax columns)
'     GE Unit Cost / GF Extended Amount (PurchasesAmount Claims D)
'     GQ Memo               (Rule 14 -- "{C} | {A} | {J} | by {G}")
'     GR Reporting Type     (second segment first char via ResolveReportingTypeFromFinalCode on normalized Budget Code)
'     GT Cost Centre        (normalized Budget Code low-hierarchy match; gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (Event Camp FALearning CommercialPartner Grant Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRow( _
    ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal dataRowCounter As Long, _
    ByVal sheetNameColumn As Long, ByVal gpIdColumn As Long, ByVal budgetCodeNamingColumn As Long, _
    ByVal paymentRunColumn As Long, _
    ByVal invoiceDateColumn As Long, ByVal receiptDateColumn As Long, _
    ByVal normalizedBudgetCode As String, ByVal purchasesAmountColumn As Long, _
    ByVal costCenterId As String, ByVal spendCategoryId As String, ByVal supplierId As String, _
    ByRef optionalWorktagBlankRows As Long, ByVal sourceCompanyCode As String, _
    ByRef invoiceGroupCounters As Collection, _
    Optional ByVal headOfApprovalColumn As Long = 0, _
    Optional ByVal processedByColumn As Long = 0, Optional ByVal supplierNameColumn As Long = 0 _
) As Boolean
    Dim amount As Double
    Dim sheetNameValue As String
    Dim gpIdValue As String
    Dim suppliersInvoiceNumberValue As String
    Dim invoiceNumberValue As String
    Dim descriptionValue As String
    Dim invoiceDateValue As String
    Dim receiptDateValue As String
    Dim headOfApprovalValue As String
    Dim processedByValue As String
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
    Dim anyWorktagBlank As Boolean
    ' ---- Rule 8 (InvoiceNumber Col K / EIB G) locals -- CHANGED 2026-08-25 ----
    Dim namingValue As String       ' source J Budget Code Naming (kept AS-IS case, NOT upper'd, because
                                    ' CFA end-users want the exact Finance-approved underscore-case label)
    Dim paymentRunValue As String   ' source I Payment Run literal (e.g. "August")
    Dim rawProcessedDate As Variant ' source H -- we need Year(); also passed to NormalizeSourceDate below
    Dim processedYearStr As String  ' 4-digit year from Processed Date H
    Dim invoiceGroupKey As String   ' composite: "{Naming}_{PaymentRun}_{YYYY}"
    Dim nextSeq As Long             ' next per-group counter for this key (starts at 1)

    sheetNameValue = Trim$(CStr(sourceTable(sourceRow, sheetNameColumn)))
    gpIdValue = Trim$(CStr(sourceTable(sourceRow, gpIdColumn)))
    namingValue = Trim$(CStr(sourceTable(sourceRow, budgetCodeNamingColumn)))                 ' Rule 8 -- kept AS-IS case
    paymentRunValue = Trim$(CStr(sourceTable(sourceRow, paymentRunColumn)))                    ' Rule 8 -- literal
    descriptionValue = UCase$(namingValue)                                                      ' Rule 4 = UPPER(namingValue)
    rawProcessedDate = sourceTable(sourceRow, invoiceDateColumn)                               ' Rule 8 -- via NormalizeSourceDate
    amount = ResolveAmount(sourceTable(sourceRow, purchasesAmountColumn))
    invoiceDateValue = NormalizeSourceDate(rawProcessedDate)      ' Rule 5 -- per-row ("YYYY-MM-DD" ISO)
    receiptDateValue = NormalizeSourceDate(rawProcessedDate)      ' Rule 6 -- = per-row H
    If headOfApprovalColumn > 0 Then headOfApprovalValue = Trim$(CStr(sourceTable(sourceRow, headOfApprovalColumn)))
    If processedByColumn > 0 Then processedByValue = Trim$(CStr(sourceTable(sourceRow, processedByColumn)))

    ' Rule 7 CHANGED 2026-08-25 per Finance v2 hand-edit witness:
    ' SuppliersInvoiceNumber (target BG) = EXACT 1:1 MIRROR of InvoiceNumber (target K)
    ' on the same non-skipped row.  Writes the SAME `invoiceNumberValue` string instance
    ' to both ColumnIndex("K") and ColumnIndex("BG") so K == BG for every row (no separate
    ' derivation, no independent counters, no separator drift).  The CFA Workday system
    ' validates cross-column consistency between the two invoice-number fields and rejects
    ' submissions where they diverge; v1 wrote a legacy "{GP ID C}-{Sheet Name A}-{dataRowCounter}"
    ' composite which caused rejection; Finance AP users hand-edited 271 BG cells of v2 to
    ' equal their K counterparts which succeeded.  Because we use the same string value we
    ' also INHERIT uniqueness FOR FREE from Rule 8's per-group counter (no separate audit needed).
    ' NOTE: the assignment must happen AFTER Rule 8 computes invoiceNumberValue (a few lines down).
    ' Rule 8 (CHANGED 2026-08-25) -- InvoiceNumber (EIB Col K / G):
    '   "{Budget Code Naming J}_{Payment Run I}_{4-digit Year(H ProcessedDate)}_{per-group counter}"
    ' Counter starts at 1 and RESETS per unique composite group key (source order within group).
    ' Year is taken from the ALREADY-NORMALIZED invoiceDateValue (ISO YYYY-MM-DD -> Left 4 chars) so
    ' it is GUARANTEED to match the year that is written to EIB cols AF/AG InvoiceDate / ReceiptDate.
    ' (NormalizeSourceDate resolves Excel serial-Double numerics, text DMY, etc. -- reusing its output
    '  prevents any CDate()-vs-serial interpretation skew on edge-case cells.)
    ' IMPLEMENTATION NOTE (Mac-safety): Uses plain VBA Collection (NOT Scripting.Dictionary) via the
    ' CollectionKeyExists / CollectionSetItemLong / CollectionGetItemLong private helpers at module
    ' bottom.  Scripting.Dictionary triggers run-time 429 on Mac Excel (no Scripting Runtime ActiveX).
    processedYearStr = Left$(invoiceDateValue, 4)
    invoiceGroupKey = namingValue & "_" & paymentRunValue & "_" & processedYearStr
    If CollectionKeyExists(invoiceGroupCounters, invoiceGroupKey) Then
        nextSeq = CollectionGetItemLong(invoiceGroupCounters, invoiceGroupKey)
    Else
        nextSeq = 1
    End If
    CollectionSetItemLong invoiceGroupCounters, invoiceGroupKey, nextSeq + 1   ' store NEXT value for future rows in this group
    invoiceNumberValue = namingValue & "_" & paymentRunValue & "_" & processedYearStr & "_" & CStr(nextSeq)
    ' Rule 7 (mirror) assignment happens NOW -- AFTER invoiceNumberValue is known; SAME value on K and BG.
    suppliersInvoiceNumberValue = invoiceNumberValue

    eventId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mEventLookup, normalizedBudgetCode)
    campId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCampLookup, normalizedBudgetCode)
    faLearningId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mFaLearningLookup, normalizedBudgetCode)
    commercialPartnerId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCommercialPartnerLookup, normalizedBudgetCode)
    grantId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mGrantLookup, normalizedBudgetCode)
    programId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProgramLookup, normalizedBudgetCode)
    projectId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProjectLookup, normalizedBudgetCode)

    anyWorktagBlank = (LenB(eventId) = 0) Or (LenB(campId) = 0) Or (LenB(faLearningId) = 0) _
        Or (LenB(commercialPartnerId) = 0) Or (LenB(grantId) = 0) Or (LenB(programId) = 0) _
        Or (LenB(projectId) = 0)
    If anyWorktagBlank Then
        optionalWorktagBlankRows = optionalWorktagBlankRows + 1
    End If

    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceNumberValue, mGrpNextKey, lineOrderValue)
    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue          ' Rule 8
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = mCompanyId
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = OUTPUT_DEFAULT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId                 ' Rule 0 -- Workday ID passthrough
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue              ' Rule 5 (per-row)
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = receiptDateValue             ' Rule 6 (= per-row H)
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = suppliersInvoiceNumberValue   ' Rule 7 (mirror K: Suppliers Invoice Number == Invoice Number)
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue    ' Row ID = 1..N within invoice group
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue    ' Line Order = 1..N within invoice group (mirror of EK)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue            ' Rule 4 -- UPPER()
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    ' Rule 13 -- Tax DS/DT/DY ALWAYS OUTSIDE / ZERO / NONRECOVERABLE (Claims source has NO Tax columns).
    ' Hard-coded zero-rated branch; conditional_on_tax_amount is never entered.  If Finance later
    ' add a TaxAmount column, copy the ap_cfa_distributions taxIsStandard conditional pattern verbatim.
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = OUTPUT_TAX_APPLICABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = OUTPUT_TAX_CODE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = OUTPUT_TAX_RECOVERABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = amount
    ' Rule 14 -- Memo: "{GP ID} | {Sheet Name} | {Budget Code Naming} | by {Processed By}"
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = gpIdValue _
        & " | " & sheetNameValue _
        & " | " & Trim$(CStr(sourceTable(sourceRow, budgetCodeNamingColumn))) _
        & IIf(LenB(processedByValue) = 0, vbNullString, " | by " & processedByValue)
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(normalizedBudgetCode)
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

' -----------------------------------------------------------------
'  Claims row-skip helpers (matches skip_rule.drop_when in transform_specs/ap_cfa_claims.yaml).
'     sheetNameColumn = A, countyFaColumn = B, gpIdColumn = C (trace-only, NOT used for resolution),
'     processedDateColumn = H (date), budgetCodeColumn = E, headOfApprovalColumn = F,
'     budgetCodeNamingColumn = J, purchasesAmountColumn = D, workdayIdColumn = L.
'     Document type is HARD-CODED "INVOICE" literal (Claims tab has no DocumentType column),
'     so no column-level gate on doc type.
' -----------------------------------------------------------------
Private Function RowIsSilentTrailerAdminClaims( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal countyFaColumn As Long, _
    ByVal budgetCodeColumn As Long, _
    ByVal headOfApprovalColumn As Long, _
    ByVal workdayIdColumn As Long _
) As Boolean
    Dim b1 As String, b2 As String, b3 As String, b4 As String
    b1 = Trim$(CStr(tableData(rowIndex, countyFaColumn)))
    b2 = Trim$(CStr(tableData(rowIndex, budgetCodeColumn)))
    b3 = Trim$(CStr(tableData(rowIndex, headOfApprovalColumn)))
    b4 = Trim$(CStr(tableData(rowIndex, workdayIdColumn)))
    If LenB(b1) = 0 And LenB(b2) = 0 And LenB(b3) = 0 And LenB(b4) = 0 Then
        RowIsSilentTrailerAdminClaims = True
    Else
        RowIsSilentTrailerAdminClaims = False
    End If
End Function

' -----------------------------------------------------------------
'  Per-row data-integrity gates (counted as reject rows if any fire).
'  MISSING-FIELD checks: Sheet Name/A blank OR County-FA-GP-ID/C blank OR
'  Processed-Date/H blank OR Budget-Code/E blank OR Budget-Code-Naming/J blank OR
'  amount D <= 0 (claim must have a positive value) OR Workday-Supplier-ID/L blank.
'  All of these indicate a malformed claim line (Finance should investigate) so they
'  ARE written to the reject sheet and increment rowsSkipped / missingFieldRows.
' -----------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal sheetNameColumn As Long, _
    ByVal countyFaColumn As Long, _
    ByVal gpIdColumn As Long, _
    ByVal processedDateColumn As Long, _
    ByVal budgetCodeColumn As Long, _
    ByVal headOfApprovalColumn As Long, _
    ByVal budgetCodeNamingColumn As Long, _
    ByVal purchasesAmountColumn As Long, _
    ByVal workdayIdColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, sheetNameColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, gpIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, processedDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, budgetCodeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, budgetCodeNamingColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, purchasesAmountColumn)) <= 0# Or _
        LenB(Trim$(CStr(tableData(rowIndex, workdayIdColumn)))) = 0
End Function

' Keyed-Collection primitives (CollectionKeyExists / *SetItemLong / *GetItemLong /
' *SetItemDouble / *GetItemDouble) are shared in modHelpers -- no per-module copy (2026-09-14).

