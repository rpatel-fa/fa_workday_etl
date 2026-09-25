Attribute VB_Name = "modApRefereeV2SupplierInvoice"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_referees_v2.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: Referee Expenses - Analysis tab, header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).
'  Historical hard-coded constants retained for documentation audit only; NOT used at runtime:
'  Private Const SOURCE_WORKSHEET_NAME As String = "Referee Expenses - Analysis"
'  Private Const SOURCE_HEADER_ROW As Long = 1
'  Private Const SOURCE_DATA_START_ROW As Long = 2

'  -- per-batch derived constants --
'  BATCH_ID_TOKEN is NOT hardcoded (unlike v1 which fixed REFS0426ME).  Instead it is
'  derived ONCE per-run from the MAX(Submission Date) across all rows:
'    REFS + MM(date) + YY(date) + "ME"
'  Submission Date (Analysis col 2/B) is chosen because it is 100% populated on every
'  row by the form system at submission time (verified in the April 2026 exemplar:
'  MAX(Submission Date) = 2026-04-16 -> REFS0426ME, identical to v1).
'  Finance Processed Date (Analysis col 32/AF) was the original candidate but it is
'  NOT populated in the Approved workbook Finance shares with the agent (the date is
'  only stamped internally by Finance at the point they run the GP batch), so it is
'  NOT used for any derivation.
'  Company is DERIVED AS CONSTANT "FA" (the Analysis tab has no explicit Company column;
'  Cost Code first segment and Category both confirm all referee rows are FA-entity only,
'  identical in effect to v1's per-row Integration col A "FA" literal).
Private m_batchIdToken As String
Private m_receiptDateValue As String   ' batch MAX(Submission Date) = latest claim-submission date

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"
Private Const OUTPUT_TAX_APPLICABILITY As String = "GBR_Outside_of_Scope"
Private Const OUTPUT_TAX_CODE As String = "GBR_Zero_Rated"
Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_RECOVERABILITY As String = "Nonrecoverable_Allocate"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"
' GE / GF cell display format (transform_specs line_amount.number_format; display only, value not rounded).
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"

'  -- Analysis-tab hardcoded sentinels --
'  The Analysis tab has NO TaxAmount or TaxDetailID columns; these are the uniform
'  v1-equivalent values written to every output row.  Tax calculation is no-op because
'  amount=0 on both TaxAmount and CALC_TAX_DUE leaves Workday with nothing to add.
Private Const OUTPUT_TAX_AMOUNT As Double = 0#
Private Const OUTPUT_TAX_DETAIL_ID As String = "POS"
Private Const OUTPUT_DOCUMENT_TYPE As String = "INVOICE"
Private Const OUTPUT_COMPANY_CODE As String = "FA"

'  -- Approval gate RETIRED September 2026 --
'  Sheet1 has no HH Final Approval column (the file is the Finance-approved extract),
'  so every row is treated as approved.  The SOLE authoritative gate for Supplier is
'  "Workday ID" (Sheet1 col 18/R; was NEW WORKDAY ID col 21/U) being non-blank.

Private mSpendCategoryLookup As Collection
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

Public Sub RunApRefereeV2SupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim blankWorkdayIdRows As Long
    Dim optionalWorktagBlankRows As Long
    Dim rowIndex As Long
    '  -- Sheet1 column indices (23 cols A..W in September 2026 refresh) --
    '  A 1=INV number, B 2=Submission Date, C 3=Reference Number, D 4=Invoice Date,
    '  E 5=Full Name, F 6=Details of Invoice, G 7=Invoice Total, H 8=VENDOR ID,
    '  I 9=Address, J 10=Postcode, K 11=Cost Code, L 12=Telephone No, M 13=Email Address,
    '  N 14=Bank Name, O 15=Sort Code, P 16=Account Number, Q 17=Finance Processed Date,
    '  R 18=Workday ID (DIRECT supplier passthrough -> EIB U), S..V 19..22=bank checks,
    '  W 23=GP ID to Supplier ID (NOT used).
    '  No HH Final Approval / PE Approval / Category columns -> approval gate retired.
    '  Columns are resolved by header name, so positions are informational only.
    '
    '  -- History: analysis tab column indices (39 cols A..AM in August 2026 refresh) --
    '  A 1=INV number, B 2=Submission Date, C 3=Reference Number, D 4=Invoice Date,
    '  E 5=Full Name (As it appears on your bank account),
    '  F 6=Details of Invoice (raw), G 7=Invoice Total,
    '  H 8=VENDOR ID, I 9=Address, J 10=Postcode, K 11=Category,
    '  L 12=Cost Code (4-6-8-4 accounting string),
    '  M 13=HH Final Approval, N 14=PE Approval, O 15=PE Approval Date,
    '  P 16=Telephone No, Q 17=Email Address, R 18=Bank Name,
    '  S 19=Sort Code, T 20=Account Number,
    '  U 21=NEW WORKDAY ID      (AP/Finance curated S-XXXXXX reference - DIRECT supplier passthrough)
    '  V 22=Account number, W 23=A N check, X 24=Sort code, Y 25=S C check,
    '  Z 26=Finance Processed Date, AA..Raj 27..39=Approver name columns (Submitted/Approved).
    '
    '  August 2026 refresh: No separate "JY Details" cleaned column anymore.
    '  Description is read directly from Details of Invoice col 6/F.
    '  GP ID Match / Account number match columns no longer exist (Supplier was previously
    '  looked up from VENDOR ID via Supplier Spreadsheet + reconciled to GP register).
    '  NEW WORKDAY ID col U 21 is now the authoritative Supplier reference.
    '  VENDOR ID col H 8 is retained for trace but is NOT used for Supplier resolution.
    '  "#NO MATCH" rows in VENDOR ID are now accepted so long as NEW WORKDAY ID is populated
    '  (there are 4 such rows in the August 2026 source batch for newly-onboarded referees).
    Dim invNumberColumn As Long              ' 1 A
    Dim submissionDateColumn As Long         ' 2 B  Submission Date -- latest date used for batch token + ReceiptDate
    Dim invoiceDateColumn As Long            ' 4 D
    Dim detailsOfInvoiceColumn As Long       ' 6 F  Details of Invoice -- raw description, no JY cleaned col
    Dim invoiceTotalColumn As Long           ' 7 G  Invoice Total = PurchasesAmount
    Dim vendorIdColumn As Long               ' 8 H  VENDOR ID (OPTIONAL header, trace only; 0 when absent)
    Dim costCodeColumn As Long               ' 11 K Cost Code = 4-6-8-4 accounting string (12 L in Aug 2026)
    Dim newWorkdayIdColumn As Long           ' 18 R Workday ID = Supplier S-XXXXXX (DIRECT passthrough; mandatory; 21 U in Aug 2026)
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
    '  NOTE: supplierSpreadsheetPath may still be present in the Runner config for legacy
    '  transforms.  August 2026 refresh of AP Supplier Referees (V1/V2) retires the old
    '  Creditor-ID -> Supplier Spreadsheet lookup; supplier is instead read directly
    '  from the NEW WORKDAY ID column (col 21/U) curated in the source file by AP/Finance.
    '  We deliberately do NOT open supplierSpreadsheetPath here.
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

    '  -- Supplier Spreadsheet lookup RETIRED August 2026 --
    '  Previously: open the Supplier reference workbook and build a lookup
    '  collection via BuildReferenceLookupFromWorkbook (two-part CreditorID/Company
    '  resolution precedence).  No longer used: AP/Finance now curate the Workday
    '  S-XXXXXX supplier reference directly in the source file as "NEW WORKDAY ID"
    '  (col 21/U), which we read per-row at write time.

    '  Required headers -- analysis tab columns used by the transform.  Optional columns
    '  (Address/Postcode/Approver names etc.) are read via FindColumnIndexInRow with
    '  LenB()=0 no-op if missing, since they are used for reject-context only.
    '  Source of truth: config/transforms/*.yaml source_layout.required_headers,
    '  written into Config "Required Headers Pipe Delimited" by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApRefereeV2SupplierInvoice", 7)

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

    '  Schema validation phase (Tier 2 VBA defence-in-depth).
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
        Err.Raise vbObjectError + 1920, "modApRefereeV2SupplierInvoice.RunApRefereeV2SupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    invNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "INV number")
    submissionDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Submission Date")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Date")
    detailsOfInvoiceColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Details of Invoice")
    invoiceTotalColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Invoice Total")
    vendorIdColumn = SafeFindOptionalColumn(sourceTable, headerRowIndex, "VENDOR ID")
    costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Cost Code")
    newWorkdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "WORKDAY ID|NEW WORKDAY ID|WORKDAY SUPPLIER ID|NEW WORKDAY SUPPLIER ID")
    '  Optional columns (FindColumnIndexInRow returns 0 if not present; we guard downstream).
    '  August 2026 refresh removed GP ID Match / Account number match / Sort code Match.
    '  NEW WORKDAY ID column is now the authoritative Supplier reference; bank-validation
    '  columns are informational only.

    '  -- PRE-PASS: Derive batch constants from Submission Date --
    '  Two batch-level values are computed from the full sourceTable BEFORE the row loop:
    '    (a) m_receiptDateValue = MAX(Submission Date) -- most-recent-form-submission date
    '        within the approved batch file (deterministic analogue of v1's GP batch run date
    '        constant; April 2026 MAX=2026-04-16 vs v1 constant 2026-04-17 (1-day offset)).
    '    (b) m_batchIdToken = REFS + MM(maxDate) + YY(maxDate) + "ME".  April 2026 -> "REFS0426ME".
    '  Submission Date is 100% populated (verified).  If NO row has a usable Submission Date
    '  (impossible given the form-system stamp) the batch-wide derivation fails fast with a
    '  clear error rather than per-row confusion downstream.
    DeriveBatchConstants sourceTable, dataStartRowIndex, submissionDateColumn

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        '  -- Silent pre-flight: trailing admin / Excel SUM footer rows.
        '     August 2026 row 291: Finance-appended summary where only col G (Invoice Total)
        '     is non-blank (a SUM over the 289 data rows); INV number / Cost Code / NEW WORKDAY ID
        '     are all blank.  Disposition: skip SILENTLY (not counted as a reject row, no counter
        '     increment, no AppendRejectRow entry) because this is Finance housekeeping, not a
        '     real referee claim.  Mirrors RowIsSilentTrailerAdminClaims for CFA Claims. --
        If RowIsSilentTrailerAdminReferees(sourceTable, rowIndex, _
            invNumberColumn, costCodeColumn, newWorkdayIdColumn) Then
            consecutiveEmptyRows = 0
        '  -- Blank / empty row gate. --
        ElseIf RowIsEmptyOrTrailer(sourceTable, rowIndex, _
            Array(invNumberColumn, vendorIdColumn, detailsOfInvoiceColumn, costCodeColumn, _
                  newWorkdayIdColumn), _
            Array(invoiceTotalColumn), _
            Array(invNumberColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(invNumberColumn, vendorIdColumn, detailsOfInvoiceColumn, costCodeColumn, newWorkdayIdColumn), _
                Array(invoiceTotalColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 3 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsSkippableV2(sourceTable, rowIndex, _
            invNumberColumn, detailsOfInvoiceColumn, _
            invoiceDateColumn, submissionDateColumn, _
            costCodeColumn, invoiceTotalColumn, newWorkdayIdColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required source field (INV number blank / " _
                    & "Details of Invoice blank / Invoice Date blank / Submission Date blank / " _
                    & "Cost Code blank / Workday ID blank / Invoice Total <= 0)", _
                OptionalCellText(sourceTable, rowIndex, vendorIdColumn), _
                "Workday ID=" & Trim$(CStr(sourceTable(rowIndex, newWorkdayIdColumn))) _
                    & " | Details=" & Trim$(CStr(sourceTable(rowIndex, detailsOfInvoiceColumn)))
        Else
            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, sourceTable(rowIndex, costCodeColumn))
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Cost centre unresolved: Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                    Trim$(CStr(sourceTable(rowIndex, costCodeColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, detailsOfInvoiceColumn)))
            Else
                spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, sourceTable(rowIndex, costCodeColumn))
                If LenB(spendCategoryId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Spend Category unresolved: Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", Core vs Prism = Core, lowest hierarchy match)", _
                        Trim$(CStr(sourceTable(rowIndex, costCodeColumn))), _
                        Trim$(CStr(sourceTable(rowIndex, detailsOfInvoiceColumn)))
                Else
                    Dim resolvedCompanyId As String
                    resolvedCompanyId = ResolveMappedCompany(runnerWorkbook, OUTPUT_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)

                    '  Supplier = DIRECT passthrough of Workday ID (Sheet1 col 18/R) -> EIB U.
                    '  AP/Finance have already curated this S-XXXXXX reference.
                    '  Old Creditor-ID -> Supplier Spreadsheet lookup is RETIRED.
                    supplierId = Trim$(CStr(sourceTable(rowIndex, newWorkdayIdColumn)))
                    If LenB(supplierId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        blankWorkdayIdRows = blankWorkdayIdRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                            "Supplier unresolved: Workday ID is blank.  AP/Finance must " _
                                & "allocate a Workday Supplier S-XXXXXX reference for each referee " _
                                & "claim row and write it into the Workday ID column (R/18).", _
                            OptionalCellText(sourceTable, rowIndex, vendorIdColumn), _
                            "Company=FA | VENDOR ID for trace=" & OptionalCellText(sourceTable, rowIndex, vendorIdColumn) _
                                & " | Workday ID=" & supplierId
                    Else
                        If WriteOutputRowV2( _
                            targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                            vendorIdColumn, detailsOfInvoiceColumn, _
                            invoiceDateColumn, invNumberColumn, _
                            costCodeColumn, invoiceTotalColumn, _
                            costCenterId, spendCategoryId, supplierId, _
                            resolvedCompanyId, optionalWorktagBlankRows) _
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
            & " source row(s) with missing INV number, Details of Invoice, Invoice Date," _
            & " Submission Date, Cost Code, Workday ID, or Invoice Total <= 0."
    End If
    If blankWorkdayIdRows > 0 Then
        issueParts.Add "Skipped " & CStr(blankWorkdayIdRows) _
            & " source row(s) whose Workday ID is blank -- AP/Finance must allocate a Workday Supplier" _
            & " S-XXXXXX reference for each such row.  This replaces the old ""#NO MATCH"" VENDOR ID" _
            & " pass gate: the Creditor/GP provisioning new-referee onboarding step is now FINISHED when" _
            & " Workday ID column has an S-XXXXXX value (Supplier reference = new authority gate)."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) _
            & " source row(s) whose Cost Code did not resolve to a Workday cost centre" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) _
            & " source row(s) whose Cost Code did not resolve to a Workday spend category" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) _
            & " source row(s) had 1 or more blank optional worktags" _
            & " (Event/Camp/FA Learning/Commercial Partner/Grant/Program/Project)" _
            & " -- Cost Code did not fully map to all 7 worktag hierarchies."
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
    Err.Raise failureNumber, "modApRefereeV2SupplierInvoice.RunApRefereeV2SupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row -- V2 (Analysis tab columns).
'  INPUTS:
'     vendorIdColumn / jyDetailsColumn / invoiceDateColumn / invNumberColumn /
'     costCodeColumn / invoiceTotalColumn
'        -- 1-based column indices into sourceTable (Analysis tab).
'     resolvedCompanyId
'        -- EIB Company (col L) entity name, resolved from constant "FA" via runner
'          company-mapping rows (belt-and-braces fallback OUTPUT_COMPANY_DEFAULT_FALLBACK
'          used when runner mapping has no match for "FA").
'  OUTPUT COLUMN MAPPING (matches mapping: in config/transform_specs/ap_referees_v2.yaml -- Submit_Supplier_Invoice_v46.2):
'     B  seq            (sequenceNumber)
'     I  submit         (Y)
'     K  InvoiceNumber  (m_batchIdToken + "-" + INV number; per-INVOICE unique)
'     R  Company        (resolvedCompanyId -- constant FA -> entity via runner map)
'     T  Currency       (GBP default)
'     U  Supplier       (supplierId = Workday ID Sheet1 col 18/R direct passthrough, NO lookup)
'     AF InvoiceDate    (per-row Analysis D Invoice Date via NormalizeSourceDate)
'     AG ReceiptDate    (Invoice Received Date, m_receiptDateValue -- batch MAX(Submission Date) precomputed)
'     BG SuppliersInvoiceNumber / voucher (INV number from Analysis A -- NOV17056 etc.)
'     EK line number 1  (line Row ID)
'     EP ItemDescription (UPPER(TRIM(JY Details Analysis G)))
'     EW Spend Category  (Cost Code low-hierarchy match; on_fail=reject, gated upstream)
'     FE FF FH FK FL  tax defaults (uniform Outside of Scope / Zero Rated / CALC_TAX_DUE)
'     GE Unit Cost / GF Extended Amount (Invoice Total; NumberFormat "0.00")
'     GQ Memo (INV number + " " + Cost Code Analysis N)
'     GR Reporting Type  (Cost Code second segment first char via ResolveReportingTypeFromFinalCode)
'     GT Cost Centre     (Cost Code low-hierarchy match; on_fail=reject, gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (Event Camp FALearning CommercialPartner Grant Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRowV2( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal vendorIdColumn As Long, _
    ByVal detailsOfInvoiceColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal invNumberColumn As Long, _
    ByVal costCodeColumn As Long, _
    ByVal invoiceTotalColumn As Long, _
    ByVal costCenterId As String, _
    ByVal spendCategoryId As String, _
    ByVal supplierId As String, _
    ByVal resolvedCompanyId As String, _
    ByRef optionalWorktagBlankRows As Long _
) As Boolean
    Dim amount As Double
    Dim costCodeValue As Variant
    Dim invNumberValue As String
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

    invNumberValue = Trim$(CStr(sourceTable(sourceRow, invNumberColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, detailsOfInvoiceColumn))))
    amount = ResolveAmount(sourceTable(sourceRow, invoiceTotalColumn))
    costCodeValue = sourceTable(sourceRow, costCodeColumn)
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))

    invoiceNumberValue = m_batchIdToken & "-" & invNumberValue

    eventId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mEventLookup, costCodeValue)
    campId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCampLookup, costCodeValue)
    faLearningId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mFaLearningLookup, costCodeValue)
    commercialPartnerId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCommercialPartnerLookup, costCodeValue)
    grantId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mGrantLookup, costCodeValue)
    programId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProgramLookup, costCodeValue)
    projectId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProjectLookup, costCodeValue)

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
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = resolvedCompanyId
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = OUTPUT_DEFAULT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = m_receiptDateValue
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = invNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue    ' Row ID = 1..N within invoice group
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue    ' Line Order = 1..N within invoice group (mirror of EK)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = OUTPUT_TAX_APPLICABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = OUTPUT_TAX_CODE
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = OUTPUT_TAX_RECOVERABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = invNumberValue & " " & CStr(costCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(costCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = costCenterId
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = eventId
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = campId
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = faLearningId
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = commercialPartnerId
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = grantId
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = programId
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = projectId
    WriteOutputRowV2 = True
End Function

'  --------------------------------------------------------------
'  DeriveBatchConstants -- pre-pass across ALL data rows to compute:
'    (a) latest Submission Date -> m_receiptDateValue (as YYYY-MM-DD text),
'    (b) batch token from that date -> m_batchIdToken.
'  Fails fast if NO row has a usable Submission Date.
'  --------------------------------------------------------------
Private Sub DeriveBatchConstants( _
    ByVal sourceTable As Variant, _
    ByVal dataStartRowIndex As Long, _
    ByVal submissionDateColumn As Long _
)
    Dim rowIndex As Long
    Dim maxSerialDate As Double
    Dim haveAnyDate As Boolean
    Dim cellValue As Variant
    Dim serialVal As Double
    Dim mmStr As String
    Dim yyStr As String

    haveAnyDate = False
    maxSerialDate = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        cellValue = sourceTable(rowIndex, submissionDateColumn)
        If Not IsEmpty(cellValue) And LenB(Trim$(CStr(cellValue))) > 0 Then
            serialVal = CDbl(CDate(cellValue))
            If Not haveAnyDate Then
                maxSerialDate = serialVal
                haveAnyDate = True
            ElseIf serialVal > maxSerialDate Then
                maxSerialDate = serialVal
            End If
        End If
    Next rowIndex

    If Not haveAnyDate Then
        Err.Raise vbObjectError + 1925, "modApRefereeV2SupplierInvoice.DeriveBatchConstants", _
            "No rows had a usable Submission Date (Analysis col 2/B). " _
                & "Cannot derive batch token (REFS{MM}{YY}ME) or ReceiptDate without at least one valid date."
    End If

    m_receiptDateValue = Format$(CDate(maxSerialDate), "yyyy-mm-dd")
    mmStr = Format$(CDate(maxSerialDate), "mm")
    yyStr = Format$(CDate(maxSerialDate), "yy")
    m_batchIdToken = "REFS" & mmStr & yyStr & "ME"
End Sub

'  --------------------------------------------------------------
'  RowIsSilentTrailerAdminReferees -- Finance-appended trailing summary / SUM
'  footer rows (not real referee claims).  Row 291 of the August 2026 exemplar
'  contains only a single non-blank cell (col G = Invoice Total = SUM of the 289
'  data rows); INV number / Cost Code / NEW WORKDAY ID (cols A, L, U) are all
'  blank.  Disposition: silent skip (no counters incremented, no AppendRejectRow
'  entry) -- these are admin housekeeping lines, not claim rows with bad data.
'  --------------------------------------------------------------
Private Function RowIsSilentTrailerAdminReferees( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal invNumberColumn As Long, _
    ByVal costCodeColumn As Long, _
    ByVal newWorkdayIdColumn As Long _
) As Boolean
    Dim invValue As String
    Dim costCodeValue As String
    Dim workdayIdValue As String
    invValue = Trim$(CStr(tableData(rowIndex, invNumberColumn)))
    costCodeValue = Trim$(CStr(tableData(rowIndex, costCodeColumn)))
    workdayIdValue = Trim$(CStr(tableData(rowIndex, newWorkdayIdColumn)))
    If LenB(invValue) = 0 And LenB(costCodeValue) = 0 And LenB(workdayIdValue) = 0 Then
        RowIsSilentTrailerAdminReferees = True
    Else
        RowIsSilentTrailerAdminReferees = False
    End If
End Function

'  --------------------------------------------------------------
'  RowIsSkippableV2 -- analysis-tab missing-field gate (August 2026 refresh).
'  Blank INV number / Details of Invoice / Invoice Date /
'  Submission Date / Cost Code / Workday ID / Invoice Total <= 0.
'  (VENDOR ID is an optional trace-only header since September 2026.)
'  (Approval gate retired September 2026.)  The Workday ID
'  blank check is a pre-condition of the direct-passthrough Supplier
'  resolution (old "#NO MATCH" Creditor-ID sentinel + Supplier Spreadsheet
'  lookup is RETIRED).
'  --------------------------------------------------------------
Private Function RowIsSkippableV2( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal invNumberColumn As Long, _
    ByVal detailsOfInvoiceColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal submissionDateColumn As Long, _
    ByVal costCodeColumn As Long, _
    ByVal invoiceTotalColumn As Long, _
    ByVal newWorkdayIdColumn As Long _
) As Boolean
    RowIsSkippableV2 = _
        LenB(Trim$(CStr(tableData(rowIndex, invNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, detailsOfInvoiceColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, submissionDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, costCodeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, newWorkdayIdColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, invoiceTotalColumn)) <= 0#
End Function

'  --------------------------------------------------------------
'  OptionalCellText -- trimmed cell text for an optional column; returns ""
'  when the column is absent (index 0).
'  --------------------------------------------------------------
Private Function OptionalCellText( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long _
) As String
    If columnIndex > 0 Then OptionalCellText = Trim$(CStr(tableData(rowIndex, columnIndex)))
End Function

'  --------------------------------------------------------------
'  SafeFindOptionalColumn -- wraps FindColumnIndexInRow for optional headers
'  (returns 0 if header not present rather than raising an error).
'  --------------------------------------------------------------
Private Function SafeFindOptionalColumn( _
    ByVal sourceTable As Variant, _
    ByVal headerRowIndex As Long, _
    ByVal headerName As String _
) As Long
    On Error Resume Next
    SafeFindOptionalColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, headerName)
    If Err.Number <> 0 Then
        SafeFindOptionalColumn = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function
