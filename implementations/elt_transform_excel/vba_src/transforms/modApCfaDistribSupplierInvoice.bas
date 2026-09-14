Attribute VB_Name = "modApCfaDistribSupplierInvoice"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_cfa_distributions.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: Finance Master sheet, header=5, data=6 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).
'  Historical hard-coded constants retained for documentation audit only; NOT used at runtime:
'  Private Const SOURCE_WORKSHEET_NAME As String = "Finance Master sheet"
'  Private Const SOURCE_HEADER_ROW As Long = 5
'  Private Const SOURCE_DATA_START_ROW As Long = 6

'  -- REMOVED per-period hard-coded constants / deprecated supplier-lookup paths (Rules 0/1/2):
'  Private Const SOURCE_COMPANY_CODE As String = "FA"
'  Private Const BATCH_ID_TOKEN As String = "CFAD0826"
'  mSupplierLookup / BuildReferenceLookupFromWorkbook("Supplier") / ResolveSupplierWithCompanyPrecedence
'     ALL DEPRECATED 2026-08-24. Finance Master col U ("Workday ID") is the AP-team-curated
'      Workday Supplier Reference ID, used DIRECTLY (no Creditor-ID spreadsheet lookup).
'  Company code is a constant "FA" (Finance Master B = County FA NAME, not legal-entity code).
'  batchIdToken is DERIVED from Finance Master col K (Batch ID = date)  "CFAD" + MMYY.

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"

Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"

'  --- Profile-local Tax Triple override (shadows modFinanceCommon shared standard-rated) ---
'  AP Supplier-Invoice standard-rated lines use the domestic PURCHASE ref-IDs, NOT the shared
'  GBR_Standard_Rated (which is the Customer-Invoice / AR value). Same local-shadow pattern as
'  ap_travel_places / ap_taxi / ap_stationery. AP confirmed 2026-09-14. Zero-rated + tax point/
'  option defaults still resolve from the shared modFinanceCommon constants.
Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Purchase_Of_Goods_And_Services"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_RECOVERABILITY_STANDARD As String = "Fully_Recoverable"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

' NEW 2026-08-24: Finance-Master-sheet layout constants.
' Kept as Private Consts (not Config-driven) because they are structural change-of-source,
' not month-on-month tuning. Mirror the finance_master_constants block in the profile YAML.
Private Const CONSTANT_COMPANY_CODE As String = "FA"
Private Const BATCH_ID_TOKEN_PREFIX As String = "CFAD"

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
    '  -- Finance Master column indices (cols A..U; header row 5) --
    '  Col A = DocNo / row key (unique per row, 54/54 distinct; used for AU + InvoiceNumber suffix)
    '  B = CFA / County FA name (informational; NOT legal-entity company)
    '  C = Type / always "CFA" (informational)
    '  D = Creditor ID / legacy AMA003 etc. -- RETAINED for skip-row gate + traceability, NOT resolution
    '  E = Document type ("invoice" lowercase; gate checks UPPER == INVOICE)
    '  F = Description - remittance (item description DI; accepts alias Decscription|Description via pipe)
    '  G = Distribution/Cost Code (accounting string: 1077-504007-...)
    '  H = Total (net line amount = PurchasesAmount)
    '  I = Tax amount  (>0 triggers standard-rated tax; conditional_on_tax_amount)
    '  J = Tax detail / POS etc. (informational)
    '  K = Batch ID / DATE (per-file date; used for InvoiceDate U, ReceiptDate AG, AND batchIdToken derivation)
    '  L = Code lookup (FA Grant / EP Workforce -- informational)
    '  M..T = gap cols (empty in source)
    '  U = Workday ID / AP-team-curated Supplier ID (S-XXXXXX -- passthrough to EIB col O, NO lookup)
    Dim docNoColumn As Long                  ' A
    Dim cfaNameColumn As Long                ' B
    Dim typeColumn As Long                   ' C
    Dim creditorIdColumn As Long             ' D (retained for skip-gate + traceability only)
    Dim documentTypeColumn As Long           ' E
    Dim descriptionColumn As Long            ' F  "Description - remittance" (pipe alias to Decscription|Description)
    Dim distributionRefCodeColumn As Long    ' G  Distribution/Cost Code (accounting string)
    Dim purchasesAmountColumn As Long        ' H  Total = PurchasesAmount (net)
    Dim taxAmountColumn As Long              ' I  Tax amount (conditional switch DS/DT/DY)
    Dim taxDetailColumn As Long              ' J  Tax detail
    Dim batchIdDateColumn As Long            ' K  Batch ID = date (InvoiceDate, ReceiptDate, batchIdToken)
    Dim codeLookupColumn As Long             ' L
    Dim codeDescColumn As Long               ' M
    Dim nameOnGpColumn As Long               ' N
    Dim workdayIdColumn As Long              ' U  *** NEW: AP-team-selected Workday Supplier ID (passthrough)
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
    ' Spreadsheet is NO LONGER OPENED and NO lookup is performed against it. Finance Master
    ' col U ("Workday ID") = the AP-team-curated Workday Supplier Reference used DIRECTLY.
    ' The line below intentionally removed:
    '   Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    '   Set mSupplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    '   supplierWorkbook.Close SaveChanges:=False

    '  Source of truth: config/transforms/ap_cfa_distributions.yaml source_layout.required_headers.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApCfaDistribSupplierInvoice", 14)

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

    ' Column A = DocNo / row key has no explicit header (our required-headers scan starts at B).
    ' DocNo A is always the column immediately before the "CFA" header, so DocNo = CFA_col - 1.
    cfaNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "CFA")
    docNoColumn = cfaNameColumn - 1
    typeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Type")
    creditorIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Creditor ID")
    documentTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Document type")
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description - remittance|Decscription|Description")
    distributionRefCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Distribution/Cost Code")
    purchasesAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Total")
    taxAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax amount")
    taxDetailColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax detail")
    batchIdDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Batch ID")
    codeLookupColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Code lookup")
    codeDescColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Code desc")
    nameOnGpColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Name on GP")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday ID")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    '  -- Read BATCH-LEVEL constants / derived tokens from Finance Master + CONSTANTS --
    '  sourceCompanyCode = CONSTANT "FA" (CFA always = THE FA GROUP).
    '  batchIdToken = "CFAD" + MMYY derived from first-valid-row col K (Batch ID = date).
    '  Iterate dataStartRowIndex..UBound to find the first non-trailer, non-skippable row.
    sourceCompanyCode = CONSTANT_COMPANY_CODE
    batchIdToken = vbNullString
    workdayIdBlankRows = 0
    Dim bootstrapRow As Long
    For bootstrapRow = dataStartRowIndex To UBound(sourceTable, 1)
        If RowIsEmptyOrTrailer(sourceTable, bootstrapRow, _
            Array(docNoColumn, creditorIdColumn, documentTypeColumn, descriptionColumn, distributionRefCodeColumn, workdayIdColumn), _
            Array(purchasesAmountColumn), _
            Array(docNoColumn)) Then
            GoTo nextBootstrapRow
        End If
        If RowIsSkippable(sourceTable, bootstrapRow, docNoColumn, creditorIdColumn, _
            documentTypeColumn, descriptionColumn, batchIdDateColumn, _
            distributionRefCodeColumn, purchasesAmountColumn, workdayIdColumn) Then
            GoTo nextBootstrapRow
        End If
        ' First valid row  derive batchIdToken from col K (Batch ID = date)
        Dim batchDateValue As Variant
        batchDateValue = sourceTable(bootstrapRow, batchIdDateColumn)
        batchIdToken = DeriveBatchIdTokenFromDate(batchDateValue)
        Exit For
nextBootstrapRow:
    Next bootstrapRow

    If LenB(batchIdToken) = 0 Then
        Err.Raise vbObjectError + 1912, "modApCfaDistribSupplierInvoice.RunApCfaDistributionsSupplierInvoice", _
            "Finance Master col K (Batch ID = date) was BLANK or unparseable on every non-skippable " _
            & "data row. Cannot build batchIdToken (InvoiceNumber prefix). Expected a date cell " _
            & "e.g. 2026-08-01  derives CFAD0826. The Finance Master Batch ID column is the " _
            & "authoritative source for both InvoiceDate/ReceiptDate AND the invoice-number prefix."
    End If

    mixedBatchRows = 0

    companyId = ResolveMappedCompany(runnerWorkbook, sourceCompanyCode, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        '  -- Blank / trailer row gate (Grand Total / summation rows below the data). --
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, _
            Array(docNoColumn, creditorIdColumn, documentTypeColumn, descriptionColumn, distributionRefCodeColumn, workdayIdColumn), _
            Array(purchasesAmountColumn), _
            Array(docNoColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(docNoColumn, creditorIdColumn, documentTypeColumn, descriptionColumn, distributionRefCodeColumn, workdayIdColumn), _
                Array(purchasesAmountColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 3 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsSkippable(sourceTable, rowIndex, docNoColumn, creditorIdColumn, _
            documentTypeColumn, descriptionColumn, batchIdDateColumn, _
            distributionRefCodeColumn, purchasesAmountColumn, workdayIdColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required Finance Master field (DocNo/A blank / Creditor ID/D blank / " _
                    & "Document type/E UPPER!=INVOICE / Description/F blank / Batch-ID/K (date) blank / " _
                    & "Distribution-Cost-Code/G blank / Total/H=0 / Workday-ID/U blank)", _
                Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
        Else
            '  -- Uniformity check: non-skipped rows must have SAME BatchDate / derived token across batch --
            '    (A Distributions workbook represents ONE period; mixing batch dates inside one
            '     run would produce inconsistent InvoiceNumber prefixes.)
            Dim rowBatchToken As String
            Dim rowBatchDate As Variant
            rowBatchDate = sourceTable(rowIndex, batchIdDateColumn)
            rowBatchToken = DeriveBatchIdTokenFromDate(rowBatchDate)
            If LenB(rowBatchToken) > 0 And rowBatchToken <> batchIdToken Then
                mixedBatchRows = mixedBatchRows + 1
            End If

            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, sourceTable(rowIndex, distributionRefCodeColumn))
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Cost centre unresolved: Distribution/Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                    Trim$(CStr(sourceTable(rowIndex, distributionRefCodeColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
            Else
                spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, sourceTable(rowIndex, distributionRefCodeColumn))
                If LenB(spendCategoryId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Spend Category unresolved: Distribution/Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", Core vs Prism = Core, lowest hierarchy match)", _
                        Trim$(CStr(sourceTable(rowIndex, distributionRefCodeColumn))), _
                        Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                Else
                    '  -- Supplier = WORKDAY ID PASSTHROUGH (Rule 0). NO Supplier Spreadsheet lookup. --
                    supplierId = Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
                    If LenB(supplierId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unresolvedSupplierRows = unresolvedSupplierRows + 1
                        workdayIdBlankRows = workdayIdBlankRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                            "Supplier unresolved: Workday ID (Finance Master col U) was BLANK. " _
                                & "As of 2026-08-24 the AP team curates the Workday Supplier Reference " _
                                & "(S-XXXXXX) directly in Finance Master col U ""Workday ID""; " _
                                & "the Creditor-ID Supplier Spreadsheet lookup is retired. " _
                                & "Creditor ID / col D (for traceability): " _
                                & Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, nameOnGpColumn)))
                    Else
                        If WriteOutputRow( _
                            targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                            docNoColumn, documentTypeColumn, descriptionColumn, _
                            batchIdDateColumn, batchIdDateColumn, distributionRefCodeColumn, _
                            purchasesAmountColumn, taxAmountColumn, _
                            costCenterId, spendCategoryId, supplierId, optionalWorktagBlankRows, _
                            sourceCompanyCode, batchIdToken, creditorIdColumn) _
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
            & " source row(s) with missing DocNo/A, Creditor ID/D, Document type/E (not INVOICE)," _
            & " Description/F, Batch ID/K (date), Distribution/Cost Code/G, Total/H=0, or Workday ID/U blank."
    End If
    If mixedBatchRows > 0 Then
        issueParts.Add CStr(mixedBatchRows) _
            & " non-skipped source row(s) had a Batch-ID/K date whose derived token DIFFERED from " _
            & "the first valid row (expected all rows of a single batch file to share the same " _
            & "col K date  same batchIdToken = " & batchIdToken & ")."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " source row(s) whose Finance Master col U ""Workday ID"" was BLANK. " _
            & "Supplier Spreadsheet / Creditor-ID lookup was RETIRED 2026-08-24; " _
            & "AP team must populate ""Workday ID"" on every row."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) _
            & " source row(s) whose Distribution/Cost Code did not resolve to a Workday cost centre" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) _
            & " source row(s) whose Distribution/Cost Code did not resolve to a Workday spend category" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) _
            & " source row(s) had 1 or more blank optional worktags" _
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
'  DeriveBatchIdTokenFromDate -- turn a Finance Master col K date-cell into "CFAD" + MMYY.
'  INPUT: any date value acceptable to CDate / IsDate (Excel serial-date Double, Date,
'         or date-string "2026-08-01" already coerced to text).
'  OUTPUT: e.g. date = 2026-08-01  "CFAD0826".
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
'     docNoColumn / documentTypeColumn / descriptionColumn /
'     invoiceDateColumn (= batchIdDateColumn, K) /
'     receiptDateColumn (= batchIdDateColumn, K -- same because no separate ReceiptDate in FM) /
'     distributionRefCodeColumn / purchasesAmountColumn / taxAmountColumn
'        -- 1-based column indices into sourceTable (Finance Master).
'     sourceCompanyCode = constant "FA" (not source-driven anymore)
'     batchIdToken = derived from Finance Master col K date on first valid row
'  OUTPUT COLUMN MAPPING (matches mapping: in config/transform_specs/ap_cfa_distributions.yaml -- Submit_Supplier_Invoice_v46.2):
'     B  seq            (sequenceNumber)
'     I  submit         (Y)
'     K  InvoiceNumber  (Rule 8 -- {batchIdToken}-{DocNo/colA}; e.g. "CFAD0826-AFAAug2026FA Grant")
'     R  Company        (mCompanyId resolved from constant "FA"  THE FA GROUP - FA)
'     T  Currency       (GBP default)
'     U  Supplier       (supplierId = Finance Master col U passthrough, NO lookup)
'     AF InvoiceDate    (Rule 5 -- Finance Master K via NormalizeSourceDate; batch-level)
'     AG ReceiptDate    (Rule 6 -- Invoice Received Date, Finance Master K via NormalizeSourceDate; =InvoiceDate)
'     BG SuppliersInvoiceNumber / voucher (Rule 7 -- Finance Master col A; unique per row)
'     EK line number 1  (line Row ID)
'     EP ItemDescription (Rule 4 -- UPPER(TRIM(Finance Master F)))
'     EW Spend Category  (Dist/CostCode low-hierarchy match; on_fail=reject, gated upstream)
'     FE FF FH FK FL  tax conditional_on_tax_amount (Rule 13: TaxAmount/I>0 -> Standard, else Zero)
'     GE Unit Cost / GF Extended Amount (PurchasesAmount Finance Master H)
'     GQ Memo (Rule 14 -- DocNo colA + " " + Distribution/CostCode colG)
'     GR Reporting Type  (second segment first char via ResolveReportingTypeFromFinalCode)
'     GT Cost Centre     (Dist/CostCode low-hierarchy match; on_fail=reject, gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (Event Camp FALearning CommercialPartner Grant Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal docNoColumn As Long, _
    ByVal documentTypeColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal receiptDateColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal purchasesAmountColumn As Long, _
    ByVal taxAmountColumn As Long, _
    ByVal costCenterId As String, _
    ByVal spendCategoryId As String, _
    ByVal supplierId As String, _
    ByRef optionalWorktagBlankRows As Long, _
    ByVal sourceCompanyCode As String, _
    ByVal batchIdToken As String, _
    Optional ByVal creditorIdColumn As Long = 0 _
) As Boolean
    Dim amount As Double
    Dim taxAmount As Double
    Dim distributionRefCodeValue As Variant
    Dim docNoValue As String
    Dim invoiceNumberValue As String
    Dim descriptionValue As String
    Dim invoiceDateValue As String
    Dim receiptDateValue As String
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
    Dim anyWorktagBlank As Boolean
    Dim taxIsStandard As Boolean

    docNoValue = Trim$(CStr(sourceTable(sourceRow, docNoColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, descriptionColumn))))   ' Rule 4
    amount = ResolveAmount(sourceTable(sourceRow, purchasesAmountColumn))
    taxAmount = ResolveAmount(sourceTable(sourceRow, taxAmountColumn))
    distributionRefCodeValue = sourceTable(sourceRow, distributionRefCodeColumn)
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))     ' Rule 5 -- FM col K
    receiptDateValue = NormalizeSourceDate(sourceTable(sourceRow, receiptDateColumn))     ' Rule 6 -- FM col K (=invoiceDate)

    ' Rule 13 -- Tax code conditional switch: TaxAmount>0 -> Standard-rated, else Zero-rated
    taxIsStandard = (taxAmount > 0#)

    ' Rule 8 -- Invoice Number = {batchIdToken}-{DocNo/colA}.
    ' batchIdToken derived on FIRST VALID DATA ROW from Finance Master K (date)  "CFADMMYY",
    ' validated uniform across batch. DocNo/colA is per-row unique (54/54 in Aug exemplar).
    invoiceNumberValue = batchIdToken & "-" & docNoValue

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
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue              ' Rule 5
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = receiptDateValue             ' Rule 6 (=K, same as U)
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = docNoValue                   ' Rule 7 -- DocNo/colA
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue    ' Row ID = 1..N within invoice group
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue    ' Line Order = 1..N within invoice group (mirror of EK)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue            ' Rule 4 -- UPPER()
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    ' Rule 13 -- Tax DS/DT/DY per-row conditional_on_tax_amount (FM col I)
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
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = docNoValue & " " & CStr(distributionRefCodeValue)   ' Rule 14 -- colA + " " + colG
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
'  RowIsSkippable -- Finance-Master-specific gate (matches skip_rule.drop_when in spec YAML).
'     docNoColumn = A, creditorIdColumn = D, documentTypeColumn = E,
'     descriptionColumn = F, batchIdDateColumn = K (date),
'     distributionRefCodeColumn = G, purchasesAmountColumn = H,
'     workdayIdColumn = U (NEW mandatory gate 2026-08-24).
'     Company code is constant "FA", not present in FM, so no column-level gate.
'  --------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal docNoColumn As Long, _
    ByVal creditorIdColumn As Long, _
    ByVal documentTypeColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal batchIdDateColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal purchasesAmountColumn As Long, _
    ByVal workdayIdColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, docNoColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, creditorIdColumn)))) = 0 Or _
        StrComp(UCase$(Trim$(CStr(tableData(rowIndex, documentTypeColumn)))), "INVOICE", vbBinaryCompare) <> 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, descriptionColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, batchIdDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, distributionRefCodeColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, purchasesAmountColumn)) = 0# Or _
        LenB(Trim$(CStr(tableData(rowIndex, workdayIdColumn)))) = 0
End Function
