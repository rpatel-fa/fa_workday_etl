Attribute VB_Name = "modApStationerySupplierInvoice"
Option Explicit

' Source of truth: config/transform_specs/ap_stationery.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: Proactis Upload tab, header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).
'  Historical hard-coded constants retained for documentation audit only; NOT used at runtime:
'  Private Const SOURCE_WORKSHEET_NAME As String = "Proactis Upload"
'  Private Const SOURCE_HEADER_ROW As Long = 1
'  Private Const SOURCE_DATA_START_ROW As Long = 2

'  -- NOTE: no fixed SOURCE_SUPPLIER_REFERENCE_ID; supplier is per-row split. --
'  No Private Const SOURCE_COMPANY_CODE. Company is DERIVED PER ROW from
'  SupplierID LEFT-segment split. OUTPUT_COMPANY_DEFAULT_FALLBACK is used only
'  when the runner's company-mapping rows have no entry for a given company
'  prefix (belt-and-braces last-resort entity name).
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"
Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

'  --- Profile-local Tax Triple overrides (hybrid scope; shadows modFinanceCommon.bas shared defaults) ---
'  These values come from ap_stationery.yaml transform_specs per_row_tax sets (P20/P0) and MUST
'  NOT be replaced with the shared Public Const TAX_APPLICABILITY_STANDARD / TAX_CODE_STANDARD from
'  modFinanceCommon.  Shared constants use legacy GBR_Standard_Rated / Recoverable_Allocate which
'  AP confirmed incorrect against Will's hand-edited Aug Ambassador Taxi EIB. Supplier Invoice
'  (Purchasing) uses the domestic Purchase ref-IDs vs the generic Standard_Rated that Customer
'  Invoice modules use.
Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Purchase_Of_Goods_And_Services"
Private Const TAX_APPLICABILITY_ZERO As String = "GBR_Outside_of_Scope"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_CODE_ZERO As String = "GBR_Zero_Rated"
Private Const TAX_RECOVERABILITY_STANDARD As String = "Fully_Recoverable"
Private Const TAX_RECOVERABILITY_ZERO As String = "Nonrecoverable_Allocate"

'  9 worktag lookups: 2 mandatory (Cost Center, Spend Category) + 7 optional.
'  Module-scope so WriteOutputRow helper can use them without being passed as params.
Private mSupplierLookup As Collection
Private mSpendCategoryLookup As Collection
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' ====================================================================
'  Entry point (called by the Runner via reflection on selected transform key).
' ====================================================================
Public Sub RunApStationerySupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim mappingSpreadsheetPath As String
    Dim mappingSheetName As String
    Dim supplierSpreadsheetPath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim mappingWorkbook As Workbook
    Dim supplierWorkbook As Workbook
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
    Dim unknownTaxCodeRows As Long
    Dim optionalWorktagBlankRows As Long
    Dim rowIndex As Long

    '  -- Proactis Upload column indices (A..P = 16 cols) --
    Dim templateColumn As Long                ' A  Template (Proactis tag, unused for EIB)
    Dim supplierIdColumn As Long              ' B  SupplierID = "{Company}-{CreditorID}"
    Dim invoiceNumberColumn As Long           ' C  InvoiceNumber (per-row = EIB G & AU)
    Dim invoiceDateColumn As Long             ' D  InvoiceDate -- per-row via NormalizeSourceDate
    Dim receivedDateColumn As Long            ' E  ReceivedDate -- per-row via NormalizeSourceDate
    Dim currencyColumn As Long                ' F  Currency -- per-row direct (fallback GBP)
    Dim trayColumn As Long                    ' G  Tray (Proactis internal, unused)
    Dim createdByColumn As Long               ' H  CreatedBy (Proactis internal, unused)
    Dim poNumberColumn As Long                ' I  PONumber -- optional
    Dim itemCodeColumn As Long                ' J  ItemCode -- header bleed guard only
    Dim costCodeColumn As Long                ' K  "New Cost Code" primary / "Final Cost Code" fallback
    Dim netColumn As Long                     ' L  "Net" primary / "Net Amount" fallback
    Dim taxColumn As Long                     ' M  Tax
    Dim grossColumn As Long                   ' N  Gross
    Dim descriptionColumn As Long             ' O  Description -- UCase for DI
    Dim taxCodeColumn As Long                 ' P  TaxCode -- per-row switch (P20/P0/PZ)

    Dim companyCodePerRow As String
    Dim creditorIdPerRow As String
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
    Dim consecutiveEmptyRows As Long
    Dim worksheetNameStr As String
    Dim worksheetResolutionStr As String
    Dim sheetToResolve As String
    Dim sourceHeaderRow As Long
    Dim sourceDataStartRow As Long
    consecutiveEmptyRows = 0

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

    '  Supplier lookup -- shared across all rows; 2-part precedence applied per-row.
    Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    Set mSupplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    supplierWorkbook.Close SaveChanges:=False
    Set supplierWorkbook = Nothing

    '  16 Proactis Upload headers A..P. Cols K + L use pipe-delimited alias groups
    '  to tolerate supplier variations (see modSourceTable.SplitHeaderAliases /
    '  FindColumnIndexInRow for the alias-resolution mechanism). Order in each
    '  alias group = lookup priority.
    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApStationerySupplierInvoice", 16)

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
    schemaMsg = ValidateSourceWorkbookSchema( sheetToResolve, sourceTable, sourceHeaderRow, sourceDataStartRow, requiredHeaders, , schemaHeaderRow, schemaDiag )
    If LenB(schemaMsg) > 0 Then
        Err.Raise vbObjectError + 1910, "modApStationerySupplierInvoice.RunApStationerySupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    '  Resolve column indices -- cols K + L use pipe-delimited alias groups.
    templateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Template")
    supplierIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "SupplierID")
    invoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "InvoiceNumber")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "InvoiceDate")
    receivedDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "ReceivedDate")
    currencyColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Currency")
    trayColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tray")
    createdByColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "CreatedBy")
    poNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "PONumber")
    itemCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "ItemCode")
    '  Col K: "New Cost Code" (3 of 4 suppliers) OR "Final Cost Code" (Addison Lee).
    '  FindColumnIndexInRow accepts the pipe-delimited alias string directly and
    '  returns the FIRST matching column (priority = alias list order).
    costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "New Cost Code|Final Cost Code")
    '  Col L: "Net" (3 of 4 suppliers) OR "Net Amount" (E-Courier).
    netColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net|Net Amount")

    taxColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax")
    grossColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Gross")
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description")
    taxCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "TaxCode")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW
      Set mGrpKey = New Collection
      Set mGrpLine = New Collection
      mGrpNextKey = 0

    '  NOTE: no module-scope mCompanyId assignment here. Company is resolved PER ROW
    '  from the SupplierID split -- see WriteOutputRow / per-row split block below.
    '  OUTPUT_COMPANY_DEFAULT_FALLBACK (the entity name literal) is used as the
    '  belt-and-braces fallback only when the runner's company-mapping rows do NOT
    '  contain an entry for a given row's companyCodePerRow derived from the split.

      '  CHANGE 2026-09-08 (Ambassador Taxi Invoice 496 multi-line: same invoice number
      '  across multiple rows -> Workday duplicate-ref rejection):
      '  Phase 1 header-consistency PREPASS over valid non-skipped rows. Produces 4 pure-VBA
      '  Collections (cross-OS safe, no Scripting.Dictionary):
      '    (A) groupCompanyCaptured    -> key = raw suppliersInvoiceNumberValue (source C DIRECT), item = CStr Rule 1 Company for guard
      '    (B) groupSupplierCaptured   -> key = raw suppliersInvoiceNumberValue (source C DIRECT), item = CStr Rule 2 Supplier for guard
      '    (C) groupCurrencyCaptured   -> key = raw suppliersInvoiceNumberValue (source C DIRECT), item = CStr Rule 12 Currency for guard
      '    (D) groupInvDateCaptured    -> key = raw suppliersInvoiceNumberValue (source C DIRECT), item = CStr Rule 4 InvoiceDate for guard
      '  NOTE: group key = RAW source InvoiceNumber passthrough (NOT the K-composite CreditorID-YYMM-RawInvNum).
      '  This way group membership is stable, plaintext-readable, and a 1:1 match to the upstream group key.
      '  First non-skipped appearance of a group populates all capturers;
      '  subsequent rows of same group assert header consistency (4
    '  header fields: Company/Supplier/Currency/InvoiceDate -- ReceivedDate excluded deliberately
    '  because same Workday invoice can legitimately span multiple receipt months when split
    '  across cost codes at line level; B+EM grouping is authoritative for Workday multi-line).
    '  If any of the 4 strict fields diverges within a group we raise a build ERROR before
    '  any EIB rows are written.
    Dim groupCompanyCaptured As New Collection    ' key=InvoiceNumberValue, item=CStr company
    Dim groupSupplierCaptured As New Collection   ' key=InvoiceNumberValue, item=CStr supplier
    Dim groupCurrencyCaptured As New Collection    ' key=InvoiceNumberValue, item=CStr currency
    Dim groupInvDateCaptured As New Collection     ' key=InvoiceNumberValue, item=CStr invdate
    Dim prepassMsg As String
    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, Array(supplierIdColumn, invoiceNumberColumn, costCodeColumn, descriptionColumn), Array(netColumn, grossColumn), Array(descriptionColumn, supplierIdColumn, invoiceNumberColumn)) Then
        ElseIf RowIsSkippable(sourceTable, rowIndex, supplierIdColumn, invoiceNumberColumn, costCodeColumn, invoiceDateColumn, taxCodeColumn, descriptionColumn, itemCodeColumn, netColumn, taxColumn, grossColumn) Then
        Else
            Dim pSidRaw As String
            Dim pHyphenPos As Long
            Dim pCcPerRow As String
            Dim pCrPerRow As String
            Dim pRCompanyId As String
            Dim pRSupplierId As String
            Dim pNAmt As Double
            Dim pGAmt As Double
            Dim pTAmt As Double
            Dim pINum As String
            Dim pIDt As String
            Dim pRDt As String
            Dim pCVal As String
            Dim pTcStr As String
            pSidRaw = Trim$(CStr(sourceTable(rowIndex, supplierIdColumn)))
            pHyphenPos = InStr(pSidRaw, "-")
            If pHyphenPos = 0 Then GoTo NextPrepassRow
            pCcPerRow = Left$(pSidRaw, pHyphenPos - 1)
            pCrPerRow = Mid$(pSidRaw, pHyphenPos + 1)
            If StrComp(pCcPerRow, "WNSL", vbTextCompare) = 0 Then GoTo NextPrepassRow
            pRCompanyId = ResolveMappedCompany(runnerWorkbook, pCcPerRow, OUTPUT_COMPANY_DEFAULT_FALLBACK)
            pRSupplierId = ResolveSupplierWithCompanyPrecedence(mSupplierLookup, pCrPerRow, pCcPerRow)
            If LenB(pRSupplierId) = 0 Then GoTo NextPrepassRow
            Dim pCcRaw As String
            pCcRaw = Trim$(CStr(sourceTable(rowIndex, costCodeColumn)))
            If LenB(ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, pCcRaw)) = 0 Then GoTo NextPrepassRow
            If LenB(ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, pCcRaw)) = 0 Then GoTo NextPrepassRow
            pTcStr = UCase$(Trim$(CStr(sourceTable(rowIndex, taxCodeColumn))))
            If pTcStr <> "P20" And pTcStr <> "P0" And pTcStr <> "PZ" Then GoTo NextPrepassRow
            pNAmt = ResolveAmount(sourceTable(rowIndex, netColumn))
            pGAmt = ResolveAmount(sourceTable(rowIndex, grossColumn))
            pTAmt = ResolveAmount(sourceTable(rowIndex, taxColumn))
            If pNAmt = 0# Then pNAmt = pGAmt - pTAmt
            If pNAmt < 0# Then pNAmt = -pNAmt
            pINum = Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
            pIDt = NormalizeSourceDate(sourceTable(rowIndex, invoiceDateColumn))
            pRDt = NormalizeSourceDate(sourceTable(rowIndex, receivedDateColumn))
            If LenB(pRDt) = 0 Then pRDt = pIDt
            pCVal = Trim$(CStr(sourceTable(rowIndex, currencyColumn)))
              If LenB(pCVal) = 0 Then pCVal = OUTPUT_DEFAULT_CURRENCY
              If CollectionKeyExists(groupCompanyCaptured, pINum) = False Then
                groupCompanyCaptured.Add pRCompanyId, pINum
                groupSupplierCaptured.Add pRSupplierId, pINum
                groupCurrencyCaptured.Add pCVal, pINum
                groupInvDateCaptured.Add pIDt, pINum
            Else
                prepassMsg = vbNullString
                If StrComp(CStr(groupCompanyCaptured(pINum)), pRCompanyId, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "Company(" & CStr(groupCompanyCaptured(pINum)) & " vs " & pRCompanyId & "); "
                If StrComp(CStr(groupSupplierCaptured(pINum)), pRSupplierId, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "Supplier(" & CStr(groupSupplierCaptured(pINum)) & " vs " & pRSupplierId & "); "
                If StrComp(CStr(groupCurrencyCaptured(pINum)), pCVal, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "Currency(" & CStr(groupCurrencyCaptured(pINum)) & " vs " & pCVal & "); "
                If StrComp(CStr(groupInvDateCaptured(pINum)), pIDt, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "InvoiceDate(" & CStr(groupInvDateCaptured(pINum)) & " vs " & pIDt & "); "
                If LenB(prepassMsg) > 0 Then
                    Err.Raise vbObjectError + 1912, "modApStationerySupplierInvoice.RunApStationerySupplierInvoice Phase1GroupPrepass", _
                        "Header Consistency Guard (4-field) failed. Same InvoiceNumber '" & pINum & "' source rows have diverged header fields. " & prepassMsg & "ReceivedDate intentionally excluded (legit per-group variance OK). Fix source data FIRST before re-running the Runner."
                End If
            End If
        End If
NextPrepassRow:
    Next rowIndex

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime

        If RowIsEmptyOrTrailer(sourceTable, rowIndex, Array(supplierIdColumn, invoiceNumberColumn, costCodeColumn, descriptionColumn), Array(netColumn, grossColumn), Array(descriptionColumn, supplierIdColumn, invoiceNumberColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, Array(supplierIdColumn, invoiceNumberColumn, costCodeColumn, descriptionColumn), Array(netColumn, grossColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 3 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsSkippable(sourceTable, rowIndex, supplierIdColumn, invoiceNumberColumn, costCodeColumn, invoiceDateColumn, taxCodeColumn, descriptionColumn, itemCodeColumn, netColumn, taxColumn, grossColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Missing required Proactis Upload field (InvoiceNumber blank / " & "SupplierID blank/no-hyphen / Cost Code blank / " & "InvoiceDate blank / TaxCode blank / ItemCode literal header bleed / " &  _
                "Net blank/zero AND Gross-Tax blank/zero).", Trim$(CStr(sourceTable(rowIndex, supplierIdColumn))), Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
        Else
            '  -- Split SupplierID on first hyphen ? company + creditor PER ROW. --
            Dim sidRaw As String
            Dim hyphenPos As Long
            sidRaw = Trim$(CStr(sourceTable(rowIndex, supplierIdColumn)))
            hyphenPos = InStr(sidRaw, "-")
            companyCodePerRow = Left$(sidRaw, hyphenPos - 1)
            creditorIdPerRow = Mid$(sidRaw, hyphenPos + 1)

            '  DELIBERATE EXCLUSION: WNSL Addison Lee handled separately.
            If StrComp(companyCodePerRow, "WNSL", vbTextCompare) = 0 Then
                rowsSkipped = rowsSkipped + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "WNSL-prefixed SupplierID excluded from GENERIC ap_stationery " & "transform (Addison Lee WNSL handled in separate dedicated " & "transform / separate session per scope rules).", sidRaw, Trim$(CStr(sourceTable(rowIndex,  _
                    descriptionColumn)))
            Else
                '  Resolve company per row via runner company-mapping (entity=company),
                '  belt-and-braces fall back to the entity literal from
                '  OUTPUT_COMPANY_DEFAULT_FALLBACK if the runner mapping has no match.
                Dim resolvedCompanyId As String
                resolvedCompanyId = ResolveMappedCompany(runnerWorkbook, companyCodePerRow, OUTPUT_COMPANY_DEFAULT_FALLBACK)

                '  Resolve supplier via 2-part precedence (per transform_spec parties.block).
                supplierId = ResolveSupplierWithCompanyPrecedence( mSupplierLookup, creditorIdPerRow, companyCodePerRow)
                If LenB(supplierId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unresolvedSupplierRows = unresolvedSupplierRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Supplier unresolved: CreditorID did not resolve via the Supplier Spreadsheet" & " (Reference ID -> Supplier ID, two-part company precedence:" & " CreditorID ""-"" Company first, then bare CreditorID." & " Company=" &  _
                        companyCodePerRow & ".)", creditorIdPerRow, "SupplierID raw=" & sidRaw
                Else
                    '  Resolve mandatory worktags first (Cost Center + Spend Category
                    '  from Rule-7 cost-code accounting string via Mapping MASTER
                    '  lowest-hierarchy match; reject entire row on miss).
                    Dim costCodeRaw As String
                    costCodeRaw = Trim$(CStr(sourceTable(rowIndex, costCodeColumn)))

                    costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( costCenterLookup, costCodeRaw)
                    If LenB(costCenterId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unmappedCostCenterRows = unmappedCostCenterRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Cost centre unresolved: Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", " & "Core vs Prism = Core, lowest hierarchy match)", costCodeRaw, Trim$(CStr(sourceTable(rowIndex,  _
                            descriptionColumn)))
                    Else
                        spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( mSpendCategoryLookup, costCodeRaw)
                        If LenB(spendCategoryId) = 0 Then
                            rowsSkipped = rowsSkipped + 1
                            unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Spend Category unresolved: Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", " & "Core vs Prism = Core, lowest hierarchy match)", costCodeRaw,  _
                                Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                        Else
                            '  TaxCode classification (Rule 11): P20 / P0 / PZ.
                            '  Any other value ? reject row (spec on_unknown_when_value:
                            '  reject_and_report).
                            Dim taxApply As String
                            Dim taxCod As String
                            Dim taxRecov As String
                            Dim tcRaw As String
                            tcRaw = UCase$(Trim$(CStr(sourceTable(rowIndex, taxCodeColumn))))
                            If tcRaw = "P20" Then
                                taxApply = TAX_APPLICABILITY_STANDARD
                                taxCod = TAX_CODE_STANDARD
                                taxRecov = TAX_RECOVERABILITY_STANDARD
                            ElseIf tcRaw = "P0" Or tcRaw = "PZ" Then
                                taxApply = TAX_APPLICABILITY_ZERO
                                taxCod = TAX_CODE_ZERO
                                taxRecov = TAX_RECOVERABILITY_ZERO
                            Else
                                taxApply = vbNullString
                                taxCod = vbNullString
                                taxRecov = vbNullString
                            End If
                            If LenB(taxApply) = 0 Then
                                rowsSkipped = rowsSkipped + 1
                                unknownTaxCodeRows = unknownTaxCodeRows + 1
                                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Unknown TaxCode '" & tcRaw & "' -- expected P20/P0/PZ. " & "Per spec on_unknown_when_value = reject_and_report.", "TaxCode raw=" & tcRaw, Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                            Else
                                If WriteOutputRow( targetSheet, targetRow, sourceTable, rowIndex, resolvedCompanyId, supplierId, taxApply, taxCod, taxRecov, invoiceNumberColumn, invoiceDateColumn, receivedDateColumn, descriptionColumn, costCodeColumn, netColumn, taxColumn,  _
                                      grossColumn, currencyColumn, poNumberColumn, supplierIdColumn, costCenterId, spendCategoryId, optionalWorktagBlankRows) Then
                                    rowsProcessed = rowsProcessed + 1
                                    targetRow = targetRow + 1
                                Else
                                    rowsSkipped = rowsSkipped + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Proactis Upload required field (Net/Gross " & "blank/zero / InvoiceNumber blank / SupplierID blank/no-hyphen / " & "Cost Code blank / InvoiceDate blank / TaxCode blank / " &  _
            "ItemCode literal 'ItemCode' header-bleed)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) & " source row(s) whose CreditorID did not resolve to a Workday Supplier ID" & " via the Supplier Spreadsheet (2-part company precedence)."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) & " source row(s) whose Cost Code did not resolve to a Workday cost centre" & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) & " source row(s) whose Cost Code did not resolve to a Workday spend category" & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unknownTaxCodeRows > 0 Then
        issueParts.Add "Skipped " & CStr(unknownTaxCodeRows) & " source row(s) with unknown TaxCode (only P20 / P0 / PZ accepted)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) & " source row(s) had 1 or more blank optional worktags" & " (Event / Camp / FA Learning / Commercial Partner / Grant / Program / Project)" & " -- Cost Code did not fully map to all 7 worktag hierarchies."
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
    If Not supplierWorkbook Is Nothing Then supplierWorkbook.Close SaveChanges:=False
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
    Err.Raise failureNumber, "modApStationerySupplierInvoice.RunApStationerySupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row.
'  INPUTS:
'     resolvedCompanyId, supplierId -- per-row SupplierID split outcomes.
'     taxApply/taxCod/taxRecov -- per-row TaxCode switch results.
   '     All other Long params are 1-based column indices into sourceTable.
'  OUTPUT COLUMN MAPPING (matches mapping: in transform_specs/ap_stationery.yaml -- Submit_Supplier_Invoice_v46.2):
'     B  Spreadsheet Key (Rule 15: per-group SHARED key, all rows of same InvoiceNumber
'                          write the SAME integer; assigned in Phase 1 prepass 1..M)
'     I  submit         (Y)
'     K  InvoiceNumber  (Rule 3 COMPOSITE -- CreditorID-YYMM-RawSourceInvNum, Workday UNIQUE key;
'                          differs from BG which remains the raw supplier-issued passthrough)
'     BG SuppliersInvoiceNumber (Rule 6 -- EXTERNAL-FACING raw passthrough; NOT a mirror of K)
'     R  Company        (Rule 1 -- per-row SupplierID split LEFT segment)
'     T  Currency       (Rule 12 -- col F DIRECT; fallback = GBP if blank)
'     U  Supplier       (Rule 2 -- per-row SupplierID split RIGHT segment 2-part precedence)
'     AF InvoiceDate    (Rule 4 -- col D per-row via NormalizeSourceDate, coerced @ format)
'     AG ReceiptDate    (Rule 5 -- col E per-row via NormalizeSourceDate; fallback=InvoiceDate if blank)
'     AS Control Amount Total (Rule 17: Leave BLANK  --  AP confirmed no value required here)
'     BG SuppliersInvoiceNumber (Rule 6 -- same value as K InvoiceNumber, bitwise mirror)
'     BH PONumber       (Rule 13 -- External PO Number, col I DIRECT if non-blank else leave empty)
'     EL Supplier Invoice Line ID (Leave EMPTY; Workday auto-assigns from Line Order EM)
'     EM Line Order     (Rule 16: per-group 1..N reset counter; increment AFTER write)
'     EK Row ID         (Rule 19: EXACT MIRROR of EM Line Order; same 1..N per-group counter)
'     EP ItemDescription (Rule 8 -- UCASE(TRIM(col O Description)))
'     EW Spend Category  (Rule 7 Seg2 -- CostCode via mapping MASTER, gated upstream)
'     FE Applicability / FF Code / FH date-type / FK Recoverability / FL CALC_TAX_DUE
'     GE Unit Cost / GF Extended Amount (Rule 10 Net for BOTH  --  GE and GF must be identical; CALC_TAX_DUE lets Workday re-derive VAT)
'     GQ Memo (Rule 9 -- InvoiceNumber + ' ' + SupplierID + ' ' + CostCodeNormalized)
'     GR Reporting Type  (second segment first char via ResolveReportingTypeFromFinalCode)
'     GT Cost Centre     (CostCode via mapping MASTER, gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (Event Camp FALearning CommercialPartner Grant Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRow( ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal resolvedCompanyId As String, ByVal supplierId As String, ByVal taxApply As String, ByVal taxCod As String,  _
    ByVal taxRecov As String, ByVal invoiceNumberColumn As Long, ByVal invoiceDateColumn As Long, ByVal receivedDateColumn As Long, ByVal descriptionColumn As Long, ByVal costCodeColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, ByVal grossColumn As Long,  _
      ByVal currencyColumn As Long, ByVal poNumberColumn As Long, ByVal supplierIdColumn As Long, ByVal costCenterId As String, ByVal spendCategoryId As String, ByRef optionalWorktagBlankRows As Long) As Boolean
    Dim netAmt As Double
    Dim grossAmt As Double
    Dim taxAmt As Double
    Dim costCodeRaw As String
    Dim costCodeNormalized As String
    Dim suppliersInvoiceNumberValue As String     ' Rule 6: raw source col C passthrough for BG (external-facing)
    Dim invoiceNumberCompositeValue As String     ' Rule 3: K composite = CreditorID-YYMM-RawInvNum (Workday-unique)
    Dim descriptionValue As String
    Dim invoiceDateValue As String
    Dim receiptDateValue As String
    Dim currencyValue As String
    Dim poValue As String
    Dim supplierIdRaw As String
    Dim supplierCreditorIdBare As String          ' Rule 1 split RIGHT segment (ADD001 / AMB001 / DHL001 / ECO003)
    Dim invoiceDateYYMM As String                 ' 4-char 2digitYear + 2digitMonth (2607)
    Dim hyphenPos As Long
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
      Dim anyWorktagBlank As Boolean
      Dim spreadsheetKeyValue As Long
      Dim lineOrderValue As Long

    '  Rule 6: Suppliers Invoice Number (BG) = raw passthrough source col C DIRECT.
    suppliersInvoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, descriptionColumn))))

    '  Rule 10: Net primary (col L). If blank/zero ? fallback compute
    '  Gross - Tax from cols N + M.
    netAmt = ResolveAmount(sourceTable(sourceRow, netColumn))
    grossAmt = ResolveAmount(sourceTable(sourceRow, grossColumn))
    taxAmt = ResolveAmount(sourceTable(sourceRow, taxColumn))
    If netAmt = 0# Then netAmt = grossAmt - taxAmt
    If netAmt < 0# Then netAmt = -netAmt

    '  Rule 7: cost code separator normalization (any space/dash/dot/underscore
    '  runs ? single dash, then trim surrounding dashes).
    costCodeRaw = Trim$(CStr(sourceTable(sourceRow, costCodeColumn)))
    costCodeNormalized = NormalizeCostCodeSeparators(costCodeRaw)

    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    receiptDateValue = NormalizeSourceDate(sourceTable(sourceRow, receivedDateColumn))
    '  Rule 5: ReceiptDate (AG)  --  CHANGED 2026-09-09. AP confirmed (vs Will's hand-edited Aug EIB)
    '  that AG SHOULD BE LEFT BLANK. Do NOT fall back to InvoiceDate; Will's file has all 14 AG cells = Empty.

    currencyValue = Trim$(CStr(sourceTable(sourceRow, currencyColumn)))
    If LenB(currencyValue) = 0 Then currencyValue = OUTPUT_DEFAULT_CURRENCY

    poValue = Trim$(CStr(sourceTable(sourceRow, poNumberColumn)))
    supplierIdRaw = Trim$(CStr(sourceTable(sourceRow, supplierIdColumn)))

    '  Rule 3: Build K composite = CreditorID_RightSegment + "-" + YYMM(InvoiceDate) + "-" + SuppliersInvoiceNumber.
    '  K remains COMPOSITE (per user instruction  --  "we want to ensure the invoice number is unique").
    '  BG (Suppliers Invoice Number) stays RAW PASHROUGH (col C direct, external-facing).
    '  Example: Ambassador Invoice 496 dated 2026-07-01 -> K=AMB001-2607-496, BG=496.
    '  (Will's hand-edited Aug EIB had K=495/BG=495 identical  --  but AP reverted: keep composite K for
    '   Workday uniqueness guard across suppliers/months; BG is still raw passthrough for AP PDF reconciliation.)
    hyphenPos = InStr(supplierIdRaw, "-")
    If hyphenPos > 0 Then
        supplierCreditorIdBare = Mid$(supplierIdRaw, hyphenPos + 1)
    Else
        supplierCreditorIdBare = supplierIdRaw
    End If
    If Len(invoiceDateValue) >= 7 Then
        invoiceDateYYMM = Mid$(invoiceDateValue, 3, 2) & Mid$(invoiceDateValue, 6, 2)
    Else
        invoiceDateYYMM = Format$(sourceTable(sourceRow, invoiceDateColumn), "YYMM")
    End If
    invoiceNumberCompositeValue = supplierCreditorIdBare & "-" & invoiceDateYYMM & "-" & suppliersInvoiceNumberValue

    '  Optional 7 worktags via CostCodeNormalized lowest-hierarchy mapping match.
    eventId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mEventLookup, costCodeNormalized)
    campId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCampLookup, costCodeNormalized)
    faLearningId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mFaLearningLookup, costCodeNormalized)
    commercialPartnerId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCommercialPartnerLookup, costCodeNormalized)
    grantId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mGrantLookup, costCodeNormalized)
    programId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProgramLookup, costCodeNormalized)
    projectId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProjectLookup, costCodeNormalized)

    anyWorktagBlank = (LenB(eventId) = 0) Or (LenB(campId) = 0) Or (LenB(faLearningId) = 0) Or (LenB(commercialPartnerId) = 0) Or (LenB(grantId) = 0) Or (LenB(programId) = 0) Or (LenB(projectId) = 0)
    If anyWorktagBlank Then
        optionalWorktagBlankRows = optionalWorktagBlankRows + 1
    End If

      spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, suppliersInvoiceNumberValue, mGrpNextKey, lineOrderValue)

      targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue         ' Rule 15 (per-invoice shared key from AssignInvoiceLineGroup)
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT                ' HOUSE DEFAULT Add Only = Y
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberCompositeValue  ' Rule 3 COMPOSITE (CreditorID-YYMM-RawInvNum, Workday UNIQUE -- retained per user instruction for uniqueness guarantee across suppliers/months)
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = resolvedCompanyId            ' Rule 1 (split LEFT)
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = currencyValue                ' Rule 12 (col F DIRECT, fallback GBP)
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId                   ' Rule 2 (split RIGHT 2-part)
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue            ' Rule 4
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = Empty                      ' Rule 5 -- CHANGED 2026-09-09: Leave BLANK per Will's Aug EIB (all 14 AG = None). Do NOT populate with source ReceivedDate.
    targetSheet.Cells(targetRow, ColumnIndex("AS")).Value = Empty                      ' Rule 17 (Control Amount Total -- leave BLANK per AP confirmation)
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = suppliersInvoiceNumberValue ' Rule 6 (EXTERNAL-FACING bare passthrough -- raw col C InvoiceNumber. Differs from K intentionally: K = composite Workday key, BG = visible supplier-issued invoice ref for AP PDF cross-check)
    If LenB(poValue) > 0 Then targetSheet.Cells(targetRow, ColumnIndex("BH")).Value = poValue   ' Rule 13 (optional)
    '  Supplier Invoice Line ID col EL (142): Intentionally BLANK per spec.
    '  Leave EMPTY -- Workday auto-allocates internal line IDs from Line Order col EM.
      targetSheet.Cells(targetRow, ColumnIndex("EL")).Value = Empty
      targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue               ' Rule 16 (per-group 1..N from AssignInvoiceLineGroup)
      targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue               ' Rule 19 (exact mirror of EM from the same shared-helper output)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue            ' Rule 8 (UCase trim)
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = taxApply
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = taxCod
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = taxRecov
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = netAmt                      ' Rule 10 (Unit Cost = net)
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = netAmt                      ' Rule 10 (Extended Amount = net SAME as GE  --  AP confirmed BOTH must match; CALC_TAX_DUE re-derives VAT)
    '  Rule 9: Memo = [K-composite or raw InvoiceNumber-visible] + ' ' + SupplierID + ' ' + CostCodeNormalized
    '  Keep the visible portion readable to AP reviewers; include K-composite so Workday internal id is traceable.
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = invoiceNumberCompositeValue & " " & supplierIdRaw & " " & costCodeNormalized
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(costCodeNormalized)
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
'  RowIsSkippable -- matches transform_spec skip_rule.drop_when 7 conditions:
'    1. Net amount blank/zero AND (Gross-Tax) blank/zero
'    2. InvoiceNumber blank
'    3. SupplierID blank OR no "-" character (can't split)
'    4. Cost Code blank
'    5. InvoiceDate blank
'    6. TaxCode blank
'    7. ItemCode = literal "ItemCode" (header bleed guard for files that have
'       accidental title rows above true header row detected by schema stage)
'  --------------------------------------------------------------
Private Function RowIsSkippable( ByVal tableData As Variant, ByVal rowIndex As Long, ByVal supplierIdColumn As Long, ByVal invoiceNumberColumn As Long, ByVal costCodeColumn As Long, ByVal invoiceDateColumn As Long, ByVal taxCodeColumn As Long, ByVal descriptionColumn As Long,  _
    ByVal itemCodeColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, ByVal grossColumn As Long ) As Boolean
    Dim netAmt As Double
    Dim sidRaw As String
    netAmt = ResolveEffectiveNet(tableData(rowIndex, netColumn), tableData(rowIndex, grossColumn), tableData(rowIndex, taxColumn))

    sidRaw = Trim$(CStr(tableData(rowIndex, supplierIdColumn)))

    RowIsSkippable = (netAmt = 0#) Or (LenB(Trim$(CStr(tableData(rowIndex, invoiceNumberColumn)))) = 0) Or (LenB(sidRaw) = 0 Or InStr(sidRaw, "-") = 0) Or (LenB(Trim$(CStr(tableData(rowIndex, costCodeColumn)))) = 0) Or (LenB(Trim$(CStr(tableData(rowIndex,  _
        invoiceDateColumn)))) = 0) Or (LenB(Trim$(CStr(tableData(rowIndex, taxCodeColumn)))) = 0) Or (StrComp(Trim$(CStr(tableData(rowIndex, itemCodeColumn))), "ItemCode", vbTextCompare) = 0)
End Function

' Keyed-Collection primitives (CollectionKeyExists / *SetItemLong / *GetItemLong /
' *SetItemDouble / *GetItemDouble) are shared in modHelpers -- no per-module copy (2026-09-14).


