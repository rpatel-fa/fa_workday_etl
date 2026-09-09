Attribute VB_Name = "modApTaxiSupplierInvoice"
Option Explicit


' Source of truth: config/transform_specs/ap_taxi.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: Upload Format tab, header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).
'  Historical hard-coded constants retained for documentation audit only; NOT used at runtime:
'  Private Const SOURCE_WORKSHEET_NAME As String = "Upload Format"
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
'  These values come from ap_taxi.yaml transform_specs per_row_tax sets (P20/P0) and MUST
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
Private mSupplierLookup As Collection
Private mSpendCategoryLookup As Collection
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

' ====================================================================
'  Entry point (called by the Runner via reflection on selected transform key).
' ====================================================================
Public Sub RunApTaxiSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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

    '  -- Upload Format column indices (A..P = 16 cols) --
    Dim templateColumn As Long                ' A  Template (FBP/Proactis tag, unused for EIB)
    Dim supplierIdColumn As Long              ' B  SupplierID = "{Company}-{CreditorID}"
    Dim invoiceNumberColumn As Long           ' C  InvoiceNumber (per-row DIRECT, multiple per file)
    Dim invoiceDateColumn As Long             ' D  InvoiceDate -- per-row via NormalizeSourceDate
    Dim receivedDateColumn As Long            ' E  ReceivedDate -- per-row via NormalizeSourceDate
    Dim currencyColumn As Long                ' F  Currency -- per-row direct (fallback GBP)
    Dim trayColumn As Long                    ' G  Tray (FBP internal, unused)
    Dim createdByColumn As Long               ' H  CreatedBy (FBP internal, unused)
    Dim poNumberColumn As Long                ' I  PONumber -- optional (all blank in FIRST CLASS)
    Dim itemCodeColumn As Long                ' J  ItemCode -- header bleed guard only
    Dim costCodeColumn As Long                ' K  "Account" primary / "New Cost Code" / "Final Cost Code"
    Dim netColumn As Long                     ' L  Net
    Dim taxColumn As Long                     ' M  Tax
    Dim grossColumn As Long                   ' N  Gross
    Dim descriptionColumn As Long             ' O  "Description/BOOKED BY" primary / "Description" fallback
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

    '  16 Upload Format headers A..P. Aliases are handled at FindColumnIndexInRow calls
    '  below (col K tries "Account" then "New Cost Code" then "Final Cost Code"; col O
    '  tries "Description/BOOKED BY" then "Description"). Listed below are for schema
    '  validation catch-most -- alias fallbacks pick up the remaining variants.
    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApTaxiSupplierInvoice", 16)

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
        Err.Raise vbObjectError + 1910, "modApTaxiSupplierInvoice.RunApTaxiSupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    '  Resolve column indices -- with alias fallbacks for the 2 non-uniform columns.
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

    '  Col K: "Account" primary (FIRST CLASS). Fallback to stationery-style names
    '  if future taxi suppliers reuse the Proactis Upload convention headers.
    costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Account")
    If costCodeColumn = 0 Then costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "New Cost Code")
    If costCodeColumn = 0 Then costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Final Cost Code")
    '  Col L (Net) FIRST CLASS uses "Net" (no alias needed -- matches). Keep for parity
    '  with stationery just in case a future E-Courier-like taxi variant has "Net Amount".
    netColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net")
    If netColumn = 0 Then netColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net Amount")
    '  Col O: "Description/BOOKED BY" primary (FIRST CLASS). Fallback to the simpler
    '  stationery-style "Description" if future taxi suppliers simplify the header name.
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description/BOOKED BY")
    If descriptionColumn = 0 Then descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description")

    taxColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax")
    grossColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Gross")
    taxCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "TaxCode")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    '  CHANGE 2026-08-25 (FIRST CLASS Jul 26 Workday duplicate-ref rejection):
    '  Phase 1 group-map PREPASS over valid non-skipped rows. Produces 7 pure-VBA
    '  Collections (no ActiveX, no Scripting.Dictionary -- cross-OS safe via the same
    '  Collection helper pattern used for CFA Rule 8 per-group counter in
    '  modApCfaClaimsSupplierInvoice.bas):
    '    (A) groupSpreadsheetKeyId   -> key = invoiceNumberValue (col C), item = CLng B grouping key (1..M)
    '    (B) groupNextLineOrder      -> key = invoiceNumberValue, item = CLng next Line Order (1..N per group)
    '    (C) groupNetSum             -> key = invoiceNumberValue, item = CDbl Round(SUM Net,2)
    '    (D) groupCompanyCaptured    -> key = invoiceNumberValue, item = CStr Rule 1 Company for guard
    '    (E) groupSupplierCaptured   -> key = invoiceNumberValue, item = CStr Rule 2 Supplier for guard
    '    (F) groupCurrencyCaptured   -> key = invoiceNumberValue, item = CStr Rule 12 Currency for guard
    '    (G) groupInvDateCaptured    -> key = invoiceNumberValue, item = CStr Rule 4 InvoiceDate for guard
    '  First non-skipped appearance of a group populates all capturers AND assigns a new B ID 1..M;
    '  subsequent rows of same group accumulate Net SUM and assert header consistency Rule 18
    '  (4 header fields: Company/Supplier/Currency/InvoiceDate  --  ReceivedDate excluded deliberately
    '  because B+EM grouping is authoritative for Workday multi-line header collapse; a single
    '  Workday invoice can span multiple receipt months when split across cost codes at line level).
    '  If any of the 4 strict fields diverges within a group we raise a build ERROR before any
    '  EIB rows are written (corrupt source data, can't emit a valid multi-line Workday submission).
    Dim nextSpreadsheetId As Long
    nextSpreadsheetId = 0
    Dim groupSpreadsheetKeyId As New Collection   ' key=InvoiceNumberValue, item=CLng B id
    Dim groupNextLineOrder As New Collection      ' key=InvoiceNumberValue, item=CLng next line 1..N
    Dim groupNetSum As New Collection             ' key=InvoiceNumberValue, item=CDbl Round(sum,2)
    Dim groupCompanyCaptured As New Collection    ' key=InvoiceNumberValue, item=CStr company
    Dim groupSupplierCaptured As New Collection   ' key=InvoiceNumberValue, item=CStr supplier
    Dim groupCurrencyCaptured As New Collection    ' key=InvoiceNumberValue, item=CStr currency
    Dim groupInvDateCaptured As New Collection     ' key=InvoiceNumberValue, item=CStr invdate
    Dim prepassMsg As String
    Dim prepassPhase As Long
    prepassPhase = 0
    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, Array(supplierIdColumn, invoiceNumberColumn, costCodeColumn, descriptionColumn), Array(netColumn, grossColumn), Array(descriptionColumn, supplierIdColumn, invoiceNumberColumn)) Then
            ' trailer/empty -- skip in prepass too
        ElseIf RowIsSkippable(sourceTable, rowIndex, supplierIdColumn, invoiceNumberColumn, costCodeColumn, invoiceDateColumn, taxCodeColumn, descriptionColumn, itemCodeColumn, netColumn, taxColumn, grossColumn) Then
            ' missing required fields -- will be rejected in write loop; not part of any group
        Else
            Dim sRaw As String
            Dim hPos As Long
            Dim ccPerRow As String
            Dim crPerRow As String
            Dim rCompanyId As String
            Dim rSupplierId As String
            Dim nAmt As Double
            Dim gAmt As Double
            Dim tAmt As Double
            Dim iNum As String
            Dim iDt As String
            Dim rDt As String
            Dim cVal As String
            Dim tcStr As String
            sRaw = Trim$(CStr(sourceTable(rowIndex, supplierIdColumn)))
            hPos = InStr(sRaw, "-")
            If hPos = 0 Then GoTo NextPrepassRow
            ccPerRow = Left$(sRaw, hPos - 1)
            crPerRow = Mid$(sRaw, hPos + 1)
            rCompanyId = ResolveMappedCompany(runnerWorkbook, ccPerRow, OUTPUT_COMPANY_DEFAULT_FALLBACK)
            rSupplierId = ResolveSupplierWithCompanyPrecedence(mSupplierLookup, crPerRow, ccPerRow)
            If LenB(rSupplierId) = 0 Then GoTo NextPrepassRow
            Dim ccRaw As String
            ccRaw = Trim$(CStr(sourceTable(rowIndex, costCodeColumn)))
            If LenB(ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, ccRaw)) = 0 Then GoTo NextPrepassRow
            If LenB(ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, ccRaw)) = 0 Then GoTo NextPrepassRow
            tcStr = UCase$(Trim$(CStr(sourceTable(rowIndex, taxCodeColumn))))
            If tcStr <> "P20" And tcStr <> "P0" And tcStr <> "PZ" Then GoTo NextPrepassRow
            nAmt = ResolveAmount(sourceTable(rowIndex, netColumn))
            gAmt = ResolveAmount(sourceTable(rowIndex, grossColumn))
            tAmt = ResolveAmount(sourceTable(rowIndex, taxColumn))
            If nAmt = 0# Then nAmt = gAmt - tAmt
            If nAmt < 0# Then nAmt = -nAmt
            iNum = Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
            iDt = NormalizeSourceDate(sourceTable(rowIndex, invoiceDateColumn))
            rDt = NormalizeSourceDate(sourceTable(rowIndex, receivedDateColumn))
            If LenB(rDt) = 0 Then rDt = iDt
            cVal = Trim$(CStr(sourceTable(rowIndex, currencyColumn)))
            If LenB(cVal) = 0 Then cVal = OUTPUT_DEFAULT_CURRENCY
            If CollectionKeyExists(groupSpreadsheetKeyId, iNum) = False Then
                nextSpreadsheetId = nextSpreadsheetId + 1
                CollectionSetItemLong groupSpreadsheetKeyId, iNum, nextSpreadsheetId
                CollectionSetItemLong groupNextLineOrder, iNum, 1
                CollectionSetItemDouble groupNetSum, iNum, Round(nAmt, 2)
                groupCompanyCaptured.Add rCompanyId, iNum
                groupSupplierCaptured.Add rSupplierId, iNum
                groupCurrencyCaptured.Add cVal, iNum
                groupInvDateCaptured.Add iDt, iNum
            Else
                '  accumulate Net, then Rule 18 header consistency guard (4 strict fields)
                CollectionSetItemDouble groupNetSum, iNum, Round(CDbl(CollectionGetItemDouble(groupNetSum, iNum)) + Round(nAmt, 2), 2)
                prepassMsg = vbNullString
                If StrComp(CStr(groupCompanyCaptured(iNum)), rCompanyId, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "Company(" & CStr(groupCompanyCaptured(iNum)) & " vs " & rCompanyId & "); "
                If StrComp(CStr(groupSupplierCaptured(iNum)), rSupplierId, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "Supplier(" & CStr(groupSupplierCaptured(iNum)) & " vs " & rSupplierId & "); "
                If StrComp(CStr(groupCurrencyCaptured(iNum)), cVal, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "Currency(" & CStr(groupCurrencyCaptured(iNum)) & " vs " & cVal & "); "
                If StrComp(CStr(groupInvDateCaptured(iNum)), iDt, vbBinaryCompare) <> 0 Then prepassMsg = prepassMsg & "InvoiceDate(" & CStr(groupInvDateCaptured(iNum)) & " vs " & iDt & "); "
                If LenB(prepassMsg) > 0 Then
                    Err.Raise vbObjectError + 1911, "modApTaxiSupplierInvoice.RunApTaxiSupplierInvoice Phase1GroupPrepass", _
                        "Rule 18 Header Consistency Guard (4-field) failed. Same InvoiceNumber '" & iNum & "' source rows have diverged header fields. " & prepassMsg & "ReceivedDate intentionally excluded (legit per-group variance OK). Fix source data FIRST before re-running the Runner."
                End If
            End If
        End If
NextPrepassRow:
    Next rowIndex

    '  Structural note vs stationery: ONE TAXI FILE = MULTIPLE INVOICES
    '  (FIRST CLASS exemplar: 21 distinct Invoice groups across 32 rows,
    '   1-5 rows per invoice group; 8 invoice groups are multi-line >1 rows.
    '   Group map assigned nextSpreadsheetId 1..21 in prepass Phase1.)
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
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Missing required Upload Format field (InvoiceNumber blank / " & "SupplierID blank/no-hyphen / Account Cost Code blank / " & "InvoiceDate blank / TaxCode blank / ItemCode literal header bleed / " &  _
                "Net blank/zero AND Gross-Tax blank/zero).", Trim$(CStr(sourceTable(rowIndex, supplierIdColumn))), Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
        Else
            '  -- Split SupplierID on first hyphen ? company + creditor PER ROW. --
            Dim sidRaw As String
            Dim hyphenPos As Long
            sidRaw = Trim$(CStr(sourceTable(rowIndex, supplierIdColumn)))
            hyphenPos = InStr(sidRaw, "-")
            companyCodePerRow = Left$(sidRaw, hyphenPos - 1)
            creditorIdPerRow = Mid$(sidRaw, hyphenPos + 1)

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
                Dim costCodeRaw As String
                costCodeRaw = Trim$(CStr(sourceTable(rowIndex, costCodeColumn)))

                costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( costCenterLookup, costCodeRaw)
                If LenB(costCenterId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedCostCenterRows = unmappedCostCenterRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Cost centre unresolved: Account Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", " & "Core vs Prism = Core, lowest hierarchy match)", costCodeRaw, Trim$(CStr(sourceTable(rowIndex,  _
                        descriptionColumn)))
                Else
                    spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( mSpendCategoryLookup, costCodeRaw)
                    If LenB(spendCategoryId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Spend Category unresolved: Account Cost Code did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", " & "Core vs Prism = Core, lowest hierarchy match)", costCodeRaw,  _
                            Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                    Else
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
                                grossColumn, currencyColumn, poNumberColumn, supplierIdColumn, costCenterId, spendCategoryId, optionalWorktagBlankRows, groupSpreadsheetKeyId, groupNextLineOrder, groupNetSum) Then
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
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Upload Format required field (Net/Gross " & "blank/zero / InvoiceNumber blank / SupplierID blank/no-hyphen / " & "Account Cost Code blank / InvoiceDate blank / TaxCode blank / " &  _
            "ItemCode literal 'ItemCode' header-bleed)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) & " source row(s) whose CreditorID did not resolve to a Workday Supplier ID" & " via the Supplier Spreadsheet (2-part company precedence)."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) & " source row(s) whose Account Cost Code did not resolve to a Workday cost centre" & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) & " source row(s) whose Account Cost Code did not resolve to a Workday spend category" & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unknownTaxCodeRows > 0 Then
        issueParts.Add "Skipped " & CStr(unknownTaxCodeRows) & " source row(s) with unknown TaxCode (only P20 / P0 / PZ accepted)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) & " source row(s) had 1 or more blank optional worktags" & " (Event / Camp / FA Learning / Commercial Partner / Grant / Program / Project)" & " -- Account Cost Code did not fully map to all 7 worktag hierarchies."
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
    Err.Raise failureNumber, "modApTaxiSupplierInvoice.RunApTaxiSupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row.
'  INPUTS:
'     resolvedCompanyId, supplierId -- per-row SupplierID split outcomes.
'     taxApply/taxCod/taxRecov -- per-row TaxCode switch results.
'     All other Long params are 1-based column indices into sourceTable.
'  OUTPUT COLUMN MAPPING (matches mapping: in transform_specs/ap_taxi.yaml -- Submit_Supplier_Invoice_v46.2):
'     B  seq            (sequenceNumber)
'     I  submit         (Y)
'     K  InvoiceNumber  (Rule 3 -- Upload Format col C DIRECT)
'     R  Company        (Rule 1 -- per-row SupplierID split LEFT segment,
'                        mapping sheet entity=company, fallback sentinel defaultCompanyId)
'     T  Currency       (Rule 12 -- col F DIRECT; fallback = GBP if blank)
'     U  Supplier       (Rule 2 -- per-row SupplierID split RIGHT segment via
'                        2-part ResolveSupplierWithCompanyPrecedence)
'     AF InvoiceDate    (Rule 4 -- col D per-row via NormalizeSourceDate, coerced @ format)
'     AG ReceiptDate    (Rule 5 -- Invoice Received Date, col E per-row via NormalizeSourceDate;
'                        fallback = InvoiceDate if blank)
'     BG SuppliersInvoiceNumber (Rule 6 -- same value as K InvoiceNumber)
'     BH PONumber       (Rule 13 -- External PO Number, col I DIRECT if non-blank else leave empty)
'     EK line number 1  (line Row ID)
'     EP ItemDescription (Rule 8 -- UCASE(TRIM(col O Description/BOOKED BY)))
'     EW Spend Category  (Rule 7 Seg2 -- CostCode via mapping MASTER, gated upstream)
'     FE Applicability / FF Code / FH date-type / FK Recoverability / FL CALC_TAX_DUE
'     GE Unit Cost / GF Extended Amount (Rule 10 Net for BOTH  --  GE and GF MUST be identical; CALC_TAX_DUE lets Workday re-derive VAT)
'     GQ Memo (Rule 9 -- InvoiceNumber + ' ' + SupplierID + ' ' + CostCodeNormalized)
'     GR Reporting Type  (second segment first char via ResolveReportingTypeFromFinalCode)
'     GT Cost Centre     (CostCode via mapping MASTER, gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (Event Camp FALearning CommercialPartner Grant Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRow( ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal resolvedCompanyId As String, ByVal supplierId As String, ByVal taxApply As String, ByVal taxCod As String,  _
    ByVal taxRecov As String, ByVal invoiceNumberColumn As Long, ByVal invoiceDateColumn As Long, ByVal receivedDateColumn As Long, ByVal descriptionColumn As Long, ByVal costCodeColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, ByVal grossColumn As Long,  _
    ByVal currencyColumn As Long, ByVal poNumberColumn As Long, ByVal supplierIdColumn As Long, ByVal costCenterId As String, ByVal spendCategoryId As String, ByRef optionalWorktagBlankRows As Long, _
    ByRef groupSpreadsheetKeyId As Collection, ByRef groupNextLineOrder As Collection, ByRef groupNetSum As Collection ) As Boolean
    Dim netAmt As Double
    Dim grossAmt As Double
    Dim taxAmt As Double
    Dim costCodeRaw As String
    Dim costCodeNormalized As String
    Dim invoiceNumberValue As String
    Dim descriptionValue As String
    Dim invoiceDateValue As String
    Dim receiptDateValue As String
    Dim currencyValue As String
    Dim poValue As String
    Dim supplierIdRaw As String
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
    Dim anyWorktagBlank As Boolean
    Dim bGroupKey As Long
    Dim lineOrderNext As Long
    Dim netSumGroup As Double

    '  INPUTS (source_table taxi "Upload Format" tab A..P):
    '    A=Template  B=SupplierID (split LEFT=Company, RIGHT=Creditor)  C=InvoiceNumber
    '    D=InvoiceDate  E=ReceivedDate  F=Currency  G=Tray  H=CreatedBy  I=PONumber
    '    J=ItemCode  K=Account (Cost Code)  L=Net  M=Tax  N=Gross  O=Description/BOOKED BY
    '    P=TaxCode
    '  OUTPUT COLUMN MAPPING (Supplier Invoice EIB v46.2 template, header row 5 labels):
    '    Rule 15 CHANGED 2026-08-25 after Finance duplicate-ref rejection: B Spreadsheet Key*
    '         = SHARED GROUP KEY across all rows of same InvoiceNumber C (1 integer per distinct C value,
    '           not per-row). Read from groupSpreadsheetKeyId collection assigned in Phase 1 prepass.
    '         All rows of one invoice share the same B value -> Workday reads one header, N lines.
    '    I=Add Only (OUTPUT_SUBMIT = Y) | Rule 3 K Invoice Number = source C DIRECT |
    '    Rule 1 R Company = split LEFT segment resolved | Rule 12 T Currency = F DIRECT |
    '    Rule 2 U Supplier = split RIGHT segment 2-part precedence |
    '    Rule 4 AF Invoice Date = D coerced YYYY-MM-DD | Rule 5 AG Receipt Date = E fallback D |
    '    Rule 17 AS Control Amount Total  --  LEAVE BLANK (2026-09-09: AP confirmed same Stationery fix applies to all AP transforms; no value here).
    '    Rule 6 BG Suppliers Invoice Number = 1:1 COPY of K (same invoiceNumberValue local,
    '         mirror_invoice_number_value_from_target_column_k spec function; cross_column_guarantee
    '         bitwise identical per row. Previously legacy AU col, now BG per v46.2 header).
    '    Rule 13 BH PO Number = I DIRECT if non-blank |
    '    NEW EL col 142 Supplier Invoice Line ID = Leave EMPTY (Workday auto-allocs from EM Line Order) |
    '    NEW Rule 16 EM col 143 Line Order = per-group 1..N reset counter (groupNextLineOrder Phase 1 map,
    '         increment AFTER write so Phase 2 writes sequentially). Mirrors AR multiline DV Line Order.
    '    EK Row ID = 1 (HOUSE DEFAULT invoice_line_row_id_constant) |
    '    Rule 8 EP Description = UCase(col O description) | Rule 7 line worktags EW..HH = resolved Cost Code
    '         segments via Mapping MASTER Hierarchy lowest specificity match (9 pack).
    '    Rule 11 tax triple FE Applicability / FF Code / FH INVOICE/ADJUSTMENT_DATE_TYPE /
    '         FK Recoverability / FL CALC_TAX_DUE (workday re-derives VAT from the triple).
    '    Rule 10 GE Unit Cost = col L Net (fallback Gross-Tax). GF Extended Amount = SAME as GE (net only; AP confirmed both columns MUST match. CALC_TAX_DUE re-derives VAT from FE/FF/FK triple).
    '    Rule 9 GQ Memo = InvoiceNumber + ' ' + SupplierID Raw + ' ' + CostCodeNormalized (traceability).

    invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, descriptionColumn))))

    '  Rule 10: Net primary (col L). If blank/zero ? fallback Gross - Tax.
    netAmt = ResolveAmount(sourceTable(sourceRow, netColumn))
    grossAmt = ResolveAmount(sourceTable(sourceRow, grossColumn))
    taxAmt = ResolveAmount(sourceTable(sourceRow, taxColumn))
    If netAmt = 0# Then netAmt = grossAmt - taxAmt
    If netAmt < 0# Then netAmt = -netAmt

    '  Rule 7: cost code separator normalization.
    costCodeRaw = Trim$(CStr(sourceTable(sourceRow, costCodeColumn)))
    costCodeNormalized = NormalizeCostCodeSeparators(costCodeRaw)

    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    receiptDateValue = NormalizeSourceDate(sourceTable(sourceRow, receivedDateColumn))
    '  Rule 5: ReceiptDate (AG)  --  CHANGED 2026-09-09. AP confirmed (Will's Aug Ambassador Taxi EIB):
    '  AG MUST BE BLANK (no read from ReceivedDate, no InvoiceDate fallback).
    receiptDateValue = vbNullString

    currencyValue = Trim$(CStr(sourceTable(sourceRow, currencyColumn)))
    If LenB(currencyValue) = 0 Then currencyValue = OUTPUT_DEFAULT_CURRENCY

    poValue = Trim$(CStr(sourceTable(sourceRow, poNumberColumn)))
    supplierIdRaw = Trim$(CStr(sourceTable(sourceRow, supplierIdColumn)))

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

    '  Rule 15, 16, 17: look up the group map (built in Phase 1 prepass).
    '  Missing key at this point is INTERNAL DEFECT (Phase1/Phase2 skip_rule parity).
    bGroupKey = CollectionGetItemLong(groupSpreadsheetKeyId, invoiceNumberValue)
    lineOrderNext = CollectionGetItemLong(groupNextLineOrder, invoiceNumberValue)
    netSumGroup = CollectionGetItemDouble(groupNetSum, invoiceNumberValue)

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = bGroupKey                   ' Rule 15 (CHANGED 2026-08-25: per-group SHARED, no longer per-row sequential)
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT                ' HOUSE DEFAULT Add Only = Y
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue           ' Rule 3 (source C DIRECT -- remains SAME as BG, both = raw passthrough; FIRST CLASS taxi already unique)
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = resolvedCompanyId            ' Rule 1
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = currencyValue                ' Rule 12
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId                   ' Rule 2
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue            ' Rule 4
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = Empty                      ' Rule 5 (CHANGED 2026-09-09: BLANK per Will's Aug EIB -- not receiptDateValue)
    targetSheet.Cells(targetRow, ColumnIndex("AS")).Value = Empty                      ' Rule 17 (Control Amount Total -- LEAVE BLANK; AP confirmed 2026-09-09 same stationery fix applies)
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = invoiceNumberValue          ' Rule 6 (mirror Rule 3 K = same string instance; cross_column_guarantee byte-identical)
    If LenB(poValue) > 0 Then targetSheet.Cells(targetRow, ColumnIndex("BH")).Value = poValue   ' Rule 13
    ' NEW: Supplier Invoice Line ID stub col 142. Intentionally BLANK per spec.
    ' Leave EMPTY -- Workday auto-allocates internal line IDs from Line Order col EM.
    ' Explicit write keeps bidirectional drift guard HYGIENE OK (no spec/code drift).
    targetSheet.Cells(targetRow, ColumnIndex("EL")).Value = Empty
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderNext               ' Rule 16 (NEW 2026-08-25: Line Order per group 1..N reset -- CORRECT for mixed single/multi-line batches (Sept). B+EM = authoritative line position; single-line groups trivially produce 1.)
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderNext               ' Rule 19 (EXACT MIRROR of EM col 143 -- per Sept mixed single/multi-line grouping requirement. Written from same lineOrderNext local to guarantee EK == EM per-row byte identity.)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue            ' Rule 8
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = taxApply
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = taxCod
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = taxRecov
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = netAmt                      ' Rule 10 (Unit Cost = net)
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = netAmt                      ' Rule 10 (Extended Amount = net SAME as GE; AP confirmed 2026-09-09 both columns MUST match  --  no VAT pre-add, CALC_TAX_DUE re-derives)
    '  Rule 9: Memo = InvoiceNumber + ' ' + SupplierID + ' ' + CostCodeNormalized
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = invoiceNumberValue & " " & supplierIdRaw & " " & costCodeNormalized
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(costCodeNormalized)
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = costCenterId
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = eventId
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = campId
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = faLearningId
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = commercialPartnerId
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = grantId
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = programId
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = projectId

    '  Phase 2 line-order advance: increment counter for THIS group AFTER successful write.
    '  (Separate prepass initialised next = 1; so write 1 then increment to 2, etc.)
    CollectionSetItemLong groupNextLineOrder, invoiceNumberValue, lineOrderNext + 1

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
'    7. ItemCode = literal "ItemCode" (header bleed guard)
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

' -----------------------------------------------------------------
'  CollectionKeyExists -- Mac-safe replacement for Scripting.Dictionary.Exists().
'  VBA Collection has no native .Exists method; attempting to retrieve an item by a
'  missing key raises error 5 "Invalid procedure call or argument".  Swallow ONLY
'  that specific error so unrelated runtime errors still surface.
'  This helper (plus the CollectionSetItem / CollectionGetItem pairs below) is used
'  INSTEAD OF Scripting.Dictionary because Mac Excel does NOT ship the Scripting
'  Runtime ActiveX library and throws run-time 429 "ActiveX component can't create
'  object" on CreateObject("Scripting.Dictionary").
' -----------------------------------------------------------------
Private Function CollectionKeyExists( _
    ByRef coll As Collection, _
    ByVal key As String _
) As Boolean
    On Error Resume Next
    Dim dummy As Variant
    dummy = coll(key)
    If Err.Number = 0 Then
        CollectionKeyExists = True
    ElseIf Err.Number = 5 Then
        CollectionKeyExists = False
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    On Error GoTo 0
End Function

' -----------------------------------------------------------------
'  Replace-or-insert a keyed Long item into a VBA Collection.
'  VBA Collections can't do in-place item replacement by key; to update we must
'  first Remove() the existing key (if present), then Add() the new value.
'  Value is stringified as CStr(value) Long on store; caller CLng()s on get.
' -----------------------------------------------------------------
Private Sub CollectionSetItemLong( _
    ByRef coll As Collection, _
    ByVal key As String, _
    ByVal value As Long _
)
    If CollectionKeyExists(coll, key) Then
        coll.Remove key
    End If
    coll.Add CStr(value), key
End Sub

' -----------------------------------------------------------------
'  Retrieve a keyed Long item from a VBA Collection.  Caller MUST have
'  verified existence with CollectionKeyExists() first (missing key raises Err 5).
' -----------------------------------------------------------------
Private Function CollectionGetItemLong( _
    ByRef coll As Collection, _
    ByVal key As String _
) As Long
    CollectionGetItemLong = CLng(coll(key))
End Function

' -----------------------------------------------------------------
'  Double variant of CollectionSetItemLong (for sterling net-sum accumulation).
'  Stores CStr(Round(value,2)) so values like 8018.05 survive the Variant->String
'  round-trip with the same 2dp precision used at write time.
' -----------------------------------------------------------------
Private Sub CollectionSetItemDouble( _
    ByRef coll As Collection, _
    ByVal key As String, _
    ByVal value As Double _
)
    If CollectionKeyExists(coll, key) Then
        coll.Remove key
    End If
    coll.Add CStr(Round(value, 2)), key
End Sub

' -----------------------------------------------------------------
'  Retrieve a keyed Double item from a VBA Collection (2dp rounded on store).
' -----------------------------------------------------------------
Private Function CollectionGetItemDouble( _
    ByRef coll As Collection, _
    ByVal key As String _
) As Double
    CollectionGetItemDouble = CDbl(coll(key))
End Function

