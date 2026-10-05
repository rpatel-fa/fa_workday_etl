Attribute VB_Name = "modApTravelPlacesSupInv"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_travel_places.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: Travel Places upload vMR tab, header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution / Header Row / Data Start Row).

'  -- NOTE: no fixed SOURCE_SUPPLIER_REFERENCE_ID; supplier is per-row DIRECT from col C "Workday Supplier ID". --
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
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"   ' GE/GF display 2dp (stored value not rounded)

'  --- Profile-local Tax Triple overrides (hybrid scope; shadows modFinanceCommon.bas shared defaults) ---
'  These values come from ap_travel_places.yaml transform_specs per_row_tax sets (P20/P0) and MUST
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
'  CHANGED 2026-10-05 (ap_travel_places ONLY): PZ/P0 recoverability = Fully_Recoverable (was
'  Nonrecoverable_Allocate). Applicability/Code unchanged. Taxi/Stationery keep Nonrecoverable_Allocate.
Private Const TAX_RECOVERABILITY_ZERO As String = "Fully_Recoverable"

'  9 worktag lookups: 2 mandatory (Cost Center, Spend Category) + 7 optional.
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
Public Sub RunApTravelPlacesSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim zeroValueRowsSkipped As Long
    Dim unmappedCostCenterRows As Long
    Dim unmappedSpendCategoryRows As Long
    Dim unresolvedSupplierRows As Long
    Dim unknownTaxCodeRows As Long
    Dim optionalWorktagBlankRows As Long
    Dim rowIndex As Long

    '  -- Travel Places upload column indices (A..R = 18 cols) --
    Dim templateColumn As Long                ' A  Template (GP tag "DINV", unused for EIB)
    Dim supplierIdColumn As Long              ' B  SupplierID = "{Company}-{CreditorID}" (Company split only)
    Dim workdaySupplierIdColumn As Long       ' C  "Workday Supplier ID" -- DIRECT to EIB U (no lookup)
    Dim invoiceNumberColumn As Long           ' D  InvoiceNumber (per-row DIRECT -> EIB K + BG)
    Dim invoiceDateColumn As Long             ' E  "SL Date" -- per-row via NormalizeSourceDate
    Dim receivedDateColumn As Long            ' F  "Date sent" -- per-row via NormalizeSourceDate
    Dim currencyColumn As Long                ' G  Currency -- per-row direct (fallback GBP)
    Dim trayColumn As Long                    ' H  Tray (agency internal, unused)
    Dim createdByColumn As Long               ' I  Createdby (agency internal, unused)
    Dim costCodeColumn As Long                ' R  String Reformatted -- THE accounting/coding string (worktag key)
    Dim itemCodeColumn As Long                ' K  Itemcode -- header bleed guard only
    Dim netColumn As Long                     ' M  Net
    Dim taxColumn As Long                     ' N  Tax
    Dim grossColumn As Long                   ' O  Gross
    Dim descriptionColumn As Long             ' P  Description
    Dim taxCodeColumn As Long                 ' Q  Taxcode -- per-row switch (PZ/P0/P20)

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

    '  No Supplier Spreadsheet lookup: Supplier (EIB U) is read DIRECT from col C "Workday Supplier ID".

    '  Travel Places headers used for header-row detection. col L "Account" is present
    '  in the header row but BLANK on every data row (coding lives in col R String Reformatted), so it
    '  is intentionally omitted here -- the remaining 17 uniquely identify the header row.
    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApTravelPlacesSupInv", 17)

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
        Err.Raise vbObjectError + 1910, "modApTravelPlacesSupInv.RunApTravelPlacesSupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    '  Resolve column indices. Delta 3: the worktag key column is String Reformatted (col R),
    '  NOT Account (col L) and NOT the legacy PONumber (col J). Delta 2: SL Date / Date sent / Description header names.
    templateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Template")
    supplierIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "SupplierID")
    workdaySupplierIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday Supplier ID")
    invoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "InvoiceNumber")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "SL Date")
    receivedDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Date sent")
    currencyColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Currency")
    trayColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tray")
    createdByColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Createdby")
    itemCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Itemcode")

    '  Delta 3: coding string is col R "String Reformatted" (finance-cleaned; supersedes col J
    '  "PONumber" which has formulas + hardcoded fixes). any columns after R are the finance user's own
    '  derived/test worktag columns and are intentionally NOT read; we re-derive via Mapping MASTER.
    costCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "String Reformatted")

    netColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net")
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description")
    taxColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax")
    grossColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Gross")
    taxCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Taxcode")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

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
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Missing required Upload field (InvoiceNumber blank / " & "SupplierID blank/no-hyphen / col R coding blank / " & "SL Date blank / Taxcode blank / Itemcode literal header bleed).",  _
                Trim$(CStr(sourceTable(rowIndex, supplierIdColumn))), Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
        ElseIf ResolveEffectiveNet(sourceTable(rowIndex, netColumn), sourceTable(rowIndex, grossColumn), sourceTable(rowIndex, taxColumn)) = 0# Then
            '  Zero-value line: Net = 0 AND (Gross - Tax) = 0 -> benign empty line, nothing to post.
            '  SILENT skip: NOT written and NOT added to the reject file (config:
            '  zero_value_line_handling: skip_silently). Counted only for the informational summary
            '  so the run stays auditable. It is deliberately NOT added to rowsSkipped (= the run
            '  log's "rejected" count). A NEGATIVE Net (credit note) is non-zero and still posts.
            zeroValueRowsSkipped = zeroValueRowsSkipped + 1
        Else
            '  -- Split SupplierID on first hyphen -> company + creditor PER ROW. --
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

            '  Supplier (EIB U) = col C "Workday Supplier ID" DIRECT passthrough (no Supplier
            '  Spreadsheet lookup). Blank -> reject as unresolved supplier.
            supplierId = Trim$(CStr(sourceTable(rowIndex, workdaySupplierIdColumn)))
            If LenB(supplierId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unresolvedSupplierRows = unresolvedSupplierRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Supplier unresolved: col C ""Workday Supplier ID"" was BLANK" & " (direct passthrough to EIB U; no Supplier Spreadsheet lookup)." & " Company=" &  _
                    companyCodePerRow & ".", creditorIdPerRow, "SupplierID raw=" & sidRaw
            Else
                Dim costCodeRaw As String
                costCodeRaw = Trim$(CStr(sourceTable(rowIndex, costCodeColumn)))
                '  col R "String Reformatted" arrives hyphen- (509 rows), space- (6) or dot-delimited (2);
                '  delimiter normalisation is owned by the SHARED resolver
                '  (ResolveLowestHierarchyLookupAcrossFinalCodeSegments / FinalCodeSegment), so we pass
                '  the raw code and every delimiter form resolves. costCodeRaw stays raw for the reject sheet.
                costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( costCenterLookup, costCodeRaw)
                If LenB(costCenterId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedCostCenterRows = unmappedCostCenterRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Cost centre unresolved: col R coding did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", " & "Core vs Prism = Core, lowest hierarchy match)", costCodeRaw, Trim$(CStr(sourceTable(rowIndex,  _
                        descriptionColumn)))
                Else
                    spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( mSpendCategoryLookup, costCodeRaw)
                    If LenB(spendCategoryId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Spend Category unresolved: col R coding did not resolve in '" & mappingSheetName & "' (" & WBO_SPEND_CATEGORY & ", " & "Core vs Prism = Core, lowest hierarchy match)", costCodeRaw,  _
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
                            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Unknown Taxcode '" & tcRaw & "' -- expected P20/P0/PZ. " & "Per spec on_unknown_when_value = reject_and_report.", "Taxcode raw=" & tcRaw, Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                        Else
                            If WriteOutputRow( targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, resolvedCompanyId, supplierId, taxApply, taxCod, taxRecov, invoiceNumberColumn, invoiceDateColumn, receivedDateColumn, descriptionColumn, costCodeColumn, netColumn, taxColumn,  _
                                grossColumn, currencyColumn, supplierIdColumn, costCenterId, spendCategoryId, optionalWorktagBlankRows) Then
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
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Upload required field (InvoiceNumber blank / " & "SupplierID blank/no-hyphen / " & "col R coding blank / SL Date blank / Taxcode blank / " &  _
            "Itemcode literal 'Itemcode' header-bleed)."
    End If
    If zeroValueRowsSkipped > 0 Then
        issueParts.Add "Skipped " & CStr(zeroValueRowsSkipped) & " zero-value line(s) (Net = 0 and Gross - Tax = 0)" & " -- benign, nothing to post; excluded from the reject file by design."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) & " source row(s) with a blank col C ""Workday Supplier ID""" & " (direct passthrough to EIB U)."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) & " source row(s) whose col R coding did not resolve to a Workday cost centre" & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) & " source row(s) whose col R coding did not resolve to a Workday spend category" & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unknownTaxCodeRows > 0 Then
        issueParts.Add "Skipped " & CStr(unknownTaxCodeRows) & " source row(s) with unknown Taxcode (only P20 / P0 / PZ accepted)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) & " source row(s) had 1 or more blank optional worktags" & " (Event / Camp / FA Learning / Commercial Partner / Grant / Program / Project)" & " -- col R coding did not fully map to all 7 worktag hierarchies."
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
    Err.Raise failureNumber, "modApTravelPlacesSupInv.RunApTravelPlacesSupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row.
'  INPUTS:
'     resolvedCompanyId -- per-row SupplierID split outcome; supplierId -- col C "Workday Supplier ID" direct.
'     taxApply/taxCod/taxRecov -- per-row Taxcode switch results.
'     All other Long params are 1-based column indices into sourceTable.
'  OUTPUT COLUMN MAPPING (matches mapping: in transform_specs/ap_travel_places.yaml -- Submit_Supplier_Invoice_v46.2):
'     B  seq            (sequenceNumber)
'     I  submit         (Y)
'     K  InvoiceNumber  (Rule 3 -- col D DIRECT)
'     R  Company        (Rule 1 -- per-row SupplierID split LEFT segment,
'                        mapping sheet entity=company, fallback sentinel)
'     T  Currency       (Rule 12 -- col G DIRECT; fallback = GBP if blank)
'     U  Supplier       (Rule 2 -- col C "Workday Supplier ID" DIRECT passthrough; no lookup)
'     AF InvoiceDate    (Rule 4 -- col E "SL Date" per-row via NormalizeSourceDate)
'     AG ReceiptDate    (Rule 5 -- Invoice Received Date, col F "Date sent" per-row; fallback = InvoiceDate)
'     BG SuppliersInvoiceNumber (Rule 6 -- col D InvoiceNumber DIRECT; same value as K)
'     (External PO Number BH intentionally NOT written -- Rule 13; col R "String Reformatted" is the coding string)
'     EK line number 1  (line Row ID)
'     EP ItemDescription (Rule 8 -- UCASE(TRIM(col P Description)))
'     EW Spend Category  (Rule 7 -- col R "String Reformatted" coding via mapping MASTER, gated upstream)
'     FE Applicability / FF Code / FH date-type / FK Recoverability / FL CALC_TAX_DUE
'     GE Unit Cost / GF Extended Amount (Rule 10 Net for BOTH — GE and GF MUST be identical;
'                         CALC_TAX_DUE lets Workday re-derive VAT; NEGATIVE net kept as-is for credit-note rows;
'                         NumberFormat "0.00" = 2dp display only, stored value not rounded)
'     GQ Memo (Rule 9 -- InvoiceNumber + ' ' + CostCode raw; mirrors ap_referees_v2)
'     GR Reporting Type  (second segment first char via ResolveReportingTypeFromFinalCode)
'     GT Cost Centre     (col R "String Reformatted" coding via mapping MASTER, gated upstream)
'     GV GX GZ HB HD HF HH 7 optional worktags (Event Camp FALearning CommercialPartner Grant Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRow( ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal sequenceNumber As Long, ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal resolvedCompanyId As String, ByVal supplierId As String, ByVal taxApply As String, ByVal taxCod As String,  _
    ByVal taxRecov As String, ByVal invoiceNumberColumn As Long, ByVal invoiceDateColumn As Long, ByVal receivedDateColumn As Long, ByVal descriptionColumn As Long, ByVal costCodeColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, ByVal grossColumn As Long,  _
    ByVal currencyColumn As Long, ByVal supplierIdColumn As Long, ByVal costCenterId As String, ByVal spendCategoryId As String, ByRef optionalWorktagBlankRows As Long ) As Boolean
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
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
    Dim anyWorktagBlank As Boolean

    invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, descriptionColumn))))

    '  Rule 10: Net primary (col M). If blank/zero -> fallback Gross - Tax.
    '  NEGATIVE net is KEPT (credit-note rows post as negative lines -- no abs()).
    netAmt = ResolveAmount(sourceTable(sourceRow, netColumn))
    grossAmt = ResolveAmount(sourceTable(sourceRow, grossColumn))
    taxAmt = ResolveAmount(sourceTable(sourceRow, taxColumn))
    If netAmt = 0# Then netAmt = grossAmt - taxAmt

    '  Rule 7: cost code separator normalization (col R String Reformatted -- finance-cleaned,
    '  no "*XX" booker suffix; already dash-separated so normalization is a safe no-op).
    costCodeRaw = Trim$(CStr(sourceTable(sourceRow, costCodeColumn)))
    costCodeNormalized = NormalizeCostCodeSeparators(costCodeRaw)

    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    receiptDateValue = NormalizeSourceDate(sourceTable(sourceRow, receivedDateColumn))
    '  Rule 5: ReceiptDate (AG) — CHANGED 2026-09-09 per Will's hand-edited Aug Ambassador Taxi EIB
    '  (14/14 AG=None). AP "same for all of them" rule: LEAVE BLANK / do not populate.
    '  Do NOT fall back to InvoiceDate. receiptDateValue is intentionally never read; the
    '  WriteOutputRow call below writes Empty directly to ColumnIndex("AG").

    currencyValue = Trim$(CStr(sourceTable(sourceRow, currencyColumn)))
    If LenB(currencyValue) = 0 Then currencyValue = OUTPUT_DEFAULT_CURRENCY


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

    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceNumberValue, mGrpNextKey, lineOrderValue)
    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue       ' Rule 3
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = resolvedCompanyId        ' Rule 1
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = currencyValue            ' Rule 12
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId               ' Rule 2
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue         ' Rule 4
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = Empty                   ' Rule 5 (CHANGED 2026-09-09: Leave BLANK per Will's Aug EIB. AG=None all rows. Do NOT write receiptDateValue fallback)
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = invoiceNumberValue      ' Rule 6
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue    ' Row ID = 1..N within invoice group (CHANGED 2026-09-14: shared grouping; single-line rows get 1)
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue    ' Line Order = 1..N within invoice group (mirror of EK; ADDED 2026-09-14)
    '  Rule 16 EM (Line Order col 143) not explicitly written — default BLANK matches Will's Aug EIB (EM=None all rows, Workday derives from B Spreadsheet Key group) (CHANGED 2026-09-09 documented match)
    '  Rule 17 AS (Control Amount Total col 45) not explicitly written — default BLANK matches AP confirmed rule (no per-group SUM)
    '  Supplier Invoice Line ID stub EL not written — default BLANK (Workday auto-assigns)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue        ' Rule 8
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = taxApply
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = taxCod
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = taxRecov
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = netAmt                  ' Rule 10 (Unit Cost = net)
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = netAmt                  ' Rule 10 (Extended Amount = net SAME as GE; AP confirmed 2026-09-09 both columns MUST match — no VAT pre-add, CALC_TAX_DUE re-derives)
    '  Rule 9: Memo = InvoiceNumber + ' ' + CostCode (CHANGED 2026-10-01: mirrors ap_referees_v2 -- raw coding string, no SupplierID)
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = invoiceNumberValue & " " & costCodeRaw
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
'  RowIsSkippable -- matches transform_spec skip_rule.drop_when:
'    1. Net==0 AND (Gross-Tax)==0 (empty-line guard; NEGATIVE net survives)
'    2. InvoiceNumber blank
'    3. SupplierID blank OR no "-" character (can't split)
'    4. col R coding blank
'    5. SL Date blank
'    6. Taxcode blank
'    7. Itemcode = literal "Itemcode" (header bleed guard)
'  --------------------------------------------------------------
Private Function RowIsSkippable( ByVal tableData As Variant, ByVal rowIndex As Long, ByVal supplierIdColumn As Long, ByVal invoiceNumberColumn As Long, ByVal costCodeColumn As Long, ByVal invoiceDateColumn As Long, ByVal taxCodeColumn As Long, ByVal descriptionColumn As Long,  _
    ByVal itemCodeColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, ByVal grossColumn As Long ) As Boolean
    Dim sidRaw As String
    '  Zero-value (Net = 0 AND Gross - Tax = 0) is NOT a reject condition -- it is a benign silent
    '  skip handled in the main loop (config: zero_value_line_handling: skip_silently), so a 0/0/0
    '  line never becomes a false negative in the reject file. A NEGATIVE Net (credit note) posts.
    sidRaw = Trim$(CStr(tableData(rowIndex, supplierIdColumn)))

    RowIsSkippable = (LenB(Trim$(CStr(tableData(rowIndex, invoiceNumberColumn)))) = 0) Or (LenB(sidRaw) = 0 Or InStr(sidRaw, "-") = 0) Or (LenB(Trim$(CStr(tableData(rowIndex, costCodeColumn)))) = 0) Or (LenB(Trim$(CStr(tableData(rowIndex,  _
        invoiceDateColumn)))) = 0) Or (LenB(Trim$(CStr(tableData(rowIndex, taxCodeColumn)))) = 0) Or (StrComp(Trim$(CStr(tableData(rowIndex, itemCodeColumn))), "Itemcode", vbTextCompare) = 0)
End Function
