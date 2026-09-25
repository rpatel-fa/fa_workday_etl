Attribute VB_Name = "modApCfaClaimsSupplierInvoice"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_cfa_claims.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

' REWORKED 2026-09-24 for the pre-resolved September-2026 CFA Claims source format.
' Worktags arrive PRE-RESOLVED as discrete Workday-ref-id columns (E..O) and Invoice Number /
' Description / Date are pre-computed (Y / Z / AA). The worktag pipeline is INACTIVE: no Mapping
' MASTER lookup, no WBO_* constants, no ResolveReportingTypeFromFinalCode, no accounting-string
' normalisation, no batch-token / per-group invoice counter. Pass-through only. Mapping MASTER
' existence of worktag values is validated OUT OF BAND (authoring / reconcile), not here.

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"

'  transform_constants.uniform_tax (CFA Claims carry NO tax columns; ALL rows = zero-rated branch).
'  Synced 1:1 to uniform_tax.applicability/code/recoverability in transforms/ap_cfa_claims.yaml.
Private Const OUTPUT_TAX_APPLICABILITY As String = "GBR_Outside_of_Scope"
Private Const OUTPUT_TAX_CODE As String = "GBR_Zero_Rated"
Private Const OUTPUT_TAX_RECOVERABILITY As String = "Nonrecoverable_Allocate"

Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
' GE / GF cell display format (transform_specs mapping number_format; display only, value not rounded).
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"

' Company is a constant "FA" (claims_constants.constant_company_code). Claims tab B = County FA
' NAME, not legal-entity code; the paying legal entity is always The FA Group.
Private Const CONSTANT_COMPANY_CODE As String = "FA"

Private mCompanyId As String

Public Sub RunApCfaClaimsSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As Variant
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unresolvedSupplierRows As Long
    Dim rowIndex As Long
    '  -- Claims (Sep-2026 pre-resolved layout) column indices --
    Dim sheetNameColumn As Long                 ' A  claim category (memo + skip gate)
    Dim countyFaColumn As Long                  ' B  claimant entity (skip gate; informational)
    Dim gpIdColumn As Long                      ' C  County FA GP ID (trace + skip gate only)
    Dim purchasesAmountColumn As Long           ' D  Total Claim (GBP) -> PurchasesAmount
    Dim costCentreColumn As Long                ' E  PRE-RESOLVED Cost Centre (mandatory)
    Dim reportingTypeColumn As Long             ' F  PRE-RESOLVED Reporting Type (display name)
    Dim spendCategoryColumn As Long             ' H  PRE-RESOLVED Spend Category (optional)
    Dim eventColumn As Long                     ' I  PRE-RESOLVED Event worktag (optional)
    Dim faLearningColumn As Long                ' J  PRE-RESOLVED FA Learning worktag (optional)
    Dim campColumn As Long                      ' K  PRE-RESOLVED Camp worktag (optional)
    Dim programColumn As Long                   ' L  PRE-RESOLVED Program worktag (optional)
    Dim grantColumn As Long                     ' M  PRE-RESOLVED Grant worktag (optional)
    Dim commercialPartnerColumn As Long         ' N  PRE-RESOLVED Commercial Partner worktag (optional)
    Dim projectColumn As Long                   ' O  PRE-RESOLVED Project worktag (optional)
    Dim headOfApprovalColumn As Long            ' Q  informational (memo)
    Dim processedByColumn As Long               ' R  informational (memo)
    Dim processedDateColumn As Long             ' S  trace / integrity gate
    Dim budgetCodeNamingColumn As Long          ' U  Finance label (memo)
    Dim workdayIdColumn As Long                 ' W  Workday Supplier ID Lookup (DIRECT passthrough)
    Dim invoiceNumberColumn As Long             ' Y  PRE-COMPUTED Invoice Number
    Dim invoiceDescriptionColumn As Long        ' Z  PRE-COMPUTED Invoice Description
    Dim invoiceDateColumn As Long               ' AA PRE-COMPUTED Invoice Date
    Dim companyId As String
    Dim supplierId As String
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

    ' Supplier resolution is DIRECT passthrough from Claims col W ("Workday Supplier ID Lookup").
    ' No Supplier Spreadsheet is opened. Worktags are pre-resolved in source, so NO Mapping MASTER
    ' is opened either (pipeline inactive).

    '  Source of truth: config/transforms/ap_cfa_claims.yaml source_layout.required_headers.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApCfaClaimsSupplierInvoice", 15)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    Set sourceSheet = ResolveSourceWorksheet(sourceWorkbook, sheetToResolve)
    sourceTable = ReadWorksheetTable(sourceSheet)
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

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

    sheetNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Sheet Name")
    countyFaColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "County FA")
    gpIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "County FA GP ID")
    purchasesAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Total Claim ()|Total Claim|Total Claim (|Total Claim (GBP)")
    costCentreColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Cost Centre")
    reportingTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Reporting Type")
    spendCategoryColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Spend Category")
    eventColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Events - Custom Org")
    faLearningColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "FA Learning - Custom Org")
    campColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Camps - Custom Org")
    programColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Program - Delivered Org")
    grantColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Grant")
    commercialPartnerColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Commercial Partners - Custom Worktag")
    projectColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Projects (Opex)")
    headOfApprovalColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Head of Approval")
    processedByColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Processed By")
    processedDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Processed Date")
    budgetCodeNamingColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Budget Code Naming")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Workday Supplier ID Lookup")
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

    companyId = ResolveMappedCompany(runnerWorkbook, CONSTANT_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, _
            Array(sheetNameColumn, gpIdColumn, costCentreColumn, budgetCodeNamingColumn, workdayIdColumn), _
            Array(purchasesAmountColumn), _
            Array(sheetNameColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(sheetNameColumn, gpIdColumn, costCentreColumn, budgetCodeNamingColumn, workdayIdColumn), _
                Array(purchasesAmountColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 8 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsSilentTrailerAdminClaims(sourceTable, rowIndex, _
            countyFaColumn, costCentreColumn, headOfApprovalColumn, workdayIdColumn) Then
            ' Intentionally blank: trailing admin / summary footer row -- skip silently.
            consecutiveEmptyRows = 0
        ElseIf RowIsSkippable(sourceTable, rowIndex, _
            sheetNameColumn, gpIdColumn, processedDateColumn, purchasesAmountColumn, _
            budgetCodeNamingColumn, invoiceNumberColumn, invoiceDateColumn, _
            costCentreColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required Claims field (Sheet Name/A / County-FA-GP-ID/C / " _
                    & "Processed-Date/S / Budget-Code-Naming/U / Invoice-Number/Y / " _
                    & "Invoice-Date/AA blank, Total-Claim/D <=0, or mandatory pre-resolved " _
                    & "Cost-Centre/E blank)", _
                Trim$(CStr(sourceTable(rowIndex, gpIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
        Else
            supplierId = Trim$(CStr(sourceTable(rowIndex, workdayIdColumn)))
            If LenB(supplierId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unresolvedSupplierRows = unresolvedSupplierRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Supplier unresolved: Workday Supplier ID Lookup (Claims col W) was BLANK. " _
                        & "AP/Finance curate the Workday Supplier Reference (S-XXXXXX) directly in " _
                        & "col W; blank means the supplier is not yet allocated. County FA GP ID " _
                        & "(col C, for traceability): " & Trim$(CStr(sourceTable(rowIndex, gpIdColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, gpIdColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
            Else
                If WriteOutputRow(targetSheet, targetRow, sourceTable, rowIndex, supplierId, _
                    sheetNameColumn, gpIdColumn, purchasesAmountColumn, _
                    costCentreColumn, reportingTypeColumn, spendCategoryColumn, _
                    eventColumn, faLearningColumn, campColumn, programColumn, grantColumn, _
                    commercialPartnerColumn, projectColumn, _
                    processedByColumn, budgetCodeNamingColumn, _
                    invoiceNumberColumn, invoiceDescriptionColumn, invoiceDateColumn) Then
                    rowsProcessed = rowsProcessed + 1
                    targetRow = targetRow + 1
                Else
                    rowsSkipped = rowsSkipped + 1
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) _
            & " source row(s) with a missing required field (Sheet Name/A, County FA GP ID/C, " _
            & "Processed Date/S, Budget Code Naming/U, Invoice Number/Y, Invoice Date/AA, " _
            & "Total Claim/D <=0, or mandatory pre-resolved Cost Centre/E blank)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " source row(s) whose Claims col W ""Workday Supplier ID Lookup"" was BLANK. " _
            & "AP/Finance must populate the Workday Supplier ID on every row."
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
'  Write one EIB output row (pre-resolved pass-through model).
'  OUTPUT COLUMN MAPPING (matches mapping: in config/transform_specs/ap_cfa_claims.yaml):
'     B  Spreadsheet Key      (AssignInvoiceLineGroup on Invoice Number)
'     I  Submit               (Y)
'     K  Invoice Number       (col Y verbatim)
'     R  Company              (mCompanyId from constant "FA")
'     T  Currency             (GBP)
'     U  Supplier             (col W Workday Supplier ID Lookup passthrough)
'     AF Invoice Date         (col AA via NormalizeSourceDate)
'     AG Invoice Received Date(col AA; = Invoice Date)
'     BG Suppliers Invoice No (1:1 mirror of K)
'     EK Row ID / EM Line Order (1..N within Invoice Number group)
'     EP Item Description      (col Z verbatim)
'     EW Spend Category        (col H passthrough; optional - blank H writes EW blank)
'     FE FF FH FK FL tax       (uniform zero-rated constants)
'     GE Unit Cost / GF Extended Amount (col D; NumberFormat "0.00")
'     GQ Memo                  ("{C} | {A} | {U} | by {R}")
'     GR Reporting Type        (col F passthrough)
'     GT Cost Centre           (col E passthrough; mandatory)
'     GV Event / GX Camp / GZ FA Learning / HB Commercial Partner /
'     HD Grant / HF Program / HH Project (cols I/K/J/N/M/L/O passthrough; optional)
'  --------------------------------------------------------------
Private Function WriteOutputRow( _
    ByVal targetSheet As Worksheet, ByVal targetRow As Long, _
    ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal supplierId As String, _
    ByVal sheetNameColumn As Long, ByVal gpIdColumn As Long, ByVal purchasesAmountColumn As Long, _
    ByVal costCentreColumn As Long, ByVal reportingTypeColumn As Long, ByVal spendCategoryColumn As Long, _
    ByVal eventColumn As Long, ByVal faLearningColumn As Long, ByVal campColumn As Long, _
    ByVal programColumn As Long, ByVal grantColumn As Long, ByVal commercialPartnerColumn As Long, _
    ByVal projectColumn As Long, ByVal processedByColumn As Long, ByVal budgetCodeNamingColumn As Long, _
    ByVal invoiceNumberColumn As Long, ByVal invoiceDescriptionColumn As Long, ByVal invoiceDateColumn As Long _
) As Boolean
    Dim amount As Double
    Dim invoiceNumberValue As String
    Dim descriptionValue As String
    Dim invoiceDateValue As String
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long

    invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = Trim$(CStr(sourceTable(sourceRow, invoiceDescriptionColumn)))
    amount = ResolveAmount(sourceTable(sourceRow, purchasesAmountColumn))
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))

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
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = Trim$(CStr(sourceTable(sourceRow, spendCategoryColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = OUTPUT_TAX_APPLICABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = OUTPUT_TAX_CODE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = OUTPUT_TAX_RECOVERABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = Trim$(CStr(sourceTable(sourceRow, gpIdColumn))) _
        & " | " & Trim$(CStr(sourceTable(sourceRow, sheetNameColumn))) _
        & " | " & Trim$(CStr(sourceTable(sourceRow, budgetCodeNamingColumn))) _
        & " | by " & Trim$(CStr(sourceTable(sourceRow, processedByColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = Trim$(CStr(sourceTable(sourceRow, reportingTypeColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = Trim$(CStr(sourceTable(sourceRow, costCentreColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = Trim$(CStr(sourceTable(sourceRow, eventColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = Trim$(CStr(sourceTable(sourceRow, campColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = Trim$(CStr(sourceTable(sourceRow, faLearningColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = Trim$(CStr(sourceTable(sourceRow, commercialPartnerColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = Trim$(CStr(sourceTable(sourceRow, grantColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = Trim$(CStr(sourceTable(sourceRow, programColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = Trim$(CStr(sourceTable(sourceRow, projectColumn)))
    WriteOutputRow = True
End Function

' -----------------------------------------------------------------
'  Trailing admin / summary footer detection (skip silently, NOT a reject):
'  ALL FOUR blank -> County FA (B) AND Cost Centre (E) AND Head of Approval (Q) AND
'  Workday Supplier ID Lookup (W).
' -----------------------------------------------------------------
Private Function RowIsSilentTrailerAdminClaims( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal countyFaColumn As Long, _
    ByVal costCentreColumn As Long, _
    ByVal headOfApprovalColumn As Long, _
    ByVal workdayIdColumn As Long _
) As Boolean
    Dim b1 As String, b2 As String, b3 As String, b4 As String
    b1 = Trim$(CStr(tableData(rowIndex, countyFaColumn)))
    b2 = Trim$(CStr(tableData(rowIndex, costCentreColumn)))
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
'  Sheet Name/A blank OR County-FA-GP-ID/C blank OR Processed-Date/S blank OR
'  Budget-Code-Naming/U blank OR Invoice-Number/Y blank OR Invoice-Date/AA blank OR
'  amount D <= 0 OR mandatory pre-resolved Cost-Centre/E blank.
'  Spend Category (H) is OPTIONAL (transform_specs/ap_cfa_claims.yaml, 2026-09-25).
' -----------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal sheetNameColumn As Long, _
    ByVal gpIdColumn As Long, _
    ByVal processedDateColumn As Long, _
    ByVal purchasesAmountColumn As Long, _
    ByVal budgetCodeNamingColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal costCentreColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, sheetNameColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, gpIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, processedDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, budgetCodeNamingColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, purchasesAmountColumn)) <= 0# Or _
        LenB(Trim$(CStr(tableData(rowIndex, costCentreColumn)))) = 0
End Function
