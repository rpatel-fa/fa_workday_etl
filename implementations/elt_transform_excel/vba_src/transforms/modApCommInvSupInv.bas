Attribute VB_Name = "modApCommInvSupInv"
Option Explicit

' Source of truth: config/transform_specs/ap_commercial_invoice.yaml
'   mapping SOT the .bas mirrors BY HAND (no auto-regen);
'   build_runner.py enforces spec<->code column agreement.

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

'  -- source layout: INTEGRATION tab, header=1, data=2 --
'  Values are read from Runner config at runtime (Source Worksheet Name / Resolution
'  / Header Row / Data Start Row). Hard-coded constants retained for documentation
'  audit only; NOT used at runtime:
'Private Const SOURCE_WORKSHEET_NAME As String = "INTEGRATION"
'Private Const SOURCE_HEADER_ROW As Long = 1
'Private Const SOURCE_DATA_START_ROW As Long = 2

'  -- per-period constants (Rules 1/2) --
'  Month-on-month maintenance:
'  (1) update BATCH_ID_TOKEN below to match Integration col A first-row value;
'  (2) all other fields read from Integration tab (dates, docno, creditor, amounts).
Private Const SOURCE_COMPANY_CODE As String = "FA"
Private Const BATCH_ID_TOKEN As String = "EFAC 2RP 2526"

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"
Private Const OUTPUT_TAX_APPLICABILITY As String = "GBR_Outside_of_Scope"
Private Const OUTPUT_TAX_CODE As String = "GBR_Zero_Rated"
Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_RECOVERABILITY As String = "Nonrecoverable_Allocate"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

Private mCompanyId As String
Private mSupplierLookup As Collection
Private mSpendCategoryLookup As Collection
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection
' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

Public Sub RunApCommInvSupInv(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim optionalWorktagBlankRows As Long
    Dim rowIndex As Long
    '  -- Integration column indices (11 cols A-K, no Company col) --
    '  Col A = BatchID, B = CreditorID, C = DocumentType, D = Decscription (typo 'c'),
    '  E = InvoiceDate, F = ReceiptDate, G = DocNo, H = DistributionReferenceCode,
    '  I = PurchasesAmount (header has leading + trailing space; NormalizeToken trims),
    '  J = TaxAmount, K = TaxDetailID.
    Dim batchIdColumn As Long
    Dim creditorIdColumn As Long
    Dim documentTypeColumn As Long
    Dim descriptionColumn As Long
    Dim invoiceDateColumn As Long
    Dim receiptDateColumn As Long
    Dim docNoColumn As Long
    Dim distributionRefCodeColumn As Long
    Dim purchasesAmountColumn As Long
    Dim taxAmountColumn As Long
    Dim taxDetailColumn As Long
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

    Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    Set mSupplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    supplierWorkbook.Close SaveChanges:=False
    Set supplierWorkbook = Nothing

    '  Integration headers A..K (11 cols -- no Company column). Header D is
    '  "Decscription" typo with extra 'c'; header I has leading + trailing space
    '  in the source header " PurchasesAmount ". NormalizeToken handles both.
    requiredHeaders = ResolveRequiredHeadersFromConfig( _
        controlSheet, "modApCommInvSupInv", 11)

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
        Err.Raise vbObjectError + 1910, _
            "modApCommInvSupInv.RunApCommInvSupInv", _
            schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    batchIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "BatchID")
    creditorIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "CreditorID")
    documentTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "DocumentType")
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Decscription")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "InvoiceDate")
    receiptDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "ReceiptDate")
    docNoColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "DocNo")
    distributionRefCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "DistributionReferenceCode")
    purchasesAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "PurchasesAmount")
    taxAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "TaxAmount")
    taxDetailColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "TaxDetailID")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    ' Fresh invoice-line grouping state for this run (Runner persists across runs).
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    companyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    Dim consecutiveEmptyRows As Long
    consecutiveEmptyRows = 0

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then _
            UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, _
                UBound(sourceTable, 1), runStartTime

        If RowIsEmptyOrTrailer(sourceTable, rowIndex, _
            Array(creditorIdColumn, documentTypeColumn, descriptionColumn, _
                docNoColumn, distributionRefCodeColumn), _
            Array(purchasesAmountColumn), _
            Array(descriptionColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(creditorIdColumn, documentTypeColumn, descriptionColumn, _
                    docNoColumn, distributionRefCodeColumn), _
                Array(purchasesAmountColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 3 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsSkippable(sourceTable, rowIndex, creditorIdColumn, _
            documentTypeColumn, descriptionColumn, docNoColumn, _
            distributionRefCodeColumn, purchasesAmountColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required Integration field (CreditorID blank / " _
                    & "DocumentType UPPER<>INVOICE / Decscription blank / " _
                    & "DocNo blank / DistributionReferenceCode blank / " _
                    & "PurchasesAmount=0)", _
                Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
        Else
            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
                costCenterLookup, sourceTable(rowIndex, distributionRefCodeColumn))
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Cost centre unresolved: DistributionReferenceCode did not resolve in '" _
                        & mappingSheetName & "' (" & WBO_COST_CENTER _
                        & ", Core vs Prism = Core, lowest hierarchy match)", _
                    Trim$(CStr(sourceTable(rowIndex, distributionRefCodeColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
            Else
                spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
                    mSpendCategoryLookup, sourceTable(rowIndex, distributionRefCodeColumn))
                If LenB(spendCategoryId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Spend Category unresolved: DistributionReferenceCode did not resolve in '" _
                            & mappingSheetName & "' (" & WBO_SPEND_CATEGORY _
                            & ", Core vs Prism = Core, lowest hierarchy match)", _
                        Trim$(CStr(sourceTable(rowIndex, distributionRefCodeColumn))), _
                        Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                Else
                    supplierId = ResolveSupplierWithCompanyPrecedence( _
                        mSupplierLookup, _
                        Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                        SOURCE_COMPANY_CODE)
                    If LenB(supplierId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unresolvedSupplierRows = unresolvedSupplierRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                            "Supplier unresolved: CreditorID did not resolve via the Supplier Spreadsheet" _
                                & " (Reference ID -> Supplier ID, two-part company precedence:" _
                                & " CreditorID ""-"" FA first, then bare CreditorID." _
                                & " Company=FA constant.)", _
                            Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                            "Company=" & SOURCE_COMPANY_CODE & " | " _
                                & Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                    Else
                        If WriteOutputRow( _
                            targetSheet, targetRow, rowsProcessed + 1, _
                            sourceTable, rowIndex, _
                            creditorIdColumn, documentTypeColumn, descriptionColumn, _
                            invoiceDateColumn, receiptDateColumn, docNoColumn, _
                            distributionRefCodeColumn, purchasesAmountColumn, _
                            costCenterId, spendCategoryId, supplierId, _
                            optionalWorktagBlankRows) _
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
            & " source row(s) with missing CreditorID, DocumentType (not INVOICE)," _
            & " Decscription, DocNo, DistributionReferenceCode, or PurchasesAmount=0."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " source row(s) whose CreditorID did not resolve to a Workday Supplier ID" _
            & " via the Supplier Spreadsheet (Reference ID -> Supplier ID, constant FA company)."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) _
            & " source row(s) whose DistributionReferenceCode did not resolve to a Workday cost centre" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) _
            & " source row(s) whose DistributionReferenceCode did not resolve to a Workday spend category" _
            & " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) _
            & " source row(s) had 1 or more blank optional worktags" _
            & " (Event/Camp/FA Learning/Commercial Partner/Grant/Program/Project)" _
            & " -- DistributionReferenceCode did not fully map to all 7 worktag hierarchies."
    End If

    outputPath = outputWorkbook.FullName
    outputWorkbook.Save
    outputWorkbook.Close SaveChanges:=False
    Set outputWorkbook = Nothing

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Completed", _
        rowsProcessed, rowsSkipped, runStartTime, runEndTime, outputPath, _
        JoinCollectionValues(issueParts, " | "))
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
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Failed", _
        rowsProcessed, rowsSkipped, runStartTime, runEndTime, "", failureMessage)
    If LenB(logIssue) > 0 Then failureMessage = failureMessage & " | " & logIssue
    EndRun

    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Failed"
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, failureMessage
    WriteControlValue controlSheet, RESULT_LABEL_SELECTION_KEY, selectionKey
    WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error Resume Next
    runnerWorkbook.Save
    On Error GoTo 0
    Err.Raise failureNumber, _
        "modApCommInvSupInv.RunApCommInvSupInv", _
        failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row.
'  INPUTS:
'     creditorIdColumn / documentTypeColumn / descriptionColumn /
'     invoiceDateColumn / receiptDateColumn / docNoColumn /
'     distributionRefCodeColumn / purchasesAmountColumn
'        -- 1-based column indices into sourceTable (Integration tab).
'  OUTPUT COLUMN MAPPING (matches mapping: in transform_specs/ap_commercial_invoice.yaml
'     -- Submit_Supplier_Invoice_v46.2):
'     B  seq                           (sequenceNumber)
'     I  submit                        (Y)
'     K  InvoiceNumber                 (Rule 8 -- EFAC 2RP 2526-{CreditorID};
'                                        DocNo is constant, batch+cred required for
'                                        per-row uniqueness)
'     R  Company                       (mCompanyId constant "FA")
'     T  Currency                      (GBP default)
'     U  Supplier                      (supplierId from ResolveSupplierWithCompanyPrecedence)
'     AF InvoiceDate                   (per-row Int E via NormalizeSourceDate)
'     AG ReceiptDate                   (per-row Int F via NormalizeSourceDate)
'     BG SuppliersInvoiceNumber        (Rule 7 -- DocNo + "-" + CreditorID)
'     EK Row ID                        (1 -- single-line invoice per Rule 19)
'     EP ItemDescription               (UPPER(TRIM(Int D Decscription));
'                                        typo 'c' preserved, matches spec header)
'     EW Spend Category                (DistributionRefCode low-hierarchy match;
'                                        on_fail=reject, gated upstream)
'     FE Applicability / FF Code / FH date-type / FK Recoverability / FL CALC_TAX_DUE
'                                       (uniform tax defaults)
'     GE Unit Cost / GF Extended Amount
'                                       (PurchasesAmount Int I; no split)
'     GQ Memo                          (traceability: DocNo + " " + CreditorID +
'                                        " " + DistributionReferenceCode)
'     GR Reporting Type                (second segment first char via
'                                        ResolveReportingTypeFromFinalCode)
'     GT Cost Centre                   (DistributionRefCode low-hierarchy match;
'                                        on_fail=reject, gated upstream)
'     GV GX GZ HB HD HF HH             (7 optional worktags -- Event Camp
'                                        FALearning CommercialPartner Grant
'                                        Program Project)
'  --------------------------------------------------------------
Private Function WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal creditorIdColumn As Long, _
    ByVal documentTypeColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal receiptDateColumn As Long, _
    ByVal docNoColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal purchasesAmountColumn As Long, _
    ByVal costCenterId As String, _
    ByVal spendCategoryId As String, _
    ByVal supplierId As String, _
    ByRef optionalWorktagBlankRows As Long _
) As Boolean
    Dim amount As Double
    Dim distributionRefCodeValue As Variant
    Dim docNoValue As String
    Dim creditorIdValue As String
    Dim invoiceNumberValue As String
    Dim suppliersInvoiceNumberValue As String
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

    creditorIdValue = Trim$(CStr(sourceTable(sourceRow, creditorIdColumn)))
    docNoValue = Trim$(CStr(sourceTable(sourceRow, docNoColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, descriptionColumn))))
    amount = ResolveAmount(sourceTable(sourceRow, purchasesAmountColumn))
    distributionRefCodeValue = sourceTable(sourceRow, distributionRefCodeColumn)
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    receiptDateValue = NormalizeSourceDate(sourceTable(sourceRow, receiptDateColumn))

    '  Rule 8 -- Invoice Number G = {BATCH_ID_TOKEN}-{CreditorID}.
    '  DocNo is NOT unique per row (all rows share "EFAC 2RP 2526" in the exemplar),
    '  so bare DocNo would fail Workday's per-supplier unique inv-no check. Creditor
    '  suffix disambiguates -- each commercial supplier has exactly one row per batch.
    invoiceNumberValue = BATCH_ID_TOKEN & "-" & creditorIdValue

    '  Rule 7 -- Suppliers Invoice Number (BG / voucher no) = {DocNo}-{CreditorID}.
    suppliersInvoiceNumberValue = docNoValue & "-" & creditorIdValue

    eventId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mEventLookup, distributionRefCodeValue)
    campId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mCampLookup, distributionRefCodeValue)
    faLearningId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mFaLearningLookup, distributionRefCodeValue)
    commercialPartnerId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mCommercialPartnerLookup, distributionRefCodeValue)
    grantId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mGrantLookup, distributionRefCodeValue)
    programId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mProgramLookup, distributionRefCodeValue)
    projectId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
        mProjectLookup, distributionRefCodeValue)

    anyWorktagBlank = (LenB(eventId) = 0) Or (LenB(campId) = 0) _
        Or (LenB(faLearningId) = 0) Or (LenB(commercialPartnerId) = 0) _
        Or (LenB(grantId) = 0) Or (LenB(programId) = 0) _
        Or (LenB(projectId) = 0)
    If anyWorktagBlank Then
        optionalWorktagBlankRows = optionalWorktagBlankRows + 1
    End If

    ' Invoice-line grouping: B increments per NEW Invoice Number; EM/EK count 1..N within group.
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long
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
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = receiptDateValue
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = suppliersInvoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue    ' Row ID = 1..N within invoice group (was hardcoded 1)
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue    ' Line Order = 1..N within invoice group (mirror of EK; was not written)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = OUTPUT_TAX_APPLICABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = OUTPUT_TAX_CODE
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = OUTPUT_TAX_RECOVERABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = docNoValue & " " _
        & creditorIdValue & " " & CStr(distributionRefCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = _
        ResolveReportingTypeFromFinalCode(distributionRefCodeValue)
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
'  RowIsSkippable -- commercial Integration gate (matches skip_rule.drop_when in spec).
'     CreditorID blank / DocumentType UPPER<>INVOICE / Decscription blank /
'     DocNo blank / DistributionReferenceCode blank / PurchasesAmount = 0.
'     NO Company-blank gate (no Company column in Commercial Integration tab).
'  --------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal creditorIdColumn As Long, _
    ByVal documentTypeColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal docNoColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal purchasesAmountColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, creditorIdColumn)))) = 0 Or _
        StrComp(Trim$(CStr(tableData(rowIndex, documentTypeColumn))), _
            "INVOICE", vbTextCompare) <> 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, descriptionColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, docNoColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, distributionRefCodeColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, purchasesAmountColumn)) = 0#
End Function
