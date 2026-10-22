Attribute VB_Name = "modApCfaDistribSupplierInvoice"
Option Explicit

' Invoice-line grouping state (shared modInvoiceGrouping): reset per run, keyed on Supplier + Invoice Number.
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Source of truth: config/transform_specs/ap_cfa_distributions.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

' REWORKED 2026-10-22 for the October-2026 "Finance Master sheet" tab. INVOICES ONLY:
' processes rows where col E "Document type" = "invoice"; other non-blank document types are
' skipped here (reserved for credit notes / modApCfaDistribAmendment). Supplier = col D
' "Old Creditor ID" looked up in the Supplier extract (CreditorID-FA, then bare CreditorID).
' Invoice Number + Item Description = col F "Description - remittance"; Extended Amount = col H
' "Total" formatted "0.00", not rounded (Unit Cost left blank); Invoice/Receipt Date = col K "Batch ID".
' Worktags/Reporting Type still resolve from Distribution/Cost Code (col G) via Mapping MASTER;
' tax stays conditional on Tax amount (I).

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

' Company is a constant "FA" (CFA always = THE FA GROUP); also the supplier-key company suffix.
Private Const CONSTANT_COMPANY_CODE As String = "FA"
Private Const DOCUMENT_TYPE_INVOICE As String = "INVOICE"
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"
Private Const GROUP_KEY_SEPARATOR As String = "|"

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
    Dim nonInvoiceRows As Long
    Dim rowIndex As Long
    '  -- Finance Master sheet column indices (header row 5) --
    Dim docNoColumn As Long                  ' A  DocNo / row key (headerless; = CFA_col - 1); memo trace
    Dim cfaNameColumn As Long                ' B  County FA name (anchor)
    Dim creditorIdColumn As Long             ' D  Old Creditor ID (Supplier extract lookup key)
    Dim documentTypeColumn As Long           ' E  Document type ("invoice" here; else skipped)
    Dim descriptionColumn As Long            ' F  Description - remittance (-> EIB K + BG + EP)
    Dim distributionRefCodeColumn As Long    ' G  Distribution/Cost Code (accounting string)
    Dim totalColumn As Long                  ' H  Total (-> EIB GF, formatted "0.00")
    Dim taxAmountColumn As Long              ' I  Tax amount (conditional switch)
    Dim batchIdColumn As Long                ' K  Batch ID (-> EIB AF/AG)
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
    EnsureNonBlank supplierSpreadsheetPath, CONTROL_LABEL_SUPPLIER_SPREADSHEET
    If LenB(mappingSheetName) = 0 Then mappingSheetName = DEFAULT_MAPPING_SHEET_NAME

    ' Supplier = col D "Old Creditor ID" resolved via the Supplier extract (Reference ID -> Supplier ID).
    Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    Set mSupplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    supplierWorkbook.Close SaveChanges:=False
    Set supplierWorkbook = Nothing

    '  Source of truth: config/transforms/ap_cfa_distributions.yaml source_layout.required_headers.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApCfaDistribSupplierInvoice", 8)

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
    creditorIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Old Creditor ID")
    documentTypeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Document type")
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description - remittance")
    distributionRefCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Distribution/Cost Code")
    totalColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Total")
    taxAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Tax amount")
    batchIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Batch ID")

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
            Array(docNoColumn, creditorIdColumn, documentTypeColumn, descriptionColumn), _
            Array(totalColumn), _
            Array(docNoColumn)) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(docNoColumn, creditorIdColumn, documentTypeColumn, descriptionColumn), _
                Array(totalColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 8 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsNonInvoiceType(sourceTable, rowIndex, documentTypeColumn) Then
            '  -- Non-invoice document types (credit notes): reserved for
            '     ap_cfa_distributions_amendments. Skip silently here (not a reject) so the two
            '     1-input->1-output transforms partition the same source by document type.
            nonInvoiceRows = nonInvoiceRows + 1
            consecutiveEmptyRows = 0
        ElseIf RowIsSkippable(sourceTable, rowIndex, docNoColumn, creditorIdColumn, _
            documentTypeColumn, descriptionColumn, distributionRefCodeColumn, totalColumn, _
            batchIdColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required invoice field (DocNo/A / Old Creditor ID/D / Document type/E / " _
                    & "Description-remittance/F / Distribution-Cost-Code/G / Batch-ID/K blank, " _
                    & "or Total/H = 0)", _
                Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
        Else
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
                    supplierId = ResolveSupplierWithCompanyPrecedence( _
                        mSupplierLookup, _
                        Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                        CONSTANT_COMPANY_CODE)
                    If LenB(supplierId) = 0 Then
                        rowsSkipped = rowsSkipped + 1
                        unresolvedSupplierRows = unresolvedSupplierRows + 1
                        AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                            "Supplier unresolved: Old Creditor ID (col D) did not resolve via the Supplier Spreadsheet" _
                                & " (Reference ID -> Supplier ID, two-part company precedence:" _
                                & " CreditorID ""-"" FA first, then bare CreditorID.)", _
                            Trim$(CStr(sourceTable(rowIndex, creditorIdColumn))), _
                            Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                    Else
                        If WriteOutputRow( _
                            targetSheet, targetRow, sourceTable, rowIndex, _
                            docNoColumn, distributionRefCodeColumn, totalColumn, taxAmountColumn, _
                            descriptionColumn, batchIdColumn, _
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
            & " invoice row(s) with a missing required field (DocNo/A, Old Creditor ID/D, Document type/E," _
            & " Description - remittance/F, Distribution/Cost Code/G, Batch ID/K, or Total/H = 0)."
    End If
    If nonInvoiceRows > 0 Then
        issueParts.Add CStr(nonInvoiceRows) _
            & " row(s) with Document type <> 'invoice' (credit notes) were skipped here -- they are" _
            & " reserved for ap_cfa_distributions_amendments (Supplier Invoice Adjustment EIB)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) _
            & " invoice row(s) whose col D ""Old Creditor ID"" did not resolve in the Supplier Spreadsheet" _
            & " (tried CreditorID-FA, then bare CreditorID)."
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
    If Not supplierWorkbook Is Nothing Then supplierWorkbook.Close SaveChanges:=False
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
'     B  Spreadsheet Key (AssignInvoiceLineGroup on Supplier + "|" + Invoice Number)
'     I  Submit (Y)
'     K  Invoice Number (col F Description - remittance)
'     R  Company (constant FA -> THE FA GROUP - FA)
'     T  Currency (GBP)
'     U  Supplier (col D Old Creditor ID via Supplier extract; resolved upstream)
'     AF Invoice Date / AG Received Date (col K Batch ID via NormalizeSourceDate)
'     BG Suppliers Invoice Number (1:1 mirror of K)
'     EK Row ID / EM Line Order (1..N within Supplier + Invoice Number group)
'     EP Item Description (col F Description - remittance)
'     EW Spend Category (col G lowest-hierarchy match; gated upstream)
'     FE FF FH FK FL tax conditional_on_tax_amount (col I > 0 -> Standard, else Zero)
'     GE Unit Cost -- intentionally NOT written (left blank)
'     GF Extended Amount (col H Total as is, NumberFormat "0.00" -- display only, not rounded)
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
    ByVal totalColumn As Long, _
    ByVal taxAmountColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal batchIdColumn As Long, _
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

    ' Col F feeds Invoice Number (K/BG) and Item Description (EP).
    invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, descriptionColumn)))
    descriptionValue = invoiceNumberValue
    amount = ResolveAmount(sourceTable(sourceRow, totalColumn))
    taxAmount = ResolveAmount(sourceTable(sourceRow, taxAmountColumn))
    distributionRefCodeValue = sourceTable(sourceRow, distributionRefCodeColumn)
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, batchIdColumn))

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

    ' Col F repeats across suppliers, so group on Supplier + Invoice Number (not Invoice Number alone).
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, _
        supplierId & GROUP_KEY_SEPARATOR & invoiceNumberValue, mGrpNextKey, lineOrderValue)
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
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = AMOUNT_NUMBER_FORMAT
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
'  RowIsNonInvoiceType -- TRUE when col E "Document type" is present and NOT "invoice"
'  (case-insensitive; e.g. a credit note). Such rows are reserved for
'  ap_cfa_distributions_amendments. A blank Document type is NOT treated as non-invoice here
'  (it falls through to the missing-field gate below so genuine data gaps are still reported).
'  --------------------------------------------------------------
Private Function RowIsNonInvoiceType( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal documentTypeColumn As Long _
) As Boolean
    Dim v As String
    v = Trim$(CStr(tableData(rowIndex, documentTypeColumn)))
    If LenB(v) = 0 Then
        RowIsNonInvoiceType = False
    Else
        RowIsNonInvoiceType = (StrComp(UCase$(v), DOCUMENT_TYPE_INVOICE, vbBinaryCompare) <> 0)
    End If
End Function

'  --------------------------------------------------------------
'  RowIsSkippable -- per-row missing-field gate (matches skip_rule in the spec YAML).
'  DocNo/A, Old Creditor ID/D, Document type/E, Description - remittance/F,
'  Distribution/Cost Code/G, Batch ID/K blank, OR Total/H = 0.
'  --------------------------------------------------------------
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal docNoColumn As Long, _
    ByVal creditorIdColumn As Long, _
    ByVal documentTypeColumn As Long, _
    ByVal descriptionColumn As Long, _
    ByVal distributionRefCodeColumn As Long, _
    ByVal totalColumn As Long, _
    ByVal batchIdColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, docNoColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, creditorIdColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, documentTypeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, descriptionColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, distributionRefCodeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, batchIdColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, totalColumn)) = 0#
End Function
