Attribute VB_Name = "modApStatAddLeeWnslSupInv"
Option Explicit

' Source of truth: config/transform_specs/ap_stationery_addisonlee_wnsl.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together

Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const MODULE_NAME As String = "modApStatAddLeeWnslSupInv"

'  -- source layout: pivot tab, header=3, data=4 (read from Runner config at runtime) --
'  Pivot total row that ends the data (transforms profile source_layout.data_end_marker).
Private Const DATA_END_MARKER As String = "Grand Total"

'  -- fixed parties (no SupplierID / Currency column in the pivot) --
Private Const SOURCE_COMPANY_CODE As String = "WNSL"
Private Const SOURCE_SUPPLIER_REFERENCE_ID As String = "ADD001"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - WNSL"

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const OUTPUT_TAX_APPLICABILITY As String = "GBR_Domestic_Purchase_Of_Goods_And_Services"
Private Const OUTPUT_TAX_CODE As String = "GBR_Standard"
Private Const OUTPUT_TAX_RECOVERABILITY As String = "Fully_Recoverable"
Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

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
Public Sub RunApStationeryAddisonLeeWnslSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim supplierLookup As Collection
    Dim costCenterLookup As Collection
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unmappedCostCenterRows As Long
    Dim unmappedSpendCategoryRows As Long
    Dim optionalWorktagBlankRows As Long
    Dim rowIndex As Long
    Dim lastDataRowIndex As Long
    Dim consecutiveEmptyRows As Long

    '  -- pivot column indices (A..G) --
    Dim accountStringColumn As Long           ' A  Row Labels (4-segment accounting string)
    Dim netColumn As Long                     ' B  Sum of Invoice Final Net
    Dim taxColumn As Long                     ' C  Sum of Invoice Total Vat
    Dim grossColumn As Long                   ' D  Sum of Invoice Amount
    Dim descriptionColumn As Long             ' E  Description | DESCRITPION
    Dim invoiceDateColumn As Long             ' F  INVOICE DATE (first data row only; filled down)
    Dim invoiceNumberColumn As Long           ' G  INVOICE NUMBER (first data row only; filled down)

    Dim resolvedCompanyId As String
    Dim supplierId As String
    Dim costCenterId As String
    Dim spendCategoryId As String
    Dim accountStringRaw As String
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

    '  Rule 2: fixed supplier, resolved ONCE with the 2-part company precedence.
    Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    Set supplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    supplierWorkbook.Close SaveChanges:=False
    Set supplierWorkbook = Nothing
    supplierId = ResolveSupplierWithCompanyPrecedence(supplierLookup, SOURCE_SUPPLIER_REFERENCE_ID, SOURCE_COMPANY_CODE)
    If LenB(supplierId) = 0 Then
        Err.Raise vbObjectError + 1913, MODULE_NAME & ".RunApStationeryAddisonLeeWnslSupplierInvoice", _
            "Unable to resolve supplier Reference ID '" & SOURCE_SUPPLIER_REFERENCE_ID & "-" & SOURCE_COMPANY_CODE & _
            "' or '" & SOURCE_SUPPLIER_REFERENCE_ID & "' to a Workday Supplier ID via the Supplier Spreadsheet."
    End If

    '  Rule 1: fixed company.
    resolvedCompanyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)

    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, MODULE_NAME, 7)

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
    schemaMsg = ValidateSourceWorkbookSchema(sheetToResolve, sourceTable, sourceHeaderRow, sourceDataStartRow, requiredHeaders, , schemaHeaderRow, schemaDiag)
    If LenB(schemaMsg) > 0 Then
        Err.Raise vbObjectError + 1910, MODULE_NAME & ".RunApStationeryAddisonLeeWnslSupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    accountStringColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Row Labels")
    netColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Sum of Invoice Final Net")
    taxColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Sum of Invoice Total Vat")
    grossColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Sum of Invoice Amount")
    descriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Description|DESCRITPION")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "INVOICE DATE")
    invoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "INVOICE NUMBER")

    '  data_end_marker: the pivot "Grand Total" row ends the data (it is never a line).
    lastDataRowIndex = UBound(sourceTable, 1)
    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If StrComp(Trim$(CStr(sourceTable(rowIndex, accountStringColumn))), DATA_END_MARKER, vbTextCompare) = 0 Then
            lastDataRowIndex = rowIndex - 1
            Exit For
        End If
    Next rowIndex

    '  fill_down_context_columns: INVOICE DATE / INVOICE NUMBER are typed on the first data row only.
    FillDownInvoiceColumn sourceTable, invoiceDateColumn, dataStartRowIndex, lastDataRowIndex
    FillDownInvoiceColumn sourceTable, invoiceNumberColumn, dataStartRowIndex, lastDataRowIndex

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    '  Rule 18 prepass: Company / Supplier / Currency are fixed, so only InvoiceDate can diverge
    '  within an INVOICE NUMBER group. Fail before any EIB row is written.
    Dim groupInvDateCaptured As New Collection
    Dim pINum As String
    Dim pIDt As String
    For rowIndex = dataStartRowIndex To lastDataRowIndex
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, Array(descriptionColumn), Array(netColumn, grossColumn), Array(accountStringColumn)) Then
        ElseIf RowIsSkippable(sourceTable, rowIndex, accountStringColumn, invoiceNumberColumn, invoiceDateColumn, netColumn, taxColumn, grossColumn) Then
        Else
            pINum = Trim$(CStr(sourceTable(rowIndex, invoiceNumberColumn)))
            pIDt = NormalizeSourceDate(sourceTable(rowIndex, invoiceDateColumn))
            If CollectionKeyExists(groupInvDateCaptured, pINum) = False Then
                groupInvDateCaptured.Add pIDt, pINum
            ElseIf StrComp(CStr(groupInvDateCaptured(pINum)), pIDt, vbBinaryCompare) <> 0 Then
                Err.Raise vbObjectError + 1912, MODULE_NAME & ".RunApStationeryAddisonLeeWnslSupplierInvoice Phase1GroupPrepass", _
                    "Header Consistency Guard failed. INVOICE NUMBER '" & pINum & "' has diverged InvoiceDate(" & _
                    CStr(groupInvDateCaptured(pINum)) & " vs " & pIDt & "). Fix source data FIRST before re-running the Runner."
            End If
        End If
    Next rowIndex

    For rowIndex = dataStartRowIndex To lastDataRowIndex
        If RowIsEmptyOrTrailer(sourceTable, rowIndex, Array(descriptionColumn), Array(netColumn, grossColumn), Array(accountStringColumn)) Then
            '  blank padding row or summation trailer -- not a line; 3 truly empty rows in a row end the data.
            If RowIsTrulyEmpty(sourceTable, rowIndex, Array(accountStringColumn, descriptionColumn), Array(netColumn, grossColumn)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 3 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        ElseIf RowIsSkippable(sourceTable, rowIndex, accountStringColumn, invoiceNumberColumn, invoiceDateColumn, netColumn, taxColumn, grossColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Missing required pivot field (Row Labels blank / " & _
                "INVOICE NUMBER blank / INVOICE DATE blank / Net blank/zero AND Gross-VAT blank/zero).", _
                Trim$(CStr(sourceTable(rowIndex, accountStringColumn))), Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
        Else
            consecutiveEmptyRows = 0
            accountStringRaw = Trim$(CStr(sourceTable(rowIndex, accountStringColumn)))
            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, accountStringRaw)
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Cost centre unresolved: Row Labels did not resolve in '" & mappingSheetName & _
                    "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", accountStringRaw, Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
            Else
                spendCategoryId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mSpendCategoryLookup, accountStringRaw)
                If LenB(spendCategoryId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedSpendCategoryRows = unmappedSpendCategoryRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, "Spend Category unresolved: Row Labels did not resolve in '" & mappingSheetName & _
                        "' (" & WBO_SPEND_CATEGORY & ", Core vs Prism = Core, lowest hierarchy match)", accountStringRaw, Trim$(CStr(sourceTable(rowIndex, descriptionColumn)))
                Else
                    WriteOutputRow targetSheet, targetRow, sourceTable, rowIndex, resolvedCompanyId, supplierId, costCenterId, spendCategoryId, _
                        accountStringColumn, netColumn, taxColumn, grossColumn, descriptionColumn, invoiceDateColumn, invoiceNumberColumn, optionalWorktagBlankRows
                    rowsProcessed = rowsProcessed + 1
                    targetRow = targetRow + 1
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with a missing required pivot field (Row Labels / " & _
            "INVOICE NUMBER / INVOICE DATE blank, or Net and Gross-VAT both blank/zero)."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) & " source row(s) whose Row Labels did not resolve to a Workday cost centre" & _
            " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If unmappedSpendCategoryRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedSpendCategoryRows) & " source row(s) whose Row Labels did not resolve to a Workday spend category" & _
            " in '" & mappingSheetName & "' (lowest hierarchy match, Core vs Prism = Core)."
    End If
    If optionalWorktagBlankRows > 0 Then
        issueParts.Add CStr(optionalWorktagBlankRows) & " source row(s) had 1 or more blank optional worktags" & _
            " (Event / Camp / FA Learning / Commercial Partner / Grant / Program / Project)" & _
            " -- Row Labels did not fully map to all 7 worktag hierarchies."
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
    Err.Raise failureNumber, MODULE_NAME & ".RunApStationeryAddisonLeeWnslSupplierInvoice", failureMessage
End Sub

'  --------------------------------------------------------------
'  Write one EIB output row (column mapping = mapping: in the transform spec).
'  --------------------------------------------------------------
Private Sub WriteOutputRow(ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal sourceTable As Variant, ByVal sourceRow As Long, _
    ByVal resolvedCompanyId As String, ByVal supplierId As String, ByVal costCenterId As String, ByVal spendCategoryId As String, _
    ByVal accountStringColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, ByVal grossColumn As Long, _
    ByVal descriptionColumn As Long, ByVal invoiceDateColumn As Long, ByVal invoiceNumberColumn As Long, ByRef optionalWorktagBlankRows As Long)
    Dim netAmt As Double
    Dim accountStringNormalized As String
    Dim suppliersInvoiceNumberValue As String
    Dim invoiceNumberCompositeValue As String
    Dim invoiceDateValue As String
    Dim invoiceDateYYMM As String
    Dim descriptionValue As String
    Dim eventId As String
    Dim campId As String
    Dim faLearningId As String
    Dim commercialPartnerId As String
    Dim grantId As String
    Dim programId As String
    Dim projectId As String
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long

    suppliersInvoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
    descriptionValue = UCase$(Trim$(CStr(sourceTable(sourceRow, descriptionColumn))))

    '  Rule 10: Net (Gross - VAT fallback), absolute value, rounded to 2dp (pivot sums carry float noise).
    netAmt = ResolveEffectiveNet(sourceTable(sourceRow, netColumn), sourceTable(sourceRow, grossColumn), sourceTable(sourceRow, taxColumn))
    If netAmt < 0# Then netAmt = -netAmt
    netAmt = Round(netAmt, 2)

    accountStringNormalized = NormalizeCostCodeSeparators(Trim$(CStr(sourceTable(sourceRow, accountStringColumn))))

    '  Rule 3: K = ADD001 + "-" + YYMM(InvoiceDate) + "-" + raw INVOICE NUMBER.
    invoiceDateValue = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    If Len(invoiceDateValue) >= 7 Then
        invoiceDateYYMM = Mid$(invoiceDateValue, 3, 2) & Mid$(invoiceDateValue, 6, 2)
    Else
        invoiceDateYYMM = Format$(sourceTable(sourceRow, invoiceDateColumn), "YYMM")
    End If
    invoiceNumberCompositeValue = SOURCE_SUPPLIER_REFERENCE_ID & "-" & invoiceDateYYMM & "-" & suppliersInvoiceNumberValue

    eventId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mEventLookup, accountStringNormalized)
    campId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCampLookup, accountStringNormalized)
    faLearningId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mFaLearningLookup, accountStringNormalized)
    commercialPartnerId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCommercialPartnerLookup, accountStringNormalized)
    grantId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mGrantLookup, accountStringNormalized)
    programId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProgramLookup, accountStringNormalized)
    projectId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProjectLookup, accountStringNormalized)
    If (LenB(eventId) = 0) Or (LenB(campId) = 0) Or (LenB(faLearningId) = 0) Or (LenB(commercialPartnerId) = 0) Or _
        (LenB(grantId) = 0) Or (LenB(programId) = 0) Or (LenB(projectId) = 0) Then
        optionalWorktagBlankRows = optionalWorktagBlankRows + 1
    End If

    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, suppliersInvoiceNumberValue, mGrpNextKey, lineOrderValue)

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue          ' Rule 15
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberCompositeValue  ' Rule 3
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = resolvedCompanyId            ' Rule 1
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = OUTPUT_CURRENCY              ' Rule 12
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId                   ' Rule 2
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDateValue            ' Rule 4
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = Empty                       ' Rule 5 (blank)
    targetSheet.Cells(targetRow, ColumnIndex("AS")).Value = Empty                       ' Rule 17 (blank)
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = suppliersInvoiceNumberValue ' Rule 6
    targetSheet.Cells(targetRow, ColumnIndex("EL")).Value = Empty
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue              ' Rule 16
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue              ' Rule 19 (mirror of EM)
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = descriptionValue            ' Rule 8
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = spendCategoryId
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = OUTPUT_TAX_APPLICABILITY    ' Rule 11
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = OUTPUT_TAX_CODE
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = OUTPUT_TAX_RECOVERABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = netAmt                      ' Rule 10
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = netAmt                      ' Rule 10 (same as GE)
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = invoiceNumberCompositeValue & " " & SOURCE_COMPANY_CODE & "-" & _
        SOURCE_SUPPLIER_REFERENCE_ID & " " & accountStringNormalized                     ' Rule 9
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(accountStringNormalized)
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = costCenterId
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = eventId
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = campId
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = faLearningId
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = commercialPartnerId
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = grantId
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = programId
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = projectId
End Sub

'  skip_rule.drop_when: Row Labels blank / INVOICE NUMBER blank / INVOICE DATE blank /
'  Net and Gross-VAT both zero.
Private Function RowIsSkippable(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal accountStringColumn As Long, _
    ByVal invoiceNumberColumn As Long, ByVal invoiceDateColumn As Long, ByVal netColumn As Long, ByVal taxColumn As Long, _
    ByVal grossColumn As Long) As Boolean
    RowIsSkippable = (ResolveEffectiveNet(tableData(rowIndex, netColumn), tableData(rowIndex, grossColumn), tableData(rowIndex, taxColumn)) = 0#) _
        Or (LenB(Trim$(CStr(tableData(rowIndex, accountStringColumn)))) = 0) _
        Or (LenB(Trim$(CStr(tableData(rowIndex, invoiceNumberColumn)))) = 0) _
        Or (LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0)
End Function

'  fill_down_context_columns: carry the last non-blank value down to blank cells, stopping at
'  the data end (the pivot Grand Total row is never filled).
Private Sub FillDownInvoiceColumn(ByRef sourceTable As Variant, ByVal columnIndex As Long, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim rowIndex As Long
    Dim carryValue As Variant
    Dim hasCarry As Boolean

    If columnIndex <= 0 Then Exit Sub
    For rowIndex = firstRow To lastRow
        If LenB(Trim$(CStr(sourceTable(rowIndex, columnIndex)))) = 0 Then
            If hasCarry Then sourceTable(rowIndex, columnIndex) = carryValue
        Else
            carryValue = sourceTable(rowIndex, columnIndex)
            hasCarry = True
        End If
    Next rowIndex
End Sub
