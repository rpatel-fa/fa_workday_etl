Attribute VB_Name = "modApAmexTrainsSupplierInvoice"
Option Explicit
' Source of truth: config/transform_specs/ap_amex_trains.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together
Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const SOURCE_COMPANY_CODE As String = "FA"
Private Const SOURCE_SUPPLIER_REFERENCE_ID As String = "AME001"

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_SPEND_CATEGORY As String = "Staff_Travel_Trains"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"


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

Private mSupplierId As String
Private mCompanyId As String
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

Public Sub RunApAmexTrainsSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim rowIndex As Long
    Dim currencyColumn As Long
    Dim invoiceNumber2Column As Long
    Dim invoiceDateColumn As Long
    Dim supplierInvoiceNumberColumn As Long
    Dim itemDescriptionColumn As Long
    Dim netAmountColumn As Long
    Dim vatPercentageColumn As Long
    Dim finalCodeColumn As Long
    Dim journalNameColumn As Long
    Dim supplierId As String
    Dim companyId As String
    Dim costCenterId As String
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
    Set supplierLookup = BuildReferenceLookupFromWorkbook(supplierWorkbook, "Supplier")
    supplierWorkbook.Close SaveChanges:=False
    Set supplierWorkbook = Nothing

    supplierId = LookupCollectionValue(supplierLookup, SOURCE_SUPPLIER_REFERENCE_ID)
    If LenB(supplierId) = 0 Then
        Err.Raise vbObjectError + 2001, "modApAmexTrainsSupplierInvoice", _
            "Unable to resolve supplier Reference ID '" & SOURCE_SUPPLIER_REFERENCE_ID & "' to a Workday Supplier ID."
    End If
    mSupplierId = supplierId

    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApAmexTrainsSupplierInvoice", 15)

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
    Set mEventLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_EVENT)
    Set mCampLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_CAMP)
    Set mFaLearningLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_FA_LEARNING)
    Set mCommercialPartnerLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_COMMERCIAL_PARTNER)
    Set mGrantLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_GRANT)
    Set mProgramLookup = BuildBusinessObjectHierarchyLookup(mappingTable, WBO_PROGRAM)
    Set mProjectLookup = BuildBusinessObjectHierarchyLookupMulti(mappingTable, WBO_PROJECT, WBO_PROJECT_CAPEX)

    ' Schema validation phase (Tier 2 VBA defence-in-depth).
    Dim schemaMsgT1 As String
    Dim schemaHeaderRowT1 As Long
    Dim schemaDiagT1 As String
    schemaMsgT1 = ValidateSourceWorkbookSchema( _
        sheetToResolve, _
        sourceTable, _
        sourceHeaderRow, _
        sourceDataStartRow, _
        requiredHeaders, _
        , _
        schemaHeaderRowT1, _
        schemaDiagT1 _
    )
    If LenB(schemaMsgT1) > 0 Then
        Err.Raise vbObjectError + 1910, "modApAmexTrainsSupplierInvoice.RunApAmexTrainsSupplierInvoice", schemaMsgT1
    End If
    If LenB(schemaDiagT1) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiagT1, vbCrLf, " | ")
    End If

    ' Detect the header row by content rather than assuming a fixed physical row, then
    ' derive the data-start row from the detected header position so the offset holds
    ' regardless of any leading title/blank rows in a given month's export.
    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    currencyColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Currency")
    invoiceNumber2Column = FindColumnIndexInRow(sourceTable, headerRowIndex, "INVOICE NUMBER2")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "INVOICE DATE")
    supplierInvoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "SUPPLIER INVOICE NUMBER")
    itemDescriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "ITEM DESCRIPTION")
    netAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net amount")
    vatPercentageColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "VAT Percentage")
    finalCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Final Code")
    journalNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Journal Name?")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    ' Company ID is invariant across output rows (fixed SOURCE_COMPANY_CODE), so resolve
    ' it once instead of re-scanning the mapping sheet inside WriteOutputRow per row.
    companyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, currencyColumn, invoiceNumber2Column, invoiceDateColumn, supplierInvoiceNumberColumn, itemDescriptionColumn, netAmountColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Currency / INVOICE NUMBER2 / INVOICE DATE / SUPPLIER INVOICE NUMBER / ITEM DESCRIPTION / Net amount)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, journalNameColumn)))
        Else
            costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, sourceTable(rowIndex, finalCodeColumn))
            If LenB(costCenterId) = 0 Then
                rowsSkipped = rowsSkipped + 1
                unmappedCostCenterRows = unmappedCostCenterRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Cost centre unresolved: Final Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                    Trim$(CStr(sourceTable(rowIndex, finalCodeColumn))), _
                    Trim$(CStr(sourceTable(rowIndex, journalNameColumn)))
            Else
                WriteOutputRow _
                    targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                    currencyColumn, invoiceNumber2Column, invoiceDateColumn, supplierInvoiceNumberColumn, itemDescriptionColumn, netAmountColumn, vatPercentageColumn, finalCodeColumn, _
                    journalNameColumn, costCenterId
                rowsProcessed = rowsProcessed + 1
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Currency, INVOICE NUMBER2, INVOICE DATE, SUPPLIER INVOICE NUMBER, ITEM DESCRIPTION, or Net amount."
    End If
    If unmappedCostCenterRows > 0 Then
        issueParts.Add "Skipped " & CStr(unmappedCostCenterRows) & " source row(s) whose Final Code did not resolve to a Workday cost centre in '" & mappingSheetName & "' using the SQL-style lowest hierarchy match."
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
    Err.Raise failureNumber, "modApAmexTrainsSupplierInvoice.RunApAmexTrainsSupplierInvoice", failureMessage
End Sub

Private Sub WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal currencyColumn As Long, _
    ByVal invoiceNumber2Column As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal supplierInvoiceNumberColumn As Long, _
    ByVal itemDescriptionColumn As Long, _
    ByVal netAmountColumn As Long, _
    ByVal vatPercentageColumn As Long, _
    ByVal finalCodeColumn As Long, _
    ByVal journalNameColumn As Long, _
    ByVal costCenterId As String _
)
    Dim currencyCode As String
    Dim invoiceDate As String
    Dim amount As Double
    Dim vatRate As Long
    Dim finalCodeValue As Variant

    currencyCode = Trim$(CStr(sourceTable(sourceRow, currencyColumn)))
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    amount = ResolveAmount(sourceTable(sourceRow, netAmountColumn))
    vatRate = ResolveRoundedRate(sourceTable(sourceRow, vatPercentageColumn))
    finalCodeValue = sourceTable(sourceRow, finalCodeColumn)

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = sequenceNumber
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = Trim$(CStr(sourceTable(sourceRow, invoiceNumber2Column)))
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = mCompanyId
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = currencyCode
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = mSupplierId
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = Trim$(CStr(sourceTable(sourceRow, supplierInvoiceNumberColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("DD")).Value = 1
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = Trim$(CStr(sourceTable(sourceRow, itemDescriptionColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = OUTPUT_SPEND_CATEGORY
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = TaxValueForRate(vatRate, TAX_APPLICABILITY_ZERO, TAX_APPLICABILITY_STANDARD)
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = TaxValueForRate(vatRate, TAX_CODE_ZERO, TAX_CODE_STANDARD)
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = TaxValueForRate(vatRate, TAX_RECOVERABILITY_ZERO, TAX_RECOVERABILITY_STANDARD)
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("IK")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("IL")).Value = Trim$(CStr(sourceTable(sourceRow, journalNameColumn)))
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = ResolveReportingTypeFromFinalCode(finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = costCenterId
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mEventLookup, finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCampLookup, finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mFaLearningLookup, finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mCommercialPartnerLookup, finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mGrantLookup, finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProgramLookup, finalCodeValue)
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(mProjectLookup, finalCodeValue)
End Sub

Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal currencyColumn As Long, _
    ByVal invoiceNumber2Column As Long, _
    ByVal invoiceDateColumn As Long, _
    ByVal supplierInvoiceNumberColumn As Long, _
    ByVal itemDescriptionColumn As Long, _
    ByVal netAmountColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, currencyColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceNumber2Column)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, supplierInvoiceNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, itemDescriptionColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, netAmountColumn)) = 0#
End Function
