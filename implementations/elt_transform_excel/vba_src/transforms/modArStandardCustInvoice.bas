Attribute VB_Name = "modArStandardCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const TARGET_LAST_COLUMN_LETTERS As String = "IC"

Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const OUTPUT_ADD_ONLY As String = "Y"
Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"
Private Const REFERENCE_ID_PREFIX As String = "ar_"

Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Sale_of_Goods_and_Services"
Private Const TAX_APPLICABILITY_ZERO As String = "GBR_Outside_of_Scope"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_CODE_ZERO As String = "GBR_Zero_Rated"

' Target letters mirror config/transforms/ar_standard_invoice.yaml.
Private Const TGT_SPREADSHEET_KEY As String = "B"
Private Const TGT_ADD_ONLY As String = "C"
Private Const TGT_CUSTOMER_INVOICE_REF_ID As String = "D"
Private Const TGT_HEADER_COMMENT As String = "E"
Private Const TGT_CUSTOMER_INVOICE_ID As String = "G"
Private Const TGT_SUBMIT As String = "H"
Private Const TGT_COMPANY As String = "J"
Private Const TGT_CURRENCY As String = "AI"
Private Const TGT_CUSTOMER As String = "AJ"
Private Const TGT_INVOICE_NUMBER As String = "BO"
Private Const TGT_INVOICE_DATE As String = "BR"
Private Const TGT_PAYMENT_STATUS As String = "CK"
Private Const TGT_DOCUMENT_STATUS As String = "CL"
Private Const TGT_AMOUNT_DUE As String = "CM"
Private Const TGT_CUSTOMER_INVOICE_TYPE As String = "CO"
Private Const TGT_HEADER_TAX_CODE As String = "CP"
Private Const TGT_CUSTOMER_PO_NUMBER As String = "CQ"
Private Const TGT_INCLUDE_CUSTOMER_WORKTAGS As String = "CY"
Private Const TGT_HEADER_MEMO As String = "DB"
Private Const TGT_ROW_ID_LINE As String = "DS"
Private Const TGT_LINE_REFERENCE As String = "DU"
Private Const TGT_LINE_ORDER As String = "DV"
Private Const TGT_SALES_ITEM As String = "EJ"
Private Const TGT_REVENUE_CATEGORY As String = "EM"
Private Const TGT_LINE_ITEM_DESCRIPTION As String = "EN"
Private Const TGT_LINE_TAX_APPLICABILITY As String = "FA"
Private Const TGT_LINE_TAX_CODE As String = "FB"
Private Const TGT_LINE_QUANTITY As String = "FH"
Private Const TGT_LINE_UNIT_OF_MEASURE As String = "FI"
Private Const TGT_LINE_UNIT_COST As String = "FL"
Private Const TGT_LINE_EXTENDED_AMOUNT As String = "FO"
Private Const TGT_LINE_TRANSACTION_DATE As String = "FP"
Private Const TGT_LINE_FROM_DATE As String = "FQ"
Private Const TGT_LINE_TO_DATE As String = "FR"
Private Const TGT_DEFERRED_REVENUE As String = "FS"
Private Const TGT_LINE_MEMO As String = "FV"
Private Const TGT_COST_CENTRE As String = "FX"
Private Const TGT_REPORTING_TYPE As String = "FY"
Private Const TGT_CAMP As String = "FZ"
Private Const TGT_EVENT As String = "GA"
Private Const TGT_COMMERCIAL_PARTNER As String = "GB"
Private Const TGT_FA_LEARNING_COURSE As String = "GC"
Private Const TGT_GRANT As String = "GD"
Private Const TGT_PROGRAMME As String = "GE"
Private Const TGT_PROJECT_LINE_WORKTAG As String = "GF"

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

Private Enum SourceColumnSlot
    SRC_WORKDAY_ID = 0
    SRC_DATE
    SRC_COMPANY
    SRC_DESCRIPTION
    SRC_QUANTITY
    SRC_NET
    SRC_GROSS
    SRC_REPORTING_TYPE
    SRC_COST_CENTRE
    SRC_REVENUE_CATEGORY
    SRC_CUSTOMER_PO
    SRC_HEADER_MEMO
    SRC_LINE_MEMO
    SRC_EVENT
    SRC_FA_LEARNING
End Enum

Public Sub RunArStandardInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceTable As Variant
    Dim sourceColumns As Variant
    Dim firstGroupValues As Collection
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim issues As Collection
    Dim outputPath As String
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
    Dim worksheetName As String
    Dim worksheetResolution As String
    Dim sheetToResolve As String
    Dim sourceHeaderRow As Long
    Dim sourceDataStartRow As Long
    Dim dateColumn As Long
    Dim companyColumn As Long
    Dim workdayIdColumn As Long
    Dim descriptionColumn As Long
    Dim qtyColumn As Long
    Dim netColumn As Long
    Dim grossColumn As Long
    Dim reportingTypeColumn As Long
    Dim costCentreColumn As Long
    Dim revenueCategoryColumn As Long
    Dim customerPoColumn As Long
    Dim headerMemoColumn As Long
    Dim lineMemoColumn As Long
    Dim eventColumn As Long
    Dim faLearningColumn As Long

    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)
    Set issues = New Collection

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH
    worksheetName = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolution = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If StrComp(worksheetResolution, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetName
    End If

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    sourceTable = ReadWorksheetTable(ResolveSourceWorksheet(sourceWorkbook, sheetToResolve))
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    dateColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "DATE")
    companyColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "COMPANY")
    workdayIdColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "WORKDAY ID")
    descriptionColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "INVOICE LINE DESCRIPTION")
    qtyColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "QTY")
    netColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "UNIT COST")
    grossColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Total")
    reportingTypeColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Reporting Type")
    costCentreColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Cost Centre")
    revenueCategoryColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Revenue Category")
    customerPoColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "Purchase Order")
    headerMemoColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "INVOICE MEMO (DB)")
    lineMemoColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "INVOICE LINE MEMO (emails) (FV)")
    eventColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "Event-OPTIONAL")
    faLearningColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "FA LEARNING COURSES-OPTIONAL")
    sourceColumns = Array( _
        workdayIdColumn, dateColumn, companyColumn, descriptionColumn, qtyColumn, _
        netColumn, grossColumn, reportingTypeColumn, costCentreColumn, revenueCategoryColumn, _
        customerPoColumn, headerMemoColumn, lineMemoColumn, eventColumn, faLearningColumn _
    )

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, TARGET_LAST_COLUMN_LETTERS
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    Set firstGroupValues = New Collection
    mGrpNextKey = 0

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then
            UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        End If

        If RowIsSkippable(sourceTable, rowIndex, workdayIdColumn, dateColumn, netColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (WORKDAY ID / DATE / UNIT COST)", _
                ReadSourceText(sourceTable, rowIndex, workdayIdColumn), _
                ReadSourceText(sourceTable, rowIndex, descriptionColumn)
        Else
            WriteOutputRow targetSheet, targetRow, sourceTable, rowIndex, sourceColumns, firstGroupValues
            rowsProcessed = rowsProcessed + 1
            targetRow = targetRow + 1
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & _
            " source row(s) with missing WORKDAY ID, DATE, or UNIT COST."
    End If

    outputPath = outputWorkbook.FullName
    outputWorkbook.Save
    outputWorkbook.Close SaveChanges:=False
    Set outputWorkbook = Nothing

    runEndTime = Now
    logIssue = AppendRunLog( _
        runnerWorkbook, selectionKey, "Completed", rowsProcessed, rowsSkipped, _
        runStartTime, runEndTime, outputPath, JoinCollectionValues(issues, " | ") _
    )
    If LenB(logIssue) > 0 Then issues.Add logIssue
    EndRun

    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Completed"
    WriteControlValue controlSheet, RESULT_LABEL_ROWS_PROCESSED, CStr(rowsProcessed)
    WriteControlValue controlSheet, RESULT_LABEL_ROWS_SKIPPED, CStr(rowsSkipped)
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, JoinCollectionValues(issues, " | ")
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
    logIssue = AppendRunLog( _
        runnerWorkbook, selectionKey, "Failed", rowsProcessed, rowsSkipped, _
        runStartTime, runEndTime, vbNullString, failureMessage _
    )
    If LenB(logIssue) > 0 Then failureMessage = failureMessage & " | " & logIssue
    EndRun

    WriteControlValue controlSheet, RESULT_LABEL_RUN_STATUS, "Failed"
    WriteControlValue controlSheet, RESULT_LABEL_ISSUES, failureMessage
    WriteControlValue controlSheet, RESULT_LABEL_SELECTION_KEY, selectionKey
    WriteControlValue controlSheet, RESULT_LABEL_LAST_RUN_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error Resume Next
    runnerWorkbook.Save
    On Error GoTo 0
    Err.Raise failureNumber, "modArStandardCustInvoice.RunArStandardInvoiceCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputRow(ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal sourceColumns As Variant, ByVal firstGroupValues As Collection)
    Dim workdayId As String
    Dim invoiceDate As String
    Dim groupingKey As String
    Dim invoiceId As String
    Dim netAmount As Double
    Dim grossAmount As Double
    Dim calculatedVat As Double
    Dim taxApplicability As String
    Dim taxCode As String
    Dim spreadsheetKey As Long
    Dim lineOrder As Long
    Dim qtyValue As Double
    Dim groupValues As Variant

    workdayId = ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_WORKDAY_ID))
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, sourceColumns(SRC_DATE)))
    netAmount = ResolveAmount(sourceTable(sourceRow, sourceColumns(SRC_NET)))
    grossAmount = ResolveAmount(sourceTable(sourceRow, sourceColumns(SRC_GROSS)))
    calculatedVat = grossAmount - netAmount
    qtyValue = ResolveAmount(sourceTable(sourceRow, sourceColumns(SRC_QUANTITY)))
    groupingKey = workdayId & Chr$(30) & invoiceDate

    spreadsheetKey = AssignInvoiceLineGroup( _
        mGrpKey, mGrpLine, groupingKey, mGrpNextKey, lineOrder _
    )
    If Not CollectionKeyExists(firstGroupValues, groupingKey) Then
        firstGroupValues.Add Array( _
            ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_CUSTOMER_PO)), _
            ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_HEADER_MEMO)) _
        ), groupingKey
    End If
    groupValues = firstGroupValues(groupingKey)

    If calculatedVat = 0# Then
        taxApplicability = TAX_APPLICABILITY_ZERO
        taxCode = TAX_CODE_ZERO
    Else
        taxApplicability = TAX_APPLICABILITY_STANDARD
        taxCode = TAX_CODE_STANDARD
    End If

    invoiceId = REFERENCE_ID_PREFIX & Replace(workdayId, "-", vbNullString) & _
        "_" & Replace(invoiceDate, "-", vbNullString) & "_" & CStr(spreadsheetKey)

    targetSheet.Cells(targetRow, ColumnIndex(TGT_SPREADSHEET_KEY)).Value = spreadsheetKey
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ADD_ONLY)).Value = OUTPUT_ADD_ONLY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_COMMENT)).Value = OUTPUT_AUTO_COMPLETE
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_REF_ID)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_ID)).Value = invoiceId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_SUBMIT)).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMPANY)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_COMPANY))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CURRENCY)).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER)).Value = workdayId
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_NUMBER)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INVOICE_DATE)).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex(TGT_PAYMENT_STATUS)).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DOCUMENT_STATUS)).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex(TGT_AMOUNT_DUE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_INVOICE_TYPE)).Value = OUTPUT_INVOICE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_TAX_CODE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CUSTOMER_PO_NUMBER)).Value = groupValues(0)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_INCLUDE_CUSTOMER_WORKTAGS)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_HEADER_MEMO)).Value = groupValues(1)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_ROW_ID_LINE)).Value = CStr(lineOrder)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_REFERENCE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ORDER)).Value = CStr(lineOrder)
    targetSheet.Cells(targetRow, ColumnIndex(TGT_REVENUE_CATEGORY)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_REVENUE_CATEGORY))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_ITEM_DESCRIPTION)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_DESCRIPTION))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_APPLICABILITY)).Value = taxApplicability
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TAX_CODE)).Value = taxCode
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_QUANTITY)).Value = qtyValue
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_OF_MEASURE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_UNIT_COST)).Value = netAmount
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_EXTENDED_AMOUNT)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TRANSACTION_DATE)).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_FROM_DATE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_TO_DATE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_DEFERRED_REVENUE)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_LINE_MEMO)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_LINE_MEMO))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COST_CENTRE)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_COST_CENTRE))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_REPORTING_TYPE)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_REPORTING_TYPE))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_CAMP)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_EVENT)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_EVENT))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_COMMERCIAL_PARTNER)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_FA_LEARNING_COURSE)).Value = _
        ReadSourceText(sourceTable, sourceRow, sourceColumns(SRC_FA_LEARNING))
    targetSheet.Cells(targetRow, ColumnIndex(TGT_GRANT)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_PROGRAMME)).Value = vbNullString
    targetSheet.Cells(targetRow, ColumnIndex(TGT_PROJECT_LINE_WORKTAG)).Value = vbNullString
End Sub

Private Function RowIsSkippable( _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal workdayIdColumn As Long, _
    ByVal dateColumn As Long, _
    ByVal netColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(ReadSourceText(sourceTable, sourceRow, workdayIdColumn)) = 0 Or _
        LenB(ReadSourceText(sourceTable, sourceRow, dateColumn)) = 0 Or _
        LenB(ReadSourceText(sourceTable, sourceRow, netColumn)) = 0
End Function

Private Function ReadSourceText( _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceColumn As Long _
) As String
    If sourceColumn <= 0 Then Exit Function
    If IsError(sourceTable(sourceRow, sourceColumn)) Or _
       IsEmpty(sourceTable(sourceRow, sourceColumn)) Or _
       IsNull(sourceTable(sourceRow, sourceColumn)) Then Exit Function
    ReadSourceText = Trim$(CStr(sourceTable(sourceRow, sourceColumn)))
End Function
