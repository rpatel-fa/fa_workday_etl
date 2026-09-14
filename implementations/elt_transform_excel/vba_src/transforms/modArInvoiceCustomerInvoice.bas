Attribute VB_Name = "modArInvoiceCustomerInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
Private Const OUTPUT_CURRENCY As String = "GBP"
' The "k" (Company) row default in config/transform_specs/ar_invoice.yaml's mapping --
' the fallback used when the live "company" mapping lookup has no match. Blank for this
' transform.
Private Const COMPANY_DEFAULT_FALLBACK As String = ""

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Direct transform: load the source invoice extract and the Workday customer extract,
' look up each invoice row's customer via the customer extract's pipe-delimited
' "Reference ID" field, and write straight into the cloned target EIB. No Interim
' workbook and no generic table-ops/keyed-collection engine -- this is a single lookup
' join, not a multi-stage migration, so it does not need one.
Public Sub RunArInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim customerSpreadsheetPath As String
    Dim selectionKey As String
    Dim customerWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim customerLookup As Collection
    Dim sourceTable As Variant
    Dim companyColumn As Long, orderDateColumn As Long, customerIdColumn As Long
    Dim courseNameColumn As Long, courseCodeColumn As Long, grossPriceColumn As Long, commentColumn As Long
    Dim invoiceNumberColumn As Long
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim sourceCustomerId As String
    Dim workdayCustomerId As String
    Dim issues As Collection
    Dim unresolvedCounts As Collection
    Dim unresolvedKeys As Collection
    Dim outputPath As String
    Dim failureMessage As String
    Dim failureNumber As Long
    Dim runStartTime As Date
    Dim runEndTime As Date
    Dim logIssue As String
    Dim worksheetNameStr As String
    Dim worksheetResolutionStr As String
    Dim sheetToResolve As String
    Dim sourceHeaderRow As Long
    Dim sourceDataStartRow As Long

    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)

    Set issues = New Collection
    Set unresolvedCounts = New Collection
    Set unresolvedKeys = New Collection

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH
    customerSpreadsheetPath = ReadControlValue(controlSheet, CONTROL_LABEL_CUSTOMER_SPREADSHEET)
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If StrComp(worksheetResolutionStr, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetNameStr
    End If

    Set customerWorkbook = OpenWorkbookReadOnly(customerSpreadsheetPath)
    Set customerLookup = BuildReferenceLookupFromWorkbook(customerWorkbook, "Customer")
    customerWorkbook.Close SaveChanges:=False
    Set customerWorkbook = Nothing

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    sourceTable = ReadWorksheetTable(ResolveSourceWorksheet(sourceWorkbook, sheetToResolve))
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    companyColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Company")
    orderDateColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Order Date")
    customerIdColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Customer ID")
    courseNameColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Course Name")
    courseCodeColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Course Code")
    grossPriceColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Gross Price")
    commentColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Comment")
    invoiceNumberColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "Invoice Number")
    If invoiceNumberColumn = 0 Then _
        invoiceNumberColumn = FindOptionalColumnIndexInRow(sourceTable, sourceHeaderRow, "Document Number")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "MT"
    targetRow = TARGET_DATA_START_ROW
      Set mGrpKey = New Collection
      Set mGrpLine = New Collection
      mGrpNextKey = 0

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, companyColumn, orderDateColumn, customerIdColumn, grossPriceColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Company / Order Date / Customer ID / Gross Price)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, customerIdColumn)))
        Else
            sourceCustomerId = Trim$(CStr(sourceTable(rowIndex, customerIdColumn)))
            workdayCustomerId = LookupCollectionValue(customerLookup, sourceCustomerId)
            If LenB(workdayCustomerId) = 0 Then
                ' on_fail = "reject" on the "cq" (Customer ID) row of the ar_invoice spec mapping:
                ' an unresolved customer lookup rejects the source row entirely rather than
                ' writing a placeholder value.
                rowsSkipped = rowsSkipped + 1
                IncrementCount unresolvedCounts, unresolvedKeys, sourceCustomerId
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Customer ID not found in customer extract", _
                    sourceCustomerId, _
                    Trim$(CStr(sourceTable(rowIndex, courseNameColumn)))
            Else
                rowsProcessed = rowsProcessed + 1
                WriteOutputRow runnerWorkbook, targetSheet, targetRow, rowsProcessed, sourceTable, rowIndex, _
                    companyColumn, orderDateColumn, courseNameColumn, courseCodeColumn, grossPriceColumn, commentColumn, invoiceNumberColumn, _
                    workdayCustomerId
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Company, Order Date, Customer ID, or Gross Price."
    End If
    If unresolvedKeys.Count > 0 Then
        issues.Add BuildRejectedLookupIssue(unresolvedCounts, unresolvedKeys, "Customer ID")
    End If

    outputPath = outputWorkbook.FullName
    outputWorkbook.Save
    outputWorkbook.Close SaveChanges:=False
    Set outputWorkbook = Nothing

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Completed", rowsProcessed, rowsSkipped, runStartTime, runEndTime, outputPath, JoinCollectionValues(issues, " | "))
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
    If Not customerWorkbook Is Nothing Then customerWorkbook.Close SaveChanges:=False
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
    Err.Raise failureNumber, "modArInvoiceCustomerInvoice.RunArInvoiceCustomerInvoice", failureMessage
End Sub
' counts/keys are deliberately two plain Collections instead of one combined keyed
' structure: `counts` is looked up and updated by key directly (native Collection
' behaviour), and `keys` is a simple append-only list used only to remember which
' distinct customer IDs were seen, so the summary message can enumerate them -- VBA's
' Collection has no API to enumerate its own keys, only to look one up.
Private Sub WriteOutputRow( _
    ByVal runnerWorkbook As Workbook, _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal companyColumn As Long, ByVal orderDateColumn As Long, _
    ByVal courseNameColumn As Long, ByVal courseCodeColumn As Long, ByVal grossPriceColumn As Long, ByVal commentColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal outputCustomerId As String _
)
    Dim companyId As String
    Dim invoiceDate As String
    Dim grossAmount As Double
      Dim invoiceNumberValue As String
      Dim spreadsheetKeyValue As Long
      Dim lineOrderValue As Long

    companyId = ResolveMappedCompany(runnerWorkbook, CStr(sourceTable(sourceRow, companyColumn)), COMPANY_DEFAULT_FALLBACK)
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, orderDateColumn))
    grossAmount = ResolveAmount(sourceTable(sourceRow, grossPriceColumn))
      If invoiceNumberColumn > 0 Then invoiceNumberValue = Trim$(CStr(sourceTable(sourceRow, invoiceNumberColumn)))
      spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceNumberValue, mGrpNextKey, lineOrderValue)

      targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex("E")).Value = OUTPUT_AUTO_COMPLETE
    If LenB(invoiceNumberValue) > 0 Then _
        targetSheet.Cells(targetRow, ColumnIndex("H")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = companyId
    targetSheet.Cells(targetRow, ColumnIndex("CP")).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex("CQ")).Value = outputCustomerId
    ' Force Text format before assigning: the target template's GF cell is formatted
    ' General, so assigning a "yyyy-mm-dd" string straight to .Value lets Excel
    ' auto-detect it as a date and silently re-render it in the system's default
    ' short-date format -- overriding the ar_invoice spec mapping's target_format = YYYY-MM-DD.
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex("GY")).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex("HA")).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex("HC")).Value = OUTPUT_INVOICE_TYPE
     targetSheet.Cells(targetRow, ColumnIndex("IH")).Value = CStr(lineOrderValue)
    ' KE mapping default is "NULL" -> leave explicitly blank (already cleared by ClearPublishedRows).
    targetSheet.Cells(targetRow, ColumnIndex("KE")).Value = ""
    targetSheet.Cells(targetRow, ColumnIndex("KF")).Value = CStr(sourceTable(sourceRow, courseNameColumn)) & "-" & CStr(sourceTable(sourceRow, courseCodeColumn))
    targetSheet.Cells(targetRow, ColumnIndex("MJ")).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex("MT")).Value = CStr(sourceTable(sourceRow, commentColumn))
End Sub
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal companyColumn As Long, _
    ByVal orderDateColumn As Long, _
    ByVal customerIdColumn As Long, _
    ByVal grossPriceColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, companyColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, orderDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, customerIdColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, grossPriceColumn)) = 0#
End Function
