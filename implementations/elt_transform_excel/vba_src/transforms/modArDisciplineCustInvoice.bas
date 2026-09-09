Attribute VB_Name = "modArDisciplineCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const SOURCE_COMPANY_CODE As String = "FA"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "DSC"
Private Const OUTPUT_DSC_INDICATOR As String = "Y"
' The "k" (Company) row default in config/transform_specs/ar_discipline_invoice.yaml's
' mapping -- the fallback used when the live "company" mapping lookup has no match.
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"

' Direct transform: load the Discipline (fines/payments) invoice extract and the
' Workday customer extract, look up each row's customer via the customer extract's
' pipe-delimited "Reference ID" field, and write straight into the cloned target EIB.
' Company and Currency are fixed values (the source extract has neither column). The
' field mapping mirrors the mapping: section in
' config/transform_specs/ar_discipline_invoice.yaml, the same pattern AR Invoice uses.
Public Sub RunArDisciplineInvoiceCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim customerNumberColumn As Long, documentNumberColumn As Long, documentDateColumn As Long, amountColumn As Long
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

    customerNumberColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Customer Number")
    documentNumberColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Document Number")
    documentDateColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Document Date")
    amountColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Original Trx Amount")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "MT"
    targetRow = TARGET_DATA_START_ROW

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, customerNumberColumn, documentDateColumn, amountColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Customer Number / Document Date / Amount)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, customerNumberColumn)))
        Else
            sourceCustomerId = Trim$(CStr(sourceTable(rowIndex, customerNumberColumn)))
            workdayCustomerId = LookupCollectionValue(customerLookup, sourceCustomerId)
            If LenB(workdayCustomerId) = 0 Then
                ' on_fail = "reject" on the "cq" (Customer ID) row of the ar_discipline_invoice
                ' spec mapping: an unresolved customer lookup rejects the source row entirely
                ' rather than writing a placeholder value.
                rowsSkipped = rowsSkipped + 1
                IncrementCount unresolvedCounts, unresolvedKeys, sourceCustomerId
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Customer Number not found in customer extract", _
                    sourceCustomerId, _
                    Trim$(CStr(sourceTable(rowIndex, documentNumberColumn)))
            Else
                rowsProcessed = rowsProcessed + 1
                WriteOutputRow runnerWorkbook, targetSheet, targetRow, rowsProcessed, sourceTable, rowIndex, _
                    documentNumberColumn, documentDateColumn, amountColumn, workdayCustomerId
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Customer Number, Document Date, or Original Trx Amount."
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
    Err.Raise failureNumber, "modArDisciplineCustInvoice.RunArDisciplineInvoiceCustomerInvoice", failureMessage
End Sub
Private Sub WriteOutputRow( _
    ByVal runnerWorkbook As Workbook, _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal documentNumberColumn As Long, ByVal documentDateColumn As Long, ByVal amountColumn As Long, _
    ByVal outputCustomerId As String _
)
    Dim companyId As String
    Dim invoiceDate As String
    Dim amount As Double

    companyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, COMPANY_DEFAULT_FALLBACK)
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, documentDateColumn))
    amount = ResolveAmount(sourceTable(sourceRow, amountColumn))

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = sequenceNumber
    targetSheet.Cells(targetRow, ColumnIndex("E")).Value = OUTPUT_AUTO_COMPLETE
    targetSheet.Cells(targetRow, ColumnIndex("H")).Value = CStr(sourceTable(sourceRow, documentNumberColumn))
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = companyId
    targetSheet.Cells(targetRow, ColumnIndex("CP")).Value = OUTPUT_CURRENCY
    targetSheet.Cells(targetRow, ColumnIndex("CQ")).Value = outputCustomerId
    ' Force Text format before assigning: the target template's GF cell is formatted
    ' General, so assigning a "yyyy-mm-dd" string straight to .Value lets Excel
    ' auto-detect it as a date and silently re-render it in the system's default
    ' short-date format -- overriding the ar_discipline_invoice spec mapping's
    ' target_format = YYYY-MM-DD.
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex("GY")).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = OUTPUT_DOCUMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex("HA")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("HC")).Value = OUTPUT_INVOICE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("HL")).Value = OUTPUT_DSC_INDICATOR
    targetSheet.Cells(targetRow, ColumnIndex("IH")).Value = "1"
    targetSheet.Cells(targetRow, ColumnIndex("MJ")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("MT")).Value = CStr(sourceTable(sourceRow, documentNumberColumn))
End Sub
Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal customerNumberColumn As Long, _
    ByVal documentDateColumn As Long, _
    ByVal amountColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, customerNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, documentDateColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, amountColumn)) = 0#
End Function
