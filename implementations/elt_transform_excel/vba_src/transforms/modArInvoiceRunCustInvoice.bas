Attribute VB_Name = "modArInvoiceRunCustInvoice"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "Submit Customer Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const SOURCE_COMPANY_CODE As String = "WNSL"
Private Const OUTPUT_CURRENCY As String = "GBP"
Private Const OUTPUT_AUTO_COMPLETE As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_PAYMENT_STATUS As String = "Unpaid"
Private Const OUTPUT_DOCUMENT_STATUS As String = "Approved"
Private Const OUTPUT_INVOICE_TYPE As String = "Standard"
' Output VAT itemisation (decided 2026-07-30, mirrors the tax: section of
' config/transform_specs/ar_invoice_run.yaml). Tax Code Ref IDs are the tenant CCS "Tenant Default
' Named Ranges" values. Fines are zero-rated; every other invoice type is standard 20%. The line
' Tax Code (LW) drives Workday's VAT calc; NET posts to Extended Amount (MJ); gross TOTAL stays on
' Amount Due (HA) as the control figure.
Private Const TAX_DRIVER_FINES As String = "FINES"
Private Const TAX_CODE_ZERO_RATED As String = "GBR_Zero_Rated"
Private Const TAX_CODE_STANDARD As String = "GBR_200_VAT"
' The "k" (Company) row default in config/transform_specs/ar_invoice_run.yaml's mapping -- the
' fallback used when the live "company" mapping lookup has no match. The whole WNL run is fixed
' to WNSL (source column E is not the company driver).
Private Const COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - WNSL"

Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

' Direct transform: load the AR "invoice run" extract and the Workday customer extract, look up
' each row's customer via the customer extract's pipe-delimited "Reference ID" field, and write
' straight into the cloned target EIB. Company and Currency are fixed values (the source has
' neither as a usable column). No line cost centre is written -- the Customer Invoice EIB does not
' require one (see ar_invoice), so the source "Cost Code" is intentionally not carried. Output VAT
' is itemised: NET posts to Extended Amount (MJ), a per-row line Tax Code (LW) derived from the
' source "FA WNL Code" drives Workday's VAT calc, and gross TOTAL posts to Amount Due (HA). The
' field mapping mirrors the mapping: section in config/transform_specs/ar_invoice_run.yaml.
Public Sub RunArInvoiceRunCustomerInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim codeColumn As Long, dateColumn As Long, invoiceTypeColumn As Long
    Dim clubNameColumn As Long, totalColumn As Long
    Dim netColumn As Long, taxDriverColumn As Long, invoiceNumberColumn As Long
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim companyId As String
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

    codeColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Code")
    dateColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Date")
    invoiceTypeColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Invoice type")
    clubNameColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "Club Name")
    netColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "NET")
    taxDriverColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "FA WNL Code")
    totalColumn = FindColumnIndexInRow(sourceTable, sourceHeaderRow, "TOTAL")
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

    ' Company ID is invariant across output rows (fixed SOURCE_COMPANY_CODE), so resolve it once
    ' instead of re-scanning the mapping sheet inside WriteOutputRow for every row.
    companyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, COMPANY_DEFAULT_FALLBACK)

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, codeColumn, dateColumn, totalColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Code / Date / TOTAL)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, codeColumn)))
        Else
            sourceCustomerId = Trim$(CStr(sourceTable(rowIndex, codeColumn)))
            workdayCustomerId = LookupCollectionValue(customerLookup, sourceCustomerId)
            If LenB(workdayCustomerId) = 0 Then
                ' on_fail = "reject" on the "cq" (Customer) row of the ar_invoice_run spec mapping:
                ' an unresolved customer lookup rejects the source row entirely rather than writing
                ' a placeholder value.
                rowsSkipped = rowsSkipped + 1
                IncrementCount unresolvedCounts, unresolvedKeys, sourceCustomerId
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Code not found in customer extract", _
                    sourceCustomerId, _
                    Trim$(CStr(sourceTable(rowIndex, clubNameColumn)))
            Else
                rowsProcessed = rowsProcessed + 1
                WriteOutputRow targetSheet, targetRow, rowsProcessed, sourceTable, rowIndex, _
                    dateColumn, invoiceTypeColumn, clubNameColumn, _
                    netColumn, taxDriverColumn, totalColumn, invoiceNumberColumn, _
                    companyId, workdayCustomerId
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issues.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Code, Date, or TOTAL."
    End If
    If unresolvedKeys.Count > 0 Then
        issues.Add BuildRejectedLookupIssue(unresolvedCounts, unresolvedKeys, "Code")
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
    Err.Raise failureNumber, "modArInvoiceRunCustInvoice.RunArInvoiceRunCustomerInvoice", failureMessage
End Sub

Private Sub WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal dateColumn As Long, ByVal invoiceTypeColumn As Long, ByVal clubNameColumn As Long, _
    ByVal netColumn As Long, ByVal taxDriverColumn As Long, ByVal totalColumn As Long, _
    ByVal invoiceNumberColumn As Long, _
    ByVal companyId As String, _
    ByVal outputCustomerId As String _
)
    Dim invoiceDate As String
    Dim grossAmount As Double
    Dim netAmount As Double
    Dim taxCode As String
    Dim invoiceNumberValue As String
      Dim spreadsheetKeyValue As Long
      Dim lineOrderValue As Long

    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, dateColumn))
    grossAmount = ResolveAmount(sourceTable(sourceRow, totalColumn))
    netAmount = ResolveAmount(sourceTable(sourceRow, netColumn))
    taxCode = DeriveTaxCode(sourceTable(sourceRow, taxDriverColumn))
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
    ' Force Text format before assigning: the target template's GF cell is formatted General, so
    ' assigning a "yyyy-mm-dd" string straight to .Value lets Excel auto-detect it as a date and
    ' silently re-render it in the system's default short-date format -- overriding the
    ' ar_invoice_run spec mapping's target_format = YYYY-MM-DD.
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = invoiceDate
    targetSheet.Cells(targetRow, ColumnIndex("GY")).Value = OUTPUT_PAYMENT_STATUS
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = OUTPUT_DOCUMENT_STATUS
    ' Amount Due (HA) stays the gross TOTAL as the control figure; Workday reconciles it against
    ' net (MJ) + calculated VAT (from the line Tax Code LW).
    targetSheet.Cells(targetRow, ColumnIndex("HA")).Value = grossAmount
    targetSheet.Cells(targetRow, ColumnIndex("HC")).Value = OUTPUT_INVOICE_TYPE
     targetSheet.Cells(targetRow, ColumnIndex("IH")).Value = CStr(lineOrderValue)
    ' KE mapping default is "NULL" -> leave explicitly blank (already cleared by ClearPublishedRows).
    targetSheet.Cells(targetRow, ColumnIndex("KE")).Value = ""
    targetSheet.Cells(targetRow, ColumnIndex("KF")).Value = CStr(sourceTable(sourceRow, invoiceTypeColumn))
    ' LW = line Tax Code; MJ = NET (Extended Amount). Workday computes the VAT from the code.
    targetSheet.Cells(targetRow, ColumnIndex("LW")).Value = taxCode
    targetSheet.Cells(targetRow, ColumnIndex("MJ")).Value = netAmount
    targetSheet.Cells(targetRow, ColumnIndex("MT")).Value = CStr(sourceTable(sourceRow, clubNameColumn))
End Sub

' Map the source "FA WNL Code" to a Workday output-VAT Tax Code Ref ID. Fines are zero-rated; every
' other invoice type is standard 20%. Mirrors the tax: sets in the ar_invoice_run transform spec.
Private Function DeriveTaxCode(ByVal taxDriverValue As Variant) As String
    If StrComp(Trim$(CStr(taxDriverValue)), TAX_DRIVER_FINES, vbTextCompare) = 0 Then
        DeriveTaxCode = TAX_CODE_ZERO_RATED
    Else
        DeriveTaxCode = TAX_CODE_STANDARD
    End If
End Function

Private Function RowIsSkippable( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal codeColumn As Long, _
    ByVal dateColumn As Long, _
    ByVal totalColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, codeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, dateColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, totalColumn)) = 0#
End Function
