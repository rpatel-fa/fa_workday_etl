Attribute VB_Name = "modApAmexHotelV3SupplierInvoice"
Option Explicit

' Source of truth: config/transform_specs/ap_amex_hotel_v3.yaml
Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const REQUIRED_HEADER_COUNT As Long = 14

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_DEFAULT_CURRENCY As String = "GBP"
Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const TEXT_NUMBER_FORMAT As String = "@"
Private Const AMOUNT_NUMBER_FORMAT As String = "0.00"
Private Const MEMO_SEPARATOR As String = "_"

Private Const TAX_APPLICABILITY_STANDARD As String = "GBR_Domestic_Purchase_Of_Goods_And_Services"
Private Const TAX_APPLICABILITY_ZERO As String = "GBR_Outside_of_Scope"
Private Const TAX_CODE_STANDARD As String = "GBR_Standard"
Private Const TAX_CODE_ZERO As String = "GBR_Zero_Rated"
Private Const TAX_RECOVERABILITY_STANDARD As String = "Fully_Recoverable"
Private Const TAX_RECOVERABILITY_ZERO As String = "Nonrecoverable_Allocate"

Private Const SOURCE_UNRESOLVED_MARKER As String = "NOT_FOUND"
Private Const SOURCE_ERROR_TEXT As String = "#ERROR"
Private Const SUPPLIER_EXTRACT_ENTITY As String = "Supplier"
Private Const SUPPLIER_EXTRACT_REFERENCE_HEADER As String = "Reference ID"
Private Const SUPPLIER_EXTRACT_ID_HEADER As String = "Supplier ID"

Private Const REJECT_MISSING_FIELD As Long = 1
Private Const REJECT_SUPPLIER As Long = 2
Private Const REJECT_WORKTAG As Long = 3
Private Const REJECT_TAX_CODE As Long = 4
Private Const REJECT_NEGATIVE_NET As Long = 5
Private Const WORKTAG_COUNT As Long = 10

Private mSupplierIds As Collection
Private mGrpKey As Collection
Private mGrpLine As Collection
Private mGrpNextKey As Long

Private mColSupplierId As Long
Private mColInvoiceNumber As Long
Private mColInvoiceDate As Long
Private mColCurrency As Long
Private mColCompany As Long
Private mColNet As Long
Private mColGross As Long
Private mColDescription As Long
Private mColTaxCode As Long
Private mColHotelStayMemo As Long
Private mWorktagHeaders(1 To 10) As String
Private mWorktagMandatory(1 To 10) As Boolean
Private mWorktagCols(1 To 10) As Long

Public Sub RunApAmexHotelV3SupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim supplierSpreadsheetPath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim supplierWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As Variant
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim missingFieldRows As Long
    Dim unresolvedSupplierRows As Long
    Dim worktagRejectRows As Long
    Dim unknownTaxCodeRows As Long
    Dim negativeNetRows As Long
    Dim rowIndex As Long
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
    Dim rejectReason As String
    Dim rejectCategory As Long
    Dim taxApply As String
    Dim taxCod As String
    Dim taxRecov As String
    Dim groupCompanyCaptured As Collection
    Dim groupSupplierCaptured As Collection
    Dim groupCurrencyCaptured As Collection
    Dim groupInvDateCaptured As Collection
    Dim prepassMsg As String
    Dim invoiceNumberValue As String
    Dim companyValue As String
    Dim supplierValue As String
    Dim resolvedCurrency As String
    Dim invoiceDateValue As String

    consecutiveEmptyRows = 0
    selectionKey = ResolveSelectedTransformKey(runnerWorkbook, controlSheet)
    Set issueParts = New Collection
    On Error GoTo HandleFailure
    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
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
    EnsureNonBlank supplierSpreadsheetPath, CONTROL_LABEL_SUPPLIER_SPREADSHEET

    Set supplierWorkbook = OpenWorkbookReadOnly(supplierSpreadsheetPath)
    Set mSupplierIds = BuildSupplierIdSet(supplierWorkbook)
    supplierWorkbook.Close SaveChanges:=False
    Set supplierWorkbook = Nothing

    requiredHeaders = ResolveRequiredHeadersFromConfig( _
        controlSheet, "modApAmexHotelV3SupplierInvoice", REQUIRED_HEADER_COUNT)

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    Set sourceSheet = ResolveSourceWorksheet(sourceWorkbook, sheetToResolve)
    sourceTable = ReadWorksheetTable(sourceSheet)
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    Dim schemaMsg As String
    Dim schemaHeaderRow As Long
    Dim schemaDiag As String
    schemaMsg = ValidateSourceWorkbookSchema(sheetToResolve, sourceTable, sourceHeaderRow, sourceDataStartRow, _
        requiredHeaders, , schemaHeaderRow, schemaDiag)
    If LenB(schemaMsg) > 0 Then
        Err.Raise vbObjectError + 1910, _
            "modApAmexHotelV3SupplierInvoice.RunApAmexHotelV3SupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)
    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders
    ResolveSourceColumns sourceTable, headerRowIndex

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "IL"
    targetRow = TARGET_DATA_START_ROW
    Set mGrpKey = New Collection
    Set mGrpLine = New Collection
    mGrpNextKey = 0

    Set groupCompanyCaptured = New Collection
    Set groupSupplierCaptured = New Collection
    Set groupCurrencyCaptured = New Collection
    Set groupInvDateCaptured = New Collection
    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If Not RowIsTrailer(sourceTable, rowIndex) Then
            If LenB(RowRejectReason(sourceTable, rowIndex, rejectCategory, taxApply, taxCod, taxRecov)) = 0 Then
                invoiceNumberValue = CellText(sourceTable(rowIndex, mColInvoiceNumber))
                companyValue = CellText(sourceTable(rowIndex, mColCompany))
                supplierValue = CellText(sourceTable(rowIndex, mColSupplierId))
                resolvedCurrency = CurrencyValue(sourceTable(rowIndex, mColCurrency))
                invoiceDateValue = NormalizeSourceDate(sourceTable(rowIndex, mColInvoiceDate))
                If Not CollectionKeyExists(groupCompanyCaptured, invoiceNumberValue) Then
                    groupCompanyCaptured.Add companyValue, invoiceNumberValue
                    groupSupplierCaptured.Add supplierValue, invoiceNumberValue
                    groupCurrencyCaptured.Add resolvedCurrency, invoiceNumberValue
                    groupInvDateCaptured.Add invoiceDateValue, invoiceNumberValue
                Else
                    prepassMsg = vbNullString
                    If StrComp(CStr(groupCompanyCaptured(invoiceNumberValue)), companyValue, vbBinaryCompare) <> 0 Then _
                        prepassMsg = prepassMsg & "Company(" & CStr(groupCompanyCaptured(invoiceNumberValue)) & _
                            " vs " & companyValue & "); "
                    If StrComp(CStr(groupSupplierCaptured(invoiceNumberValue)), supplierValue, vbBinaryCompare) <> 0 Then _
                        prepassMsg = prepassMsg & "Supplier(" & CStr(groupSupplierCaptured(invoiceNumberValue)) & _
                            " vs " & supplierValue & "); "
                    If StrComp(CStr(groupCurrencyCaptured(invoiceNumberValue)), resolvedCurrency, vbBinaryCompare) <> 0 Then _
                        prepassMsg = prepassMsg & "Currency(" & CStr(groupCurrencyCaptured(invoiceNumberValue)) & _
                            " vs " & resolvedCurrency & "); "
                    If StrComp(CStr(groupInvDateCaptured(invoiceNumberValue)), invoiceDateValue, vbBinaryCompare) <> 0 Then _
                        prepassMsg = prepassMsg & "InvoiceDate(" & CStr(groupInvDateCaptured(invoiceNumberValue)) & _
                            " vs " & invoiceDateValue & "); "
                    If LenB(prepassMsg) > 0 Then
                        Err.Raise vbObjectError + 1912, _
                            "modApAmexHotelV3SupplierInvoice.RunApAmexHotelV3SupplierInvoice", _
                            "Header Consistency Guard failed for InvoiceNumber '" & invoiceNumberValue & _
                            "'. Source rows have diverged header fields: " & prepassMsg
                    End If
                End If
            End If
        End If
    Next rowIndex

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then _
            UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime

        If RowIsTrailer(sourceTable, rowIndex) Then
            If RowIsTrulyEmpty(sourceTable, rowIndex, _
                Array(mColSupplierId, mColInvoiceNumber, mColDescription), Array(mColNet, mColGross)) Then
                consecutiveEmptyRows = consecutiveEmptyRows + 1
                If consecutiveEmptyRows >= 3 Then Exit For
            Else
                consecutiveEmptyRows = 0
            End If
        Else
            rejectReason = RowRejectReason(sourceTable, rowIndex, rejectCategory, taxApply, taxCod, taxRecov)
            If LenB(rejectReason) > 0 Then
                rowsSkipped = rowsSkipped + 1
                Select Case rejectCategory
                    Case REJECT_MISSING_FIELD: missingFieldRows = missingFieldRows + 1
                    Case REJECT_SUPPLIER: unresolvedSupplierRows = unresolvedSupplierRows + 1
                    Case REJECT_WORKTAG: worktagRejectRows = worktagRejectRows + 1
                    Case REJECT_TAX_CODE: unknownTaxCodeRows = unknownTaxCodeRows + 1
                    Case REJECT_NEGATIVE_NET: negativeNetRows = negativeNetRows + 1
                End Select
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, rejectReason, _
                    CellText(sourceTable(rowIndex, mColSupplierId)), CellText(sourceTable(rowIndex, mColDescription))
            Else
                WriteOutputRow targetSheet, targetRow, sourceTable, rowIndex, taxApply, taxCod, taxRecov
                rowsProcessed = rowsProcessed + 1
                targetRow = targetRow + 1
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & _
            " source row(s) with a missing required field (InvoiceNumber / SupplierID / InvoiceDate / Company / " & _
            "TaxCode blank, or Net blank/zero)."
    End If
    If unresolvedSupplierRows > 0 Then
        issueParts.Add "Skipped " & CStr(unresolvedSupplierRows) & _
            " source row(s) whose SupplierID is not in the Supplier Spreadsheet '" & _
            SUPPLIER_EXTRACT_ID_HEADER & "' column."
    End If
    If worktagRejectRows > 0 Then
        issueParts.Add "Skipped " & CStr(worktagRejectRows) & _
            " source row(s) with a blank mandatory worktag or a worktag = " & SOURCE_UNRESOLVED_MARKER & "."
    End If
    If unknownTaxCodeRows > 0 Then
        issueParts.Add "Skipped " & CStr(unknownTaxCodeRows) & _
            " source row(s) with unknown TaxCode (only P20 / P0 / PZ accepted)."
    End If
    If negativeNetRows > 0 Then
        issueParts.Add "Rejected " & CStr(negativeNetRows) & " source row(s) because Net is negative."
    End If

    outputPath = outputWorkbook.FullName
    outputWorkbook.Save
    outputWorkbook.Close SaveChanges:=False
    Set outputWorkbook = Nothing

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Completed", rowsProcessed, rowsSkipped, _
        runStartTime, runEndTime, outputPath, JoinCollectionValues(issueParts, " | "))
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
    If Not outputWorkbook Is Nothing Then outputWorkbook.Close SaveChanges:=False
    On Error GoTo 0

    runEndTime = Now
    logIssue = AppendRunLog(runnerWorkbook, selectionKey, "Failed", rowsProcessed, rowsSkipped, _
        runStartTime, runEndTime, vbNullString, failureMessage)
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
        "modApAmexHotelV3SupplierInvoice.RunApAmexHotelV3SupplierInvoice", failureMessage
End Sub

Private Sub ResolveSourceColumns(ByVal sourceTable As Variant, ByVal headerRowIndex As Long)
    Dim i As Long

    mColSupplierId = RequiredColumn(sourceTable, headerRowIndex, "SupplierID")
    mColInvoiceNumber = RequiredColumn(sourceTable, headerRowIndex, "InvoiceNumber")
    mColInvoiceDate = RequiredColumn(sourceTable, headerRowIndex, "InvoiceDate")
    mColCurrency = RequiredColumn(sourceTable, headerRowIndex, "Currency")
    mColCompany = RequiredColumn(sourceTable, headerRowIndex, "Company")
    mColNet = RequiredColumn(sourceTable, headerRowIndex, "Net")
    mColGross = RequiredColumn(sourceTable, headerRowIndex, "Gross")
    mColDescription = RequiredColumn(sourceTable, headerRowIndex, "Description")
    mColTaxCode = RequiredColumn(sourceTable, headerRowIndex, "TaxCode")
    mColHotelStayMemo = RequiredColumn(sourceTable, headerRowIndex, "Hotel Stay Memo")

    SetWorktag 1, "Spend Category", True
    SetWorktag 2, "Reporting Type", True
    SetWorktag 3, "Cost Centre", True
    SetWorktag 4, "Event", False
    SetWorktag 5, "Camp", False
    SetWorktag 6, "FA Learning Course", False
    SetWorktag 7, "Commercial Partner", False
    SetWorktag 8, "Grant", False
    SetWorktag 9, "Programme", False
    SetWorktag 10, "Project", False
    For i = 1 To WORKTAG_COUNT
        mWorktagCols(i) = RequiredColumn(sourceTable, headerRowIndex, mWorktagHeaders(i))
    Next i
End Sub

Private Sub SetWorktag(ByVal slot As Long, ByVal headerText As String, ByVal isMandatory As Boolean)
    mWorktagHeaders(slot) = headerText
    mWorktagMandatory(slot) = isMandatory
End Sub

Private Function RequiredColumn(ByVal sourceTable As Variant, ByVal headerRowIndex As Long, _
    ByVal headerText As String) As Long
    RequiredColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, headerText)
    If RequiredColumn = 0 Then
        Err.Raise vbObjectError + 1911, "modApAmexHotelV3SupplierInvoice.RequiredColumn", _
            "Workday_Upload header '" & headerText & "' not found on the header row."
    End If
End Function

Private Function BuildSupplierIdSet(ByVal supplierWorkbook As Workbook) As Collection
    Dim extractTable As Variant
    Dim headerRow As Long
    Dim supplierIdColumn As Long

    extractTable = ReadWorksheetTable(ResolveSourceWorksheet(supplierWorkbook, vbNullString))
    headerRow = FindHeaderRow(extractTable, _
        Array(SUPPLIER_EXTRACT_ENTITY, SUPPLIER_EXTRACT_REFERENCE_HEADER, SUPPLIER_EXTRACT_ID_HEADER))
    supplierIdColumn = FindColumnIndexInRow(extractTable, headerRow, SUPPLIER_EXTRACT_ID_HEADER)
    If supplierIdColumn = 0 Then
        Err.Raise vbObjectError + 1913, "modApAmexHotelV3SupplierInvoice.BuildSupplierIdSet", _
            "Supplier Spreadsheet has no '" & SUPPLIER_EXTRACT_ID_HEADER & "' column on its header row."
    End If
    Set BuildSupplierIdSet = BuildDelimitedReferenceLookup(extractTable, headerRow, _
        supplierIdColumn, supplierIdColumn)
End Function

Private Function RowIsTrailer(ByVal sourceTable As Variant, ByVal rowIndex As Long) As Boolean
    RowIsTrailer = RowIsEmptyOrTrailer(sourceTable, rowIndex, _
        Array(mColSupplierId, mColInvoiceNumber, mColDescription), _
        Array(mColNet, mColGross), Array(mColDescription, mColSupplierId, mColInvoiceNumber))
End Function

Private Function RowRejectReason(ByVal sourceTable As Variant, ByVal rowIndex As Long, _
    ByRef rejectCategory As Long, ByRef taxApply As String, ByRef taxCod As String, _
    ByRef taxRecov As String) As String
    Dim supplierId As String
    Dim worktagValue As String
    Dim taxCodeValue As String
    Dim netValue As Double
    Dim i As Long

    rejectCategory = 0
    supplierId = CellText(sourceTable(rowIndex, mColSupplierId))
    taxCodeValue = UCase$(CellText(sourceTable(rowIndex, mColTaxCode)))
    If LenB(CellText(sourceTable(rowIndex, mColNet))) > 0 Then
        netValue = ResolveAmount(sourceTable(rowIndex, mColNet))
        If netValue < 0# Then
            rejectCategory = REJECT_NEGATIVE_NET
            RowRejectReason = "Net amount is negative; row rejected by skip_rule.drop_when Net < 0."
            Exit Function
        End If
    End If

    If LenB(CellText(sourceTable(rowIndex, mColInvoiceNumber))) = 0 Or LenB(supplierId) = 0 _
        Or LenB(CellText(sourceTable(rowIndex, mColInvoiceDate))) = 0 _
        Or LenB(CellText(sourceTable(rowIndex, mColCompany))) = 0 _
        Or LenB(taxCodeValue) = 0 Or LenB(CellText(sourceTable(rowIndex, mColNet))) = 0 _
        Or ResolveAmount(sourceTable(rowIndex, mColNet)) = 0# Then
        rejectCategory = REJECT_MISSING_FIELD
        RowRejectReason = "Missing required Workday_Upload field (InvoiceNumber / SupplierID / InvoiceDate / " & _
            "Company / TaxCode blank, or Net blank/zero)."
        Exit Function
    End If

    If LenB(LookupCollectionValue(mSupplierIds, supplierId)) = 0 Then
        rejectCategory = REJECT_SUPPLIER
        RowRejectReason = "Supplier unresolved: SupplierID '" & supplierId & "' is not in the Supplier Spreadsheet '" & _
            SUPPLIER_EXTRACT_ID_HEADER & "' column."
        Exit Function
    End If

    For i = 1 To WORKTAG_COUNT
        worktagValue = CellText(sourceTable(rowIndex, mWorktagCols(i)))
        If StrComp(worktagValue, SOURCE_UNRESOLVED_MARKER, vbTextCompare) = 0 _
            Or worktagValue = SOURCE_ERROR_TEXT Then
            rejectCategory = REJECT_WORKTAG
            RowRejectReason = "Worktag '" & mWorktagHeaders(i) & "' is " & worktagValue & _
                " (unresolved in the source workbook)."
            Exit Function
        End If
        If mWorktagMandatory(i) And LenB(worktagValue) = 0 Then
            rejectCategory = REJECT_WORKTAG
            RowRejectReason = "Mandatory worktag '" & mWorktagHeaders(i) & "' is blank."
            Exit Function
        End If
    Next i

    If Not ResolveTaxTriple(taxCodeValue, taxApply, taxCod, taxRecov) Then
        rejectCategory = REJECT_TAX_CODE
        RowRejectReason = "Unknown TaxCode '" & taxCodeValue & _
            "' -- expected P20/P0/PZ. Per spec on_unknown_when_value = reject_and_report."
        Exit Function
    End If
End Function

Private Function ResolveTaxTriple(ByVal taxCodeValue As String, ByRef taxApply As String, _
    ByRef taxCod As String, ByRef taxRecov As String) As Boolean
    If taxCodeValue = "P20" Then
        taxApply = TAX_APPLICABILITY_STANDARD
        taxCod = TAX_CODE_STANDARD
        taxRecov = TAX_RECOVERABILITY_STANDARD
        ResolveTaxTriple = True
    ElseIf taxCodeValue = "P0" Or taxCodeValue = "PZ" Then
        taxApply = TAX_APPLICABILITY_ZERO
        taxCod = TAX_CODE_ZERO
        taxRecov = TAX_RECOVERABILITY_ZERO
        ResolveTaxTriple = True
    Else
        taxApply = vbNullString
        taxCod = vbNullString
        taxRecov = vbNullString
        ResolveTaxTriple = False
    End If
End Function

Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        CellText = SOURCE_ERROR_TEXT
    Else
        CellText = Trim$(CStr(cellValue))
    End If
End Function

Private Function CurrencyValue(ByVal cellValue As Variant) As String
    CurrencyValue = CellText(cellValue)
    If LenB(CurrencyValue) = 0 Then CurrencyValue = OUTPUT_DEFAULT_CURRENCY
End Function

Private Sub WriteOutputRow(ByVal targetSheet As Worksheet, ByVal targetRow As Long, _
    ByVal sourceTable As Variant, ByVal sourceRow As Long, ByVal taxApply As String, _
    ByVal taxCod As String, ByVal taxRecov As String)
    Dim supplierId As String
    Dim invoiceNumberValue As String
    Dim spreadsheetKeyValue As Long
    Dim lineOrderValue As Long

    supplierId = CellText(sourceTable(sourceRow, mColSupplierId))
    invoiceNumberValue = CellText(sourceTable(sourceRow, mColInvoiceNumber))
    spreadsheetKeyValue = AssignInvoiceLineGroup(mGrpKey, mGrpLine, invoiceNumberValue, _
        mGrpNextKey, lineOrderValue)

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = spreadsheetKeyValue
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).NumberFormat = TEXT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = CellText(sourceTable(sourceRow, mColCompany))
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = CurrencyValue(sourceTable(sourceRow, mColCurrency))
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = supplierId
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = TEXT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = NormalizeSourceDate(sourceTable(sourceRow, mColInvoiceDate))
    targetSheet.Cells(targetRow, ColumnIndex("AG")).NumberFormat = TEXT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("AG")).Value = Empty
    targetSheet.Cells(targetRow, ColumnIndex("AS")).Value = Empty
    targetSheet.Cells(targetRow, ColumnIndex("BG")).NumberFormat = TEXT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("EL")).Value = Empty
    targetSheet.Cells(targetRow, ColumnIndex("EM")).Value = lineOrderValue
    targetSheet.Cells(targetRow, ColumnIndex("EK")).Value = lineOrderValue
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = UCase$(CellText(sourceTable(sourceRow, mColDescription)))
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = taxApply
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = taxCod
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = taxRecov
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = Empty
    targetSheet.Cells(targetRow, ColumnIndex("GF")).NumberFormat = AMOUNT_NUMBER_FORMAT
    targetSheet.Cells(targetRow, ColumnIndex("GF")).Value = ResolveAmount(sourceTable(sourceRow, mColNet))
    targetSheet.Cells(targetRow, ColumnIndex("GQ")).Value = supplierId & MEMO_SEPARATOR & invoiceNumberValue
    targetSheet.Cells(targetRow, ColumnIndex("IL")).Value = CellText(sourceTable(sourceRow, mColHotelStayMemo))
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = CellText(sourceTable(sourceRow, mWorktagCols(1)))
    targetSheet.Cells(targetRow, ColumnIndex("GR")).Value = CellText(sourceTable(sourceRow, mWorktagCols(2)))
    targetSheet.Cells(targetRow, ColumnIndex("GT")).Value = CellText(sourceTable(sourceRow, mWorktagCols(3)))
    targetSheet.Cells(targetRow, ColumnIndex("GV")).Value = CellText(sourceTable(sourceRow, mWorktagCols(4)))
    targetSheet.Cells(targetRow, ColumnIndex("GX")).Value = CellText(sourceTable(sourceRow, mWorktagCols(5)))
    targetSheet.Cells(targetRow, ColumnIndex("GZ")).Value = CellText(sourceTable(sourceRow, mWorktagCols(6)))
    targetSheet.Cells(targetRow, ColumnIndex("HB")).Value = CellText(sourceTable(sourceRow, mWorktagCols(7)))
    targetSheet.Cells(targetRow, ColumnIndex("HD")).Value = CellText(sourceTable(sourceRow, mWorktagCols(8)))
    targetSheet.Cells(targetRow, ColumnIndex("HF")).Value = CellText(sourceTable(sourceRow, mWorktagCols(9)))
    targetSheet.Cells(targetRow, ColumnIndex("HH")).Value = CellText(sourceTable(sourceRow, mWorktagCols(10)))
End Sub
