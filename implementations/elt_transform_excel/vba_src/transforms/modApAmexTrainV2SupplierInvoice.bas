Attribute VB_Name = "modApAmexTrainV2SupplierInvoice"
Option Explicit
' Source of truth: config/transform_specs/ap_amex_train_v2.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together
Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const SOURCE_SUPPLIER_REFERENCE_ID As String = "AME001"

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_SPEND_CATEGORY As String = "Staff_Travel_Trains"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"


Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"

Private Const FINAL_CODE_PADDED_LENGTH As Long = 22
Private Const FINAL_CODE_SEG1_LEN As Long = 4
Private Const FINAL_CODE_SEG2_LEN As Long = 6
Private Const FINAL_CODE_SEG3_LEN As Long = 8
Private Const FINAL_CODE_SEG4_LEN As Long = 4
Private Const FINAL_CODE_ERROR As String = "Error"

Private mSupplierId As String
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

Private mHeaderInvoiceNumber2 As String
Private mHeaderInvoiceDate As String
Private mHeaderSupplierInvoiceNumber As String
Private mHeaderItemDescription As String

Public Sub RunApAmexTrainsV2SupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim finalCodeErrorRows As Long
    Dim rowIndex As Long
    Dim accountNameColumn As Long
    Dim ref2Column As Long
    Dim statementIdColumn As Long
    Dim statementDateColumn As Long
    Dim currencyColumn As Long
    Dim netAmountColumn As Long
    Dim vatColumn As Long
    Dim travellerNameColumn As Long
    Dim travelRef4Column As Long
    Dim departureAirportColumn As Long
    Dim departureDateColumn As Long
    Dim destinationAirportColumn As Long
    Dim travellerFirstNameColumn As Long
    Dim bookingSourceColumn As Long
    Dim helperSeg1Column As Long
    Dim helperSeg2Column As Long
    Dim helperSeg3Column As Long
    Dim helperSeg4Column As Long
    Dim seg1OvColumn As Long
    Dim seg2OvColumn As Long
    Dim seg3OvColumn As Long
    Dim seg4OvColumn As Long
    Dim supplierId As String
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
    Dim fillDownColumns As Variant
    Dim fillIndex As Long
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
        Err.Raise vbObjectError + 2101, "modApAmexTrainV2SupplierInvoice", _
            "Unable to resolve supplier Reference ID '" & SOURCE_SUPPLIER_REFERENCE_ID & "' to a Workday Supplier ID."
    End If
    mSupplierId = supplierId

    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApAmexTrainV2SupplierInvoice", 15)

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
        Err.Raise vbObjectError + 1910, "modApAmexTrainV2SupplierInvoice.RunApAmexTrainV2SupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    accountNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Account name")
    ref2Column = FindColumnIndexInRow(sourceTable, headerRowIndex, "Ref 2")
    statementIdColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Statement ID")
    statementDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Statement date")
    currencyColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Currency")
    netAmountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Net amount")
    vatColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "VAT")
    travellerNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Traveller name")

    travelRef4Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Travel ref 4")
    departureAirportColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Departure airport")
    departureDateColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Departure date")
    destinationAirportColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Destination airport")
    travellerFirstNameColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Traveller first name")
    bookingSourceColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Booking Source")

    helperSeg1Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg1")
    helperSeg2Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg2")
    helperSeg3Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg3")
    helperSeg4Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg4")

    seg1OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg1 Overwrite")
    seg2OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg2 Overwrite")
    seg3OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg3 Overwrite3")
    If seg3OvColumn = 0 Then
        seg3OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg3 Overwrite")
    End If
    seg4OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg4 Overwrite")

    fillDownColumns = Array( _
        travelRef4Column, _
        departureAirportColumn, _
        departureDateColumn, _
        destinationAirportColumn, _
        travellerFirstNameColumn, _
        bookingSourceColumn _
    )
    For fillIndex = LBound(fillDownColumns) To UBound(fillDownColumns)
        If CLng(fillDownColumns(fillIndex)) > 0 Then
            FillDownColumn sourceTable, CLng(fillDownColumns(fillIndex)), dataStartRowIndex
        End If
    Next fillIndex

    BuildHeaderLevelValues sourceTable, dataStartRowIndex, sourceDataPath, _
        statementIdColumn, statementDateColumn

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, currencyColumn, ref2Column, netAmountColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Currency / Ref 2 / Net amount)", _
                Trim$(CStr(sourceTable(rowIndex, ref2Column))), _
                BuildTravelMemo(sourceTable, rowIndex, travellerNameColumn, departureDateColumn, destinationAirportColumn)
        Else
            Dim finalCodeValue As String
            Dim hS1 As Variant
            Dim hS2 As Variant
            Dim hS3 As Variant
            Dim hS4 As Variant
            Dim s1Ov As Variant
            Dim s2Ov As Variant
            Dim s3Ov As Variant
            Dim s4Ov As Variant
            If helperSeg1Column > 0 Then hS1 = sourceTable(rowIndex, helperSeg1Column) Else hS1 = Empty
            If helperSeg2Column > 0 Then hS2 = sourceTable(rowIndex, helperSeg2Column) Else hS2 = Empty
            If helperSeg3Column > 0 Then hS3 = sourceTable(rowIndex, helperSeg3Column) Else hS3 = Empty
            If helperSeg4Column > 0 Then hS4 = sourceTable(rowIndex, helperSeg4Column) Else hS4 = Empty
            If seg1OvColumn > 0 Then s1Ov = sourceTable(rowIndex, seg1OvColumn) Else s1Ov = Empty
            If seg2OvColumn > 0 Then s2Ov = sourceTable(rowIndex, seg2OvColumn) Else s2Ov = Empty
            If seg3OvColumn > 0 Then s3Ov = sourceTable(rowIndex, seg3OvColumn) Else s3Ov = Empty
            If seg4OvColumn > 0 Then s4Ov = sourceTable(rowIndex, seg4OvColumn) Else s4Ov = Empty
            finalCodeValue = DeriveFinalCodeFromRef2( _
                sourceTable(rowIndex, ref2Column), _
                hS1, hS2, hS3, hS4, _
                s1Ov, s2Ov, s3Ov, s4Ov)
            If InStr(1, finalCodeValue, FINAL_CODE_ERROR, vbTextCompare) > 0 Then
                rowsSkipped = rowsSkipped + 1
                finalCodeErrorRows = finalCodeErrorRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Final Code derivation error: Ref 2 Seg2 first character did not map to a 6-digit accounting code (expected 3/4/5/6)", _
                    Trim$(CStr(sourceTable(rowIndex, ref2Column))), _
                    BuildTravelMemo(sourceTable, rowIndex, travellerNameColumn, departureDateColumn, destinationAirportColumn)
            Else
                costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, finalCodeValue)
                If LenB(costCenterId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedCostCenterRows = unmappedCostCenterRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Cost centre unresolved: Final Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                        finalCodeValue, _
                        BuildTravelMemo(sourceTable, rowIndex, travellerNameColumn, departureDateColumn, destinationAirportColumn)
                Else
                    WriteOutputRow _
                        targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                        runnerWorkbook, accountNameColumn, currencyColumn, netAmountColumn, _
                        vatColumn, travellerNameColumn, departureDateColumn, destinationAirportColumn, _
                        finalCodeValue, costCenterId
                    rowsProcessed = rowsProcessed + 1
                    targetRow = targetRow + 1
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Currency, Ref 2, or Net amount."
    End If
    If finalCodeErrorRows > 0 Then
        issueParts.Add "Skipped " & CStr(finalCodeErrorRows) & " source row(s) whose Ref 2 Seg2 first character did not map to a 6-digit accounting code (expected 3/4/5/6)."
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
    Err.Raise failureNumber, "modApAmexTrainV2SupplierInvoice.RunApAmexTrainsV2SupplierInvoice", failureMessage
End Sub

Private Sub BuildHeaderLevelValues( _
    ByVal sourceTable As Variant, _
    ByVal dataStartRowIndex As Long, _
    ByVal sourceDataPath As String, _
    ByVal statementIdColumn As Long, _
    ByVal statementDateColumn As Long _
)
    Dim invoiceLabel As String
    Dim firstStatementDateValue As Variant
    Dim firstStatementIdValue As Variant
    Dim ddmmyy As String
    Dim parsedDate As Date

    invoiceLabel = ExtractAmexTrainsInvoiceLabelFromFileName(sourceDataPath)
    mHeaderInvoiceNumber2 = invoiceLabel
    mHeaderItemDescription = invoiceLabel

    firstStatementDateValue = sourceTable(dataStartRowIndex, statementDateColumn)
    mHeaderInvoiceDate = NormalizeSourceDate(firstStatementDateValue)
    If LenB(mHeaderInvoiceDate) = 0 Then
        Err.Raise vbObjectError + 2102, "modApAmexTrainV2SupplierInvoice", _
            "Unable to derive INVOICE DATE from first data row's Statement date. Value='" & CStr(firstStatementDateValue) & "'."
    End If

    If IsDate(firstStatementDateValue) Then
        parsedDate = CDate(firstStatementDateValue)
    Else
        On Error Resume Next
        parsedDate = CDate(firstStatementDateValue)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise vbObjectError + 2103, "modApAmexTrainV2SupplierInvoice", _
                "Unable to parse first data row's Statement date as a Date for SUPPLIER INVOICE NUMBER construction. Value='" & CStr(firstStatementDateValue) & "'."
        End If
        On Error GoTo 0
    End If
    ddmmyy = Format$(parsedDate, "ddmmyy")

    firstStatementIdValue = sourceTable(dataStartRowIndex, statementIdColumn)
    mHeaderSupplierInvoiceNumber = Trim$(CStr(firstStatementIdValue)) & "_" & ddmmyy
End Sub

Private Sub FillDownColumn( _
    ByVal sourceTable As Variant, _
    ByVal columnIndex As Long, _
    ByVal dataStartRowIndex As Long _
)
    Dim rowIndex As Long
    Dim firstValue As String

    If columnIndex <= 0 Then Exit Sub
    If dataStartRowIndex > UBound(sourceTable, 1) Then Exit Sub

    firstValue = Trim$(CStr(sourceTable(dataStartRowIndex, columnIndex)))
    If LenB(firstValue) = 0 Then Exit Sub

    For rowIndex = dataStartRowIndex + 1 To UBound(sourceTable, 1)
        If LenB(Trim$(CStr(sourceTable(rowIndex, columnIndex)))) = 0 Then
            sourceTable(rowIndex, columnIndex) = firstValue
        End If
    Next rowIndex
End Sub

Private Function DeriveFinalCodeFromRef2( _
    ByVal ref2Value As Variant, _
    ByVal helperSeg1 As Variant, _
    ByVal helperSeg2 As Variant, _
    ByVal helperSeg3 As Variant, _
    ByVal helperSeg4 As Variant, _
    ByVal seg1Override As Variant, _
    ByVal seg2Override As Variant, _
    ByVal seg3Override As Variant, _
    ByVal seg4Override As Variant _
) As String
    Dim cleaned As String
    Dim padded As String
    Dim seg1 As String
    Dim seg2 As String
    Dim seg3 As String
    Dim seg4 As String
    Dim seg2Left As String
    Dim seg2Mapped As String
    Dim useHS1 As String
    Dim useHS2 As String
    Dim useHS3 As String
    Dim useHS4 As String
    Dim useSeg1Ov As String
    Dim useSeg2Ov As String
    Dim useSeg3Ov As String
    Dim useSeg4Ov As String
    Dim seg2IsError As Boolean

    DeriveFinalCodeFromRef2 = vbNullString

    cleaned = Trim$(CStr(ref2Value))
    If LenB(cleaned) = 0 Then
        Exit Function
    End If
    cleaned = Replace(cleaned, "-", vbNullString)
    cleaned = Replace(cleaned, " ", vbNullString)
    cleaned = Replace(cleaned, ".", vbNullString)

    If Len(cleaned) < FINAL_CODE_PADDED_LENGTH Then
        padded = cleaned & String$(FINAL_CODE_PADDED_LENGTH - Len(cleaned), "0")
    Else
        padded = Left$(cleaned, FINAL_CODE_PADDED_LENGTH)
    End If

    seg1 = Mid$(padded, 1, FINAL_CODE_SEG1_LEN)
    seg2 = Mid$(padded, FINAL_CODE_SEG1_LEN + 1, FINAL_CODE_SEG2_LEN)
    seg3 = Mid$(padded, FINAL_CODE_SEG1_LEN + FINAL_CODE_SEG2_LEN + 1, FINAL_CODE_SEG3_LEN)
    seg4 = Mid$(padded, FINAL_CODE_SEG1_LEN + FINAL_CODE_SEG2_LEN + FINAL_CODE_SEG3_LEN + 1, FINAL_CODE_SEG4_LEN)

    seg2Left = Left$(seg2, 1)
    seg2Mapped = MapSeg2LeftToAccountingCode(seg2Left)
    seg2IsError = (LenB(seg2Mapped) = 0)

    ' Step 2 -- V1 helper Seg override. Same format contract as Hotels:
    '   * helperSeg2 = RAW 6-char placeholder (e.g. "4XXXXX"), run seg2Left remap.
    '   * helperSeg1/3/4 = direct slice replacement.
    useHS1 = Trim$(CStr(helperSeg1))
    useHS2 = Trim$(CStr(helperSeg2))
    useHS3 = Trim$(CStr(helperSeg3))
    useHS4 = Trim$(CStr(helperSeg4))

    If LenB(useHS1) > 0 Then seg1 = useHS1

    If LenB(useHS2) > 0 Then
        seg2 = useHS2
        seg2Left = Left$(seg2, 1)
        seg2Mapped = MapSeg2LeftToAccountingCode(seg2Left)
        seg2IsError = (LenB(seg2Mapped) = 0)
    End If

    If LenB(useHS3) > 0 Then seg3 = useHS3

    If LenB(useHS4) > 0 Then seg4 = useHS4

    ' Step 3 -- Tier-2 Seg Overwrite cols (Seg2 Overwrite = final accounting code, clears error).
    useSeg1Ov = Trim$(CStr(seg1Override))
    useSeg2Ov = Trim$(CStr(seg2Override))
    useSeg3Ov = Trim$(CStr(seg3Override))
    useSeg4Ov = Trim$(CStr(seg4Override))

    If LenB(useSeg1Ov) > 0 Then seg1 = useSeg1Ov

    If LenB(useSeg2Ov) > 0 Then
        seg2Mapped = useSeg2Ov
        seg2IsError = False
    End If

    If LenB(useSeg3Ov) > 0 Then seg3 = useSeg3Ov

    If LenB(useSeg4Ov) > 0 Then seg4 = useSeg4Ov

    If seg2IsError Then
        DeriveFinalCodeFromRef2 = seg1 & "-" & FINAL_CODE_ERROR & "-" & seg3 & "-" & seg4
        Exit Function
    End If

    DeriveFinalCodeFromRef2 = seg1 & "-" & seg2Mapped & "-" & seg3 & "-" & seg4
End Function

Private Function MapSeg2LeftToAccountingCode(ByVal seg2LeftChar As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(seg2LeftChar))

    Select Case normalized
        Case "3"
            MapSeg2LeftToAccountingCode = "302101"
        Case "4"
            MapSeg2LeftToAccountingCode = "402201"
        Case "5"
            MapSeg2LeftToAccountingCode = "502201"
        Case "6"
            MapSeg2LeftToAccountingCode = "602201"
        Case Else
            MapSeg2LeftToAccountingCode = vbNullString
    End Select
End Function

Private Function DeriveVatPercentage(ByVal vatValue As Variant, ByVal netAmountValue As Variant) As Long
    Dim vatAmount As Double
    Dim netAmount As Double

    vatAmount = ResolveAmount(vatValue)
    netAmount = ResolveAmount(netAmountValue)

    If netAmount <= 0# Then
        DeriveVatPercentage = 0
        Exit Function
    End If

    DeriveVatPercentage = CLng(Int((vatAmount / netAmount) * 100# + 0.5))
End Function

Private Function BuildTravelMemo( _
    ByVal sourceTable As Variant, _
    ByVal rowIndex As Long, _
    ByVal travellerNameColumn As Long, _
    ByVal departureDateColumn As Long, _
    ByVal destinationAirportColumn As Long _
) As String
    Dim travellerName As String
    Dim departureDate As Variant
    Dim destinationAirport As String
    Dim ddmmyy As String
    Dim parsedDate As Date

    travellerName = Trim$(CStr(sourceTable(rowIndex, travellerNameColumn)))
    If departureDateColumn > 0 Then
        departureDate = sourceTable(rowIndex, departureDateColumn)
    Else
        departureDate = vbNullString
    End If
    If destinationAirportColumn > 0 Then
        destinationAirport = Trim$(CStr(sourceTable(rowIndex, destinationAirportColumn)))
    Else
        destinationAirport = vbNullString
    End If

    If LenB(CStr(departureDate)) > 0 Then
        If IsDate(departureDate) Then
            parsedDate = CDate(departureDate)
            ddmmyy = Format$(parsedDate, "ddmmyy")
        Else
            On Error Resume Next
            parsedDate = CDate(departureDate)
            If Err.Number = 0 Then
                ddmmyy = Format$(parsedDate, "ddmmyy")
            Else
                ddmmyy = Replace$(Replace$(Left$(CStr(departureDate), 10), "/", vbNullString), "-", vbNullString)
            End If
            On Error GoTo 0
        End If
    End If

    BuildTravelMemo = Trim$(travellerName & IIf(LenB(ddmmyy) > 0, " " & ddmmyy, vbNullString) & IIf(LenB(destinationAirport) > 0, " " & destinationAirport, vbNullString))
End Function

Private Function ExtractAmexTrainsInvoiceLabelFromFileName(ByVal sourceDataPath As String) As String
    Dim fileName As String
    Dim normalized As String
    Dim tokens As Variant
    Dim index As Long
    Dim token As String
    Dim monthIndex As Long
    Dim monthToken As String
    Dim yearToken As String
    Dim prefix As String
    Dim foundMonth As Boolean

    fileName = UCase$(LocalFileNameWithoutExtensionV2(sourceDataPath))
    normalized = Replace(fileName, "_", " ")
    normalized = Replace(normalized, "-", " ")
    normalized = Replace(normalized, ".", " ")
    normalized = Replace(normalized, ",", " ")
    normalized = Application.WorksheetFunction.Trim(normalized)
    tokens = Split(normalized, " ")

    monthIndex = -1
    For index = LBound(tokens) To UBound(tokens)
        token = CStr(tokens(index))
        monthToken = NormalizeMonthTokenV2(token)
        If LenB(monthToken) > 0 And index < UBound(tokens) Then
            yearToken = Trim$(CStr(tokens(index + 1)))
            If Len(yearToken) = 2 And IsNumeric(yearToken) Then
                monthIndex = index
                foundMonth = True
                Exit For
            End If
        End If
    Next index

    If foundMonth Then
        prefix = vbNullString
        For index = LBound(tokens) To monthIndex - 1
            If LenB(prefix) > 0 Then prefix = prefix & " "
            prefix = prefix & StrConv(CStr(tokens(index)), vbProperCase)
        Next index
        ExtractAmexTrainsInvoiceLabelFromFileName = prefix & " " & CStr(tokens(monthIndex)) & " " & CStr(tokens(monthIndex + 1))
    Else
        ExtractAmexTrainsInvoiceLabelFromFileName = "AMEX TRAINS UNKNOWN"
    End If
End Function

Private Function NormalizeMonthTokenV2(ByVal tokenText As String) As String
    Dim normalized As String
    normalized = Left$(UCase$(Trim$(tokenText)), 3)
    Select Case normalized
        Case "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
            NormalizeMonthTokenV2 = normalized
    End Select
End Function

Private Function LocalFileNameWithoutExtensionV2(ByVal filePath As String) As String
    Dim fileName As String
    Dim slashPosition As Long
    Dim colonPosition As Long
    Dim splitPosition As Long
    Dim dotPosition As Long

    slashPosition = InStrRev(filePath, "/")
    colonPosition = InStrRev(filePath, ":")
    splitPosition = slashPosition
    If colonPosition > splitPosition Then
        splitPosition = colonPosition
    End If

    If splitPosition = 0 Then
        fileName = filePath
    Else
        fileName = Mid$(filePath, splitPosition + 1)
    End If

    dotPosition = InStrRev(fileName, ".")
    If dotPosition = 0 Then
        LocalFileNameWithoutExtensionV2 = fileName
    Else
        LocalFileNameWithoutExtensionV2 = Left$(fileName, dotPosition - 1)
    End If
End Function

Private Sub WriteOutputRow( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceTable As Variant, _
    ByVal sourceRow As Long, _
    ByVal runnerWorkbook As Workbook, _
    ByVal accountNameColumn As Long, _
    ByVal currencyColumn As Long, _
    ByVal netAmountColumn As Long, _
    ByVal vatColumn As Long, _
    ByVal travellerNameColumn As Long, _
    ByVal departureDateColumn As Long, _
    ByVal destinationAirportColumn As Long, _
    ByVal finalCodeValue As String, _
    ByVal costCenterId As String _
)
    Dim currencyCode As String
    Dim amount As Double
    Dim vatRate As Long
    Dim companyId As String
    Dim memo As String
    Dim accountName As String

    currencyCode = Trim$(CStr(sourceTable(sourceRow, currencyColumn)))
    amount = ResolveAmount(sourceTable(sourceRow, netAmountColumn))
    vatRate = DeriveVatPercentage(sourceTable(sourceRow, vatColumn), sourceTable(sourceRow, netAmountColumn))

    accountName = Trim$(CStr(sourceTable(sourceRow, accountNameColumn)))
    companyId = ResolveMappedCompany(runnerWorkbook, accountName, OUTPUT_COMPANY_DEFAULT_FALLBACK)

    memo = BuildTravelMemo(sourceTable, sourceRow, travellerNameColumn, departureDateColumn, destinationAirportColumn)

    targetSheet.Cells(targetRow, ColumnIndex("B")).Value = sequenceNumber
    targetSheet.Cells(targetRow, ColumnIndex("I")).Value = OUTPUT_SUBMIT
    targetSheet.Cells(targetRow, ColumnIndex("K")).Value = mHeaderInvoiceNumber2
    targetSheet.Cells(targetRow, ColumnIndex("R")).Value = companyId
    targetSheet.Cells(targetRow, ColumnIndex("T")).Value = currencyCode
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = mSupplierId
    targetSheet.Cells(targetRow, ColumnIndex("AF")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("AF")).Value = mHeaderInvoiceDate
    targetSheet.Cells(targetRow, ColumnIndex("BG")).Value = mHeaderSupplierInvoiceNumber
    targetSheet.Cells(targetRow, ColumnIndex("DD")).Value = 1
    targetSheet.Cells(targetRow, ColumnIndex("EP")).Value = mHeaderItemDescription
    targetSheet.Cells(targetRow, ColumnIndex("EW")).Value = OUTPUT_SPEND_CATEGORY
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = TaxValueForRate(vatRate, TAX_APPLICABILITY_ZERO, TAX_APPLICABILITY_STANDARD)
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = TaxValueForRate(vatRate, TAX_CODE_ZERO, TAX_CODE_STANDARD)
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = TaxValueForRate(vatRate, TAX_RECOVERABILITY_ZERO, TAX_RECOVERABILITY_STANDARD)
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("IK")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("IL")).Value = memo
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
    ByVal ref2Column As Long, _
    ByVal netAmountColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, currencyColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, ref2Column)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, netAmountColumn)) = 0#
End Function
