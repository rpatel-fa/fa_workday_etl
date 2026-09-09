Attribute VB_Name = "modApAmexHotelSupplierInvoice"
Option Explicit
' Source of truth: config/transform_specs/ap_amex_hotel.yaml -- mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner.py enforces spec<->code column agreement -- edit both together
Private Const TARGET_SHEET_NAME As String = "Submit Supplier Invoice"
Private Const TARGET_DATA_START_ROW As Long = 6

Private Const SOURCE_COMPANY_CODE As String = "FA"
Private Const SOURCE_SUPPLIER_REFERENCE_ID As String = "AME001"

Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_SPEND_CATEGORY As String = "Hotel_Staff_Costs"
Private Const OUTPUT_TAX_APPLICABILITY As String = "GBR_Outside_of_Scope"
Private Const OUTPUT_TAX_CODE As String = "GBR_Zero_Rated"
Private Const OUTPUT_TAX_POINT_DATE_TYPE As String = "INVOICE/ADJUSTMENT_DATE"
Private Const OUTPUT_TAX_RECOVERABILITY As String = "Nonrecoverable_Allocate"
Private Const OUTPUT_TAX_OPTION As String = "CALC_TAX_DUE"
Private Const OUTPUT_COMPANY_DEFAULT_FALLBACK As String = "THE FA GROUP - FA"
Private Const DEFAULT_MAPPING_SHEET_NAME As String = "MAPPING MASTER"

Private Const FINAL_CODE_PADDED_LENGTH As Long = 22
Private Const FINAL_CODE_SEG1_LEN As Long = 4
Private Const FINAL_CODE_SEG2_LEN As Long = 6
Private Const FINAL_CODE_SEG3_LEN As Long = 8
Private Const FINAL_CODE_SEG4_LEN As Long = 4
Private Const FINAL_CODE_ERROR As String = "Error"

Private mSupplierId As String
Private mCompanyId As String
Private mEventLookup As Collection
Private mCampLookup As Collection
Private mFaLearningLookup As Collection
Private mCommercialPartnerLookup As Collection
Private mGrantLookup As Collection
Private mProgramLookup As Collection
Private mProjectLookup As Collection

Public Sub RunApAmexHotelSupplierInvoice(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    Dim currencyColumn As Long
    Dim invoiceNumber2Column As Long
    Dim invoiceDateColumn As Long
    Dim supplierInvoiceNumberColumn As Long
    Dim itemDescriptionColumn As Long
    Dim amountColumn As Long
    Dim memoColumn As Long
    Dim travellerNameColumn As Long
    Dim hotelAhColumn As Long
    Dim hotelArrivalDateColumn As Long
    Dim hotelDepartureDateColumn As Long
    Dim defaultCostCodeColumn As Long
    Dim helperSeg1Column As Long
    Dim helperSeg2Column As Long
    Dim helperSeg3Column As Long
    Dim helperSeg4Column As Long
    Dim seg1OvColumn As Long
    Dim seg2OvColumn As Long
    Dim seg3OvColumn As Long
    Dim seg4OvColumn As Long
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
        Err.Raise vbObjectError + 1901, "modApAmexHotelSupplierInvoice", _
            "Unable to resolve supplier Reference ID '" & SOURCE_SUPPLIER_REFERENCE_ID & "' to a Workday Supplier ID."
    End If
    mSupplierId = supplierId

    '  Source of truth: config/transforms/*.yaml source_layout.required_headers, written into Config 'Required Headers Pipe Delimited' by build_runner.py.
    requiredHeaders = ResolveRequiredHeadersFromConfig(controlSheet, "modApAmexHotelSupplierInvoice", 15)

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

    ' ---------------------------------------------------------------------
    ' Schema validation phase (Tier 2 VBA defence-in-depth).
    ' Consolidates all file/sheet/dimension/header-row/required-header/optional-header
    ' problems into ONE error dialog (instead of fail-fix, fail-fix).
    ' The CLI validate_input.py script is Tier 1 (pre-Excel, faster).
    ' ---------------------------------------------------------------------
    Dim schemaMsg As String
    Dim schemaHeaderRow As Long
    Dim schemaDiag As String
    Dim optionalHeadersArr As Variant
    optionalHeadersArr = Array("Seg1 Overwrite", "Seg2 Overwrite", "Seg3 Overwrite", "Seg4 Overwrite")
    schemaMsg = ValidateSourceWorkbookSchema( _
        sheetToResolve, _
        sourceTable, _
        sourceHeaderRow, _
        sourceDataStartRow, _
        requiredHeaders, _
        optionalHeadersArr, _
        schemaHeaderRow, _
        schemaDiag _
    )
    If LenB(schemaMsg) > 0 Then
        Err.Raise vbObjectError + 1910, "modApAmexHotelSupplierInvoice.RunApAmexHotelSupplierInvoice", schemaMsg
    End If
    If LenB(schemaDiag) > 0 Then
        issueParts.Add "[Schema] " & Replace$(schemaDiag, vbCrLf, " | ")
    End If

    headerRowIndex = FindHeaderRow(sourceTable, requiredHeaders)
    dataStartRowIndex = headerRowIndex + (sourceDataStartRow - sourceHeaderRow)

    ValidateHeaderRow sourceTable, headerRowIndex, requiredHeaders

    currencyColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Transaction Account Currency Code")
    invoiceNumber2Column = FindColumnIndexInRow(sourceTable, headerRowIndex, "Transaction ID")
    invoiceDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Transaction Date")
    supplierInvoiceNumberColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Supplier Reference")
    itemDescriptionColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Transaction Merchant Narrative")
    amountColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Transaction Account Amount")
    defaultCostCodeColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Default Expense Cost Code")
    travellerNameColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Traveller Name")
    hotelAhColumn = ColumnIndex("AS")
    hotelArrivalDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Hotel Arrival Date")
    hotelDepartureDateColumn = FindColumnIndexInRow(sourceTable, headerRowIndex, "Hotel Departure Date")
    memoColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Journal Name?")

    helperSeg1Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg1")
    helperSeg2Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg2")
    helperSeg3Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg3")
    helperSeg4Column = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg4")

    seg1OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg1 Overwrite")
    seg2OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg2 Overwrite")
    seg3OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg3 Overwrite")
    seg4OvColumn = FindOptionalColumnIndexInRow(sourceTable, headerRowIndex, "Seg4 Overwrite")

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, "HH"
    targetRow = TARGET_DATA_START_ROW

    companyId = ResolveMappedCompany(runnerWorkbook, SOURCE_COMPANY_CODE, OUTPUT_COMPANY_DEFAULT_FALLBACK)
    mCompanyId = companyId

    For rowIndex = dataStartRowIndex To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If RowIsSkippable(sourceTable, rowIndex, currencyColumn, invoiceNumber2Column, invoiceDateColumn, _
            supplierInvoiceNumberColumn, itemDescriptionColumn, amountColumn, defaultCostCodeColumn, _
            travellerNameColumn, hotelAhColumn, hotelArrivalDateColumn, hotelDepartureDateColumn) Then
            rowsSkipped = rowsSkipped + 1
            missingFieldRows = missingFieldRows + 1
            Dim fallbackMemo As String
            If memoColumn > 0 Then
                fallbackMemo = Trim$(CStr(sourceTable(rowIndex, memoColumn)))
            Else
                fallbackMemo = vbNullString
            End If
            AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                "Missing required field (Currency / Transaction ID / Transaction Date / Supplier Reference / " & _
                "Transaction Merchant Narrative / Amount / Default Expense Cost Code / Traveller Name / Hotel column AH / " & _
                "Hotel Arrival Date / Hotel Departure Date)", _
                vbNullString, _
                Trim$(CStr(sourceTable(rowIndex, itemDescriptionColumn))) & " | " & fallbackMemo
        Else
            Dim personText As String
            Dim hotelText As String
            Dim stayRangeText As String
            Dim memoText As String
            personText = StripHonorificTitle(sourceTable(rowIndex, travellerNameColumn))
            hotelText = Trim$(CStr(sourceTable(rowIndex, hotelAhColumn)))
            stayRangeText = StayDateRangeText(sourceTable(rowIndex, hotelArrivalDateColumn), sourceTable(rowIndex, hotelDepartureDateColumn))
            memoText = personText & " " & hotelText & " " & stayRangeText
            If memoColumn > 0 Then
                Dim journalFallback As String
                journalFallback = Trim$(CStr(sourceTable(rowIndex, memoColumn)))
                If LenB(journalFallback) > 0 And LenB(Trim$(memoText)) = 0 Then
                    memoText = journalFallback
                End If
            End If
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
            finalCodeValue = DeriveFinalCodeFromDefaultCostCode( _
                sourceTable(rowIndex, defaultCostCodeColumn), _
                hS1, hS2, hS3, hS4, _
                s1Ov, s2Ov, s3Ov, s4Ov)
            If InStr(1, finalCodeValue, FINAL_CODE_ERROR, vbTextCompare) > 0 Then
                rowsSkipped = rowsSkipped + 1
                finalCodeErrorRows = finalCodeErrorRows + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                    "Final Code derivation error: Default Expense Cost Code Seg2 first character did not map to a 6-digit accounting code (expected 3/4/5/6) and no Seg2 Override was supplied", _
                    Trim$(CStr(sourceTable(rowIndex, defaultCostCodeColumn))), _
                    memoText
            Else
                costCenterId = ResolveLowestHierarchyLookupAcrossFinalCodeSegments(costCenterLookup, finalCodeValue)
                If LenB(costCenterId) = 0 Then
                    rowsSkipped = rowsSkipped + 1
                    unmappedCostCenterRows = unmappedCostCenterRows + 1
                    AppendRejectRow runnerWorkbook, selectionKey, rowIndex, _
                        "Cost centre unresolved: Final Code did not resolve in '" & mappingSheetName & "' (" & WBO_COST_CENTER & ", Core vs Prism = Core, lowest hierarchy match)", _
                        finalCodeValue, _
                        memoText
                Else
                    WriteOutputRow _
                        targetSheet, targetRow, rowsProcessed + 1, sourceTable, rowIndex, _
                        currencyColumn, invoiceNumber2Column, invoiceDateColumn, supplierInvoiceNumberColumn, itemDescriptionColumn, amountColumn, memoText, _
                        finalCodeValue, costCenterId
                    rowsProcessed = rowsProcessed + 1
                    targetRow = targetRow + 1
                End If
            End If
        End If
    Next rowIndex

    If missingFieldRows > 0 Then
        issueParts.Add "Skipped " & CStr(missingFieldRows) & " source row(s) with missing Currency, Transaction ID, " & _
            "Transaction Date, Supplier Reference, Transaction Merchant Narrative, Transaction Account Amount, " & _
            "Default Expense Cost Code, Traveller Name, Hotel column AH, Hotel Arrival Date, or Hotel Departure Date."
    End If
    If finalCodeErrorRows > 0 Then
        issueParts.Add "Skipped " & CStr(finalCodeErrorRows) & " source row(s) whose Default Expense Cost Code Seg2 first character did not map to a 6-digit accounting code (expected 3/4/5/6) and no Seg2 Override was supplied."
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
    Err.Raise failureNumber, "modApAmexHotelSupplierInvoice.RunApAmexHotelSupplierInvoice", failureMessage
End Sub

Private Function DeriveFinalCodeFromDefaultCostCode( _
    ByVal defaultCostCodeValue As Variant, _
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

    DeriveFinalCodeFromDefaultCostCode = vbNullString

    ' Step 1 -- coerce the raw Variant into a digits-only string with NO scientific
    ' notation.  Two paths, both Mac-safe (no CDec/CDbl trickery):
    '   A) Numeric types (Double/Long/Currency/Decimal) -> Format$(value, "0")
    '      always renders fixed-point, never "1.115302101E+16".
    '   B) String type -> Trim$, then defensive pure-string expander if E+/E- is
    '      present (copes with a cell that arrived pre-formatted).
    If IsNumeric(defaultCostCodeValue) Then
        On Error Resume Next
        cleaned = Format$(defaultCostCodeValue, "0")
        If Err.Number <> 0 Then
            cleaned = Trim$(CStr(defaultCostCodeValue))
        End If
        Err.Clear
        On Error GoTo 0
    Else
        cleaned = Trim$(CStr(defaultCostCodeValue))
    End If
    If LenB(cleaned) = 0 Then
        Exit Function
    End If
    ' Defensive pure-string scientific-notation expander.
    If InStr(1, cleaned, "E+", vbTextCompare) > 0 Or InStr(1, cleaned, "E-", vbTextCompare) > 0 Then
        cleaned = ExpandScientificNotation(cleaned)
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

    ' Step 2 -- V1 helper Seg override. These are Finance-added "Seg1", "Seg2", "Seg3", "Seg4"
    ' columns that sit between raw Default Cost Code / Ref 2 and the later "Overwrite" escape
    ' hatch.  Critical format difference:
    '   * helperSeg2 stores the RAW 6-character placeholder (e.g. "4XXXXX"), NOT the 6-digit
    '     accounting code.  After replacing seg2, we RE-RUN seg2Left -> MapSeg2Left -> seg2Mapped
    '     recomputation so "4XXXXX" -> left char "4" -> "402201" accounting code.
    '   * helperSeg1/Seg3/Seg4 always replace the corresponding slice outright.
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

    ' Step 3 -- Tier-2 Seg Override cols (Seg1 Overwrite / Seg2 Overwrite / Seg3 Overwrite /
    ' Seg4 Overwrite).  Format note: Seg2 Overwrite stores the FINAL 6-digit accounting code
    ' (e.g. "402201"), NOT the 6-char placeholder.  Using it always clears seg2IsError.
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
        DeriveFinalCodeFromDefaultCostCode = seg1 & "-" & FINAL_CODE_ERROR & "-" & seg3 & "-" & seg4
        Exit Function
    End If

    DeriveFinalCodeFromDefaultCostCode = seg1 & "-" & seg2Mapped & "-" & seg3 & "-" & seg4
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

' Pure-string scientific-notation expander (no CDec/CDbl, works on Mac VBA).
' Only called when the incoming string already contains "E+" or "E-".
' Handles: integer mantissa ("1E+16"), dotted mantissa ("1.115302101E+16"), positive
' and negative exponents, leading +/- signs.  If the shape is malformed returns
' the input untouched so downstream strippers still do their best.
Private Function ExpandScientificNotation(ByVal s As String) As String
    Dim u As String
    Dim ePos As Long
    Dim signPos As Long
    Dim mantissa As String
    Dim expStr As String
    Dim exponent As Long
    Dim dotPos As Long
    Dim intPart As String
    Dim fracPart As String
    Dim digits As String
    Dim newDotPos As Long
    Dim padLeft As Long
    Dim padRight As Long
    Dim i As Long

    u = UCase$(Trim$(s))
    ExpandScientificNotation = u

    ePos = InStr(1, u, "E", vbBinaryCompare)
    If ePos <= 1 Then Exit Function

    mantissa = Left$(u, ePos - 1)
    expStr = Mid$(u, ePos + 1)

    ' Accept optional leading +/- on exponent and mantissa
    If Left$(mantissa, 1) = "+" Or Left$(mantissa, 1) = "-" Then
        mantissa = Mid$(mantissa, 2)
    End If
    signPos = 1
    If Left$(expStr, 1) = "+" Then expStr = Mid$(expStr, 2)
    If Left$(expStr, 1) = "-" Then
        signPos = -1
        expStr = Mid$(expStr, 2)
    End If
    If Not IsNumeric(expStr) Then Exit Function
    exponent = CLng(expStr)

    dotPos = InStr(1, mantissa, ".", vbBinaryCompare)
    If dotPos > 0 Then
        intPart = Left$(mantissa, dotPos - 1)
        fracPart = Mid$(mantissa, dotPos + 1)
    Else
        intPart = mantissa
        fracPart = vbNullString
    End If
    ' Reject non-digit mantissa (defensive -- should never happen if the guard
    ' in the caller only invokes on numeric-origin strings, but belt-and-braces).
    For i = 1 To Len(intPart)
        If Mid$(intPart, i, 1) < "0" Or Mid$(intPart, i, 1) > "9" Then Exit Function
    Next i
    For i = 1 To Len(fracPart)
        If Mid$(fracPart, i, 1) < "0" Or Mid$(fracPart, i, 1) > "9" Then Exit Function
    Next i

    digits = intPart & fracPart
    If Len(digits) = 0 Then
        ExpandScientificNotation = "0"
        Exit Function
    End If

    ' If there had been a dot, its original position was 1 BEFORE the frac part starts
    ' = Len(intPart) in the "digits" string (indexed 1-based from first mantissa digit).
    If dotPos > 0 Then
        newDotPos = Len(intPart) + (signPos * exponent)
    Else
        ' integer mantissa: implicit dot is at position Len(intPart), i.e. "1234" = "1234."
        newDotPos = Len(intPart) + (signPos * exponent)
    End If

    If newDotPos <= 0 Then
        ' Exponent pushed dot fully left of the first digit: zero-padded result.
        padLeft = 1 - newDotPos
        ExpandScientificNotation = "0" & String$(padLeft, "0") & digits
    ElseIf newDotPos >= Len(digits) Then
        padRight = newDotPos - Len(digits)
        ExpandScientificNotation = digits & String$(padRight, "0")
    Else
        ExpandScientificNotation = Left$(digits, newDotPos) & Mid$(digits, newDotPos + 1)
    End If
End Function

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
    ByVal amountColumn As Long, _
    ByVal memoText As String, _
    ByVal finalCodeValue As String, _
    ByVal costCenterId As String _
)
    Dim currencyCode As String
    Dim invoiceDate As String
    Dim amount As Double

    currencyCode = Trim$(CStr(sourceTable(sourceRow, currencyColumn)))
    invoiceDate = NormalizeSourceDate(sourceTable(sourceRow, invoiceDateColumn))
    amount = ResolveAmount(sourceTable(sourceRow, amountColumn))

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
    targetSheet.Cells(targetRow, ColumnIndex("FE")).Value = OUTPUT_TAX_APPLICABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FF")).Value = OUTPUT_TAX_CODE
    targetSheet.Cells(targetRow, ColumnIndex("FH")).Value = OUTPUT_TAX_POINT_DATE_TYPE
    targetSheet.Cells(targetRow, ColumnIndex("FK")).Value = OUTPUT_TAX_RECOVERABILITY
    targetSheet.Cells(targetRow, ColumnIndex("FL")).Value = OUTPUT_TAX_OPTION
    targetSheet.Cells(targetRow, ColumnIndex("GE")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("IK")).Value = amount
    targetSheet.Cells(targetRow, ColumnIndex("IL")).Value = Trim$(memoText)
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
    ByVal amountColumn As Long, _
    ByVal defaultCostCodeColumn As Long, _
    ByVal travellerNameColumn As Long, _
    ByVal hotelAhColumn As Long, _
    ByVal hotelArrivalDateColumn As Long, _
    ByVal hotelDepartureDateColumn As Long _
) As Boolean
    RowIsSkippable = _
        LenB(Trim$(CStr(tableData(rowIndex, currencyColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceNumber2Column)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, invoiceDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, supplierInvoiceNumberColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, itemDescriptionColumn)))) = 0 Or _
        ResolveAmount(tableData(rowIndex, amountColumn)) = 0# Or _
        LenB(Trim$(CStr(tableData(rowIndex, defaultCostCodeColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, travellerNameColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, hotelAhColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, hotelArrivalDateColumn)))) = 0 Or _
        LenB(Trim$(CStr(tableData(rowIndex, hotelDepartureDateColumn)))) = 0
End Function

Private Function StripHonorificTitle(ByVal nameCellValue As Variant) As String
    Dim rawText As String
    Dim i As Long
    Dim PREFIXES As Variant
    Dim thisPrefix As String
    Dim compareText As String

    If IsError(nameCellValue) Or IsEmpty(nameCellValue) Or IsNull(nameCellValue) Then
        StripHonorificTitle = vbNullString
        Exit Function
    End If
    rawText = Trim$(CStr(nameCellValue))
    If LenB(rawText) = 0 Then
        StripHonorificTitle = vbNullString
        Exit Function
    End If

    PREFIXES = Array("DR.", "DR ", "DOCTOR ", "MRS.", "MRS ", "MISS ", "MR.", "MR ", "MS.", "MS ")
    compareText = UCase$(rawText)
    For i = LBound(PREFIXES) To UBound(PREFIXES)
        thisPrefix = CStr(PREFIXES(i))
        If Len(compareText) >= Len(thisPrefix) Then
            If Left$(compareText, Len(thisPrefix)) = thisPrefix Then
                StripHonorificTitle = Trim$(Mid$(rawText, Len(thisPrefix) + 1))
                Exit Function
            End If
        End If
    Next i
    StripHonorificTitle = rawText
End Function

Private Function StayDateRangeText(ByVal arrivalValue As Variant, ByVal departureValue As Variant) As String
    Dim arrivalText As String
    Dim departureText As String

    arrivalText = FormatHotelDateToDdmmyy(arrivalValue)
    departureText = FormatHotelDateToDdmmyy(departureValue)
    If LenB(arrivalText) > 0 And LenB(departureText) > 0 Then
        StayDateRangeText = arrivalText & "-" & departureText
    Else
        StayDateRangeText = arrivalText & departureText
    End If
End Function

Private Function FormatHotelDateToDdmmyy(ByVal sourceValue As Variant) As String
    Dim dtValue As Date
    Dim hasDate As Boolean
    Dim sText As String
    Dim testParts() As String
    Dim MM As Long, DD As Long, YY As Long, YYlong As Long

    hasDate = False
    FormatHotelDateToDdmmyy = vbNullString

    If IsError(sourceValue) Or IsEmpty(sourceValue) Or IsNull(sourceValue) Then
        Exit Function
    End If

    If IsDate(sourceValue) Then
        dtValue = CDate(sourceValue)
        hasDate = True
    Else
        sText = Trim$(CStr(sourceValue))
        If LenB(sText) = 0 Then Exit Function

        If IsNumeric(sText) Then
            On Error Resume Next
            dtValue = CDate(CDbl(sText))
            If Err.Number = 0 Then hasDate = True
            Err.Clear
            On Error GoTo 0
        End If

        If Not hasDate Then
            testParts = Split(Replace$(Replace$(sText, "-", "/"), ".", "/"), "/")
            If UBound(testParts) = 2 Then
                If Len(testParts(2)) = 4 Or Len(testParts(2)) = 2 Then
                    On Error Resume Next
                    MM = CLng(testParts(0)): DD = CLng(testParts(1)): YYlong = CLng(testParts(2))
                    If Err.Number = 0 Then
                        If YYlong < 100 Then YYlong = YYlong + 2000
                        If MM >= 1 And MM <= 12 And DD >= 1 And DD <= 31 And YYlong >= 1990 And YYlong <= 2100 Then
                            dtValue = DateSerial(YYlong, MM, DD)
                            hasDate = True
                        End If
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        End If

        If Not hasDate Then
            On Error Resume Next
            dtValue = CDate(sText)
            If Err.Number = 0 Then hasDate = True
            Err.Clear
            On Error GoTo 0
        End If
    End If

    If hasDate Then
        Dim yrStr As String
        DD = Day(dtValue): MM = Month(dtValue): YY = Year(dtValue) Mod 100
        FormatHotelDateToDdmmyy = Right$("0" & CStr(DD), 2) & Right$("0" & CStr(MM), 2) & Right$("0" & CStr(YY), 2)
    Else
        FormatHotelDateToDdmmyy = Trim$(CStr(sourceValue))
    End If
End Function

