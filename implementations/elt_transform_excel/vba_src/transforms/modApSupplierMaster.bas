Attribute VB_Name = "modApSupplierMaster"
Option Explicit

' AP Supplier input template (sheet "Supplier") -> Workday "Submit Supplier" master EIB (3+ Master
' Suppliers workbook, data from row 6). One input row = one supplier = one EIB block of rows that
' share the Spreadsheet Key (column B). Built like modArCustomerMaster.
'
' 1:1 RULE (operator, 2026-10-12): "only address - i thought we said we only support 1 address 1
' email 1 phone number 1 bank account details and you have stated that for addresses - it means
' spilling to multiple lines since an address is multipe address lines". So ONE EIB row per
' supplier EXCEPT the address lines: each NON-BLANK Address Line 1..3 is written on its own EIB
' row, exactly like the 3+ sample (Master Suppliers rows 6-7: ADDRESS_LINE_1 then ADDRESS_LINE_2).
' Row k of the block carries B (same key), DD (address Row ID = 1, repeated), DM (= k), DO
' (ADDRESS_LINE_k) and DP (the line text); every other field stays on the first row only. Blank
' lines are skipped and k is renumbered contiguously from 1 (same as modArCustomerMaster). The block
' is written by the shared modEntityBlockWriter (registry concern entity_row_block); the letters it
' is given are pinned in config/transforms/ap_supplier_master.yaml supplier_master_eib_columns.
' Single values (operator, 2026-10-12): Use For BILLING on address / phone / email (the sample's
' SHIPPING / REMIT / PROCUREMENT rows are NOT produced), Type BUSINESS, Public / Primary Y, every
' Row ID 1, Supplier Group = the automatic alphabetical band only, Payment Types Accepted = the
' Default Payment Type only, one email, one phone, one settlement (bank) account.
' Supplier ID and Reference ID are typed by AP (required / copied as typed): never generated.
'
' NO FALLBACKS: a missing required value, or a value with no Workday reference ID in the runtime
' "lookups" sheet, rejects the WHOLE supplier (one AppendRejectRow, no EIB rows).
' Status: code-complete, pending manual Excel import + verify (not run in Excel yet).

Private Const TARGET_SHEET_NAME As String = "Submit Supplier"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const CLEAR_END_COLUMN As String = "RH"
Private Const OUTPUT_ADD_ONLY As String = "Y"
Private Const OUTPUT_APPROVAL_STATUS As String = "APPROVED"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_FLAG_YES As String = "Y"
Private Const OUTPUT_FLAG_NO As String = "N"
Private Const OUTPUT_ROW_ID As String = "1"
Private Const OUTPUT_USAGE_TYPE As String = "BUSINESS"
Private Const OUTPUT_USE_FOR As String = "BILLING"
Private Const OUTPUT_ADDRESS_FORMAT As String = "BASIC"
Private Const OUTPUT_EFFECTIVE_DATE As String = "1900-01-01"
Private Const ADDRESS_LINE_TYPE_PREFIX As String = "ADDRESS_LINE_"
Private Const MAX_ADDRESS_LINES As Long = 3
Private Const ADDRESS_ID_SUFFIX As String = "_MAIN_1"
Private Const EMAIL_ID_SUFFIX As String = "_EM1_1"
Private Const GROUP_ASSIGNMENT_AUTO_BAND As String = "auto_band"

Private Const LK_CATEGORY As String = "ref_supplier_category"
Private Const LK_GROUP As String = "ref_supplier_group"
Private Const LK_STATUS As String = "ref_supplier_status"
Private Const LK_PAYMENT_TERM As String = "ref_payment_term"
Private Const LK_PAYMENT_TYPE As String = "ref_payment_type"
Private Const LK_CURRENCY As String = "ref_supplier_currency"
Private Const LK_TAX_STATUS As String = "ref_supplier_transaction_tax_status"
Private Const LK_COUNTRY As String = "ref_country"
Private Const LK_REGION As String = "ref_country_region"
Private Const LK_TAX_ID_TYPE As String = "ref_supplier_country_tax_id_type"
Private Const LK_PHONE_DEVICE As String = "ref_phone_device"
Private Const LK_PO_ISSUE As String = "ref_po_issue_option"
Private Const LK_BANK_TYPE As String = "ref_bank_account_type"

' Source (input template) column positions, resolved once per run from the header row.
Private mColSupId As Long, mColSupName As Long, mColReferenceId As Long
Private mColCategory As Long, mColStatus As Long, mColPaymentTerms As Long, mColPaymentType As Long
Private mColCurrency As Long, mColTaxStatus As Long, mColTaxStatusCountry As Long
Private mColTaxId As Long, mColTaxIdType As Long, mColTaxIdCountry As Long
Private mColPoIssue As Long, mColPayMemo As Long
Private mColBankCountry As Long, mColBankCurrency As Long, mColBankType As Long
Private mColBankName As Long, mColBankAcctName As Long, mColBankAcctNo As Long
Private mColSortCode As Long, mColIban As Long, mColSwift As Long
Private mColRollNo As Long, mColCheckDigit As Long, mColBranchId As Long
Private mColAddrCountry As Long, mColAddrLine1 As Long, mColAddrLine2 As Long, mColAddrLine3 As Long
Private mColCity As Long, mColRegion As Long, mColPostcode As Long
Private mColPhoneCountry As Long, mColPhone As Long, mColPhoneDevice As Long, mColEmail As Long

' The supplier currently being resolved (module level so the stages below stay short).
Private mSupId As String, mName As String, mRefId As String
Private mCategoryId As String, mStatusId As String, mGroupId As String, mTermId As String
Private mPayTypeId As String, mCurrencyId As String, mTaxStatusId As String, mTaxStatusCountry As String
Private mTaxId As String, mTaxIdTypeId As String
Private mPoIssueId As String, mPayMemo As String
Private mBankCountry As String, mBankCurrency As String, mBankType As String, mBankTypeId As String
Private mBankName As String, mBankAcctName As String, mBankAcctNo As String
Private mSortCode As String, mIban As String, mSwift As String
Private mRollNo As String, mCheckDigit As String, mBranchId As String
Private mAddrCountry As String, mCity As String, mRegionId As String, mRegionName As String
Private mPostcode As String
Private mAddrLine(1 To MAX_ADDRESS_LINES) As String
Private mPhone As String, mPhoneIso3 As String, mPhoneCode As String, mPhoneDeviceId As String
Private mEmail As String
Private mRejectValue As String

' Supplier-group band rows, loaded once per run from the runtime "lookups" sheet.
Private mBandCount As Long
Private mBandRefId() As String
Private mBandFrom() As String
Private mBandTo() As String
Private mFallbackRefId As String

Public Sub RunApSupplierMaster(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    Dim sourceDataPath As String
    Dim targetEibPath As String
    Dim outputFilePath As String
    Dim selectionKey As String
    Dim sourceWorkbook As Workbook
    Dim outputWorkbook As Workbook
    Dim sourceTable As Variant
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim rowIndex As Long
    Dim rowsProcessed As Long
    Dim rowsSkipped As Long
    Dim rejectReason As String
    Dim issues As Collection
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

    On Error GoTo HandleFailure

    runStartTime = Now
    BeginRun runnerWorkbook, controlSheet, selectionKey

    sourceDataPath = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA)
    targetEibPath = ReadControlValue(controlSheet, CONTROL_LABEL_TARGET_EIB)
    outputFilePath = ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)
    EnsureNonBlank outputFilePath, CONTROL_LABEL_OUTPUT_FILE_PATH
    worksheetNameStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_NAME)
    worksheetResolutionStr = ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_WORKSHEET_RESOLUTION)
    sourceHeaderRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_HEADER_ROW)))
    sourceDataStartRow = CLng(Val(ReadControlValue(controlSheet, CONTROL_LABEL_SOURCE_DATA_START_ROW)))
    If StrComp(worksheetResolutionStr, "first_sheet", vbTextCompare) = 0 Then
        sheetToResolve = vbNullString
    Else
        sheetToResolve = worksheetNameStr
    End If

    Set sourceWorkbook = OpenWorkbookReadOnly(sourceDataPath)
    sourceTable = ReadWorksheetTable(ResolveSourceWorksheet(sourceWorkbook, sheetToResolve))
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing

    ResolveSourceColumns sourceTable, sourceHeaderRow
    LoadSupplierGroupBands runnerWorkbook

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, CLEAR_END_COLUMN
    targetRow = TARGET_DATA_START_ROW

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If Not RowIsBlankTemplateRow(sourceTable, rowIndex) Then
            rejectReason = ResolveSupplierRow(runnerWorkbook, sourceTable, rowIndex)
            If LenB(rejectReason) > 0 Then
                rowsSkipped = rowsSkipped + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, rejectReason, mRejectValue, mSupId & " " & mName
            Else
                rowsProcessed = rowsProcessed + 1
                targetRow = WriteSupplierBlock(targetSheet, targetRow, rowsProcessed)
            End If
        End If
    Next rowIndex

    If rowsSkipped > 0 Then
        issues.Add "Rejected " & CStr(rowsSkipped) & " supplier(s): missing required value or no Workday reference ID (see the .rejects.csv beside the output EIB)."
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
    Err.Raise failureNumber, "modApSupplierMaster.RunApSupplierMaster", failureMessage
End Sub

' --- Source columns -------------------------------------------------------------------------

Private Sub ResolveSourceColumns(ByVal sourceTable As Variant, ByVal headerRow As Long)
    mColSupId = FindColumnIndexInRow(sourceTable, headerRow, "Supplier ID")
    mColSupName = FindColumnIndexInRow(sourceTable, headerRow, "Supplier Name")
    mColReferenceId = FindColumnIndexInRow(sourceTable, headerRow, "Reference ID")
    mColCategory = FindColumnIndexInRow(sourceTable, headerRow, "Supplier Category")
    mColStatus = FindColumnIndexInRow(sourceTable, headerRow, "Supplier Status")
    mColPaymentTerms = FindColumnIndexInRow(sourceTable, headerRow, "Payment Terms")
    mColPaymentType = FindColumnIndexInRow(sourceTable, headerRow, "Default Payment Type")
    mColCurrency = FindColumnIndexInRow(sourceTable, headerRow, "Default Currency")
    mColTaxStatus = FindColumnIndexInRow(sourceTable, headerRow, "Transaction Tax Status")
    mColTaxStatusCountry = FindColumnIndexInRow(sourceTable, headerRow, "Tax Status Country Code")
    mColTaxId = FindColumnIndexInRow(sourceTable, headerRow, "Tax ID")
    mColTaxIdType = FindColumnIndexInRow(sourceTable, headerRow, "Tax ID Type")
    mColTaxIdCountry = FindColumnIndexInRow(sourceTable, headerRow, "Tax ID Country Code")
    mColPoIssue = FindColumnIndexInRow(sourceTable, headerRow, "Purchase Order Issue Option")
    mColPayMemo = FindColumnIndexInRow(sourceTable, headerRow, "Use Supplier Reference as Default Supplier Payment Memo")
    mColBankCountry = FindColumnIndexInRow(sourceTable, headerRow, "Bank Country Code")
    mColBankCurrency = FindColumnIndexInRow(sourceTable, headerRow, "Bank Currency")
    mColBankType = FindColumnIndexInRow(sourceTable, headerRow, "Bank Account Type")
    mColBankName = FindColumnIndexInRow(sourceTable, headerRow, "Bank Name")
    mColBankAcctName = FindColumnIndexInRow(sourceTable, headerRow, "Bank Account Name")
    mColBankAcctNo = FindColumnIndexInRow(sourceTable, headerRow, "Bank Account Number")
    mColSortCode = FindColumnIndexInRow(sourceTable, headerRow, "Sort Code/Routing Number")
    mColIban = FindColumnIndexInRow(sourceTable, headerRow, "IBAN")
    mColSwift = FindColumnIndexInRow(sourceTable, headerRow, "SWIFT/BIC")
    mColRollNo = FindColumnIndexInRow(sourceTable, headerRow, "Roll Number")
    mColCheckDigit = FindColumnIndexInRow(sourceTable, headerRow, "Check Digit")
    mColBranchId = FindColumnIndexInRow(sourceTable, headerRow, "Branch ID")
    mColAddrCountry = FindColumnIndexInRow(sourceTable, headerRow, "Address Country Code")
    mColAddrLine1 = FindColumnIndexInRow(sourceTable, headerRow, "Address Line 1")
    mColAddrLine2 = FindColumnIndexInRow(sourceTable, headerRow, "Address Line 2")
    mColAddrLine3 = FindColumnIndexInRow(sourceTable, headerRow, "Address Line 3")
    mColCity = FindColumnIndexInRow(sourceTable, headerRow, "City")
    mColRegion = FindColumnIndexInRow(sourceTable, headerRow, "Region")
    mColPostcode = FindColumnIndexInRow(sourceTable, headerRow, "Postcode")
    mColPhoneCountry = FindColumnIndexInRow(sourceTable, headerRow, "Phone Country Code")
    mColPhone = FindColumnIndexInRow(sourceTable, headerRow, "Phone")
    mColPhoneDevice = FindColumnIndexInRow(sourceTable, headerRow, "Phone Device Type")
    mColEmail = FindColumnIndexInRow(sourceTable, headerRow, "Email Address")
End Sub

Private Function CellText(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal columnIdx As Long) As String
    Dim rawValue As Variant

    If columnIdx < 1 Then Exit Function
    rawValue = tableData(rowIndex, columnIdx)
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    CellText = Trim$(Replace(CStr(rawValue), Chr$(160), " "))
End Function

' A template row with no Supplier ID, no Supplier Name and no Address Line 1 is an unused row
' (the template pre-fills defaults such as GB on row 2): skipped silently, not rejected.
Private Function RowIsBlankTemplateRow(ByVal tableData As Variant, ByVal rowIndex As Long) As Boolean
    RowIsBlankTemplateRow = _
        LenB(CellText(tableData, rowIndex, mColSupId)) = 0 And _
        LenB(CellText(tableData, rowIndex, mColSupName)) = 0 And _
        LenB(CellText(tableData, rowIndex, mColAddrLine1)) = 0
End Function

' --- Per-supplier resolution (any non-empty return value = reject the whole supplier) ---------

Private Function ResolveSupplierRow(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    ResetSupplierRecord
    ReadSupplierFields tableData, rowIndex

    ResolveSupplierRow = CheckRequiredFields()
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolveCoreReferences(runnerWorkbook, tableData, rowIndex)
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolveProcurement(runnerWorkbook, tableData, rowIndex)
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolveBank(runnerWorkbook)
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolveAddress(runnerWorkbook)
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolvePhone(runnerWorkbook, tableData, rowIndex)
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolveEmail()
    If LenB(ResolveSupplierRow) > 0 Then Exit Function
    ResolveSupplierRow = ResolveTaxId(runnerWorkbook, tableData, rowIndex)
End Function

Private Sub ResetSupplierRecord()
    Dim lineIndex As Long

    mSupId = vbNullString: mName = vbNullString: mRefId = vbNullString
    mCategoryId = vbNullString: mStatusId = vbNullString: mGroupId = vbNullString: mTermId = vbNullString
    mPayTypeId = vbNullString: mCurrencyId = vbNullString: mTaxStatusId = vbNullString: mTaxStatusCountry = vbNullString
    mTaxId = vbNullString: mTaxIdTypeId = vbNullString
    mPoIssueId = vbNullString: mPayMemo = vbNullString
    mBankCountry = vbNullString: mBankCurrency = vbNullString: mBankType = vbNullString: mBankTypeId = vbNullString
    mBankName = vbNullString: mBankAcctName = vbNullString: mBankAcctNo = vbNullString
    mSortCode = vbNullString: mIban = vbNullString: mSwift = vbNullString
    mRollNo = vbNullString: mCheckDigit = vbNullString: mBranchId = vbNullString
    mAddrCountry = vbNullString: mCity = vbNullString: mRegionId = vbNullString: mRegionName = vbNullString
    mPostcode = vbNullString
    For lineIndex = 1 To MAX_ADDRESS_LINES
        mAddrLine(lineIndex) = vbNullString
    Next lineIndex
    mPhone = vbNullString: mPhoneIso3 = vbNullString: mPhoneCode = vbNullString: mPhoneDeviceId = vbNullString
    mEmail = vbNullString
    mRejectValue = vbNullString
End Sub

Private Sub ReadSupplierFields(ByVal tableData As Variant, ByVal rowIndex As Long)
    mSupId = CellText(tableData, rowIndex, mColSupId)
    mName = CellText(tableData, rowIndex, mColSupName)
    mRefId = CellText(tableData, rowIndex, mColReferenceId)
    mAddrCountry = CellText(tableData, rowIndex, mColAddrCountry)
    mAddrLine(1) = CellText(tableData, rowIndex, mColAddrLine1)
    mAddrLine(2) = CellText(tableData, rowIndex, mColAddrLine2)
    mAddrLine(3) = CellText(tableData, rowIndex, mColAddrLine3)
    mCity = CellText(tableData, rowIndex, mColCity)
    mRegionName = CellText(tableData, rowIndex, mColRegion)
    mPostcode = CellText(tableData, rowIndex, mColPostcode)
    mPhone = CellText(tableData, rowIndex, mColPhone)
    mEmail = CellText(tableData, rowIndex, mColEmail)
    mTaxId = CellText(tableData, rowIndex, mColTaxId)
    mBankCountry = CellText(tableData, rowIndex, mColBankCountry)
    mBankCurrency = CellText(tableData, rowIndex, mColBankCurrency)
    mBankType = CellText(tableData, rowIndex, mColBankType)
    mBankName = CellText(tableData, rowIndex, mColBankName)
    mBankAcctName = CellText(tableData, rowIndex, mColBankAcctName)
    mBankAcctNo = CellText(tableData, rowIndex, mColBankAcctNo)
    mSortCode = CellText(tableData, rowIndex, mColSortCode)
    mIban = CellText(tableData, rowIndex, mColIban)
    mSwift = CellText(tableData, rowIndex, mColSwift)
    mRollNo = CellText(tableData, rowIndex, mColRollNo)
    mCheckDigit = CellText(tableData, rowIndex, mColCheckDigit)
    mBranchId = CellText(tableData, rowIndex, mColBranchId)
End Sub

Private Function CheckRequiredFields() As String
    If LenB(mSupId) = 0 Then
        CheckRequiredFields = "Missing required field: Supplier ID"
    ElseIf LenB(mName) = 0 Then
        CheckRequiredFields = "Missing required field: Supplier Name"
    ElseIf LenB(mAddrCountry) = 0 Then
        CheckRequiredFields = "Missing required field: Address Country Code"
    ElseIf LenB(mAddrLine(1)) = 0 And LenB(mAddrLine(2)) = 0 And LenB(mAddrLine(3)) = 0 Then
        CheckRequiredFields = "Missing required field: Address Line 1"
    ElseIf LenB(mCity) = 0 Then
        CheckRequiredFields = "Missing required field: City"
    ElseIf LenB(mBankAcctName) = 0 Then
        CheckRequiredFields = "Missing required field: Bank Account Name"
    ElseIf LenB(mBankAcctNo) = 0 Then
        CheckRequiredFields = "Missing required field: Bank Account Number"
    ElseIf LenB(mBankName) = 0 Then
        CheckRequiredFields = "Missing required field: Bank Name"
    End If
End Function

Private Function ResolveLookup( _
    ByVal runnerWorkbook As Workbook, _
    ByVal lookupName As String, _
    ByVal keyValue As String, _
    ByVal fieldName As String, _
    ByVal fieldLabel As String, _
    ByRef resolved As String _
) As String
    If LenB(keyValue) = 0 Then
        mRejectValue = vbNullString
        ResolveLookup = "Missing required field: " & fieldLabel
        Exit Function
    End If
    resolved = LookupRuntimeValue(runnerWorkbook, lookupName, keyValue, fieldName)
    If LenB(resolved) = 0 Then
        mRejectValue = keyValue
        ResolveLookup = fieldLabel & " '" & keyValue & "' has no Workday " & fieldName & " in lookup " & lookupName
    End If
End Function

' A currency code is valid as a DEFAULT / account currency only when ref_supplier_currency marks it
' default_allowed = Yes ('All Currencies' is valid for Accepted Currencies only). The code itself
' is what Workday takes.
Private Function ResolveCurrency(ByVal runnerWorkbook As Workbook, ByVal currencyCode As String, ByVal fieldLabel As String) As String
    Dim allowed As String

    ResolveCurrency = ResolveLookup(runnerWorkbook, LK_CURRENCY, currencyCode, "default_allowed", fieldLabel, allowed)
    If LenB(ResolveCurrency) > 0 Then Exit Function
    If StrComp(allowed, "Yes", vbTextCompare) <> 0 Then
        mRejectValue = currencyCode
        ResolveCurrency = fieldLabel & " '" & currencyCode & "' is not allowed as a default currency"
    End If
End Function

Private Function ResolveCoreReferences(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    Dim taxCountryName As String

    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_CATEGORY, CellText(tableData, rowIndex, mColCategory), "reference_id", "Supplier Category", mCategoryId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_STATUS, CellText(tableData, rowIndex, mColStatus), "reference_id", "Supplier Status", mStatusId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    mGroupId = ResolveSupplierGroupBand(mName)
    If LenB(mGroupId) = 0 Then
        mRejectValue = Left$(mName, 1)
        ResolveCoreReferences = "No supplier group band for the first character of Supplier Name"
        Exit Function
    End If
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_PAYMENT_TERM, CellText(tableData, rowIndex, mColPaymentTerms), "reference_id", "Payment Terms", mTermId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_PAYMENT_TYPE, CellText(tableData, rowIndex, mColPaymentType), "reference_id", "Default Payment Type", mPayTypeId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    mCurrencyId = CellText(tableData, rowIndex, mColCurrency)
    ResolveCoreReferences = ResolveCurrency(runnerWorkbook, mCurrencyId, "Default Currency")
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_TAX_STATUS, CellText(tableData, rowIndex, mColTaxStatus), "reference_id", "Transaction Tax Status", mTaxStatusId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    mTaxStatusCountry = CellText(tableData, rowIndex, mColTaxStatusCountry)
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_COUNTRY, mTaxStatusCountry, "country_name", "Tax Status Country Code", taxCountryName)
End Function

' Purchase Order Issue Option and the payment-memo flag are optional; when typed they must resolve.
Private Function ResolveProcurement(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    Dim poIssue As String
    Dim memoText As String

    poIssue = CellText(tableData, rowIndex, mColPoIssue)
    If LenB(poIssue) > 0 Then
        ResolveProcurement = ResolveLookup(runnerWorkbook, LK_PO_ISSUE, poIssue, "reference_id", "Purchase Order Issue Option", mPoIssueId)
        If LenB(ResolveProcurement) > 0 Then Exit Function
    End If
    memoText = CellText(tableData, rowIndex, mColPayMemo)
    If LenB(memoText) = 0 Then Exit Function
    If StrComp(memoText, "Yes", vbTextCompare) = 0 Then
        mPayMemo = OUTPUT_FLAG_YES
    ElseIf StrComp(memoText, "No", vbTextCompare) = 0 Then
        mPayMemo = OUTPUT_FLAG_NO
    Else
        mRejectValue = memoText
        ResolveProcurement = "Use Supplier Reference as Default Supplier Payment Memo must be Yes or No"
    End If
End Function

' The one settlement (bank) account. Country is written as typed (ISO2, as in the 3+ sample).
Private Function ResolveBank(ByVal runnerWorkbook As Workbook) As String
    Dim bankCountryName As String

    ResolveBank = ResolveLookup(runnerWorkbook, LK_COUNTRY, mBankCountry, "country_name", "Bank Country Code", bankCountryName)
    If LenB(ResolveBank) > 0 Then Exit Function
    ResolveBank = ResolveCurrency(runnerWorkbook, mBankCurrency, "Bank Currency")
    If LenB(ResolveBank) > 0 Then Exit Function
    ResolveBank = ResolveLookup(runnerWorkbook, LK_BANK_TYPE, mBankType, "reference_id", "Bank Account Type", mBankTypeId)
End Function

Private Function ResolveAddress(ByVal runnerWorkbook As Workbook) As String
    Dim countryName As String

    ResolveAddress = ResolveLookup(runnerWorkbook, LK_COUNTRY, mAddrCountry, "country_name", "Address Country Code", countryName)
    If LenB(ResolveAddress) > 0 Then Exit Function
    ' Region is optional; when typed it must resolve to a Workday region ID (key "<ISO2>|<region>").
    If LenB(mRegionName) > 0 Then
        ResolveAddress = ResolveLookup(runnerWorkbook, LK_REGION, mAddrCountry & "|" & mRegionName, "reference_id", "Region", mRegionId)
    End If
End Function

Private Function ResolvePhone(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    Dim phoneCountryCode As String

    If LenB(mPhone) = 0 Then Exit Function
    phoneCountryCode = CellText(tableData, rowIndex, mColPhoneCountry)
    ResolvePhone = ResolveLookup(runnerWorkbook, LK_COUNTRY, phoneCountryCode, "iso3", "Phone Country Code", mPhoneIso3)
    If LenB(ResolvePhone) > 0 Then Exit Function
    ResolvePhone = ResolveLookup(runnerWorkbook, LK_COUNTRY, phoneCountryCode, "phone_code", "Phone Country Code", mPhoneCode)
    If LenB(ResolvePhone) > 0 Then Exit Function
    mPhoneCode = Replace(mPhoneCode, "+", vbNullString)
    ResolvePhone = ResolveLookup(runnerWorkbook, LK_PHONE_DEVICE, CellText(tableData, rowIndex, mColPhoneDevice), "reference_id", "Phone Device Type", mPhoneDeviceId)
End Function

' 1:1 rule: exactly one email address per supplier.
Private Function ResolveEmail() As String
    If InStr(1, mEmail, ";", vbBinaryCompare) > 0 Then
        mRejectValue = mEmail
        ResolveEmail = "More than one email address (1:1 rule allows a single Email Address)"
    End If
End Function

' The tax-identifier group is all-or-nothing: Tax ID, Tax ID Type and Tax ID Country Code.
Private Function ResolveTaxId(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    Dim taxIdType As String
    Dim taxIdCountry As String

    taxIdType = CellText(tableData, rowIndex, mColTaxIdType)
    taxIdCountry = CellText(tableData, rowIndex, mColTaxIdCountry)
    If LenB(mTaxId) = 0 And LenB(taxIdType) = 0 Then Exit Function
    If LenB(mTaxId) = 0 Then
        ResolveTaxId = "Tax ID Type given without a Tax ID"
        Exit Function
    End If
    If LenB(taxIdType) = 0 Or LenB(taxIdCountry) = 0 Then
        ResolveTaxId = "Tax ID needs both Tax ID Type and Tax ID Country Code"
        Exit Function
    End If
    ResolveTaxId = ResolveLookup(runnerWorkbook, LK_TAX_ID_TYPE, taxIdCountry & "|" & taxIdType, "reference_id", "Tax ID Type", mTaxIdTypeId)
End Function

' --- Supplier-group band (assignment = auto_band rows of ref_supplier_group) -----------------

Private Sub LoadSupplierGroupBands(ByVal runnerWorkbook As Workbook)
    Dim lookupSheet As Worksheet
    Dim lookupValues As Variant
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim groupKey As String

    mBandCount = 0
    mFallbackRefId = vbNullString
    ReDim mBandRefId(1 To 8)
    ReDim mBandFrom(1 To 8)
    ReDim mBandTo(1 To 8)

    Set lookupSheet = runnerWorkbook.Worksheets(LOOKUP_SHEET_NAME)
    lastRow = lookupSheet.Cells(lookupSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    lookupValues = lookupSheet.Range(lookupSheet.Cells(2, 1), lookupSheet.Cells(lastRow, 4)).Value

    For rowIndex = 1 To UBound(lookupValues, 1)
        If StrComp(Trim$(CStr(lookupValues(rowIndex, 1))), LK_GROUP, vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(lookupValues(rowIndex, 3))), "assignment", vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(lookupValues(rowIndex, 4))), GROUP_ASSIGNMENT_AUTO_BAND, vbTextCompare) = 0 Then
            groupKey = Trim$(CStr(lookupValues(rowIndex, 2)))
            AddBand runnerWorkbook, groupKey
        End If
    Next rowIndex
End Sub

Private Sub AddBand(ByVal runnerWorkbook As Workbook, ByVal groupKey As String)
    Dim refId As String

    refId = LookupRuntimeValue(runnerWorkbook, LK_GROUP, groupKey, "reference_id")
    If LenB(refId) = 0 Then Exit Sub
    If StrComp(LookupRuntimeValue(runnerWorkbook, LK_GROUP, groupKey, "is_fallback"), "Y", vbTextCompare) = 0 Then
        mFallbackRefId = refId
    End If
    If LenB(LookupRuntimeValue(runnerWorkbook, LK_GROUP, groupKey, "first_char_from")) = 0 Then Exit Sub
    If mBandCount >= UBound(mBandRefId) Then Exit Sub
    mBandCount = mBandCount + 1
    mBandRefId(mBandCount) = refId
    mBandFrom(mBandCount) = UCase$(LookupRuntimeValue(runnerWorkbook, LK_GROUP, groupKey, "first_char_from"))
    mBandTo(mBandCount) = UCase$(LookupRuntimeValue(runnerWorkbook, LK_GROUP, groupKey, "first_char_to"))
End Sub

' The band whose first_char_from..first_char_to holds Supplier Name's first character; the
' is_fallback row when none does. Empty (reject) when neither exists. Band boundaries are config
' (config/ingest_lookups/ref_supplier_group.json), never code.
Private Function ResolveSupplierGroupBand(ByVal supplierName As String) As String
    Dim firstChar As String
    Dim bandIndex As Long

    firstChar = UCase$(Left$(supplierName, 1))
    For bandIndex = 1 To mBandCount
        If firstChar >= mBandFrom(bandIndex) And firstChar <= mBandTo(bandIndex) Then
            ResolveSupplierGroupBand = mBandRefId(bandIndex)
            Exit Function
        End If
    Next bandIndex
    ResolveSupplierGroupBand = mFallbackRefId
End Function

' --- Writing --------------------------------------------------------------------------------

' Writes one supplier as a block of rows and returns the next free row. Address lines (the only
' multi-row item) go through modEntityBlockWriter; every other field is written on the first row.
Private Function WriteSupplierBlock(ByVal targetSheet As Worksheet, ByVal startRow As Long, ByVal spreadsheetKey As Long) As Long
    Dim lineIndex As Long
    Dim itemIndex As Long

    EntityBlockReset
    itemIndex = 0
    For lineIndex = 1 To MAX_ADDRESS_LINES
        If LenB(mAddrLine(lineIndex)) > 0 Then
            itemIndex = itemIndex + 1
            EntityBlockAddItem "DM", itemIndex, CStr(itemIndex)
            EntityBlockAddItem "DO", itemIndex, ADDRESS_LINE_TYPE_PREFIX & CStr(itemIndex)
            EntityBlockAddItem "DP", itemIndex, mAddrLine(lineIndex)
        End If
    Next lineIndex
    EntityBlockAddAncestor "DD", OUTPUT_ROW_ID, "DM,DO,DP"
    WriteSupplierBlock = EntityBlockFlush(targetSheet, startRow, spreadsheetKey)

    WriteIdentityFields targetSheet, startRow
    WriteSettingsFields targetSheet, startRow
    WriteAddressFields targetSheet, startRow
    WritePhoneFields targetSheet, startRow
    WriteEmailFields targetSheet, startRow
    WriteTaxFields targetSheet, startRow
    WriteBankFields targetSheet, startRow
End Function

Private Sub PutValue(ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal columnIdx As Long, ByVal cellValue As String)
    If LenB(cellValue) = 0 Then Exit Sub
    targetSheet.Cells(targetRow, columnIdx).Value = cellValue
End Sub

' Text format first, so IDs / postcodes / phone numbers / account numbers keep leading zeros and are not re-typed by Excel.
Private Sub PutText(ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal columnIdx As Long, ByVal cellValue As String)
    If LenB(cellValue) = 0 Then Exit Sub
    targetSheet.Cells(targetRow, columnIdx).NumberFormat = "@"
    targetSheet.Cells(targetRow, columnIdx).Value = cellValue
End Sub

Private Sub WriteIdentityFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("C"), OUTPUT_ADD_ONLY
    PutText targetSheet, targetRow, ColumnIndex("H"), mSupId
    PutText targetSheet, targetRow, ColumnIndex("I"), mRefId
    PutValue targetSheet, targetRow, ColumnIndex("K"), mName
    PutValue targetSheet, targetRow, ColumnIndex("Q"), OUTPUT_APPROVAL_STATUS
    PutValue targetSheet, targetRow, ColumnIndex("R"), OUTPUT_SUBMIT
    PutValue targetSheet, targetRow, ColumnIndex("CZ"), mName
    ' Supplier Status: the 3+ sample writes ACTIVE (upper case) where the tenant range's Ref ID is
    ' "Active"; the sample is the authority for the output (first-Workday-upload check in docs/TODO.md).
    PutValue targetSheet, targetRow, ColumnIndex("MH"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("MI"), UCase$(mStatusId)
End Sub

Private Sub WriteSettingsFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("AU"), mCategoryId
    PutValue targetSheet, targetRow, ColumnIndex("AV"), mGroupId
    PutValue targetSheet, targetRow, ColumnIndex("BW"), mPoIssueId
    PutValue targetSheet, targetRow, ColumnIndex("CR"), mTermId
    PutValue targetSheet, targetRow, ColumnIndex("CT"), mPayTypeId
    PutValue targetSheet, targetRow, ColumnIndex("CU"), mPayTypeId
    PutValue targetSheet, targetRow, ColumnIndex("MQ"), mPayMemo
    PutValue targetSheet, targetRow, ColumnIndex("MT"), mCurrencyId
    PutValue targetSheet, targetRow, ColumnIndex("AP"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("AR"), mTaxStatusCountry
    PutValue targetSheet, targetRow, ColumnIndex("AS"), mTaxStatusId
End Sub

Private Sub WriteAddressFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("DD"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("DF"), OUTPUT_ADDRESS_FORMAT
    PutText targetSheet, targetRow, ColumnIndex("DJ"), OUTPUT_EFFECTIVE_DATE
    PutValue targetSheet, targetRow, ColumnIndex("DK"), mAddrCountry
    PutValue targetSheet, targetRow, ColumnIndex("DQ"), mCity
    PutValue targetSheet, targetRow, ColumnIndex("DS"), mAddrCountry
    PutValue targetSheet, targetRow, ColumnIndex("DX"), mRegionId
    PutValue targetSheet, targetRow, ColumnIndex("DY"), mRegionName
    PutText targetSheet, targetRow, ColumnIndex("ED"), mPostcode
    PutValue targetSheet, targetRow, ColumnIndex("EE"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("EF"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("EG"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("EH"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("EI"), OUTPUT_USAGE_TYPE
    PutValue targetSheet, targetRow, ColumnIndex("EJ"), OUTPUT_USE_FOR
    PutText targetSheet, targetRow, ColumnIndex("EP"), mSupId & ADDRESS_ID_SUFFIX
End Sub

Private Sub WritePhoneFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    If LenB(mPhone) = 0 Then Exit Sub
    PutValue targetSheet, targetRow, ColumnIndex("EQ"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("FA"), mPhoneIso3
    PutValue targetSheet, targetRow, ColumnIndex("FB"), mPhoneCode
    PutText targetSheet, targetRow, ColumnIndex("FC"), mPhone
    PutValue targetSheet, targetRow, ColumnIndex("FE"), mPhoneDeviceId
    PutValue targetSheet, targetRow, ColumnIndex("FF"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("FG"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("FH"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("FI"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("FJ"), OUTPUT_USAGE_TYPE
    PutValue targetSheet, targetRow, ColumnIndex("FK"), OUTPUT_USE_FOR
End Sub

' One email: it is both the Default PO Issue Email (BZ / CC) and the supplier email (FP.. / GD).
Private Sub WriteEmailFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    If LenB(mEmail) = 0 Then Exit Sub
    PutValue targetSheet, targetRow, ColumnIndex("BZ"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("CC"), mEmail
    PutValue targetSheet, targetRow, ColumnIndex("FP"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("FS"), mEmail
    PutValue targetSheet, targetRow, ColumnIndex("FU"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("FV"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("FW"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("FX"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("FY"), OUTPUT_USAGE_TYPE
    PutValue targetSheet, targetRow, ColumnIndex("FZ"), OUTPUT_USE_FOR
    PutText targetSheet, targetRow, ColumnIndex("GD"), mSupId & EMAIL_ID_SUFFIX
End Sub

' Tax-identifier group: only when a Tax ID was typed (ResolveTaxId guarantees type + country resolved).
' AE / AF are fixed "Y" (Transaction Tax ID / Primary Tax ID), as in modArCustomerMaster.
Private Sub WriteTaxFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    If LenB(mTaxId) = 0 Then Exit Sub
    PutValue targetSheet, targetRow, ColumnIndex("AA"), OUTPUT_ROW_ID
    PutText targetSheet, targetRow, ColumnIndex("AC"), mTaxId
    PutValue targetSheet, targetRow, ColumnIndex("AD"), mTaxIdTypeId
    PutValue targetSheet, targetRow, ColumnIndex("AE"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("AF"), OUTPUT_FLAG_YES
End Sub

Private Sub WriteBankFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("HM"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("HP"), mBankCountry
    PutValue targetSheet, targetRow, ColumnIndex("HQ"), mBankCurrency
    PutValue targetSheet, targetRow, ColumnIndex("HS"), mBankTypeId
    PutValue targetSheet, targetRow, ColumnIndex("HT"), mBankName
    PutText targetSheet, targetRow, ColumnIndex("HU"), mSortCode
    PutText targetSheet, targetRow, ColumnIndex("HV"), mBranchId
    PutText targetSheet, targetRow, ColumnIndex("HX"), mBankAcctNo
    PutText targetSheet, targetRow, ColumnIndex("HY"), mCheckDigit
    PutValue targetSheet, targetRow, ColumnIndex("HZ"), mBankAcctName
    PutText targetSheet, targetRow, ColumnIndex("IA"), mRollNo
    PutText targetSheet, targetRow, ColumnIndex("IB"), mIban
    PutText targetSheet, targetRow, ColumnIndex("IC"), mSwift
End Sub
