Attribute VB_Name = "modArCustomerMaster"
Option Explicit

' AR Customer input template (sheet "Customer") -> Workday "Submit Customer" master EIB (3+ Master
' Customers workbook, data from row 6). One input row = one customer = one EIB block of rows that
' share the Spreadsheet Key (column B).
'
' 1:1 RULE (operator, 2026-10-12): a single value per area -- Use For = BILLING only, one customer
' group (the automatic alphabetical band), one email, one phone, every Row ID = 1. The ONE exception
' (operator, 2026-10-12) is the address lines: each NON-BLANK Address Line 1..3 is written on its own
' EIB row, exactly like the 3+ sample (Master Suppliers rows 6-7: ADDRESS_LINE_1 then ADDRESS_LINE_2).
' Row k of the block carries B (same key), BL (address Row ID = 1, repeated), BU (= k), BW
' (ADDRESS_LINE_k) and BX (the line text); every other field stays on the first row only. Blank lines
' are skipped and k is renumbered contiguously from 1. The block is written by the shared
' modEntityBlockWriter (registry concern entity_row_block); the letters it is given are pinned in
' config/transforms/ar_customer_master.yaml customer_master_eib_columns.
'
' NO FALLBACKS: a missing required value, or a value with no Workday reference ID in the runtime
' "lookups" sheet, rejects the WHOLE customer (one AppendRejectRow, no EIB rows).
' Status: code-complete, pending manual Excel import + verify (not run in Excel yet).

Private Const TARGET_SHEET_NAME As String = "Submit Customer"
Private Const TARGET_DATA_START_ROW As Long = 6
Private Const CLEAR_END_COLUMN As String = "NT"
Private Const OUTPUT_ADD_ONLY As String = "Y"
Private Const OUTPUT_SUBMIT As String = "Y"
Private Const OUTPUT_FLAG_YES As String = "Y"
Private Const OUTPUT_ROW_ID As String = "1"
Private Const OUTPUT_USAGE_TYPE As String = "BUSINESS"
Private Const OUTPUT_USE_FOR As String = "BILLING"
Private Const ADDRESS_LINE_TYPE_PREFIX As String = "ADDRESS_LINE_"
Private Const MAX_ADDRESS_LINES As Long = 3
Private Const ADDRESS_ID_SUFFIX As String = "_MAIN_1"
Private Const PHONE_ID_SUFFIX As String = "_PH1_1"
Private Const EMAIL_ID_SUFFIX As String = "_EM1_1"
Private Const GROUP_ASSIGNMENT_AUTO_BAND As String = "auto_band"

Private Const LK_CATEGORY As String = "ref_customer_category"
Private Const LK_GROUP As String = "ref_customer_group"
Private Const LK_PAYMENT_TERM As String = "ref_payment_term"
Private Const LK_PAYMENT_TYPE As String = "ref_payment_type"
Private Const LK_CURRENCY As String = "ref_currency_default"
Private Const LK_TAX_STATUS As String = "ref_transaction_tax_status"
Private Const LK_COUNTRY As String = "ref_country"
Private Const LK_REGION As String = "ref_country_region"
Private Const LK_TAX_ID_TYPE As String = "ref_country_tax_id_type"
Private Const LK_PHONE_DEVICE As String = "ref_phone_device"
Private Const LK_DELIVERY As String = "ref_document_delivery_method"
Private Const LK_RECIPIENT_INVOICE As String = "ref_document_email_recipient_invoice"
Private Const LK_RECIPIENT_STATEMENT As String = "ref_document_email_recipient_statement"

' Source (input template) column positions, resolved once per run from the header row.
Private mColCustomerId As Long, mColCustomerName As Long, mColReferenceId As Long
Private mColCategory As Long, mColPaymentTerms As Long, mColPaymentType As Long
Private mColCurrency As Long, mColTaxStatus As Long, mColTaxStatusCountry As Long
Private mColTaxId As Long, mColTaxIdType As Long, mColTaxIdCountry As Long
Private mColAddrCountry As Long, mColAddrLine1 As Long, mColAddrLine2 As Long, mColAddrLine3 As Long
Private mColCity As Long, mColRegion As Long, mColPostcode As Long
Private mColPhoneCountry As Long, mColPhone As Long, mColPhoneDevice As Long, mColEmail As Long
Private mColInvDelivery As Long, mColInvRecipients As Long
Private mColStmtDelivery As Long, mColStmtRecipients As Long

' The customer currently being resolved (module level so the stages below stay short).
Private mCustId As String, mName As String, mRefId As String
Private mCategoryId As String, mGroupId As String, mTermId As String, mPayTypeId As String
Private mCurrencyId As String, mTaxStatusId As String, mTaxStatusCountry As String
Private mTaxId As String, mTaxIdTypeId As String
Private mAddrCountry As String, mCity As String, mRegionId As String, mRegionName As String
Private mPostcode As String
Private mAddrLine(1 To MAX_ADDRESS_LINES) As String
Private mPhone As String, mPhoneIso3 As String, mPhoneCode As String, mPhoneDeviceId As String
Private mEmail As String
Private mInvDeliveryId As String, mInvRecipientsId As String
Private mStmtDeliveryId As String, mStmtRecipientsId As String
Private mRejectValue As String

' Customer-group band rows, loaded once per run from the runtime "lookups" sheet.
Private mBandCount As Long
Private mBandRefId() As String
Private mBandFrom() As String
Private mBandTo() As String
Private mFallbackRefId As String

Public Sub RunArCustomerMaster(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
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
    LoadCustomerGroupBands runnerWorkbook

    Set outputWorkbook = CloneTemplateWorkbook(targetEibPath, outputFilePath)
    Set targetSheet = outputWorkbook.Worksheets(TARGET_SHEET_NAME)
    ClearPublishedRows targetSheet, TARGET_DATA_START_ROW, CLEAR_END_COLUMN
    targetRow = TARGET_DATA_START_ROW

    For rowIndex = sourceDataStartRow To UBound(sourceTable, 1)
        If (rowIndex Mod 200) = 0 Then UpdateRunProgress runnerWorkbook, selectionKey, rowIndex, UBound(sourceTable, 1), runStartTime
        If Not RowIsBlankTemplateRow(sourceTable, rowIndex) Then
            rejectReason = ResolveCustomerRow(runnerWorkbook, sourceTable, rowIndex)
            If LenB(rejectReason) > 0 Then
                rowsSkipped = rowsSkipped + 1
                AppendRejectRow runnerWorkbook, selectionKey, rowIndex, rejectReason, mRejectValue, mCustId & " " & mName
            Else
                rowsProcessed = rowsProcessed + 1
                targetRow = WriteCustomerBlock(targetSheet, targetRow, rowsProcessed)
            End If
        End If
    Next rowIndex

    If rowsSkipped > 0 Then
        issues.Add "Rejected " & CStr(rowsSkipped) & " customer(s): missing required value or no Workday reference ID (see the .rejects.csv beside the output EIB)."
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
    Err.Raise failureNumber, "modArCustomerMaster.RunArCustomerMaster", failureMessage
End Sub

' --- Source columns -------------------------------------------------------------------------

Private Sub ResolveSourceColumns(ByVal sourceTable As Variant, ByVal headerRow As Long)
    mColCustomerId = FindColumnIndexInRow(sourceTable, headerRow, "Customer ID")
    mColCustomerName = FindColumnIndexInRow(sourceTable, headerRow, "Customer Name")
    mColReferenceId = FindColumnIndexInRow(sourceTable, headerRow, "Reference ID")
    mColCategory = FindColumnIndexInRow(sourceTable, headerRow, "Customer Category")
    mColPaymentTerms = FindColumnIndexInRow(sourceTable, headerRow, "Payment Terms")
    mColPaymentType = FindColumnIndexInRow(sourceTable, headerRow, "Default Payment Type")
    mColCurrency = FindColumnIndexInRow(sourceTable, headerRow, "Default Currency")
    mColTaxStatus = FindColumnIndexInRow(sourceTable, headerRow, "Transaction Tax Status")
    mColTaxStatusCountry = FindColumnIndexInRow(sourceTable, headerRow, "Tax Status Country Code")
    mColTaxId = FindColumnIndexInRow(sourceTable, headerRow, "Tax ID")
    mColTaxIdType = FindColumnIndexInRow(sourceTable, headerRow, "Tax ID Type")
    mColTaxIdCountry = FindColumnIndexInRow(sourceTable, headerRow, "Tax ID Country Code")
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
    mColInvDelivery = FindColumnIndexInRow(sourceTable, headerRow, "Invoice Delivery Method")
    mColInvRecipients = FindColumnIndexInRow(sourceTable, headerRow, "Invoice Notification Email Recipients")
    mColStmtDelivery = FindColumnIndexInRow(sourceTable, headerRow, "Statement Delivery Method")
    mColStmtRecipients = FindColumnIndexInRow(sourceTable, headerRow, "Statement Notification Email Recipients")
End Sub

Private Function CellText(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal columnIdx As Long) As String
    Dim rawValue As Variant

    If columnIdx < 1 Then Exit Function
    rawValue = tableData(rowIndex, columnIdx)
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    CellText = Trim$(Replace(CStr(rawValue), Chr$(160), " "))
End Function

' A template row with no Customer ID, no Customer Name and no Address Line 1 is an unused row
' (the template pre-fills defaults such as GB on row 2): skipped silently, not rejected.
Private Function RowIsBlankTemplateRow(ByVal tableData As Variant, ByVal rowIndex As Long) As Boolean
    RowIsBlankTemplateRow = _
        LenB(CellText(tableData, rowIndex, mColCustomerId)) = 0 And _
        LenB(CellText(tableData, rowIndex, mColCustomerName)) = 0 And _
        LenB(CellText(tableData, rowIndex, mColAddrLine1)) = 0
End Function

' --- Per-customer resolution (any non-empty return value = reject the whole customer) --------

Private Function ResolveCustomerRow(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    ResetCustomerRecord
    ReadCustomerFields tableData, rowIndex

    ResolveCustomerRow = CheckRequiredFields(tableData, rowIndex)
    If LenB(ResolveCustomerRow) > 0 Then Exit Function
    ResolveCustomerRow = ResolveCoreReferences(runnerWorkbook, tableData, rowIndex)
    If LenB(ResolveCustomerRow) > 0 Then Exit Function
    ResolveCustomerRow = ResolveAddress(runnerWorkbook, tableData, rowIndex)
    If LenB(ResolveCustomerRow) > 0 Then Exit Function
    ResolveCustomerRow = ResolvePhone(runnerWorkbook, tableData, rowIndex)
    If LenB(ResolveCustomerRow) > 0 Then Exit Function
    ResolveCustomerRow = ResolveEmail(tableData, rowIndex)
    If LenB(ResolveCustomerRow) > 0 Then Exit Function
    ResolveCustomerRow = ResolveTaxId(runnerWorkbook, tableData, rowIndex)
End Function

Private Sub ResetCustomerRecord()
    Dim lineIndex As Long

    mCustId = vbNullString: mName = vbNullString: mRefId = vbNullString
    mCategoryId = vbNullString: mGroupId = vbNullString: mTermId = vbNullString: mPayTypeId = vbNullString
    mCurrencyId = vbNullString: mTaxStatusId = vbNullString: mTaxStatusCountry = vbNullString
    mTaxId = vbNullString: mTaxIdTypeId = vbNullString
    mAddrCountry = vbNullString: mCity = vbNullString: mRegionId = vbNullString: mRegionName = vbNullString
    mPostcode = vbNullString
    For lineIndex = 1 To MAX_ADDRESS_LINES
        mAddrLine(lineIndex) = vbNullString
    Next lineIndex
    mPhone = vbNullString: mPhoneIso3 = vbNullString: mPhoneCode = vbNullString: mPhoneDeviceId = vbNullString
    mEmail = vbNullString
    mInvDeliveryId = vbNullString: mInvRecipientsId = vbNullString
    mStmtDeliveryId = vbNullString: mStmtRecipientsId = vbNullString
    mRejectValue = vbNullString
End Sub

Private Sub ReadCustomerFields(ByVal tableData As Variant, ByVal rowIndex As Long)
    mCustId = CellText(tableData, rowIndex, mColCustomerId)
    mName = CellText(tableData, rowIndex, mColCustomerName)
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
End Sub

Private Function CheckRequiredFields(ByVal tableData As Variant, ByVal rowIndex As Long) As String
    If LenB(mCustId) = 0 Then
        CheckRequiredFields = "Missing required field: Customer ID"
    ElseIf LenB(mName) = 0 Then
        CheckRequiredFields = "Missing required field: Customer Name"
    ElseIf LenB(mAddrCountry) = 0 Then
        CheckRequiredFields = "Missing required field: Address Country Code"
    ElseIf LenB(mAddrLine(1)) = 0 And LenB(mAddrLine(2)) = 0 And LenB(mAddrLine(3)) = 0 Then
        CheckRequiredFields = "Missing required field: Address Line 1"
    ElseIf LenB(mCity) = 0 Then
        CheckRequiredFields = "Missing required field: City"
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

Private Function ResolveCoreReferences(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    Dim taxCountryName As String

    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_CATEGORY, CellText(tableData, rowIndex, mColCategory), "reference_id", "Customer Category", mCategoryId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    mGroupId = ResolveCustomerGroupBand(mName)
    If LenB(mGroupId) = 0 Then
        mRejectValue = Left$(mName, 1)
        ResolveCoreReferences = "No customer group band for the first character of Customer Name"
        Exit Function
    End If
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_PAYMENT_TERM, CellText(tableData, rowIndex, mColPaymentTerms), "reference_id", "Payment Terms", mTermId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_PAYMENT_TYPE, CellText(tableData, rowIndex, mColPaymentType), "reference_id", "Default Payment Type", mPayTypeId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_CURRENCY, CellText(tableData, rowIndex, mColCurrency), "currency_id", "Default Currency", mCurrencyId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_TAX_STATUS, CellText(tableData, rowIndex, mColTaxStatus), "reference_id", "Transaction Tax Status", mTaxStatusId)
    If LenB(ResolveCoreReferences) > 0 Then Exit Function
    mTaxStatusCountry = CellText(tableData, rowIndex, mColTaxStatusCountry)
    ResolveCoreReferences = ResolveLookup(runnerWorkbook, LK_COUNTRY, mTaxStatusCountry, "country_name", "Tax Status Country Code", taxCountryName)
    If LenB(ResolveCoreReferences) > 0 Then
        mRejectValue = mTaxStatusCountry
        Exit Function
    End If
    ResolveCoreReferences = ResolveDeliverySettings(runnerWorkbook, tableData, rowIndex)
End Function

Private Function ResolveDeliverySettings(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
    ResolveDeliverySettings = ResolveLookup(runnerWorkbook, LK_DELIVERY, CellText(tableData, rowIndex, mColInvDelivery), "reference_id", "Invoice Delivery Method", mInvDeliveryId)
    If LenB(ResolveDeliverySettings) > 0 Then Exit Function
    ResolveDeliverySettings = ResolveLookup(runnerWorkbook, LK_RECIPIENT_INVOICE, CellText(tableData, rowIndex, mColInvRecipients), "reference_id", "Invoice Notification Email Recipients", mInvRecipientsId)
    If LenB(ResolveDeliverySettings) > 0 Then Exit Function
    ResolveDeliverySettings = ResolveLookup(runnerWorkbook, LK_DELIVERY, CellText(tableData, rowIndex, mColStmtDelivery), "reference_id", "Statement Delivery Method", mStmtDeliveryId)
    If LenB(ResolveDeliverySettings) > 0 Then Exit Function
    ResolveDeliverySettings = ResolveLookup(runnerWorkbook, LK_RECIPIENT_STATEMENT, CellText(tableData, rowIndex, mColStmtRecipients), "reference_id", "Statement Notification Email Recipients", mStmtRecipientsId)
End Function

Private Function ResolveAddress(ByVal runnerWorkbook As Workbook, ByVal tableData As Variant, ByVal rowIndex As Long) As String
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

' 1:1 rule: exactly one email address per customer.
Private Function ResolveEmail(ByVal tableData As Variant, ByVal rowIndex As Long) As String
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

' --- Customer-group band (assignment = auto_band rows of ref_customer_group) -----------------

Private Sub LoadCustomerGroupBands(ByVal runnerWorkbook As Workbook)
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

' The band whose first_char_from..first_char_to holds Customer Name's first character; the
' is_fallback row when none does. Empty (reject) when neither exists. Band boundaries are config
' (config/ingest_lookups/ref_customer_group.json), never code.
Private Function ResolveCustomerGroupBand(ByVal customerName As String) As String
    Dim firstChar As String
    Dim bandIndex As Long

    firstChar = UCase$(Left$(customerName, 1))
    For bandIndex = 1 To mBandCount
        If firstChar >= mBandFrom(bandIndex) And firstChar <= mBandTo(bandIndex) Then
            ResolveCustomerGroupBand = mBandRefId(bandIndex)
            Exit Function
        End If
    Next bandIndex
    ResolveCustomerGroupBand = mFallbackRefId
End Function

' --- Writing --------------------------------------------------------------------------------

' Writes one customer as a block of rows and returns the next free row. Address lines (the only
' multi-row item) go through modEntityBlockWriter; every other field is written on the first row.
Private Function WriteCustomerBlock(ByVal targetSheet As Worksheet, ByVal startRow As Long, ByVal spreadsheetKey As Long) As Long
    Dim lineIndex As Long
    Dim itemIndex As Long

    EntityBlockReset
    itemIndex = 0
    For lineIndex = 1 To MAX_ADDRESS_LINES
        If LenB(mAddrLine(lineIndex)) > 0 Then
            itemIndex = itemIndex + 1
            EntityBlockAddItem "BU", itemIndex, CStr(itemIndex)
            EntityBlockAddItem "BW", itemIndex, ADDRESS_LINE_TYPE_PREFIX & CStr(itemIndex)
            EntityBlockAddItem "BX", itemIndex, mAddrLine(lineIndex)
        End If
    Next lineIndex
    EntityBlockAddAncestor "BL", OUTPUT_ROW_ID, "BU,BW,BX"
    WriteCustomerBlock = EntityBlockFlush(targetSheet, startRow, spreadsheetKey)

    WriteIdentityFields targetSheet, startRow
    WriteAddressFields targetSheet, startRow
    WritePhoneFields targetSheet, startRow
    WriteEmailFields targetSheet, startRow
    WriteTaxFields targetSheet, startRow
    WriteSettingsFields targetSheet, startRow
End Function

Private Sub PutValue(ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal columnIdx As Long, ByVal cellValue As String)
    If LenB(cellValue) = 0 Then Exit Sub
    targetSheet.Cells(targetRow, columnIdx).Value = cellValue
End Sub

' Text format first, so IDs / postcodes / phone numbers keep leading zeros and are not re-typed by Excel.
Private Sub PutText(ByVal targetSheet As Worksheet, ByVal targetRow As Long, ByVal columnIdx As Long, ByVal cellValue As String)
    If LenB(cellValue) = 0 Then Exit Sub
    targetSheet.Cells(targetRow, columnIdx).NumberFormat = "@"
    targetSheet.Cells(targetRow, columnIdx).Value = cellValue
End Sub

Private Sub WriteIdentityFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("C"), OUTPUT_ADD_ONLY
    PutText targetSheet, targetRow, ColumnIndex("H"), mCustId
    PutText targetSheet, targetRow, ColumnIndex("I"), mRefId
    PutValue targetSheet, targetRow, ColumnIndex("J"), mName
    PutValue targetSheet, targetRow, ColumnIndex("L"), OUTPUT_SUBMIT
    PutValue targetSheet, targetRow, ColumnIndex("BI"), mName
End Sub

Private Sub WriteAddressFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("BL"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("BS"), mAddrCountry
    PutValue targetSheet, targetRow, ColumnIndex("BY"), mCity
    PutValue targetSheet, targetRow, ColumnIndex("CA"), mAddrCountry
    PutValue targetSheet, targetRow, ColumnIndex("CF"), mRegionId
    PutValue targetSheet, targetRow, ColumnIndex("CG"), mRegionName
    PutText targetSheet, targetRow, ColumnIndex("CL"), mPostcode
    PutValue targetSheet, targetRow, ColumnIndex("CM"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("CN"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("CO"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("CP"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("CQ"), OUTPUT_USAGE_TYPE
    PutValue targetSheet, targetRow, ColumnIndex("CR"), OUTPUT_USE_FOR
    PutText targetSheet, targetRow, ColumnIndex("CX"), mCustId & ADDRESS_ID_SUFFIX
End Sub

Private Sub WritePhoneFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    If LenB(mPhone) = 0 Then Exit Sub
    PutValue targetSheet, targetRow, ColumnIndex("CY"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("DI"), mPhoneIso3
    PutValue targetSheet, targetRow, ColumnIndex("DJ"), mPhoneCode
    PutText targetSheet, targetRow, ColumnIndex("DK"), mPhone
    PutValue targetSheet, targetRow, ColumnIndex("DM"), mPhoneDeviceId
    PutValue targetSheet, targetRow, ColumnIndex("DN"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("DO"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("DP"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("DQ"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("DR"), OUTPUT_USAGE_TYPE
    PutValue targetSheet, targetRow, ColumnIndex("DS"), OUTPUT_USE_FOR
    PutText targetSheet, targetRow, ColumnIndex("DW"), mCustId & PHONE_ID_SUFFIX
End Sub

Private Sub WriteEmailFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    If LenB(mEmail) = 0 Then Exit Sub
    PutValue targetSheet, targetRow, ColumnIndex("DX"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("EA"), mEmail
    PutValue targetSheet, targetRow, ColumnIndex("EC"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("ED"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("EE"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("EF"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("EG"), OUTPUT_USAGE_TYPE
    PutValue targetSheet, targetRow, ColumnIndex("EH"), OUTPUT_USE_FOR
    PutText targetSheet, targetRow, ColumnIndex("EL"), mCustId & EMAIL_ID_SUFFIX
End Sub

' Tax-identifier group: only when a Tax ID was typed (ResolveTaxId guarantees type + country resolved).
' AV / AW are fixed "Y": the template's Transaction Tax ID / Primary Tax ID default is Yes and are not loaded.
Private Sub WriteTaxFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    If LenB(mTaxId) = 0 Then Exit Sub
    PutValue targetSheet, targetRow, ColumnIndex("AR"), OUTPUT_ROW_ID
    PutText targetSheet, targetRow, ColumnIndex("AT"), mTaxId
    PutValue targetSheet, targetRow, ColumnIndex("AU"), mTaxIdTypeId
    PutValue targetSheet, targetRow, ColumnIndex("AV"), OUTPUT_FLAG_YES
    PutValue targetSheet, targetRow, ColumnIndex("AW"), OUTPUT_FLAG_YES
End Sub

Private Sub WriteSettingsFields(ByVal targetSheet As Worksheet, ByVal targetRow As Long)
    PutValue targetSheet, targetRow, ColumnIndex("Q"), mCategoryId
    PutValue targetSheet, targetRow, ColumnIndex("S"), mGroupId
    PutValue targetSheet, targetRow, ColumnIndex("V"), mTermId
    PutValue targetSheet, targetRow, ColumnIndex("W"), mPayTypeId
    PutValue targetSheet, targetRow, ColumnIndex("LF"), mCurrencyId
    PutValue targetSheet, targetRow, ColumnIndex("LI"), mInvDeliveryId
    PutValue targetSheet, targetRow, ColumnIndex("LJ"), mInvRecipientsId
    PutValue targetSheet, targetRow, ColumnIndex("LK"), mStmtDeliveryId
    PutValue targetSheet, targetRow, ColumnIndex("LL"), mStmtRecipientsId
    PutValue targetSheet, targetRow, ColumnIndex("MB"), OUTPUT_ROW_ID
    PutValue targetSheet, targetRow, ColumnIndex("MD"), mTaxStatusCountry
    PutValue targetSheet, targetRow, ColumnIndex("ME"), mTaxStatusId
End Sub
