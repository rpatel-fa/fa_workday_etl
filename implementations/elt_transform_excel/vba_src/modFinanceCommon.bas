Attribute VB_Name = "modFinanceCommon"
Option Explicit

' Mapping MASTER v35 column contract (header row 1 on tab "MAPPING MASTER"). These are
' single-sourced here from the YAML file contract in config/reference_paths.yaml. The
' coding agent (not VBA runtime) reads the YAML to keep these constants in sync.  Core
' macro reads use only the 5 columns below; remaining v35 columns (OLD DIMENSION Name,
' Capex Entity, Code_Entity, Splits/Combos, WORKDAY NAME, cross-instance checks,
' parents, Reporting Type matrices) are Finance-only and not consumed.
Private Const MAPCOL_OLD_SEGMENT As String = "OLD SEGMENT"
Private Const MAPCOL_OLD_CODE As String = "OLD CODE"
Private Const MAPCOL_CORE_VS_PRISM As String = "Core vs Prism"
Private Const MAPCOL_WORKDAY_BUSINESS_OBJECT As String = "WORKDAY BUSINESS OBJECT"
Private Const MAPCOL_WORKDAY_REF_ID As String = "WORKDAY REF ID"
Private Const MAPVAL_CORE_ONLY As String = "Core"

Public Const WBO_COST_CENTER As String = "Cost Center"
Public Const WBO_SPEND_CATEGORY As String = "Spend Category"
Public Const WBO_EVENT As String = "Events - Custom Org"
Public Const WBO_CAMP As String = "Camps - Custom Org"
Public Const WBO_FA_LEARNING As String = "FA Learning - Custom Org"
Public Const WBO_COMMERCIAL_PARTNER As String = "Commercial Partners - Custom Worktag"
Public Const WBO_GRANT As String = "Grant"
Public Const WBO_PROGRAM As String = "Program - Delivered Org"
Public Const WBO_PROJECT As String = "Projects (Opex)"
Public Const WBO_PROJECT_CAPEX As String = "Projects (Capex)"

Public Const TAX_APPLICABILITY_ZERO As String = "GBR_Outside_of_Scope"
Public Const TAX_CODE_ZERO As String = "GBR_Zero_Rated"
Public Const TAX_RECOVERABILITY_ZERO As String = "Nonrecoverable_Allocate"
' Shared standard-rated triple. This is the CUSTOMER-INVOICE (AR) value -- GBR_Standard_Rated is
' a Sale/Customer ref-id. AR transforms (bulk, GBE, invoice_run, multi_line) resolve their
' standard-rated lines against these. AP SUPPLIER-INVOICE standard-rated lines use the Domestic
' PURCHASE triple instead (GBR_Domestic_Purchase_Of_Goods_And_Services / GBR_Standard /
' Fully_Recoverable); each AP module carries that as a LOCAL Private Const shadow of the same
' name (see ap_travel_places / ap_taxi / ap_stationery / ap_cfa_distributions / ap_amex_trains*).
' Centralising this split under an AP-specific shared constant is a build_runner tax-gate change
' tracked in docs/todo/unify-invoice-line-grouping.md -- do NOT change these shared AR values.
Public Const TAX_APPLICABILITY_STANDARD As String = "GBR_Standard_Rated"
Public Const TAX_CODE_STANDARD As String = "GBR_Standard_Rated"
Public Const TAX_RECOVERABILITY_STANDARD As String = "Recoverable_Allocate"

' Shared finance-domain helpers reused across transform profiles. Keep generic workbook/runtime
' plumbing in modHelpers; place reusable finance coding rules here when they apply to multiple
' transforms.

' Build a lookup from a tabular extract whose reference IDs are pipe-delimited in one column and
' whose resolved Workday ID sits in another. The caller owns finding the correct header row and
' source/target columns for its specific extract shape; this function owns the repeated row-walk.
Public Function BuildDelimitedReferenceLookup( _
    ByVal tableData As Variant, _
    ByVal headerRow As Long, _
    ByVal referenceIdColumn As Long, _
    ByVal targetIdColumn As Long _
) As Collection
    Dim rowIndex As Long
    Dim referenceIds As Variant
    Dim referenceIndex As Long
    Dim referenceId As String
    Dim targetId As String
    Dim lookup As Collection

    Set lookup = New Collection

    For rowIndex = headerRow + 1 To UBound(tableData, 1)
        targetId = Trim$(CStr(tableData(rowIndex, targetIdColumn)))
        If LenB(targetId) > 0 Then
            referenceIds = Split(CStr(tableData(rowIndex, referenceIdColumn)), "|")
            For referenceIndex = LBound(referenceIds) To UBound(referenceIds)
                referenceId = Trim$(CStr(referenceIds(referenceIndex)))
                If LenB(referenceId) > 0 Then
                    AddLookupIfMissing lookup, referenceId, targetId
                End If
            Next referenceIndex
        End If
    Next rowIndex

    Set BuildDelimitedReferenceLookup = lookup
End Function

Public Sub AddLookupIfMissing(ByVal lookup As Collection, ByVal keyText As String, ByVal value As String)
    On Error Resume Next
    lookup.Add Item:=value, Key:=LCase$(Trim$(keyText))
    On Error GoTo 0
End Sub

Public Function LookupCollectionValue(ByVal lookup As Collection, ByVal keyText As String) As String
    On Error Resume Next
    LookupCollectionValue = CStr(lookup.Item(LCase$(Trim$(keyText))))
    If Err.Number <> 0 Then
        LookupCollectionValue = vbNullString
    End If
    On Error GoTo 0
End Function

' Resolve a source vendor code against a pipe-delimited supplier lookup using two-part
' precedence:
'   1. vendorId + "-" + companyCode    -- exact supplier-per-company match (higher precedence)
'   2. vendorId alone                  -- bare fallback (when the vendor id maps to a single
'                                        supplier regardless of company)
' Example: AMP001-FA -> AMPTHILL TOWN FC (S-000386) wins over a bare AMP001 fallback, and
' AMP001-WNSL -> AMPERSAND (S-000385) for the same vendor prefix under a different company.
' The lookup parameter must have been built by BuildDelimitedReferenceLookup (so reference_ids
' are already split by pipe, trimmed, and keyed lower-case).
Public Function ResolveSupplierWithCompanyPrecedence( _
    ByVal lookup As Collection, _
    ByVal vendorId As String, _
    ByVal companyCode As String _
) As String
    Dim exactKey As String
    Dim bareKey As String
    Dim resolved As String

    exactKey = LCase$(Trim$(vendorId)) & "-" & LCase$(Trim$(companyCode))
    bareKey = LCase$(Trim$(vendorId))

    resolved = LookupCollectionValue(lookup, exactKey)
    If LenB(resolved) > 0 Then
        ResolveSupplierWithCompanyPrecedence = resolved
        Exit Function
    End If

    ResolveSupplierWithCompanyPrecedence = LookupCollectionValue(lookup, bareKey)
End Function

' Build a Reference-ID -> Workday-ID lookup from a Workday CCS extract workbook for the given entity
' ("Customer", "Supplier", ...). Assumes the standard extract shape: a header row carrying the entity
' name, a pipe-delimited "Reference ID" column, and an "<Entity> ID" column.
Public Function BuildReferenceLookupFromWorkbook(ByVal extractWorkbook As Workbook, ByVal entityLabel As String) As Collection
    Dim extractTable As Variant
    Dim headerRow As Long
    Dim referenceIdColumn As Long
    Dim workdayIdColumn As Long

    extractTable = ReadWorksheetTable(ResolveSourceWorksheet(extractWorkbook, vbNullString))

    headerRow = FindHeaderRow(extractTable, Array(entityLabel, "Reference ID", entityLabel & " ID"))
    referenceIdColumn = FindColumnIndexInRow(extractTable, headerRow, "Reference ID")
    workdayIdColumn = FindColumnIndexInRow(extractTable, headerRow, entityLabel & " ID")

    Set BuildReferenceLookupFromWorkbook = BuildDelimitedReferenceLookup(extractTable, headerRow, referenceIdColumn, workdayIdColumn)
End Function

'  ---------------------------------------------------------------------------
'  CUSTOMER NAME -> CUSTOMER ID RESOLUTION (two-level cascade)
'  Shared engine for sources (e.g. GBE) that carry customer NAME only, with
'  no GP Reference ID / Workday Customer ID / pipe-delimited ID. PRECISION-
'  FIRST at every level: ambiguous or absent names REJECT, never guess.
'  ---------------------------------------------------------------------------

'  Build a normalised customer-display-name -> Customer ID lookup, plus an
'  ambiguous-name set (a normalised name mapping to more than one Customer
'  ID). Reads the Customer Spreadsheet (Customer extract) workbook. The
'  caller resolves rows by calling ResolveCustomerIdMultiLevel().
Public Sub BuildCustomerNameLookupFromWorkbook( _
    ByVal extractWorkbook As Workbook, _
    ByRef nameToId As Collection, _
    ByRef ambiguousNames As Collection _
)
    Dim extractTable As Variant
    Dim headerRow As Long
    Dim nameColumn As Long
    Dim customerIdColumn As Long
    Dim rowIndex As Long
    Dim rawName As String
    Dim customerId As String
    Dim normKey As String
    Dim existingId As String

    extractTable = ReadWorksheetTable(ResolveSourceWorksheet(extractWorkbook, vbNullString))

    headerRow = FindHeaderRow(extractTable, Array("Customer", "Reference ID", "Customer ID"))
    nameColumn = FindColumnIndexInRow(extractTable, headerRow, "Customer")
    customerIdColumn = FindColumnIndexInRow(extractTable, headerRow, "Customer ID")

    For rowIndex = headerRow + 1 To UBound(extractTable, 1)
        rawName = Trim$(CStr(extractTable(rowIndex, nameColumn)))
        customerId = Trim$(CStr(extractTable(rowIndex, customerIdColumn)))
        If LenB(rawName) > 0 And LenB(customerId) > 0 Then
            normKey = NormalizeClubName(rawName)
            If LenB(normKey) > 0 Then
                existingId = LookupCollectionValue(nameToId, normKey)
                If LenB(existingId) = 0 Then
                    nameToId.Add customerId, normKey
                ElseIf StrComp(existingId, customerId, vbTextCompare) <> 0 Then
                    If LenB(LookupCollectionValue(ambiguousNames, normKey)) = 0 Then
                        ambiguousNames.Add "1", normKey
                    End If
                End If
            End If
        End If
    Next rowIndex
End Sub

'  Normalise a customer name for matching: NormalizeToken (lower-case, ASCII-
'  only, whitespace collapsed) then reconcile women's-team spellings so
'  "... WFC" matches "... WOMEN FC" and "... WOMENS" matches "... WOMEN".
Public Function NormalizeClubName(ByVal rawName As Variant) As String
    Dim base As String
    Dim tokens() As String
    Dim i As Long
    Dim tok As String
    Dim out As String

    base = NormalizeToken(rawName)
    If LenB(base) = 0 Then
        NormalizeClubName = vbNullString
        Exit Function
    End If

    tokens = Split(base, " ")
    out = vbNullString
    For i = LBound(tokens) To UBound(tokens)
        tok = tokens(i)
        Select Case tok
            Case "wfc"
                out = out & " women fc"
            Case "womens"
                out = out & " women"
            Case Else
                out = out & " " & tok
        End Select
    Next i
    NormalizeClubName = Trim$(out)
End Function

'  Apply the SAME transforms that the GP data-migration SQL pipeline runs on
'  customer names (elt_ingest_excel/.../src_fin_customer_raw.sql lines 105-
'  128) so that a GBE/non-GP source name can be L2-matched against a Workday
'  extract that already has those transforms baked in. Returns the raw
'  uppercase result; the caller then runs NormalizeClubName() on the result.
'
'  Order matches the SQL exactly:
'    1. Trim
'    2. UCase
'    3. Strip dots (.)
'    4. WOMEN'S -> WOMEN  (also WOMENS -> WOMEN, idempotent)
'    5. FOOTBALL CLUB -> FC
'    6. Strip any parenthesised content (...) including nested parens
'    7. Word-boundary token-level: UNITED -> UTD, AND -> &
'    8. Peel trailing legal suffixes (repeatable): LIMITED/LTD/COMPANY/PLC/LLP/INC
Public Function ApplySqlNameTransforms(ByVal rawName As String) As String
    Dim temp As String
    Dim i As Long
    Dim ch As String
    Dim parenDepth As Long
    Dim sb As String
    Dim tokens() As String
    Dim tok As String
    Dim out As String
    Dim stripped As String

    temp = Trim$(rawName)
    If LenB(temp) = 0 Then
        ApplySqlNameTransforms = vbNullString
        Exit Function
    End If

    temp = UCase$(temp)

    temp = Replace$(temp, ".", "")

    temp = Replace$(temp, "WOMEN'S", "WOMEN")
    temp = Replace$(temp, "WOMENS", "WOMEN")

    temp = Replace$(temp, "FOOTBALL CLUB", "FC")

    sb = vbNullString
    parenDepth = 0
    For i = 1 To Len(temp)
        ch = Mid$(temp, i, 1)
        If ch = "(" Then
            parenDepth = parenDepth + 1
        ElseIf ch = ")" Then
            If parenDepth > 0 Then parenDepth = parenDepth - 1
        Else
            If parenDepth = 0 Then
                sb = sb & ch
            End If
        End If
    Next i
    temp = sb

    tokens = Split(Trim$(temp), " ")
    out = vbNullString
    For i = LBound(tokens) To UBound(tokens)
        tok = Trim$(tokens(i))
        If LenB(tok) > 0 Then
            Select Case tok
                Case "UNITED"
                    out = out & " UTD"
                Case "AND"
                    out = out & " &"
                Case Else
                    out = out & " " & tok
            End Select
        End If
    Next i
    temp = Trim$(out)

    tokens = Split(temp, " ")
    Do While True
        If UBound(tokens) < LBound(tokens) Then Exit Do
        tok = UCase$(Trim$(tokens(UBound(tokens))))
        If tok = "LIMITED" Or tok = "LTD" Or tok = "COMPANY" _
            Or tok = "PLC" Or tok = "LLP" Or tok = "INC" Then
            stripped = vbNullString
            For i = LBound(tokens) To UBound(tokens) - 1
                If LenB(Trim$(tokens(i))) > 0 Then
                    stripped = stripped & " " & tokens(i)
                End If
            Next i
            temp = Trim$(stripped)
            tokens = Split(temp, " ")
        Else
            Exit Do
        End If
    Loop

    ApplySqlNameTransforms = Trim$(temp)
End Function

'  Two-level PRECISION-FIRST name resolver: L1 exact normalized match wins
'  immediately; an L1 MISS (NOT an L1 ambiguous) falls through to L2 (SQL-
'  transformed then normalized); an ambiguous name at EITHER level rejects
'  immediately with no fallback. Returns the resolved Customer ID, or "" if
'  both levels miss / any level is ambiguous, plus out-flags for the match
'  level (1/2/0), ambiguous flag, and the level where ambiguity triggered.
Public Function ResolveCustomerIdMultiLevel( _
    ByVal clubName As String, _
    ByVal nameToId As Collection, _
    ByVal ambiguousNames As Collection, _
    ByRef outMatchLevel As Long, _
    ByRef outIsAmbiguous As Boolean, _
    ByRef outAmbiguousLevel As Long _
) As String
    Dim keyL1 As String
    Dim keyL2 As String
    Dim sqlTransformed As String
    Dim result As String

    outMatchLevel = 0
    outIsAmbiguous = False
    outAmbiguousLevel = 0
    result = vbNullString

    keyL1 = NormalizeClubName(clubName)
    If LenB(keyL1) > 0 Then
        If LenB(LookupCollectionValue(ambiguousNames, keyL1)) > 0 Then
            outIsAmbiguous = True
            outAmbiguousLevel = 1
            ResolveCustomerIdMultiLevel = vbNullString
            Exit Function
        End If
        result = LookupCollectionValue(nameToId, keyL1)
        If LenB(result) > 0 Then
            outMatchLevel = 1
            ResolveCustomerIdMultiLevel = result
            Exit Function
        End If
    End If

    sqlTransformed = ApplySqlNameTransforms(clubName)
    keyL2 = NormalizeClubName(sqlTransformed)
    If LenB(keyL2) > 0 Then
        If LenB(LookupCollectionValue(ambiguousNames, keyL2)) > 0 Then
            outIsAmbiguous = True
            outAmbiguousLevel = 2
            ResolveCustomerIdMultiLevel = vbNullString
            Exit Function
        End If
        result = LookupCollectionValue(nameToId, keyL2)
        If LenB(result) > 0 Then
            outMatchLevel = 2
            ResolveCustomerIdMultiLevel = result
            Exit Function
        End If
    End If

    ResolveCustomerIdMultiLevel = vbNullString
End Function

' Resolve a source company code to its Workday company name via the runner's live "company" mapping
' rows, falling back to the caller's default when no mapping row matches.
Public Function ResolveMappedCompany(ByVal runnerWorkbook As Workbook, ByVal companyCode As String, ByVal fallback As String) As String
    Dim resolved As String

    resolved = LookupMappingValue(runnerWorkbook, "company", Trim$(companyCode))
    If LenB(resolved) > 0 Then
        ResolveMappedCompany = resolved
    Else
        ResolveMappedCompany = fallback
    End If
End Function

' Build an OLD CODE -> "old_segment<tab>WORKDAY REF ID" lookup from the finance mapping master,
' filtered both to one Workday business object and to "Core vs Prism" = Core rows only (whitelist
' Core rather than blacklist Prism, so a future third value never leaks in). Keeping old_segment in
' the shared lookup lets callers reproduce the SQL behavior of choosing the lowest / most specific
' hierarchy when multiple Final Code segments match the same business object.
Public Function BuildBusinessObjectHierarchyLookup( _
    ByVal mappingTable As Variant, _
    ByVal workdayBusinessObject As String _
) As Collection
    Dim headerRow As Long
    Dim oldCodeColumn As Long
    Dim oldSegmentColumn As Long
    Dim coreVsPrismColumn As Long
    Dim businessObjectColumn As Long
    Dim workdayRefIdColumn As Long
    Dim rowIndex As Long
    Dim sourceCode As String
    Dim oldSegment As String
    Dim coreVsPrism As String
    Dim businessObject As String
    Dim targetRefId As String
    Dim lookup As Collection

    Set lookup = New Collection

    headerRow = FindHeaderRow(mappingTable, Array(MAPCOL_OLD_CODE, MAPCOL_OLD_SEGMENT, MAPCOL_CORE_VS_PRISM, MAPCOL_WORKDAY_BUSINESS_OBJECT, MAPCOL_WORKDAY_REF_ID))
    oldCodeColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_OLD_CODE)
    oldSegmentColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_OLD_SEGMENT)
    coreVsPrismColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_CORE_VS_PRISM)
    businessObjectColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_WORKDAY_BUSINESS_OBJECT)
    workdayRefIdColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_WORKDAY_REF_ID)

    For rowIndex = headerRow + 1 To UBound(mappingTable, 1)
        sourceCode = Trim$(CStr(mappingTable(rowIndex, oldCodeColumn)))
        oldSegment = Trim$(CStr(mappingTable(rowIndex, oldSegmentColumn)))
        coreVsPrism = Trim$(CStr(mappingTable(rowIndex, coreVsPrismColumn)))
        businessObject = Trim$(CStr(mappingTable(rowIndex, businessObjectColumn)))
        targetRefId = Trim$(CStr(mappingTable(rowIndex, workdayRefIdColumn)))

        If LenB(sourceCode) > 0 _
            And LenB(targetRefId) > 0 _
            And StrComp(coreVsPrism, MAPVAL_CORE_ONLY, vbTextCompare) = 0 _
            And StrComp(businessObject, workdayBusinessObject, vbTextCompare) = 0 Then
            UpdateHierarchyLookup lookup, sourceCode, oldSegment, targetRefId
        End If
    Next rowIndex

    Set BuildBusinessObjectHierarchyLookup = lookup
End Function

Public Function BuildBusinessObjectHierarchyLookupMulti( _
    ByVal mappingTable As Variant, _
    ParamArray workdayBusinessObjects() As Variant _
) As Collection
    Dim headerRow As Long
    Dim oldCodeColumn As Long
    Dim oldSegmentColumn As Long
    Dim coreVsPrismColumn As Long
    Dim businessObjectColumn As Long
    Dim workdayRefIdColumn As Long
    Dim rowIndex As Long
    Dim sourceCode As String
    Dim oldSegment As String
    Dim coreVsPrism As String
    Dim businessObject As String
    Dim targetRefId As String
    Dim lookup As Collection
    Dim wboIndex As Long
    Dim matchWBO As Boolean
    Dim candidateWBO As String

    Set lookup = New Collection

    If UBound(workdayBusinessObjects) < LBound(workdayBusinessObjects) Then
        Set BuildBusinessObjectHierarchyLookupMulti = lookup
        Exit Function
    End If

    headerRow = FindHeaderRow(mappingTable, Array(MAPCOL_OLD_CODE, MAPCOL_OLD_SEGMENT, MAPCOL_CORE_VS_PRISM, MAPCOL_WORKDAY_BUSINESS_OBJECT, MAPCOL_WORKDAY_REF_ID))
    oldCodeColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_OLD_CODE)
    oldSegmentColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_OLD_SEGMENT)
    coreVsPrismColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_CORE_VS_PRISM)
    businessObjectColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_WORKDAY_BUSINESS_OBJECT)
    workdayRefIdColumn = FindColumnIndexInRow(mappingTable, headerRow, MAPCOL_WORKDAY_REF_ID)

    For rowIndex = headerRow + 1 To UBound(mappingTable, 1)
        sourceCode = Trim$(CStr(mappingTable(rowIndex, oldCodeColumn)))
        oldSegment = Trim$(CStr(mappingTable(rowIndex, oldSegmentColumn)))
        coreVsPrism = Trim$(CStr(mappingTable(rowIndex, coreVsPrismColumn)))
        businessObject = Trim$(CStr(mappingTable(rowIndex, businessObjectColumn)))
        targetRefId = Trim$(CStr(mappingTable(rowIndex, workdayRefIdColumn)))

        If LenB(sourceCode) > 0 _
            And LenB(targetRefId) > 0 _
            And StrComp(coreVsPrism, MAPVAL_CORE_ONLY, vbTextCompare) = 0 Then
            matchWBO = False
            For wboIndex = LBound(workdayBusinessObjects) To UBound(workdayBusinessObjects)
                candidateWBO = Trim$(CStr(workdayBusinessObjects(wboIndex)))
                If StrComp(businessObject, candidateWBO, vbTextCompare) = 0 Then
                    matchWBO = True
                    Exit For
                End If
            Next wboIndex
            If matchWBO Then
                UpdateHierarchyLookup lookup, sourceCode, oldSegment, targetRefId
            End If
        End If
    Next rowIndex

    Set BuildBusinessObjectHierarchyLookupMulti = lookup
End Function

Public Function FinalCodeSegment(ByVal finalCodeValue As Variant, ByVal segmentIndex As Long) As String
    Dim segments As Variant
    Dim normalized As String

    If segmentIndex < 1 Then
        Exit Function
    End If

    '  Normalise separators (hyphen/space/dot/underscore/...) to canonical "-" HERE, in the
    '  shared segment splitter, so EVERY caller (all transforms, gate + write) is delimiter-robust
    '  without re-normalising. NormalizeCostCodeSeparators is idempotent on already-dashed codes.
    normalized = NormalizeCostCodeSeparators(Trim$(CStr(finalCodeValue)))
    If LenB(normalized) = 0 Then
        Exit Function
    End If

    segments = Split(normalized, "-")
    If UBound(segments) < segmentIndex - 1 Then
        Exit Function
    End If

    FinalCodeSegment = Trim$(CStr(segments(segmentIndex - 1)))
End Function

' Resolve the best matching worktag across the Final Code accounting-string segments. This mirrors
' the finance SQL reference joins: check every accounting element, then choose the row with the
' lowest / most specific hierarchy (highest OLD SEGMENT).
Public Function ResolveLowestHierarchyLookupAcrossFinalCodeSegments( _
    ByVal lookup As Collection, _
    ByVal finalCodeValue As Variant _
) As String
    Dim normalized As String
    Dim segments As Variant
    Dim segmentIndex As Long
    Dim candidate As String
    Dim rawEntry As String
    Dim bestEntry As String
    Dim bestOldSegment As String
    Dim candidateOldSegment As String
    Dim candidateTargetRefId As String

    '  Normalise separators to canonical "-" HERE (shared resolver) so the accounting-string ->
    '  Workday-hierarchy lookup is delimiter-robust for EVERY caller. Idempotent on dashed codes.
    normalized = NormalizeCostCodeSeparators(Trim$(CStr(finalCodeValue)))
    If LenB(normalized) = 0 Then
        Exit Function
    End If

    segments = Split(normalized, "-")
    For segmentIndex = LBound(segments) To UBound(segments)
        candidate = Trim$(CStr(segments(segmentIndex)))
        If LenB(candidate) > 0 Then
            rawEntry = LookupCollectionValue(lookup, candidate)
            If LenB(rawEntry) > 0 Then
                candidateOldSegment = HierarchyLookupOldSegment(rawEntry)
                candidateTargetRefId = HierarchyLookupTargetRefId(rawEntry)
                If LenB(candidateTargetRefId) > 0 Then
                    If Not HierarchyTargetIsPlaceholder(candidateTargetRefId, candidateOldSegment) Then
                        If LenB(bestEntry) = 0 Or CompareHierarchySegments(candidateOldSegment, bestOldSegment) > 0 Then
                            bestEntry = rawEntry
                            bestOldSegment = candidateOldSegment
                            ResolveLowestHierarchyLookupAcrossFinalCodeSegments = candidateTargetRefId
                        End If
                    End If
                End If
            End If
        End If
    Next segmentIndex
End Function

Private Function HierarchyTargetIsPlaceholder( _
    ByVal targetRefId As String, _
    ByVal oldSegment As String _
) As Boolean
    Dim cleanTarget As String
    Dim cleanSegment As String

    cleanTarget = UCase$(Trim$(targetRefId))
    cleanSegment = UCase$(Trim$(oldSegment))

    If cleanTarget = "STRING OVERRIDE" Then
        HierarchyTargetIsPlaceholder = True
        Exit Function
    End If

    If cleanSegment = "STRING OVERRIDE" Then
        HierarchyTargetIsPlaceholder = True
        Exit Function
    End If

    HierarchyTargetIsPlaceholder = False
End Function

Public Function ResolveLookupFromFinalCodeSegment( _
    ByVal lookup As Collection, _
    ByVal finalCodeValue As Variant, _
    ByVal segmentIndex As Long _
) As String
    Dim rawEntry As String
    Dim candidateTargetRefId As String
    Dim candidateOldSegment As String

    rawEntry = LookupCollectionValue(lookup, FinalCodeSegment(finalCodeValue, segmentIndex))
    If LenB(rawEntry) = 0 Then
        Exit Function
    End If

    candidateOldSegment = HierarchyLookupOldSegment(rawEntry)
    candidateTargetRefId = HierarchyLookupTargetRefId(rawEntry)

    If Not HierarchyTargetIsPlaceholder(candidateTargetRefId, candidateOldSegment) Then
        ResolveLookupFromFinalCodeSegment = candidateTargetRefId
    End If
End Function

Public Function ResolveRoundedRate(ByVal sourceValue As Variant) As Long
    Dim normalized As String

    If IsNumeric(sourceValue) Then
        ResolveRoundedRate = CLng(Int(CDbl(sourceValue) + 0.5))
        Exit Function
    End If

    normalized = Trim$(CStr(sourceValue))
    If LenB(normalized) = 0 Then
        ResolveRoundedRate = 0
    ElseIf IsNumeric(normalized) Then
        ResolveRoundedRate = CLng(Int(CDbl(normalized) + 0.5))
    Else
        ResolveRoundedRate = 0
    End If
End Function

Public Function TaxValueForRate( _
    ByVal vatRate As Long, _
    ByVal zeroRatedValue As String, _
    ByVal standardRatedValue As String _
) As String
    If vatRate >= 20 Then
        TaxValueForRate = standardRatedValue
    Else
        TaxValueForRate = zeroRatedValue
    End If
End Function

Private Sub UpdateHierarchyLookup( _
    ByVal lookup As Collection, _
    ByVal sourceCode As String, _
    ByVal oldSegment As String, _
    ByVal targetRefId As String _
)
    Dim existingEntry As String
    Dim existingOldSegment As String
    Dim normalizedKey As String

    normalizedKey = LCase$(Trim$(sourceCode))
    existingEntry = LookupCollectionValue(lookup, sourceCode)
    If LenB(existingEntry) = 0 Then
        lookup.Add Item:=oldSegment & vbTab & targetRefId, Key:=normalizedKey
        Exit Sub
    End If

    existingOldSegment = HierarchyLookupOldSegment(existingEntry)
    If CompareHierarchySegments(oldSegment, existingOldSegment) > 0 Then
        lookup.Remove normalizedKey
        lookup.Add Item:=oldSegment & vbTab & targetRefId, Key:=normalizedKey
    End If
End Sub

Private Function HierarchyLookupOldSegment(ByVal rawEntry As String) As String
    Dim parts As Variant

    parts = Split(rawEntry, vbTab, 2)
    HierarchyLookupOldSegment = Trim$(CStr(parts(0)))
End Function

Private Function HierarchyLookupTargetRefId(ByVal rawEntry As String) As String
    Dim parts As Variant

    parts = Split(rawEntry, vbTab, 2)
    If UBound(parts) >= 1 Then
        HierarchyLookupTargetRefId = Trim$(CStr(parts(1)))
    End If
End Function

Private Function CompareHierarchySegments(ByVal leftValue As String, ByVal rightValue As String) As Long
    Dim leftNumber As Double
    Dim rightNumber As Double

    leftValue = Trim$(leftValue)
    rightValue = Trim$(rightValue)

    If LenB(leftValue) = 0 And LenB(rightValue) = 0 Then
        Exit Function
    End If
    If LenB(leftValue) = 0 Then
        CompareHierarchySegments = -1
        Exit Function
    End If
    If LenB(rightValue) = 0 Then
        CompareHierarchySegments = 1
        Exit Function
    End If

    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        leftNumber = CDbl(leftValue)
        rightNumber = CDbl(rightValue)
        If leftNumber > rightNumber Then
            CompareHierarchySegments = 1
        ElseIf leftNumber < rightNumber Then
            CompareHierarchySegments = -1
        End If
        Exit Function
    End If

    CompareHierarchySegments = StrComp(leftValue, rightValue, vbTextCompare)
End Function

Public Sub IncrementCount(ByVal counts As Collection, ByVal keys As Collection, ByVal keyText As String)
    Dim normalizedKey As String
    Dim currentValue As Long
    Dim isNew As Boolean

    normalizedKey = LCase$(Trim$(keyText))

    On Error Resume Next
    currentValue = CLng(counts.Item(normalizedKey))
    isNew = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    If isNew Then
        counts.Add Item:=CLng(1), Key:=normalizedKey
        keys.Add normalizedKey
    Else
        counts.Remove normalizedKey
        counts.Add Item:=currentValue + 1, Key:=normalizedKey
    End If
End Sub

Public Function BuildRejectedLookupIssue( _
    ByVal counts As Collection, _
    ByVal keys As Collection, _
    ByVal referenceLabel As String _
) As String
    Dim currentKey As Variant
    Dim parts As String
    Dim displayed As Long

    For Each currentKey In keys
        If LenB(parts) > 0 Then
            parts = parts & ", "
        End If
        parts = parts & CStr(currentKey) & " (" & CStr(counts.Item(CStr(currentKey))) & ")"
        displayed = displayed + 1
        If displayed = 5 Then
            Exit For
        End If
    Next currentKey

    BuildRejectedLookupIssue = _
        "Rejected source row(s) with unresolved " & referenceLabel & " mapping: " & parts
    If keys.Count > displayed Then
        BuildRejectedLookupIssue = BuildRejectedLookupIssue & ", ..."
    End If
End Function

' Supplier PO reporting-type mapping: split Final Code on "-", take accounting element 2,
' then map its first character.
Public Function ResolveReportingTypeFromFinalCode(ByVal finalCodeValue As Variant) As String
    Dim accountingElement2 As String

    accountingElement2 = AccountingElement2FromFinalCode(finalCodeValue)
    If LenB(accountingElement2) = 0 Then
        Exit Function
    End If

    ResolveReportingTypeFromFinalCode = ResolveReportingTypeCode(Left$(accountingElement2, 1))
End Function

Public Function ResolveInvoiceNumberFromSource( _
    ByVal sourceDataPath As String, _
    ByVal tableData As Variant, _
    ByVal sourceDateColumn As Long, _
    ByVal dataStartRowIndex As Long, _
    ByVal invoiceSuffix As String, _
    ByVal sourceColumnLabel As String, _
    ByVal errorSource As String, _
    ByVal errorNumber As Long _
) As String
    Dim periodToken As String

    periodToken = ExtractInvoicePeriodFromFileName(sourceDataPath)
    If LenB(periodToken) = 0 Then
        periodToken = DeriveDominantPeriodFromColumn(tableData, sourceDateColumn, dataStartRowIndex)
    End If
    If LenB(periodToken) = 0 Then
        Err.Raise errorNumber, errorSource, _
            "Unable to derive the invoice period for columns G/AU from the source file name or " & sourceColumnLabel & " column."
    End If

    ResolveInvoiceNumberFromSource = periodToken & invoiceSuffix
End Function

Private Function AccountingElement2FromFinalCode(ByVal finalCodeValue As Variant) As String
    AccountingElement2FromFinalCode = FinalCodeSegment(finalCodeValue, 2)
End Function

Private Function ResolveReportingTypeCode(ByVal codeText As String) As String
    Select Case Left$(Trim$(codeText), 1)
        Case "1", "2", "3", "7", "8", "9"
            ResolveReportingTypeCode = "FA Operations"
        Case "4"
            ResolveReportingTypeCode = "FA Investments"
        Case "5"
            ResolveReportingTypeCode = "National Game Board (NGB)"
        Case "6"
            ResolveReportingTypeCode = "Professional Game Board (PGB)"
    End Select
End Function

Private Function ExtractInvoicePeriodFromFileName(ByVal sourceDataPath As String) As String
    Dim fileName As String
    Dim normalized As String
    Dim tokens As Variant
    Dim index As Long
    Dim candidateMonth As String
    Dim candidateYear As String

    fileName = UCase$(LocalFileNameWithoutExtension(sourceDataPath))
    normalized = Replace(fileName, "_", " ")
    normalized = Replace(normalized, "-", " ")
    normalized = Replace(normalized, ".", " ")
    tokens = Split(Application.WorksheetFunction.Trim(normalized), " ")

    For index = LBound(tokens) To UBound(tokens) - 1
        candidateMonth = NormalizeMonthToken(CStr(tokens(index)))
        candidateYear = Trim$(CStr(tokens(index + 1)))
        If LenB(candidateMonth) > 0 And Len(candidateYear) = 2 And IsNumeric(candidateYear) Then
            ExtractInvoicePeriodFromFileName = candidateMonth & candidateYear
            Exit Function
        End If
    Next index
End Function

Private Function NormalizeMonthToken(ByVal tokenText As String) As String
    Dim normalized As String

    normalized = Left$(UCase$(Trim$(tokenText)), 3)
    Select Case normalized
        Case "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
            NormalizeMonthToken = normalized
    End Select
End Function

Private Function LocalFileNameWithoutExtension(ByVal filePath As String) As String
    Dim fileName As String
    Dim slashPosition As Long
    Dim colonPosition As Long
    Dim backslashPosition As Long
    Dim splitPosition As Long
    Dim dotPosition As Long

    slashPosition = InStrRev(filePath, "/")
    colonPosition = InStrRev(filePath, ":")
    backslashPosition = InStrRev(filePath, "\")
    splitPosition = slashPosition
    If colonPosition > splitPosition Then
        splitPosition = colonPosition
    End If
    If backslashPosition > splitPosition Then
        splitPosition = backslashPosition
    End If

    If splitPosition = 0 Then
        fileName = filePath
    Else
        fileName = Mid$(filePath, splitPosition + 1)
    End If

    dotPosition = InStrRev(fileName, ".")
    If dotPosition = 0 Then
        LocalFileNameWithoutExtension = fileName
    Else
        LocalFileNameWithoutExtension = Left$(fileName, dotPosition - 1)
    End If
End Function

Private Function DeriveDominantPeriodFromColumn( _
    ByVal tableData As Variant, _
    ByVal columnIndex As Long, _
    ByVal startRow As Long _
) As String
    Dim counts As Collection
    Dim keys As Collection
    Dim rowIndex As Long
    Dim currentToken As String
    Dim bestToken As String
    Dim bestCount As Long
    Dim currentKey As Variant

    If columnIndex <= 0 Then
        Exit Function
    End If

    Set counts = New Collection
    Set keys = New Collection

    For rowIndex = startRow To UBound(tableData, 1)
        currentToken = PeriodTokenFromDateValue(tableData(rowIndex, columnIndex))
        If LenB(currentToken) > 0 Then
            IncrementCount counts, keys, currentToken
        End If
    Next rowIndex

    For Each currentKey In keys
        If CLng(counts.Item(CStr(currentKey))) > bestCount Then
            bestCount = CLng(counts.Item(CStr(currentKey)))
            bestToken = CStr(currentKey)
        End If
    Next currentKey

    DeriveDominantPeriodFromColumn = bestToken
End Function

Private Function PeriodTokenFromDateValue(ByVal sourceValue As Variant) As String
    Dim normalized As String
    Dim parsedDate As Date

    If IsDate(sourceValue) Then
        parsedDate = CDate(sourceValue)
        PeriodTokenFromDateValue = UCase$(Format$(parsedDate, "mmmyy"))
        Exit Function
    End If

    normalized = NormalizeDateText(sourceValue)
    If LenB(normalized) = 10 And Mid$(normalized, 5, 1) = "-" Then
        On Error Resume Next
        parsedDate = DateSerial(CLng(Left$(normalized, 4)), CLng(Mid$(normalized, 6, 2)), CLng(Right$(normalized, 2)))
        If Err.Number = 0 Then
            PeriodTokenFromDateValue = UCase$(Format$(parsedDate, "mmmyy"))
        End If
        Err.Clear
        On Error GoTo 0
    End If
End Function

Private Function NormalizeDateText(ByVal sourceValue As Variant) As String
    Dim normalized As String
    Dim parts As Variant
    Dim dayPart As String
    Dim monthPart As String
    Dim yearPart As String
    Dim dayNumber As Long
    Dim monthNumber As Long

    normalized = Trim$(CStr(sourceValue))
    If LenB(normalized) = 0 Then
        Exit Function
    End If

    normalized = Replace(normalized, "-", "/")
    normalized = Replace(normalized, ".", "/")
    parts = Split(normalized, "/")

    If UBound(parts) = 2 Then
        dayPart = Right$("00" & Trim$(CStr(parts(0))), 2)
        monthPart = Right$("00" & Trim$(CStr(parts(1))), 2)
        yearPart = Trim$(CStr(parts(2)))
        If IsNumeric(dayPart) And IsNumeric(monthPart) And IsNumeric(yearPart) Then
            dayNumber = CLng(dayPart)
            monthNumber = CLng(monthPart)
            If dayNumber >= 1 And dayNumber <= 31 And monthNumber >= 1 And monthNumber <= 12 Then
                If Len(yearPart) = 2 Then
                    yearPart = "20" & yearPart
                End If
                NormalizeDateText = yearPart & "-" & monthPart & "-" & dayPart
                Exit Function
            End If
        End If
    End If

    If IsDate(sourceValue) Then
        NormalizeDateText = Format$(CDate(sourceValue), "yyyy-mm-dd")
        Exit Function
    End If

    NormalizeDateText = normalized
End Function
