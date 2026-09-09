Attribute VB_Name = "modSourceTable"
Option Explicit

' Generic, domain-free operations on a source table already read into a 1-based 2D Variant array
' by modHelpers.ReadWorksheetTable: locate headers/columns, validate required headers, and coerce
' cell values (amounts, dates). Shared by every transform module. No column name, sheet name, or
' business value specific to one transform belongs here -- those stay in the per-profile module.

' Convert an Excel column-letter reference ("A", "AU", ...) to its 1-based column number.
Public Function ColumnIndex(ByVal columnLetters As String) As Long
    Dim characterIndex As Long
    Dim result As Long

    For characterIndex = 1 To Len(columnLetters)
        result = (result * 26) + (Asc(UCase$(Mid$(columnLetters, characterIndex, 1))) - Asc("A") + 1)
    Next characterIndex

    ColumnIndex = result
End Function

' Case/whitespace-insensitive cell token used for header matching; empty for error/blank cells.
' Internal runs of whitespace (multiple spaces, tabs, non-breaking spaces etc.) are
' collapsed to a single ordinary space so source files with accidental double-spaces
' in header names (e.g. "INVOICE  LINE DESCRIPTION") still match the declared
' requiredHeaders text ("INVOICE LINE DESCRIPTION").
' Non-ASCII characters (chr() above 127, e.g. currency-symbol doubled UTF-8 bytes,
' smart-quotes, em-dashes) are DROPPED entirely.  This is deliberate defence-in-depth
' against two failure modes:
'   (a) Byte-order-mark / UTF-8 vs Windows-1252 vs Latin-1 encoding drift when .bas
'       modules are re-imported.  Symptom: a single "pound-currency" char literal in a
'       module's source turns into two VISIBLE chars (capital-A-circumflex + pound)
'       at runtime in the VBA host on Mac Excel.
'   (b) Finance putting currency symbols or typographic quotes into header text in
'       some months but not others.
' After stripping, tokens like "Total Claim (capital-A-circumflex + pound)",
' "Total Claim (pound)" and "Total Claim ()" all collapse to the same normalised
' ASCII form "total claim ()".
Public Function NormalizeToken(ByVal value As Variant) As String
    Dim raw As String
    Dim i As Long
    Dim ch As String
    Dim codePt As Long
    Dim inWS As Boolean
    Dim sb As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        NormalizeToken = vbNullString
        Exit Function
    End If
    raw = Trim$(CStr(value))
    If LenB(raw) = 0 Then
        NormalizeToken = vbNullString
        Exit Function
    End If
    sb = vbNullString
    inWS = False
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        codePt = AscW(ch)
        If codePt < 32 Or codePt > 126 Then
            ' Strip control chars, DEL, and all non-ASCII Unicode (currency marks,
            ' smart quotes, em dashes, UTF-8 doubled-byte residues from VBA host
            ' import-encoding drift etc.).
            GoTo nextCharNormalize
        End If
        Select Case ch
            Case " ", vbTab, Chr$(160), Chr$(10), Chr$(13)
                If Not inWS Then
                    sb = sb & " "
                    inWS = True
                End If
            Case Else
                sb = sb & ch
                inWS = False
        End Select
nextCharNormalize:
    Next i
    NormalizeToken = LCase$(Trim$(sb))
End Function

' Split a header specifier into its alias array.
'   - Plain text "Description" -> one-element array ("Description")
'   - Pipe-delimited "Decscription|Description" -> two elements ("Decscription", "Description")
' Leading/trailing spaces around each alternative are stripped; empty alternatives are skipped.
Public Function SplitHeaderAliases(ByVal headerSpec As String) As Variant
    Dim parts() As String
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim alt As String
    Dim out() As String

    If InStr(headerSpec, "|") = 0 Then
        ReDim out(0 To 0) As String
        out(0) = Trim$(headerSpec)
        SplitHeaderAliases = out
        Exit Function
    End If
    parts = Split(headerSpec, "|")
    n = UBound(parts) - LBound(parts) + 1
    ReDim out(0 To n - 1) As String
    j = 0
    For i = LBound(parts) To UBound(parts)
        alt = Trim$(parts(i))
        If LenB(alt) > 0 Then
            out(j) = alt
            j = j + 1
        End If
    Next i
    If j = 0 Then
        ReDim out(0 To 0) As String
        out(0) = vbNullString
        j = 1
    End If
    If j - 1 < UBound(out) Then
        ReDim Preserve out(0 To j - 1) As String
    End If
    SplitHeaderAliases = out
End Function

' True iff ANY of the aliases in headerSpec (pipe-delimited or plain) matches
' a cell on rowIndex (normalized compare). Single source of truth for alias OR-matching.
Public Function RowContainsHeader( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal headerText As String _
) As Boolean
    Dim columnIndex As Long
    Dim aliases As Variant
    Dim i As Long
    Dim aliasNorm As String

    aliases = SplitHeaderAliases(headerText)
    For columnIndex = 1 To UBound(tableData, 2)
        Dim cellNorm As String
        cellNorm = NormalizeToken(tableData(rowIndex, columnIndex))
        If LenB(cellNorm) = 0 Then GoTo nextCol
        For i = LBound(aliases) To UBound(aliases)
            aliasNorm = NormalizeToken(CStr(aliases(i)))
            If LenB(aliasNorm) = 0 Then GoTo nextAliasI
            If cellNorm = aliasNorm Then
                RowContainsHeader = True
                Exit Function
            End If
nextAliasI:
        Next i
nextCol:
    Next columnIndex
    RowContainsHeader = False
End Function

' First column on rowIndex whose header matches any alias in headerText;
' raises if none (required header). Accepts pipe-delimited alias lists
' ("Decscription|Description") OR plain single-name strings.
Public Function FindColumnIndexInRow( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal headerText As String _
) As Long
    Dim columnIndex As Long
    Dim aliases As Variant
    Dim i As Long
    Dim aliasNorm As String

    aliases = SplitHeaderAliases(headerText)
    For columnIndex = 1 To UBound(tableData, 2)
        Dim cellNorm As String
        cellNorm = NormalizeToken(tableData(rowIndex, columnIndex))
        If LenB(cellNorm) = 0 Then GoTo nextCol2
        For i = LBound(aliases) To UBound(aliases)
            aliasNorm = NormalizeToken(CStr(aliases(i)))
            If LenB(aliasNorm) = 0 Then GoTo nextAliasI2
            If cellNorm = aliasNorm Then
                FindColumnIndexInRow = columnIndex
                Exit Function
            End If
nextAliasI2:
        Next i
nextCol2:
    Next columnIndex

    Err.Raise vbObjectError + 1611, "modSourceTable.FindColumnIndexInRow", "Header not found: " & headerText
End Function

' As FindColumnIndexInRow but returns 0 when the header is absent (optional column).
Public Function FindOptionalColumnIndexInRow( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal headerText As String _
) As Long
    Dim columnIndex As Long
    Dim aliases As Variant
    Dim i As Long
    Dim aliasNorm As String

    aliases = SplitHeaderAliases(headerText)
    For columnIndex = 1 To UBound(tableData, 2)
        Dim cellNorm As String
        cellNorm = NormalizeToken(tableData(rowIndex, columnIndex))
        If LenB(cellNorm) = 0 Then GoTo nextCol3
        For i = LBound(aliases) To UBound(aliases)
            aliasNorm = NormalizeToken(CStr(aliases(i)))
            If LenB(aliasNorm) = 0 Then GoTo nextAliasI3
            If cellNorm = aliasNorm Then
                FindOptionalColumnIndexInRow = columnIndex
                Exit Function
            End If
nextAliasI3:
        Next i
nextCol3:
    Next columnIndex
End Function

' Detect the header row by content (the first row carrying every required header), so a transform
' does not assume a fixed physical row when source files vary in leading title/blank rows.
' If no row matches 100%, raise a detailed error identifying:
'   * the row whose match count was highest ("best candidate row"),
'   * EXACTLY which required headers were missing on that row,
'   * the list of all source file rows that had >= 25% match (so the user can eyeball where the
'     header row might really be if the config is misnamed).
' Each requiredHeaders entry may be a pipe-delimited alias list; any one alternative counts as present.
Public Function FindHeaderRow(ByVal tableData As Variant, ByVal requiredHeaders As Variant) As Long
    Const CANDIDATE_MIN_PCT As Single = 0.25
    Dim totalRows As Long
    Dim totalCols As Long
    Dim requiredCount As Long
    Dim headerIndex As Long
    Dim rowIndex As Long
    Dim matchedHeaders As Long
    Dim bestRow As Long
    Dim bestRowMatches As Long
    Dim bestMissing As String
    Dim candidateRows As String
    Dim missingList As String
    Dim thisMissing As String
    Dim presentThisRow() As Boolean
    Dim minCandidateMatch As Long

    totalRows = UBound(tableData, 1)
    totalCols = UBound(tableData, 2)
    requiredCount = UBound(requiredHeaders) - LBound(requiredHeaders) + 1
    minCandidateMatch = CLng(CSng(requiredCount) * CANDIDATE_MIN_PCT)
    If minCandidateMatch < 1 Then minCandidateMatch = 1
    ReDim presentThisRow(LBound(requiredHeaders) To UBound(requiredHeaders))

    bestRow = 0
    bestRowMatches = 0
    bestMissing = vbNullString
    candidateRows = vbNullString

    For rowIndex = 1 To totalRows
        matchedHeaders = 0
        missingList = vbNullString
        For headerIndex = LBound(requiredHeaders) To UBound(requiredHeaders)
            presentThisRow(headerIndex) = RowContainsHeader(tableData, rowIndex, CStr(requiredHeaders(headerIndex)))
            If presentThisRow(headerIndex) Then
                matchedHeaders = matchedHeaders + 1
            Else
                If Len(missingList) = 0 Then
                    missingList = "  - " & CStr(requiredHeaders(headerIndex))
                Else
                    missingList = missingList & vbCrLf & "  - " & CStr(requiredHeaders(headerIndex))
                End If
            End If
        Next headerIndex

        If matchedHeaders = requiredCount Then
            FindHeaderRow = rowIndex
            Exit Function
        End If

        If matchedHeaders > bestRowMatches Then
            bestRow = rowIndex
            bestRowMatches = matchedHeaders
            bestMissing = missingList
        End If
        If matchedHeaders >= minCandidateMatch Then
            If Len(candidateRows) = 0 Then
                candidateRows = CStr(rowIndex) & " (" & CStr(matchedHeaders) & "/" & CStr(requiredCount) & ")"
            Else
                candidateRows = candidateRows & ", " & CStr(rowIndex) & " (" & CStr(matchedHeaders) & "/" & CStr(requiredCount) & ")"
            End If
        End If
    Next rowIndex

    ' No full match -- produce the consolidated diagnostic
    thisMissing = vbNullString
    If bestRow > 0 Then
        thisMissing = _
            "Unable to find expected header row." & vbCrLf & _
            "  Required header count  : " & CStr(requiredCount) & vbCrLf & _
            "  Best candidate row     : row " & CStr(bestRow) & " (matched " & CStr(bestRowMatches) & "/" & CStr(requiredCount) & ")" & vbCrLf & _
            "  Missing on row " & CStr(bestRow) & ":" & vbCrLf & bestMissing
    Else
        thisMissing = _
            "Unable to find expected header row." & vbCrLf & _
            "  Required header count: " & CStr(requiredCount)
    End If
    If Len(candidateRows) > 0 Then
        thisMissing = thisMissing & vbCrLf & "  Other rows with >=25% match: rows " & candidateRows
    End If
    thisMissing = thisMissing & vbCrLf & _
        "  Tip: run `validate_input.py --transform <key>` from the CLI for a richer " & _
        "pre-flight schema check that enumerates every header and its column position."

    Err.Raise vbObjectError + 1612, "modSourceTable.FindHeaderRow", thisMissing
End Function

' Assert every required header is present on rowIndex; raises otherwise.
' Each requiredHeaders entry may be a pipe-delimited alias list; any one alternative counts.
Public Sub ValidateHeaderRow(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal requiredHeaders As Variant)
    Dim headerIndex As Long

    For headerIndex = LBound(requiredHeaders) To UBound(requiredHeaders)
        If Not RowContainsHeader(tableData, rowIndex, CStr(requiredHeaders(headerIndex))) Then
            Err.Raise vbObjectError + 1613, "modSourceTable.ValidateHeaderRow", _
                "Expected header '" & CStr(requiredHeaders(headerIndex)) & "' not found on source row " & CStr(rowIndex) & "."
        End If
    Next headerIndex
End Sub

' --- Schema validation phase (Tier 2 VBA defence-in-depth) ----------------
' Consolidates multiple schema failures into one consolidated error.
' Called by each transform module AFTER ResolveSourceWorksheet + ReadWorksheetTable
' but BEFORE FindHeaderRow, so the user sees the full list of problems with their
' source input file in a single dialog (rather than error-1-fix, error-2-fix, ...)
'
' Failures produced:
'   - Sheet has no rows or no columns (empty file)
'   - declared header_row (if any) is outside bounds
'   - declared data_start_row is <= header_row
'   - Optional: explicit headerRow passed (literal_row mode); verifies that exact row has ALL required headers.
'     If headerRow = 0 we just rely on FindHeaderRow to scan + enumerate.
'   - Optional headers: enumerated with col positions, informational (never fail).
'
' Public return: schema failures as newline-delimited string; empty string = OK.
' Each transform module calls this; if non-empty, raises one consolidated error
' via the transform's global error handler so the dialog is branded with transform key.
Public Function ValidateSourceWorkbookSchema( _
    ByVal sheetName As String, _
    ByVal tableData As Variant, _
    ByVal declaredHeaderRow As Long, _
    ByVal declaredDataStartRow As Long, _
    ByVal requiredHeaders As Variant, _
    Optional ByVal optionalHeaders As Variant = Empty, _
    Optional ByRef outHeaderRow As Long = 0, _
    Optional ByRef outHeaderDiagnostics As String = vbNullString _
) As String
    Dim totalRows As Long
    Dim totalCols As Long
    Dim requiredCount As Long
    Dim failures As String
    Dim diagnostics As String
    Dim headerIndex As Long
    Dim presentRequired As Long
    Dim missingList As String
    Dim infoList As String
    Dim foundCol As Long
    Dim headerRowToUse As Long
    Dim scanEnd As Long
    Dim ri As Long
    Dim hm As Long
    Dim bestHR As Long
    Dim bestHm As Long
    Dim optInfo As String
    Dim hasOpt As Boolean
    Dim optText As String
    Dim optionalHeadersProvided As Boolean

    failures = vbNullString
    diagnostics = vbNullString
    outHeaderDiagnostics = vbNullString

    ' Distinguish caller-omitted optional (Empty, not an array) from explicit empty arrays
    optionalHeadersProvided = IsArray(optionalHeaders)

    ' --- Dimensional checks
    totalRows = UBound(tableData, 1)
    totalCols = UBound(tableData, 2)
    If totalRows <= 0 Then
        failures = failures & vbCrLf & "  - Source sheet '" & sheetName & "' has no rows."
    End If
    If totalCols <= 0 Then
        failures = failures & vbCrLf & "  - Source sheet '" & sheetName & "' has no columns."
    End If
    If Len(failures) > 0 Then GoTo done_fast

    requiredCount = UBound(requiredHeaders) - LBound(requiredHeaders) + 1

    ' declared rows are sensible (only check if declared / non-zero - caller passes 0 if unknown)
    If declaredHeaderRow > 0 Then
        If declaredHeaderRow < 1 Or declaredHeaderRow > totalRows Then
            failures = failures & vbCrLf & _
                "  - Declared source_layout.header_row (" & CStr(declaredHeaderRow) & _
                ") is outside the source sheet row range (1-" & CStr(totalRows) & ")."
        End If
    End If
    If declaredDataStartRow > 0 And declaredHeaderRow > 0 Then
        If declaredDataStartRow <= declaredHeaderRow Then
            failures = failures & vbCrLf & _
                "  - source_layout.data_start_row (" & CStr(declaredDataStartRow) & _
                ") must be strictly greater than header_row (" & CStr(declaredHeaderRow) & ")."
        End If
    End If

    ' --- Required headers on the declared or scanned header row
    headerRowToUse = 0
    If declaredHeaderRow > 0 Then
        headerRowToUse = declaredHeaderRow
    Else
        ' Fallback: scan up to 40 rows for the first row with 100% required header match;
        ' fall back to best match (highest count) if no row hits 100%.
        scanEnd = 40
        If totalRows < scanEnd Then scanEnd = totalRows
        bestHR = 0
        bestHm = 0
        For ri = 1 To scanEnd
            hm = 0
            For headerIndex = LBound(requiredHeaders) To UBound(requiredHeaders)
                If RowContainsHeader(tableData, ri, CStr(requiredHeaders(headerIndex))) Then
                    hm = hm + 1
                End If
            Next headerIndex
            If hm = requiredCount Then
                headerRowToUse = ri
                Exit For
            End If
            If hm > bestHm Then
                bestHm = hm
                bestHR = ri
            End If
        Next ri
        If headerRowToUse = 0 And bestHR > 0 Then
            headerRowToUse = bestHR
        End If
    End If
    outHeaderRow = headerRowToUse

    ' Check required headers on the target row
    If headerRowToUse > 0 Then
        presentRequired = 0
        missingList = vbNullString
        For headerIndex = LBound(requiredHeaders) To UBound(requiredHeaders)
            If RowContainsHeader(tableData, headerRowToUse, CStr(requiredHeaders(headerIndex))) Then
                presentRequired = presentRequired + 1
            Else
                If Len(missingList) = 0 Then
                    missingList = "  - Missing: " & CStr(requiredHeaders(headerIndex))
                Else
                    missingList = missingList & ", " & CStr(requiredHeaders(headerIndex))
                End If
            End If
        Next headerIndex
        If presentRequired < requiredCount Then
            failures = failures & vbCrLf & _
                "  - Row " & CStr(headerRowToUse) & " matched only " & CStr(presentRequired) & _
                "/" & CStr(requiredCount) & " required headers." & vbCrLf & missingList
        End If
        ' Append informational diagnostics
        infoList = vbNullString
        For headerIndex = LBound(requiredHeaders) To UBound(requiredHeaders)
            If RowContainsHeader(tableData, headerRowToUse, CStr(requiredHeaders(headerIndex))) Then
                foundCol = FindColumnIndexInRow(tableData, headerRowToUse, CStr(requiredHeaders(headerIndex)))
                If Len(infoList) = 0 Then
                    infoList = "Row " & CStr(headerRowToUse) & " required: " & _
                        CStr(requiredHeaders(headerIndex)) & "=col" & CStr(foundCol)
                Else
                    infoList = infoList & "; " & _
                        CStr(requiredHeaders(headerIndex)) & "=col" & CStr(foundCol)
                End If
            End If
        Next headerIndex
        diagnostics = infoList
    Else
        failures = failures & vbCrLf & _
            "  - Unable to locate any row containing required headers (rows 1-40 scanned)."
    End If

    ' --- Optional headers (informational, never a failure).
    '     Only run when the caller explicitly passed an array of optional header names;
    '     skip when Optional argument was omitted by caller (Empty, non-array).
    If optionalHeadersProvided Then
        optInfo = vbNullString
        hasOpt = False
        If headerRowToUse > 0 Then
            For headerIndex = LBound(optionalHeaders) To UBound(optionalHeaders)
                optText = CStr(optionalHeaders(headerIndex))
                If RowContainsHeader(tableData, headerRowToUse, optText) Then
                    hasOpt = True
                    foundCol = FindOptionalColumnIndexInRow(tableData, headerRowToUse, optText)
                    If Len(optInfo) = 0 Then
                        optInfo = "  Optional: " & optText & "=col" & CStr(foundCol)
                    Else
                        optInfo = optInfo & "; " & optText & "=col" & CStr(foundCol)
                    End If
                Else
                    If Len(optInfo) = 0 Then
                        optInfo = "  Optional: " & optText & "=<absent>"
                    Else
                        optInfo = optInfo & "; " & optText & "=<absent>"
                    End If
                End If
            Next headerIndex
        End If
        If hasOpt Or Len(optInfo) > 0 Then
            If Len(diagnostics) = 0 Then
                diagnostics = optInfo
            Else
                diagnostics = diagnostics & vbCrLf & optInfo
            End If
        End If
    End If

    outHeaderDiagnostics = diagnostics

done_fast:
    If Len(failures) > 0 Then
        ValidateSourceWorkbookSchema = _
            "Source schema validation failed for sheet '" & sheetName & "':" & vbCrLf & _
            failures & vbCrLf & vbCrLf & _
            "  Hint: run `uv run --with openpyxl --with pyyaml --no-project python " & _
            "elt_transform_excel/scripts/validate_input.py -t <transform_key>` from CLI for " & _
            "the full 14-point schema check report."
    Else
        ValidateSourceWorkbookSchema = vbNullString
    End If
End Function

' Coerce a source amount cell to Double: numeric as-is, blank to 0, otherwise strip thousands commas.
Public Function ResolveAmount(ByVal sourceValue As Variant) As Double
    If IsNumeric(sourceValue) Then
        ResolveAmount = CDbl(sourceValue)
    ElseIf LenB(Trim$(CStr(sourceValue))) = 0 Then
        ResolveAmount = 0#
    Else
        ResolveAmount = CDbl(Replace(CStr(sourceValue), ",", vbNullString))
    End If
End Function

' Parse a source date value to canonical YYYY-MM-DD. Manual DMY parsing runs before CDate so the
' host machine's regional date settings can't silently swap day/month (see the OpenText rationale
' on modHelpers.OpenWorkbookReadOnly); falls back to IsDate/CDate for genuine Date-typed cells,
' then to the normalized source text.
Public Function NormalizeSourceDate(ByVal sourceValue As Variant) As String
    Dim normalized As String
    Dim parts As Variant
    Dim dayPart As String
    Dim monthPart As String
    Dim yearPart As String
    Dim dayNumber As Long
    Dim monthNumber As Long
    Dim serialCandidate As Double
    Dim isNumericSerial As Boolean

    normalized = Trim$(CStr(sourceValue))
    If LenB(normalized) = 0 Then
        NormalizeSourceDate = vbNullString
        Exit Function
    End If

    isNumericSerial = False
    If VarType(sourceValue) = vbDouble Or VarType(sourceValue) = vbCurrency Or _
       VarType(sourceValue) = vbInteger Or VarType(sourceValue) = vbLong Then
        serialCandidate = CDbl(sourceValue)
        isNumericSerial = (serialCandidate >= 1# And serialCandidate <= 100000#)
    End If
    If Not isNumericSerial Then
        Dim decSep As String
        Dim serialText As String
        Dim dotCount As Long
        Dim i As Long
        Dim ch As String
        decSep = Application.International(xlDecimalSeparator)
        serialText = normalized
        dotCount = 0
        For i = 1 To Len(serialText)
            ch = Mid$(serialText, i, 1)
            If (ch < "0" Or ch > "9") And ch <> decSep Then
                Exit For
            End If
            If ch = decSep Then dotCount = dotCount + 1
        Next i
        If i > Len(serialText) And (dotCount = 0 Or dotCount = 1) Then
            If IsNumeric(serialText) Then
                serialCandidate = CDbl(serialText)
                isNumericSerial = (serialCandidate >= 1# And serialCandidate <= 100000#)
            End If
        End If
    End If
    If isNumericSerial Then
        NormalizeSourceDate = Format$(CDate(Fix(serialCandidate)), "yyyy-mm-dd")
        Exit Function
    End If

    normalized = Replace(normalized, "-", "/")
    Dim normalizedDot As String
    normalizedDot = Replace(normalized, ".", "/")
    Dim dotParts As Variant
    dotParts = Split(normalizedDot, "/")
    If UBound(dotParts) = 2 Then
        Dim allThreeNumeric As Boolean
        allThreeNumeric = IsNumeric(Trim$(CStr(dotParts(0)))) And _
                          IsNumeric(Trim$(CStr(dotParts(1)))) And _
                          IsNumeric(Trim$(CStr(dotParts(2))))
        If allThreeNumeric Then
            normalized = normalizedDot
        End If
    End If
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
                NormalizeSourceDate = yearPart & "-" & monthPart & "-" & dayPart
                Exit Function
            End If
        End If
    End If

    If IsDate(sourceValue) Then
        NormalizeSourceDate = Format$(CDate(sourceValue), "yyyy-mm-dd")
        Exit Function
    End If

    If UBound(parts) = 2 Then
        If Len(yearPart) = 2 Then
            yearPart = "20" & yearPart
        End If
        NormalizeSourceDate = yearPart & "-" & monthPart & "-" & dayPart
    Else
        NormalizeSourceDate = normalized
    End If
End Function

Public Function NormaliseToWordBoundaries(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim ch As String
    Dim prevWasSpace As Boolean
    out = vbNullString
    prevWasSpace = False
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
            prevWasSpace = False
        Else
            If Not prevWasSpace Then
                out = out & " "
                prevWasSpace = True
            End If
        End If
    Next i
    NormaliseToWordBoundaries = out
End Function

Public Function NormalizeCostCodeSeparators(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim ch As String
    Dim prevWasDash As Boolean
    Dim dashCount As Long
    out = vbNullString
    prevWasDash = False
    dashCount = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "-", " ", ".", "_", "|", ":", "/", "\", ";", ","
                If Not prevWasDash Then
                    out = out & "-"
                    prevWasDash = True
                    dashCount = dashCount + 1
                End If
            Case Else
                out = out & ch
                prevWasDash = False
        End Select
    Next i
    Do While Left$(out, 1) = "-"
        out = Mid$(out, 2)
    Loop
    Do While Right$(out, 1) = "-"
        out = Left$(out, Len(out) - 1)
    Loop
    '  ---- 2026-08-24: AP Finance sometimes enter the CFA accounting string without
    '     any separators, e.g. "1177504007PGRW0039G030" when the standard canonical form
    '     is "1177-504007-PGRW0039-G030" (4/6/8/5 -- four segments total).  The generic
    '     separator-collapse block above only re-normalises existing dashes; it cannot
    '     introduce dashes where there were none, so the no-separator form came through
    '     as ONE long token -> Split("-") is a single segment -> lowest-hierarchy
    '     segment-lookup failed even though every individual 4/6/8/5 part exists in
    '     Mapping MASTER.  Declared valid in transform_specs/ap_cfa_claims.yaml:
    '       source_value_examples:
    '         - "1177-504007-PGRW0039-G030"   (dashes -- works already)
    '         - "1177504007PGRW0039G030"      (no dashes -- fixed below)
    If dashCount = 0 And Len(out) > 0 Then
        Dim L As Long
        L = Len(out)
        If L = 23 Then
            ' Mask NNNN-NNNNNN-AAAAAAAA-AAAAA (4/6/8/5) - total 23 chars, 3 internal dashes
            out = Left$(out, 4) & "-" & Mid$(out, 5, 6) & "-" & Mid$(out, 11, 8) & "-" & Mid$(out, 19, 5)
        ElseIf L = 22 Then
            ' 22 chars total = 4/6/8/4 or 4/6/7/5; use digit/letter transitions to figure it out
            out = AutoSegmentNoDashCode(out, L)
        ElseIf L = 20 Then
            ' Second common mask: "1224407003000000000000" -> "1224-407003-00000000-0000" (4/6/8/4)
            ' All digits; split at fixed offsets.
            out = Left$(out, 4) & "-" & Mid$(out, 5, 6) & "-" & Mid$(out, 11, 8) & "-" & Mid$(out, 19, 4)
        Else
            out = AutoSegmentNoDashCode(out, L)
        End If
    End If
    NormalizeCostCodeSeparators = out
End Function

' For no-separator accounting strings of non-standard length: walk left to right,
' inserting a dash every time the character class (digit vs letter) changes or every
' 6/8 digits when remaining digits are all-numeric -- whichever comes first.  This is
' intentionally conservative: it will never collapse an existing dash, and when no
' class transitions exist for an all-digit code we fall back to 4/6/8/rest slicing.
Private Function AutoSegmentNoDashCode(ByVal raw As String, ByVal L As Long) As String
    If L < 4 Then
        AutoSegmentNoDashCode = raw
        Exit Function
    End If
    Dim out As String
    Dim i As Long
    Dim prevClass As Long
    Dim runLen As Long
    Dim curCh As String
    Dim curClass As Long
    out = Left$(raw, 1)
    prevClass = CharClassD(AscW(Mid$(raw, 1, 1)))
    runLen = 1
    For i = 2 To L
        curCh = Mid$(raw, i, 1)
        curClass = CharClassD(AscW(curCh))
        Dim transition As Boolean
        transition = False
        If curClass <> prevClass And curClass <> 0 And prevClass <> 0 Then
            transition = True
        End If
        Dim sliceAllDigit As Boolean
        sliceAllDigit = False
        If prevClass = 1 And Not transition Then
            If runLen = 4 Or runLen = 6 Or runLen = 8 Then
                sliceAllDigit = True
            End If
        End If
        If transition Or sliceAllDigit Then
            out = out & "-" & curCh
            runLen = 1
        Else
            out = out & curCh
            runLen = runLen + 1
        End If
        prevClass = curClass
    Next i
    AutoSegmentNoDashCode = out
End Function

' 0 = other/non-alnum, 1 = digit, 2 = ASCII letter.
Private Function CharClassD(ByVal codePt As Long) As Long
    If codePt >= 48 And codePt <= 57 Then
        CharClassD = 1
    ElseIf (codePt >= 65 And codePt <= 90) Or (codePt >= 97 And codePt <= 122) Then
        CharClassD = 2
    Else
        CharClassD = 0
    End If
End Function

Public Function RowIsTrulyEmpty( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal textCheckCols As Variant, _
    ByVal amountCheckCols As Variant _
) As Boolean
    Dim i As Long
    Dim col As Long

    If IsArray(textCheckCols) Then
        For i = LBound(textCheckCols) To UBound(textCheckCols)
            col = CLng(textCheckCols(i))
            If col > 0 Then
                If LenB(Trim$(CStr(tableData(rowIndex, col)))) <> 0 Then
                    RowIsTrulyEmpty = False
                    Exit Function
                End If
            End If
        Next i
    End If

    If IsArray(amountCheckCols) Then
        For i = LBound(amountCheckCols) To UBound(amountCheckCols)
            col = CLng(amountCheckCols(i))
            If col > 0 Then
                If ResolveAmount(tableData(rowIndex, col)) <> 0# Then
                    RowIsTrulyEmpty = False
                    Exit Function
                End If
            End If
        Next i
    End If

    RowIsTrulyEmpty = True
End Function

Public Function RowIsEmptyOrTrailer( _
    ByVal tableData As Variant, _
    ByVal rowIndex As Long, _
    ByVal textCheckCols As Variant, _
    ByVal amountCheckCols As Variant, _
    ByVal labelCheckCols As Variant _
) As Boolean
    If RowIsTrulyEmpty(tableData, rowIndex, textCheckCols, amountCheckCols) Then
        RowIsEmptyOrTrailer = True
        Exit Function
    End If

    Dim i As Long
    Dim col As Long
    Dim blankCount As Long
    Dim textTotal As Long
    Dim labelText As String
    Dim normalised As String

    textTotal = 0
    blankCount = 0
    If IsArray(textCheckCols) Then
        For i = LBound(textCheckCols) To UBound(textCheckCols)
            col = CLng(textCheckCols(i))
            If col > 0 Then
                textTotal = textTotal + 1
                If LenB(Trim$(CStr(tableData(rowIndex, col)))) = 0 Then
                    blankCount = blankCount + 1
                End If
            End If
        Next i
    End If

    If textTotal > 0 And blankCount = textTotal Then
        labelText = vbNullString
        If IsArray(labelCheckCols) Then
            For i = LBound(labelCheckCols) To UBound(labelCheckCols)
                col = CLng(labelCheckCols(i))
                If col > 0 Then
                    labelText = labelText & " " & UCase$(Trim$(CStr(tableData(rowIndex, col))))
                End If
            Next i
        End If
        labelText = Trim$(labelText)
        If LenB(labelText) > 0 Then
            normalised = NormaliseToWordBoundaries(" " & labelText & " ")
            If (InStr(normalised, " TOTAL ") > 0) Or (InStr(normalised, " GRAND ") > 0) Or _
               (InStr(normalised, " SUBTOTAL ") > 0) Or (InStr(normalised, " SUB TOTAL ") > 0) Or _
               (InStr(normalised, " SUB-TOTAL ") > 0) Or (InStr(normalised, " SUM OF ") > 0) Or _
               (InStr(normalised, " SUM: ") > 0) Then
                RowIsEmptyOrTrailer = True
                Exit Function
            End If
        End If
    End If

    RowIsEmptyOrTrailer = False
End Function

Public Function ResolveEffectiveNet( _
    ByVal netValue As Variant, _
    ByVal grossValue As Variant, _
    ByVal taxValue As Variant _
) As Double
    Dim netAmt As Double
    Dim grossAmt As Double
    Dim taxAmt As Double
    netAmt = ResolveAmount(netValue)
    grossAmt = ResolveAmount(grossValue)
    taxAmt = ResolveAmount(taxValue)
    If netAmt = 0# Then netAmt = grossAmt - taxAmt
    ResolveEffectiveNet = netAmt
End Function
