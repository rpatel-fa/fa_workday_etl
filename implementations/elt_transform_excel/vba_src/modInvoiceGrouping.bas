Attribute VB_Name = "modInvoiceGrouping"
Option Explicit

' Shared invoice-line grouping for EIB population (single-line = the N=1 degenerate case).
'
' Every supplier/customer invoice transform should treat its source as MULTI-LINE: rows that
' share one source invoice number collapse to ONE Workday invoice header (Spreadsheet Key), each
' carrying a 1..N Line Order / Row ID for its line within that header. Assigning a per-ROW
' Spreadsheet Key instead makes Workday read N separate headers that share one external invoice
' reference and reject the load (the duplicate-reference bug). This module owns the increment
' logic so transforms do not each re-implement it. Plain-VBA Collections only (no ActiveX /
' Scripting.Dictionary), so it behaves identically on Mac and Windows Excel.
'
' Usage: the caller keeps two Collections plus a running key counter, and calls this once per
' line it is about to WRITE, in source-file order, for the rows that passed its own skip logic:
'
'     Dim grpKey As New Collection, grpLine As New Collection, grpNextKey As Long
'     ...
'     Dim lineOrder As Long, spreadsheetKey As Long
'     spreadsheetKey = AssignInvoiceLineGroup(grpKey, grpLine, invoiceNumber, grpNextKey, lineOrder)
'     ' write Spreadsheet Key (B) = spreadsheetKey; Line Order (EM) = Row ID (EK) = lineOrder
'
' keyMap       : invoiceNumber -> assigned Spreadsheet Key (Long). Caller-owned, ByRef.
' lineOrderMap : invoiceNumber -> NEXT line order to hand out (Long). Caller-owned, ByRef.
' nextKey      : running maximum Spreadsheet Key assigned so far. Caller-owned, ByRef.
' lineOrder    : OUT -- this line's 1-based position within its invoice group.
' Returns the Spreadsheet Key (B) this line belongs to.
Public Function AssignInvoiceLineGroup( _
    ByRef keyMap As Collection, _
    ByRef lineOrderMap As Collection, _
    ByVal invoiceNumber As String, _
    ByRef nextKey As Long, _
    ByRef lineOrder As Long) As Long

    If LenB(invoiceNumber) = 0 Then
        ' No invoice number to group on: treat the row as its own single line. Defensive only --
        ' a blank invoice number should already have been rejected by the transform's skip logic;
        ' this just avoids a Collection empty-key error and never silently merges blank-key rows.
        nextKey = nextKey + 1
        lineOrder = 1
        AssignInvoiceLineGroup = nextKey
        Exit Function
    End If

    ' Uses the shared keyed-Collection primitives in modHelpers (one implementation for the whole
    ' project) -- CollectionKeyExists / CollectionSetItemLong / CollectionGetItemLong.
    If CollectionKeyExists(keyMap, invoiceNumber) = False Then
        nextKey = nextKey + 1
        CollectionSetItemLong keyMap, invoiceNumber, nextKey
        CollectionSetItemLong lineOrderMap, invoiceNumber, 1
    End If

    AssignInvoiceLineGroup = CollectionGetItemLong(keyMap, invoiceNumber)
    lineOrder = CollectionGetItemLong(lineOrderMap, invoiceNumber)
    CollectionSetItemLong lineOrderMap, invoiceNumber, lineOrder + 1
End Function
