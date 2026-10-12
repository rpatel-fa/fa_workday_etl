Attribute VB_Name = "modEntityBlockWriter"
Option Explicit

' Shared "one entity -> block of EIB rows" writer (registry concern entity_row_block, variant
' item_index_block). Master-data EIBs (Customer, Supplier, ...) carry ONE entity as several
' consecutive rows that share a Spreadsheet Key (column B): row 1 holds the single-valued fields
' plus item 1 of every multi-value list / child area, and row k (k > 1) holds item k of each list
' that has at least k items. This module owns that layout so transforms do not each re-implement
' it. Generic by design: it knows nothing about customers or suppliers -- the transform supplies
' the column letters (pinned in its profile YAML column map), item indexes and values.
'
' Usage, once per entity:
'
'     EntityBlockReset
'     EntityBlockAddItem "BI", 1, customerName               ' single-valued field = item 1
'     EntityBlockAddItem "CR", 1, "BILLING"                  ' list "Use For": item 1
'     EntityBlockAddItem "CR", 2, "SHIPPING"                 ' list "Use For": item 2 -> row 2
'     EntityBlockAddAncestor "BL", "1", "CR,CS"              ' parent Row ID repeated on rows where CR/CS hold data
'     nextRow = EntityBlockFlush(ws, nextRow, spreadsheetKey)
'
' Plain VBA arrays only (no ActiveX / Scripting.Dictionary): identical on Mac and Windows Excel.
' Status: code-complete, pending manual Excel import (not run in Excel yet).

Private Const BLOCK_KEY_COLUMN_LETTERS As String = "B"
Private Const BLOCK_INITIAL_CAPACITY As Long = 64

Private mItemCount As Long
Private mItemColumn() As Long
Private mItemIndex() As Long
Private mItemValue() As Variant

Private mAncestorCount As Long
Private mAncestorColumn() As Long
Private mAncestorValue() As Variant
Private mAncestorAreaColumns() As String

' Clears the per-entity buffers. Call once before collecting each entity.
Public Sub EntityBlockReset()
    mItemCount = 0
    mAncestorCount = 0
    ReDim mItemColumn(1 To BLOCK_INITIAL_CAPACITY)
    ReDim mItemIndex(1 To BLOCK_INITIAL_CAPACITY)
    ReDim mItemValue(1 To BLOCK_INITIAL_CAPACITY)
    ReDim mAncestorColumn(1 To BLOCK_INITIAL_CAPACITY)
    ReDim mAncestorValue(1 To BLOCK_INITIAL_CAPACITY)
    ReDim mAncestorAreaColumns(1 To BLOCK_INITIAL_CAPACITY)
End Sub

' Records value as item itemIndex (1-based) of column colLetter. Blank values and itemIndex < 1
' are ignored, so a transform can call this unconditionally for optional fields.
Public Sub EntityBlockAddItem(ByVal colLetter As String, ByVal itemIndex As Long, ByVal itemValue As Variant)
    If itemIndex < 1 Then Exit Sub
    If IsEmpty(itemValue) Then Exit Sub
    If LenB(Trim$(CStr(itemValue))) = 0 Then Exit Sub

    If mItemCount = 0 Then
        If Not IsBlockInitialised() Then EntityBlockReset
    End If
    If mItemCount >= UBound(mItemColumn) Then
        ReDim Preserve mItemColumn(1 To mItemCount * 2)
        ReDim Preserve mItemIndex(1 To mItemCount * 2)
        ReDim Preserve mItemValue(1 To mItemCount * 2)
    End If

    mItemCount = mItemCount + 1
    mItemColumn(mItemCount) = ColumnIndex(colLetter)
    mItemIndex(mItemCount) = itemIndex
    mItemValue(mItemCount) = itemValue
End Sub

' Registers a parent-area Row ID column whose value must be repeated on every continuation row
' (row k > 1) on which at least one of the area's item columns holds data. areaColLetters is a
' comma-separated list of item column letters, e.g. "CR,CS".
Public Sub EntityBlockAddAncestor(ByVal colLetter As String, ByVal ancestorValue As Variant, ByVal areaColLetters As String)
    Dim areaParts() As String
    Dim partIndex As Long
    Dim areaList As String

    If IsEmpty(ancestorValue) Then Exit Sub
    If LenB(Trim$(CStr(ancestorValue))) = 0 Then Exit Sub

    If mAncestorCount = 0 Then
        If Not IsBlockInitialised() Then EntityBlockReset
    End If
    If mAncestorCount >= UBound(mAncestorColumn) Then
        ReDim Preserve mAncestorColumn(1 To mAncestorCount * 2)
        ReDim Preserve mAncestorValue(1 To mAncestorCount * 2)
        ReDim Preserve mAncestorAreaColumns(1 To mAncestorCount * 2)
    End If

    ' Normalise the area list to ",<index>,<index>," so membership is a plain InStr test.
    areaParts = Split(areaColLetters, ",")
    areaList = ","
    For partIndex = LBound(areaParts) To UBound(areaParts)
        If LenB(Trim$(areaParts(partIndex))) > 0 Then
            areaList = areaList & CStr(ColumnIndex(Trim$(areaParts(partIndex)))) & ","
        End If
    Next partIndex

    mAncestorCount = mAncestorCount + 1
    mAncestorColumn(mAncestorCount) = ColumnIndex(colLetter)
    mAncestorValue(mAncestorCount) = ancestorValue
    mAncestorAreaColumns(mAncestorCount) = areaList
End Sub

' Writes the collected entity as a block starting at startRow and returns the next free row.
' Block height = the largest item index added (0 items -> nothing written, returns startRow).
' Spreadsheet Key goes in column B of every block row.
Public Function EntityBlockFlush(ByVal ws As Worksheet, ByVal startRow As Long, ByVal spreadsheetKey As Variant) As Long
    Dim blockHeight As Long
    Dim itemPos As Long
    Dim ancestorPos As Long
    Dim rowOffset As Long
    Dim targetRow As Long
    Dim keyColumn As Long

    For itemPos = 1 To mItemCount
        If mItemIndex(itemPos) > blockHeight Then blockHeight = mItemIndex(itemPos)
    Next itemPos

    If blockHeight = 0 Then
        EntityBlockFlush = startRow
        Exit Function
    End If

    keyColumn = ColumnIndex(BLOCK_KEY_COLUMN_LETTERS)
    For rowOffset = 1 To blockHeight
        targetRow = startRow + rowOffset - 1
        ws.Cells(targetRow, keyColumn).Value = spreadsheetKey

        ' Ancestor Row IDs first, only on continuation rows that hold data for their area; item
        ' writes below then win if an ancestor column is also supplied as an item on this row.
        If rowOffset > 1 Then
            For ancestorPos = 1 To mAncestorCount
                If RowHasAreaData(rowOffset, mAncestorAreaColumns(ancestorPos)) Then
                    ws.Cells(targetRow, mAncestorColumn(ancestorPos)).Value = mAncestorValue(ancestorPos)
                End If
            Next ancestorPos
        End If

        For itemPos = 1 To mItemCount
            If mItemIndex(itemPos) = rowOffset Then
                ws.Cells(targetRow, mItemColumn(itemPos)).Value = mItemValue(itemPos)
            End If
        Next itemPos
    Next rowOffset

    EntityBlockFlush = startRow + blockHeight
End Function

Private Function IsBlockInitialised() As Boolean
    On Error Resume Next
    IsBlockInitialised = (UBound(mItemColumn) >= 1)
    If Err.Number <> 0 Then IsBlockInitialised = False
    On Error GoTo 0
End Function

Private Function RowHasAreaData(ByVal rowOffset As Long, ByVal areaList As String) As Boolean
    Dim itemPos As Long

    For itemPos = 1 To mItemCount
        If mItemIndex(itemPos) = rowOffset Then
            If InStr(1, areaList, "," & CStr(mItemColumn(itemPos)) & ",", vbBinaryCompare) > 0 Then
                RowHasAreaData = True
                Exit Function
            End If
        End If
    Next itemPos
End Function
