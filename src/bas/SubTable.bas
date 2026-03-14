Option Explicit

' =============================================================================
' MODULE: SubTable
' Purpose: Table-specific operations (formulas, borders, margins)
'          Requires cursor to be inside a table.
'
' Contents:
'   - SelSumColumn / SelAverageColumn / SelCountColumn
'   - SelTableBorder
'   - SelTableMargin
'   - DocTableMargin
' =============================================================================


' ===== FORMULAS ==============================================================

Public Sub SelSumColumn()
    InsertTableFormula "SUM"
End Sub

Public Sub SelAverageColumn()
    InsertTableFormula "AVERAGE"
End Sub

Public Sub SelCountColumn()
    InsertTableFormula "COUNT"
End Sub

Private Sub InsertTableFormula(funcName As String)

    Dim cel As Cell
    Dim rng As Range
    Dim fld As Field
    Dim tbl As Table
    Dim col As Long
    Dim targetRow As Long
    Dim cellText As String
    Dim val As Double
    Dim total As Double
    Dim cnt As Long
    Dim fldResult As String

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place your cursor in a table cell.", vbExclamation
        Exit Sub
    End If

    Set cel = Selection.Cells(1)

    ' If cell already has a field, just refresh it
    If cel.Range.Fields.Count > 0 Then
        For Each fld In cel.Range.Fields
            fld.Update
        Next fld
        Exit Sub
    End If

    ' Get position
    Set tbl = cel.Range.Tables(1)
    targetRow = cel.RowIndex
    col = cel.ColumnIndex

    ' Manually calculate by reading cells above
    total = 0
    cnt = 0

    Dim r As Long
    For r = 1 To targetRow - 1
        On Error Resume Next
        Set rng = tbl.Cell(r, col).Range
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextRow

        rng.End = rng.End - 1
        cellText = Trim(rng.Text)

        ' Strip formatting
        cellText = CleanNumericText(cellText)

        ' Also read formula field results
        If tbl.Cell(r, col).Range.Fields.Count > 0 Then
            fldResult = tbl.Cell(r, col).Range.Fields(1).Result.Text
            fldResult = CleanNumericText(fldResult)
            If IsNumeric(fldResult) Then
                cellText = fldResult
            End If
        End If

        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            total = total + val
            cnt = cnt + 1
        End If

NextRow:
        Set rng = Nothing
    Next r

    ' Calculate result
    Dim finalVal As Double
    Select Case UCase(funcName)
        Case "SUM"
            finalVal = total
        Case "AVERAGE"
            If cnt > 0 Then
                finalVal = total / cnt
            Else
                finalVal = 0
            End If
        Case "COUNT"
            finalVal = cnt
    End Select

    ' Clear cell and write result
    Set rng = cel.Range
    rng.End = rng.End - 1
    rng.Delete

    Set rng = cel.Range
    rng.End = rng.End - 1
    rng.Collapse wdCollapseStart

    ' Insert as field so it can be formatted later
    Set fld = rng.Fields.Add( _
        Range:=rng, _
        Type:=wdFieldEmpty, _
        Text:="= " & Format(finalVal, "0.00"), _
        PreserveFormatting:=False)

    fld.Update

End Sub


' ===== BORDERS ===============================================================

Sub SelTableBorder()

    Dim tbl As Table
    Dim bTypes As Variant
    Dim i As Long

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)

    bTypes = Array(wdBorderLeft, wdBorderRight, wdBorderTop, _
                   wdBorderBottom, wdBorderHorizontal, wdBorderVertical)

    For i = LBound(bTypes) To UBound(bTypes)
        With tbl.Borders(bTypes(i))
            .LineStyle = wdLineStyleSingle
            .Color = wdColorAutomatic
            .LineWidth = wdLineWidth025pt
        End With
    Next i

    tbl.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    tbl.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone

End Sub


' ===== MARGINS — SELECTED TABLE ==============================================

Sub SelTableMargin()

    Dim tbl As Table

    Const PAD_TOP_CM As Double = 0.05
    Const PAD_BOTTOM_CM As Double = 0.05
    Const PAD_LEFT_CM As Double = 0.19
    Const PAD_RIGHT_CM As Double = 0.19

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Place the cursor inside the table you want to format.", vbExclamation
        Exit Sub
    End If

    If Selection.Cells.Count > 0 Then
        Set tbl = Selection.Cells(1).Range.Tables(1)
    ElseIf Selection.Range.Tables.Count > 0 Then
        Set tbl = Selection.Range.Tables(Selection.Range.Tables.Count)
    Else
        MsgBox "Couldn't resolve the table from the selection.", vbExclamation
        Exit Sub
    End If

    With tbl
        .TopPadding = CentimetersToPoints(PAD_TOP_CM)
        .BottomPadding = CentimetersToPoints(PAD_BOTTOM_CM)
        .LeftPadding = CentimetersToPoints(PAD_LEFT_CM)
        .RightPadding = CentimetersToPoints(PAD_RIGHT_CM)
    End With

End Sub


' ===== MARGINS — ALL TABLES IN DOCUMENT ======================================

Sub DocTableMargin()

    Dim tbl As Table

    Application.ScreenUpdating = False

    For Each tbl In ActiveDocument.Tables
        With tbl
            .TopPadding = CentimetersToPoints(0.1)
            .BottomPadding = CentimetersToPoints(0.1)
            .LeftPadding = CentimetersToPoints(0.19)
            .RightPadding = CentimetersToPoints(0.19)
        End With
    Next tbl

    Application.ScreenUpdating = True

End Sub


' ===== HELPER ================================================================

Private Function CleanNumericText(s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", "")
    t = Replace(t, "$", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, Chr(7), "")  ' Word table cell end marker
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If
    CleanNumericText = Trim$(t)
End Function
