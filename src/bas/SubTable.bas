Attribute VB_Name = "SubTable"
Option Explicit

' =============================================================================
' MODULE: ModTable
' Purpose: Table cell operations - formulas, number formatting, date formatting
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
    Dim row As Long
    Dim targetRow As Long
    Dim cellText As String
    Dim val As Double
    Dim total As Double
    Dim count As Long
    Dim formulaText As String

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place your cursor in a table cell.", vbExclamation
        Exit Sub
    End If

    Set cel = Selection.Cells(1)

    ' If cell already has a field, just refresh it
    If cel.Range.Fields.count > 0 Then
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
    count = 0

    Dim r As Long
    For r = 1 To targetRow - 1
        On Error Resume Next
        Set rng = tbl.Cell(r, col).Range
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextRow

        rng.End = rng.End - 1
        cellText = Trim(rng.text)

        ' Strip formatting: commas, dollar signs, brackets to minus
        cellText = Replace(cellText, ",", "")
        cellText = Replace(cellText, "$", "")
        cellText = Replace(cellText, vbTab, "")
        If InStr(cellText, "(") > 0 And InStr(cellText, ")") > 0 Then
            cellText = Replace(cellText, "(", "-")
            cellText = Replace(cellText, ")", "")
        End If
        cellText = Trim(cellText)

        ' Also read formula field results
        If tbl.Cell(r, col).Range.Fields.count > 0 Then
            Dim fldResult As String
            fldResult = tbl.Cell(r, col).Range.Fields(1).result.text
            fldResult = Replace(fldResult, ",", "")
            fldResult = Replace(fldResult, "$", "")
            If InStr(fldResult, "(") > 0 And InStr(fldResult, ")") > 0 Then
                fldResult = Replace(fldResult, "(", "-")
                fldResult = Replace(fldResult, ")", "")
            End If
            fldResult = Trim(fldResult)
            If IsNumeric(fldResult) Then
                cellText = fldResult
            End If
        End If

        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            total = total + val
            count = count + 1
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
            If count > 0 Then
                finalVal = total / count
            Else
                finalVal = 0
            End If
        Case "COUNT"
            finalVal = count
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
        text:="= " & Format(finalVal, "0.00"), _
        PreserveFormatting:=False)

    fld.Update

End Sub

' ===== NUMBER FORMATTING =====================================================

Public Sub SelFormatNumDecimal()
    FormatSelectedNumbers "#,##0.00", ""
End Sub

Public Sub SelFormatNumNoDecimal()
    FormatSelectedNumbers "#,##0", ""
End Sub

Public Sub SelFormatNumDollar()
    FormatSelectedNumbers "#,##0.00", "$"
End Sub

Private Sub FormatSelectedNumbers(fmt As String, prefix As String)

    Dim sel As Selection
    Set sel = ActiveWindow.Selection

    ' --- Text is highlighted (works in text box, table cell, title, etc.) ---
    If sel.Type = ppSelectionText Then
        FormatNumInTextRange sel.TextRange, fmt, prefix
        Exit Sub
    End If

    ' --- Whole shape(s) selected — format all text inside ---
    If sel.Type = ppSelectionShapes Then
        Dim i As Long
        For i = 1 To sel.ShapeRange.Count
            FormatNumInShape sel.ShapeRange(i), fmt, prefix
        Next i
        Exit Sub
    End If

    MsgBox "Please select text or a shape containing numbers.", vbExclamation, "Number Format"

End Sub

Private Sub FormatNumInShape(shp As Shape, fmt As String, prefix As String)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            FormatNumInShape subShp, fmt, prefix
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                FormatNumInTextRange shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fmt, prefix
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            FormatNumInTextRange shp.TextFrame.TextRange, fmt, prefix
        End If
    End If

    On Error GoTo 0

End Sub

Private Sub FormatNumInTextRange(tr As TextRange, fmt As String, prefix As String)

    On Error Resume Next

    ' Try the whole range as a single number first
    Dim txt As String
    txt = CleanNumericText(tr.Text)

    If IsNumeric(txt) And Len(txt) > 0 Then
        Dim val As Double
        val = CDbl(txt)
        tr.Text = FormatValue(val, fmt, prefix)
        tr.ParagraphFormat.Alignment = ppAlignRight
        Exit Sub
    End If

    ' Otherwise try each paragraph individually
    Dim p As Long
    For p = tr.Paragraphs.Count To 1 Step -1
        Dim pTxt As String
        pTxt = CleanNumericText(tr.Paragraphs(p).Text)
        If IsNumeric(pTxt) And Len(pTxt) > 0 Then
            tr.Paragraphs(p).Text = FormatValue(CDbl(pTxt), fmt, prefix)
            tr.Paragraphs(p).ParagraphFormat.Alignment = ppAlignRight
        End If
    Next p

    On Error GoTo 0

End Sub

Private Function FormatValue(val As Double, fmt As String, prefix As String) As String
    Dim result As String
    If val < 0 Then
        result = "(" & Format(Abs(val), fmt) & ")"
    Else
        result = Format(val, fmt)
    End If
    If Len(prefix) > 0 Then
        result = prefix & " " & result
    End If
    FormatValue = result
End Function

' ===== DATE FORMATTING — WORKS ON ANY SELECTED TEXT ==========================

Public Sub SelFormatDateShort()
    FormatSelectedDates "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    FormatSelectedDates "DD-MMMM-YYYY"
End Sub

Private Sub FormatSelectedDates(fmt As String)

    Dim sel As Selection
    Set sel = ActiveWindow.Selection

    ' --- Text is highlighted ---
    If sel.Type = ppSelectionText Then
        FormatDateInTextRange sel.TextRange, fmt
        Exit Sub
    End If

    ' --- Whole shape(s) selected ---
    If sel.Type = ppSelectionShapes Then
        Dim i As Long
        For i = 1 To sel.ShapeRange.Count
            FormatDateInShape sel.ShapeRange(i), fmt
        Next i
        Exit Sub
    End If

    MsgBox "Please select text or a shape containing dates.", vbExclamation, "Date Format"

End Sub

Private Sub FormatDateInShape(shp As Shape, fmt As String)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            FormatDateInShape subShp, fmt
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                FormatDateInTextRange shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fmt
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            FormatDateInTextRange shp.TextFrame.TextRange, fmt
        End If
    End If

    On Error GoTo 0

End Sub

Private Sub FormatDateInTextRange(tr As TextRange, fmt As String)

    On Error Resume Next

    ' Try the whole range as a single date
    Dim txt As String
    txt = Trim$(tr.Text)

    If IsDate(txt) Then
        tr.Text = Format(CDate(txt), fmt)
        Exit Sub
    End If

    ' Otherwise try each paragraph
    Dim p As Long
    For p = tr.Paragraphs.Count To 1 Step -1
        Dim pTxt As String
        pTxt = Trim$(tr.Paragraphs(p).Text)
        If IsDate(pTxt) Then
            tr.Paragraphs(p).Text = Format(CDate(pTxt), fmt)
        End If
    Next p

    On Error GoTo 0

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
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If
    CleanNumericText = Trim$(t)
End Function