Option Explicit

' =============================================================================
' MODULE: SubTable
' Purpose: Table cell operations - formulas, number formatting, date formatting
'          Number and date formatting works on ANY selected text (table cells,
'          text boxes, headers, footers, body text, etc.)
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
    cnt = 0

    Dim r As Long
    For r = 1 To targetRow - 1
        On Error Resume Next
        Set rng = tbl.Cell(r, col).Range
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextRow

        rng.End = rng.End - 1
        cellText = Trim(rng.text)

        ' Strip formatting: commas, dollar signs, brackets to minus
        cellText = CleanNumericText(cellText)

        ' Also read formula field results
        If tbl.Cell(r, col).Range.Fields.count > 0 Then
            fldResult = tbl.Cell(r, col).Range.Fields(1).result.text
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
        text:="= " & Format(finalVal, "0.00"), _
        PreserveFormatting:=False)

    fld.Update

End Sub

' ===== NUMBER FORMATTING — WORKS ON ANY SELECTED TEXT ========================

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

    Dim rng As Range
    Dim fld As Field
    Dim code As String
    Dim fieldFormat As String

    Set rng = Selection.Range

    ' Build Word field format string with bracket negatives
    If Len(prefix) > 0 Then
        fieldFormat = " \# ""'" & prefix & " '" & fmt & ";'" & prefix & " '(" & fmt & ")"""
    Else
        fieldFormat = " \# """ & fmt & ";(" & fmt & ")"""
    End If

    ' --- Case 1: Inside a table with cells selected ---
    If Selection.Information(wdWithInTable) Then
        If Selection.Cells.count > 0 Then
            Dim cel As Cell
            For Each cel In Selection.Cells
                Set rng = cel.Range
                rng.End = rng.End - 1

                If rng.Fields.count > 0 Then
                    ' Update field format switches
                    For Each fld In rng.Fields
                        code = fld.code.text
                        If InStr(code, "\#") > 0 Then
                            code = Left(code, InStr(code, "\#") - 1)
                        End If
                        If InStr(code, "\*") > 0 Then
                            code = Left(code, InStr(code, "\*") - 1)
                        End If
                        fld.code.text = Trim(code) & fieldFormat
                        fld.Update
                    Next fld
                    Set rng = cel.Range
                    rng.End = rng.End - 1
                    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
                Else
                    FormatNumInRange rng, fmt, prefix
                End If
            Next cel
            Exit Sub
        End If
    End If

    ' --- Case 2: Any other text selection (body, text box, header, etc.) ---
    If rng.Fields.count > 0 Then
        For Each fld In rng.Fields
            code = fld.code.text
            If InStr(code, "\#") > 0 Then
                code = Left(code, InStr(code, "\#") - 1)
            End If
            If InStr(code, "\*") > 0 Then
                code = Left(code, InStr(code, "\*") - 1)
            End If
            fld.code.text = Trim(code) & fieldFormat
            fld.Update
        Next fld
        rng.ParagraphFormat.Alignment = wdAlignParagraphRight
    Else
        FormatNumInRange rng, fmt, prefix
    End If

End Sub

Private Sub FormatNumInRange(rng As Range, fmt As String, prefix As String)

    Dim cellText As String
    Dim val As Double

    cellText = CleanNumericText(rng.text)

    If IsNumeric(cellText) And Len(cellText) > 0 Then
        val = CDbl(cellText)
        rng.text = FormatValue(val, fmt, prefix)
        rng.Font.Color = wdColorAutomatic
        rng.ParagraphFormat.Alignment = wdAlignParagraphRight
    End If

End Sub

' ===== DATE FORMATTING — WORKS ON ANY SELECTED TEXT ==========================

Public Sub SelFormatDateShort()
    FormatSelectedDates "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    FormatSelectedDates "DD-MMMM-YYYY"
End Sub

Private Sub FormatSelectedDates(fmt As String)

    Dim rng As Range
    Set rng = Selection.Range

    ' --- Case 1: Inside a table with cells selected ---
    If Selection.Information(wdWithInTable) Then
        If Selection.Cells.count > 0 Then
            Dim cel As Cell
            For Each cel In Selection.Cells
                Set rng = cel.Range
                rng.End = rng.End - 1
                FormatDateInRange rng, fmt
            Next cel
            Exit Sub
        End If
    End If

    ' --- Case 2: Any other text selection ---
    FormatDateInRange rng, fmt

End Sub

Private Sub FormatDateInRange(rng As Range, fmt As String)

    Dim cellText As String
    cellText = Trim$(rng.text)

    If IsDate(cellText) Then
        rng.text = Format(CDate(cellText), fmt)
    End If

End Sub

' ===== HELPERS ===============================================================

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


