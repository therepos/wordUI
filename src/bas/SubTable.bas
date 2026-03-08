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
    FormatNumber "#,##0.00", ""
End Sub

Public Sub SelFormatNumNoDecimal()
    FormatNumber "#,##0", ""
End Sub

Public Sub SelFormatNumDollar()
    FormatNumber "#,##0.00", "$"
End Sub

Private Sub FormatNumber(fmt As String, prefix As String)

    Dim cel As Cell
    Dim rng As Range
    Dim fld As Field
    Dim val As Double
    Dim cellText As String
    Dim code As String
    Dim result As String
    Dim fieldFormat As String

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place your cursor in a table cell.", vbExclamation
        Exit Sub
    End If

    ' Build Word field format string with bracket negatives
    If Len(prefix) > 0 Then
        fieldFormat = " \# ""'" & prefix & " '" & fmt & ";'" & prefix & " '(" & fmt & ")"""
    Else
        fieldFormat = " \# """ & fmt & ";(" & fmt & ")"""
    End If

    For Each cel In Selection.Cells
        Set rng = cel.Range
        rng.End = rng.End - 1

        ' ---- FORMULA FIELDS ----
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

            Set rng = cel.Range
            rng.End = rng.End - 1
            rng.ParagraphFormat.Alignment = wdAlignParagraphRight

        ' ---- PLAIN TEXT ----
        Else
            cellText = rng.text
            cellText = Replace(cellText, ",", "")
            cellText = Replace(cellText, "$", "")
            cellText = Replace(cellText, vbTab, "")

            If InStr(cellText, "(") > 0 And InStr(cellText, ")") > 0 Then
                cellText = Replace(cellText, "(", "-")
                cellText = Replace(cellText, ")", "")
            End If

            cellText = Trim(cellText)

            If IsNumeric(cellText) And Len(cellText) > 0 Then
                val = CDbl(cellText)

                If val < 0 Then
                    result = "(" & Format(Abs(val), fmt) & ")"
                Else
                    result = Format(val, fmt)
                End If

                If Len(prefix) > 0 Then
                    result = prefix & " " & result
                End If

                rng.text = result
                Set rng = cel.Range
                rng.End = rng.End - 1
                rng.Font.Color = wdColorAutomatic
                rng.ParagraphFormat.Alignment = wdAlignParagraphRight
            End If
        End If
    Next cel

End Sub

' ===== DATE FORMATTING =======================================================

Public Sub SelFormatDateShort()
    FormatDate "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    FormatDate "DD-MMMM-YYYY"
End Sub

Private Sub FormatDate(fmt As String)

    Dim cel As Cell
    Dim rng As Range
    Dim dt As Date
    Dim cellText As String

    ' Works both inside and outside tables
    If Not Selection.Information(wdWithInTable) Then
        If Len(Selection.text) > 0 Then
            cellText = Trim(Selection.text)
            If IsDate(cellText) Then
                dt = CDate(cellText)
                Selection.text = Format(dt, fmt)
            End If
        End If
        Exit Sub
    End If

    For Each cel In Selection.Cells
        Set rng = cel.Range
        rng.End = rng.End - 1
        cellText = Trim(rng.text)

        If IsDate(cellText) Then
            dt = CDate(cellText)
            rng.text = Format(dt, fmt)
        End If
    Next cel

End Sub

