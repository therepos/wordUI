Option Explicit

' =============================================================================
' MODULE: Subs
' Purpose: General formatting operations that work anywhere in the document
'          (body text, text boxes, headers, footers, etc.)
'
' Contents:
'   - DocFontSizeDecrease / DocFontSizeIncrease
'   - DocSpacingSingle
'   - SelListAlphaRoman
'   - SelFormatNumDecimal / SelFormatNumNoDecimal / SelFormatNumDollar
'   - SelFormatNumRepeat  (splitButton — repeats last-used number format)
'   - SelFormatDateShort / SelFormatDateLong
'   - SelFormatDateRepeat (splitButton — repeats last-used date format)
'   - ViewSplitVerticalToggle
' =============================================================================


' ===== FONT SIZE =============================================================

Sub DocFontSizeDecrease()

    Dim para As Paragraph
    Dim rng As Range

    Application.ScreenUpdating = False
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        If rng.Font.Size > 1 Then
            rng.Font.Size = rng.Font.Size - 1
        End If
    Next para
    Application.ScreenUpdating = True

End Sub

Sub DocFontSizeIncrease()

    Dim para As Paragraph
    Dim rng As Range

    Application.ScreenUpdating = False
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        rng.Font.Size = rng.Font.Size + 1
    Next para
    Application.ScreenUpdating = True

End Sub


' ===== SPACING ===============================================================

Sub DocSpacingSingle()

    Dim sr As Range
    Dim storyType As Variant
    Dim storyTypes As Variant

    Application.ScreenUpdating = False
    storyTypes = Array(wdMainTextStory, wdFootnotesStory, wdEndnotesStory, _
                       wdPrimaryHeaderStory, wdPrimaryFooterStory, _
                       wdFirstPageHeaderStory, wdFirstPageFooterStory, _
                       wdEvenPagesHeaderStory, wdEvenPagesFooterStory, _
                       wdTextFrameStory)

    For Each storyType In storyTypes
        Set sr = Nothing
        On Error Resume Next
        Set sr = ActiveDocument.StoryRanges(storyType)
        On Error GoTo 0

        If Not sr Is Nothing Then
            With sr.ParagraphFormat
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
            End With
        End If
    Next storyType
    Application.ScreenUpdating = True

End Sub


' ===== LIST — ALPHA / ROMAN ==================================================

Sub SelListAlphaRoman()

    Dim lt As ListTemplate
    Dim rng As Range
    Dim startLevel As Long
    Dim prevPara As Paragraph

    Set rng = Selection.Range
    startLevel = 0
    If rng.ListFormat.ListType <> wdListNoNumbering Then
        startLevel = rng.ListFormat.ListLevelNumber
        rng.ListFormat.RemoveNumbers
    Else
        If rng.Paragraphs(1).Range.Start > 0 Then
            Set prevPara = rng.Paragraphs(1).Previous
            If Not prevPara Is Nothing Then
                If prevPara.Range.ListFormat.ListType <> wdListNoNumbering Then
                    startLevel = prevPara.Range.ListFormat.ListLevelNumber + 1
                End If
            End If
        End If
    End If

    If startLevel < 1 Then startLevel = 1
    If startLevel > 8 Then startLevel = 8

    Set lt = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)

    With lt.ListLevels(startLevel)
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberFormat = "(%" & startLevel & ")"
        .TrailingCharacter = wdTrailingTab
        .NumberPosition = CentimetersToPoints(0.63 * (startLevel - 1))
        .TextPosition = CentimetersToPoints(0.63 * startLevel)
        .TabPosition = wdUndefined
        .ResetOnHigher = startLevel - 1
        .StartAt = 1
    End With

    If startLevel + 1 <= 9 Then
        With lt.ListLevels(startLevel + 1)
            .NumberStyle = wdListNumberStyleLowercaseRoman
            .NumberFormat = "(%" & (startLevel + 1) & ")"
            .TrailingCharacter = wdTrailingTab
            .NumberPosition = CentimetersToPoints(0.63 * startLevel)
            .TextPosition = CentimetersToPoints(0.63 * (startLevel + 1))
            .TabPosition = wdUndefined
            .ResetOnHigher = startLevel
            .StartAt = 1
        End With
    End If

    rng.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=lt, _
        ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToSelection, _
        DefaultListBehavior:=wdWord10ListBehavior

End Sub


' ===== NUMBER FORMATTING — WORKS ON ANY SELECTED TEXT ========================

Public Sub SelFormatNumNoDecimal()
    SavePref "LastNumFmt", "#,##0"
    SavePref "LastNumPrefix", ""
    FormatSelectedNumbers "#,##0", ""
End Sub

Public Sub SelFormatNumDecimal()
    SavePref "LastNumFmt", "#,##0.00"
    SavePref "LastNumPrefix", ""
    FormatSelectedNumbers "#,##0.00", ""
End Sub

Public Sub SelFormatNumDollar()
    SavePref "LastNumFmt", "#,##0.00"
    SavePref "LastNumPrefix", "$"
    FormatSelectedNumbers "#,##0.00", "$"
End Sub

Public Sub SelFormatNumRepeat()
    FormatSelectedNumbers _
        GetPref("LastNumFmt", "#,##0.00"), _
        GetPref("LastNumPrefix", "")
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
        If Selection.Cells.Count > 0 Then
            Dim cel As Cell
            For Each cel In Selection.Cells
                Set rng = cel.Range
                rng.End = rng.End - 1

                If rng.Fields.Count > 0 Then
                    For Each fld In rng.Fields
                        code = fld.code.Text
                        If InStr(code, "\#") > 0 Then
                            code = Left(code, InStr(code, "\#") - 1)
                        End If
                        If InStr(code, "\*") > 0 Then
                            code = Left(code, InStr(code, "\*") - 1)
                        End If
                        fld.code.Text = Trim(code) & fieldFormat
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
    If rng.Fields.Count > 0 Then
        For Each fld In rng.Fields
            code = fld.code.Text
            If InStr(code, "\#") > 0 Then
                code = Left(code, InStr(code, "\#") - 1)
            End If
            If InStr(code, "\*") > 0 Then
                code = Left(code, InStr(code, "\*") - 1)
            End If
            fld.code.Text = Trim(code) & fieldFormat
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

    cellText = CleanNumericText(rng.Text)

    If IsNumeric(cellText) And Len(cellText) > 0 Then
        val = CDbl(cellText)
        rng.Text = FormatValue(val, fmt, prefix)
        rng.Font.Color = wdColorAutomatic
        rng.ParagraphFormat.Alignment = wdAlignParagraphRight
    End If

End Sub


' ===== DATE FORMATTING — WORKS ON ANY SELECTED TEXT ==========================

Public Sub SelFormatDateShort()
    SavePref "LastDateFmt", "DD-MMM-YY"
    FormatSelectedDates "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    SavePref "LastDateFmt", "DD-MMMM-YYYY"
    FormatSelectedDates "DD-MMMM-YYYY"
End Sub

Public Sub SelFormatDateRepeat()
    FormatSelectedDates GetPref("LastDateFmt", "DD-MMM-YY")
End Sub

Private Sub FormatSelectedDates(fmt As String)

    Dim rng As Range
    Set rng = Selection.Range

    ' --- Case 1: Inside a table with cells selected ---
    If Selection.Information(wdWithInTable) Then
        If Selection.Cells.Count > 0 Then
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
    cellText = Trim$(rng.Text)

    If IsDate(cellText) Then
        rng.Text = Format(CDate(cellText), fmt)
    End If

End Sub


' ===== VIEW ==================================================================

Sub ViewSplitVerticalToggle()

    Dim doc As Document
    Set doc = ActiveDocument

    If doc.Windows.Count > 1 Then
        Do While doc.Windows.Count > 1
            doc.Windows(doc.Windows.Count).Close
        Loop
        doc.Windows(1).Activate
        doc.Windows(1).WindowState = wdWindowStateMaximize
        Exit Sub
    End If

    Dim Win1 As Long
    Dim Win2 As Long
    Dim halfW As Long
    Dim fullH As Long
    Dim overlap As Long

    ActiveWindow.WindowState = wdWindowStateMaximize

    halfW = Application.UsableWidth / 2
    fullH = Application.UsableHeight
    overlap = 8

    Win1 = ActiveWindow.Index

    ActiveWindow.NewWindow
    Win2 = ActiveWindow.Index

    With Windows(Win1)
        .WindowState = wdWindowStateNormal
        .Left = 0
        .Top = 0
        .Width = halfW + overlap
        .Height = fullH
    End With

    With Windows(Win2)
        .WindowState = wdWindowStateNormal
        .Left = halfW - overlap
        .Top = 0
        .Width = halfW + overlap
        .Height = fullH
    End With

    Windows(Win1).Activate

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


' ===== PREFERENCE STORAGE ====================================================
' Uses Windows registry (SaveSetting/GetSetting) so preferences persist
' across all documents for the user.

Private Sub SavePref(key As String, val As String)
    SaveSetting "WordUI", "Preferences", key, val
End Sub

Private Function GetPref(key As String, defaultVal As String) As String
    GetPref = GetSetting("WordUI", "Preferences", key, defaultVal)
End Function
