Attribute VB_Name = "Subs"
Option Explicit

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

Sub DocTableMargin()

    Dim tbl As Table
    Dim topMargin As Single, bottomMargin As Single
    Dim leftMargin As Single, rightMargin As Single

    Application.ScreenUpdating = False
    topMargin = CentimetersToPoints(0.1)
    bottomMargin = CentimetersToPoints(0.1)
    leftMargin = CentimetersToPoints(0.19)
    rightMargin = CentimetersToPoints(0.19)

    For Each tbl In ActiveDocument.Tables
        With tbl
            .TopPadding = topMargin
            .BottomPadding = bottomMargin
            .LeftPadding = leftMargin
            .RightPadding = rightMargin
        End With
    Next tbl
    Application.ScreenUpdating = True

End Sub

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

Sub SelTableBorder()

    Dim tbl As Table
    Dim bTypes As Variant
    Dim i As Long

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)

    bTypes = Array(wdBorderLeft, wdBorderRight, wdBorderTop, wdBorderBottom, wdBorderHorizontal, wdBorderVertical)

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

Sub SelTableMargin()

    Dim tbl As Table

    Const PAD_TOP_CM As Double = 0.05
    Const PAD_BOTTOM_CM As Double = 0.05
    Const PAD_LEFT_CM As Double = 0.19
    Const PAD_RIGHT_CM As Double = 0.19

    If Selection.Information(wdWithInTable) Then

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

    Else
        MsgBox "Place the cursor inside the table you want to format.", vbExclamation
    End If

End Sub

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

