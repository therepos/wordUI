# VBA modules

_This file is generated automatically from `.bas` files in `src/bas`._

## Module `Ribbon`

### `RibbonOnLoad`

```vbnet
Public Sub RibbonOnLoad(r As IRibbonUI)
    Set Ribbon = r
End Sub
```

### `RunByName`

```vbnet
Public Sub RunByName(control As IRibbonControl)
    Dim macro As String
    macro = control.Tag
    If Len(macro) = 0 Then macro = control.ID
    On Error GoTo errh
    Application.Run macro
    Exit Sub
errh:
    MsgBox "Macro not found: " & macro, vbExclamation
End Sub
```

## Module `SubConvert`

### `ConvertUStoUK`

```vbnet
Public Sub ConvertUStoUK()

    ' Guard: document has no words
    If ActiveDocument.ComputeStatistics(wdStatisticWords) = 0 Then
        MsgBox "Document has no words.", vbInformation, "US to UK English"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "US to UK Conversion"

    m_count = 0
    m_totalReplaced = 0
    Erase m_us
    Erase m_uk

    ' ----- -ize / -ise -----
    R "recognize", "recognise"
    R "organize", "organise"
    R "realize", "realise"
    R "minimize", "minimise"
    R "maximize", "maximise"
    R "optimize", "optimise"
    R "utilize", "utilise"
    R "authorize", "authorise"
    R "categorize", "categorise"
    R "characterize", "characterise"
    R "customize", "customise"
    R "emphasize", "emphasise"
    R "finalize", "finalise"
    R "globalize", "globalise"
    R "harmonize", "harmonise"
    R "initialize", "initialise"
    R "legalize", "legalise"
    R "memorize", "memorise"
    R "modernize", "modernise"
    R "neutralize", "neutralise"
    R "normalize", "normalise"
    R "prioritize", "prioritise"
    R "specialize", "specialise"
    R "standardize", "standardise"
    R "summarize", "summarise"
    R "symbolize", "symbolise"
    R "synchronize", "synchronise"
    R "apologize", "apologise"
    R "capitalize", "capitalise"
    R "centralize", "centralise"
    R "criticize", "criticise"
    R "digitize", "digitise"
    R "dramatize", "dramatise"
    R "familiarize", "familiarise"
    R "fertilize", "fertilise"
    R "generalize", "generalise"
    R "hospitalize", "hospitalise"
    R "hypothesize", "hypothesise"
    R "idealize", "idealise"
    R "immunize", "immunise"
    R "itemize", "itemise"
    R "jeopardize", "jeopardise"
    R "liberalize", "liberalise"
    R "localize", "localise"
    R "marginalize", "marginalise"
    R "materialize", "materialise"
    R "mechanize", "mechanise"
    R "mobilize", "mobilise"
    R "monopolize", "monopolise"
    R "nationalize", "nationalise"
    R "penalize", "penalise"
    R "polarize", "polarise"
    R "privatize", "privatise"
    R "revolutionize", "revolutionise"
    R "scrutinize", "scrutinise"
    R "sensitize", "sensitise"
    R "socialize", "socialise"
    R "stabilize", "stabilise"
    R "sterilize", "sterilise"
    R "subsidize", "subsidise"
    R "terrorize", "terrorise"
    R "traumatize", "traumatise"
    R "trivialize", "trivialise"
    R "vandalize", "vandalise"
    R "vaporize", "vaporise"
    R "visualize", "visualise"

    ' ----- -or / -our -----
    R "color", "colour"
    R "favor", "favour"
    R "honor", "honour"
    R "humor", "humour"
    R "labor", "labour"
    R "neighbor", "neighbour"
    R "behavior", "behaviour"
    R "flavor", "flavour"
    R "harbor", "harbour"
    R "rumor", "rumour"
    R "tumor", "tumour"
    R "valor", "valour"
    R "vigor", "vigour"

    ' ----- -er / -re -----
    R "center", "centre"
    R "fiber", "fibre"
    R "liter", "litre"
    R "meter", "metre"
    R "theater", "theatre"

    ' ----- exact words -----
    X "aging", "ageing"
    X "airplane", "aeroplane"
    X "airplanes", "aeroplanes"
    X "aluminum", "aluminium"
    X "cozy", "cosy"
    X "gray", "grey"
    X "judgment", "judgement"
    X "math", "maths"
    X "program", "programme"
    X "programs", "programmes"
    X "check", "cheque"
    X "checks", "cheques"
    X "curb", "kerb"
    X "curbs", "kerbs"
    X "jewelry", "jewellery"
    X "skillful", "skilful"
    X "skillfully", "skilfully"

    RunAllReplacements

    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True

    If m_totalReplaced > 0 Then
        MsgBox m_totalReplaced & " word(s) converted. Ctrl+Z to undo.", vbInformation
    Else
        MsgBox "No US English words found.", vbInformation
    End If

End Sub
```

### `X`

```vbnet
Private Sub X(us As String, uk As String)

    m_count = m_count + 1
    ReDim Preserve m_us(1 To m_count)
    ReDim Preserve m_uk(1 To m_count)

    m_us(m_count) = us
    m_uk(m_count) = uk

End Sub
```

### `R`

```vbnet
Private Sub R(us As String, uk As String)

    If EndsWith(us, "ize") And EndsWith(uk, "ise") Then

        Dim stem1 As String: stem1 = Left(us, Len(us) - 3)
        Dim stem2 As String: stem2 = Left(uk, Len(uk) - 3)

        X us, uk
        X stem1 & "izes", stem2 & "ises"
        X stem1 & "ized", stem2 & "ised"
        X stem1 & "izing", stem2 & "ising"
        X stem1 & "izer", stem2 & "iser"
        X stem1 & "ization", stem2 & "isation"

        Exit Sub
    End If

    X us, uk

End Sub
```

### `RunAllReplacements`

```vbnet
Private Sub RunAllReplacements()

    Dim sr As Range
    Dim storyType As Variant
    Dim i As Long
    Dim storyTypes As Variant

    storyTypes = Array( _
        wdMainTextStory, _
        wdFootnotesStory, _
        wdEndnotesStory, _
        wdPrimaryHeaderStory, _
        wdPrimaryFooterStory, _
        wdFirstPageHeaderStory, _
        wdFirstPageFooterStory, _
        wdEvenPagesHeaderStory, _
        wdEvenPagesFooterStory, _
        wdTextFrameStory)

    For Each storyType In storyTypes

        On Error Resume Next
        Set sr = ActiveDocument.StoryRanges(storyType)
        On Error GoTo 0

        Do While Not sr Is Nothing

            If sr.StoryLength > 1 Then

                For i = 1 To m_count

                    With sr.Find

                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = m_us(i)
                        .Replacement.text = m_uk(i)

                        .Forward = True
                        .Wrap = wdFindStop
                        .MatchCase = False
                        .MatchWholeWord = True

                        If .Execute(Replace:=wdReplaceAll) Then
                            m_totalReplaced = m_totalReplaced + 1
                        End If

                    End With

                Next i

            End If

            Set sr = sr.NextStoryRange

        Loop

    Next storyType

End Sub
```

### `EndsWith`

```vbnet
Private Function EndsWith(s As String, suffix As String) As Boolean

    If Len(s) >= Len(suffix) Then
        EndsWith = (Right$(s, Len(suffix)) = suffix)
    End If

End Function
```

## Module `SubReset`

### `ResetAll`

```vbnet
Public Sub ResetAll()
    Application.ScreenUpdating = False
    ResetFormat
    ResetList
    ResetObject
    ResetTables
    ResetHyperlinks
    ResetStylesCustom
    ResetStylesDefault
    Application.ScreenUpdating = True
    MsgBox "Reset complete:" & vbCrLf & vbCrLf & _
           "Formatting, Lists, Objects, Tables, Hyperlinks, Styles (All)", _
           vbInformation, "Reset"
End Sub
```

### `RunResetFormat`

```vbnet
Public Sub RunResetFormat()
    Application.ScreenUpdating = False
    ResetFormat
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Formatting", vbInformation, "Reset"
End Sub
```

### `RunResetList`

```vbnet
Public Sub RunResetList()
    Application.ScreenUpdating = False
    ResetList
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Lists", vbInformation, "Reset"
End Sub
```

### `RunResetObject`

```vbnet
Public Sub RunResetObject()
    Application.ScreenUpdating = False
    ResetObject
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Objects", vbInformation, "Reset"
End Sub
```

### `RunResetTables`

```vbnet
Public Sub RunResetTables()
    Application.ScreenUpdating = False
    ResetTables
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Tables", vbInformation, "Reset"
End Sub
```

### `RunResetStylesAll`

```vbnet
Public Sub RunResetStylesAll()
    Application.ScreenUpdating = False
    ResetStylesCustom
    ResetStylesDefault
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Styles (All)", vbInformation, "Reset"
End Sub
```

### `RunResetStylesDefault`

```vbnet
Public Sub RunResetStylesDefault()
    Application.ScreenUpdating = False
    ResetStylesDefault
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Styles (Default)", vbInformation, "Reset"
End Sub
```

### `RunResetHyperlinks`

```vbnet
Public Sub RunResetHyperlinks()
    Application.ScreenUpdating = False
    ResetHyperlinks
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Hyperlinks", vbInformation, "Reset"
End Sub
```

### `ResetFormat`

```vbnet
Private Sub ResetFormat()
    ActiveDocument.Content.Select
    Selection.ClearFormatting
End Sub
```

### `ResetList`

```vbnet
Private Sub ResetList()
    Dim para As Paragraph
    Dim lt As WdListType
    Dim lvl As Long

    For Each para In ActiveDocument.Paragraphs
        lt = para.Range.ListFormat.ListType
        If lt <> wdListNoNumbering Then
            lvl = para.Range.ListFormat.ListLevelNumber
            para.Range.ListFormat.RemoveNumbers

            If lt = wdListBullet Then
                para.Range.ListFormat.ApplyBulletDefault
            Else
                para.Range.ListFormat.ApplyNumberDefault
            End If

            para.Range.ListFormat.ListLevelNumber = lvl
        End If
    Next para
End Sub
```

### `ResetObject`

```vbnet
Private Sub ResetObject()
    Dim shp As InlineShape

    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
End Sub
```

### `ResetTables`

```vbnet
Private Sub ResetTables()
    Dim tbl As Word.Table
    Dim s As Word.Style
    Dim tableNormal As Word.Style

    For Each s In ActiveDocument.Styles
        If s.BuiltIn Then
            If s.Type = wdStyleTypeTable Then
                If LCase$(s.NameLocal) Like "*normal*" Or _
                   LCase$(s.NameLocal) Like "*table normal*" Then
                    Set tableNormal = s
                    Exit For
                End If
            End If
        End If
    Next s

    For Each tbl In ActiveDocument.Tables
        On Error Resume Next
        If Not tableNormal Is Nothing Then
            tbl.Style = tableNormal
        Else
            tbl.Style = "Table Normal"
            If Err.Number <> 0 Then
                Err.Clear
                tbl.Style = "Table Grid"
            End If
        End If
        On Error GoTo 0

        With tbl
            .TopPadding = 0
            .BottomPadding = 0
            .LeftPadding = 0
            .RightPadding = 0
            .Borders.Enable = True
        End With
    Next tbl
End Sub
```

### `ResetStylesDefault`

```vbnet
Private Sub ResetStylesDefault()
    Dim docOrig As Object
    Dim docTemp As Object
    Dim sty As Object

    Set docOrig = ActiveDocument
    Set docTemp = Documents.Add(Visible:=False)

    On Error Resume Next
    For Each sty In docTemp.Styles
        If sty.BuiltIn Then
            docOrig.Styles(sty.NameLocal).Font = sty.Font
            docOrig.Styles(sty.NameLocal).ParagraphFormat = sty.ParagraphFormat
        End If
    Next sty
    On Error GoTo 0

    docTemp.Close SaveChanges:=False
End Sub
```

### `ResetStylesCustom`

```vbnet
Private Sub ResetStylesCustom()
    Dim oStyle As Style

    On Error Resume Next
    For Each oStyle In ActiveDocument.Styles
        If oStyle.BuiltIn = False Then
            oStyle.Delete
        End If
    Next oStyle
    On Error GoTo 0
End Sub
```

### `ResetHyperlinks`

```vbnet
Private Sub ResetHyperlinks()
    With ActiveDocument
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
End Sub
```

## Module `Subs`

### `DocFontArial`

```vbnet
Sub DocFontArial()

    Dim sr As Range
    Application.ScreenUpdating = False
    For Each sr In ActiveDocument.StoryRanges
        Do
            sr.Font.Name = "Arial"
            Set sr = sr.NextStoryRange
        Loop Until sr Is Nothing
    Next
    Application.ScreenUpdating = True

End Sub
```

### `DocFontEYInterstateLight`

```vbnet
Sub DocFontEYInterstateLight()

    Dim sr As Range
    Application.ScreenUpdating = False
    For Each sr In ActiveDocument.StoryRanges
        Do
            sr.Font.Name = "EYInterstate Light"
            Set sr = sr.NextStoryRange
        Loop Until sr Is Nothing
    Next
    Application.ScreenUpdating = True

End Sub
```

### `DocFontSizeDecrease`

```vbnet
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
```

### `DocFontSizeIncrease`

```vbnet
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
```

### `DocSpacingSingle`

```vbnet
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
```

### `DocTableMargin`

```vbnet
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
```

### `SelListAlphaRoman`

```vbnet
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
```

### `SelTableBorder`

```vbnet
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
```

### `SelTableMargin`

```vbnet
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
```

### `ViewSplitVerticalToggle`

```vbnet
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
```
