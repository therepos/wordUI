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

## Module `SubConvertUStoUK`

### `ConvertUStoUK`

```vbnet
Public Sub ConvertUStoUK()
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "US to UK Conversion"
    
    ' ----- -ize / -ise (auto-covers -izes, -ized, -izing, -izer, -ization) -----
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
    
    ' ----- -or / -our (auto-covers -ors, -ored, -oring, -orous, -orable, -orful, -orless) -----
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
    R "armor", "armour"
    R "candor", "candour"
    R "clamor", "clamour"
    R "endeavor", "endeavour"
    R "fervor", "fervour"
    R "odor", "odour"
    R "parlor", "parlour"
    R "rancor", "rancour"
    R "rigor", "rigour"
    R "savior", "saviour"
    R "splendor", "splendour"
    R "demeanor", "demeanour"
    
    ' ----- -er / -re (auto-covers -ers, -ered) -----
    R "center", "centre"
    R "fiber", "fibre"
    R "liter", "litre"
    R "meter", "metre"
    R "theater", "theatre"
    R "caliber", "calibre"
    R "scepter", "sceptre"
    R "somber", "sombre"
    R "specter", "spectre"
    R "luster", "lustre"
    R "meager", "meagre"
    R "saber", "sabre"
    
    ' ----- -ense / -ence (auto-covers -enses, -enseless, -ensible) -----
    R "defense", "defence"
    R "offense", "offence"
    R "license", "licence"
    R "pretense", "pretence"
    
    ' ----- -og / -ogue (auto-covers -ogs, -oged, -oging) -----
    R "analog", "analogue"
    R "catalog", "catalogue"
    R "dialog", "dialogue"
    R "monolog", "monologue"
    R "prolog", "prologue"
    R "epilog", "epilogue"
    
    ' ----- -yze / -yse (auto-covers -yzed, -yzing, -yzes, -yzer) -----
    R "analyze", "analyse"
    R "paralyze", "paralyse"
    R "catalyze", "catalyse"
    
    ' ----- doubled consonants (auto-covers -ed, -ing, -er, -ers) -----
    R "travel", "travell"
    R "cancel", "cancell"
    R "model", "modell"
    R "label", "labell"
    R "level", "levell"
    R "fuel", "fuell"
    R "counsel", "counsell"
    R "dial", "diall"
    
    ' ----- exact matches only (no inflection) -----
    X "aging", "ageing"
    X "airplane", "aeroplane"
    X "airplanes", "aeroplanes"
    X "aluminum", "aluminium"
    X "artifact", "artefact"
    X "artifacts", "artefacts"
    X "cozy", "cosy"
    X "donut", "doughnut"
    X "donuts", "doughnuts"
    X "gray", "grey"
    X "inquire", "enquire"
    X "inquired", "enquired"
    X "inquiring", "enquiring"
    X "inquiry", "enquiry"
    X "inquiries", "enquiries"
    X "judgment", "judgement"
    X "judgments", "judgements"
    X "maneuver", "manoeuvre"
    X "maneuvers", "manoeuvres"
    X "maneuvered", "manoeuvred"
    X "maneuvering", "manoeuvring"
    X "mold", "mould"
    X "molds", "moulds"
    X "molded", "moulded"
    X "molding", "moulding"
    X "molt", "moult"
    X "molts", "moults"
    X "molted", "moulted"
    X "molting", "moulting"
    X "mustache", "moustache"
    X "mustaches", "moustaches"
    X "pajamas", "pyjamas"
    X "plow", "plough"
    X "plows", "ploughs"
    X "plowed", "ploughed"
    X "plowing", "ploughing"
    X "skeptic", "sceptic"
    X "skeptics", "sceptics"
    X "skeptical", "sceptical"
    X "skepticism", "scepticism"
    X "tire", "tyre"
    X "tires", "tyres"
    X "fetus", "foetus"
    X "diarrhea", "diarrhoea"
    X "anemia", "anaemia"
    X "anemic", "anaemic"
    X "anesthetic", "anaesthetic"
    X "anesthetics", "anaesthetics"
    X "archeology", "archaeology"
    X "archeological", "archaeological"
    X "estrogen", "oestrogen"
    X "pediatric", "paediatric"
    X "pediatrics", "paediatrics"
    X "pediatrician", "paediatrician"
    X "leukemia", "leukaemia"
    X "math", "maths"
    X "program", "programme"
    X "programs", "programmes"
    X "check", "cheque"
    X "checks", "cheques"
    X "checkbook", "chequebook"
    X "curb", "kerb"
    X "curbs", "kerbs"
    X "draft", "draught"
    X "drafts", "draughts"
    X "drafty", "draughty"
    X "enrollment", "enrolment"
    X "enrollments", "enrolments"
    X "enroll", "enrol"
    X "enrolls", "enrols"
    X "fulfill", "fulfil"
    X "fulfills", "fulfils"
    X "fulfillment", "fulfilment"
    X "instill", "instil"
    X "instills", "instils"
    X "installment", "instalment"
    X "installments", "instalments"
    X "jewelry", "jewellery"
    X "skillful", "skilful"
    X "skillfully", "skilfully"
    X "willful", "wilful"
    X "willfully", "wilfully"
    
    ' ===== ADD YOUR OWN HERE =====
    ' R "usword", "ukword"      <- smart: auto-handles inflections
    ' X "usword", "ukword"      <- exact: only replaces this specific word
    
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "Done. Ctrl+Z to undo.", vbInformation, "US to UK English"
End Sub
```

### `R`

```vbnet
Private Sub R(us As String, uk As String)
    ' Always do the root word itself
    X us, uk
    
    ' -ize / -ise pattern
    If EndsWith(us, "ize") And EndsWith(uk, "ise") Then
        Dim izeStem As String: izeStem = Left(us, Len(us) - 3)
        Dim iseStem As String: iseStem = Left(uk, Len(uk) - 3)
        X izeStem & "izes", iseStem & "ises"
        X izeStem & "ized", iseStem & "ised"
        X izeStem & "izing", iseStem & "ising"
        X izeStem & "izer", iseStem & "iser"
        X izeStem & "izers", iseStem & "isers"
        X izeStem & "ization", iseStem & "isation"
        X izeStem & "izations", iseStem & "isations"
        Exit Sub
    End If
    
    ' -yze / -yse pattern
    If EndsWith(us, "yze") And EndsWith(uk, "yse") Then
        Dim yzeStem As String: yzeStem = Left(us, Len(us) - 3)
        Dim yseStem As String: yseStem = Left(uk, Len(uk) - 3)
        X yzeStem & "yzes", yseStem & "yses"
        X yzeStem & "yzed", yseStem & "ysed"
        X yzeStem & "yzing", yseStem & "ysing"
        X yzeStem & "yzer", yseStem & "yser"
        X yzeStem & "yzers", yseStem & "ysers"
        Exit Sub
    End If
    
    ' -or / -our pattern
    If EndsWith(us, "or") And EndsWith(uk, "our") Then
        X us & "s", uk & "s"
        X us & "ed", uk & "ed"
        X us & "ing", uk & "ing"
        X us & "ful", uk & "ful"
        X us & "fully", uk & "fully"
        X us & "less", uk & "less"
        X us & "able", uk & "able"
        X us & "ous", uk & "ous"
        X us & "ite", uk & "ite"
        X us & "ites", uk & "ites"
        X us & "al", uk & "al"
        Exit Sub
    End If
    
    ' -er / -re pattern
    If EndsWith(us, "er") And EndsWith(uk, "re") Then
        X us & "s", uk & "s"
        X us & "ed", uk & "d"
        Exit Sub
    End If
    
    ' -ense / -ence pattern
    If EndsWith(us, "ense") And EndsWith(uk, "ence") Then
        Dim enseStem As String: enseStem = Left(us, Len(us) - 4)
        Dim enceStem As String: enceStem = Left(uk, Len(uk) - 4)
        X enseStem & "enses", enceStem & "ences"
        X enseStem & "enseless", enceStem & "enceless"
        X enseStem & "ensible", enceStem & "encible"
        Exit Sub
    End If
    
    ' -og / -ogue pattern
    If EndsWith(us, "og") And EndsWith(uk, "ogue") Then
        X us & "s", uk & "s"
        X us & "ed", uk & "d"
        X us & "ing", uk & "ing"
        Exit Sub
    End If
    
    ' doubled consonant pattern (e.g. travel -> travell)
    ' The root won't match whole words usefully, but the suffixes will
    If Len(uk) = Len(us) + 1 And Left(uk, Len(us)) = us Then
        X us & "ed", uk & "ed"
        X us & "ing", uk & "ing"
        X us & "er", uk & "er"
        X us & "ers", uk & "ers"
        X us & "or", uk & "or"
        X us & "ors", uk & "ors"
        ' Don't replace the bare root (travel is valid in both)
        Exit Sub
    End If
End Sub
```

### `X`

```vbnet
Private Sub X(us As String, uk As String)
    Dim sr As Range
    For Each sr In ActiveDocument.StoryRanges
        Do
            With sr.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .text = us
                .Replacement.text = uk
                .MatchWholeWord = True
                .MatchCase = False
                .Forward = True
                .Wrap = wdFindStop
                .Execute Replace:=wdReplaceAll
            End With
            Set sr = sr.NextStoryRange
        Loop Until sr Is Nothing
    Next sr
End Sub
```

### `EndsWith`

```vbnet
Private Function EndsWith(s As String, suffix As String) As Boolean
    If Len(s) >= Len(suffix) Then
        EndsWith = (Right(s, Len(suffix)) = suffix)
    End If
End Function
```

## Module `SubResetPicker`

### `ResetPicker`

```vbnet
Public Sub ResetPicker()
    Dim frm As frmResetPicker
    Set frm = New frmResetPicker
    frm.Show vbModal

    If frm.Tag <> "OK" Then
        Unload frm
        Exit Sub
    End If

    ' --- Read checkbox values before unloading ---
    Dim doFormat    As Boolean: doFormat = frm.Controls("chkFormat").Value
    Dim doList      As Boolean: doList = frm.Controls("chkList").Value
    Dim doObjects   As Boolean: doObjects = frm.Controls("chkObjects").Value
    Dim doTables    As Boolean: doTables = frm.Controls("chkTables").Value
    Dim doStylesAll As Boolean: doStylesAll = frm.Controls("chkStyleAll").Value
    Dim doStylesBI  As Boolean: doStylesBI = frm.Controls("chkStyleBI").Value

    Unload frm
    Set frm = Nothing

    ' --- Execute selected resets ---
    Application.ScreenUpdating = False

    Dim actions As String
    actions = ""

    If doFormat Then
        ResetFormat
        actions = actions & "Formatting, "
    End If

    If doList Then
        ResetList
        actions = actions & "Lists, "
    End If

    If doObjects Then
        ResetObject
        actions = actions & "Objects, "
    End If

    If doTables Then
        ResetTables
        actions = actions & "Tables, "
    End If

    If doStylesAll Then
        ResetStylesCustom
        ResetStylesDefault
        actions = actions & "Styles (All), "
    ElseIf doStylesBI Then
        ResetStylesDefault
        actions = actions & "Styles (Built-in), "
    End If

    Application.ScreenUpdating = True

    If Len(actions) > 0 Then
        actions = Left$(actions, Len(actions) - 2)
        MsgBox "Reset complete:" & vbCrLf & vbCrLf & actions, _
               vbInformation, "Reset Picker"
    End If
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

## Module `Subs`

### `ResetStyles`

```vbnet
Sub ResetStyles()
    Dim docOrig As Object
    Dim docTemp As Object
    Dim sty As Object
    
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
End Sub
```

### `ListAlphaRoman`

```vbnet
Sub ListAlphaRoman()
    Dim lt As ListTemplate
    Dim rng As Range
    Dim startLevel As Long
    Dim prevPara As Paragraph
    
    Set rng = Selection.Range
    
    ' Detect current list level
    startLevel = 0
    If rng.ListFormat.ListType <> wdListNoNumbering Then
        ' Cursor is inside an existing list  use that level
        startLevel = rng.ListFormat.ListLevelNumber
        ' Remove existing list formatting for clean replacement
        rng.ListFormat.RemoveNumbers
    Else
        ' Cursor is NOT in a list  check the previous paragraph
        If rng.Paragraphs(1).Range.Start > 0 Then
            Set prevPara = rng.Paragraphs(1).Previous
            If Not prevPara Is Nothing Then
                If prevPara.Range.ListFormat.ListType <> wdListNoNumbering Then
                    ' Previous para is in a list, so we're continuing one level deeper
                    startLevel = prevPara.Range.ListFormat.ListLevelNumber + 1
                End If
            End If
        End If
    End If
    
    ' Fallback to level 1 if we couldn't determine context
    If startLevel < 1 Then startLevel = 1
    If startLevel > 8 Then startLevel = 8  ' cap so startLevel+1 <= 9
    
    Set lt = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)
    
    ' Alpha level: (a), (b), (c)...
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
    
    ' Roman sublevel: (i), (ii), (iii)...
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

### `ResetFormat`

```vbnet
Sub ResetFormat()
    ActiveDocument.Content.Select
    Selection.ClearFormatting
End Sub
```

### `ResetTables`

```vbnet
Sub ResetTables()
    Dim tbl As Word.Table
    Dim s As Word.Style
    Dim tableNormal As Word.Style

    ' Find the built-in "Table Normal" style without using the enum
    For Each s In ActiveDocument.Styles
        If s.BuiltIn Then
            ' Built-in table styles have Type = wdStyleTypeTable
            If s.Type = wdStyleTypeTable Then
                ' "NameLocal" is the localized display name; look for the "Normal" base table style
                ' This heuristic avoids hardcoding the English name
                If LCase$(s.NameLocal) Like "*normal*" Or LCase$(s.NameLocal) Like "*table normal*" Then
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
            ' Fallbacks:
            ' 1) English name (works if UI is English)
            tbl.Style = "Table Normal"
            If Err.Number <> 0 Then
                Err.Clear
                ' 2) A very common built-in table style (English UI)
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

### `ResetList`

```vbnet
Sub ResetList()

    Dim para As Paragraph
    Dim lt As WdListType
    Dim lvl As Long

    For Each para In ActiveDocument.Paragraphs
        lt = para.Range.ListFormat.ListType
        If lt <> wdListNoNumbering Then
            ' Remember the current nesting level
            lvl = para.Range.ListFormat.ListLevelNumber

            ' Remove then reapply to reset formatting
            para.Range.ListFormat.RemoveNumbers

            If lt = wdListBullet Then
                para.Range.ListFormat.ApplyBulletDefault
            Else
                para.Range.ListFormat.ApplyNumberDefault
            End If

            ' Restore the original nesting level
            para.Range.ListFormat.ListLevelNumber = lvl
        End If
    Next para

End Sub
```

### `FormatDocTable`

```vbnet
Sub FormatDocTable()
    Dim tbl As Table
    Dim topMargin As Single, bottomMargin As Single
    Dim leftMargin As Single, rightMargin As Single
    
    ' Convert cm to points (1 cm = 28.35 points approx)
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
End Sub
```

### `FormatSelectedTableMargins`

```vbnet
Sub FormatSelectedTableMargins()
    Dim tbl As Table
    Const PAD_TOP_CM As Double = 0.05
    Const PAD_BOTTOM_CM As Double = 0.05
    Const PAD_LEFT_CM As Double = 0.19
    Const PAD_RIGHT_CM As Double = 0.19

    If Selection.Information(wdWithInTable) Then
        ' Prefer resolving via selected cell if available
        If Selection.Cells.Count > 0 Then
            ' Parent table of the first selected cell
            Set tbl = Selection.Cells(1).Range.Tables(1)
        ElseIf Selection.Range.Tables.Count > 0 Then
            ' Innermost table within the current selection range
            Set tbl = Selection.Range.Tables(Selection.Range.Tables.Count)
        Else
            MsgBox "Couldn't resolve the table from the selection.", vbExclamation
            Exit Sub
        End If
        
        With tbl
            .TopPadding = Application.CentimetersToPoints(PAD_TOP_CM)
            .BottomPadding = Application.CentimetersToPoints(PAD_BOTTOM_CM)
            .LeftPadding = Application.CentimetersToPoints(PAD_LEFT_CM)
            .RightPadding = Application.CentimetersToPoints(PAD_RIGHT_CM)
            
            ' Optional:
            ' .CellSpacing = 0
            ' .AllowAutoFit = False
        End With
        
        MsgBox "Cell margins (padding) applied to the selected (innermost) table.", vbInformation
    Else
        MsgBox "Place the cursor inside the table you want to format, then run the macro.", vbExclamation
    End If
End Sub
```

### `FormatSelectedTableBorders`

```vbnet
Sub FormatSelectedTableBorders()
    Dim tbl As Table
    Dim bTypes As Variant
    Dim i As Long

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)

    ' Only these borders (no diagonals)
    bTypes = Array( _
        wdBorderLeft, _
        wdBorderRight, _
        wdBorderTop, _
        wdBorderBottom, _
        wdBorderHorizontal, _
        wdBorderVertical)

    With tbl
        For i = LBound(bTypes) To UBound(bTypes)
            With .Borders(bTypes(i))
                .LineStyle = wdLineStyleSingle
                .Color = wdColorAutomatic
                .LineWidth = wdLineWidth025pt
            End With
        Next i

        ' make sure diagonals are off
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
End Sub
```

### `FormatAllFontsArial`

```vbnet
Sub FormatAllFontsArial()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "Arial"
    Next para
End Sub
```

### `FormatAllFontsEYInterstateLight`

```vbnet
Sub FormatAllFontsEYInterstateLight()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "EYInterstate Light"
    Next para
End Sub
```

### `AllFontSizesMinusOne`

```vbnet
Sub AllFontSizesMinusOne()
    Dim para As Paragraph
    Dim rng As Range
    
    ' Loop through all paragraphs
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        
        ' Reduce font size by 1 point if greater than 1
        If rng.Font.Size > 1 Then
            rng.Font.Size = rng.Font.Size - 1
        End If
    Next para
End Sub
```

### `AllFontSizesPlusOne`

```vbnet
Sub AllFontSizesPlusOne()
    Dim para As Paragraph
    Dim rng As Range
    
    ' Loop through all paragraphs
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        ' Increase font size by 1 point
        rng.Font.Size = rng.Font.Size + 1
    Next para
End Sub
```

### `Document_ContentControlOnExit`

```vbnet
Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
'   Purpose: Change Textbox content as Dropdown List change.

    Dim oCC As ContentControl
    Dim oRng As Word.Range

    Select Case ContentControl.Title
      Case "Client" 'The "Client" Dropdown CC in document.
        Set oCC = ActiveDocument.SelectContentControlsByTitle("RegNo").Item(1) 'Richtext CC in Header
        Select Case True
          Case ContentControl.ShowingPlaceholderText:
            oCC.Range.text = vbNullString
          Case ContentControl.Range.text = "AC ALLIANCES (PAC)"
            oCC.Range.text = "201118268H"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdRed
          Case ContentControl.Range.text = "Y M WOO & CO"
            oCC.Range.text = "S88PF0309G"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdBlue
          Case ContentControl.Range.text = "GAAP PAC"
            oCC.Range.text = "201831129C"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdGreen
        End Select
    End Select
lbl_Exit:
    Exit Sub
    
End Sub
```

### `CopyHyperlink`

```vbnet
Sub CopyHyperlink()
'   Purpose: Copy hyperlinks
'   Reference: https://www.msofficeforums.com/word-vba/38223-how-extract-selected-hyperlink-address-clipboard.html
'   Reference: https://software-solutions-online.com/word-vba-move-cursor-to-end-of-document/
'   Reference: https://gregmaxey.com/word_tips.html
'   Reference: https://www.thespreadsheetguru.com/blog/dynamically-populating-array-vba-variables
'   Reference: https://stackoverflow.com/questions/39690078/vba-output-contents-of-array-to-word-document
'   =========================================
'    For i = 1 To Selection.Hyperlinks.Count
'        With Selection.Hyperlinks(i)
'          StrTxt = .Address
'          If .SubAddress <> "" Then StrTxt = StrTxt & "#" & .SubAddress
'          With .Range.Fields(1).Code
'            .Text = StrTxt
'            .Copy
'          End With
'        End With
'        ActiveDocument.Undo
'    Next
'    Selection.EndKey Unit:=wdStory
'    Selection.Range.Text = vbNewLine
'    Selection.Paste

    Dim StrTxt As String
    Dim results() As Variant
    Dim inputWord As Variant
    Dim i As Long
    Dim insertPos As Range
    Set insertPos = Selection.Range
    
    ReDim results(Selection.Hyperlinks.Count)
    For i = 1 To Selection.Hyperlinks.Count
        results(i) = Selection.Hyperlinks(i).TextToDisplay & ": " & Selection.Hyperlinks(i).Address
    Next

    For Each inputWord In results
        insertPos.Collapse wdCollapseEnd
        insertPos = inputWord & vbCrLf
    Next
     
End Sub
```

### `Document_ContentControlOnExit`

```vbnet
Sub Document_ContentControlOnExit(, ByVal ContentControl As ContentControl, Cancel As Boolean)
'   Purpose: Change Textbox content as Dropdown List change.

    Dim oCC As ContentControl
    Dim oRng As Word.Range

    Select Case ContentControl.Title
      Case "Client" 'The "Client" Dropdown CC in document.
        Set oCC = ActiveDocument.SelectContentControlsByTitle("RegNo").Item(1) 'Richtext CC in Header
        Select Case True
          Case ContentControl.ShowingPlaceholderText:
            oCC.Range.text = vbNullString
          Case ContentControl.Range.text = "AC ALLIANCES (PAC)"
            oCC.Range.text = "201118268H"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdRed
          Case ContentControl.Range.text = "Y M WOO & CO"
            oCC.Range.text = "S88PF0309G"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdBlue
          Case ContentControl.Range.text = "GAAP PAC"
            oCC.Range.text = "201831129C"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdGreen
        End Select
    End Select
lbl_Exit:
    Exit Sub
    
End Sub
```

### `EditLinks`

```vbnet
Sub EditLinks()
'   Purpose: Edit hyperlinks
'   Reference: https://stackoverflow.com/questions/3355266/how-to-programmatically-edit-all-hyperlinks-in-a-word-document
'   Reference: http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.hyperlink_members.aspx
'   Reference: https://shaunakelly.com/word/word-development/selecting-or-referring-to-a-page-in-the-word-object-model.html
'   =========================================
'   doc.Hyperlinks(i).Address = Replace(doc.Hyperlinks(i).Address, "gopher:", "https://")
'   If LCase(doc.Hyperlinks(i).Address) Like "*partOfHyperlinkHere*" Then
'   doc.Hyperlinks(i).Address = Mid(doc.Hyperlinks(i).Address, 70,20)

    Dim i As Long
    For i = 1 To Selection.Hyperlinks.Count
        Selection.Hyperlinks(i).TextToDisplay = "[" & Selection.Hyperlinks(i).TextToDisplay & "]"
    Next
    
    Call CopyHyperlink(control)
    
End Sub
```

### `InsertReference`

```vbnet
Sub InsertReference()
'   Purpose: Paste clipboard content as hyperlink
'   References: https://www.slipstick.com/developer/code-samples/paste-clipboard-contents-vba/
'   Notes:
'   - https://excel-macro.tutorialhorizon.com/vba-excel-reference-libraries-in-excel-workbook/

    Dim DataObj As MSForms.DataObject
    Set DataObj = New MSForms.DataObject
    Dim strPaste As Variant
    DataObj.GetFromClipboard
    
    Application.ScreenUpdating = False
    
    strPaste = DataObj.GetText(1)
    If strPaste = False Then Exit Sub
    If strPaste = "" Then Exit Sub

    Selection.TypeText text:="["
    ActiveDocument.Hyperlinks.Add _
        Anchor:=Selection.Range, _
        Address:=strPaste, _
        TextToDisplay:=ChrW(664)
    Selection.TypeText text:="]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Color = wdColorBlue
    Selection.Move
    Selection.MoveRight Unit:=wdCharacter, Count:=1

    Set DataObj = Nothing
    
    Application.ScreenUpdating = True
    
End Sub
```

### `InsertSymbol`

```vbnet
Sub InsertSymbol()
'   Purpose: Insert symbol for reference

    Application.ScreenUpdating = False
    Selection.InsertSymbol Font:="Arial", CharacterNumber:=664, Unicode:=True
    Application.ScreenUpdating = True
        
End Sub
```

### `KillTheHyperlinks`

```vbnet
Sub KillTheHyperlinks()
'   Purpose: Removes all hyperlinks from the document

    With ThisDocument
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    
End Sub
```

### `KillTheHyperlinksInAllOpenDocuments`

```vbnet
Sub KillTheHyperlinksInAllOpenDocuments()
'   Purpose: Removes all hyperlinks from all opened document

    Dim doc As Document
    Dim szOpenDocName As String
     
    For Each doc In Application.Documents
        szOpenDocName = doc.Name
        With Documents(szOpenDocName)
            While .Hyperlinks.Count > 0
                .Hyperlinks(1).Delete
            Wend
        End With
        Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    Next doc
    
End Sub
```

### `OpenEmbeddedExcelInWord`

```vbnet
Sub OpenEmbeddedExcelInWord()
'   Purpose: Remove LockAspectRatio from linked Excel objects

    Dim shp As InlineShape
    
    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
    
End Sub
```

### `RemoveContentControl`

```vbnet
Sub RemoveContentControl()
'   Purpose: Remove all content controls
    
    Dim oRng As Range
    Dim CC   As ContentControl
    Dim LC   As Integer
    Dim LRCC As Integer
    Dim LTCC As Integer
    Dim LE   As Boolean

    Set oRng = ActiveDocument.Content
    LTCC = LTCC + oRng.ContentControls.Count
    For LC = oRng.ContentControls.Count To 1 Step -1
    
    Set CC = oRng.ContentControls(LC)
    If CC.LockContentControl = True Then
        CC.LockContentControl = False
    End If
    CC.Delete
    If Not LE Then
        LRCC = LRCC + 1
        End If
        LE = False
    Next
    
End Sub
```

### `RemoveCrossReferences`

```vbnet
Sub RemoveCrossReferences()
'   Purpose: Remove all cross-references

    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldRef Then
            fld.Unlink
        End If
    Next
 
End Sub
```

### `ResetObject`

```vbnet
Sub ResetObject()
'   Purpose: Reset WorkbookObject sizes

    Dim shp As InlineShape
    
    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
    
End Sub
```

### `ResizeImage`

```vbnet
Sub ResizeImage()
'   Purpose: Resize selected image
'   Source: https://www.extendoffice.com/documents/word/1207-word-resize-all-multiple-images.html

    Dim shp As Word.Shape
    Dim ishp As Word.InlineShape
    If Word.Selection.Type <> wdSelectionInlineShape And _
        Word.Selection.Type <> wdSelectionShape Then
            Exit Sub
    End If
    If Word.Selection.Type = wdSelectionInlineShape Then
        Set ishp = Word.Selection.Range.InlineShapes(1)
        ishp.LockAspectRatio = False
        ishp.Height = CentimetersToPoints(5)
        ishp.Width = CentimetersToPoints(5)
    Else
        If Word.Selection.Type = wdSelectionShape Then
            Set shp = Word.Selection.ShapeRange(1)
            shp.LockAspectRatio = False
            shp.Height = CentimetersToPoints(5)
            shp.Width = CentimetersToPoints(5)
        End If
    End If
    
End Sub
```

### `SetPageLayout`

```vbnet
Sub SetPageLayout()
'   Purpose: Set page margin and edge distance

    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .topMargin = CentimetersToPoints(1)
        .bottomMargin = CentimetersToPoints(1)
        .leftMargin = CentimetersToPoints(3)
        .rightMargin = CentimetersToPoints(1.8)
        .HeaderDistance = CentimetersToPoints(1)
        .FooterDistance = CentimetersToPoints(1)
        .PaperSize = wdPaperA4
    End With

End Sub
```

### `SetParagraph`

```vbnet
Sub SetParagraph()
'   Purpose: Set paragraph spacing

    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
  
End Sub
```

### `SetTablesBordersColor`

```vbnet
Sub SetTablesBordersColor(, varColor As Long)
'   Purpose: Standardise all table borders color in a document at 1/2 pt single-line

    Application.ScreenUpdating = False
    
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Borders.InsideColor = varColor
        tbl.Borders.InsideLineStyle = wdLineStyleSingle
'        Tbl.Borders.InsideLineWidth = wdLineWidth050pt
        tbl.Borders.OutsideColor = varColor
        tbl.Borders.OutsideLineStyle = wdLineStyleSingle
'        Tbl.Borders.OutsideLineWidth = wdLineWidth050pt
    Next
        
    Application.ScreenUpdating = True
    Application.ScreenRefresh

End Sub
```

### `SetTablesMargin`

```vbnet
Sub SetTablesMargin(, varPadding As Double)
'   Purpose: Standardise all table paddings in a document
'   varPadding: Measured in centimeters
'   Notes:

    Application.ScreenUpdating = False
    
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.AutoFitBehavior (wdAutoFitWindow)
        tbl.AllowAutoFit = True
        tbl.LeftPadding = CentimetersToPoints(varPadding)
        tbl.RightPadding = CentimetersToPoints(varPadding)
        tbl.TopPadding = CentimetersToPoints(varPadding)
        tbl.BottomPadding = CentimetersToPoints(varPadding)
    Next
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
End Sub
```

### `SplitVertically`

```vbnet
Sub SplitVertically()
'   Purpose: Split WORD active window vertically to view side by side
'   Source: https://dharma-records.buddhasasana.net/computing/ms-word-split-windows-vertically
 
    Dim Win1 As Integer
    Dim Win2 As Integer
     
    Dim WinWidth As Integer
    Dim WinHeight As Integer
 
'   Check for duplicated window
 
    Dim Win As Word.Window
    Dim DocString As String
    Dim FirstString
    Dim SecondString
    Dim StringLength As Long
     
    For Each Win In Application.Windows
    DocString = Win
    FirstString = Right(DocString, 1)
     
        If FirstString = "1" Then
         
'   If there is a duplicate
         
        StringLength = Len(DocString) - 2
        SecondString = Left(DocString, StringLength)
         
'   Close the copy
     
        Windows(SecondString & ":2").Close
         
'   Activate and maximise the identified document window
     
        Windows(Win).Activate
        Windows(Win).WindowState = wdWindowStateMaximize
         
        GoTo TheEnd
         
        End If

'   Otherwise check the next window
     
    Next Win
     
'   If there are no duplicates, get the dimensions
     
    ActiveWindow.WindowState = wdWindowStateMaximize
     
'   Find the serial number of the window and set the variable
     
    Win1 = ActiveWindow.Index
     
'   Set the dimension variables
     
    WinHeight = ActiveWindow.Height - 20
    WinWidth = ActiveWindow.Width

'   Make a new window from the first
     
    NewWindow
     
'   Find the serial number of the new window
         
    Win2 = ActiveWindow.Index
         
'   Arrange all windows (window must be in maximised state)
     
    Windows.Arrange
     
'   Set the size of the two windows we found
     
    With Windows(Win1)
        .Left = 0
        .Top = 0
        .Height = WinHeight
        .Width = WinWidth / 2
    End With
     
    With Windows(Win2)
        .Left = WinWidth / 2
        .Top = 0
        .Height = WinHeight
        .Width = WinWidth / 2
    End With
        
'   Return to the first window
     
    Windows(Win1).Activate
         
'   Added by Ryan
'   Resize windows
'   Width = 486

    Application.Resize Width:=WinWidth / 2, Height:=Application.UsableHeight
    Windows(Win2).Activate
    Application.Resize Width:=WinWidth / 2, Height:=Application.UsableHeight
    
TheEnd:
         
End Sub
```

### `StyleKill`

```vbnet
Sub StyleKill()
'   Purpose: Delete unwanted styles
'   Source: https://word.tips.net/T001337_Removing_Unused_Styles.html

    Dim oStyle As Style
    For Each oStyle In ActiveDocument.Styles
        'Only check out non-built-in styles
        If oStyle.BuiltIn = False Then
                oStyle.Delete
        End If
    Next oStyle
     
End Sub
```
