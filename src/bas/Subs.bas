Attribute VB_Name = "Subs"
Option Explicit

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

Sub ListAlphaRoman()
    Dim lt As ListTemplate
    Dim rng As Range
    Dim startLevel As Long
    Dim prevPara As Paragraph
    
    Set rng = Selection.Range
    
    ' Detect current list level
    startLevel = 0
    If rng.ListFormat.ListType <> wdListNoNumbering Then
        ' Cursor is inside an existing list — use that level
        startLevel = rng.ListFormat.ListLevelNumber
        ' Remove existing list formatting for clean replacement
        rng.ListFormat.RemoveNumbers
    Else
        ' Cursor is NOT in a list — check the previous paragraph
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

Sub ResetFormat()
    ActiveDocument.Content.Select
    Selection.ClearFormatting
End Sub

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

Sub FormatAllFontsArial()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "Arial"
    Next para
End Sub

Sub FormatAllFontsEYInterstateLight()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "EYInterstate Light"
    Next para
End Sub

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

' ================================ New No Ribbon
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

Sub InsertSymbol()
'   Purpose: Insert symbol for reference

    Application.ScreenUpdating = False
    Selection.InsertSymbol Font:="Arial", CharacterNumber:=664, Unicode:=True
    Application.ScreenUpdating = True
        
End Sub

Sub KillTheHyperlinks()
'   Purpose: Removes all hyperlinks from the document

    With ThisDocument
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    
End Sub

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

Sub RemoveCrossReferences()
'   Purpose: Remove all cross-references

    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldRef Then
            fld.Unlink
        End If
    Next
 
End Sub

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






