Attribute VB_Name = "Subs"
Sub FormatDocTable()
    Dim Tbl As Table
    Dim topMargin As Single, bottomMargin As Single
    Dim leftMargin As Single, rightMargin As Single
    
    ' Convert cm to points (1 cm = 28.35 points approx)
    topMargin = CentimetersToPoints(0.1)
    bottomMargin = CentimetersToPoints(0.1)
    leftMargin = CentimetersToPoints(0.19)
    rightMargin = CentimetersToPoints(0.19)
    
    For Each Tbl In ActiveDocument.Tables
        With Tbl
            .TopPadding = topMargin
            .BottomPadding = bottomMargin
            .LeftPadding = leftMargin
            .RightPadding = rightMargin
        End With
    Next Tbl
End Sub

Sub FormatSelectedTableBorders()
    Dim Tbl As Table
    Dim brd As Border
    
    ' Check if selection is inside a table
    If Selection.Information(wdWithInTable) Then
        Set Tbl = Selection.Tables(1)
        
        ' Apply borders to all sides and internal lines
        With Tbl
            ' Loop through all borders
            For Each brd In .Borders
                brd.LineStyle = wdLineStyleSingle   ' Solid line
                brd.Color = wdColorAutomatic        ' Automatic color
                brd.LineWidth = wdLineWidth025pt    ' ½ pt width
            Next brd
        End With
    Else
        MsgBox "Please place the cursor inside a table.", vbExclamation
    End If
End Sub

Sub FormatAllFontsEYInterstateLight()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "EYInterstate Light"
        para.Range.Font.Size = 11
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
            oCC.Range.Text = vbNullString
          Case ContentControl.Range.Text = "AC ALLIANCES (PAC)"
            oCC.Range.Text = "201118268H"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdRed
          Case ContentControl.Range.Text = "Y M WOO & CO"
            oCC.Range.Text = "S88PF0309G"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdBlue
          Case ContentControl.Range.Text = "GAAP PAC"
            oCC.Range.Text = "201831129C"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdGreen
        End Select
    End Select
lbl_Exit:
    Exit Sub
    
End Sub

Sub ClearTableStyles(control As IRibbonControl)
'   Purpose: Clear table styles

    Dim objTable As Table
    Dim objDoc As Document
    
    Application.ScreenUpdating = False
    Set objDoc = ActiveDocument
    
    For Each objTable In objDoc.Tables
      objTable.Style = "Table Normal"
      objTable.Borders.Enable = True
    Next objTable
    
    Application.ScreenUpdating = True
    
    Set objDoc = Nothing
  
End Sub

Sub CopyHyperlink(control As IRibbonControl)
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

Sub Document_ContentControlOnExit(control As IRibbonControl, ByVal ContentControl As ContentControl, Cancel As Boolean)
'   Purpose: Change Textbox content as Dropdown List change.

    Dim oCC As ContentControl
    Dim oRng As Word.Range

    Select Case ContentControl.Title
      Case "Client" 'The "Client" Dropdown CC in document.
        Set oCC = ActiveDocument.SelectContentControlsByTitle("RegNo").Item(1) 'Richtext CC in Header
        Select Case True
          Case ContentControl.ShowingPlaceholderText:
            oCC.Range.Text = vbNullString
          Case ContentControl.Range.Text = "AC ALLIANCES (PAC)"
            oCC.Range.Text = "201118268H"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdRed
          Case ContentControl.Range.Text = "Y M WOO & CO"
            oCC.Range.Text = "S88PF0309G"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdBlue
          Case ContentControl.Range.Text = "GAAP PAC"
            oCC.Range.Text = "201831129C"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdGreen
        End Select
    End Select
lbl_Exit:
    Exit Sub
    
End Sub

Sub EditLinks(control As IRibbonControl)
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

Sub InsertReference(control As IRibbonControl)
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

    Selection.TypeText Text:="["
    ActiveDocument.Hyperlinks.Add _
        Anchor:=Selection.Range, _
        Address:=strPaste, _
        TextToDisplay:=ChrW(664)
    Selection.TypeText Text:="]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Color = wdColorBlue
    Selection.Move
    Selection.MoveRight Unit:=wdCharacter, Count:=1

    Set DataObj = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Sub InsertSymbol(control As IRibbonControl)
'   Purpose: Insert symbol for reference

    Application.ScreenUpdating = False
    Selection.InsertSymbol Font:="Arial", CharacterNumber:=664, Unicode:=True
    Application.ScreenUpdating = True
        
End Sub

Sub KillTheHyperlinks(control As IRibbonControl)
'   Purpose: Removes all hyperlinks from the document

    With ThisDocument
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    
End Sub

Sub KillTheHyperlinksInAllOpenDocuments(control As IRibbonControl)
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

Sub OpenEmbeddedExcelInWord(control As IRibbonControl)
'   Purpose: Remove LockAspectRatio from linked Excel objects

    Dim shp As InlineShape
    
    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
    
End Sub

Sub RemoveContentControl(control As IRibbonControl)
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

Sub RemoveCrossReferences(control As IRibbonControl)
'   Purpose: Remove all cross-references

    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldRef Then
            fld.Unlink
        End If
    Next
 
End Sub

Sub ResetObject(control As IRibbonControl)
'   Purpose: Reset WorkbookObject sizes

    Dim shp As InlineShape
    
    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
    
End Sub

Sub ResizeImage(control As IRibbonControl)
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

Sub SetPageLayout(control As IRibbonControl)
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

Sub SetParagraph(control As IRibbonControl)
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

Sub SetTablesBordersColor(control As IRibbonControl, varColor As Long)
'   Purpose: Standardise all table borders color in a document at 1/2 pt single-line

    Application.ScreenUpdating = False
    
    Dim Tbl As Table
    For Each Tbl In ActiveDocument.Tables
        Tbl.Borders.InsideColor = varColor
        Tbl.Borders.InsideLineStyle = wdLineStyleSingle
'        Tbl.Borders.InsideLineWidth = wdLineWidth050pt
        Tbl.Borders.OutsideColor = varColor
        Tbl.Borders.OutsideLineStyle = wdLineStyleSingle
'        Tbl.Borders.OutsideLineWidth = wdLineWidth050pt
    Next
        
    Application.ScreenUpdating = True
    Application.ScreenRefresh

End Sub

Sub SetTablesMargin(control As IRibbonControl, varPadding As Double)
'   Purpose: Standardise all table paddings in a document
'   varPadding: Measured in centimeters
'   Notes:

    Application.ScreenUpdating = False
    
    Dim Tbl As Table
    For Each Tbl In ActiveDocument.Tables
        Tbl.AutoFitBehavior (wdAutoFitWindow)
        Tbl.AllowAutoFit = True
        Tbl.LeftPadding = CentimetersToPoints(varPadding)
        Tbl.RightPadding = CentimetersToPoints(varPadding)
        Tbl.TopPadding = CentimetersToPoints(varPadding)
        Tbl.BottomPadding = CentimetersToPoints(varPadding)
    Next
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
End Sub

Sub SplitVertically(control As IRibbonControl)
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

Sub StyleKill(control As IRibbonControl)
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






