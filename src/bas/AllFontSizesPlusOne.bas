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