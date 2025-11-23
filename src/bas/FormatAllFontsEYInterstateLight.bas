Sub FormatAllFontsEYInterstateLight()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "EYInterstate Light"
        para.Range.Font.Size = 11
    Next para
End Sub