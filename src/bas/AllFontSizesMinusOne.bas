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