# VBA modules

_This file is generated automatically from `.bas` files in `src/bas`._

## Module `AllFontSizesMinusOne`

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

## Module `AllFontSizesPlusOne`

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

## Module `FormatAllFontsEYInterstateLight`

### `FormatAllFontsEYInterstateLight`

```vbnet
Sub FormatAllFontsEYInterstateLight()
    Dim para As Paragraph
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        para.Range.Font.Name = "EYInterstate Light"
        para.Range.Font.Size = 11
    Next para
End Sub
```

## Module `FormatDocTable`

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

## Module `FormatSelectedTableBorders`

### `FormatSelectedTableBorders`

```vbnet
Sub FormatSelectedTableBorders()
    Dim tbl As Table
    Dim brd As Border
    
    ' Check if selection is inside a table
    If Selection.Information(wdWithInTable) Then
        Set tbl = Selection.Tables(1)
        
        ' Apply borders to all sides and internal lines
        With tbl
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
```
