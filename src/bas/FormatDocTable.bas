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