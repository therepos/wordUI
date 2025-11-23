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