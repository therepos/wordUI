Option Explicit

' =============================================================================
' Reset
' -----------
' Ribbon dropdown menu entry points for reset operations.
' Each Public Sub is called directly from a ribbon menu button.
' =============================================================================

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

Public Sub RunResetFormat()
    Application.ScreenUpdating = False
    ResetFormat
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Formatting", vbInformation, "Reset"
End Sub

Public Sub RunResetList()
    Application.ScreenUpdating = False
    ResetList
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Lists", vbInformation, "Reset"
End Sub

Public Sub RunResetObject()
    Application.ScreenUpdating = False
    ResetObject
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Objects", vbInformation, "Reset"
End Sub

Public Sub RunResetTables()
    Application.ScreenUpdating = False
    ResetTables
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Tables", vbInformation, "Reset"
End Sub

Public Sub RunResetStylesAll()
    Application.ScreenUpdating = False
    ResetStylesCustom
    ResetStylesDefault
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Styles (All)", vbInformation, "Reset"
End Sub

Public Sub RunResetStylesDefault()
    Application.ScreenUpdating = False
    ResetStylesDefault
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Styles (Default)", vbInformation, "Reset"
End Sub

Public Sub RunResetHyperlinks()
    Application.ScreenUpdating = False
    ResetHyperlinks
    Application.ScreenUpdating = True
    MsgBox "Reset complete: Hyperlinks", vbInformation, "Reset"
End Sub

' ===========================================================================
'  RESET SUBS (Private)
' ===========================================================================

Private Sub ResetFormat()
    ActiveDocument.Content.Select
    Selection.ClearFormatting
End Sub

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

Private Sub ResetObject()
    Dim shp As InlineShape

    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
End Sub

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

Private Sub ResetHyperlinks()
    With ActiveDocument
        While .Hyperlinks.count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
End Sub
