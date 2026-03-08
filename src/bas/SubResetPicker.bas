Attribute VB_Name = "ResetPicker"
Option Explicit

' =============================================================================
' ResetPicker
' -----------
' Single-button entry point that opens frmResetPicker allowing the user
' to choose which reset operations to run.
'
' Consolidates: ResetFormat, ResetList, ResetObject, ResetTables,
'               ResetStylesDefault (formerly ResetStyles),
'               ResetStylesCustom  (formerly StyleKill)
'
' Import both files:
'   1. ResetPicker.bas    (this module)
'   2. frmResetPicker.frm (the UserForm)
' =============================================================================

' ---------------------------------------------------------------------------
' Public entry point - wire this to your ribbon / QAT button
' ---------------------------------------------------------------------------
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

' ===========================================================================
'  RESET SUBS
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
