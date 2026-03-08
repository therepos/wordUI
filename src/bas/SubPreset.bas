Attribute VB_Name = "SubPreset"
Option Explicit

' =============================================================================
' Font Presets
' =============================================================================

Public Sub RunPresetFontArial()
    ApplyPreset "Arial"
End Sub

Public Sub RunPresetFontEY()
    ApplyPreset "EYInterstate Light"
End Sub

Public Sub RunPresetFontTimes()
    ApplyPreset "Times New Roman"
End Sub

Public Sub RunPresetFontCalibri()
    ApplyPreset "Calibri"
End Sub


' ===========================================================================
' INTERNAL
' ===========================================================================

Private Sub ApplyPreset(f As String)

    Application.ScreenUpdating = False
    SetPresetFont f
    ApplyFontPreset
    Application.ScreenUpdating = True
    MsgBox "Font applied: " & f, vbInformation, "Font"

End Sub


Private Sub ApplyFontPreset()

    Dim f As String
    Dim sr As Range

    On Error Resume Next
    f = ActiveDocument.CustomDocumentProperties("PresetFont").Value
    If f = "" Then f = "Arial"
    On Error GoTo 0
    For Each sr In ActiveDocument.StoryRanges
        Do
            sr.Font.Name = f
            Set sr = sr.NextStoryRange
        Loop Until sr Is Nothing
    Next sr

End Sub


Private Sub SetPresetFont(f As String)

    On Error Resume Next
    If ActiveDocument.CustomDocumentProperties("PresetFont").Name = "" Then
        ActiveDocument.CustomDocumentProperties.Add _
            Name:="PresetFont", _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=f
    Else
        ActiveDocument.CustomDocumentProperties("PresetFont").Value = f
    End If

End Sub

