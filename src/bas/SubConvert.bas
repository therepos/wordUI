Attribute VB_Name = "SubConvert"
' ============================================================
' US to UK English Converter (Smart Inflection) - Enhanced
' ============================================================

Private m_us() As String
Private m_uk() As String
Private m_count As Long
Private m_totalReplaced As Long

Public Sub ConvertUStoUK()

    ' Guard: document has no words
    If ActiveDocument.ComputeStatistics(wdStatisticWords) = 0 Then
        MsgBox "Document has no words.", vbInformation, "US to UK English"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "US to UK Conversion"

    m_count = 0
    m_totalReplaced = 0
    Erase m_us
    Erase m_uk

    ' ----- -ize / -ise -----
    r "recognize", "recognise"
    r "organize", "organise"
    r "realize", "realise"
    r "minimize", "minimise"
    r "maximize", "maximise"
    r "optimize", "optimise"
    r "utilize", "utilise"
    r "authorize", "authorise"
    r "categorize", "categorise"
    r "characterize", "characterise"
    r "customize", "customise"
    r "emphasize", "emphasise"
    r "finalize", "finalise"
    r "globalize", "globalise"
    r "harmonize", "harmonise"
    r "initialize", "initialise"
    r "legalize", "legalise"
    r "memorize", "memorise"
    r "modernize", "modernise"
    r "neutralize", "neutralise"
    r "normalize", "normalise"
    r "prioritize", "prioritise"
    r "specialize", "specialise"
    r "standardize", "standardise"
    r "summarize", "summarise"
    r "symbolize", "symbolise"
    r "synchronize", "synchronise"
    r "apologize", "apologise"
    r "capitalize", "capitalise"
    r "centralize", "centralise"
    r "criticize", "criticise"
    r "digitize", "digitise"
    r "dramatize", "dramatise"
    r "familiarize", "familiarise"
    r "fertilize", "fertilise"
    r "generalize", "generalise"
    r "hospitalize", "hospitalise"
    r "hypothesize", "hypothesise"
    r "idealize", "idealise"
    r "immunize", "immunise"
    r "itemize", "itemise"
    r "jeopardize", "jeopardise"
    r "liberalize", "liberalise"
    r "localize", "localise"
    r "marginalize", "marginalise"
    r "materialize", "materialise"
    r "mechanize", "mechanise"
    r "mobilize", "mobilise"
    r "monopolize", "monopolise"
    r "nationalize", "nationalise"
    r "penalize", "penalise"
    r "polarize", "polarise"
    r "privatize", "privatise"
    r "revolutionize", "revolutionise"
    r "scrutinize", "scrutinise"
    r "sensitize", "sensitise"
    r "socialize", "socialise"
    r "stabilize", "stabilise"
    r "sterilize", "sterilise"
    r "subsidize", "subsidise"
    r "terrorize", "terrorise"
    r "traumatize", "traumatise"
    r "trivialize", "trivialise"
    r "vandalize", "vandalise"
    r "vaporize", "vaporise"
    r "visualize", "visualise"

    ' ----- -or / -our -----
    r "color", "colour"
    r "favor", "favour"
    r "honor", "honour"
    r "humor", "humour"
    r "labor", "labour"
    r "neighbor", "neighbour"
    r "behavior", "behaviour"
    r "flavor", "flavour"
    r "harbor", "harbour"
    r "rumor", "rumour"
    r "tumor", "tumour"
    r "valor", "valour"
    r "vigor", "vigour"

    ' ----- -er / -re -----
    r "center", "centre"
    r "fiber", "fibre"
    r "liter", "litre"
    r "meter", "metre"
    r "theater", "theatre"

    ' ----- exact words -----
    X "aging", "ageing"
    X "airplane", "aeroplane"
    X "airplanes", "aeroplanes"
    X "aluminum", "aluminium"
    X "cozy", "cosy"
    X "gray", "grey"
    X "judgment", "judgement"
    X "math", "maths"
    X "program", "programme"
    X "programs", "programmes"
    X "check", "cheque"
    X "checks", "cheques"
    X "curb", "kerb"
    X "curbs", "kerbs"
    X "jewelry", "jewellery"
    X "skillful", "skilful"
    X "skillfully", "skilfully"

    RunAllReplacements

    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True

    If m_totalReplaced > 0 Then
        MsgBox m_totalReplaced & " word(s) converted. Ctrl+Z to undo.", vbInformation
    Else
        MsgBox "No US English words found.", vbInformation
    End If

End Sub

' ============================================================
' ENGINE
' ============================================================

Private Sub X(us As String, uk As String)

    m_count = m_count + 1
    ReDim Preserve m_us(1 To m_count)
    ReDim Preserve m_uk(1 To m_count)

    m_us(m_count) = us
    m_uk(m_count) = uk

End Sub

Private Sub r(us As String, uk As String)

    If EndsWith(us, "ize") And EndsWith(uk, "ise") Then

        Dim stem1 As String: stem1 = Left(us, Len(us) - 3)
        Dim stem2 As String: stem2 = Left(uk, Len(uk) - 3)

        X us, uk
        X stem1 & "izes", stem2 & "ises"
        X stem1 & "ized", stem2 & "ised"
        X stem1 & "izing", stem2 & "ising"
        X stem1 & "izer", stem2 & "iser"
        X stem1 & "ization", stem2 & "isation"

        Exit Sub
    End If

    X us, uk

End Sub

Private Sub RunAllReplacements()

    Dim sr As Range
    Dim storyType As Variant
    Dim i As Long
    Dim storyTypes As Variant

    storyTypes = Array( _
        wdMainTextStory, _
        wdFootnotesStory, _
        wdEndnotesStory, _
        wdPrimaryHeaderStory, _
        wdPrimaryFooterStory, _
        wdFirstPageHeaderStory, _
        wdFirstPageFooterStory, _
        wdEvenPagesHeaderStory, _
        wdEvenPagesFooterStory, _
        wdTextFrameStory)

    For Each storyType In storyTypes

        On Error Resume Next
        Set sr = ActiveDocument.StoryRanges(storyType)
        On Error GoTo 0

        Do While Not sr Is Nothing

            If sr.StoryLength > 1 Then

                For i = 1 To m_count

                    With sr.Find

                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = m_us(i)
                        .Replacement.text = m_uk(i)

                        .Forward = True
                        .Wrap = wdFindStop
                        .MatchCase = False
                        .MatchWholeWord = True

                        If .Execute(Replace:=wdReplaceAll) Then
                            m_totalReplaced = m_totalReplaced + 1
                        End If

                    End With

                Next i

            End If

            Set sr = sr.NextStoryRange

        Loop

    Next storyType

End Sub

Private Function EndsWith(s As String, suffix As String) As Boolean

    If Len(s) >= Len(suffix) Then
        EndsWith = (Right$(s, Len(suffix)) = suffix)
    End If

End Function

