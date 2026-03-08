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
    R "recognize", "recognise"
    R "organize", "organise"
    R "realize", "realise"
    R "minimize", "minimise"
    R "maximize", "maximise"
    R "optimize", "optimise"
    R "utilize", "utilise"
    R "authorize", "authorise"
    R "categorize", "categorise"
    R "characterize", "characterise"
    R "customize", "customise"
    R "emphasize", "emphasise"
    R "finalize", "finalise"
    R "globalize", "globalise"
    R "harmonize", "harmonise"
    R "initialize", "initialise"
    R "legalize", "legalise"
    R "memorize", "memorise"
    R "modernize", "modernise"
    R "neutralize", "neutralise"
    R "normalize", "normalise"
    R "prioritize", "prioritise"
    R "specialize", "specialise"
    R "standardize", "standardise"
    R "summarize", "summarise"
    R "symbolize", "symbolise"
    R "synchronize", "synchronise"
    R "apologize", "apologise"
    R "capitalize", "capitalise"
    R "centralize", "centralise"
    R "criticize", "criticise"
    R "digitize", "digitise"
    R "dramatize", "dramatise"
    R "familiarize", "familiarise"
    R "fertilize", "fertilise"
    R "generalize", "generalise"
    R "hospitalize", "hospitalise"
    R "hypothesize", "hypothesise"
    R "idealize", "idealise"
    R "immunize", "immunise"
    R "itemize", "itemise"
    R "jeopardize", "jeopardise"
    R "liberalize", "liberalise"
    R "localize", "localise"
    R "marginalize", "marginalise"
    R "materialize", "materialise"
    R "mechanize", "mechanise"
    R "mobilize", "mobilise"
    R "monopolize", "monopolise"
    R "nationalize", "nationalise"
    R "penalize", "penalise"
    R "polarize", "polarise"
    R "privatize", "privatise"
    R "revolutionize", "revolutionise"
    R "scrutinize", "scrutinise"
    R "sensitize", "sensitise"
    R "socialize", "socialise"
    R "stabilize", "stabilise"
    R "sterilize", "sterilise"
    R "subsidize", "subsidise"
    R "terrorize", "terrorise"
    R "traumatize", "traumatise"
    R "trivialize", "trivialise"
    R "vandalize", "vandalise"
    R "vaporize", "vaporise"
    R "visualize", "visualise"

    ' ----- -or / -our -----
    R "color", "colour"
    R "favor", "favour"
    R "honor", "honour"
    R "humor", "humour"
    R "labor", "labour"
    R "neighbor", "neighbour"
    R "behavior", "behaviour"
    R "flavor", "flavour"
    R "harbor", "harbour"
    R "rumor", "rumour"
    R "tumor", "tumour"
    R "valor", "valour"
    R "vigor", "vigour"

    ' ----- -er / -re -----
    R "center", "centre"
    R "fiber", "fibre"
    R "liter", "litre"
    R "meter", "metre"
    R "theater", "theatre"

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

Private Sub R(us As String, uk As String)

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

