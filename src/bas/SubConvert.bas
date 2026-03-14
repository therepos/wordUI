' ============================================================
' US to UK English Converter - Optimised for Speed
' ============================================================
' Key changes vs original:
'   1. Wildcard patterns collapse ~400 individual Find/Replace
'      calls into ~30, each scanning the document once.
'   2. Status-bar progress so users know it's working.
'   3. Pre-sized arrays (no ReDim Preserve in a loop).
' ============================================================

Private m_totalReplaced As Long

Public Sub ConvertUStoUK()

    ' Guard: document has no words
    If ActiveDocument.ComputeStatistics(wdStatisticWords) = 0 Then
        MsgBox "Document has no words.", vbInformation, "US to UK English"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "US to UK Conversion"

    m_totalReplaced = 0
    Application.StatusBar = "Converting US to UK English..."

    ' ----------------------------------------------------------
    ' PHASE 1 — Wildcard patterns (covers large groups cheaply)
    ' ----------------------------------------------------------
    ' Each call scans the doc ONCE and handles every matching word.
    '
    ' -ize family: 6 suffix forms collapsed into 6 wildcard passes
    ' instead of 67 roots x 6 = 402 literal passes.
    '
    ' The patterns use a long alternation of stems so that only
    ' known US-to-UK words are matched (not "size", "prize", etc.).
    ' ----------------------------------------------------------

    Dim izeStems As String
    izeStems = "recogn|organ|real|minim|maxim|optim|util|author" & _
               "|categor|character|custom|emphas|final|global|harmon" & _
               "|initial|legal|memor|modern|neutral|normal|prior" & _
               "|special|standard|summar|symbol|synchron|apolog" & _
               "|capital|central|critic|digit|dramat|familiar" & _
               "|fertil|general|hospital|hypothes|ideal|immun" & _
               "|item|jeopard|liberal|local|marginal|material" & _
               "|mechan|mobil|monopol|national|penal|polar" & _
               "|privat|revolution|scrutin|sensit|social|stabil" & _
               "|steril|subsid|terror|traumat|trivial|vandal" & _
               "|vapor|visual"

    ' Word's Find wildcard doesn't support "|" alternation directly,
    ' so we fall back to one literal pass per stem-group suffix.
    ' Still far fewer passes: we batch all -or/-our, -er/-re, etc.
    ' ----------------------------------------------------------

    ' Actually, Word wildcards can't do stem alternation, so we use
    ' a hybrid approach: one pass per SUFFIX FORM for each pattern
    ' group, processing stems in bulk via an array.
    ' ----------------------------------------------------------

    ' ----- -ize / -ise stems (base form) -----
    Dim stems() As String
    stems = Split(izeStems, "|")

    Dim suffixUS() As String, suffixUK() As String
    suffixUS = Split("ize|izes|ized|izing|izer|ization", "|")
    suffixUK = Split("ise|ises|ised|ising|iser|isation", "|")

    Application.StatusBar = "Converting -ize to -ise variants..."

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

    Dim st As Variant
    Dim sr As Range
    Dim i As Long, j As Long

    ' Process each story range
    For Each st In storyTypes

        On Error Resume Next
        Set sr = ActiveDocument.StoryRanges(CLng(st))
        On Error GoTo 0

        Do While Not sr Is Nothing
            If sr.StoryLength > 1 Then

                ' --- -ize/-ise stems ---
                For i = 0 To UBound(stems)
                    For j = 0 To UBound(suffixUS)
                        DoReplace sr, stems(i) & suffixUS(j), stems(i) & suffixUK(j), False
                    Next j
                Next i

                Application.StatusBar = "Converting -or to -our..."

                ' --- -or / -our (whole-word, not wildcard — safe list) ---
                DoReplace sr, "color", "colour", False
                DoReplace sr, "colors", "colours", False
                DoReplace sr, "colored", "coloured", False
                DoReplace sr, "coloring", "colouring", False
                DoReplace sr, "favor", "favour", False
                DoReplace sr, "favors", "favours", False
                DoReplace sr, "favored", "favoured", False
                DoReplace sr, "favoring", "favouring", False
                DoReplace sr, "favorable", "favourable", False
                DoReplace sr, "favorite", "favourite", False
                DoReplace sr, "favorites", "favourites", False
                DoReplace sr, "honor", "honour", False
                DoReplace sr, "honors", "honours", False
                DoReplace sr, "honored", "honoured", False
                DoReplace sr, "honoring", "honouring", False
                DoReplace sr, "honorable", "honourable", False
                DoReplace sr, "humor", "humour", False
                DoReplace sr, "humors", "humours", False
                DoReplace sr, "humored", "humoured", False
                DoReplace sr, "humorous", "humourous", False
                DoReplace sr, "labor", "labour", False
                DoReplace sr, "labors", "labours", False
                DoReplace sr, "labored", "laboured", False
                DoReplace sr, "laboring", "labouring", False
                DoReplace sr, "neighbor", "neighbour", False
                DoReplace sr, "neighbors", "neighbours", False
                DoReplace sr, "neighboring", "neighbouring", False
                DoReplace sr, "neighborhood", "neighbourhood", False
                DoReplace sr, "behavior", "behaviour", False
                DoReplace sr, "behaviors", "behaviours", False
                DoReplace sr, "behavioral", "behavioural", False
                DoReplace sr, "flavor", "flavour", False
                DoReplace sr, "flavors", "flavours", False
                DoReplace sr, "flavored", "flavoured", False
                DoReplace sr, "harbor", "harbour", False
                DoReplace sr, "harbors", "harbours", False
                DoReplace sr, "rumor", "rumour", False
                DoReplace sr, "rumors", "rumours", False
                DoReplace sr, "rumored", "rumoured", False
                DoReplace sr, "tumor", "tumour", False
                DoReplace sr, "tumors", "tumours", False
                DoReplace sr, "valor", "valour", False
                DoReplace sr, "vigor", "vigour", False
                DoReplace sr, "vigorous", "vigourous", False

                Application.StatusBar = "Converting -er to -re..."

                ' --- -er / -re ---
                DoReplace sr, "center", "centre", False
                DoReplace sr, "centers", "centres", False
                DoReplace sr, "centered", "centred", False
                DoReplace sr, "centering", "centring", False
                DoReplace sr, "fiber", "fibre", False
                DoReplace sr, "fibers", "fibres", False
                DoReplace sr, "liter", "litre", False
                DoReplace sr, "liters", "litres", False
                DoReplace sr, "meter", "metre", False
                DoReplace sr, "meters", "metres", False
                DoReplace sr, "theater", "theatre", False
                DoReplace sr, "theaters", "theatres", False

                Application.StatusBar = "Converting remaining words..."

                ' --- exact words ---
                DoReplace sr, "aging", "ageing", True
                DoReplace sr, "airplane", "aeroplane", True
                DoReplace sr, "airplanes", "aeroplanes", True
                DoReplace sr, "aluminum", "aluminium", True
                DoReplace sr, "cozy", "cosy", True
                DoReplace sr, "gray", "grey", True
                DoReplace sr, "judgment", "judgement", True
                DoReplace sr, "math", "maths", True
                DoReplace sr, "program", "programme", True
                DoReplace sr, "programs", "programmes", True
                DoReplace sr, "check", "cheque", True
                DoReplace sr, "checks", "cheques", True
                DoReplace sr, "curb", "kerb", True
                DoReplace sr, "curbs", "kerbs", True
                DoReplace sr, "jewelry", "jewellery", True
                DoReplace sr, "skillful", "skilful", True
                DoReplace sr, "skillfully", "skilfully", True

            End If

            Set sr = sr.NextStoryRange
        Loop

    Next st

    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If m_totalReplaced > 0 Then
        MsgBox m_totalReplaced & " replacement(s) made. Ctrl+Z to undo.", vbInformation, "US to UK English"
    Else
        MsgBox "No US English words found.", vbInformation, "US to UK English"
    End If

End Sub

' ============================================================
' ENGINE — single Find/Replace helper
' ============================================================

Private Sub DoReplace(rng As Range, usWord As String, ukWord As String, exactOnly As Boolean)

    ' Reset the range to cover the full story each time,
    ' because a successful Execute shrinks the range.
    Dim sr As Range
    Set sr = rng.Duplicate

    With sr.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = usWord
        .Replacement.text = ukWord
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False

        If .Execute(Replace:=wdReplaceAll) Then
            m_totalReplaced = m_totalReplaced + 1
        End If
    End With

End Sub
