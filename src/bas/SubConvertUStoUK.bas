Attribute VB_Name = "SubConvertUStoUK"
' ============================================================
' US to UK English Converter (Smart Inflection)
' ============================================================
' INSTALL: Word > Alt+F11 > Insert > Module > Paste this code
' RUN:     Alt+F8 > ConvertUStoUK
' ADD:     Just add the ROOT form:  R "organize", "organise"
'          The code auto-handles: organizes, organized, organizing,
'          organizer, organizers, organization, organizations, etc.
'
' Case is automatic — Color -> Colour, COLOR -> COLOUR, etc.
' Ctrl+Z to undo everything in one step.
' ============================================================

Public Sub ConvertUStoUK()
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "US to UK Conversion"
    
    ' ----- -ize / -ise (auto-covers -izes, -ized, -izing, -izer, -ization) -----
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
    
    ' ----- -or / -our (auto-covers -ors, -ored, -oring, -orous, -orable, -orful, -orless) -----
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
    R "armor", "armour"
    R "candor", "candour"
    R "clamor", "clamour"
    R "endeavor", "endeavour"
    R "fervor", "fervour"
    R "odor", "odour"
    R "parlor", "parlour"
    R "rancor", "rancour"
    R "rigor", "rigour"
    R "savior", "saviour"
    R "splendor", "splendour"
    R "demeanor", "demeanour"
    
    ' ----- -er / -re (auto-covers -ers, -ered) -----
    R "center", "centre"
    R "fiber", "fibre"
    R "liter", "litre"
    R "meter", "metre"
    R "theater", "theatre"
    R "caliber", "calibre"
    R "scepter", "sceptre"
    R "somber", "sombre"
    R "specter", "spectre"
    R "luster", "lustre"
    R "meager", "meagre"
    R "saber", "sabre"
    
    ' ----- -ense / -ence (auto-covers -enses, -enseless, -ensible) -----
    R "defense", "defence"
    R "offense", "offence"
    R "license", "licence"
    R "pretense", "pretence"
    
    ' ----- -og / -ogue (auto-covers -ogs, -oged, -oging) -----
    R "analog", "analogue"
    R "catalog", "catalogue"
    R "dialog", "dialogue"
    R "monolog", "monologue"
    R "prolog", "prologue"
    R "epilog", "epilogue"
    
    ' ----- -yze / -yse (auto-covers -yzed, -yzing, -yzes, -yzer) -----
    R "analyze", "analyse"
    R "paralyze", "paralyse"
    R "catalyze", "catalyse"
    
    ' ----- doubled consonants (auto-covers -ed, -ing, -er, -ers) -----
    R "travel", "travell"
    R "cancel", "cancell"
    R "model", "modell"
    R "label", "labell"
    R "level", "levell"
    R "fuel", "fuell"
    R "counsel", "counsell"
    R "dial", "diall"
    
    ' ----- exact matches only (no inflection) -----
    X "aging", "ageing"
    X "airplane", "aeroplane"
    X "airplanes", "aeroplanes"
    X "aluminum", "aluminium"
    X "artifact", "artefact"
    X "artifacts", "artefacts"
    X "cozy", "cosy"
    X "donut", "doughnut"
    X "donuts", "doughnuts"
    X "gray", "grey"
    X "inquire", "enquire"
    X "inquired", "enquired"
    X "inquiring", "enquiring"
    X "inquiry", "enquiry"
    X "inquiries", "enquiries"
    X "judgment", "judgement"
    X "judgments", "judgements"
    X "maneuver", "manoeuvre"
    X "maneuvers", "manoeuvres"
    X "maneuvered", "manoeuvred"
    X "maneuvering", "manoeuvring"
    X "mold", "mould"
    X "molds", "moulds"
    X "molded", "moulded"
    X "molding", "moulding"
    X "molt", "moult"
    X "molts", "moults"
    X "molted", "moulted"
    X "molting", "moulting"
    X "mustache", "moustache"
    X "mustaches", "moustaches"
    X "pajamas", "pyjamas"
    X "plow", "plough"
    X "plows", "ploughs"
    X "plowed", "ploughed"
    X "plowing", "ploughing"
    X "skeptic", "sceptic"
    X "skeptics", "sceptics"
    X "skeptical", "sceptical"
    X "skepticism", "scepticism"
    X "tire", "tyre"
    X "tires", "tyres"
    X "fetus", "foetus"
    X "diarrhea", "diarrhoea"
    X "anemia", "anaemia"
    X "anemic", "anaemic"
    X "anesthetic", "anaesthetic"
    X "anesthetics", "anaesthetics"
    X "archeology", "archaeology"
    X "archeological", "archaeological"
    X "estrogen", "oestrogen"
    X "pediatric", "paediatric"
    X "pediatrics", "paediatrics"
    X "pediatrician", "paediatrician"
    X "leukemia", "leukaemia"
    X "math", "maths"
    X "program", "programme"
    X "programs", "programmes"
    X "check", "cheque"
    X "checks", "cheques"
    X "checkbook", "chequebook"
    X "curb", "kerb"
    X "curbs", "kerbs"
    X "draft", "draught"
    X "drafts", "draughts"
    X "drafty", "draughty"
    X "enrollment", "enrolment"
    X "enrollments", "enrolments"
    X "enroll", "enrol"
    X "enrolls", "enrols"
    X "fulfill", "fulfil"
    X "fulfills", "fulfils"
    X "fulfillment", "fulfilment"
    X "instill", "instil"
    X "instills", "instils"
    X "installment", "instalment"
    X "installments", "instalments"
    X "jewelry", "jewellery"
    X "skillful", "skilful"
    X "skillfully", "skilfully"
    X "willful", "wilful"
    X "willfully", "wilfully"
    
    ' ===== ADD YOUR OWN HERE =====
    ' R "usword", "ukword"      <- smart: auto-handles inflections
    ' X "usword", "ukword"      <- exact: only replaces this specific word
    
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "Done. Ctrl+Z to undo.", vbInformation, "US to UK English"
End Sub

' ============================================================
' ENGINE (no need to edit below here)
' ============================================================

' Smart replace — detects the pattern and auto-generates inflected forms
Private Sub R(us As String, uk As String)
    ' Always do the root word itself
    X us, uk
    
    ' -ize / -ise pattern
    If EndsWith(us, "ize") And EndsWith(uk, "ise") Then
        Dim izeStem As String: izeStem = Left(us, Len(us) - 3)
        Dim iseStem As String: iseStem = Left(uk, Len(uk) - 3)
        X izeStem & "izes", iseStem & "ises"
        X izeStem & "ized", iseStem & "ised"
        X izeStem & "izing", iseStem & "ising"
        X izeStem & "izer", iseStem & "iser"
        X izeStem & "izers", iseStem & "isers"
        X izeStem & "ization", iseStem & "isation"
        X izeStem & "izations", iseStem & "isations"
        Exit Sub
    End If
    
    ' -yze / -yse pattern
    If EndsWith(us, "yze") And EndsWith(uk, "yse") Then
        Dim yzeStem As String: yzeStem = Left(us, Len(us) - 3)
        Dim yseStem As String: yseStem = Left(uk, Len(uk) - 3)
        X yzeStem & "yzes", yseStem & "yses"
        X yzeStem & "yzed", yseStem & "ysed"
        X yzeStem & "yzing", yseStem & "ysing"
        X yzeStem & "yzer", yseStem & "yser"
        X yzeStem & "yzers", yseStem & "ysers"
        Exit Sub
    End If
    
    ' -or / -our pattern
    If EndsWith(us, "or") And EndsWith(uk, "our") Then
        X us & "s", uk & "s"
        X us & "ed", uk & "ed"
        X us & "ing", uk & "ing"
        X us & "ful", uk & "ful"
        X us & "fully", uk & "fully"
        X us & "less", uk & "less"
        X us & "able", uk & "able"
        X us & "ous", uk & "ous"
        X us & "ite", uk & "ite"
        X us & "ites", uk & "ites"
        X us & "al", uk & "al"
        Exit Sub
    End If
    
    ' -er / -re pattern
    If EndsWith(us, "er") And EndsWith(uk, "re") Then
        X us & "s", uk & "s"
        X us & "ed", uk & "d"
        Exit Sub
    End If
    
    ' -ense / -ence pattern
    If EndsWith(us, "ense") And EndsWith(uk, "ence") Then
        Dim enseStem As String: enseStem = Left(us, Len(us) - 4)
        Dim enceStem As String: enceStem = Left(uk, Len(uk) - 4)
        X enseStem & "enses", enceStem & "ences"
        X enseStem & "enseless", enceStem & "enceless"
        X enseStem & "ensible", enceStem & "encible"
        Exit Sub
    End If
    
    ' -og / -ogue pattern
    If EndsWith(us, "og") And EndsWith(uk, "ogue") Then
        X us & "s", uk & "s"
        X us & "ed", uk & "d"
        X us & "ing", uk & "ing"
        Exit Sub
    End If
    
    ' doubled consonant pattern (e.g. travel -> travell)
    ' The root won't match whole words usefully, but the suffixes will
    If Len(uk) = Len(us) + 1 And Left(uk, Len(us)) = us Then
        X us & "ed", uk & "ed"
        X us & "ing", uk & "ing"
        X us & "er", uk & "er"
        X us & "ers", uk & "ers"
        X us & "or", uk & "or"
        X us & "ors", uk & "ors"
        ' Don't replace the bare root (travel is valid in both)
        Exit Sub
    End If
End Sub

' Exact replace — single word, no inflection
Private Sub X(us As String, uk As String)
    Dim sr As Range
    For Each sr In ActiveDocument.StoryRanges
        Do
            With sr.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .text = us
                .Replacement.text = uk
                .MatchWholeWord = True
                .MatchCase = False
                .Forward = True
                .Wrap = wdFindStop
                .Execute Replace:=wdReplaceAll
            End With
            Set sr = sr.NextStoryRange
        Loop Until sr Is Nothing
    Next sr
End Sub

Private Function EndsWith(s As String, suffix As String) As Boolean
    If Len(s) >= Len(suffix) Then
        EndsWith = (Right(s, Len(suffix)) = suffix)
    End If
End Function

