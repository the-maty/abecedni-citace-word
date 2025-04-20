Sub PrecislovatCitaceCesky()
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")

    ' Původní čísla → Nové podle českého řazení
    mapping.Add "11", "1"
    mapping.Add "1", "2"
    mapping.Add "6", "3"
    mapping.Add "12", "4"
    mapping.Add "13", "5"
    mapping.Add "8", "6"
    mapping.Add "10", "7"
    mapping.Add "19", "8"
    mapping.Add "21", "9"
    mapping.Add "7", "10"
    mapping.Add "2", "11"
    mapping.Add "15", "12"
    mapping.Add "3", "13"
    mapping.Add "16", "14"
    mapping.Add "20", "15"
    mapping.Add "18", "16"
    mapping.Add "14", "17"
    mapping.Add "17", "18"
    mapping.Add "9", "19"
    mapping.Add "22", "20"
    mapping.Add "4", "21"
    mapping.Add "5", "22"

    Dim key As Variant
    For Each key In mapping.Keys
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "{{" & key & "}}"   ' Vsechny jsem si precisloval jako {{1}}, {{2}} Taky by se dal udelat skriptik ale whatever :D 
            .Replacement.Text = "[" & mapping(key) & "]"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next key

    MsgBox "Citace přečíslovány podle českého abecedního pořadí!", vbInformation
End Sub
