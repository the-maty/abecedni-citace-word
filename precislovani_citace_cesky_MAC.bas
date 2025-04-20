Sub PrecislovatCitaceCesky_Mac()

    Dim keys As Variant
    Dim values As Variant
    Dim i As Integer

    ' Ručně vytvořená pole pro převod
    keys = Array("11", "1", "6", "12", "13", "8", "10", "19", "21", "7", "2", "15", "3", "16", "20", "18", "14", "17", "9", "22", "4", "5")
    values = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22")

    For i = LBound(keys) To UBound(keys)
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "{{" & keys(i) & "}}"
            .Replacement.Text = "[" & values(i) & "]"
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
    Next i

    MsgBox "Citace přečíslovány (macOS verze).", vbInformation

End Sub
