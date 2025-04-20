Sub PrevestNaZavorky()
    Dim i As Integer

    ' Upravit dle počtu citací – např. 1 až 99
    For i = 1 To 99
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "[" & i & "]"
            .Replacement.Text = "{{" & i & "}}"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    MsgBox "Hotovo! Všechny citace převedeny na {{x}} formát.", vbInformation
End Sub