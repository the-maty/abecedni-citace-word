# 🔄 Převod číslovaných citací pomocí VBA makra ve Wordu

Tento návod ti ukáže, jak jednoduše přečíslovat zdroje ve Word dokumentu podle nového (např. abecedního) pořadí pomocí VBA makra. Funguje na Windows i macOS.

---

## 📁 Struktura citací v dokumentu

V dokumentu musí být citace zapsané dočasně v tomto tvaru:

```text
{{1}}, {{2}}, {{3}}, ...
```

Například:
```text
Jak uvádí {{3}}, databáze jsou klíčové...
```

Po spuštění makra se tyto značky automaticky přepíšou na:

```text
[13], [5], [21], ...
```

---

## 🪟 Windows: Jak spustit VBA makro

### 1. Otevři Word dokument

### 2. Stiskni `Alt + F11`
- Tím se otevře editor jazyka **Visual Basic for Applications** (VBA).

### 3. Vlevo v panelu „Project“
- Pravým klikni na `Normal` nebo název dokumentu
- Zvol **Insert → Module**

### 4. Vlož kód makra
- Zkopíruj a vlož celý kód níže do okna modulu:

```vba
Sub PrecislovatCitaceCesky()
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")

    mapping.Add "11", "1" ' První reprezentuje {{11}} druhý nový [1]
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
            .Text = "{{" & key & "}}"
            .Replacement.Text = "[" & mapping(key) & "]"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next key

    MsgBox "Citace přečíslovány podle českého abecedního pořadí!", vbInformation
End Sub
```

### 5. Zavři editor (Alt + Q)

### 6. Stiskni `Alt + F8`
- Vyber makro `PrecislovatCitaceCesky`
- Klikni na **Spustit**

---

## 🍎 macOS: Jak spustit VBA makro

### 1. Otevři Word dokument

### 2. Horní lišta → **Nástroje → Editor maker**  
_(v angličtině Tools → Visual Basic Editor)_

### 3. V levém panelu „Project“:
- Pravým klikni na `Normal` nebo název dokumentu
- Zvol **Insert → Module**

### 4. Vlož MAC-kompatibilní kód (bez Windows-only objektů):

```vba
Sub PrecislovatCitaceCesky_Mac()
    Dim keys As Variant
    Dim values As Variant
    Dim i As Integer

    keys = Array("11", "1", "6", "12", "13", "8", "10", "19", "21", "7", "2", "15", "3", "16", "20", "18", "14", "17", "9", "22", "4", "5")    ' První reprezentuje {{11}}
    values = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22") ' Druhý reprezentuje nový [1]

    For i = LBound(keys) To UBound(keys)
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "{{" & keys(i) & "}}"
            .Replacement.Text = "[" & values(i) & "]"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    MsgBox "Citace přečíslovány (macOS verze).", vbInformation
End Sub
```

### 5. Zavři editor (`Cmd + W`)

### 6. Spusť makro:
- Horní lišta: **Nástroje → Makro → Makra…**
- Vyber `PrecislovatCitaceCesky_Mac`
- Klikni **Spustit**

---

## 🧼 Tipy na závěr

- Doporučuji si před spuštěním makra **uložit zálohu dokumentu**.
- Funkční v rámci ruční citace [1], [2],...

---

> Vytvořeno s láskou ♥️ xdd

