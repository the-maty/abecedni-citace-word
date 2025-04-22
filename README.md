# 📚 Citace Sorter – přečíslování zdrojů ve Wordu

## 🚀 Jak na to? – (Doporučený postup)

🔄 **Krok 1:** Zkopíruj si bibliografické odkazy ještě s  `[1]` označením než je převedeme na `{{1}}`

🤖 **Krok 2:** Použij full automatický nástroj pro vygenerování skriptu:

👉 **[➡️ Otevřít tool ⬅️](https://the-maty.github.io/citace-sorter/)**

**(modul si jen připrav spustit až na konec!)**

🔄 **Krok 3:** Spusť makro `PrevestNaZavorky` ve Wordu – převede `[1]` → `{{1}}`

🤖 **Krok 4:** Spusť makro jež bylo vygenerování zde 👉 **[➡️ Otevřít tool ⬅️](https://the-maty.github.io/citace-sorter/)**:

---

## 🔁 Pomocný skript: převod z `[x]` na `{{x}}`

Pokud máš v dokumentu citace ve formátu `[1]`, `[2]`, atd., můžeš je jednoduše převést zpět na `{{1}}`, `{{2}}` pomocí tohoto VBA makra:

```vba
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
```

---

## 📁 Struktura citací v dokumentu

Citace musí být dočasně zapsané v tomto tvaru:

```text
{{1}}, {{2}}, {{3}}, ...
```

Například:
```text
Jak uvádí {{3}}, databáze jsou klíčové...
```

Po spuštění makra pro přečíslování se tyto značky automaticky přepíšou na:

```text
[13], [5], [21], ...
```

---

## 🪟 Windows: Jak spustit VBA makro

### 1. Otevři Word dokument  
### 2. Stiskni `Alt + F11` – otevře se VBA editor  
### 3. Vlevo v panelu „Project“:
- Pravým klikni na `Normal` nebo název dokumentu  
- Zvol **Insert → Module**

### 4. Vlož kód makra:

```vba
Sub PrecislovatCitaceCesky()
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")

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
            .Text = "{{" & key & "}}"
            .Replacement.Text = "[" & mapping(key) & "]"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next key

    MsgBox "Citace zmeneny na pozadovany format z {{cislo}} na [cislo] podle pozadovaneho poradi", vbInformation
End Sub
```

### 5. Zavři editor (`Alt + Q`)  
### 6. Spusť makro pomocí `Alt + F8`  
- Vyber `PrecislovatCitaceCesky`  
- Klikni **Spustit**

---

## 🍎 macOS: Jak spustit VBA makro

### 1. Otevři Word dokument  
### 2. Horní lišta → **Nástroje → Editor maker**  
_(v angličtině Tools → Visual Basic Editor)_

### 3. V levém panelu „Project“:
- Pravým klikni na `Normal` nebo název dokumentu  
- Zvol **Insert → Module**

### 4. Vlož MAC-kompatibilní kód:

```vba
Sub PrecislovatCitaceCesky_Mac()
    Dim keys As Variant
    Dim values As Variant
    Dim i As Integer

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
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    MsgBox "Citace přečíslovány.", vbInformation
End Sub
```

### 5. Zavři editor (`Cmd + W`)  
### 6. Spusť makro:
- Horní lišta: **Nástroje → Makro → Makra…**
- Vyber `PrecislovatCitaceCesky_Mac`
- Klikni **Spustit**

---

## 🫼 Tipy na závěr

- 💾 **Před spuštěním si ulož zálohu dokumentu**
- ✅ Tool funguje pro ručně psané citace ve Wordu ve formátu `[1]`, `[2]`, …
- 🌐 Full automatický nástroj: [citace-sorter](https://the-maty.github.io/citace-sorter/)

---

> Created by MaTy ♥️

