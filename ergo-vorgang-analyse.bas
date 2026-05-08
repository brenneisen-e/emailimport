Attribute VB_Name = "ErgoVorgangAnalyse"
' ============================================================================
' ERGO VORGANG-ANALYSE - Excel-VBA-Tool (v1.0)
' ============================================================================
' Liest die Email-Vorgaenge aus dem Tabellenblatt "Emails" (Output von
' ergo-email-batch.bas / Ordner_Scannen) und laesst ErgoGPT jeden Vorgang
' klassifizieren. Schreibt die Ergebnisse als Zusatzspalten I-O zurueck.
'
' VORAUSSETZUNGEN:
'   1) Modul mit Funktion ASK_ErgoGPT(prompt) ist bereits in der Mappe
'      importiert (aus Test.txt im Repo).
'   2) Tabellenblatt "GPT" mit Konfig:
'        A6  = Modell-Name
'        A9  = Temperature (z.B. 0.0)
'        A12 = Tone
'        A20 = (optional) AI_RULES
'   3) Cookie-Datei vorhanden (Standard: F:\ExcelGPT-Cookie\Cookie.txt)
'   4) Tabellenblatt "Emails" mit Spalten A-H (Nr/Datum/Datum_ISO/
'      Absender_Name/Absender_Email/Betreff/Text/Anhaenge)
'
' WORKFLOW:
'   1. Outlook-Mails per ergo-email-batch.bas einlesen (Sheet "Emails")
'   2. Alt+F8 -> Vorgaenge_Analysieren -> Ausfuehren
'   3. Pro Zeile wird ErgoGPT befragt (~3-10s pro Zeile)
'   4. Spalten I-O werden gefuellt (Maklerpool, Makler-Name,
'      Klassifikation, Geschaefts-Typ, Unterlagen-Anfrage, Sonderfall)
'
' RESUME:
'   Bereits gefuellte Zeilen (Spalte I gefuellt) werden uebersprungen.
'   Mit Vorgaenge_Analysieren_Reset kannst du Spalten I-O leeren.
' ============================================================================

Option Explicit

Public Const SHEET_VORGANG As String = "Emails"

' Spaltenpositionen Output (1-basiert)
Private Const COL_MAKLERPOOL As Long = 9    ' I
Private Const COL_NACHNAME   As Long = 10   ' J
Private Const COL_VORNAME    As Long = 11   ' K
Private Const COL_KLASSIFIK  As Long = 12   ' L
Private Const COL_GESCHTYP   As Long = 13   ' M
Private Const COL_UNTERLAGEN As Long = 14   ' N
Private Const COL_SONDERFALL As Long = 15   ' O
Private Const COL_SPARTE     As Long = 16   ' P

' Pfad zur Cookie-Datei (wird an mehreren Stellen genutzt)
Private Const COOKIE_PATH As String = "F:\ExcelGPT-Cookie\Cookie.txt"

' === HAUPTEINSTIEG ==========================================================
Public Sub Vorgaenge_Analysieren()
    Dim ws As Worksheet
    On Error GoTo Fehler

    ' 0) Cookie-Setup pruefen / abfragen
    If Not PruefeCookieMitDialog() Then Exit Sub

    Set ws = HoleEmailsSheet()
    If ws Is Nothing Then Exit Sub

    Dim letzteZeile As Long
    letzteZeile = ws.cells(ws.Rows.count, 1).End(xlUp).row
    If letzteZeile < 2 Then
        MsgBox "Keine Daten im Tabellenblatt '" & SHEET_VORGANG & "' gefunden.", vbExclamation
        Exit Sub
    End If

    HeaderSchreiben ws

    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "Es werden " & (letzteZeile - 1) & " Zeilen analysiert." & vbCrLf & _
        "Pro Zeile dauert das ca. 3-10 Sekunden." & vbCrLf & vbCrLf & _
        "Geschaetzte Laufzeit: ~" & Int((letzteZeile - 1) * 6 / 60) + 1 & " Minuten." & vbCrLf & vbCrLf & _
        "Bereits gefuellte Zeilen (Spalte I gefuellt) werden uebersprungen." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Vorgang-Analyse starten")
    If antwort <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim verarbeitet As Long: verarbeitet = 0
    Dim uebersprungen As Long: uebersprungen = 0
    Dim fehler As Long: fehler = 0
    Dim row As Long

    For row = 2 To letzteZeile
        ' Skip wenn bereits analysiert
        If Trim(CStr(ws.cells(row, COL_MAKLERPOOL).Value)) <> "" Then
            uebersprungen = uebersprungen + 1
            GoTo NextRow
        End If

        Application.StatusBar = "Analysiere Zeile " & row & " von " & letzteZeile & _
                                " (verarbeitet: " & verarbeitet & ", Fehler: " & fehler & ")"
        DoEvents

        On Error Resume Next
        Err.Clear
        Dim ergebnis As Object
        Set ergebnis = AnalysiereEineZeile(ws, row)
        If Err.Number <> 0 Or ergebnis Is Nothing Then
            ws.cells(row, COL_MAKLERPOOL).Value = "[FEHLER]"
            ws.cells(row, COL_NACHNAME).Value = Err.Description
            fehler = fehler + 1
            Err.Clear
        Else
            SchreibeErgebnis ws, row, ergebnis
            verarbeitet = verarbeitet + 1
        End If
        On Error GoTo Fehler

        ' Alle 5 Zeilen Calc kurz updaten + Speichern
        If verarbeitet Mod 5 = 0 And verarbeitet > 0 Then
            Application.Calculation = xlCalculationAutomatic
            Application.Calculation = xlCalculationManual
            On Error Resume Next
            ThisWorkbook.Save
            On Error GoTo Fehler
        End If
NextRow:
    Next row

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Analyse abgeschlossen." & vbCrLf & vbCrLf & _
           "Verarbeitet: " & verarbeitet & vbCrLf & _
           "Uebersprungen (bereits analysiert): " & uebersprungen & vbCrLf & _
           "Fehler: " & fehler, _
           vbInformation, "Vorgaenge_Analysieren"
    Exit Sub

Fehler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

' === RESET (Spalten I-O leeren) =============================================
Public Sub Vorgaenge_Analysieren_Reset()
    Dim ws As Worksheet
    Set ws = HoleEmailsSheet()
    If ws Is Nothing Then Exit Sub

    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("Spalten I-O (Analyse-Ergebnisse) im Sheet '" & _
                     SHEET_VORGANG & "' wirklich loeschen?", _
                     vbYesNo + vbExclamation, "Reset")
    If antwort <> vbYes Then Exit Sub

    Dim letzteZeile As Long
    letzteZeile = ws.cells(ws.Rows.count, 1).End(xlUp).row
    If letzteZeile < 2 Then Exit Sub

    ws.Range(ws.cells(2, COL_MAKLERPOOL), ws.cells(letzteZeile, COL_SPARTE)).ClearContents
    MsgBox "Analyse-Spalten geleert.", vbInformation
End Sub

' === EINZELZEILE ANALYSIEREN ================================================
Private Function AnalysiereEineZeile(ws As Worksheet, row As Long) As Object
    Dim datum As String, absName As String, absMail As String
    Dim betreff As String, text As String, anhaenge As String

    datum = CStr(ws.cells(row, 2).Value)        ' B
    absName = CStr(ws.cells(row, 4).Value)      ' D
    absMail = CStr(ws.cells(row, 5).Value)      ' E
    betreff = CStr(ws.cells(row, 6).Value)      ' F
    text = CStr(ws.cells(row, 7).Value)         ' G
    anhaenge = CStr(ws.cells(row, 8).Value)     ' H

    ' Body kappen, damit Prompt nicht explodiert
    If Len(text) > 6000 Then text = Left(text, 6000) & " [...gekuerzt]"

    Dim prompt As String
    prompt = BuildVorgangPrompt(datum, absName, absMail, betreff, text, anhaenge)

    Dim antwort As String
    antwort = ASK_ErgoGPT(prompt)

    Set AnalysiereEineZeile = ParseGptJsonAntwort(antwort)
End Function

' === DER PROMPT (Kern des Tools!) ===========================================
Private Function BuildVorgangPrompt(datum As String, absName As String, absMail As String, _
                                     betreff As String, text As String, anhaenge As String) As String
    Dim p As String
    p = ""
    p = p & "Du bist ein erfahrener Sachbearbeiter im Innendienst eines deutschen" & vbCrLf
    p = p & "Komposit-Versicherers (SHUK: Sach/Haftpflicht/Unfall/Kfz)." & vbCrLf
    p = p & "Du bekommst eine Email aus dem Posteingang (typischerweise von einem" & vbCrLf
    p = p & "Versicherungsmakler oder Maklerpool) und klassifizierst den Vorgang." & vbCrLf
    p = p & vbCrLf
    p = p & "===== EMAIL =====" & vbCrLf
    p = p & "Datum: " & datum & vbCrLf
    p = p & "Absender: " & absName & " <" & absMail & ">" & vbCrLf
    p = p & "Betreff: " & betreff & vbCrLf
    p = p & "Anhaenge: " & IIf(Len(anhaenge) = 0, "(keine)", anhaenge) & vbCrLf
    p = p & vbCrLf
    p = p & "Body:" & vbCrLf
    p = p & "<<<" & vbCrLf
    p = p & text & vbCrLf
    p = p & ">>>" & vbCrLf
    p = p & "===== ENDE EMAIL =====" & vbCrLf
    p = p & vbCrLf
    p = p & "AUFGABE: Klassifiziere den Vorgang strikt anhand der unten" & vbCrLf
    p = p & "definierten Felder und Wertelisten." & vbCrLf
    p = p & vbCrLf
    p = p & "REGELN:" & vbCrLf
    p = p & "- Nutze die Email-Signatur, die From-Domain und den Body als Quellen." & vbCrLf
    p = p & "- Wenn Information nicht eindeutig ablesbar ist, nimm den leeren String" & vbCrLf
    p = p & "  bzw. den Default-Wert. Rate NIE." & vbCrLf
    p = p & "- Antworte AUSSCHLIESSLICH mit einem einzeiligen JSON-Objekt mit genau" & vbCrLf
    p = p & "  diesen 7 Schluesseln (alle Werte als String). Kein Markdown, keine" & vbCrLf
    p = p & "  Code-Fences, kein Vorwort, keine Erklaerung." & vbCrLf
    p = p & vbCrLf
    p = p & "FELDER:" & vbCrLf
    p = p & vbCrLf
    p = p & "1) maklerpool" & vbCrLf
    p = p & "   Welcher Maklerpool / welche Plattform steht hinter dem Absender?" & vbCrLf
    p = p & "   Erkennungs-Hinweise: Email-Domain, Footer/Disclaimer, Vertriebsorga." & vbCrLf
    p = p & "   Erlaubte Werte:" & vbCrLf
    p = p & "     'Fonds Finanz', 'BCA', 'JDC', 'blau direkt', 'Fondsnet'," & vbCrLf
    p = p & "     'Maxpool', 'WIFO', 'Aruna', 'Apella', 'Netfonds'," & vbCrLf
    p = p & "     'VEMA', 'Mr.Money', 'Direktmakler' (= eigenstaendiger Makler ohne Pool)," & vbCrLf
    p = p & "     '' (= unbekannt / Privatperson / Versicherer-intern)" & vbCrLf
    p = p & vbCrLf
    p = p & "2) makler_nachname" & vbCrLf
    p = p & "   Nachname des konkreten Bearbeiters/Maklers (aus Signatur oder" & vbCrLf
    p = p & "   Anschreiben). Bei 'Mustermann, Max' -> 'Mustermann'." & vbCrLf
    p = p & "   Default: ''" & vbCrLf
    p = p & vbCrLf
    p = p & "3) makler_vorname" & vbCrLf
    p = p & "   Vorname analog. Default: ''" & vbCrLf
    p = p & vbCrLf
    p = p & "4) klassifikation" & vbCrLf
    p = p & "   - 'Standardvorgang': Routine-Tagesgeschaeft (Antrag, Aenderungs-" & vbCrLf
    p = p & "     anzeige, Schadenmeldung, einfache Ruecksprache, Beleg-Nachreichung)." & vbCrLf
    p = p & "   - 'Nicht-Standard': Eskalation, Beschwerde, Sonderkonstrukt," & vbCrLf
    p = p & "     juristische Themen, ungewoehnliche Anliegen." & vbCrLf
    p = p & vbCrLf
    p = p & "5) geschaefts_typ" & vbCrLf
    p = p & "   Hinweis: 'BUe' (Bestandsuebertragung) bedeutet, dass Vertraege" & vbCrLf
    p = p & "   ohne inhaltliche Aenderung vom bisherigen Makler/Pool auf einen" & vbCrLf
    p = p & "   anderen Makler/Pool umgehaengt werden sollen (Maklerwechsel)." & vbCrLf
    p = p & "   In der Email ist meistens eine Vertragsnummer (VNR) genannt." & vbCrLf
    p = p & "   Entscheidend ist NICHT, ob eine VNR drin steht, sondern OB der" & vbCrLf
    p = p & "   Makler nur diesen einen Vertrag oder die gesamte Kundenverbindung" & vbCrLf
    p = p & "   (= alle Vertraege dieses Kunden) uebernehmen will." & vbCrLf
    p = p & "   - 'BUe einfacher Vertrag': Es soll explizit nur DER eine konkret" & vbCrLf
    p = p & "     genannte Vertrag (oder eine eng begrenzte Liste einzelner" & vbCrLf
    p = p & "     Vertraege) uebertragen werden. Andere Vertraege des Kunden" & vbCrLf
    p = p & "     bleiben unberuehrt." & vbCrLf
    p = p & "   - 'BUe Kundenverbindung': Die gesamte Kundenverbindung soll" & vbCrLf
    p = p & "     uebertragen werden, also alle Vertraege dieses Kunden -" & vbCrLf
    p = p & "     auch wenn als 'Aufhaenger' nur eine einzige VNR genannt ist." & vbCrLf
    p = p & "     Hinweise: Worte wie 'gesamte Kundenverbindung', 'alle" & vbCrLf
    p = p & "     Vertraege', 'kompletter Bestand des Kunden', 'Kundenwechsel'," & vbCrLf
    p = p & "     'wir uebernehmen den Kunden komplett'." & vbCrLf
    p = p & "   - 'Keine BUe': Der Vorgang ist KEINE Bestandsuebertragung," & vbCrLf
    p = p & "     sondern z.B. eine Schadenmeldung, ein Antrag, eine Vertrags-" & vbCrLf
    p = p & "     aenderung, eine Beleg-Nachreichung, eine Ruecksprache o.ae." & vbCrLf
    p = p & "   Erkennungs-Hinweise fuer BUe insgesamt: Worte wie 'Bestands-" & vbCrLf
    p = p & "   uebertragung', 'Maklerwechsel', 'BUe', 'Courtagezession'," & vbCrLf
    p = p & "   'Bestandsumdeckung', beigefuegte Maklervollmacht / Courtagezusage." & vbCrLf
    p = p & vbCrLf
    p = p & "6) unterlagen_angefragt" & vbCrLf
    p = p & "   - 'ja': Der Versicherer muss zusaetzliche Unterlagen anfordern" & vbCrLf
    p = p & "     ODER der Makler bittet darum, dass weitere Unterlagen nachgereicht" & vbCrLf
    p = p & "     werden ODER es ist offensichtlich, dass der Vorgang ohne weitere" & vbCrLf
    p = p & "     Unterlagen nicht abgeschlossen werden kann." & vbCrLf
    p = p & "   - 'nein': Vorgang ist mit den vorliegenden Infos/Anhaengen vollstaendig." & vbCrLf
    p = p & vbCrLf
    p = p & "7) sonderfall" & vbCrLf
    p = p & "   - 'Flottengeschaeft': Kfz-Flotten, mehrere Fahrzeuge eines" & vbCrLf
    p = p & "     Gewerbekunden, Flottentarif." & vbCrLf
    p = p & "   - 'Sondertarif': ungewoehnliche Tarifkonstruktion, Konsortium," & vbCrLf
    p = p & "     Grossrisiko, Industrie-Police, individueller Wording-Anhang." & vbCrLf
    p = p & "   - 'Kein Sonderfall': normaler Privat- oder Kleingewerbe-Vorgang." & vbCrLf
    p = p & vbCrLf
    p = p & "8) sparte" & vbCrLf
    p = p & "   Um welche Versicherungssparte geht es im Vorgang?" & vbCrLf
    p = p & "   - 'Komposit': SHUK-Komposit (Sach, Haftpflicht, Unfall, Kfz," & vbCrLf
    p = p & "     Hausrat, Wohngebaeude, Rechtsschutz, Tierhalter, Gewerbe etc.)." & vbCrLf
    p = p & "   - 'Leben': Lebensversicherung, Rentenversicherung, BU, Risiko-LV," & vbCrLf
    p = p & "     fondsgebundene LV." & vbCrLf
    p = p & "   - 'KV': Krankenversicherung (PKV, Zusatzversicherung, KT, KZV)." & vbCrLf
    p = p & "   - 'Mehrere': Vorgang erstreckt sich nachweislich auf mehrere der" & vbCrLf
    p = p & "     drei Sparten gleichzeitig (z.B. komplette Kundenverbindung mit" & vbCrLf
    p = p & "     Kfz + LV + KV)." & vbCrLf
    p = p & "   - 'Unbekannt': Sparte aus der Email nicht eindeutig erkennbar." & vbCrLf
    p = p & vbCrLf
    p = p & "AUSGABE-FORMAT (eine Zeile, gueltiges JSON):" & vbCrLf
    p = p & "{""maklerpool"":""..."",""makler_nachname"":""..."",""makler_vorname"":""...""," & vbCrLf
    p = p & " ""klassifikation"":""..."",""geschaefts_typ"":""..."",""unterlagen_angefragt"":""...""," & vbCrLf
    p = p & " ""sonderfall"":""..."",""sparte"":""...""}" & vbCrLf
    p = p & vbCrLf
    p = p & "Antworte JETZT, nur das JSON-Objekt:" & vbCrLf
    BuildVorgangPrompt = p
End Function

' === JSON-ANTWORT PARSEN ====================================================
Private Function ParseGptJsonAntwort(antwort As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim clean As String: clean = antwort
    ' Markdown-Code-Fences entfernen, falls das Modell sie doch ausgibt
    clean = Replace(clean, "```json", "")
    clean = Replace(clean, "```", "")
    clean = Trim(clean)

    ' Inhalt zwischen erster { und letzter } extrahieren (robust gegen Vorwort)
    Dim p1 As Long: p1 = InStr(clean, "{")
    Dim p2 As Long: p2 = InStrRev(clean, "}")
    If p1 = 0 Or p2 <= p1 Then
        Set ParseGptJsonAntwort = Nothing
        Exit Function
    End If
    clean = Mid(clean, p1, p2 - p1 + 1)

    Dim keys As Variant
    keys = Array("maklerpool", "makler_nachname", "makler_vorname", _
                 "klassifikation", "geschaefts_typ", "unterlagen_angefragt", _
                 "sonderfall", "sparte")

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        dict(keys(i)) = ExtrahiereJsonString(clean, CStr(keys(i)))
    Next i

    Set ParseGptJsonAntwort = dict
End Function

Private Function ExtrahiereJsonString(json As String, key As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    ' "key" : "value" - value darf escaped Quotes enthalten
    re.Pattern = """" & key & """\s*:\s*""((?:\\.|[^""\\])*)"""
    If re.test(json) Then
        Dim ms As Object: Set ms = re.Execute(json)
        Dim raw As String: raw = ms(0).SubMatches(0)
        ' Einfache JSON-Unescape
        raw = Replace(raw, "\""", """")
        raw = Replace(raw, "\\", "\")
        raw = Replace(raw, "\n", " ")
        raw = Replace(raw, "\r", " ")
        raw = Replace(raw, "\t", " ")
        ExtrahiereJsonString = Trim(raw)
    Else
        ExtrahiereJsonString = ""
    End If
End Function

' === SHEET HELPER ===========================================================
Private Function HoleEmailsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_VORGANG)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Tabellenblatt '" & SHEET_VORGANG & "' nicht gefunden." & vbCrLf & _
               "Bitte zuerst Mails per ergo-email-batch.bas (Ordner_Scannen / " & _
               "Outlook_Live_Abrufen) einlesen.", _
               vbExclamation
    End If
    Set HoleEmailsSheet = ws
End Function

Private Sub HeaderSchreiben(ws As Worksheet)
    ws.cells(1, COL_MAKLERPOOL).Value = "Maklerpool"
    ws.cells(1, COL_NACHNAME).Value = "Makler_Nachname"
    ws.cells(1, COL_VORNAME).Value = "Makler_Vorname"
    ws.cells(1, COL_KLASSIFIK).Value = "Klassifikation"
    ws.cells(1, COL_GESCHTYP).Value = "Geschaefts_Typ"
    ws.cells(1, COL_UNTERLAGEN).Value = "Unterlagen_Angefragt"
    ws.cells(1, COL_SONDERFALL).Value = "Sonderfall"
    ws.cells(1, COL_SPARTE).Value = "Sparte"

    With ws.Range(ws.cells(1, COL_MAKLERPOOL), ws.cells(1, COL_SPARTE))
        .Font.Bold = True
        .Interior.Color = RGB(30, 64, 175) ' kraeftiges Blau zur Abgrenzung von ERGO-rot
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(COL_MAKLERPOOL).ColumnWidth = 18
    ws.Columns(COL_NACHNAME).ColumnWidth = 18
    ws.Columns(COL_VORNAME).ColumnWidth = 14
    ws.Columns(COL_KLASSIFIK).ColumnWidth = 16
    ws.Columns(COL_GESCHTYP).ColumnWidth = 24
    ws.Columns(COL_UNTERLAGEN).ColumnWidth = 16
    ws.Columns(COL_SONDERFALL).ColumnWidth = 18
    ws.Columns(COL_SPARTE).ColumnWidth = 12
End Sub

Private Sub SchreibeErgebnis(ws As Worksheet, row As Long, dict As Object)
    ws.cells(row, COL_MAKLERPOOL).Value = SafeGet(dict, "maklerpool")
    ws.cells(row, COL_NACHNAME).Value = SafeGet(dict, "makler_nachname")
    ws.cells(row, COL_VORNAME).Value = SafeGet(dict, "makler_vorname")
    ws.cells(row, COL_KLASSIFIK).Value = SafeGet(dict, "klassifikation")
    ws.cells(row, COL_GESCHTYP).Value = SafeGet(dict, "geschaefts_typ")
    ws.cells(row, COL_UNTERLAGEN).Value = SafeGet(dict, "unterlagen_angefragt")
    ws.cells(row, COL_SONDERFALL).Value = SafeGet(dict, "sonderfall")
    ws.cells(row, COL_SPARTE).Value = SafeGet(dict, "sparte")

    ' Farbliche Hervorhebung: Sonderfall (rot) / Nicht-Standard (orange)
    Dim sonder As String: sonder = LCase(SafeGet(dict, "sonderfall"))
    Dim klass As String: klass = LCase(SafeGet(dict, "klassifikation"))

    If InStr(sonder, "flotten") > 0 Or InStr(sonder, "sondertarif") > 0 Then
        ws.cells(row, COL_SONDERFALL).Interior.Color = RGB(252, 165, 165) ' hellrot
    End If
    If InStr(klass, "nicht-standard") > 0 Then
        ws.cells(row, COL_KLASSIFIK).Interior.Color = RGB(253, 230, 138) ' hellgelb
    End If
End Sub

Private Function SafeGet(dict As Object, key As String) As String
    On Error Resume Next
    If dict.Exists(key) Then SafeGet = CStr(dict(key)) Else SafeGet = ""
End Function

' === COOKIE-DIALOG ==========================================================
Private Function PruefeCookieMitDialog() As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim hatDatei As Boolean: hatDatei = fso.FileExists(COOKIE_PATH)
    Dim alt As String: alt = ""

    If hatDatei Then
        On Error Resume Next
        Dim ts As Object
        Set ts = fso.OpenTextFile(COOKIE_PATH, 1, False, -2)
        alt = Trim(ts.ReadAll)
        ts.Close
        On Error GoTo 0
    End If

    Dim msg As String, ans As VbMsgBoxResult
    If hatDatei And Len(alt) > 0 Then
        Dim preview As String: preview = Left(alt, 100)
        If Len(alt) > 100 Then preview = preview & "..."
        msg = "Cookie-Datei gefunden:" & vbCrLf & COOKIE_PATH & vbCrLf & vbCrLf & _
              "Aktueller Inhalt (Auszug):" & vbCrLf & preview & vbCrLf & vbCrLf & _
              "Cookies laufen periodisch ab. Wenn du heute bei gpt.ergo.com schon" & vbCrLf & _
              "eingeloggt warst, ist der Cookie wahrscheinlich noch gueltig." & vbCrLf & vbCrLf & _
              "JA      = Cookie-Datei verwenden (weiter)" & vbCrLf & _
              "NEIN    = Cookie neu eingeben (z.B. wenn 401-Fehler kommt)" & vbCrLf & _
              "ABBRECH = Beenden"
        ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "Cookie-Setup")
        If ans = vbCancel Then PruefeCookieMitDialog = False: Exit Function
        If ans = vbYes Then PruefeCookieMitDialog = True: Exit Function
    Else
        msg = "Es ist noch keine Cookie-Datei vorhanden:" & vbCrLf & COOKIE_PATH & vbCrLf & vbCrLf & _
              "JA  = Cookie jetzt eingeben (wird automatisch in die Datei geschrieben)" & vbCrLf & _
              "NEIN = Beenden"
        ans = MsgBox(msg, vbYesNo + vbQuestion, "Cookie fehlt")
        If ans <> vbYes Then PruefeCookieMitDialog = False: Exit Function
    End If

    ' Manuelle Eingabe
    Dim cookie As String
    cookie = InputBox( _
        "Cookie-String einfuegen." & vbCrLf & vbCrLf & _
        "So bekommst du den Cookie:" & vbCrLf & _
        "  1) gpt.ergo.com im Browser oeffnen + einloggen" & vbCrLf & _
        "  2) F12 (DevTools) -> Tab 'Network' -> beliebigen Request anklicken" & vbCrLf & _
        "  3) Headers -> Request Headers -> Wert hinter 'Cookie:' kopieren" & vbCrLf & _
        "  4) Hier einfuegen und OK." & vbCrLf & vbCrLf & _
        "Hinweis: Sehr lange Strings sind normal.", _
        "Cookie eingeben")
    cookie = Trim(cookie)
    If Len(cookie) = 0 Then
        MsgBox "Kein Cookie eingegeben - Abbruch.", vbInformation
        PruefeCookieMitDialog = False
        Exit Function
    End If

    ' Verzeichnis sicherstellen
    Dim dirPath As String: dirPath = Left(COOKIE_PATH, InStrRev(COOKIE_PATH, "\") - 1)
    If Len(dirPath) > 0 And Not fso.FolderExists(dirPath) Then
        On Error Resume Next
        fso.CreateFolder dirPath
        If Err.Number <> 0 Then
            MsgBox "Verzeichnis konnte nicht erstellt werden:" & vbCrLf & dirPath & vbCrLf & _
                   "Bitte manuell anlegen und Schreibrechte pruefen.", vbCritical
            Err.Clear
            PruefeCookieMitDialog = False
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If

    ' Cookie schreiben (UTF-8 ohne BOM via ADODB.Stream waere overkill - ASCII reicht)
    On Error Resume Next
    Dim out As Object
    Set out = fso.CreateTextFile(COOKIE_PATH, True, False) ' overwrite, ASCII
    If Err.Number <> 0 Then
        MsgBox "Cookie-Datei konnte nicht geschrieben werden:" & vbCrLf & COOKIE_PATH & vbCrLf & _
               Err.Description, vbCritical
        Err.Clear
        PruefeCookieMitDialog = False
        On Error GoTo 0
        Exit Function
    End If
    out.Write cookie
    out.Close
    On Error GoTo 0

    MsgBox "Cookie gespeichert. Analyse startet jetzt.", vbInformation
    PruefeCookieMitDialog = True
End Function

' === SETUP / ANLEITUNG IN EXCEL =============================================
Public Sub Vorgaenge_Setup()
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "Dieses Setup richtet die Excel-Mappe fuer die Vorgang-Analyse ein:" & vbCrLf & vbCrLf & _
        "  - Tabellenblatt 'GPT' mit Default-Konfig anlegen (falls fehlt)" & vbCrLf & _
        "  - Tabellenblatt 'Anleitung' mit Schritt-fuer-Schritt-Beschreibung" & vbCrLf & _
        "  - Cookie-Setup pruefen (Dialog kommt im Anschluss)" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Vorgaenge_Setup")
    If antwort <> vbYes Then Exit Sub

    SetupGptSheet
    SetupAnleitungSheet

    ' Cookie-Dialog
    PruefeCookieMitDialog

    On Error Resume Next
    ThisWorkbook.Sheets("Anleitung").Activate
    On Error GoTo 0
End Sub

Private Sub SetupGptSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GPT")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "GPT"
    End If

    ' Header in Spalte B (Beschriftung), Werte in A
    If Trim(CStr(ws.Range("B6").Value)) = "" Then ws.Range("B6").Value = "<- Modell-Name (z.B. gpt-5.1, gpt-4o)"
    If Trim(CStr(ws.Range("B7").Value)) = "" Then ws.Range("B7").Value = "<- (nicht genutzt - Cookie kommt aus File)"
    If Trim(CStr(ws.Range("B9").Value)) = "" Then ws.Range("B9").Value = "<- Temperature (0.0 = deterministisch)"
    If Trim(CStr(ws.Range("B12").Value)) = "" Then ws.Range("B12").Value = "<- Tone (z.B. Sachlich, oder leer)"
    If Trim(CStr(ws.Range("B20").Value)) = "" Then ws.Range("B20").Value = "<- (optional) AI_RULES fuer AutoCode"

    ' Defaults nur setzen, wenn leer
    If Trim(CStr(ws.Range("A6").Value)) = "" Then ws.Range("A6").Value = "gpt-5.1"
    If Trim(CStr(ws.Range("A9").Value)) = "" Then ws.Range("A9").Value = 0
    If Trim(CStr(ws.Range("A12").Value)) = "" Then ws.Range("A12").Value = "Sachlich"

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 50
End Sub

Private Sub SetupAnleitungSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Anleitung")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.name = "Anleitung"
    Else
        ws.cells.Clear
    End If

    Dim r As Long: r = 1
    ws.cells(r, 1).Value = "ERGO Vorgang-Analyse - Anleitung": Bold ws, r, True, 16: r = r + 2

    ws.cells(r, 1).Value = "VORAUSSETZUNGEN": Bold ws, r, True, 12: r = r + 1
    ws.cells(r, 1).Value = "1. Modul ASK_ErgoGPT (aus Test.txt im Repo) ist importiert.": r = r + 1
    ws.cells(r, 1).Value = "2. Tabellenblatt 'GPT' mit:": r = r + 1
    ws.cells(r, 1).Value = "     A6  = Modell-Name  (Standard: gpt-5.1)": r = r + 1
    ws.cells(r, 1).Value = "     A9  = Temperature  (Standard: 0)": r = r + 1
    ws.cells(r, 1).Value = "     A12 = Tone         (z.B. 'Sachlich')": r = r + 1
    ws.cells(r, 1).Value = "3. Cookie-Datei: " & COOKIE_PATH: r = r + 1
    ws.cells(r, 1).Value = "   (kann beim Start des Tools auch manuell eingegeben werden)": r = r + 2

    ws.cells(r, 1).Value = "WORKFLOW": Bold ws, r, True, 12: r = r + 1
    ws.cells(r, 1).Value = "Schritt 1: Mails einlesen": Bold ws, r, True, 11: r = r + 1
    ws.cells(r, 1).Value = "  Alt+F8 -> 'Outlook_Live_Abrufen' oder 'Ordner_Scannen' aus": r = r + 1
    ws.cells(r, 1).Value = "  ergo-email-batch.bas. Dadurch entsteht das Tabellenblatt 'Emails'.": r = r + 2

    ws.cells(r, 1).Value = "Schritt 2: Vorgang-Analyse starten": Bold ws, r, True, 11: r = r + 1
    ws.cells(r, 1).Value = "  Alt+F8 -> 'Vorgaenge_Analysieren' -> Ausfuehren": r = r + 1
    ws.cells(r, 1).Value = "  -> Cookie-Dialog erscheint: Datei nutzen ODER Cookie neu eingeben.": r = r + 1
    ws.cells(r, 1).Value = "  -> Bestaetigung mit geschaetzter Laufzeit erscheint.": r = r + 1
    ws.cells(r, 1).Value = "  -> Pro Zeile ca. 3-10 Sekunden. Bei jeder 5. Zeile wird automatisch": r = r + 1
    ws.cells(r, 1).Value = "     zwischengespeichert.": r = r + 2

    ws.cells(r, 1).Value = "Schritt 3: Ergebnisse pruefen": Bold ws, r, True, 11: r = r + 1
    ws.cells(r, 1).Value = "  Spalten I-P im Sheet 'Emails' sind gefuellt:": r = r + 1
    ws.cells(r, 1).Value = "    I  Maklerpool         (Fonds Finanz / BCA / JDC / blau direkt / ...)": r = r + 1
    ws.cells(r, 1).Value = "    J  Makler_Nachname": r = r + 1
    ws.cells(r, 1).Value = "    K  Makler_Vorname": r = r + 1
    ws.cells(r, 1).Value = "    L  Klassifikation     (Standardvorgang / Nicht-Standard)": r = r + 1
    ws.cells(r, 1).Value = "    M  Geschaefts_Typ     (BUe einfacher Vertrag / BUe Kundenverbindung / Keine BUe)": r = r + 1
    ws.cells(r, 1).Value = "    N  Unterlagen_Angefragt (ja / nein)": r = r + 1
    ws.cells(r, 1).Value = "    O  Sonderfall         (Flottengeschaeft / Sondertarif / Kein Sonderfall)": r = r + 1
    ws.cells(r, 1).Value = "    P  Sparte             (Komposit / Leben / KV / Mehrere / Unbekannt)": r = r + 2

    ws.cells(r, 1).Value = "WIEDER-AUFNAHME": Bold ws, r, True, 12: r = r + 1
    ws.cells(r, 1).Value = "Schon analysierte Zeilen (Spalte I gefuellt) werden uebersprungen.": r = r + 1
    ws.cells(r, 1).Value = "Du kannst Vorgaenge_Analysieren also beliebig oft starten -": r = r + 1
    ws.cells(r, 1).Value = "die Analyse macht da weiter, wo sie aufgehoert hat.": r = r + 2

    ws.cells(r, 1).Value = "ZURUECKSETZEN": Bold ws, r, True, 12: r = r + 1
    ws.cells(r, 1).Value = "Alt+F8 -> 'Vorgaenge_Analysieren_Reset' loescht die Spalten I-P.": r = r + 2

    ws.cells(r, 1).Value = "DOMAIN-LOGIK BUe": Bold ws, r, True, 12: r = r + 1
    ws.cells(r, 1).Value = "BUe = Bestandsuebertragung. Makler A uebernimmt Bestand des Kunden": r = r + 1
    ws.cells(r, 1).Value = "von Makler B. Eine VNR ist meistens in der Email genannt - entscheidend": r = r + 1
    ws.cells(r, 1).Value = "ist aber, ob NUR DIESER Vertrag oder die GANZE Kundenverbindung": r = r + 1
    ws.cells(r, 1).Value = "(alle Vertraege des Kunden) uebertragen werden soll.": r = r + 2

    ws.cells(r, 1).Value = "PROMPT ANPASSEN": Bold ws, r, True, 12: r = r + 1
    ws.cells(r, 1).Value = "VBA-Editor (Alt+F11) -> Modul ErgoVorgangAnalyse -> BuildVorgangPrompt.": r = r + 1
    ws.cells(r, 1).Value = "Dort koennen Maklerpool-Liste, Wertelisten und Definitionen direkt": r = r + 1
    ws.cells(r, 1).Value = "editiert werden.": r = r + 2

    ws.Columns("A").ColumnWidth = 110
End Sub

Private Sub Bold(ws As Worksheet, row As Long, b As Boolean, fontSize As Long)
    With ws.cells(row, 1).Font
        .Bold = b
        .Size = fontSize
    End With
End Sub
