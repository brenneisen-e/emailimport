Attribute VB_Name = "ErgoVorgangAnalyse"
' ============================================================================
' ERGO VORGANG-ANALYSE - Excel-VBA-Tool (v2.9)
' ============================================================================
' v2.9: NEUE Spalte Y 'Ist_Reminder' (ja/nein/nicht_pruefbar) - erkennt
'       Erinnerungen / Mahnungen / Wiedervorlagen anhand von Betreff-/
'       Body-Indikatoren (Erinnerung, Reminder, WV, Sachstandsanfrage,
'       'wir warten noch', '2. Mahnung', etc.). Unabhaengig von Vorgangs-
'       typ - ein Reminder kann gleichzeitig ein Makler-Vorgang sein.
'       Cell-Highlight in Orange wenn 'ja'.
'
'       Unterschrift_Kunde + Unterschrift_Makler werden jetzt rein
'       FAKTISCH geprueft (ist eine sichtbar?). Pool-Ausnahmen (Check 24
'       digital, Verifox / Fonds Finanz / Impuls / Watson) wurden aus
'       der Unterschrift-Pruefung entfernt - die gehoeren semantisch in
'       die mv_einschraenkungen, nicht in die Unterschrift-Spalte.
' v2.8: GROSSE Ueberarbeitung der Maklervollmacht-Pruefung. Bisherige Heuristik
'       (6 Form-Pflichtfelder: Name, Geburtsdatum, Anschrift, Datum, Unter-
'       schrift, Maklerangaben) entspricht nicht der echten Pruefpraxis im
'       Innendienst. Neue Logik basiert auf der juristischen Vorgabe aus
'       AAW-Allg + PVC2D + Hospitationen T2/T3:
'         - INHALTLICHE Vollumfaenglichkeit: Schlagwort-Suche
'           ("abschliessen, aendern und kuendigen", "aktiv und passiv
'           vertreten", "bevollmaechtigt zu vertreten" etc.)
'         - EINSCHRAENKUNGEN als Ablehnungsgrund: "KV ausgeschlossen",
'           "schriftliche Zustimmung VN" etc.
'         - 4 flankierende Pflichten getrennt prueft: VN-Unterschrift,
'           Makler-Unterschrift (optional), Ausstellung auf VN, Makler
'           namentlich.
'       Spalten Q+R semantisch umbenannt:
'         Q: Vollmacht_Vollstaendig -> MV_Vollumfaenglich
'         R: Vollmacht_Fehlt -> MV_Einschraenkungen
'       Neue Spalten W (Auf_VN_Ausgestellt) und X (Makler_Namentlich).
'       Schlagwort-Liste als Doku: output/Maklervollmacht_..._Schlagworte.md.
'
'       Default-Modell auf 'gpt-51-chat' (war 'gpt-51-reasoning'):
'       Studien (Vellum, Galileo) zeigen, dass Reasoning-Modelle bei
'       strukturierter Klassifikation NICHT besser sind als Chat-Modelle,
'       aber 5-30x teurer/langsamer + halluzinieren MEHR bei Faktenextrak-
'       tion (PersonQA: o3 33% vs o1 16%). Bei Maklervollmacht-Feldern
'       ist ein leeres Feld besser als ein erfundener Name.
'
'       Optionale Textdatei pro Vorgang: am Anfang der Analyse fragt das
'       Tool, ob fuer jeden Vorgang eine .txt mit allen Klassifikations-
'       Daten im Unterordner '_KI-Analyse' angelegt werden soll.
' v2.7.2: Default-Modell auf 'gpt-51-reasoning' aktualisiert. Modelle_Testen-
'         Lauf vom 11/2025 hat gezeigt: ErgoGPT akzeptiert exakt drei Namen,
'         alle ohne Punkt:
'           - gpt-51-reasoning  (empfohlen, denkt nach)
'           - gpt-51-chat       (schneller, oberflaechlicher)
'           - gpt-41            (alt, Fallback)
'         Alle anderen Varianten (gpt-5.1, gpt-5.1-chat, gpt-4o, o3, ...)
'         antwortet die API mit HTTP 500 GENERAL_ERROR.
' v2.7.1: BUGFIX - Private Const ERGO_BASE_URL stand zwischen den Funktionen,
'         was VBA mit "Variable nicht definiert" beim Kompilieren ablehnt.
'         Module-Level-Declarations duerfen nur am Modul-KOPF stehen, vor
'         der ersten Sub/Function. Konstanten nach oben verschoben.
' v2.7: ASK_ErgoGPT + alle HTTP/Upload/JSON-Helfer sind direkt in dieses
'       Modul integriert. Kein separater Test.txt-Import mehr noetig - die
'       alten Module 'Agent', 'AI_Excel_Functions' und 'GPT' koennen aus
'       der Mappe geloescht werden.
' v2.6: NEUE Routine 'Modelle_Testen' - probiert systematisch eine Liste
'       von Modellnamen-Varianten gegen ErgoGPT (gpt-51, gpt-5.1,
'       gpt-5.1-chat, gpt-5.1-reasoning, gpt-4o ...) und schreibt das
'       Ergebnis in Sheet 'ModellTest'. So kann man sehen, welcher Name
'       die aktuelle ErgoGPT-API anspricht. Original-Modell in GPT!A6
'       wird vor und nach dem Test gesichert/restauriert.
' v2.5: NEUES Triage-Feld 'vorgangstyp' (Spalte T). Bounce-NDR-Mails,
'       ergo-Outbound, System-Mails, Werbung etc. werden NICHT mehr als
'       'Standardvorgang / BUe einfacher Vertrag' missklassifiziert,
'       sondern bekommen den korrekten Typ und alle anderen Felder
'       bleiben leer. Zeilen ohne echten Makler-Vorgang werden grau
'       hinterlegt, vorgangstyp-Zelle in Orange.
' v2.4: Modell-Fehler ("nicht mehr unterstuetzt") werden erkannt und brechen
'       sofort ab, mit Hinweis auf Sheet GPT!A6. Kleine Pause zwischen den
'       Calls (1.2s) verhindert das Rate-Limit von gpt.ergo.com (Azure-Gateway
'       schickt sonst 403). Setup-Spalte B bekommt einen Hinweis, dass der
'       Modellname EXAKT mit dem ErgoGPT-API-Wert uebereinstimmen muss.
' v2.3: Auth-/403-Fehler werden erkannt und brechen die Analyse sofort ab,
'       mit Dialog zum Cookie-Erneuern. HTML-Antworten (Azure-403-Seite)
'       werden nicht mehr roh in die Zelle geschrieben, sondern auf eine
'       lesbare Kurzform reduziert. Nach 5 Fehlern in Folge bricht der
'       Lauf ab, statt ueber alle Vorgaenge weiter zu hageln.
' ============================================================================
' WORKFLOW:
'   1) Diese .xlsm in den Ordner legen, in dem die .msg-Dateien liegen
'      (z.B. Output von ergo-email-batch / Mails_Speichern als MSG).
'   2) Alt+F8 -> Vorgaenge_Setup -> Ausfuehren (einmalig)
'        - legt Sheets 'GPT', 'Anleitung', 'Analyse' an
'        - fragt Cookie ab (wird in Sheet GPT!A7 ODER File hinterlegt)
'        - Default-Modell: gpt-51-reasoning  (bestaetigt 11/2025)
'   3) Alt+F8 -> Vorgaenge_Analysieren -> Ausfuehren
'        - fragt: Welcher Ordner? (Default: dieser Workbook-Pfad)
'        - fragt: Wie viele Vorgaenge maximal? (leer = alle)
'        - laedt pro Vorgang: Email-Text + PDF-Anhaenge zu ErgoGPT hoch
'        - GPT klassifiziert + prueft Maklervollmacht-Vollstaendigkeit
'        - Ergebnis landet als Zeile im Sheet 'Analyse'
'
' VORAUSSETZUNGEN:
'   - Outlook installiert (fuer .msg-Dateien -> OpenSharedItem).
'   - Cookie fuer gpt.ergo.com - 4 unterstuetzte Quellen:
'       1) Sheet GPT!A7 als Text-String
'       2) Sheet GPT!A8 als Pfad zu einer beliebigen Cookie-Datei
'       3) Default-Datei F:\ExcelGPT-Cookie\Cookie.txt
'       4) Dialog beim Start: Datei waehlen ODER Text einfuegen
'
' OUTPUT-SPALTEN im Sheet 'Analyse':
'   A  Datei
'   B  Datum
'   C  Absender_Name
'   D  Absender_Email
'   E  Betreff
'   F  Anhang_Namen
'   G  Maklerpool
'   H  Makler_Nachname
'   I  Makler_Vorname
'   J  Klassifikation              (Standardvorgang / Nicht-Standard)
'   K  Geschaefts_Typ              (BUe einfacher Vertrag / BUe Kundenverbindung / Keine BUe)
'   L  Unterlagen_Angefragt        (ja / nein)
'   M  Sonderfall                  (Flottengeschaeft / Sondertarif / Kein Sonderfall)
'   N  Sparte                      (Komposit / Leben / KV / Mehrere / Unbekannt)
'   O  Anhang_Typen                (Maklervollmacht, Police, Antrag, Schadenmeldung, ...)
'   P  Maklervollmacht_Enthalten   (ja / nein)
'   Q  MV_Vollumfaenglich          (ja / nein / teilweise / nicht_pruefbar -
'                                   basiert auf Schlagwort-Suche, NICHT auf
'                                   Form-Pflichtfeldern)
'   R  MV_Einschraenkungen         (Klartext: gefundene Einschraenkung oder
'                                   fehlendes Schlagwort)
'   S  Hinweis                     (kurzer GPT-Hinweis zum Vorgang)
'   T  Vorgangstyp                 (Triage)
'   U  Unterschrift_Kunde          (ja / nein / nicht_pruefbar - Pflicht)
'   V  Unterschrift_Makler         (ja / nein / nicht_pruefbar - optional)
'   W  Auf_VN_Ausgestellt          (ja / nein / nicht_pruefbar - Pflicht)
'   X  Makler_Namentlich           (ja / nein / nicht_pruefbar - Pflicht)
'   Y  Ist_Reminder                (ja / nein / nicht_pruefbar -
'                                   Erinnerung/Mahnung/Wiedervorlage)
'
' Pruefgrundlage Maklervollmacht: AAW-Allg 2.1.1+2.1.2, AAW-GES Kap. 5,
' PVC2D-Vorgabe (Rechtsabteilung). Schlagwort-Liste:
' output/Maklervollmacht_Vollumfaenglichkeit_Schlagworte.md
' ============================================================================

Option Explicit

Public Const SHEET_ANALYSE As String = "Analyse"
Public Const SHEET_GPT As String = "GPT"
Public Const SHEET_ANLEITUNG As String = "Anleitung"

Private Const COOKIE_PATH As String = "F:\ExcelGPT-Cookie\Cookie.txt"
Private Const COOKIE_CELL As String = "A7"            ' im Sheet GPT: Cookie als String
Private Const COOKIE_PATH_CELL As String = "A8"       ' im Sheet GPT: alternativer Pfad zur Cookie-Datei
Private Const MAX_PDFS_PRO_VORGANG As Long = 3
Private Const MAX_PDF_GROESSE_MB As Long = 20

' ErgoGPT-API-Konfig (siehe inlined ASK_ErgoGPT-Block am Modulende)
Private Const ERGO_BASE_URL As String = "https://gpt.ergo.com/api"
Private Const MAX_PARALLEL_PDF_UPLOADS As Long = 3

' Spaltenpositionen Output
Private Const COL_DATEI       As Long = 1
Private Const COL_DATUM       As Long = 2
Private Const COL_ABS_NAME    As Long = 3
Private Const COL_ABS_MAIL    As Long = 4
Private Const COL_BETREFF     As Long = 5
Private Const COL_ANHANG_LST  As Long = 6
Private Const COL_MAKLERPOOL  As Long = 7
Private Const COL_NACHNAME    As Long = 8
Private Const COL_VORNAME     As Long = 9
Private Const COL_KLASSIFIK   As Long = 10
Private Const COL_GESCHTYP    As Long = 11
Private Const COL_UNTERLAGEN  As Long = 12
Private Const COL_SONDERFALL  As Long = 13
Private Const COL_SPARTE      As Long = 14
Private Const COL_ANH_TYPEN   As Long = 15
Private Const COL_VM_VORHAND  As Long = 16
Private Const COL_VM_VOLLST   As Long = 17
Private Const COL_VM_FEHLT    As Long = 18
Private Const COL_HINWEIS     As Long = 19
Private Const COL_VORGANGSTYP As Long = 20    ' T: Triage - "Makler-Vorgang" / "Bounce-NDR" / "Ergo-Outbound" / "System-Mail" / "Werbung-Spam" / "Unklar"
Private Const COL_UNTERSCHR_K As Long = 21    ' U: Unterschrift Kunde vorhanden  (ja / nein / nicht_pruefbar) - Pflicht
Private Const COL_UNTERSCHR_M As Long = 22    ' V: Unterschrift Makler vorhanden (ja / nein / nicht_pruefbar) - optional
Private Const COL_AUF_VN      As Long = 23    ' W: MV auf VN ausgestellt          (ja / nein / nicht_pruefbar)
Private Const COL_MAKL_NAM    As Long = 24    ' X: Makler/Pool namentlich genannt (ja / nein / nicht_pruefbar)
Private Const COL_REMINDER    As Long = 25    ' Y: Reminder / Mahnung / Wiedervorlage (ja / nein / nicht_pruefbar)

' === HAUPTEINSTIEG ==========================================================
Public Sub Vorgaenge_Analysieren()
    On Error GoTo Fehler

    ' 0) Cookie-Setup pruefen
    If Not PruefeCookieMitDialog() Then Exit Sub

    ' 1) Ordner-Auswahl
    Dim folderPath As String
    folderPath = WaehleOrdner()
    If Len(folderPath) = 0 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Ordner nicht gefunden: " & folderPath, vbCritical
        Exit Sub
    End If

    ' 2) MSG-Dateien einsammeln
    Dim msgFiles As Collection
    Set msgFiles = SammleMsgDateien(fso, folderPath)
    If msgFiles.count = 0 Then
        MsgBox "Keine .msg-Dateien im Ordner gefunden:" & vbCrLf & folderPath, vbExclamation
        Exit Sub
    End If

    ' 3) Limit abfragen
    Dim limit As Long: limit = FrageLimit(msgFiles.count)
    If limit < 0 Then Exit Sub
    If limit = 0 Or limit > msgFiles.count Then limit = msgFiles.count

    ' 4) Bestaetigung
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Es werden " & limit & " von " & msgFiles.count & " .msg-Vorgaengen analysiert." & vbCrLf & _
                 "Pro Vorgang ca. 10-30 Sekunden (mit Anhang-Upload)." & vbCrLf & vbCrLf & _
                 "Geschaetzte Laufzeit: ~" & Int(limit * 20 / 60) + 1 & " Minuten." & vbCrLf & vbCrLf & _
                 "Fortfahren?", vbYesNo + vbQuestion, "Vorgang-Analyse starten")
    If ans <> vbYes Then Exit Sub

    ' 4b) Pro-Vorgang-Textdatei?
    Dim txtAns As VbMsgBoxResult
    txtAns = MsgBox( _
        "Soll fuer JEDEN Vorgang zusaetzlich eine Textdatei mit allen KI-Analyse-" & vbCrLf & _
        "Ergebnissen erzeugt werden?" & vbCrLf & vbCrLf & _
        "Speicherort: Unterordner '_KI-Analyse' im Vorgangs-Ordner." & vbCrLf & _
        "Dateiname:   <msg-Name>.txt" & vbCrLf & vbCrLf & _
        "JA = Textdatei je Vorgang erzeugen" & vbCrLf & _
        "NEIN = nur Excel-Sheet befuellen", _
        vbYesNo + vbQuestion, "Textdatei pro Vorgang?")
    Dim txtOrdner As String: txtOrdner = ""
    If txtAns = vbYes Then
        txtOrdner = folderPath & "\_KI-Analyse"
        On Error Resume Next
        If Not fso.FolderExists(txtOrdner) Then fso.CreateFolder txtOrdner
        On Error GoTo Fehler
        If Not fso.FolderExists(txtOrdner) Then
            MsgBox "Konnte Unterordner nicht anlegen: " & txtOrdner & vbCrLf & _
                   "Es wird ohne Textdatei-Erzeugung weitergemacht.", vbExclamation
            txtOrdner = ""
        End If
    End If

    ' 5) Outlook holen
    Dim olApp As Object
    Set olApp = HoleOutlook()
    If olApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden.", vbCritical
        Exit Sub
    End If

    ' 6) Analyse-Sheet vorbereiten
    Dim ws As Worksheet
    Set ws = HoleOderErzeugeAnalyseSheet()
    HeaderSchreiben ws

    Dim startRow As Long
    startRow = ws.cells(ws.Rows.count, 1).End(xlUp).row + 1
    If startRow < 2 Then startRow = 2

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim verarbeitet As Long: verarbeitet = 0
    Dim fehler As Long: fehler = 0
    Dim fehlerInFolge As Long: fehlerInFolge = 0
    Dim abbruchGrund As String: abbruchGrund = ""
    Dim row As Long: row = startRow
    Dim i As Long
    Dim tempBase As String: tempBase = Environ$("TEMP") & "\ergo_va_" & Format(Now, "yyyymmddhhnnss")

    For i = 1 To limit
        Dim msgPath As String: msgPath = msgFiles(i)
        Application.StatusBar = "Analysiere " & i & "/" & limit & " - " & fso.GetFileName(msgPath) & _
                                "  (verarbeitet:" & verarbeitet & " fehler:" & fehler & ")"
        DoEvents

        On Error Resume Next
        Err.Clear
        Dim erfolg As Boolean
        erfolg = AnalysiereEineDatei(olApp, fso, msgPath, ws, row, tempBase, i, txtOrdner)
        Dim einzelErrNum As Long: einzelErrNum = Err.Number
        Dim einzelErrDesc As String: einzelErrDesc = Err.Description
        Err.Clear
        On Error GoTo Fehler

        If einzelErrNum <> 0 Or Not erfolg Then
            Dim hint As String
            hint = BereinigeFehlerHinweis(einzelErrNum, einzelErrDesc)
            ws.cells(row, COL_DATEI).Value = fso.GetFileName(msgPath)
            ws.cells(row, COL_HINWEIS).Value = hint
            ws.cells(row, COL_HINWEIS).Interior.Color = RGB(252, 165, 165)
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Zeile " & row & " " & _
                        fso.GetFileName(msgPath) & ": " & hint
            fehler = fehler + 1
            fehlerInFolge = fehlerInFolge + 1
            row = row + 1

            ' Modell-Fehler? -> Sofort abbrechen, Hinweis auf Sheet GPT!A6
            If IstModellFehler(einzelErrDesc) Then
                abbruchGrund = "MODELL"
                Exit For
            End If

            ' Auth-/403-Fehler? -> Sofort abbrechen, Cookie-Dialog anbieten
            If IstAuthFehler(einzelErrDesc) Then
                abbruchGrund = "AUTH"
                Exit For
            End If

            ' 5 Fehler in Folge -> Etwas ist grundlegend kaputt, weiter macht keinen Sinn
            If fehlerInFolge >= 5 Then
                abbruchGrund = "REPEAT"
                Exit For
            End If
        Else
            verarbeitet = verarbeitet + 1
            fehlerInFolge = 0
            row = row + 1
        End If

        ' Kleine Pause gegen Rate-Limit / Azure-Gateway-403
        On Error Resume Next: Application.Wait Now + TimeSerial(0, 0, 1): On Error GoTo Fehler

        ' Auto-Save alle 3 Vorgaenge
        If i Mod 3 = 0 Then
            Application.Calculation = xlCalculationAutomatic
            Application.Calculation = xlCalculationManual
            On Error Resume Next
            ThisWorkbook.Save
            On Error GoTo Fehler
        End If
    Next i

    ' Temp-Verzeichnis aufraeumen
    On Error Resume Next
    If Len(tempBase) > 0 And fso.FolderExists(tempBase) Then
        fso.DeleteFolder tempBase, True
    End If
    On Error GoTo Fehler

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error Resume Next: ThisWorkbook.Save: On Error GoTo Fehler

    SpaltenbreitenSetzen ws

    Select Case abbruchGrund
        Case "MODELL"
            MsgBox _
                "Analyse abgebrochen: MODELL-FEHLER." & vbCrLf & vbCrLf & _
                "Der ErgoGPT-Server lehnt das in Sheet GPT!A6 eingetragene Modell ab" & vbCrLf & _
                "(z.B. 'gpt-51' oder 'gpt-5.1' wird nicht mehr unterstuetzt)." & vbCrLf & vbCrLf & _
                "Aktion: Im ErgoGPT-Browser oben rechts den aktuellen Modell-Namen" & vbCrLf & _
                "ablesen und in Sheet GPT!A6 EXAKT so eintragen.  Verarbeitet: " & verarbeitet & _
                "  Fehler: " & fehler, _
                vbExclamation, "Modell wird nicht unterstuetzt"
        Case "AUTH"
            Dim ansAuth As VbMsgBoxResult
            ansAuth = MsgBox( _
                "Analyse abgebrochen: AUTH-FEHLER (403 Forbidden)." & vbCrLf & vbCrLf & _
                "Der Cookie fuer gpt.ergo.com ist sehr wahrscheinlich abgelaufen oder ungueltig." & vbCrLf & _
                "Verarbeitet: " & verarbeitet & "  Fehler: " & fehler & vbCrLf & vbCrLf & _
                "Jetzt Cookie neu setzen und gleich weitermachen?" & vbCrLf & _
                "(JA = Cookie-Dialog oeffnen, NEIN = nur Hinweis)", _
                vbYesNo + vbExclamation, "Cookie abgelaufen")
            If ansAuth = vbYes Then PruefeCookieMitDialog
        Case "REPEAT"
            MsgBox _
                "Analyse abgebrochen: 5 Fehler in Folge." & vbCrLf & vbCrLf & _
                "Verarbeitet: " & verarbeitet & "  Fehler: " & fehler & vbCrLf & vbCrLf & _
                "Bitte Vorgaenge_Diagnose laufen lassen (Alt+F8) und das Direktfenster" & vbCrLf & _
                "im VBA-Editor (Strg+G) pruefen.", _
                vbExclamation, "Vorgaenge_Analysieren - Abbruch"
        Case Else
            MsgBox "Analyse abgeschlossen." & vbCrLf & vbCrLf & _
                   "Verarbeitet: " & verarbeitet & vbCrLf & _
                   "Fehler:      " & fehler, _
                   vbInformation, "Vorgaenge_Analysieren"
    End Select
    Exit Sub

Fehler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

' === EINE DATEI ANALYSIEREN =================================================
Private Function AnalysiereEineDatei(olApp As Object, fso As Object, msgPath As String, _
                                     ws As Worksheet, row As Long, tempBase As String, _
                                     vorgangNr As Long, txtOrdner As String) As Boolean
    Dim it As Object
    Dim schritt As String: schritt = "Init"
    On Error GoTo FehlerLokal

    ' --- Schritt 1: Datei pruefen ---
    schritt = "Datei pruefen"
    If Len(Trim(msgPath)) = 0 Then Err.Raise 5, , "msgPath ist leer"
    If Not fso.FileExists(msgPath) Then Err.Raise 53, , "Datei nicht gefunden: " & msgPath

    ' --- Schritt 2: MSG via Outlook oeffnen ---
    schritt = "Outlook OpenSharedItem"
    Set it = olApp.Session.OpenSharedItem(msgPath)
    If it Is Nothing Then Err.Raise 91, , "OpenSharedItem hat Nothing zurueckgegeben fuer: " & msgPath

    ' --- Schritt 3: Properties extrahieren ---
    schritt = "Mail-Properties lesen"
    Dim datum As String: datum = ""
    Dim absName As String: absName = ""
    Dim absMail As String: absMail = ""
    Dim betreff As String: betreff = ""
    Dim body As String: body = ""

    On Error Resume Next
    datum = Format(it.ReceivedTime, "dd.mm.yyyy hh:nn")
    absName = CStr(it.SenderName)
    absMail = HoleSenderEmail(it)
    betreff = CStr(it.Subject)
    body = CStr(it.Body)
    On Error GoTo FehlerLokal

    ' --- Schritt 4: Anhaenge extrahieren (PDFs hochladen) ---
    schritt = "Anhaenge extrahieren"
    Dim alleAnhangNamen As String: alleAnhangNamen = ""
    Dim pdfPaths As Collection: Set pdfPaths = New Collection
    Dim pdfListe As String: pdfListe = ""

    Dim tempDir As String: tempDir = tempBase & "\v" & vorgangNr
    On Error Resume Next
    If Not fso.FolderExists(tempBase) Then fso.CreateFolder tempBase
    If Not fso.FolderExists(tempDir) Then fso.CreateFolder tempDir
    On Error GoTo FehlerLokal

    Dim a As Long
    Dim attCount As Long: attCount = 0
    On Error Resume Next: attCount = it.Attachments.count: On Error GoTo FehlerLokal

    For a = 1 To attCount
        On Error Resume Next
        Dim att As Object: Set att = it.Attachments.Item(a)
        Dim fname As String: fname = ""
        If Not att Is Nothing Then fname = CStr(att.FileName)
        If Len(fname) > 0 Then
            If alleAnhangNamen <> "" Then alleAnhangNamen = alleAnhangNamen & "; "
            alleAnhangNamen = alleAnhangNamen & fname

            If LCase(fso.GetExtensionName(fname)) = "pdf" And pdfPaths.count < MAX_PDFS_PRO_VORGANG Then
                Dim safe As String: safe = SaeubereDateiname(fname)
                Dim outPath As String: outPath = tempDir & "\" & safe
                att.SaveAsFile outPath
                If fso.FileExists(outPath) Then
                    Dim fSize As Double: fSize = fso.GetFile(outPath).Size / 1048576
                    If fSize <= MAX_PDF_GROESSE_MB Then
                        pdfPaths.Add outPath
                        If pdfListe <> "" Then pdfListe = pdfListe & "; "
                        pdfListe = pdfListe & fname
                    Else
                        fso.DeleteFile outPath, True
                    End If
                End If
            End If
        End If
        Err.Clear
        On Error GoTo FehlerLokal
    Next a

    ' --- Schritt 5: Item schliessen ---
    schritt = "Mail schliessen"
    On Error Resume Next: it.Close 1: On Error GoTo FehlerLokal ' olDiscard
    Set it = Nothing

    ' --- Schritt 6: Body kappen + Prompt bauen ---
    schritt = "Prompt bauen"
    If Len(body) > 6000 Then body = Left(body, 6000) & " [...gekuerzt]"

    Dim prompt As String
    prompt = BuildVorgangPrompt(datum, absName, absMail, betreff, body, _
                                 alleAnhangNamen, pdfListe)

    ' --- Schritt 7: GPT befragen ---
    schritt = "ASK_ErgoGPT"
    Dim antwort As String
    If pdfPaths.count > 0 Then
        Dim pathArr() As String
        ReDim pathArr(0 To pdfPaths.count - 1)
        Dim k As Long
        For k = 1 To pdfPaths.count
            pathArr(k - 1) = pdfPaths(k)
        Next k
        antwort = ASK_ErgoGPT(prompt, pathArr)
    Else
        antwort = ASK_ErgoGPT(prompt)
    End If
    If Len(Trim(antwort)) = 0 Then Err.Raise 5, , "ASK_ErgoGPT hat leere Antwort geliefert"

    ' --- Schritt 8: JSON parsen + Schreiben ---
    schritt = "JSON parsen"
    Dim dict As Object: Set dict = ParseGptJsonAntwort(antwort)
    If dict Is Nothing Then Err.Raise 5, , "JSON-Antwort konnte nicht geparst werden: " & Left(antwort, 200)

    schritt = "Sheet schreiben"
    ws.cells(row, COL_DATEI).Value = fso.GetFileName(msgPath)
    ws.cells(row, COL_DATUM).Value = datum
    ws.cells(row, COL_ABS_NAME).Value = absName
    ws.cells(row, COL_ABS_MAIL).Value = absMail
    ws.cells(row, COL_BETREFF).Value = betreff
    ws.cells(row, COL_ANHANG_LST).Value = alleAnhangNamen
    SchreibeGptErgebnis ws, row, dict

    ' --- Schritt 9: optionale Textdatei pro Vorgang ---
    If Len(txtOrdner) > 0 Then
        On Error Resume Next
        SchreibeKiTextDatei txtOrdner, fso.GetFileName(msgPath), datum, absName, _
                            absMail, betreff, alleAnhangNamen, pdfListe, dict
        On Error GoTo FehlerLokal
    End If

    AnalysiereEineDatei = True
    Exit Function

FehlerLokal:
    ' Fehler-Info VOR dem Aufraeumen sichern
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Dim errSrc As String: errSrc = Err.Source
    On Error Resume Next
    If Not it Is Nothing Then it.Close 1
    Set it = Nothing
    On Error GoTo 0
    ' An Caller weitergeben (Caller hat On Error Resume Next aktiv -> kein Crash)
    Err.Raise errN, "AnalysiereEineDatei[" & schritt & "]", errD
    AnalysiereEineDatei = False
End Function

' === DIAGNOSE (testet 1 Datei und gibt jeden Schritt aus) ===================
Public Sub Vorgaenge_Diagnose()
    On Error GoTo Fehler
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim out As String: out = ""

    out = out & "=== Vorgang-Analyse Diagnose ===" & vbCrLf & vbCrLf
    Debug.Print out

    ' 1) Outlook
    Dim olApp As Object: Set olApp = HoleOutlook()
    If olApp Is Nothing Then
        out = out & "[X] Outlook NICHT erreichbar" & vbCrLf
        MsgBox out, vbCritical, "Diagnose": Exit Sub
    End If
    out = out & "[OK] Outlook erreichbar" & vbCrLf

    ' 2) ASK_ErgoGPT existiert?
    Dim hasASK As Boolean
    On Error Resume Next
    hasASK = (Not IsEmpty(Application.Run("ASK_ErgoGPT", "ping")))
    On Error GoTo Fehler
    If Not hasASK Then
        out = out & "[?] ASK_ErgoGPT-Test ergebnislos (siehe Direktfenster)" & vbCrLf
    Else
        out = out & "[OK] ASK_ErgoGPT antwortet" & vbCrLf
    End If

    ' 3) Ordner und Files
    Dim folderPath As String: folderPath = ThisWorkbook.Path
    out = out & "[?] Ordner: " & folderPath & vbCrLf
    Dim msgFiles As Collection: Set msgFiles = SammleMsgDateien(fso, folderPath)
    out = out & "[?] MSG-Dateien gefunden: " & msgFiles.count & vbCrLf

    If msgFiles.count = 0 Then
        MsgBox out & vbCrLf & "Lege MSG-Dateien in den Ordner und versuche erneut.", vbInformation, "Diagnose"
        Exit Sub
    End If

    ' 4) Erste MSG-Datei testen
    Dim path1 As String: path1 = msgFiles(1)
    out = out & vbCrLf & "Teste erste Datei: " & path1 & vbCrLf

    Dim it As Object
    On Error Resume Next: Set it = olApp.Session.OpenSharedItem(path1): On Error GoTo Fehler
    If it Is Nothing Then
        out = out & "[X] OpenSharedItem fehlgeschlagen: " & Err.Description & vbCrLf
        MsgBox out, vbCritical, "Diagnose": Exit Sub
    End If
    out = out & "[OK] OpenSharedItem ok" & vbCrLf

    Dim subj As String: subj = ""
    Dim attC As Long: attC = 0
    On Error Resume Next
    subj = CStr(it.Subject)
    attC = it.Attachments.count
    On Error GoTo Fehler
    out = out & "    Betreff: " & Left(subj, 60) & vbCrLf
    out = out & "    Anhaenge: " & attC & vbCrLf

    On Error Resume Next: it.Close 1: On Error GoTo Fehler

    out = out & vbCrLf & "Diagnose ok - falls Vorgaenge_Analysieren immer noch fehlschlaegt," & vbCrLf
    out = out & "Direktfenster (Strg+G) im VBA-Editor pruefen, dort wird jeder Fehler protokolliert."
    Debug.Print out
    MsgBox out, vbInformation, "Diagnose"
    Exit Sub
Fehler:
    out = out & vbCrLf & "[X] Diagnose-Fehler: " & Err.Number & " - " & Err.Description
    Debug.Print out
    MsgBox out, vbCritical, "Diagnose"
End Sub

' === MODELL-TEST ============================================================
' Probiert systematisch verschiedene Modellnamen-Varianten gegen ErgoGPT,
' loggt fuer jede Variante in Sheet 'ModellTest', ob sie funktioniert und
' welche Antwort/Fehler zurueckkommt. Ziel: rausfinden, welcher Name das
' GPT-5.1-Chat- bzw. das GPT-5.1-Reasoning-Modell anspricht.
'
' Der Original-Wert in Sheet GPT!A6 wird vor dem Test gespeichert und
' am Ende wiederhergestellt - egal ob der Test erfolgreich oder mit
' Fehler endet.
Public Sub Modelle_Testen()
    Const PING_PROMPT As String = "Antworte mit einem einzigen Wort: OK"

    ' Kandidaten-Liste: alles plausibel, geordnet von 'wahrscheinlich aktuell'
    ' nach 'eher alt'. Wenn die ErgoGPT-API neue Namen einfuehrt, hier ergaenzen.
    Dim kandidaten As Variant
    kandidaten = Array( _
        "gpt-41", "gpt-4.1", "gpt-4-1", _
        "gpt-51", "gpt-5.1", "gpt-5-1", _
        "gpt-51-chat", "gpt-5.1-chat", "gpt-5-1-chat", _
        "gpt-5-chat", "gpt-5chat", _
        "gpt-51-reasoning", "gpt-5.1-reasoning", "gpt-5-1-reasoning", _
        "gpt-5-reasoning", "gpt-51-thinking", "gpt-5.1-thinking", _
        "gpt-5", "gpt-5-mini", "gpt-5-nano", _
        "gpt-4o", "gpt-4o-mini", "gpt-4-turbo", _
        "o1", "o1-mini", "o3", "o3-mini", _
        "claude-3-5-sonnet", "claude-3.5-sonnet" _
    )

    Dim total As Long: total = UBound(kandidaten) - LBound(kandidaten) + 1

    Dim ans As VbMsgBoxResult
    ans = MsgBox( _
        "Modell-Test: pingt " & total & " Modellnamen mit einem Mini-Prompt." & vbCrLf & _
        "Pro Modell ca. 5-15s + 2s Pause -> insgesamt grob " & _
        Int(total * 12 / 60) + 1 & " Minuten." & vbCrLf & vbCrLf & _
        "Cookie muss aktuell sein. Sheet 'ModellTest' wird angelegt/geleert." & _
        vbCrLf & vbCrLf & "Fortfahren?", _
        vbYesNo + vbQuestion, "Modelle_Testen")
    If ans <> vbYes Then Exit Sub

    If Not PruefeCookieMitDialog() Then Exit Sub

    Dim wsGpt As Worksheet
    On Error Resume Next: Set wsGpt = ThisWorkbook.Sheets(SHEET_GPT): On Error GoTo 0
    If wsGpt Is Nothing Then
        MsgBox "Sheet '" & SHEET_GPT & "' fehlt. Bitte vorher Vorgaenge_Setup laufen lassen.", vbCritical
        Exit Sub
    End If

    Dim originalModell As String: originalModell = CStr(wsGpt.Range("A6").Value)

    ' Test-Sheet
    Dim wsTest As Worksheet
    On Error Resume Next: Set wsTest = ThisWorkbook.Sheets("ModellTest"): On Error GoTo 0
    If wsTest Is Nothing Then
        Set wsTest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsTest.name = "ModellTest"
    Else
        wsTest.cells.Clear
    End If

    wsTest.cells(1, 1).Value = "Modellname"
    wsTest.cells(1, 2).Value = "Status"
    wsTest.cells(1, 3).Value = "Dauer (s)"
    wsTest.cells(1, 4).Value = "Antwort / Fehler"
    With wsTest.Range(wsTest.cells(1, 1), wsTest.cells(1, 4))
        .Font.Bold = True
        .Interior.Color = RGB(30, 64, 175)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    wsTest.Columns(1).ColumnWidth = 28
    wsTest.Columns(2).ColumnWidth = 14
    wsTest.Columns(3).ColumnWidth = 12
    wsTest.Columns(4).ColumnWidth = 100
    wsTest.Rows(1).RowHeight = 22
    On Error Resume Next: ActiveWindow.FreezePanes = False: ActiveWindow.FreezePanes = True: On Error GoTo 0

    Application.ScreenUpdating = False

    Dim r As Long: r = 2
    Dim erfolgreich As Long: erfolgreich = 0
    Dim authFehler As Long: authFehler = 0
    Dim i As Long
    For i = LBound(kandidaten) To UBound(kandidaten)
        Dim modellName As String: modellName = CStr(kandidaten(i))
        Application.StatusBar = "Modell-Test " & (i - LBound(kandidaten) + 1) & "/" & total & ": " & modellName
        DoEvents

        wsGpt.Range("A6").Value = modellName

        Dim t0 As Double: t0 = Timer
        Dim antwort As String: antwort = ""
        Dim eN As Long: eN = 0
        Dim errDesc As String: errDesc = ""

        On Error Resume Next
        Err.Clear
        antwort = ASK_ErgoGPT(PING_PROMPT)
        eN = Err.Number
        errDesc = Err.Description
        Err.Clear
        On Error GoTo 0

        Dim dauer As Double: dauer = Timer - t0

        wsTest.cells(r, 1).Value = modellName
        wsTest.cells(r, 3).Value = Round(dauer, 1)

        If eN = 0 And Len(Trim(antwort)) > 0 Then
            wsTest.cells(r, 2).Value = "OK"
            wsTest.cells(r, 2).Interior.Color = RGB(187, 247, 208)
            wsTest.cells(r, 4).Value = Left(antwort, 500)
            erfolgreich = erfolgreich + 1
        Else
            wsTest.cells(r, 2).Value = "FEHLER"
            wsTest.cells(r, 2).Interior.Color = RGB(252, 165, 165)
            wsTest.cells(r, 4).Value = Left(BereinigeFehlerHinweis(eN, errDesc), 500)
            If IstAuthFehler(errDesc) Then
                authFehler = authFehler + 1
                ' Bei 3 Auth-Fehlern in Folge: Cookie scheint tot, Test abbrechen
                If authFehler >= 3 Then
                    r = r + 1
                    wsTest.cells(r, 1).Value = "(Test abgebrochen)"
                    wsTest.cells(r, 4).Value = "3 AUTH-Fehler in Folge - Cookie pruefen, dann Modelle_Testen erneut starten."
                    wsTest.cells(r, 4).Interior.Color = RGB(254, 215, 170)
                    Exit For
                End If
            Else
                authFehler = 0
            End If
        End If

        Debug.Print "[Modell-Test] " & modellName & " -> " & wsTest.cells(r, 2).Value & " (" & Round(dauer, 1) & "s)"
        r = r + 1

        ' Pause gegen Rate-Limit
        On Error Resume Next: Application.Wait Now + TimeSerial(0, 0, 2): On Error GoTo 0
    Next i

    ' Original wiederherstellen
    wsGpt.Range("A6").Value = originalModell

    Application.ScreenUpdating = True
    Application.StatusBar = False

    On Error Resume Next: wsTest.Activate: wsTest.Range("A2").Select: On Error GoTo 0

    MsgBox "Modell-Test fertig." & vbCrLf & vbCrLf & _
           "Erfolgreich: " & erfolgreich & " / " & total & vbCrLf & _
           "Original-Modell '" & originalModell & "' wiederhergestellt." & vbCrLf & vbCrLf & _
           "Ergebnisse: Sheet 'ModellTest'.  Den fuer Reasoning/Chat passenden" & vbCrLf & _
           "Namen einfach in Sheet GPT!A6 uebernehmen.", _
           vbInformation, "Modelle_Testen"
End Sub

' === FEHLER-AUFBEREITUNG ====================================================
' Erkennt Modell-Fehler in der Server-Antwort (z.B. "gpt-51 wird nicht
' mehr unterstuetzt"). Setzt voraus, dass ASK_ErgoGPT in Test.txt v2 den
' Server-Hinweis in die Fehlerbeschreibung uebernimmt.
Private Function IstModellFehler(desc As String) As Boolean
    Dim s As String: s = LCase(desc)
    If InStr(s, "nicht mehr unterst") > 0 Then IstModellFehler = True: Exit Function
    If InStr(s, "modell") > 0 And InStr(s, "unterstuetzt") > 0 Then IstModellFehler = True: Exit Function
    If InStr(s, "model") > 0 And InStr(s, "not supported") > 0 Then IstModellFehler = True: Exit Function
    If InStr(s, "unknown model") > 0 Then IstModellFehler = True: Exit Function
    IstModellFehler = False
End Function

' Erkennt Auth/403-Fehler (Cookie abgelaufen oder Rate-Limit vom
' Azure-Application-Gateway) - dann lohnt es sich nicht, alle weiteren
' Vorgaenge zu probieren.
Private Function IstAuthFehler(desc As String) As Boolean
    Dim s As String: s = LCase(desc)
    If InStr(s, "create failed: 403") > 0 Then IstAuthFehler = True: Exit Function
    If InStr(s, "403 forbidden") > 0 Then IstAuthFehler = True: Exit Function
    If InStr(s, "azure-application-gateway") > 0 Then IstAuthFehler = True: Exit Function
    If InStr(s, "401") > 0 And InStr(s, "unauthorized") > 0 Then IstAuthFehler = True: Exit Function
    IstAuthFehler = False
End Function

' Macht aus rohen Err.Description-Strings (inkl. eingebettetem HTML der
' Azure-403-Seite) eine kurze, lesbare Zelle. Verhindert mehrzeilige
' <html>-Bloecke in der Spalte 'Hinweis'.
Private Function BereinigeFehlerHinweis(num As Long, desc As String) As String
    Dim d As String: d = desc
    If Len(Trim(d)) = 0 Then
        BereinigeFehlerHinweis = "[FEHLER " & num & "] (keine Beschreibung - Schritt unklar)"
        Exit Function
    End If

    ' Modell-Fehler -> klare Klartext-Meldung (Sheet GPT!A6 anpassen)
    If IstModellFehler(d) Then
        BereinigeFehlerHinweis = "[MODELL-FEHLER] Modell in Sheet GPT!A6 wird vom Server " & _
                                 "abgelehnt. Aktuellen Modellnamen aus dem ErgoGPT-Browser " & _
                                 "(oben rechts) eintragen."
        Exit Function
    End If

    ' Auth-Fehler -> klare Klartext-Meldung
    If IstAuthFehler(d) Then
        BereinigeFehlerHinweis = "[AUTH-FEHLER 403] Cookie abgelaufen oder Rate-Limit. " & _
                                 "Bitte Cookie erneuern (Vorgaenge_Setup oder Sheet GPT!A7)."
        Exit Function
    End If

    ' HTML/Tags entfernen + Whitespace normalisieren
    If InStr(d, "<") > 0 And InStr(d, ">") > 0 Then
        Dim re As Object: Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.Pattern = "<[^>]+>"
        d = re.Replace(d, " ")
    End If
    d = Replace(d, vbCrLf, " ")
    d = Replace(d, vbCr, " ")
    d = Replace(d, vbLf, " ")
    d = Replace(d, vbTab, " ")
    Do While InStr(d, "  ") > 0
        d = Replace(d, "  ", " ")
    Loop
    d = Trim(d)
    If Len(d) > 280 Then d = Left(d, 277) & "..."

    BereinigeFehlerHinweis = "[FEHLER " & num & "] " & d
End Function

' === DER PROMPT (Kern) ======================================================
Private Function BuildVorgangPrompt(datum As String, absName As String, absMail As String, _
                                     betreff As String, text As String, _
                                     anhangAlle As String, anhangPdfsHochgeladen As String) As String
    Dim p As String
    p = ""
    p = p & "Du bist ein erfahrener Sachbearbeiter im Innendienst eines deutschen" & vbCrLf
    p = p & "Komposit-Versicherers (SHUK: Sach/Haftpflicht/Unfall/Kfz)." & vbCrLf
    p = p & "Du analysierst eingehende Makler-Emails inklusive Anhaengen." & vbCrLf
    p = p & vbCrLf
    p = p & "===== EMAIL =====" & vbCrLf
    p = p & "Datum: " & datum & vbCrLf
    p = p & "Absender: " & absName & " <" & absMail & ">" & vbCrLf
    p = p & "Betreff: " & betreff & vbCrLf
    p = p & "Anhaenge (alle): " & IIf(Len(anhangAlle) = 0, "(keine)", anhangAlle) & vbCrLf
    p = p & "Anhaenge HOCHGELADEN als PDF: " & IIf(Len(anhangPdfsHochgeladen) = 0, "(keine)", anhangPdfsHochgeladen) & vbCrLf
    p = p & vbCrLf
    p = p & "Body:" & vbCrLf
    p = p & "<<<" & vbCrLf
    p = p & text & vbCrLf
    p = p & ">>>" & vbCrLf
    p = p & "===== ENDE EMAIL =====" & vbCrLf
    p = p & vbCrLf
    p = p & "AUFGABE: Klassifiziere die Mail strikt anhand der unten definierten" & vbCrLf
    p = p & "Felder und Wertelisten." & vbCrLf
    p = p & vbCrLf
    p = p & "WICHTIG - SCHRITT 0 (Triage): Pruefe ZUERST das Feld 'vorgangstyp'." & vbCrLf
    p = p & "Wenn die Mail KEIN echter eingehender Makler-Vorgang ist (also Bounce-" & vbCrLf
    p = p & "NDR, von ergo selbst rausgegangen, automatische System-Mail, Werbung)," & vbCrLf
    p = p & "dann setze ALLE anderen 13 Felder auf '' (leerer String) - nur" & vbCrLf
    p = p & "'vorgangstyp' und 'hinweis' werden befuellt. NICHT raten, NICHT klassi-" & vbCrLf
    p = p & "fizieren. Hinweis erklaert in einem Satz, was die Mail wirklich ist." & vbCrLf
    p = p & vbCrLf
    p = p & "Nur wenn vorgangstyp == 'Makler-Vorgang': die uebrigen 13 Felder" & vbCrLf
    p = p & "befuellen. Nutze die HOCHGELADENEN PDFs (falls vorhanden) als" & vbCrLf
    p = p & "zusaetzliche Quelle - insbesondere fuer die Maklervollmacht-Pruefung." & vbCrLf
    p = p & vbCrLf
    p = p & "REGELN:" & vbCrLf
    p = p & "- Nutze Email-Signatur, From-Domain, Body und PDF-Inhalte als Quellen." & vbCrLf
    p = p & "- Wenn Information nicht eindeutig ablesbar ist: leerer String oder" & vbCrLf
    p = p & "  'nicht_pruefbar'. Rate NIE." & vbCrLf
    p = p & "- Antworte AUSSCHLIESSLICH mit einem einzeiligen JSON-Objekt mit GENAU" & vbCrLf
    p = p & "  14 Schluesseln (alle Werte als String). Kein Markdown, keine" & vbCrLf
    p = p & "  Code-Fences, kein Vorwort, keine Erklaerung." & vbCrLf
    p = p & vbCrLf
    p = p & "FELDER:" & vbCrLf
    p = p & vbCrLf
    p = p & "0) vorgangstyp  (TRIAGE - immer ausfuellen)" & vbCrLf
    p = p & "   - 'Makler-Vorgang': echte eingehende Mail von einem Makler/Maklerpool" & vbCrLf
    p = p & "     an ergo (BUe, Antrag, Schaden, Aenderung, Ruecksprache)." & vbCrLf
    p = p & "   - 'Bounce-NDR': Nichtzustellbarkeits-/Delivery-Failure-Benachrichti-" & vbCrLf
    p = p & "     gung. Indizien: Absender enthaelt 'postmaster', 'mailer-daemon'," & vbCrLf
    p = p & "     'ITERGO-Security', 'noreply'; Betreff 'Nachricht nicht zustellbar'," & vbCrLf
    p = p & "     'Undelivered Mail', 'Mail Delivery Failure', 'Returned mail';" & vbCrLf
    p = p & "     Body enthaelt 'Relay access denied', '5.7.1', 'delivery failed'," & vbCrLf
    p = p & "     'konnte nicht zugestellt werden'." & vbCrLf
    p = p & "   - 'Ergo-Outbound': Mail wurde von ergo selbst verschickt (Absender-" & vbCrLf
    p = p & "     Domain @ergo.de / @itergo.com / @ergo.com), z.B. eine Antwort des" & vbCrLf
    p = p & "     Bestandsuebertragungs-Teams. KEIN eingehender Vorgang." & vbCrLf
    p = p & "   - 'System-Mail': automatische Lese-/Empfangsbestaetigung, Out-of-" & vbCrLf
    p = p & "     Office, Kalender-Einladung, Newsletter, Calendar/iCal-Termin." & vbCrLf
    p = p & "   - 'Werbung-Spam': klare Werbung, Phishing, Externer-Anbieter-Pitch." & vbCrLf
    p = p & "   - 'Unklar': passt in keine Kategorie - dann ALLE Felder leer." & vbCrLf
    p = p & "   ACHTUNG: Wenn die Mail eine Bounce-Benachrichtigung im Anhang die" & vbCrLf
    p = p & "   urspruengliche ergo-Outbound-Mail enthaelt, ist DAS hier dennoch" & vbCrLf
    p = p & "   ein Bounce-NDR und kein Vorgang - die Inhalte des Anhangs sind nur" & vbCrLf
    p = p & "   Beweis fuer die Zustellung, nicht der Vorgang selbst." & vbCrLf
    p = p & vbCrLf
    p = p & "1) maklerpool" & vbCrLf
    p = p & "   Erlaubt: 'Fonds Finanz', 'BCA', 'JDC', 'blau direkt', 'Fondsnet'," & vbCrLf
    p = p & "   'Maxpool', 'WIFO', 'Aruna', 'Apella', 'Netfonds', 'VEMA', 'Mr.Money'," & vbCrLf
    p = p & "   'Direktmakler', '' (= unbekannt)" & vbCrLf
    p = p & vbCrLf
    p = p & "2) makler_nachname / 3) makler_vorname" & vbCrLf
    p = p & "   Aus Signatur, Anschreiben oder Vollmacht. Default: ''" & vbCrLf
    p = p & vbCrLf
    p = p & "4) klassifikation" & vbCrLf
    p = p & "   - 'Standardvorgang': Routine-Tagesgeschaeft (BUe, Antrag, Aenderung," & vbCrLf
    p = p & "     Schadenmeldung, einfache Ruecksprache)." & vbCrLf
    p = p & "   - 'Nicht-Standard': Eskalation, Beschwerde, Sonderkonstrukt, juristisch." & vbCrLf
    p = p & vbCrLf
    p = p & "5) geschaefts_typ" & vbCrLf
    p = p & "   BUe = Bestandsuebertragung. Vertraege wandern ohne inhaltliche" & vbCrLf
    p = p & "   Aenderung von Makler B zu Makler A. Eine VNR ist meistens drin," & vbCrLf
    p = p & "   ABER entscheidend ist: nur DER eine Vertrag oder ALLE Vertraege" & vbCrLf
    p = p & "   des Kunden?" & vbCrLf
    p = p & "   - 'BUe einfacher Vertrag': nur die genannten einzelnen Vertraege." & vbCrLf
    p = p & "   - 'BUe Kundenverbindung': gesamte Kundenverbindung / alle Vertraege" & vbCrLf
    p = p & "     dieses Kunden (Hinweise: 'gesamte Kundenverbindung', 'alle" & vbCrLf
    p = p & "     Vertraege', 'Kunde komplett')." & vbCrLf
    p = p & "   - 'Keine BUe': Vorgang ist was anderes (Schaden, Antrag, Aenderung)." & vbCrLf
    p = p & "   Erkennungs-Hinweise BUe: 'Bestandsuebertragung', 'Maklerwechsel'," & vbCrLf
    p = p & "   'BUe', 'Courtagezession', beigefuegte Maklervollmacht." & vbCrLf
    p = p & vbCrLf
    p = p & "6) unterlagen_angefragt" & vbCrLf
    p = p & "   - 'ja': Versicherer muss noch Unterlagen anfordern ODER Makler bittet" & vbCrLf
    p = p & "     um Nachreichung ODER Vorgang ohne weitere Unterlagen unvollstaendig." & vbCrLf
    p = p & "   - 'nein': mit den vorliegenden Anhaengen abschliessbar." & vbCrLf
    p = p & vbCrLf
    p = p & "7) sonderfall" & vbCrLf
    p = p & "   - 'Flottengeschaeft': Kfz-Flotten / mehrere Fahrzeuge Gewerbe." & vbCrLf
    p = p & "   - 'Sondertarif': ungewoehnliche Tarifkonstruktion, Konsortium," & vbCrLf
    p = p & "     Grossrisiko, Industrie-Police." & vbCrLf
    p = p & "   - 'Kein Sonderfall': normaler Privat-/Kleingewerbe-Vorgang." & vbCrLf
    p = p & vbCrLf
    p = p & "8) sparte" & vbCrLf
    p = p & "   - 'Komposit' (SHUK), 'Leben' (LV/RV/BU), 'KV' (PKV/Zusatz)," & vbCrLf
    p = p & "     'Mehrere', 'Unbekannt'" & vbCrLf
    p = p & vbCrLf
    p = p & "9) anhang_typen" & vbCrLf
    p = p & "   Kommagetrennte Liste der erkannten Anhang-Typen (basierend auf" & vbCrLf
    p = p & "   Dateinamen UND PDF-Inhalt). Erlaubte Werte (Mehrfachnennung ok):" & vbCrLf
    p = p & "   'Maklervollmacht', 'Police', 'Antrag', 'Schadenmeldung'," & vbCrLf
    p = p & "   'Beitragsrechnung', 'Anschreiben', 'Kuendigung', 'Datenblatt'," & vbCrLf
    p = p & "   'Sonstiges'. Default '' wenn keine Anhaenge oder nicht erkennbar." & vbCrLf
    p = p & vbCrLf
    p = p & "10) enthaelt_maklervollmacht" & vbCrLf
    p = p & "    'ja' wenn mindestens ein Anhang eine Maklervollmacht ist," & vbCrLf
    p = p & "    sonst 'nein'." & vbCrLf
    p = p & vbCrLf
    p = p & "11) mv_vollumfaenglich  (juristisch maßgeblich für Bestandsuebertragung)" & vbCrLf
    p = p & "    Pruefe inhaltlich auf SCHLAGWOERTER. Eine der folgenden Formulie-" & vbCrLf
    p = p & "    rungen muss in der MV erkennbar sein, damit sie vollumfaenglich ist:" & vbCrLf
    p = p & "      (a) 'abschliessen, aendern und kuendigen' (in Bezug auf VV)" & vbCrLf
    p = p & "      (b) 'Willenserklaerungen abgeben und entgegen[nehmen]' /" & vbCrLf
    p = p & "          'aktiv und passiv vertreten'" & vbCrLf
    p = p & "      (c) 'bevollmaechtigt zu vertreten' (uneingeschraenkt)" & vbCrLf
    p = p & "      (d) 'Abgabe und Entgegennahme von Willenserklaerungen' kombiniert" & vbCrLf
    p = p & "          mit Aenderung, Kuendigung, Abschluss eines Folgevertrages" & vbCrLf
    p = p & "      (e) 'uneingeschraenkte aktive und passive Vertretung ...'" & vbCrLf
    p = p & "          (juristischer Minimalstandard PVC2D)" & vbCrLf
    p = p & "    Werte:" & vbCrLf
    p = p & "      - 'ja':              mindestens eine Schlagwort-Formulierung erkennbar" & vbCrLf
    p = p & "      - 'teilweise':       Schlagwort vorhanden ABER mit Einschraenkung" & vbCrLf
    p = p & "                           (z.B. 'KV ausgeschlossen', 'schriftliche Zustimmung VN')" & vbCrLf
    p = p & "      - 'nein':            Schlagwoerter fehlen oder MV ist explizit eingeschraenkt" & vbCrLf
    p = p & "      - 'nicht_pruefbar':  keine MV im Anhang ODER PDF nicht ausgewertet" & vbCrLf
    p = p & "    WICHTIG: 'Vollumfaenglich' meint NICHT die Form (Name/Geburtsdatum/...)," & vbCrLf
    p = p & "    sondern den INHALT der Bevollmaechtigung. Eine MV ohne Geburtsdatum" & vbCrLf
    p = p & "    kann trotzdem 'vollumfaenglich = ja' sein, wenn die Schlagwoerter passen." & vbCrLf
    p = p & vbCrLf
    p = p & "12) mv_einschraenkungen" & vbCrLf
    p = p & "    Klartext-Liste der gefundenen Einschraenkungen oder fehlenden" & vbCrLf
    p = p & "    Schlagwoerter. Beispiele: 'KV ausgeschlossen', 'schriftliche" & vbCrLf
    p = p & "    Zustimmung VN gefordert', 'kein Schlagwort gefunden', 'Vollmacht" & vbCrLf
    p = p & "    nur fuer Schadensbearbeitung'. Leer wenn vollumfaenglich oder" & vbCrLf
    p = p & "    nicht pruefbar." & vbCrLf
    p = p & vbCrLf
    p = p & "13) unterschrift_kunde  (rein faktisch: ist eine sichtbar?)" & vbCrLf
    p = p & "    Ist auf der MV eine eigenhaendige oder digitale Unterschrift" & vbCrLf
    p = p & "    des KUNDEN/VN klar erkennbar (Schriftzug, Unterschriftenbild," & vbCrLf
    p = p & "    'digital signed'-Block)?" & vbCrLf
    p = p & "    Werte: 'ja' / 'nein' / 'nicht_pruefbar'" & vbCrLf
    p = p & "    HINWEIS: Bei Pool-Templates Check 24 digital, Verifox, Fonds" & vbCrLf
    p = p & "    Finanz, Impuls, Watson kann die Unterschrift in einem separaten" & vbCrLf
    p = p & "    Datenschutz-Anhang stehen - in dem Fall hier 'nein' (faktisch" & vbCrLf
    p = p & "    nicht auf der MV) und in 'mv_einschraenkungen' ergaenzen, dass" & vbCrLf
    p = p & "    es ein digital-signiertes Pool-Template ist (kein Ablehnungs-" & vbCrLf
    p = p & "    grund laut AAW-Allg)." & vbCrLf
    p = p & vbCrLf
    p = p & "14) unterschrift_makler  (rein faktisch)" & vbCrLf
    p = p & "    Ist zusaetzlich eine Unterschrift / Stempel des MAKLERS auf der" & vbCrLf
    p = p & "    MV sichtbar?  Werte: 'ja' / 'nein' / 'nicht_pruefbar'" & vbCrLf
    p = p & "    Optional - kein Ablehnungsgrund laut AAW-Allg." & vbCrLf
    p = p & vbCrLf
    p = p & "15) auf_vn_ausgestellt  (Pflicht)" & vbCrLf
    p = p & "    Ist die MV auf den VERSICHERUNGSNEHMER ausgestellt (nicht auf eine" & vbCrLf
    p = p & "    versicherte Person, nicht auf einen Bevollmaechtigten Dritter)?" & vbCrLf
    p = p & "    Werte: 'ja' / 'nein' / 'nicht_pruefbar'" & vbCrLf
    p = p & vbCrLf
    p = p & "16) makler_namentlich_genannt  (Pflicht)" & vbCrLf
    p = p & "    Ist der Makler ODER der Maklerpool namentlich in der MV benannt?" & vbCrLf
    p = p & "    Werte: 'ja' / 'nein' / 'nicht_pruefbar'" & vbCrLf
    p = p & vbCrLf
    p = p & "17) ist_reminder" & vbCrLf
    p = p & "    Ist diese Mail eine Erinnerung / Mahnung / Wiedervorlage zu einem" & vbCrLf
    p = p & "    bereits laufenden Vorgang? Indikatoren:" & vbCrLf
    p = p & "      - Betreff: 'Erinnerung', 'Reminder', 'Wiedervorlage', 'Mahnung'," & vbCrLf
    p = p & "        '2. Mahnung', 'zweite Anfrage', 'Sachstandsanfrage', 'Nach-" & vbCrLf
    p = p & "        frage', 'Bearbeitungsstatus', 'WV', 'AW:' bzw. 'Re:' mit" & vbCrLf
    p = p & "        zeitlichem Bezug auf eine frueher Mail" & vbCrLf
    p = p & "      - Body: 'wir warten noch', 'haben wir noch nichts erhalten'," & vbCrLf
    p = p & "        'bitten um Rueckmeldung', 'weiterhin keine Antwort', 'wie" & vbCrLf
    p = p & "        ist der Stand', 'in obiger Angelegenheit', 'haben wir" & vbCrLf
    p = p & "        bereits am ... angeschrieben'" & vbCrLf
    p = p & "    Werte: 'ja' / 'nein' / 'nicht_pruefbar'" & vbCrLf
    p = p & "    Ein Reminder kann gleichzeitig ein Makler-Vorgang sein - das" & vbCrLf
    p = p & "    Feld ist unabhaengig von 'vorgangstyp' und 'klassifikation'." & vbCrLf
    p = p & vbCrLf
    p = p & "18) hinweis" & vbCrLf
    p = p & "    EIN kurzer Satz (max 200 Zeichen) was an dem Vorgang auffaellig" & vbCrLf
    p = p & "    ist - z.B. fehlende Unterlagen, ungewoehnliche Konstellation," & vbCrLf
    p = p & "    Eskalationspotenzial. Leer wenn nichts auffaellt." & vbCrLf
    p = p & vbCrLf
    p = p & "AUSGABE-FORMAT (eine Zeile, gueltiges JSON, alle 19 Schluessel):" & vbCrLf
    p = p & "{""vorgangstyp"":""..."",""maklerpool"":""..."",""makler_nachname"":""...""," & vbCrLf
    p = p & " ""makler_vorname"":""..."",""klassifikation"":""..."",""geschaefts_typ"":""...""," & vbCrLf
    p = p & " ""unterlagen_angefragt"":""..."",""sonderfall"":""..."",""sparte"":""...""," & vbCrLf
    p = p & " ""anhang_typen"":""..."",""enthaelt_maklervollmacht"":""...""," & vbCrLf
    p = p & " ""mv_vollumfaenglich"":""..."",""mv_einschraenkungen"":""...""," & vbCrLf
    p = p & " ""unterschrift_kunde"":""..."",""unterschrift_makler"":""...""," & vbCrLf
    p = p & " ""auf_vn_ausgestellt"":""..."",""makler_namentlich_genannt"":""...""," & vbCrLf
    p = p & " ""ist_reminder"":""...""," & vbCrLf
    p = p & " ""hinweis"":""...""}" & vbCrLf
    p = p & vbCrLf
    p = p & "Antworte JETZT, nur das JSON-Objekt:" & vbCrLf
    BuildVorgangPrompt = p
End Function

' === JSON-PARSING ===========================================================
Private Function ParseGptJsonAntwort(antwort As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim clean As String: clean = antwort
    clean = Replace(clean, "```json", "")
    clean = Replace(clean, "```", "")
    clean = Trim(clean)

    Dim p1 As Long: p1 = InStr(clean, "{")
    Dim p2 As Long: p2 = InStrRev(clean, "}")
    If p1 = 0 Or p2 <= p1 Then
        Set ParseGptJsonAntwort = Nothing
        Exit Function
    End If
    clean = Mid(clean, p1, p2 - p1 + 1)

    Dim keys As Variant
    keys = Array("vorgangstyp", _
                 "maklerpool", "makler_nachname", "makler_vorname", _
                 "klassifikation", "geschaefts_typ", "unterlagen_angefragt", _
                 "sonderfall", "sparte", "anhang_typen", _
                 "enthaelt_maklervollmacht", "mv_vollumfaenglich", _
                 "mv_einschraenkungen", _
                 "unterschrift_kunde", "unterschrift_makler", _
                 "auf_vn_ausgestellt", "makler_namentlich_genannt", _
                 "ist_reminder", _
                 "hinweis")
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        dict(keys(i)) = ExtrahiereJsonString(clean, CStr(keys(i)))
    Next i
    Set ParseGptJsonAntwort = dict
End Function

Private Function ExtrahiereJsonString(json As String, key As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.Pattern = """" & key & """\s*:\s*""((?:\\.|[^""\\])*)"""
    If re.test(json) Then
        Dim raw As String: raw = re.Execute(json)(0).SubMatches(0)
        raw = Replace(raw, "\""", """")
        raw = Replace(raw, "\\", "\")
        raw = Replace(raw, "\n", " ")
        raw = Replace(raw, "\r", " ")
        raw = Replace(raw, "\t", " ")
        ExtrahiereJsonString = Trim(raw)
    End If
End Function

' === COOKIE-DIALOG ==========================================================
' Vier Cookie-Quellen werden unterstuetzt (in dieser Reihenfolge):
'   1) Sheet GPT!A7 enthaelt einen Cookie-String (>50 Zeichen)
'   2) Sheet GPT!A8 enthaelt einen Pfad zu einer Cookie-Datei
'   3) Default-Pfad F:\ExcelGPT-Cookie\Cookie.txt
'   4) Manuell im Dialog: Datei waehlen ODER String einfuegen
' Egal woher der Cookie kommt: er wird IMMER nach COOKIE_PATH gespiegelt,
' weil ASK_ErgoGPT (Test.txt) hartkodiert von dort liest.
Private Function PruefeCookieMitDialog() As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    ' (1) Cookie-String aus Sheet GPT!A7?
    Dim cookieAusSheet As String: cookieAusSheet = HoleCookieAusSheet()
    If Len(cookieAusSheet) > 50 Then
        SchreibeCookieDatei cookieAusSheet, fso
        PruefeCookieMitDialog = True
        Exit Function
    End If

    ' (2) Custom-Pfad aus Sheet GPT!A8?
    Dim customPath As String: customPath = HoleCookiePfadAusSheet()
    If Len(customPath) > 0 And fso.FileExists(customPath) Then
        Dim contentCustom As String: contentCustom = LeseCookieDatei(customPath, fso)
        If Len(contentCustom) > 50 Then
            SchreibeCookieDatei contentCustom, fso
            PruefeCookieMitDialog = True
            Exit Function
        End If
    End If

    ' (3) Default-Pfad pruefen
    Dim hatDefault As Boolean: hatDefault = fso.FileExists(COOKIE_PATH)
    Dim defaultInhalt As String
    If hatDefault Then defaultInhalt = LeseCookieDatei(COOKIE_PATH, fso)

    Dim ans As VbMsgBoxResult, msg As String
    If hatDefault And Len(defaultInhalt) > 50 Then
        Dim preview As String: preview = Left(defaultInhalt, 100) & IIf(Len(defaultInhalt) > 100, "...", "")
        msg = "Cookie-Datei gefunden:" & vbCrLf & COOKIE_PATH & vbCrLf & vbCrLf & _
              "Auszug: " & preview & vbCrLf & vbCrLf & _
              "JA      = diesen Cookie verwenden" & vbCrLf & _
              "NEIN    = anderen Cookie laden (Datei waehlen oder neu eingeben)" & vbCrLf & _
              "ABBRECH = Beenden"
        ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "Cookie-Setup")
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then PruefeCookieMitDialog = True: Exit Function
    End If

    ' (4) Quellen-Auswahl
    Dim quelle As VbMsgBoxResult
    quelle = MsgBox( _
        "Wie soll der Cookie eingelesen werden?" & vbCrLf & vbCrLf & _
        "JA      = Cookie-DATEI auswaehlen (Browse-Dialog)" & vbCrLf & _
        "NEIN    = Cookie als TEXT einfuegen (Browser-Copy-Paste)" & vbCrLf & _
        "ABBRECH = Beenden", _
        vbYesNoCancel + vbQuestion, "Cookie-Quelle waehlen")
    If quelle = vbCancel Then Exit Function

    Dim cookie As String

    If quelle = vbYes Then
        ' --- Datei waehlen ---
        Dim pfad As String: pfad = WaehleCookieDatei()
        If Len(pfad) = 0 Then Exit Function
        If Not fso.FileExists(pfad) Then
            MsgBox "Datei nicht gefunden: " & pfad, vbCritical
            Exit Function
        End If
        cookie = LeseCookieDatei(pfad, fso)
        If Len(cookie) < 50 Then
            MsgBox "Datei enthaelt keinen gueltigen Cookie (zu kurz):" & vbCrLf & pfad, vbExclamation
            Exit Function
        End If
        ' Pfad fuer naechstes Mal merken
        SetzeCookiePfadInSheet pfad
    Else
        ' --- Manuell eingeben ---
        cookie = InputBox( _
            "Cookie-String einfuegen." & vbCrLf & vbCrLf & _
            "Browser -> gpt.ergo.com einloggen -> F12 (DevTools) ->" & vbCrLf & _
            "Tab Network -> beliebigen Request -> Request Headers ->" & vbCrLf & _
            "Wert hinter 'Cookie:' kopieren und hier einfuegen.", _
            "Cookie eingeben")
        cookie = Trim(cookie)
        If Len(cookie) < 50 Then
            MsgBox "Cookie zu kurz oder leer - Abbruch.", vbExclamation
            Exit Function
        End If
    End If

    SchreibeCookieDatei cookie, fso
    MsgBox "Cookie gespeichert. Analyse startet jetzt.", vbInformation
    PruefeCookieMitDialog = True
End Function

Private Function WaehleCookieDatei() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Cookie-Datei waehlen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Text-Dateien (*.txt)", "*.txt"
        .Filters.Add "Alle Dateien (*.*)", "*.*"
        ' Default-Folder: Workbook-Pfad
        On Error Resume Next: .InitialFileName = ThisWorkbook.Path & "\": On Error GoTo 0
        If .Show <> -1 Then WaehleCookieDatei = "": Exit Function
        WaehleCookieDatei = .SelectedItems(1)
    End With
End Function

Private Function LeseCookieDatei(pfad As String, fso As Object) As String
    On Error Resume Next
    Dim ts As Object
    Set ts = fso.OpenTextFile(pfad, 1, False, -2) ' ForReading, default Codepage
    If Not ts Is Nothing Then
        LeseCookieDatei = Trim(ts.ReadAll)
        ts.Close
    End If
    On Error GoTo 0
End Function

Private Function HoleCookiePfadAusSheet() As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_GPT)
    If ws Is Nothing Then HoleCookiePfadAusSheet = "": Exit Function
    HoleCookiePfadAusSheet = Trim(CStr(ws.Range(COOKIE_PATH_CELL).Value))
    On Error GoTo 0
End Function

Private Sub SetzeCookiePfadInSheet(pfad As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_GPT)
    If Not ws Is Nothing Then ws.Range(COOKIE_PATH_CELL).Value = pfad
    On Error GoTo 0
End Sub

Private Sub SchreibeCookieDatei(cookie As String, fso As Object)
    Dim dirPath As String: dirPath = Left(COOKIE_PATH, InStrRev(COOKIE_PATH, "\") - 1)
    On Error Resume Next
    If Len(dirPath) > 0 And Not fso.FolderExists(dirPath) Then fso.CreateFolder dirPath
    Dim out As Object
    Set out = fso.CreateTextFile(COOKIE_PATH, True, False)
    out.Write cookie
    out.Close
    On Error GoTo 0
End Sub

Private Function HoleCookieAusSheet() As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_GPT)
    If ws Is Nothing Then HoleCookieAusSheet = "": Exit Function
    HoleCookieAusSheet = Trim(CStr(ws.Range(COOKIE_CELL).Value))
    On Error GoTo 0
End Function

' === SETUP ===================================================================
Public Sub Vorgaenge_Setup()
    Dim ans As VbMsgBoxResult
    ans = MsgBox( _
        "Setup richtet die Excel-Mappe fuer die Vorgang-Analyse v2.0 ein:" & vbCrLf & vbCrLf & _
        "  - Sheet 'GPT' (Modell, Cookie-Zelle, Temperature, Tone)" & vbCrLf & _
        "  - Sheet 'Anleitung' (Step-by-Step in Excel)" & vbCrLf & _
        "  - Sheet 'Analyse' (Output-Spalten)" & vbCrLf & _
        "  - Cookie abfragen / pruefen" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Vorgaenge_Setup")
    If ans <> vbYes Then Exit Sub

    SetupGptSheet
    SetupAnleitungSheet
    Dim ws As Worksheet: Set ws = HoleOderErzeugeAnalyseSheet()
    HeaderSchreiben ws
    SpaltenbreitenSetzen ws

    PruefeCookieMitDialog

    On Error Resume Next: ThisWorkbook.Sheets(SHEET_ANLEITUNG).Activate
End Sub

Private Sub SetupGptSheet()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SHEET_GPT): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = SHEET_GPT
    End If

    ws.Range("B6").Value = "<- Modell. Empfehlung: 'gpt-51-chat' (Default, schnell+guenstig, weniger Halluzinationen bei Extraktion). Reasoning ('gpt-51-reasoning') nur bei schlechten Scans/komplexen Sonderfaellen - kostet 5-10x mehr und halluziniert mehr Felder. Studien dazu siehe Anleitung. 'gpt-41' = alt."
    ws.Range("B7").Value = "<- Cookie als Text (langer String) - ODER leer lassen und A8/Dialog nutzen"
    ws.Range("B8").Value = "<- Pfad zu Cookie-Datei (z.B. C:\Users\...\Desktop\cookie.txt) - leer = Default F:\ExcelGPT-Cookie\Cookie.txt"
    ws.Range("B9").Value = "<- Temperature (0 = deterministisch)"
    ws.Range("B12").Value = "<- Tone (z.B. 'Sachlich' oder leer)"

    If Trim(CStr(ws.Range("A6").Value)) = "" Then ws.Range("A6").Value = "gpt-51-chat"
    If Trim(CStr(ws.Range("A9").Value)) = "" Then ws.Range("A9").Value = 0
    If Trim(CStr(ws.Range("A12").Value)) = "" Then ws.Range("A12").Value = "Sachlich"

    ws.Columns("A").ColumnWidth = 32
    ws.Columns("B").ColumnWidth = 90
    ws.Range("A7").WrapText = False
    ws.Range("A8").WrapText = False
End Sub

Private Sub SetupAnleitungSheet()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SHEET_ANLEITUNG): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.name = SHEET_ANLEITUNG
    Else
        ws.cells.Clear
    End If

    Dim r As Long: r = 1
    ws.cells(r, 1).Value = "ERGO Vorgang-Analyse v2.0 - Anleitung": Bold ws, r, 16: r = r + 2

    ws.cells(r, 1).Value = "WORKFLOW (kurz)": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "1. Diese .xlsm in den Ordner legen, in dem .msg-Dateien liegen.": r = r + 1
    ws.cells(r, 1).Value = "2. Vorgaenge_Setup einmalig ausfuehren (Cookie hinterlegen).": r = r + 1
    ws.cells(r, 1).Value = "3. Vorgaenge_Analysieren ausfuehren -> Ordner + Anzahl waehlen.": r = r + 2

    ws.cells(r, 1).Value = "VORAUSSETZUNGEN": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "- ASK_ErgoGPT ist seit v2.7 in diesem Modul integriert (kein Test.txt noetig).": r = r + 1
    ws.cells(r, 1).Value = "- Outlook installiert (zum Oeffnen der .msg-Dateien).": r = r + 1
    ws.cells(r, 1).Value = "- Cookie fuer gpt.ergo.com - 4 Quellen werden unterstuetzt:": r = r + 1
    ws.cells(r, 1).Value = "    1) Sheet GPT!A7 als Text-String (lang)": r = r + 1
    ws.cells(r, 1).Value = "    2) Sheet GPT!A8 als Pfad zu einer Cookie-Datei": r = r + 1
    ws.cells(r, 1).Value = "    3) Default-Datei F:\ExcelGPT-Cookie\Cookie.txt": r = r + 1
    ws.cells(r, 1).Value = "    4) Dialog: Datei waehlen (Browse) ODER Text einfuegen": r = r + 2

    ws.cells(r, 1).Value = "AUSFUEHRUNG VORGAENGE_ANALYSIEREN": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "Schritt a) Cookie-Dialog (falls keine Quelle vorbelegt):": r = r + 1
    ws.cells(r, 1).Value = "   - vorhandenen Cookie nutzen ODER andere Datei waehlen ODER neu eingeben": r = r + 1
    ws.cells(r, 1).Value = "Schritt b) Ordner-Dialog (Default: dieser Workbook-Pfad)": r = r + 1
    ws.cells(r, 1).Value = "Schritt c) Anzahl-Eingabe (leer = alle .msg-Dateien)": r = r + 1
    ws.cells(r, 1).Value = "Schritt d) Bestaetigung mit geschaetzter Laufzeit": r = r + 1
    ws.cells(r, 1).Value = "Schritt e) Pro Vorgang: Email + max. 3 PDFs (<20 MB) -> ErgoGPT.": r = r + 1
    ws.cells(r, 1).Value = "           Auto-Save alle 3 Vorgaenge.": r = r + 2

    ws.cells(r, 1).Value = "OUTPUT-SPALTEN im Sheet 'Analyse'": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "A Datei | B Datum | C Absender_Name | D Absender_Email | E Betreff": r = r + 1
    ws.cells(r, 1).Value = "F Anhang_Namen | G Maklerpool | H Makler_Nachname | I Makler_Vorname": r = r + 1
    ws.cells(r, 1).Value = "J Klassifikation | K Geschaefts_Typ | L Unterlagen_Angefragt": r = r + 1
    ws.cells(r, 1).Value = "M Sonderfall | N Sparte | O Anhang_Typen": r = r + 1
    ws.cells(r, 1).Value = "P Maklervollmacht_Enthalten | Q Vollmacht_Vollstaendig": r = r + 1
    ws.cells(r, 1).Value = "R Vollmacht_Fehlt | S Hinweis": r = r + 2

    ws.cells(r, 1).Value = "BUE-LOGIK": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "BUe = Bestandsuebertragung (Maklerwechsel ohne Vertragsaenderung).": r = r + 1
    ws.cells(r, 1).Value = "Eine VNR ist meistens da. Frage ist: nur DER eine Vertrag oder": r = r + 1
    ws.cells(r, 1).Value = "die GANZE Kundenverbindung des Kunden?": r = r + 2

    ws.cells(r, 1).Value = "MAKLERVOLLMACHT-PRUEFUNG": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "GPT prueft im PDF: Kundenname, Geburtsdatum, Anschrift, Datum,": r = r + 1
    ws.cells(r, 1).Value = "Unterschrift, Maklerangaben (Firma + Vermittlernummer).": r = r + 1
    ws.cells(r, 1).Value = "Fehlende Felder werden in Spalte R aufgelistet.": r = r + 2

    ws.cells(r, 1).Value = "ZURUECKSETZEN": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "Alt+F8 -> 'Vorgaenge_Analysieren_Reset' loescht das Sheet 'Analyse'.": r = r + 2

    ws.cells(r, 1).Value = "DIAGNOSE BEI FEHLERN": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "Wenn die Analyse fehlschlaegt: Alt+F8 -> 'Vorgaenge_Diagnose'.": r = r + 1
    ws.cells(r, 1).Value = "Testet Outlook, MSG-Dateien und gibt fuer die erste Datei jeden Schritt aus.": r = r + 1
    ws.cells(r, 1).Value = "Detail-Log: VBA-Editor (Alt+F11) -> Direktfenster (Strg+G).": r = r + 2

    ws.cells(r, 1).Value = "MODELL-TEST (welcher Modellname klappt?)": Bold ws, r, 12: r = r + 1
    ws.cells(r, 1).Value = "Alt+F8 -> 'Modelle_Testen' probiert ca. 25 Modellnamen-Varianten durch": r = r + 1
    ws.cells(r, 1).Value = "(gpt-41, gpt-51, gpt-5.1-chat, gpt-5.1-reasoning, gpt-4o, o3 ...) und": r = r + 1
    ws.cells(r, 1).Value = "schreibt das Ergebnis pro Modell ins Sheet 'ModellTest'.": r = r + 1
    ws.cells(r, 1).Value = "Den passenden Namen einfach in GPT!A6 uebernehmen.": r = r + 2

    ws.Columns("A").ColumnWidth = 110
End Sub

' === RESET ==================================================================
Public Sub Vorgaenge_Analysieren_Reset()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SHEET_ANALYSE): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim ans As VbMsgBoxResult
    ans = MsgBox("Sheet '" & SHEET_ANALYSE & "' komplett leeren?", vbYesNo + vbExclamation, "Reset")
    If ans <> vbYes Then Exit Sub

    ws.cells.Clear
    HeaderSchreiben ws
    SpaltenbreitenSetzen ws
    MsgBox "Sheet '" & SHEET_ANALYSE & "' geleert.", vbInformation
End Sub

' === HELFER =================================================================
Private Function WaehleOrdner() As String
    Dim defaultPath As String
    defaultPath = ThisWorkbook.Path
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Welcher Ordner soll gescannt werden?" & vbCrLf & vbCrLf & _
                 "JA      = aktueller Workbook-Pfad:" & vbCrLf & "  " & defaultPath & vbCrLf & vbCrLf & _
                 "NEIN    = anderen Ordner waehlen" & vbCrLf & _
                 "ABBRECH = Beenden", _
                 vbYesNoCancel + vbQuestion, "Ordner-Auswahl")
    If ans = vbCancel Then WaehleOrdner = "": Exit Function
    If ans = vbYes Then WaehleOrdner = defaultPath: Exit Function

    ' Anderen Ordner waehlen
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner mit .msg-Dateien waehlen"
        .AllowMultiSelect = False
        If .Show <> -1 Then WaehleOrdner = "": Exit Function
        WaehleOrdner = .SelectedItems(1)
    End With
End Function

Private Function FrageLimit(maxAnzahl As Long) As Long
    Dim eing As String
    eing = InputBox("Wie viele Vorgaenge sollen analysiert werden?" & vbCrLf & vbCrLf & _
                    "  Im Ordner gefunden: " & maxAnzahl & " .msg-Dateien" & vbCrLf & _
                    "  Leer oder 0 = ALLE" & vbCrLf & _
                    "  Sonst: maximale Anzahl (z.B. 10 zum Testen)", _
                    "Anzahl Vorgaenge", "10")
    If eing = "" Then FrageLimit = -1: Exit Function
    On Error Resume Next
    FrageLimit = CLng(Val(eing))
    If FrageLimit < 0 Then FrageLimit = 0
End Function

Private Function SammleMsgDateien(fso As Object, folderPath As String) As Collection
    Dim col As Collection: Set col = New Collection
    Dim folder As Object: Set folder = fso.GetFolder(folderPath)
    Dim file As Object
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.name)) = "msg" Then
            col.Add file.path
        End If
    Next file
    Set SammleMsgDateien = col
End Function

Private Function HoleOutlook() As Object
    On Error Resume Next
    Dim app As Object
    Set app = GetObject(, "Outlook.Application")
    If app Is Nothing Then Set app = CreateObject("Outlook.Application")
    Set HoleOutlook = app
End Function

Private Function HoleSenderEmail(it As Object) As String
    On Error Resume Next
    Dim typ As String: typ = it.SenderEmailType
    If typ = "EX" Then
        Dim sender As Object: Set sender = it.sender
        If Not sender Is Nothing Then
            Dim eu As Object: Set eu = sender.GetExchangeUser()
            If Not eu Is Nothing Then
                HoleSenderEmail = eu.PrimarySmtpAddress
                If HoleSenderEmail <> "" Then Exit Function
            End If
        End If
    End If
    HoleSenderEmail = it.SenderEmailAddress
End Function

Private Function SaeubereDateiname(s As String) As String
    Dim r As String: r = s
    Dim verboten As String: verboten = "\/:*?""<>|"
    Dim i As Long
    For i = 1 To Len(verboten)
        r = Replace(r, Mid(verboten, i, 1), "_")
    Next i
    If Len(r) > 100 Then
        Dim ext As String
        Dim posDot As Long: posDot = InStrRev(r, ".")
        If posDot > 0 Then ext = Mid(r, posDot)
        r = Left(r, 100 - Len(ext)) & ext
    End If
    SaeubereDateiname = r
End Function

Private Function HoleOderErzeugeAnalyseSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SHEET_ANALYSE): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = SHEET_ANALYSE
    End If
    Set HoleOderErzeugeAnalyseSheet = ws
End Function

Private Sub HeaderSchreiben(ws As Worksheet)
    ws.cells(1, COL_DATEI).Value = "Datei"
    ws.cells(1, COL_DATUM).Value = "Datum"
    ws.cells(1, COL_ABS_NAME).Value = "Absender_Name"
    ws.cells(1, COL_ABS_MAIL).Value = "Absender_Email"
    ws.cells(1, COL_BETREFF).Value = "Betreff"
    ws.cells(1, COL_ANHANG_LST).Value = "Anhang_Namen"
    ws.cells(1, COL_MAKLERPOOL).Value = "Maklerpool"
    ws.cells(1, COL_NACHNAME).Value = "Makler_Nachname"
    ws.cells(1, COL_VORNAME).Value = "Makler_Vorname"
    ws.cells(1, COL_KLASSIFIK).Value = "Klassifikation"
    ws.cells(1, COL_GESCHTYP).Value = "Geschaefts_Typ"
    ws.cells(1, COL_UNTERLAGEN).Value = "Unterlagen_Angefragt"
    ws.cells(1, COL_SONDERFALL).Value = "Sonderfall"
    ws.cells(1, COL_SPARTE).Value = "Sparte"
    ws.cells(1, COL_ANH_TYPEN).Value = "Anhang_Typen"
    ws.cells(1, COL_VM_VORHAND).Value = "Maklervollmacht_Enthalten"
    ws.cells(1, COL_VM_VOLLST).Value = "MV_Vollumfaenglich"
    ws.cells(1, COL_VM_FEHLT).Value = "MV_Einschraenkungen"
    ws.cells(1, COL_HINWEIS).Value = "Hinweis"
    ws.cells(1, COL_VORGANGSTYP).Value = "Vorgangstyp"
    ws.cells(1, COL_UNTERSCHR_K).Value = "Unterschrift_Kunde"
    ws.cells(1, COL_UNTERSCHR_M).Value = "Unterschrift_Makler"
    ws.cells(1, COL_AUF_VN).Value = "Auf_VN_Ausgestellt"
    ws.cells(1, COL_MAKL_NAM).Value = "Makler_Namentlich"
    ws.cells(1, COL_REMINDER).Value = "Ist_Reminder"

    With ws.Range(ws.cells(1, 1), ws.cells(1, COL_REMINDER))
        .Font.Bold = True
        .Interior.Color = RGB(30, 64, 175) ' kraeftiges Blau
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22
    ws.Range("A2").Select
    On Error Resume Next: ActiveWindow.FreezePanes = False: ActiveWindow.FreezePanes = True: On Error GoTo 0
End Sub

Private Sub SpaltenbreitenSetzen(ws As Worksheet)
    ws.Columns(COL_DATEI).ColumnWidth = 30
    ws.Columns(COL_DATUM).ColumnWidth = 16
    ws.Columns(COL_ABS_NAME).ColumnWidth = 22
    ws.Columns(COL_ABS_MAIL).ColumnWidth = 28
    ws.Columns(COL_BETREFF).ColumnWidth = 40
    ws.Columns(COL_ANHANG_LST).ColumnWidth = 28
    ws.Columns(COL_MAKLERPOOL).ColumnWidth = 18
    ws.Columns(COL_NACHNAME).ColumnWidth = 18
    ws.Columns(COL_VORNAME).ColumnWidth = 14
    ws.Columns(COL_KLASSIFIK).ColumnWidth = 16
    ws.Columns(COL_GESCHTYP).ColumnWidth = 24
    ws.Columns(COL_UNTERLAGEN).ColumnWidth = 16
    ws.Columns(COL_SONDERFALL).ColumnWidth = 18
    ws.Columns(COL_SPARTE).ColumnWidth = 12
    ws.Columns(COL_ANH_TYPEN).ColumnWidth = 28
    ws.Columns(COL_VM_VORHAND).ColumnWidth = 18
    ws.Columns(COL_VM_VOLLST).ColumnWidth = 18
    ws.Columns(COL_VM_FEHLT).ColumnWidth = 32
    ws.Columns(COL_HINWEIS).ColumnWidth = 50
    ws.Columns(COL_VORGANGSTYP).ColumnWidth = 18
    ws.Columns(COL_UNTERSCHR_K).ColumnWidth = 14
    ws.Columns(COL_UNTERSCHR_M).ColumnWidth = 14
    ws.Columns(COL_AUF_VN).ColumnWidth = 14
    ws.Columns(COL_MAKL_NAM).ColumnWidth = 14
    ws.Columns(COL_REMINDER).ColumnWidth = 12
    ws.Columns(COL_VM_FEHLT).ColumnWidth = 36
End Sub

Private Sub SchreibeGptErgebnis(ws As Worksheet, row As Long, dict As Object)
    Dim vtyp As String: vtyp = SafeGet(dict, "vorgangstyp")
    ws.cells(row, COL_VORGANGSTYP).Value = vtyp
    ws.cells(row, COL_HINWEIS).Value = SafeGet(dict, "hinweis")

    Dim istMakler As Boolean: istMakler = (LCase(vtyp) = "makler-vorgang")

    If istMakler Then
        ws.cells(row, COL_MAKLERPOOL).Value = SafeGet(dict, "maklerpool")
        ws.cells(row, COL_NACHNAME).Value = SafeGet(dict, "makler_nachname")
        ws.cells(row, COL_VORNAME).Value = SafeGet(dict, "makler_vorname")
        ws.cells(row, COL_KLASSIFIK).Value = SafeGet(dict, "klassifikation")
        ws.cells(row, COL_GESCHTYP).Value = SafeGet(dict, "geschaefts_typ")
        ws.cells(row, COL_UNTERLAGEN).Value = SafeGet(dict, "unterlagen_angefragt")
        ws.cells(row, COL_SONDERFALL).Value = SafeGet(dict, "sonderfall")
        ws.cells(row, COL_SPARTE).Value = SafeGet(dict, "sparte")
        ws.cells(row, COL_ANH_TYPEN).Value = SafeGet(dict, "anhang_typen")
        ws.cells(row, COL_VM_VORHAND).Value = SafeGet(dict, "enthaelt_maklervollmacht")
        ws.cells(row, COL_VM_VOLLST).Value = SafeGet(dict, "mv_vollumfaenglich")
        ws.cells(row, COL_VM_FEHLT).Value = SafeGet(dict, "mv_einschraenkungen")
        ws.cells(row, COL_UNTERSCHR_K).Value = SafeGet(dict, "unterschrift_kunde")
        ws.cells(row, COL_UNTERSCHR_M).Value = SafeGet(dict, "unterschrift_makler")
        ws.cells(row, COL_AUF_VN).Value = SafeGet(dict, "auf_vn_ausgestellt")
        ws.cells(row, COL_MAKL_NAM).Value = SafeGet(dict, "makler_namentlich_genannt")
        ws.cells(row, COL_REMINDER).Value = SafeGet(dict, "ist_reminder")
        If LCase(SafeGet(dict, "ist_reminder")) = "ja" Then
            ws.cells(row, COL_REMINDER).Interior.Color = RGB(254, 215, 170) ' helles Orange
            ws.cells(row, COL_REMINDER).Font.Bold = True
        End If

        ' Hervorhebungen NUR fuer echte Makler-Vorgaenge
        Dim sonder As String: sonder = LCase(SafeGet(dict, "sonderfall"))
        Dim klass As String: klass = LCase(SafeGet(dict, "klassifikation"))
        Dim vmVollst As String: vmVollst = LCase(SafeGet(dict, "mv_vollumfaenglich"))
        Dim vmEnthalten As String: vmEnthalten = LCase(SafeGet(dict, "enthaelt_maklervollmacht"))

        If InStr(sonder, "flotten") > 0 Or InStr(sonder, "sondertarif") > 0 Then
            ws.cells(row, COL_SONDERFALL).Interior.Color = RGB(252, 165, 165)
        End If
        If InStr(klass, "nicht-standard") > 0 Then
            ws.cells(row, COL_KLASSIFIK).Interior.Color = RGB(253, 230, 138)
        End If
        If vmEnthalten = "ja" Then
            If vmVollst = "teilweise" Or vmVollst = "nein" Then
                ws.cells(row, COL_VM_VOLLST).Interior.Color = RGB(252, 165, 165)
            ElseIf vmVollst = "ja" Then
                ws.cells(row, COL_VM_VOLLST).Interior.Color = RGB(187, 247, 208)
            End If

            ' Unterschriften: rot wenn 'nein', gruen wenn 'ja'
            Dim uk As String: uk = LCase(SafeGet(dict, "unterschrift_kunde"))
            Dim um As String: um = LCase(SafeGet(dict, "unterschrift_makler"))
            If uk = "nein" Then
                ws.cells(row, COL_UNTERSCHR_K).Interior.Color = RGB(252, 165, 165)
            ElseIf uk = "ja" Then
                ws.cells(row, COL_UNTERSCHR_K).Interior.Color = RGB(187, 247, 208)
            End If
            If um = "nein" Then
                ws.cells(row, COL_UNTERSCHR_M).Interior.Color = RGB(252, 165, 165)
            ElseIf um = "ja" Then
                ws.cells(row, COL_UNTERSCHR_M).Interior.Color = RGB(187, 247, 208)
            End If
        End If
    Else
        ' Kein Makler-Vorgang: Felder G-Y leer lassen, Zeile grau einfaerben
        Dim grau As Long: grau = RGB(229, 231, 235)
        ws.Range(ws.cells(row, COL_MAKLERPOOL), ws.cells(row, COL_REMINDER)).Interior.Color = grau
        ws.cells(row, COL_VORGANGSTYP).Interior.Color = RGB(254, 215, 170) ' helles Orange
        ws.cells(row, COL_VORGANGSTYP).Font.Bold = True
    End If
End Sub

Private Function SafeGet(dict As Object, key As String) As String
    On Error Resume Next
    If dict.Exists(key) Then SafeGet = CStr(dict(key)) Else SafeGet = ""
End Function

Private Sub Bold(ws As Worksheet, row As Long, fontSize As Long)
    With ws.cells(row, 1).Font: .Bold = True: .Size = fontSize: End With
End Sub

' Schreibt fuer einen Vorgang eine .txt-Datei mit allen KI-Analyse-Daten.
' Dateiname: <msg-Basename>.txt im uebergebenen Ordner (typischerweise
' '_KI-Analyse' im Vorgangs-Ordner).
Private Sub SchreibeKiTextDatei(ordner As String, msgName As String, _
                                datum As String, absName As String, absMail As String, _
                                betreff As String, anhangAlle As String, pdfListe As String, _
                                dict As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String: base = msgName
    If LCase(Right(base, 4)) = ".msg" Then base = Left(base, Len(base) - 4)
    base = SaeubereDateiname(base)
    Dim outPath As String: outPath = ordner & "\" & base & ".txt"

    Dim modellAktuell As String: modellAktuell = ""
    On Error Resume Next
    modellAktuell = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_GPT).Range("A6").Value2))
    On Error GoTo 0

    Dim s As String
    Dim sep As String: sep = String(72, "=") & vbCrLf
    Dim sub_ As String: sub_ = String(40, "-") & vbCrLf

    s = sep
    s = s & "ERGO Vorgang-Analyse  -  KI-Klassifikation" & vbCrLf
    s = s & sep & vbCrLf

    s = s & "Datei:           " & msgName & vbCrLf
    s = s & "Datum:           " & datum & vbCrLf
    s = s & "Absender:        " & absName & "  <" & absMail & ">" & vbCrLf
    s = s & "Betreff:         " & betreff & vbCrLf
    s = s & "Anhaenge (alle): " & IIf(Len(anhangAlle) = 0, "(keine)", anhangAlle) & vbCrLf
    s = s & "Anhaenge (PDF):  " & IIf(Len(pdfListe) = 0, "(keine)", pdfListe) & vbCrLf
    s = s & vbCrLf

    s = s & sub_
    s = s & "TRIAGE" & vbCrLf
    s = s & sub_
    s = s & "Vorgangstyp:           " & SafeGet(dict, "vorgangstyp") & vbCrLf
    s = s & "Ist Reminder:          " & SafeGet(dict, "ist_reminder") & vbCrLf
    s = s & vbCrLf

    s = s & sub_
    s = s & "KLASSIFIKATION" & vbCrLf
    s = s & sub_
    s = s & "Klassifikation:        " & SafeGet(dict, "klassifikation") & vbCrLf
    s = s & "Geschaefts_Typ:        " & SafeGet(dict, "geschaefts_typ") & vbCrLf
    s = s & "Unterlagen_Angefragt:  " & SafeGet(dict, "unterlagen_angefragt") & vbCrLf
    s = s & "Sonderfall:            " & SafeGet(dict, "sonderfall") & vbCrLf
    s = s & "Sparte:                " & SafeGet(dict, "sparte") & vbCrLf
    s = s & vbCrLf

    s = s & sub_
    s = s & "MAKLER" & vbCrLf
    s = s & sub_
    s = s & "Maklerpool:            " & SafeGet(dict, "maklerpool") & vbCrLf
    s = s & "Nachname:              " & SafeGet(dict, "makler_nachname") & vbCrLf
    s = s & "Vorname:               " & SafeGet(dict, "makler_vorname") & vbCrLf
    s = s & vbCrLf

    s = s & sub_
    s = s & "ANHANG-ANALYSE" & vbCrLf
    s = s & sub_
    s = s & "Anhang_Typen:          " & SafeGet(dict, "anhang_typen") & vbCrLf
    s = s & vbCrLf

    s = s & sub_
    s = s & "MAKLERVOLLMACHT (juristische Pruefung nach AAW-Allg)" & vbCrLf
    s = s & sub_
    s = s & "Enthalten:             " & SafeGet(dict, "enthaelt_maklervollmacht") & vbCrLf
    s = s & "Vollumfaenglich:       " & SafeGet(dict, "mv_vollumfaenglich") & vbCrLf
    s = s & "Einschraenkungen:      " & SafeGet(dict, "mv_einschraenkungen") & vbCrLf
    s = s & "Unterschrift Kunde:    " & SafeGet(dict, "unterschrift_kunde") & "  (Pflicht)" & vbCrLf
    s = s & "Unterschrift Makler:   " & SafeGet(dict, "unterschrift_makler") & "  (optional)" & vbCrLf
    s = s & "Auf VN ausgestellt:    " & SafeGet(dict, "auf_vn_ausgestellt") & vbCrLf
    s = s & "Makler namentlich:     " & SafeGet(dict, "makler_namentlich_genannt") & vbCrLf
    s = s & vbCrLf

    s = s & sub_
    s = s & "HINWEIS" & vbCrLf
    s = s & sub_
    s = s & SafeGet(dict, "hinweis") & vbCrLf
    s = s & vbCrLf

    s = s & sep
    s = s & "Erstellt: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    s = s & "Modell:   " & modellAktuell & vbCrLf
    s = s & sep

    Dim ts As Object: Set ts = fso.CreateTextFile(outPath, True, True) ' True=Unicode
    ts.Write s
    ts.Close
End Sub

' ============================================================================
' === INLINE ASK_ErgoGPT (frueher in separatem Modul / Test.txt) =============
' ============================================================================
' Ab v2.7 ist ASK_ErgoGPT direkt hier integriert, damit die Mappe nur noch
' EIN Modul braucht. Die alten Module 'Agent', 'AI_Excel_Functions' und
' 'GPT' koennen geloescht werden - sofern dort keine andere Funktionalitaet
' (XVERWEIS-AI etc.) drinsteht, die dauerhaft gebraucht wird.
' ============================================================================

' === Haupteinstieg ==========================================================
Public Function ASK_ErgoGPT(userPrompt As String, Optional pdfs As Variant) As String
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim cookie As String, bearer As String
    Dim payload As String, convId As String, resp As String, answer As String
    Dim docIdsJson As String

    cookie = ReadCookieFromFile()
    bearer = ""

    ' 1) Conversation anlegen
    payload = BuildPayload(CleanForJson(userPrompt), "[]")
    http.Open "PUT", ERGO_BASE_URL & "/conversation", False
    SetCommonHeaders http, cookie, bearer
    http.SetRequestHeader "Content-Type", "application/json;charset=UTF-8"
    http.Send payload
    If http.status <> 200 Then Err.Raise vbObjectError + 2, , "Create failed: " & http.status & " - " & http.ResponseText

    convId = ParseJsonStr(http.ResponseText, "id")
    If Len(convId) = 0 Then Err.Raise vbObjectError + 3, , "Keine conversation_id gefunden."

    ' 2) PDFs hochladen (parallel, bis zu 3)
    docIdsJson = UploadManyPdfsReturnJsonArray(convId, pdfs, cookie, bearer)

    ' 3) Antwort holen
    payload = BuildPayload(CleanForJson(userPrompt), docIdsJson)
    http.Open "PUT", ERGO_BASE_URL & "/conversation/" & convId, False
    SetCommonHeaders http, cookie, bearer
    http.SetRequestHeader "Content-Type", "application/json;charset=UTF-8"
    http.SetRequestHeader "Accept", "*/*"
    http.Send payload

    Dim status2 As Long: status2 = http.status
    resp = http.ResponseText
    answer = ExtractAssistantFromNdjson(resp)

    If Len(answer) = 0 Then
        ' Conversation aufraeumen, damit keine Karteileichen in ErgoGPT bleiben
        On Error Resume Next
        Dim delPayload As String: delPayload = BuildDeletePayload(convId)
        http.Open "DELETE", ERGO_BASE_URL & "/conversation", False
        SetCommonHeaders http, cookie, bearer
        http.SetRequestHeader "Accept", "*/*"
        http.SetRequestHeader "Content-Type", "application/json;charset=UTF-8"
        http.Send delPayload
        On Error GoTo 0

        Dim hint As String: hint = ExtractServerErrorHint(resp)
        Dim msg As String
        If status2 <> 200 Then
            msg = "Antwort-Call: HTTP " & status2 & ". " & hint
        Else
            msg = "Keine Assistant-Nachricht gefunden. " & hint
        End If
        Err.Raise vbObjectError + 4, , msg
    End If

    ASK_ErgoGPT = Replace(JsonUnescape(answer), "\n", vbCrLf)

    ' 4) Aufraeumen
    payload = BuildDeletePayload(convId)
    http.Open "DELETE", ERGO_BASE_URL & "/conversation", False
    SetCommonHeaders http, cookie, bearer
    http.SetRequestHeader "Accept", "*/*"
    http.SetRequestHeader "Content-Type", "application/json;charset=UTF-8"
    http.Send payload
End Function

' === Cookie ==================================================================
Private Function ReadCookieFromFile() As String
    Dim fso As Object, ts As Object, s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(COOKIE_PATH) Then
        Err.Raise vbObjectError + 1, , "Cookie-Datei nicht gefunden: " & COOKIE_PATH
    End If
    Set ts = fso.OpenTextFile(COOKIE_PATH, 1, False, -2)
    s = ts.ReadAll
    ts.Close
    s = Trim$(s)
    If Len(s) = 0 Then Err.Raise vbObjectError + 1, , "Cookie-Datei ist leer: " & COOKIE_PATH
    ReadCookieFromFile = s
End Function

' === Payload-Bau =============================================================
Private Function BuildPayload(prompt As String, Optional docIdsJson As String = "[]") As String
    Dim model As String, t As Double, tone As String, tempStr As String
    Dim raw As Variant

    raw = ThisWorkbook.Worksheets(SHEET_GPT).Range("A6").Value2
    If IsError(raw) Or IsNull(raw) Or IsEmpty(raw) Then
        model = "gpt-51-chat"
    Else
        model = Trim$(CStr(raw))
        model = Replace(model, vbCr, ""): model = Replace(model, vbLf, ""): model = Replace(model, vbTab, "")
        If Len(model) = 0 Then model = "gpt-4o"
    End If

    raw = ThisWorkbook.Worksheets(SHEET_GPT).Range("A9").Value2
    If IsError(raw) Or IsNull(raw) Or IsEmpty(raw) Then
        t = 0
    ElseIf IsNumeric(raw) Then
        t = CDbl(raw)
    Else
        t = 0
    End If
    tempStr = Replace(CStr(t), ",", ".")

    raw = ThisWorkbook.Worksheets(SHEET_GPT).Range("A12").Value2
    If IsError(raw) Or IsNull(raw) Or IsEmpty(raw) Then
        tone = ""
    Else
        tone = Trim$(CStr(raw))
    End If

    BuildPayload = "{""conversation_title"":""New conversation"",""messages"":[{" & _
                   """id"":""" & Uuid4() & """,""role"":""user"",""content"":""" & prompt & """,""date"":""" & IsoNow() & """}]," & _
                   """model"":""" & model & """,""temperature"":" & tempStr & ",""pinned"":false,""tone"":""" & JsonEscape(tone) & """,""document_ids"":" & docIdsJson & "}"
End Function

Private Function BuildDeletePayload(conv_id As String) As String
    BuildDeletePayload = "{""conversation_ids"":[""" & conv_id & """]}"
End Function

' === HTTP-Header =============================================================
Private Sub SetCommonHeaders(http As Object, ByVal cookie As String, ByVal bearer As String)
    If Len(cookie) > 0 Then http.SetRequestHeader "Cookie", cookie
    Dim xsrf As String: xsrf = ExtractCsrfFromCookie(cookie)
    If Len(xsrf) > 0 Then
        http.SetRequestHeader "X-XSRF-TOKEN", xsrf
        http.SetRequestHeader "X-CSRF-TOKEN", xsrf
    End If
    If Len(bearer) > 0 Then http.SetRequestHeader "Authorization", bearer
    http.SetRequestHeader "User-Agent", "Mozilla/5.0"
    http.SetRequestHeader "Accept-Language", "de-DE,de;q=0.9,en;q=0.8"
    http.SetRequestHeader "Origin", "https://gpt.ergo.com"
    http.SetRequestHeader "Referer", "https://gpt.ergo.com/"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
End Sub

' === NDJSON / Server-Antwort ================================================
Private Function ExtractAssistantFromNdjson(ByVal raw As String) As String
    Dim lines() As String, i As Long, ln As String, lastText As String
    raw = Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(raw, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If LCase$(Left$(ln, 5)) = "data:" Then ln = Trim$(Mid$(ln, 6))
            If Left$(ln, 1) = "{" And Right$(ln, 1) = "}" Then
                Dim t As String: t = ExtractAssistantFromChunk(ln)
                If Len(t) > 0 Then lastText = t
            End If
        End If
    Next
    ExtractAssistantFromNdjson = lastText
End Function

Private Function ExtractAssistantFromChunk(json As String) As String
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.MultiLine = True: re.IgnoreCase = True

    re.Pattern = """choices""\s*:\s*\[\s*\{[\s\S]*?""messages""\s*:\s*\[\s*\{[\s\S]*?""role""\s*:\s*""assistant""[\s\S]*?""content""\s*:\s*""((?:\\.|[^""])*)"""
    If re.test(json) Then
        Set ms = re.Execute(json)
        ExtractAssistantFromChunk = ms(ms.count - 1).SubMatches(0)
        Exit Function
    End If

    re.Pattern = """choices""\s*:\s*\[[\s\S]*?""delta""[\s\S]*?""content""\s*:\s*""((?:\\.|[^""])*)"""
    If re.test(json) Then
        Set ms = re.Execute(json)
        ExtractAssistantFromChunk = ms(ms.count - 1).SubMatches(0)
    End If
End Function

Private Function ExtractServerErrorHint(ByVal raw As String) As String
    If Len(raw) = 0 Then ExtractServerErrorHint = "(leere Antwort)": Exit Function

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True

    Dim keys As Variant: keys = Array("message", "error", "detail", "error_description")
    Dim k As Variant, m As String
    For Each k In keys
        re.Pattern = """" & k & """\s*:\s*""((?:\\.|[^""]){0,500})"""
        If re.test(raw) Then
            Dim ms As Object: Set ms = re.Execute(raw)
            m = ms(0).SubMatches(0)
            If Len(m) > 0 Then
                ExtractServerErrorHint = "Server-Hinweis: " & Left$(m, 240)
                Exit Function
            End If
        End If
    Next k

    Dim s As String: s = LCase$(raw)
    If InStr(s, "nicht mehr unterst") > 0 Or InStr(s, "not supported") > 0 _
       Or InStr(s, "deprecated") > 0 Or InStr(s, "unknown model") > 0 Then
        ExtractServerErrorHint = "Server-Hinweis: Modell wird nicht (mehr) unterstuetzt - Sheet GPT!A6 anpassen."
        Exit Function
    End If

    Dim plain As String: plain = raw
    re.Pattern = "<[^>]+>": plain = re.Replace(plain, " ")
    plain = Replace(plain, vbCrLf, " "): plain = Replace(plain, vbLf, " "): plain = Replace(plain, vbTab, " ")
    Do While InStr(plain, "  ") > 0: plain = Replace(plain, "  ", " "): Loop
    plain = Trim$(plain)
    If Len(plain) > 200 Then plain = Left$(plain, 197) & "..."
    ExtractServerErrorHint = "Antwort-Anfang: " & plain
End Function

' === JSON-Helpers ===========================================================
Private Function ParseJsonStr(json As String, key As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & key & """" & "\s*:\s*""([^""]*)"""
    re.IgnoreCase = True
    If re.test(json) Then ParseJsonStr = re.Execute(json)(0).SubMatches(0)
End Function

Private Function JsonEscape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, " ")
    JsonEscape = s
End Function

Private Function JsonUnescape(s As String) As String
    Dim re As Object, mc As Object, it As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.Pattern = "\\u([0-9a-fA-F]{4})"
    Set mc = re.Execute(s)
    For Each it In mc
        s = Replace(s, it.Value, ChrW(CLng("&H" & it.SubMatches(0))))
    Next
    s = Replace(s, "\""", """")
    s = Replace(s, "\\/", "/")
    s = Replace(s, "\\n", vbCrLf)
    s = Replace(s, "\\r", "")
    s = Replace(s, "\\t", vbTab)
    s = Replace(s, "\\\\", "\")
    JsonUnescape = s
End Function

Private Function CleanForJson(ByVal s As String) As String
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    CleanForJson = JsonEscape(s)
End Function

Private Function IsoNow() As String
    IsoNow = Format$(Now, "yyyy-mm-dd\THH:nn:ss") & ".000Z"
End Function

Private Function Uuid4() As String
    Dim a(): a = Array(8, 4, 4, 4, 12)
    Dim i As Long, j As Long, s As String
    Randomize
    For i = 0 To UBound(a)
        For j = 1 To a(i): s = s & LCase$(Hex$(Int(Rnd() * 16))): Next
        If i < UBound(a) Then s = s & "-"
    Next
    Uuid4 = s
End Function

Private Function ExtractCsrfFromCookie(ByVal cookie As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = "(?:XSRF-TOKEN|CSRF-TOKEN|csrfToken|_csrf)=([^;]+)"
    If re.test(cookie) Then ExtractCsrfFromCookie = UrlDecode(re.Execute(cookie)(0).SubMatches(0))
End Function

Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long, r As String, ch As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "+" Then
            r = r & " "
        ElseIf ch = "%" And i + 2 <= Len(s) Then
            r = r & Chr(CLng("&H" & Mid$(s, i + 1, 2))): i = i + 2
        Else
            r = r & ch
        End If
        i = i + 1
    Loop
    UrlDecode = r
End Function

' === PDF-Upload (parallel) ==================================================
Private Function UploadManyPdfsReturnJsonArray(convId As String, pdfs As Variant, cookie As String, bearer As String) As String
    Dim paths As Collection: Set paths = PathsFromVariant(pdfs)
    If paths Is Nothing Or paths.count = 0 Then
        UploadManyPdfsReturnJsonArray = "[]"
        Exit Function
    End If

    Dim xsrf As String: xsrf = ExtractCsrfFromCookie(cookie)
    Dim url As String: url = "https://gpt.ergo.com/api/document"

    Dim ids() As String: ReDim ids(1 To paths.count)
    Dim inflight As Object: Set inflight = CreateObject("Scripting.Dictionary")
    Dim nextIdx As Long: nextIdx = 1

    Do While nextIdx <= paths.count Or inflight.count > 0
        Do While nextIdx <= paths.count And inflight.count < MAX_PARALLEL_PDF_UPLOADS
            Dim p As String: p = Trim$(paths(nextIdx))
            If Len(p) = 0 Then
                ids(nextIdx) = ""
                nextIdx = nextIdx + 1
            Else
                Dim win As Object: Set win = CreateObject("WinHttp.WinHttpRequest.5.1")
                Dim boundary As String: boundary = "----WebKitFormBoundary" & Left$(Replace(Uuid4(), "-", ""), 16)
                Dim filename As String: filename = CreateObject("Scripting.FileSystemObject").GetFileName(p)

                Dim pre As String, post As String, body As Variant
                pre = "--" & boundary & vbCrLf & _
                      "Content-Disposition: form-data; name=""file_content""; filename=""" & filename & """" & vbCrLf & _
                      "Content-Type: application/pdf" & vbCrLf & vbCrLf
                post = vbCrLf & _
                       "--" & boundary & vbCrLf & _
                       "Content-Disposition: form-data; name=""filename""" & vbCrLf & vbCrLf & _
                       filename & vbCrLf & _
                       "--" & boundary & vbCrLf & _
                       "Content-Disposition: form-data; name=""conversation_id""" & vbCrLf & vbCrLf & _
                       convId & vbCrLf & _
                       "--" & boundary & "--" & vbCrLf
                body = BuildMultipartBinary_NoBom(pre, p, post)

                win.Open "POST", url, True
                On Error Resume Next: win.SetProxy 0: On Error GoTo 0
                win.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129 Safari/537.36"
                win.SetRequestHeader "Accept", "*/*"
                win.SetRequestHeader "Accept-Language", "de-DE,de;q=0.9,en;q=0.8"
                win.SetRequestHeader "Cookie", cookie
                win.SetRequestHeader "Origin", "https://gpt.ergo.com"
                win.SetRequestHeader "Referer", "https://gpt.ergo.com/"
                If Len(xsrf) > 0 Then
                    win.SetRequestHeader "X-XSRF-TOKEN", xsrf
                    win.SetRequestHeader "X-CSRF-TOKEN", xsrf
                End If
                win.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
                On Error Resume Next
                win.SetRequestHeader "Content-Length", CStr(UBound(body) - LBound(body) + 1)
                On Error GoTo 0

                win.Send body
                inflight.Add CStr(nextIdx), Array(win, p)
                nextIdx = nextIdx + 1
            End If
        Loop

        Dim k As Variant, finished As Collection: Set finished = New Collection
        For Each k In inflight.Keys
            Dim req As Object, meta As Variant
            meta = inflight(k): Set req = meta(0)
            If req.WaitForResponse(0) Then
                If req.status = 200 Then
                    Dim did As String: did = ParseJsonStr(req.ResponseText, "document_id")
                    If Len(did) = 0 Then
                        Err.Raise vbObjectError + 211, , "document_id fehlt bei " & meta(1) & " - " & Left$(req.ResponseText, 200)
                    End If
                    ids(CLng(k)) = did
                Else
                    Err.Raise vbObjectError + 210, , "Upload fehlgeschlagen (" & meta(1) & "): " & req.status & " - " & Left$(req.ResponseText, 300)
                End If
                finished.Add k
            End If
        Next
        For Each k In finished
            inflight.Remove k
        Next

        If inflight.count > 0 Then DoEvents
    Loop

    Dim i As Long, s As String
    For i = 1 To UBound(ids)
        If Len(ids(i)) > 0 Then
            If Len(s) > 0 Then s = s & ","
            s = s & """" & ids(i) & """"
        End If
    Next
    UploadManyPdfsReturnJsonArray = "[" & s & "]"
End Function

Private Function PathsFromVariant(v As Variant) As Collection
    Dim col As Collection: Set col = New Collection
    Dim i As Long

    If IsMissing(v) Then GoTo done
    If IsEmpty(v) Then GoTo done
    If TypeName(v) = "Nothing" Then GoTo done

    Select Case TypeName(v)
        Case "String"
            Dim s As String: s = Trim$(CStr(v))
            If Len(s) > 0 Then
                Dim arr As Variant
                s = Replace(s, vbCrLf, vbLf): s = Replace(s, vbCr, vbLf)
                s = Replace(Replace(Replace(s, ";", vbLf), "|", vbLf), ",", vbLf)
                arr = Split(s, vbLf)
                For i = LBound(arr) To UBound(arr)
                    If Len(Trim$(arr(i))) > 0 Then col.Add Trim$(arr(i))
                Next
            End If
        Case "Range"
            Dim c As Range
            For Each c In v.cells
                If Len(Trim$(CStr(c.Value))) > 0 Then col.Add Trim$(CStr(c.Value))
            Next
        Case "Collection"
            For i = 1 To v.count
                If Len(Trim$(CStr(v(i)))) > 0 Then col.Add Trim$(CStr(v(i)))
            Next
        Case Else
            If IsArray(v) Then
                For i = LBound(v) To UBound(v)
                    If Len(Trim$(CStr(v(i)))) > 0 Then col.Add Trim$(CStr(v(i)))
                Next
            Else
                col.Add Trim$(CStr(v))
            End If
    End Select
done:
    Set PathsFromVariant = col
End Function

Private Function BuildMultipartBinary_NoBom(pre As String, filePath As String, post As String) As Variant
    Dim prefix() As Byte, suffix() As Byte, fileBytes() As Byte
    prefix = Utf8NoBomBytes(pre)
    suffix = Utf8NoBomBytes(post)
    fileBytes = ReadAllBytes(filePath)

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    If (Not Not prefix) <> 0 Then stm.Write prefix
    If (Not Not fileBytes) <> 0 Then stm.Write fileBytes
    If (Not Not suffix) <> 0 Then stm.Write suffix
    stm.Position = 0
    BuildMultipartBinary_NoBom = stm.Read
    stm.Close
End Function

Private Function Utf8NoBomBytes(ByVal s As String) As Byte()
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s, 0
    st.Position = 0
    st.Type = 1
    Dim b() As Byte: b = st.Read
    st.Close

    If UBound(b) >= 2 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            Dim res() As Byte, i As Long
            ReDim res(0 To UBound(b) - 3)
            For i = 3 To UBound(b)
                res(i - 3) = b(i)
            Next
            Utf8NoBomBytes = res
            Exit Function
        End If
    End If
    Utf8NoBomBytes = b
End Function

Private Function ReadAllBytes(fp As String) As Byte()
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.LoadFromFile fp
    ReadAllBytes = st.Read
    st.Close
End Function
