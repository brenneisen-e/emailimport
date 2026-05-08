Attribute VB_Name = "ErgoVorgangAnalyse"
' ============================================================================
' ERGO VORGANG-ANALYSE - Excel-VBA-Tool (v2.5)
' ============================================================================
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
'        - Default-Modell: gpt-5.1
'   3) Alt+F8 -> Vorgaenge_Analysieren -> Ausfuehren
'        - fragt: Welcher Ordner? (Default: dieser Workbook-Pfad)
'        - fragt: Wie viele Vorgaenge maximal? (leer = alle)
'        - laedt pro Vorgang: Email-Text + PDF-Anhaenge zu ErgoGPT hoch
'        - GPT klassifiziert + prueft Maklervollmacht-Vollstaendigkeit
'        - Ergebnis landet als Zeile im Sheet 'Analyse'
'
' VORAUSSETZUNGEN:
'   - Modul mit Funktion ASK_ErgoGPT(prompt, pdfs) muss bereits in der
'     Mappe importiert sein (aus Test.txt im Repo - die volle Variante
'     mit PDF-Upload-Faehigkeit!).
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
'   Q  Vollmacht_Vollstaendig      (ja / nein / teilweise / nicht_pruefbar)
'   R  Vollmacht_Fehlt             (Liste fehlender Felder, kommagetrennt)
'   S  Hinweis                     (kurzer GPT-Hinweis zum Vorgang)
'   T  Vorgangstyp                 (Triage: Makler-Vorgang / Bounce-NDR /
'                                   Ergo-Outbound / System-Mail /
'                                   Werbung-Spam / Unklar)
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
        erfolg = AnalysiereEineDatei(olApp, fso, msgPath, ws, row, tempBase, i)
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
                                     vorgangNr As Long) As Boolean
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
    p = p & "11) maklervollmacht_vollstaendig" & vbCrLf
    p = p & "    Pruefung der Pflichtfelder einer Maklervollmacht:" & vbCrLf
    p = p & "      a) Vollstaendiger Kundenname (Vor- + Nachname)" & vbCrLf
    p = p & "      b) Geburtsdatum des Kunden" & vbCrLf
    p = p & "      c) Anschrift des Kunden" & vbCrLf
    p = p & "      d) Datum der Vollmacht" & vbCrLf
    p = p & "      e) Eigenhaendige Unterschrift des Kunden (sichtbar)" & vbCrLf
    p = p & "      f) Maklerangaben (Name + Firma + Vermittlernummer/BD-Nummer)" & vbCrLf
    p = p & "    Werte:" & vbCrLf
    p = p & "      - 'ja': alle 6 Pflichtfelder erkennbar vorhanden" & vbCrLf
    p = p & "      - 'teilweise': Vollmacht da, einzelne Felder fehlen" & vbCrLf
    p = p & "      - 'nein': als Vollmacht erkennbar, aber Pflichtfelder grosszuegig fehlen" & vbCrLf
    p = p & "      - 'nicht_pruefbar': keine Vollmacht im Anhang oder PDF nicht ausgewertet" & vbCrLf
    p = p & vbCrLf
    p = p & "12) maklervollmacht_fehlt" & vbCrLf
    p = p & "    Komma-getrennte Liste der fehlenden Felder (a-f aus Punkt 11)," & vbCrLf
    p = p & "    benannt z.B.: 'Geburtsdatum, Unterschrift Kunde'. Leer wenn" & vbCrLf
    p = p & "    Vollmacht vollstaendig oder nicht pruefbar." & vbCrLf
    p = p & vbCrLf
    p = p & "13) hinweis" & vbCrLf
    p = p & "    EIN kurzer Satz (max 200 Zeichen) was an dem Vorgang auffaellig" & vbCrLf
    p = p & "    ist - z.B. fehlende Unterlagen, ungewoehnliche Konstellation," & vbCrLf
    p = p & "    Eskalationspotenzial. Leer wenn nichts auffaellt." & vbCrLf
    p = p & vbCrLf
    p = p & "AUSGABE-FORMAT (eine Zeile, gueltiges JSON, alle 14 Schluessel):" & vbCrLf
    p = p & "{""vorgangstyp"":""..."",""maklerpool"":""..."",""makler_nachname"":""...""," & vbCrLf
    p = p & " ""makler_vorname"":""..."",""klassifikation"":""..."",""geschaefts_typ"":""...""," & vbCrLf
    p = p & " ""unterlagen_angefragt"":""..."",""sonderfall"":""..."",""sparte"":""...""," & vbCrLf
    p = p & " ""anhang_typen"":""..."",""enthaelt_maklervollmacht"":""...""," & vbCrLf
    p = p & " ""maklervollmacht_vollstaendig"":""..."",""maklervollmacht_fehlt"":""...""," & vbCrLf
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
                 "enthaelt_maklervollmacht", "maklervollmacht_vollstaendig", _
                 "maklervollmacht_fehlt", "hinweis")
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

    ws.Range("B6").Value = "<- Modell-Name EXAKT wie im ErgoGPT-Browser (oben rechts ablesen). Beispiele: 'gpt-5.1', 'gpt-4o'. Aenderung sofort uebernehmen."
    ws.Range("B7").Value = "<- Cookie als Text (langer String) - ODER leer lassen und A8/Dialog nutzen"
    ws.Range("B8").Value = "<- Pfad zu Cookie-Datei (z.B. C:\Users\...\Desktop\cookie.txt) - leer = Default F:\ExcelGPT-Cookie\Cookie.txt"
    ws.Range("B9").Value = "<- Temperature (0 = deterministisch)"
    ws.Range("B12").Value = "<- Tone (z.B. 'Sachlich' oder leer)"

    If Trim(CStr(ws.Range("A6").Value)) = "" Then ws.Range("A6").Value = "gpt-5.1"
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
    ws.cells(r, 1).Value = "- Modul ASK_ErgoGPT (volle Variante mit PDF-Upload aus Test.txt) ist importiert.": r = r + 1
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
    ws.cells(1, COL_VM_VOLLST).Value = "Vollmacht_Vollstaendig"
    ws.cells(1, COL_VM_FEHLT).Value = "Vollmacht_Fehlt"
    ws.cells(1, COL_HINWEIS).Value = "Hinweis"
    ws.cells(1, COL_VORGANGSTYP).Value = "Vorgangstyp"

    With ws.Range(ws.cells(1, 1), ws.cells(1, COL_VORGANGSTYP))
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
        ws.cells(row, COL_VM_VOLLST).Value = SafeGet(dict, "maklervollmacht_vollstaendig")
        ws.cells(row, COL_VM_FEHLT).Value = SafeGet(dict, "maklervollmacht_fehlt")

        ' Hervorhebungen NUR fuer echte Makler-Vorgaenge
        Dim sonder As String: sonder = LCase(SafeGet(dict, "sonderfall"))
        Dim klass As String: klass = LCase(SafeGet(dict, "klassifikation"))
        Dim vmVollst As String: vmVollst = LCase(SafeGet(dict, "maklervollmacht_vollstaendig"))
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
        End If
    Else
        ' Kein Makler-Vorgang: Felder G-S leer lassen, Zeile grau einfaerben
        Dim grau As Long: grau = RGB(229, 231, 235)
        ws.Range(ws.cells(row, COL_MAKLERPOOL), ws.cells(row, COL_HINWEIS)).Interior.Color = grau
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
