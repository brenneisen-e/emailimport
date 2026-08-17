# BGAV Hypercare - Email Review Tool

Tool zur Überprüfung und Kategorisierung von Hypercare-Emails für das Barmenia/Gothaer-Projekt.

**Version 1.5.0**

---

## Übersicht

Das Tool besteht aus zwei Komponenten:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           WORKFLOW ÜBERSICHT                                 │
└─────────────────────────────────────────────────────────────────────────────┘

  ┌──────────────┐         ┌──────────────┐         ┌──────────────┐
  │   OUTLOOK    │         │   WEB-APP    │         │    EXCEL     │
  │   Postfach   │         │   Browser    │         │   Tracking   │
  └──────┬───────┘         └──────┬───────┘         └──────┬───────┘
         │                        │                        │
         │  1. HTA Export Tool    │                        │
         │  ─────────────────►    │                        │
         │     JSON-Datei         │                        │
         │                        │                        │
         │                        │  2. Kategorisieren     │
         │                        │  ◄────────────────►    │
         │                        │     Bearbeiten         │
         │                        │                        │
         │                        │  3. HTA Import         │
         │                        │  ─────────────────►    │
         │                        │     Excel Update       │
         │                        │                        │
         └────────────────────────┴────────────────────────┘
```

---

## Teil 1: Email-Export (HTA Tool)

### Was ist das HTA Tool?

Das **HTA (HTML Application)** ist ein Windows-Desktop-Tool, das direkt auf Outlook zugreift und Emails exportiert. Es läuft nur auf Windows mit installiertem Outlook.

### Download

Die Datei `Email-Export-Tool.zip` enthält:
```
Email-Export-Tool/
├── outlook-export-modular.hta    # Hauptdatei (Doppelklick zum Starten)
├── js/                           # JavaScript-Module
└── css/                          # Stylesheets
```

### Funktionsweise Export

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         HTA EXPORT PROZESS                                   │
└─────────────────────────────────────────────────────────────────────────────┘

  SCHRITT 1: Excel verbinden
  ┌────────────────────────────────────────────┐
  │ • Tracking_Hypercare.xlsm muss offen sein  │
  │ • Tool erkennt automatisch die Datei       │
  │ • Liest existierende Vorgänge aus Excel    │
  │   (zur Duplikat-Erkennung)                 │
  └────────────────────────────────────────────┘
                      ↓
  SCHRITT 2: Outlook Postfach wählen
  ┌────────────────────────────────────────────┐
  │ • Postfach auswählen (z.B. Hypercare)      │
  │ • Ordner wählen (Inbox)                    │
  │ • Zeitraum festlegen (z.B. 30 Tage)        │
  └────────────────────────────────────────────┘
                      ↓
  SCHRITT 3: Export durchführen
  ┌────────────────────────────────────────────┐
  │ Phase 1: Inbox-Emails lesen                │
  │ Phase 2: Sent-Items lesen                  │
  │ Phase 3: Konversationen gruppieren         │
  │ Phase 4: JSON-Datei speichern              │
  └────────────────────────────────────────────┘
                      ↓
  ERGEBNIS: hypercare_conversations_2025-12-15.json
```

### Konversations-Gruppierung

Das Tool gruppiert Emails automatisch nach **ConversationID** (Outlook-intern):

```
Beispiel: Konversation über Provisionsreklamation

┌─────────────────────────────────────────────────────────────────────────────┐
│ ConversationID: AAA-BBB-CCC                                                  │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  01.12. 10:00  [INBOX]  Kunde → Uns     "Provisionsreklamation..."          │
│       │        depth=0  ← ANFRAGE (Root)                                    │
│       │                                                                      │
│       └─► 01.12. 14:00  [SENT]  Wir → Kunde  "Danke für Ihre Anfrage..."   │
│                 depth=1  ← ANTWORT                                          │
│                    │                                                         │
│                    └─► 02.12. 09:00  [INBOX]  Kunde → Uns  "Rückfrage..."  │
│                              depth=2  ← ANTWORT                             │
│                                 │                                            │
│                                 └─► 02.12. 11:00  [SENT]  "Erledigt..."    │
│                                           depth=3  ← ANTWORT                │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘

JSON-Ausgabe:
{
  "conversationID": "AAA-BBB-CCC",
  "messages": [
    { "folder": "inbox", "depth": 0, "subject": "Provisionsreklamation..." },
    { "folder": "sent",  "depth": 1, "subject": "AW: Provisionsreklamation..." },
    { "folder": "inbox", "depth": 2, "subject": "AW: AW: Provisionsreklamation..." },
    { "folder": "sent",  "depth": 3, "subject": "AW: AW: AW: Provisionsreklamation..." }
  ]
}
```

### Sent-Only Konversationen

Manchmal liegt die ursprüngliche Anfrage außerhalb des Export-Zeitraums:

```
Szenario: Export der letzten 7 Tage, aber Anfrage ist 30 Tage alt

┌─────────────────────────────────────────────────────────────────────────────┐
│ ConversationID: XYZ-123                                                      │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  15.11. 10:00  [INBOX]  Kunde → Uns     "Anfrage..."                        │
│       │        ↑ AUSSERHALB DES EXPORT-ZEITRAUMS (nicht im JSON!)           │
│       │                                                                      │
│       └─► 10.12. 14:00  [SENT]  Wir → Kunde  "Antwort..."                  │
│                 depth=1  ← Nur diese Nachricht im Export                    │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘

→ Diese Konversation wird als "sentOnly: true" markiert
→ Beim Import: Nur zu EXISTIERENDEN Fällen hinzufügen, keine neue Anfrage erstellen
```

---

## Teil 2: Web-App (Kategorisierung)

### Was ist die Web-App?

Die **Web-App** (`index.html`) ist ein Browser-Tool zum Anzeigen, Kategorisieren und Bearbeiten der exportierten Emails.

### Zugang

- **Online**: https://emailimport.pages.dev (oder eigene Cloudflare Pages URL)
- **Lokal**: `index.html` direkt im Browser öffnen
- **Als iframe in einer anderen Website**: erlaubt (siehe unten)

### Einbetten als iframe

Die Seiten senden kein `X-Frame-Options` und erlauben Framing per
`Content-Security-Policy: frame-ancestors *` — konfiguriert in der Datei
`_headers` im Repo-Root (Cloudflare Pages liest sie beim Deploy).

```html
<iframe src="https://emailimport.pages.dev/downloads.html?embed=1"
        style="width:100%;height:900px;border:0" loading="lazy"></iframe>
```

URL-Schalter für das eingebettete Layout:

| Parameter | Wirkung |
|-----------|---------|
| *(keiner)* | Framing wird automatisch erkannt: klebende Kopfzeilen werden statisch, Aussenränder kleiner |
| `?embed=1` | erzwingt das kompakte Einbett-Layout (auch ohne iframe, zum Testen) |
| `?embed=0` | schaltet das Einbett-Layout ab |
| `?chrome=0` | blendet zusätzlich die Kopfzeile der Seite aus (wenn die einbettende Seite eigene Navigation hat) |

Nur bestimmte Host-Seiten zulassen? In `_headers` das `*` bei `frame-ancestors`
durch die erlaubten Origins ersetzen. Hinweis: In einem iframe auf fremder Domain
können Browser `localStorage` blockieren — die API-Key-Speicherung in
`claude-analyse.html` funktioniert dann nur beim direkten Aufruf.

### Funktionen

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         WEB-APP FUNKTIONEN                                   │
└─────────────────────────────────────────────────────────────────────────────┘

  1. JSON LADEN
  ┌────────────────────────────────────────────┐
  │ • Drag & Drop der JSON-Datei               │
  │ • Oder Datei-Auswahl per Button            │
  │ • Automatische Cluster-Erkennung           │
  └────────────────────────────────────────────┘

  2. EMAILS DURCHSEHEN
  ┌────────────────────────────────────────────┐
  │ • Liste aller Vorgänge links               │
  │ • Detail-Ansicht rechts                    │
  │ • Anfrage + alle Antworten sichtbar        │
  │ • Schlagwort-Highlighting                  │
  └────────────────────────────────────────────┘

  3. KATEGORISIEREN
  ┌────────────────────────────────────────────┐
  │ Cluster:                                   │
  │ • KV (Krankenversicherung)                 │
  │ • SHUK (Sach/Haftpflicht/Unfall/KFZ)      │
  │ • LV (Lebensversicherung)                  │
  │ • Provisionierung                          │
  │ • Bestand                                  │
  │ • Absatzeinheiten                          │
  │                                            │
  │ Status: Neu / In Bearbeitung / Erledigt    │
  │ Bearbeiter: Dropdown-Auswahl               │
  └────────────────────────────────────────────┘

  4. EXPORTIEREN
  ┌────────────────────────────────────────────┐
  │ • Bearbeitete JSON speichern               │
  │ • Nur ausgewählte Emails exportieren       │
  └────────────────────────────────────────────┘
```

### Schlagwort-Erkennung

Die Web-App erkennt automatisch Cluster anhand von Schlagwörtern:

| Cluster | Erkannte Begriffe |
|---------|-------------------|
| **KV** | kv, kranken, pflege, zahnzusatz, ambulant, stationär... |
| **SHUK** | shuk, haftpflicht, hausrat, kfz, kasko, unfall, rechtsschutz... |
| **LV** | lv, leben, rente, riester, rürup, berufsunfähigkeit... |
| **Provisionierung** | provision, courtage, storno, vergütung... |
| **Bestand** | bestand, bestandskunde, bestandsvertrag... |
| **Absatzeinheiten** | ae, absatzeinheit... |

---

## Teil 3: Excel-Import (HTA Tool)

### Funktionsweise Import

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         HTA IMPORT PROZESS                                   │
└─────────────────────────────────────────────────────────────────────────────┘

  SCHRITT 1: Excel verbinden
  ┌────────────────────────────────────────────┐
  │ • Tracking_Hypercare.xlsm muss offen sein  │
  │ • Blatt "Uebersicht" wird verwendet        │
  └────────────────────────────────────────────┘
                      ↓
  SCHRITT 2: JSON laden
  ┌────────────────────────────────────────────┐
  │ • JSON-Datei auswählen                     │
  │ • Konversationen werden geladen            │
  └────────────────────────────────────────────┘
                      ↓
  SCHRITT 3: Import durchführen
  ┌────────────────────────────────────────────┐
  │ Für jede Konversation:                     │
  │                                            │
  │ 1. Suche in Excel nach ConversationID      │
  │    oder Datum+Betreff                      │
  │                                            │
  │ 2a. GEFUNDEN:                              │
  │     → Prüfe auf neue Antworten (ReplyID)   │
  │     → Füge nur NEUE Antworten hinzu        │
  │                                            │
  │ 2b. NICHT GEFUNDEN:                        │
  │     → sentOnly=true? → Überspringen        │
  │     → sonst: Neue Zeile anlegen            │
  └────────────────────────────────────────────┘
```

### Duplikat-Erkennung (ReplyID)

Jede Antwort hat eine eindeutige **ReplyID** (= Outlook EntryID):

```
┌─────────────────────────────────────────────────────────────────────────────┐
│ BEISPIEL: Zweiter Import der gleichen Konversation                          │
└─────────────────────────────────────────────────────────────────────────────┘

  EXCEL (nach erstem Import):
  ┌──────────────────────────────────────────────────────────────────────────┐
  │ Zeile 5:                                                                  │
  │ Spalte 14 (Antworten): "=== 02.12. [Von: Ich] === Antwort 1..."          │
  │ Spalte 24 (ReplyIDs):  "AAA111,BBB222"                                   │
  └──────────────────────────────────────────────────────────────────────────┘

  JSON (neuer Export):
  ┌──────────────────────────────────────────────────────────────────────────┐
  │ Konversation mit 3 Antworten:                                            │
  │ • Antwort 1: replyId="AAA111"  → bereits in Excel ✗                      │
  │ • Antwort 2: replyId="BBB222"  → bereits in Excel ✗                      │
  │ • Antwort 3: replyId="CCC333"  → NEU! ✓                                  │
  └──────────────────────────────────────────────────────────────────────────┘

  ERGEBNIS:
  ┌──────────────────────────────────────────────────────────────────────────┐
  │ Zeile 5 (aktualisiert):                                                  │
  │ Spalte 14: "=== 05.12. [Von: Ich] === Antwort 3..." + alte Antworten    │
  │ Spalte 24: "AAA111,BBB222,CCC333"                                        │
  └──────────────────────────────────────────────────────────────────────────┘
```

---

## Typischer Arbeitsablauf

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                      TÄGLICHER WORKFLOW                                      │
└─────────────────────────────────────────────────────────────────────────────┘

  MORGENS:
  ────────
  1. Excel öffnen (Tracking_Hypercare.xlsm)
  2. HTA Tool starten (outlook-export-modular.hta)
  3. Export durchführen (letzte 7 Tage)
     → JSON wird in Downloads gespeichert

  BEARBEITUNG:
  ────────────
  4. Web-App öffnen (emailimport.pages.dev)
  5. JSON per Drag & Drop laden
  6. Neue Vorgänge durchsehen und kategorisieren:
     • Cluster zuweisen
     • Status setzen
     • Bearbeiter zuweisen
     • Kommentare hinzufügen
  7. Bearbeitete JSON speichern

  ABENDS:
  ───────
  8. HTA Tool → Import-Tab
  9. Bearbeitete JSON laden
  10. Import durchführen
      → Excel wird aktualisiert

  ┌────────────────────────────────────────────┐
  │ Ergebnis:                                  │
  │ • X neue Vorgänge importiert              │
  │ • Y Antworten aktualisiert                │
  │ • Z Duplikate übersprungen                │
  └────────────────────────────────────────────┘
```

---

## Projektstruktur

```
emailimport/
├── index.html                    # Web-App (Cloudflare Pages)
├── outlook-export-modular.hta    # HTA Export/Import Tool
├── Email-Export-Tool.zip         # Download-Paket für HTA
│
├── css/
│   └── main.css                  # Stylesheet für HTA
│
├── js/
│   ├── config.js                 # Globale Konfiguration
│   │
│   ├── utils/                    # Hilfsfunktionen
│   │   ├── json-polyfill.js      # JSON für IE/HTA
│   │   ├── date-utils.js         # Datumsformatierung
│   │   ├── text-utils.js         # Text-Bereinigung, Encoding
│   │   └── html-parser.js        # HTML-Extraktion
│   │
│   ├── outlook/                  # Outlook-Zugriff
│   │   ├── init.js               # Initialisierung
│   │   ├── mapi.js               # MAPI-Properties
│   │   └── extract.js            # Email-Extraktion
│   │
│   ├── conversation/             # Konversations-Logik
│   │   ├── builder.js            # Konversations-Aufbau
│   │   └── extractor.js          # ConversationIndex-Parsing
│   │
│   ├── export/                   # Export-Funktionen
│   │   ├── core.js               # Export-Hauptlogik
│   │   ├── sent-cache.js         # Sent-Items-Cache
│   │   └── file-save.js          # Datei-Speicherung
│   │
│   └── import/                   # Import-Funktionen
│       ├── core.js               # Import-Hauptlogik
│       ├── excel-connect.js      # Excel-Verbindung
│       └── excel-write.js        # Excel-Schreibfunktionen
│
└── README.md                     # Diese Datei
```

---

## JSON-Format

### Konversations-Format (Export)

```json
{
  "exportDate": "2025-12-15T10:00:00.000Z",
  "mailboxName": "Hypercare Postfach",
  "totalEmails": 150,
  "conversationCount": 45,
  "conversations": {
    "ConvID-123": {
      "conversationID": "ConvID-123",
      "subject": "Provisionsreklamation",
      "messageCount": 3,
      "firstMessageDate": "2025-12-10T08:00:00.000Z",
      "lastMessageDate": "2025-12-12T14:30:00.000Z",
      "messages": [
        {
          "entryID": "Email-ABC",
          "folder": "inbox",
          "depth": 0,
          "subject": "Provisionsreklamation",
          "senderEmail": "kunde@example.de",
          "senderName": "Max Kunde",
          "receivedTime": "2025-12-10T08:00:00.000Z",
          "body": "Nachrichtentext..."
        },
        {
          "entryID": "Email-DEF",
          "folder": "sent",
          "depth": 1,
          "subject": "AW: Provisionsreklamation",
          "sentOn": "2025-12-10T10:30:00.000Z",
          "body": "Antwort..."
        }
      ]
    }
  }
}
```

---

## Datenschutz

- **Alle Daten bleiben lokal** - Keine Übertragung an externe Server
- **Web-App läuft im Browser** - Kein Backend, keine Datenbank
- **JSON-Dateien** - Werden nur lokal gespeichert (Downloads-Ordner)

---

## Fehlerbehebung

### HTA startet nicht
- Windows-Sicherheitswarnung bestätigen
- Rechtsklick → "Eigenschaften" → "Zulassen" aktivieren

### Outlook-Verbindung fehlgeschlagen
- Outlook muss gestartet sein
- Postfach muss eingerichtet sein
- Als Administrator ausführen probieren

### Excel wird nicht erkannt
- Excel-Datei muss geöffnet sein
- Dateiname muss "Tracking" oder "Hypercare" enthalten

### Import-Fehler
- Excel-Datei nicht schreibgeschützt?
- Tabelle "Uebersicht" vorhanden?
- JSON-Format korrekt?

---

---

## Teil 4: BGAV Testmail-Extraktion + SST-Workflow-Upload

Neben dem Hypercare-Review-Workflow gibt es einen separaten Prozess fuer die **automatisierte Ablage von BGAV-Titelbezeichnung-Mails** in der SST-Workflow-API.

### Gesamtablauf

```
┌─────────────────────────────────────────────────────────────────────────────┐
│              BGAV TESTMAIL-EXTRAKTION + UPLOAD WORKFLOW                      │
└─────────────────────────────────────────────────────────────────────────────┘

  ┌──────────────┐         ┌──────────────────┐         ┌──────────────────┐
  │  EML-Ordner  │         │  HTA-Tool        │         │  Batch-Upload    │
  │  (Dateien)   │         │  (Verarbeitung)  │         │  (Java)          │
  └──────┬───────┘         └──────┬───────────┘         └──────┬───────────┘
         │                        │                             │
         │  1. EML-Dateien        │                             │
         │     einlesen           │                             │
         │  ──────────────►       │                             │
         │                        │                             │
         │                        │  2. Zuordnen + Umbenennen   │
         │                        │     + PDF + CSV             │
         │                        │  ──────────────────►        │
         │                        │                             │
         │                        │                             │  3. SST-API
         │  (nicht zugeordnete    │                             │  ──────────►
         │   bleiben im Ordner)   │                             │  Upload
         └────────────────────────┴─────────────────────────────┘
```

### Schritt 1: HTA-Tool (bgav-testmail-extraktion.hta)

Das HTA-Tool verarbeitet EML-Dateien aus einem Ordner (kein Outlook noetig):

1. **EML-Ordner auswaehlen** - Ordner mit den .eml Dateien angeben
2. **Excel-Metadaten laden** - Liest BD-Nummern, Namen, E-Mails aus `Uebersicht_Titelaenderung.xlsx`
3. **EML-Header parsen** - Liest Empfaenger (To) aus jeder EML-Datei
4. **BD-Nummer zuordnen** - Empfaenger-E-Mail wird gegen Excel gematcht
5. **Zugeordnete Dateien umbenennen** - EML wird nach `{BD}_BGAV_Titelbezeichnung_{nr}.eml` umbenannt und in den Ausgabe-Ordner verschoben
6. **PDF erstellen** - Mail als PDF via Word-Konvertierung (optional)
7. **CSV erzeugen** - Metadaten-Datei fuer den Batch-Upload
8. **Nicht zugeordnete Dateien** bleiben im Quell-Ordner fuer manuelle Pruefung und koennen beim naechsten Durchlauf erneut verarbeitet werden

**Ausgabe-Struktur:**
```
ausgabe/
├── eml/
│   ├── 00010091_BGAV_Titelbezeichnung_001.eml
│   ├── 00550179_BGAV_Titelbezeichnung_002.eml
│   └── ...
├── pdf/
│   ├── 00010091_BGAV_Titelbezeichnung_001.pdf
│   ├── 00550179_BGAV_Titelbezeichnung_002.pdf
│   └── ...
└── bgav_metadaten.csv

quell-ordner/
├── nicht_zugeordnet_1.eml   (bleibt fuer naechsten Durchlauf)
└── nicht_zugeordnet_2.eml
```

**CSV-Format:**
```
Nr;Dateiname_EML;Dateiname_PDF;Empfaenger_Email;Vorname;Nachname;BD_Nummer;Klammerbegriff;Original_Dateiname
1;00010091_BGAV_Titelbezeichnung_001.eml;00010091_BGAV_Titelbezeichnung_001.pdf;name@firma.de;Max;Muster;00010091;ADM Vertrag;original.eml
```

### Schritt 2: Batch-Upload (Java)

Das Java-Tool `BgavBatchUpload` liest die CSV und laedt die Dateien ueber die SST-Workflow-API hoch.

**Verwendung:**
```bash
java -jar bgav-batch-upload.jar <pfad-zum-output-ordner>

# Dry-Run (nur pruefen, kein Upload):
java -jar bgav-batch-upload.jar <pfad> --dry-run
```

**Was passiert pro Testfall:**
```
1. Originaldokument archivieren    → EML wird ins SST-Archiv gespeichert
2. Arbeitsdokument erzeugen        → PDF wird als Arbeitskopie angelegt + Vorgang erstellt
3. Vorgang abschliessen            → Vorgang wird geschlossen
4. Dokument klammern               → Klammerbegriff "ADM Vertrag" wird gesetzt
```

### Aktuelle Konfiguration (WorkflowAdapter)

| Parameter | Wert | Beschreibung |
|-----------|------|-------------|
| **Klammer** | `ADM Vertrag` | Klammerbegriff am Dokument |
| **Dokumentenart** | VSW (Schriftwechsel) | TODO: numerische ID noetig |
| **Vorgangsart** | Sonstige | TODO: numerische ID noetig |
| **Hinweis** | `Titel ab 01.01.26` | Hinweistext am Dokument |
| **Fremdschluesselsystem** | `EV_Technischer_Vertragsnachtrag` | Externe System-Kennung |
| **Technischer User** | `TBD` | Erstellender User (PersNr) |
| **Geschnotyp** | `BD_VERMITTLER_VERTRAG` | Aktentyp |

### Offene TODOs

- [ ] Numerische ID fuer Vorgangsart "Sonstige" (`VORGANGS_TYP`)
- [ ] Numerische ID fuer Dokumentenart "VSW" (`DOKUMENT_TYP`)
- [ ] Echte PersNr fuer den technischen User (`ERSTELLER`)

### Weitere Dokumentation

- **[SST-WORKFLOW-API.md](SST-WORKFLOW-API.md)** - Referenz-Code und API-Details der SST-Workflow-Schnittstelle
- **[HTA-ANLEITUNG.md](HTA-ANLEITUNG.md)** - Allgemeine HTA-Anleitung

---

## Lizenz

Internes Tool fuer Barmenia/Gothaer BGAV Hypercare Projekt.
