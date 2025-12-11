# BGAV Hypercare - Email Review Tool

Tool zur Überprüfung und Kategorisierung von Hypercare-Emails für das Barmenia/Gothaer-Projekt.

## Projektstruktur

```
emailimport/
├── index.html                    # Web-App für Cloudflare Pages
├── outlook-export.hta            # Monolithische HTA (Legacy)
├── outlook-export-modular.hta    # Modulare HTA (Neu)
│
├── css/
│   └── main.css                  # Haupt-Stylesheet
│
├── js/
│   ├── config.js                 # Globale Konfiguration und State
│   │
│   ├── utils/                    # Utility-Funktionen
│   │   ├── json-polyfill.js      # JSON für IE/HTA
│   │   ├── date-utils.js         # Datumsformatierung und -parsing
│   │   ├── text-utils.js         # Text-Bereinigung und Encoding
│   │   └── html-parser.js        # HTML-Extraktion für Emails
│   │
│   ├── ui/                       # UI-Komponenten
│   │   ├── step-indicator.js     # Workflow-Schrittanzeige
│   │   ├── tabs.js               # Tab-Wechsel
│   │   └── progress.js           # Fortschrittsanzeigen
│   │
│   ├── outlook/                  # Outlook-Integration
│   │   ├── init.js               # Outlook-Initialisierung
│   │   ├── mapi.js               # MAPI-Property-Extraktion
│   │   └── extract.js            # Email-Daten-Extraktion
│   │
│   ├── threads/                  # Thread-Verarbeitung
│   │   ├── grouping.js           # Konversations-Gruppierung
│   │   ├── depth.js              # Thread-Tiefenberechnung
│   │   └── matching.js           # Reply-Matching
│   │
│   ├── export/                   # Export-Funktionalität
│   │   ├── core.js               # Export-Hauptlogik
│   │   ├── sent-cache.js         # Gesendete-Items-Cache
│   │   └── file-save.js          # JSON-Speicherung
│   │
│   └── import/                   # Import-Funktionalität
│       ├── core.js               # Import-Hauptlogik
│       ├── excel-connect.js      # Excel-Verbindung
│       └── excel-write.js        # Excel-Schreibfunktionen
│
└── Export-OutlookEmails.ps1      # PowerShell Export-Script
```

## Versionen

### Web-Version (Cloudflare Pages)

**`index.html`** - Moderne Web-App mit Dark Mode Design

- Modernes Dark-Mode UI mit Animationen
- Funktioniert in jedem Browser
- Kann auf Cloudflare Pages gehostet werden
- JSON-Import per Drag & Drop
- CSV-Export für bearbeitete Emails

### HTA-Version (Windows-Desktop)

Zwei Varianten verfügbar:

#### 1. Modulare Version (Empfohlen)
**`outlook-export-modular.hta`** - Neue Version mit modularem Code

- Lädt JavaScript-Module zur Laufzeit
- Bessere Wartbarkeit und Übersichtlichkeit
- Einfacher zu debuggen
- Benötigt alle Dateien im `js/` Ordner

#### 2. Monolithische Version (Legacy)
**`outlook-export.hta`** - Alles in einer Datei (~4800 Zeilen)

- Einzelne Datei, keine Abhängigkeiten
- Einfacher zu verteilen
- Schwerer zu warten

### Features beider HTA-Versionen

- **Direkter Outlook-Zugriff** - Liest Emails direkt aus Outlook
- **Excel-Duplikaterkennung** - Prüft bestehende Einträge
- **Thread-Erkennung** - Gruppiert Konversationen automatisch
- **Antwort-Matching** - Findet zugehörige gesendete Emails
- **Auto-Erledigt** - Markiert gelöste Vorgänge automatisch

## Workflow

1. **Excel verbinden** - Tracking_Hypercare.xlsm öffnen und verbinden
2. **Postfach auswählen** - Mailbox und Ordner wählen
3. **Emails exportieren** - JSON-Datei wird erstellt
4. **Web-App** - Emails kategorisieren und bearbeiten
5. **Excel Import** - Bearbeitete Daten importieren

## Module-Dokumentation

### js/utils/ - Utilities

| Modul | Funktionen |
|-------|-----------|
| `json-polyfill.js` | `JSON.stringify()`, `JSON.parse()` für IE |
| `date-utils.js` | `formatDate()`, `formatDateKey()`, `extractOldestDateFromQuotes()`, `extractDateFromHtmlHeader()` |
| `text-utils.js` | `sanitizeText()`, `fixEncoding()`, `normalizeSubject()`, `removeEmailQuotes()`, `htmlToPlainText()` |
| `html-parser.js` | `extractNewContentFromHtml()` - Entfernt zitierte Texte aus HTML |

### js/outlook/ - Outlook-Integration

| Modul | Funktionen |
|-------|-----------|
| `init.js` | `initOutlook()`, `loadMailboxes()`, `loadFolders()` |
| `mapi.js` | `getEmailHeaders()` - Extrahiert Message-ID, In-Reply-To, References |
| `extract.js` | `extractEmail()`, `extractEmailUnified()` |

### js/threads/ - Thread-Verarbeitung

| Modul | Funktionen |
|-------|-----------|
| `grouping.js` | `groupByConversation()`, `removeDuplicateEmails()`, `assignParentRelationships()` |
| `depth.js` | `calculateThreadDepthsUnified()`, `assignThreadPositionsUnified()`, `processConversationUnified()` |
| `matching.js` | `findRepliesMultiLayer()`, `filterNewReplies()`, `createReplyObject()` |

### js/export/ - Export

| Modul | Funktionen |
|-------|-----------|
| `core.js` | `startExport()`, `doExport()`, `processEmailBatch()`, `finishExport()` |
| `sent-cache.js` | `cacheSentItems()`, `processSentBatch()` |
| `file-save.js` | `saveExportFile()`, `openWebApp()`, `openDownloads()` |

### js/import/ - Import

| Modul | Funktionen |
|-------|-----------|
| `core.js` | `startImport()`, `doImport()`, `loadJsonFile()` |
| `excel-connect.js` | `detectExcelForExport()`, `detectExcel()`, `checkImportReady()` |
| `excel-write.js` | `writeEmailRow()`, `formatReplies()`, `boldTimestamps()` |

## JSON-Format

```json
[
  {
    "emailId": "EntryID",
    "conversationId": "ConversationID",
    "internetMessageId": "<message-id@domain>",
    "datum": "2025-12-09T14:16:36",
    "conversationStartDate": "2025-12-09T14:16:36",
    "lastActivityDate": "2025-12-11T10:30:00",
    "von_email": "max.mueller@example.de",
    "von_name": "Max Müller",
    "betreff": "Anfrage zur Provisionsabrechnung",
    "text": "Nur der neue Inhalt (ohne Zitate)",
    "threadPosition": 1,
    "threadDepth": 0,
    "isThreadRoot": true,
    "messageCount": 5,
    "antworten": [
      {
        "datum": "2025-12-09T15:30:00",
        "von": "Support Team",
        "text": "Antwort-Text",
        "threadPosition": 2,
        "threadDepth": 1,
        "isIncoming": false,
        "replyId": "EntryID"
      }
    ]
  }
]
```

## Cloudflare Pages Deployment

### Via GitHub

1. Repository zu GitHub pushen
2. Cloudflare Dashboard > Pages > Create a project
3. "Connect to Git" > Repository auswählen
4. Settings:
   - Build command: (leer)
   - Build output directory: `/`
5. Deploy!

### Direct Upload

1. Cloudflare Dashboard > Pages > Create a project
2. "Upload assets" auswählen
3. `index.html` hochladen
4. Deploy!

## Datenschutz

Alle Daten bleiben lokal. Es werden keine Daten an externe Server übertragen.
Die Web-App läuft komplett im Browser.

## Lizenz

Internes Tool für Barmenia/Gothaer BGAV Hypercare Projekt.
