# BGAV Hypercare - Email Review Tool

Eine standalone HTML-Anwendung zur Überprüfung und Kategorisierung von Hypercare-Emails für das Barmenia/Gothaer-Projekt.

## 📋 Features

- **Email-Import**: Lade Emails als JSON-Datei hoch
- **Intelligente Kategorisierung**: Automatische Erkennung von:
  - Kategorie (Incident / Fachliche Rückfrage)
  - Cluster (SHUK, LV, KV, Provisionierung, Produktzuordnung, Allgemein)
  - Agentur
- **Filterung & Suche**: Durchsuche und filtere Emails nach verschiedenen Kriterien
- **Batch-Bearbeitung**: Bearbeite mehrere Emails gleichzeitig
- **Excel-Export**: Exportiere ausgewählte Emails in eine Excel-Datei für weitere Verarbeitung

## 🚀 Verwendung

### Lokale Ausführung

1. Lade die Datei `bgav-hypercare-email-review.html` herunter
2. Öffne die Datei in einem modernen Webbrowser (Chrome, Firefox, Edge, Safari)
3. Die App funktioniert komplett offline - keine Server-Verbindung erforderlich!

### Email-Daten importieren

Die App erwartet JSON-Dateien mit folgendem Format:

```json
[
  {
    "datum": "2025-11-14T10:30:00",
    "von_email": "max.mueller@agentur.de",
    "von_name": "Max Müller",
    "betreff": "Frage zur SHUK Provision",
    "text": "Email Text hier...",
    "anhänge": ["dokument.pdf"]
  }
]
```

### Workflow

1. **Emails laden**: Klicke auf "📤 Emails laden" und wähle deine JSON-Datei
2. **Überprüfen**: Die App kategorisiert die Emails automatisch
3. **Anpassen**: Korrigiere bei Bedarf die automatische Kategorisierung
4. **Exportieren**: Exportiere die kategorisierten Emails nach Excel

## 🛠️ Technische Details

- **Standalone**: Alle CSS und JavaScript sind inline - keine externen Abhängigkeiten außer der XLSX-Bibliothek (via CDN)
- **Framework**: Vanilla JavaScript (kein Framework erforderlich)
- **Browser-Kompatibilität**: Alle modernen Browser
- **Offline-fähig**: Funktioniert ohne Internetverbindung (Excel-Export benötigt einmalig CDN-Zugriff)

## 📦 Excel-Export-Format

Der Export erstellt eine Excel-Datei mit folgenden Spalten:

- ID (HC-001, HC-002, ...)
- Datum
- Kanal (immer "Email")
- Agentur
- Kategorie
- Status
- In Bearbeitung von
- Cluster
- Anfrage
- Diagnose
- Nächste Schritte
- Lösung
- JIRA
- Kommunikation
- Kommentar

## 🔒 Datenschutz

Alle Daten bleiben lokal im Browser. Es werden keine Daten an externe Server übertragen.

## 📝 Lizenz

Internes Tool für Barmenia/Gothaer BGAV Hypercare Projekt.
