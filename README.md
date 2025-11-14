# BGAV Hypercare - Email Review Tool

Tool zur Überprüfung und Kategorisierung von Hypercare-Emails für das Barmenia/Gothaer-Projekt.

## 🎯 Zwei Versionen verfügbar

### 🚀 HTA-Version (EMPFOHLEN für Windows)

**`bgav-hypercare-standalone.hta`** - Die All-in-One-Lösung!

✅ **Direkter Outlook-Zugriff** - Keine Scripts, keine Exports
✅ **Direkter Excel-Export** - Öffnet sich automatisch in Excel
✅ **One-Click-Workflow** - Emails laden → Review → Export
✅ **Komplett offline** - Keine Internetverbindung nötig

**Quick Start:**
1. Doppelklick auf `bgav-hypercare-standalone.hta`
2. "📥 Ausgewählte Emails laden" oder "📂 Ordner laden"
3. Review & Kategorisierung
4. "💾 Nach Excel exportieren"
5. Fertig!

👉 **[Vollständige Anleitung](HTA-ANLEITUNG.md)**

---

### 🌐 HTML-Version (Plattformübergreifend)

**`bgav-hypercare-email-review.html`** - Läuft in jedem Browser

- Für Mac, Linux, Windows
- Benötigt JSON-Export aus Outlook (via PowerShell/VBA)
- XLSX-Download statt direktem Excel-Export

**Quick Start:**
1. Öffne `bgav-hypercare-email-review.html` im Browser
2. JSON-Datei hochladen (siehe Outlook-Integration)
3. Review & Kategorisierung
4. Excel-Datei wird heruntergeladen

---

## 📋 Features

- **Intelligente Kategorisierung**: Automatische Erkennung von:
  - Kategorie (Incident / Fachliche Rückfrage)
  - Cluster (SHUK, LV, KV, Provisionierung, Produktzuordnung, Allgemein)
  - Agentur
- **Filterung & Suche**: Durchsuche und filtere Emails nach verschiedenen Kriterien
- **Batch-Bearbeitung**: Bearbeite mehrere Emails gleichzeitig
- **Excel-Export**: Exportiere im Hypercare-Template-Format

## 🚀 Verwendung

### HTA-Version (Windows - EMPFOHLEN)

1. Doppelklick auf `bgav-hypercare-standalone.hta`
2. Emails direkt aus Outlook laden
3. Review & Export - Fertig!

Siehe **[HTA-ANLEITUNG.md](HTA-ANLEITUNG.md)** für Details.

### HTML-Version (Alle Plattformen)

1. Öffne `bgav-hypercare-email-review.html` im Browser
2. Lade eine JSON-Datei hoch
3. Die App funktioniert komplett offline

### 📧 Outlook-Integration

**NEU:** Du kannst Emails jetzt direkt aus Outlook exportieren!

Siehe **[OUTLOOK-INTEGRATION.md](OUTLOOK-INTEGRATION.md)** für eine detaillierte Anleitung zu:
- **PowerShell-Script** (empfohlen) - Einfacher Export ausgewählter Emails oder ganzer Ordner
- **VBA-Makro** - Direkte Integration in Outlook mit Button im Ribbon

Beide Methoden exportieren Emails im korrekten JSON-Format für diese App.

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

### 🧪 Testen mit Beispiel-Daten

Möchtest du die App erst einmal ausprobieren? Nutze die Datei **`beispiel-emails.json`**:
- Enthält 10 realistische Test-Emails
- Verschiedene Kategorien (Incidents, Rückfragen)
- Verschiedene Cluster (SHUK, LV, KV, Provisionierung, etc.)
- Perfekt zum Testen aller Features der App

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
