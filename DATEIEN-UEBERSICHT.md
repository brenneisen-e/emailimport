# Dateien-Übersicht

Dieses Repository enthält alle Dateien für das BGAV Hypercare Email Review Tool.

## 📁 Hauptdateien

### `bgav-hypercare-standalone.hta` ⭐ EMPFOHLEN
**Die All-in-One Windows-Anwendung** - Direkter Outlook & Excel Zugriff

- **Was macht sie?** Kompletter Workflow: Outlook → Review → Excel
- **Wie verwenden?** Doppelklick (läuft als Windows-App)
- **Abhängigkeiten:** Windows, Outlook, Excel
- **Offline-fähig:** Ja
- **Größe:** ~75 KB
- **Plattform:** Nur Windows

**Features:**
- ✅ **Direkter Outlook-Zugriff** (keine separate Export-Datei nötig)
- ✅ **Direkter Excel-Export** (öffnet sich automatisch)
- Automatische Kategorisierung (Incident/Rückfrage, Cluster)
- Ordner-Browser für Outlook-Ordner
- Such- und Filterfunktionen
- Batch-Bearbeitung
- Läuft als native Windows-Anwendung

**👉 Siehe [HTA-ANLEITUNG.md](HTA-ANLEITUNG.md) für Details**

---

### `bgav-hypercare-email-review.html`
**Browser-basierte Anwendung** - Plattformübergreifend

- **Was macht sie?** Email-Review-Tool mit automatischer Kategorisierung und Excel-Export
- **Wie verwenden?** Im Browser öffnen (Chrome, Firefox, Edge, Safari)
- **Abhängigkeiten:** Moderner Browser, XLSX-Library via CDN
- **Offline-fähig:** Ja (nach erstem Laden)
- **Größe:** ~49 KB
- **Plattform:** Alle (Windows, Mac, Linux)

**Features:**
- JSON-Import von Emails
- Automatische Kategorisierung (Incident/Rückfrage, Cluster)
- Such- und Filterfunktionen
- Batch-Bearbeitung
- Excel-Export als Download (XLSX-Datei)

---

## 🔧 Outlook-Export-Tools

### `Export-OutlookEmails.ps1`
**PowerShell-Script für Outlook-Export** (EMPFOHLEN)

- **Was macht es?** Exportiert ausgewählte Outlook-Emails oder ganze Ordner als JSON
- **Wie verwenden?**
  ```powershell
  .\Export-OutlookEmails.ps1
  ```
- **Voraussetzungen:** Windows PowerShell (vorinstalliert), Outlook
- **Parameter:**
  - `-OutputPath` - Direkter Export zu Datei
  - `-FolderPath` - Export ganzer Outlook-Ordner

**Vorteile:**
- ✅ Keine Installation erforderlich
- ✅ Sehr gute Performance
- ✅ Detaillierte Fehlerbehandlung
- ✅ Funktioniert mit großen Datenmengen (>10.000 Emails)

**Verwendungsbeispiele:**
```powershell
# Ausgewählte Emails exportieren
.\Export-OutlookEmails.ps1

# Ganzen Ordner exportieren
.\Export-OutlookEmails.ps1 -FolderPath "Posteingang\Hypercare"

# Direkt zu Datei exportieren
.\Export-OutlookEmails.ps1 -OutputPath "C:\Temp\emails.json"
```

---

### `OutlookExportMacro.vba`
**VBA-Makro für direkte Outlook-Integration**

- **Was macht es?** Exportiert Emails direkt aus Outlook heraus
- **Wie verwenden?** In Outlook-VBA-Editor importieren (ALT+F11)
- **Voraussetzungen:** Outlook mit aktivierten Makros

**Verfügbare Makros:**
1. `ExportSelectedEmailsToJSON` - Exportiert ausgewählte Emails
2. `ExportFolderToJSON` - Exportiert aktuellen Ordner komplett

**Vorteile:**
- ✅ Direkt in Outlook integriert
- ✅ Kann als Button im Ribbon hinzugefügt werden
- ✅ Schneller Zugriff mit ALT+F8

**Installation:**
1. Outlook öffnen → ALT+F11
2. Einfügen → Modul
3. Code aus VBA-Datei einfügen
4. Speichern

---

## 📖 Dokumentation

### `README.md`
**Hauptdokumentation**

- Überblick über das Projekt
- Features und Verwendung
- Technische Details
- Schnellstart-Anleitung

### `OUTLOOK-INTEGRATION.md`
**Detaillierte Anleitung zur Outlook-Integration**

- Schritt-für-Schritt-Anleitungen für beide Export-Methoden
- Troubleshooting
- Best Practices
- Vergleich PowerShell vs. VBA
- Häufige Fehler und Lösungen

### `DATEIEN-UEBERSICHT.md` (diese Datei)
**Übersicht aller Dateien im Repository**

---

## 🧪 Test-Daten

### `beispiel-emails.json`
**Beispiel-Datendatei zum Testen**

- **Was enthält sie?** 10 realistische Test-Emails
- **Wie verwenden?** In die HTML-App laden zum Testen
- **Inhalt:**
  - Verschiedene Kategorien (Incidents, Rückfragen)
  - Verschiedene Cluster (SHUK, LV, KV, Provisionierung, Produktzuordnung, Allgemein)
  - Mit und ohne Anhänge
  - Realistische Szenarien aus dem Hypercare-Alltag

**Perfekt um:**
- Die App auszuprobieren ohne echte Daten
- Features zu testen
- Automatische Kategorisierung zu verstehen
- Excel-Export zu prüfen

---

## 🗂️ Dateistruktur

```
emailimport/
│
├── bgav-hypercare-standalone.hta        # ⭐ HTA-App (Windows, empfohlen)
├── bgav-hypercare-email-review.html    # Browser-App (alle Plattformen)
│
├── Export-OutlookEmails.ps1             # PowerShell Export-Script
├── OutlookExportMacro.vba               # VBA Makro für Outlook
│
├── beispiel-emails.json                 # Test-Daten
│
├── README.md                            # Haupt-Dokumentation
├── HTA-ANLEITUNG.md                     # HTA-Version Anleitung
├── OUTLOOK-INTEGRATION.md               # Outlook-Anleitung (HTML-Version)
└── DATEIEN-UEBERSICHT.md               # Diese Datei
```

---

## 🚀 Quick Start

### ⭐ Empfohlen: HTA-Version (Windows)

**Schnellster Weg von Outlook zu Excel:**

1. **App starten:**
   ```
   Doppelklick auf: bgav-hypercare-standalone.hta
   ```

2. **Emails laden:**
   - Wähle in Outlook Emails aus
   - In HTA-App: "📥 Ausgewählte Emails laden"
   - Oder: "📂 Ordner laden" für ganzen Ordner

3. **Review & Export:**
   - Kategorien prüfen/anpassen
   - "💾 Nach Excel exportieren"
   - Fertig! Excel öffnet sich automatisch

**Siehe [HTA-ANLEITUNG.md](HTA-ANLEITUNG.md) für Details**

---

### Alternative: HTML-Version (Alle Plattformen)

**Mit JSON-Export:**

1. **Outlook-Export einrichten:**
   - Methode wählen (PowerShell oder VBA)
   - Siehe `OUTLOOK-INTEGRATION.md` für Details

2. **Emails aus Outlook exportieren:**
   ```powershell
   .\Export-OutlookEmails.ps1
   ```

3. **In App laden und verarbeiten:**
   - `bgav-hypercare-email-review.html` im Browser öffnen
   - Exportierte JSON-Datei laden
   - Review durchführen
   - Nach Excel exportieren (Download)

---

### Für Erstnutzer (Testing)

1. **Browser-App öffnen:**
   ```
   Doppelklick auf: bgav-hypercare-email-review.html
   ```

2. **Test-Daten laden:**
   - In der App: "📤 Emails laden" klicken
   - Datei auswählen: `beispiel-emails.json`

3. **Ausprobieren:**
   - Emails durchsehen
   - Filter testen
   - Kategorien ändern
   - Excel exportieren

---

## 💾 Workflow-Diagramme

### HTA-Version (Empfohlen)

```
┌─────────────────┐
│     Outlook     │
│  (Emails/Ordner)│
└────────┬────────┘
         │
         ▼
┌─────────────────────────────────┐
│    HTA-App (One-Click)          │
│                                 │
│  1. Emails direkt laden         │
│  2. Auto-Kategorisierung        │
│  3. Review & Anpassung          │
│  4. Batch-Bearbeitung           │
└────────┬────────────────────────┘
         │
         ▼
┌─────────────────┐
│  Excel (Auto)   │
│  Tracking_HC    │
│                 │
│  Speichern &    │
│  Fertig! ✅     │
└─────────────────┘
```

**Vorteile:** Direkter Workflow, keine Zwischenschritte, alles in einer App!

---

### HTML-Version (Plattformübergreifend)

```
┌─────────────┐
│   Outlook   │
│   Emails    │
└──────┬──────┘
       │
       ├─────────────────────────────────────┐
       │                                     │
       ▼                                     ▼
┌─────────────┐                    ┌─────────────┐
│  PowerShell │                    │  VBA-Makro  │
│   Export    │                    │   Export    │
└──────┬──────┘                    └──────┬──────┘
       │                                  │
       └────────────┬─────────────────────┘
                    ▼
            ┌───────────────┐
            │  JSON-Datei   │
            └───────┬───────┘
                    ▼
        ┌───────────────────────┐
        │   HTML Review-App     │
        │   (im Browser)        │
        │                       │
        │ - Kategorisierung     │
        │ - Filterung           │
        │ - Batch-Bearbeitung   │
        └───────┬───────────────┘
                ▼
        ┌───────────────┐
        │ Excel-Download│
        │  (XLSX-Datei) │
        └───────────────┘
```

**Vorteile:** Plattformübergreifend, mehr Kontrolle über Export-Prozess

---

## 📊 Dateigrößen & Performance

| Datei | Größe | Ladezeit | Performance-Hinweis |
|-------|-------|----------|---------------------|
| `bgav-hypercare-standalone.hta` | ~75 KB | < 1s | ⭐ Sehr schnell, native Windows-App |
| `bgav-hypercare-email-review.html` | ~49 KB | < 1s | Auch mit 10.000 Emails flüssig |
| `Export-OutlookEmails.ps1` | ~9 KB | - | Sehr schnell, auch große Exports |
| `OutlookExportMacro.vba` | ~12 KB | - | OK für <1.000 Emails |
| `beispiel-emails.json` | ~6 KB | < 1s | Instant |

**Performance-Empfehlungen:**
- **HTA-App**: Ideal für 100-5.000 Emails, direkter Outlook-Zugriff
- **PowerShell-Script**: Ideal für sehr große Mengen (>5.000 Emails)
- **VBA-Makro**: Gut für kleine/mittlere Mengen (<1.000 Emails)
- **HTML-App**: Performant bis ~5.000 Emails im Browser

**Outlook-Zugriff:**
- **HTA**: Direkt, kein Export nötig
- **PowerShell/VBA**: Separater Export-Schritt erforderlich

---

## 🔄 Versions-Historie

### Version 2.0 (2025-11-14)
- ✨ **NEU: HTA-Version** - Direkter Outlook & Excel Zugriff
- ✅ All-in-One Windows-Anwendung
- ✅ Ordner-Browser für Outlook
- ✅ Direkter Excel-Export (öffnet sich automatisch)
- ✅ Umfassende HTA-Anleitung
- ✅ Aktualisierte Dokumentation

### Version 1.0 (2025-11-14)
- ✨ Initiale Version
- ✅ Standalone HTML-App (Browser)
- ✅ PowerShell Export-Script
- ✅ VBA-Makro
- ✅ Vollständige Dokumentation
- ✅ Beispiel-Daten

---

## 🆘 Support & Hilfe

**Bei Problemen, siehe:**
1. `README.md` - Grundlegende Informationen
2. `OUTLOOK-INTEGRATION.md` - Detaillierte Anleitungen und Troubleshooting
3. Fehlermeldungen im Script/App - oft selbsterklärend

**Häufigste Probleme:**
- PowerShell ExecutionPolicy → Siehe OUTLOOK-INTEGRATION.md
- Outlook COM-Fehler → Outlook neu starten
- JSON-Format-Fehler → Script erneut ausführen

---

## 🔐 Sicherheit

**Alle Tools sind lokal und sicher:**
- ✅ Keine Cloud-Verbindungen
- ✅ Keine Telemetrie
- ✅ Keine Datenübertragung
- ✅ Open Source (Code einsehbar)
- ✅ Läuft komplett auf deinem PC

**Behandle die JSON-Exports wie Email-Daten:**
- Enthalten die gleichen sensiblen Informationen wie deine Outlook-Emails
- Auf sicherem Speicherort ablegen
- Nach Verwendung ggf. löschen

---

## 📝 Lizenz

Internes Tool für Barmenia/Gothaer BGAV Hypercare Projekt.

---

**Letzte Aktualisierung:** 2025-11-14
