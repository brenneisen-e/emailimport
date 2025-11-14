# Dateien-Übersicht

Dieses Repository enthält alle Dateien für das BGAV Hypercare Email Review Tool.

## 📁 Hauptdateien

### `bgav-hypercare-email-review.html`
**Die Hauptanwendung** - Standalone HTML-App für Email-Review

- **Was macht sie?** Email-Review-Tool mit automatischer Kategorisierung und Excel-Export
- **Wie verwenden?** Einfach im Browser öffnen (Doppelklick)
- **Abhängigkeiten:** Keine (außer XLSX-Library via CDN)
- **Offline-fähig:** Ja
- **Größe:** ~45 KB

**Features:**
- JSON-Import von Emails
- Automatische Kategorisierung (Incident/Rückfrage, Cluster)
- Such- und Filterfunktionen
- Batch-Bearbeitung
- Excel-Export für Hypercare-Template

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
├── bgav-hypercare-email-review.html    # Haupt-App (Standalone)
│
├── Export-OutlookEmails.ps1             # PowerShell Export-Script
├── OutlookExportMacro.vba               # VBA Makro für Outlook
│
├── beispiel-emails.json                 # Test-Daten
│
├── README.md                            # Haupt-Dokumentation
├── OUTLOOK-INTEGRATION.md               # Outlook-Anleitung
└── DATEIEN-UEBERSICHT.md               # Diese Datei
```

---

## 🚀 Quick Start

### Für Erstnutzer (Testing)

1. **App öffnen:**
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

### Für Produktivnutzung

1. **Outlook-Export einrichten:**
   - Methode wählen (PowerShell oder VBA)
   - Siehe `OUTLOOK-INTEGRATION.md` für Details

2. **Emails aus Outlook exportieren:**
   ```powershell
   # PowerShell-Methode
   .\Export-OutlookEmails.ps1
   ```

3. **In App laden und verarbeiten:**
   - HTML-App öffnen
   - Exportierte JSON-Datei laden
   - Review durchführen
   - Nach Excel exportieren

---

## 💾 Workflow-Diagramm

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
        │                       │
        │ - Kategorisierung     │
        │ - Filterung           │
        │ - Batch-Bearbeitung   │
        └───────┬───────────────┘
                ▼
        ┌───────────────┐
        │  Excel-Export │
        │  (Hypercare)  │
        └───────────────┘
```

---

## 📊 Dateigrößen & Performance

| Datei | Größe | Ladezeit | Performance-Hinweis |
|-------|-------|----------|---------------------|
| `bgav-hypercare-email-review.html` | ~45 KB | < 1s | Auch mit 10.000 Emails flüssig |
| `Export-OutlookEmails.ps1` | ~12 KB | - | Sehr schnell, auch große Exports |
| `OutlookExportMacro.vba` | ~11 KB | - | OK für <1.000 Emails |
| `beispiel-emails.json` | ~3 KB | < 1s | Instant |

**Performance-Empfehlungen:**
- PowerShell-Script: Ideal für große Mengen (>1.000 Emails)
- VBA-Makro: Gut für kleine/mittlere Mengen (<1.000 Emails)
- HTML-App: Performant bis ~5.000 Emails im Browser

---

## 🔄 Versions-Historie

### Version 1.0 (2025-11-14)
- ✨ Initiale Version
- ✅ Standalone HTML-App
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
