# BGAV Hypercare - Email Review Tool

Tool zur Uberprufung und Kategorisierung von Hypercare-Emails fur das Barmenia/Gothaer-Projekt.

## Drei Versionen verfugbar

### Web-Version (Cloudflare Pages) - NEU!

**`index.html`** - Moderne Web-App mit Dark Mode Design

- Modernes Dark-Mode UI mit Animationen
- Funktioniert in jedem Browser
- Kann auf Cloudflare Pages gehostet werden
- JSON-Import per Drag & Drop
- Fortschrittsanzeigen bei allen Operationen
- CSV-Export

**Deployment auf Cloudflare Pages:**
1. Repository mit GitHub verbinden
2. Build command: (leer lassen - statische Site)
3. Build output directory: `/` oder `.`
4. Fertig! Die App ist unter deiner Cloudflare-URL erreichbar

**Live Demo:** Lade einfach `index.html` im Browser

---

### HTA-Version (Windows-Desktop)

**`bgav-hypercare-standalone.hta`** - Fur Windows mit direktem Outlook-Zugriff

- Direkter Outlook-Zugriff - Keine Scripts, keine Exports
- Direkter Excel-Export - Offnet sich automatisch in Excel
- Komplett offline - Keine Internetverbindung notig

**Quick Start:**
1. Doppelklick auf `bgav-hypercare-standalone.hta`
2. "Emails laden" klicken
3. Postfach auswahlen
4. Review & Kategorisierung
5. "Excel Export" klicken

---

### HTML-Version (Legacy)

**`bgav-hypercare-email-review.html`** - Altere Version

- Fur Mac, Linux, Windows
- Benotigt JSON-Export aus Outlook

---

## Features v3.0

### Provisionsreklamation b24 Template
- **Automatische Erkennung**: Betreff "Provisionsreklamation b24" oder Emails von redaktion@barmenia.de
- **Parser fur strukturierte Daten**: Extrahiert alle Felder wie:
  - `vermittlernummer_vermittler`, `vorname_vermittler`, `nachname_vermittler`
  - `versicherungsnummer_kunde`, `vorname_kunde`, `nachname_kunde`
  - `reklamation_provision`, `reklamation_absatz`, `datenschutz`, `nachricht`

### Mehrfachkategorisierung
- Emails konnen **mehrere Cluster** gleichzeitig haben (z.B. KV + KFZ)
- **Klickbare Tags** zum An-/Abwahlen der Cluster
- Filter funktioniert mit Mehrfachauswahl

### Outlook-ahnliches Design
- Email-Ansicht wie in Outlook
- Header mit Von / Datum / Betreff
- Anhange-Leiste mit Datei-Icons

### Email-Thread Baumansicht
- Original-Email oben angezeigt
- Antworten darunter als eingeruckte Zweige
- Auf-/Zuklappbar per Klick

### Moderne UI
- Dark Mode mit Glow-Effekten
- Smooth Animationen uberall
- Fortschrittsanzeigen bei allen Operationen
- Skeleton Loading States

---

## Outlook-Integration

### PowerShell-Script (empfohlen)

```powershell
# Ausfuhren in PowerShell
.\Export-OutlookEmails.ps1
```

Das Script:
1. Verbindet sich mit Outlook
2. Exportiert ausgewahlte Emails oder einen ganzen Ordner
3. Speichert als JSON-Datei

### JSON-Format

```json
[
  {
    "datum": "2025-11-14T10:30:00",
    "von_email": "max.mueller@agentur.de",
    "von_name": "Max Muller",
    "betreff": "Provisionsreklamation b24",
    "text": "reklamation_provision : true\nvermittlernummer_vermittler : 00400288\n...",
    "anhaenge": ["dokument.pdf"]
  }
]
```

---

## Cloudflare Pages Deployment

### Option 1: Via GitHub

1. Repository zu GitHub pushen
2. Cloudflare Dashboard > Pages > Create a project
3. "Connect to Git" > Repository auswahlen
4. Settings:
   - Build command: (leer)
   - Build output directory: `/`
5. Deploy!

### Option 2: Direct Upload

1. Cloudflare Dashboard > Pages > Create a project
2. "Upload assets" auswahlen
3. `index.html` hochladen
4. Deploy!

---

## Dateien-Ubersicht

| Datei | Beschreibung |
|-------|-------------|
| `index.html` | **Moderne Web-App** - Fur Cloudflare Pages |
| `bgav-hypercare-standalone.hta` | Windows HTA mit Outlook-Integration |
| `bgav-hypercare-email-review.html` | Legacy HTML-Version |
| `Export-OutlookEmails.ps1` | PowerShell Export-Script |
| `beispiel-emails.json` | Test-Daten |

---

## Datenschutz

Alle Daten bleiben lokal im Browser. Es werden keine Daten an externe Server ubertragen.

## Lizenz

Internes Tool fur Barmenia/Gothaer BGAV Hypercare Projekt.
