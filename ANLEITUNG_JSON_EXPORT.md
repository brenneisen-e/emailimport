# Anleitung: JSON-Export Workflow

## Übersicht

Dieser Workflow ermöglicht es, Outlook Emails über Excel VBA als JSON zu exportieren und dann im HTA Tool zu laden. Das ist **viel schneller** als der direkte Outlook-Zugriff im HTA.

## Workflow

```
1. Excel öffnen (Tracking_Hypercare.xlsm)
   ↓
2. Makro "ExportOutlookToJSON" ausführen
   ↓
3. Outlook-Ordner auswählen
   ↓
4. JSON-Datei wird in Downloads gespeichert
   ↓
5. HTA Tool öffnen
   ↓
6. Button "📄 JSON aus Downloads laden" klicken
   ↓
7. Emails werden geladen und können bearbeitet werden
   ↓
8. "Nach Excel exportieren" klicken
```

## Installation

### Schritt 1: Excel VBA Makro importieren

1. Öffne **Tracking_Hypercare.xlsm** in Excel
2. Drücke **ALT + F11** (öffnet VBA Editor)
3. Klicke auf **Datei → Datei importieren**
4. Wähle die Datei **OutlookEmailExport.bas** aus
5. Das Modul "OutlookEmailExport" wird im VBA-Projekt angezeigt
6. Schließe den VBA Editor

### Schritt 2: Button für Makro erstellen (optional)

1. In Excel: **Entwicklertools → Einfügen → Schaltfläche**
2. Zeichne eine Schaltfläche auf das Tabellenblatt
3. Wähle das Makro **ExportOutlookToJSON** aus
4. Benenne die Schaltfläche z.B. "Outlook Emails exportieren"

## Verwendung

### Excel: Emails exportieren

1. Öffne **Tracking_Hypercare.xlsm**
2. Klicke auf die Schaltfläche **"Outlook Emails exportieren"**
   *(oder führe das Makro manuell aus: ALT+F8 → ExportOutlookToJSON)*
3. **Wähle den Outlook-Ordner** aus (z.B. Posteingang, Hypercare, etc.)
4. Das Makro:
   - Lädt die letzten 7 Tage (max. 100 Emails)
   - Zeigt Fortschritt in Excel-Statusleiste
   - Speichert JSON in Downloads als: `outlook_emails_YYYYMMDD_HHMMSS.json`
5. **Erfolgsmeldung** erscheint mit Dateipfad

### HTA: Emails laden

1. Öffne **bgav-hypercare-standalone.hta**
2. Klicke auf **"📄 JSON aus Downloads laden"**
3. Das Tool:
   - Sucht automatisch die neueste `outlook_emails_*.json` Datei
   - Lädt alle Emails
   - Zeigt sie in der Liste an
4. **Bearbeite die Emails** wie gewohnt:
   - Kategorie setzen
   - Cluster zuweisen
   - Status ändern
   - etc.
5. Wähle Emails aus und klicke **"Nach Excel exportieren"**

## Vorteile

✅ **Schnell**: JSON-Import dauert nur 1 Sekunde statt 30+ Sekunden
✅ **Zuverlässig**: Keine COM-Probleme oder Timeouts
✅ **Flexibel**: JSON-Dateien können gespeichert und später wiederverwendet werden
✅ **Offline**: HTA kann auch ohne Outlook-Verbindung arbeiten
✅ **Fortschritt**: Excel zeigt Fortschritt während des Exports an

## Technische Details

### JSON-Format

```json
[
  {
    "von_email": "sender@example.com",
    "von_name": "Max Mustermann",
    "betreff": "Anfrage zu Hypercare",
    "text": "Hallo,\n\nich habe eine Frage...",
    "datum": "2025-01-15 14:30:00",
    "anhänge": "attachment1.pdf, attachment2.xlsx",
    "kategorie": "Incident",
    "status": "Offen",
    "cluster": "Allgemein",
    "agentur": "Unbekannt"
  }
]
```

### Dateipfade

- **VBA Makro**: `OutlookEmailExport.bas`
- **JSON Export**: `%USERPROFILE%\Downloads\outlook_emails_*.json`
- **HTA Tool**: `bgav-hypercare-standalone.hta`

## Troubleshooting

### "Keine JSON-Dateien gefunden"
→ Führe zuerst das Excel-Makro aus

### "Downloads-Ordner nicht gefunden"
→ Überprüfe ob `%USERPROFILE%\Downloads` existiert

### "JSON-Datei ist leer"
→ Das Excel-Makro hat keine Emails gefunden (prüfe Datumsfilter)

### Excel-Makro startet nicht
→ Stelle sicher, dass Makros aktiviert sind (Datei → Optionen → Trust Center)

## Alte Methode vs. Neue Methode

| Kriterium | Alt (Direkt) | Neu (JSON) |
|-----------|-------------|------------|
| **Geschwindigkeit** | 30-60 Sekunden | 1-2 Sekunden |
| **Zuverlässigkeit** | Kann hängen | Sehr stabil |
| **Offline** | Nein | Ja |
| **Fortschritt** | Unklar | Sichtbar in Excel |
| **Outlook nötig** | Ja (HTA) | Ja (Excel) |

## Weitere Hinweise

- Die JSON-Dateien können als Backup aufbewahrt werden
- Alte JSON-Dateien können manuell aus Downloads gelöscht werden
- Das HTA lädt immer die **neueste** JSON-Datei
- Max. 100 Emails pro Export (kann im VBA-Code angepasst werden)
- Emails der letzten 7 Tage (kann im VBA-Code angepasst werden)
