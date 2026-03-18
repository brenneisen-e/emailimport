# CLAUDE.md - Projekt-Hinweise fuer Claude Code

## Wichtige Regeln bei Aenderungen

### 1. Globale Versionsnummer hochsetzen
Bei jeder Aenderung MUSS die Versionsnummer in `index.html` (Zeile ~1420) hochgesetzt werden (z.B. von "8.0" auf "8.1").
Das ist die zentrale Versionsnummer fuer das gesamte Projekt, angezeigt unten auf der Homepage.

### 2. ZIP-Dateien aktualisieren
Nach jeder Aenderung an einer HTA-Datei MUSS die zugehoerige ZIP-Datei neu erstellt werden:

| HTA-Datei | ZIP-Datei | Befehl |
|-----------|-----------|--------|
| `bgav-testmail-extraktion.hta` | `BGAV-Testmail-Extraktion.zip` | `zip -j BGAV-Testmail-Extraktion.zip bgav-testmail-extraktion.hta` |
| `outlook-export-modular.hta` + `js/` + `css/` | `Email-Export-Tool.zip` | `zip -r Email-Export-Tool.zip outlook-export-modular.hta js/ css/` |
| `outlook-selection-export.hta` | `Selection-Export-Tool.zip` | `zip -j Selection-Export-Tool.zip outlook-selection-export.hta` |
| `pdf-massenupload.hta` | `PDF-Massenupload.zip` | `zip -j PDF-Massenupload.zip pdf-massenupload.hta` |

### 3. Reihenfolge
1. Code-Aenderung in der HTA-Datei / index.html / js/ / css/
2. Globale Versionsnummer in `index.html` hochsetzen
3. ZIP aktualisieren (falls HTA betroffen)
4. Alles zusammen committen (HTA + ZIP + index.html)
