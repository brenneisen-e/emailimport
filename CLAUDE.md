# CLAUDE.md - Projekt-Hinweise fuer Claude Code

## Wichtige Regeln bei Aenderungen

### 1. Versionsnummer hochsetzen
Bei jeder Aenderung an einer HTA-Datei MUSS die VERSION im `<HTA:APPLICATION>`-Tag hochgesetzt werden (z.B. von "7.0" auf "7.1").

HTA-Dateien und ihre Versionen:
- `bgav-testmail-extraktion.hta` → VERSION in Zeile ~19
- `outlook-export-modular.hta` → VERSION in Zeile ~19
- `outlook-selection-export.hta` → VERSION in Zeile ~19
- `pdf-massenupload.hta` → VERSION in Zeile ~19

### 2. ZIP-Dateien aktualisieren
Nach jeder Aenderung an einer HTA-Datei MUSS die zugehoerige ZIP-Datei neu erstellt werden:

| HTA-Datei | ZIP-Datei | Befehl |
|-----------|-----------|--------|
| `bgav-testmail-extraktion.hta` | `BGAV-Testmail-Extraktion.zip` | `zip -j BGAV-Testmail-Extraktion.zip bgav-testmail-extraktion.hta` |
| `outlook-export-modular.hta` + `js/` + `css/` | `Email-Export-Tool.zip` | `zip -r Email-Export-Tool.zip outlook-export-modular.hta js/ css/` |
| `outlook-selection-export.hta` | `Selection-Export-Tool.zip` | `zip -j Selection-Export-Tool.zip outlook-selection-export.hta` |
| `pdf-massenupload.hta` | `PDF-Massenupload.zip` | `zip -j PDF-Massenupload.zip pdf-massenupload.hta` |

### 3. Reihenfolge
1. Code-Aenderung in der HTA-Datei
2. VERSION hochsetzen
3. ZIP aktualisieren
4. Alles zusammen committen (HTA + ZIP)
