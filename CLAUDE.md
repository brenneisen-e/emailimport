# CLAUDE.md - Projekt-Hinweise fuer Claude Code

## Wichtige Regeln bei Aenderungen

### 1. Versionsnummern hochsetzen
Bei jeder Aenderung an einem Tool MUESSEN folgende Versionsnummern hochgesetzt werden:

**a) HTA-VERSION** im `<HTA:APPLICATION>`-Tag der jeweiligen Datei:
- `bgav-testmail-extraktion.hta` → VERSION in Zeile ~19
- `outlook-export-modular.hta` → VERSION in Zeile ~19
- `outlook-selection-export.hta` → VERSION in Zeile ~19
- `pdf-massenupload.hta` → VERSION in Zeile ~19
- `ergo-email-batch.hta` → VERSION in Zeile ~19 (aktuell **1.1**)

**b) Globale Page-Version** der Homepage (`index.html`, Footer):
Im `<div class="version-info">` ganz am Ende (Zeile ~1550) steht die Page-Version
(aktuell **8.6**). **Bei JEDER Aenderung im Repo MUSS diese hochgezaehlt werden**,
damit auf der Homepage sofort sichtbar ist, dass es eine neue Version gibt.
Schema: Major.Minor — bei kleinen Aenderungen Minor +1, bei groesseren Major +1.

Fuer reine VBA-Tools (`ergo-email-batch.bas`, `ergo-vorgang-analyse.bas`) gibt es
keine VERSION-Konstante im Code; die Versionsnummer steckt **im Dateinamen** der
ZIP-Inhalte (siehe Punkt 2). Aktuell:
- `ergo-vorgang-analyse.bas` → ZIP-Version **2.5**

### 2. ZIP-Dateien aktualisieren — PFLICHT bei jeder Aenderung
**Nach jeder Code-Aenderung an einer HTA-Datei oder am `.bas`-Modul MUSS die zugehoerige
ZIP-Datei neu gebaut und mit-committed werden.** Sonst zeigt die Homepage (`index.html`)
weiterhin die alte Version zum Download. Die Links in `index.html` muessen nach
ZIP-Umbenennungen ebenfalls angepasst werden.

| Quelldatei | ZIP-Datei | Befehl |
|------------|-----------|--------|
| `bgav-testmail-extraktion.hta` | `BGAV-Testmail-Extraktion.zip` | `zip -j BGAV-Testmail-Extraktion.zip bgav-testmail-extraktion.hta` |
| `outlook-export-modular.hta` + `js/` + `css/` | `Email-Export-Tool.zip` | `zip -r Email-Export-Tool.zip outlook-export-modular.hta js/ css/` |
| `outlook-selection-export.hta` | `Selection-Export-Tool.zip` | `zip -j Selection-Export-Tool.zip outlook-selection-export.hta` |
| `pdf-massenupload.hta` | `PDF-Massenupload.zip` | `zip -j PDF-Massenupload.zip pdf-massenupload.hta` |
| `ergo-email-batch.hta` | `ERGO-Email-Batch-v<VER>.zip` | siehe unten (versionierte Inhalte) |
| `ergo-email-batch.bas` + `ANLEITUNG-ERGO-EXCEL.txt` | `ERGO-Excel-Tool-v<VER>.zip` | siehe unten (versionierte Inhalte) |
| `ergo-vorgang-analyse.bas` + `ANLEITUNG-ERGO-VORGANG-ANALYSE.txt` | `ERGO-Vorgang-Analyse-v<VER>.zip` | siehe unten (versionierte Inhalte) |

#### Versionierte ERGO-ZIPs bauen
Das ERGO-Tool fuehrt die Versionsnummer **im ZIP-Dateinamen UND im Dateinamen der
enthaltenen Dateien**. Der Source-Dateiname im Repo bleibt unversioniert
(`ergo-email-batch.bas` / `ergo-email-batch.hta`), beim Zippen wird umbenannt:

```bash
VER=1.1   # an aktuelle Version anpassen!

# Alte ZIPs entfernen (verhindert verwaiste Versionen im Repo)
rm -f ERGO-Email-Batch-v*.zip ERGO-Excel-Tool-v*.zip ERGO-Email-Batch.zip ERGO-Excel-Tool.zip

# Mit Rename in Tempfolder zippen
mkdir -p /tmp/ergo-zip-build
cp ergo-email-batch.bas /tmp/ergo-zip-build/ergo-email-batch-v$VER.bas
cp ergo-email-batch.hta /tmp/ergo-zip-build/ergo-email-batch-v$VER.hta
cp ANLEITUNG-ERGO-EXCEL.txt /tmp/ergo-zip-build/

cd /tmp/ergo-zip-build
zip -j "$OLDPWD/ERGO-Email-Batch-v$VER.zip" ergo-email-batch-v$VER.hta
zip -j "$OLDPWD/ERGO-Excel-Tool-v$VER.zip" ergo-email-batch-v$VER.bas ANLEITUNG-ERGO-EXCEL.txt
cd "$OLDPWD"
```

Fuer das Vorgang-Analyse-Tool analog (eigene Versionsnummer):

```bash
VA_VER=1.0
rm -f ERGO-Vorgang-Analyse-v*.zip
mkdir -p /tmp/ergo-va-build
cp ergo-vorgang-analyse.bas /tmp/ergo-va-build/ergo-vorgang-analyse-v$VA_VER.bas
cp ANLEITUNG-ERGO-VORGANG-ANALYSE.txt /tmp/ergo-va-build/
(cd /tmp/ergo-va-build && \
   zip -j "$OLDPWD/ERGO-Vorgang-Analyse-v$VA_VER.zip" \
       ergo-vorgang-analyse-v$VA_VER.bas ANLEITUNG-ERGO-VORGANG-ANALYSE.txt)
```

**Anschliessend in `index.html` die Download-Links auf die neue Versionsnummer anpassen**
(drei Stellen: `ERGO-Email-Batch-v<VER>.zip`, `ERGO-Excel-Tool-v<VER>.zip`,
`ERGO-Vorgang-Analyse-v<VER>.zip`).

### 3. Reihenfolge bei Aenderungen
1. Code-Aenderung in HTA-/BAS-Datei
2. Tool-VERSION hochsetzen (HTA-Tag bzw. ERGO `$VER`-Variable)
3. **Globale Page-Version** in `index.html` (Footer) hochzaehlen
4. ZIP neu bauen (siehe Tabelle/Skript oben) — auch bei reinen `.bas`-Aenderungen
5. Bei Versions-Bump der ERGO-ZIPs: Links in `index.html` anpassen
6. Alles zusammen committen (Source + ZIP + index.html)
