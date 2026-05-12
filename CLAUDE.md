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
- `ergo-vorgang-analyse.bas` → ZIP-Version **2.12** (34 Spalten, Auftrag_Datum, Triage vorn, Hinweis hinten, Kunde Petrol, Open Sans, Maklernummer-Priorisierung Agentur-Nr. > Pool-IDs)

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

### 4. Sonderfall `ergo-vorgang-analyse.hta` — KEIN ZIP / KEINE .hta.txt mehr
Ab v1.25 wird die `ergo-vorgang-analyse.hta` **NICHT** mehr als ZIP oder `.hta.txt`
ausgeliefert. Stattdessen gibt es eine **Standalone-Quellcode-Seite**
`hta-source.html`, die den kompletten HTA-Quellcode inline in einer Textarea
einbettet, plus einen Copy-Button und einen Blob-basierten Download-Button.
Vorteil: umgeht Corporate-Proxy-Filter und Cloudflare-Glitches.

**Bei jeder Aenderung an `ergo-vorgang-analyse.hta` muss `hta-source.html` neu
gebaut werden.** Dafuer gibt es das Skript `build-hta-source.sh` im Repo:

```bash
./build-hta-source.sh
```

Es liest die aktuelle HTA, escaped `&` zu `&amp;` (sonst werden die Ampersands
in der Textarea als Entity-Start interpretiert), wraps das Ganze in einen
HTML-Header (Styling + Anleitung + Copy-Button + Download-Button) und einen
Footer (JS fuer copyCode / downloadCode / selectAll). Die VERSION wird aus
dem `<HTA:APPLICATION VERSION="X.Y"/>`-Tag automatisch ausgelesen.

In `index.html` verlinkt **nur noch** `hta-source.html` — die Tile fuehrt
direkt auf die Quellcode-Seite. Die ZIP-Variante `ERGO-Vorgang-Analyse-HTA-*.zip`
und die `ergo-vorgang-analyse-v*.hta.txt`-Variante wurden bewusst entfernt.

VERSION im `<HTA:APPLICATION>`-Tag und sichtbarer Header-String (`<div class="ver">v1.X - ...`)
hochzaehlen wie gewohnt.

### 5. Sonderfall `claude-analyse.html` — Browser-KI mit Claude-API direkt
Standalone-Webseite (kein HTA, kein Outlook noetig). Laesst Anwender:
1. Eigenen Anthropic-API-Key eingeben (LocalStorage),
2. `.msg`-Dateien per Drag-Drop hochladen (msgreader via CDN parst CFB im Browser),
3. Pool-Erkennung lokal (POOL_RULES portiert aus HTA),
4. Bulk-Triage per direktem `fetch()` an `api.anthropic.com/v1/messages` mit
   Header `anthropic-dangerous-direct-browser-access: true` und Prompt-Caching
   (`cache_control: { type: "ephemeral" }`) auf dem System-Prompt,
5. Excel-Export via SheetJS, JSON-Export als Blob.

Default-Modell: `claude-sonnet-4-6`. Auswahl auch fuer `claude-haiku-4-5` (billig)
und `claude-opus-4-7` (Premium). Kostenschaetzung wird live angezeigt.

**Datenschutz-Warnung im UI:** Mail-Inhalte verlassen den Browser an Anthropic
(USA). Nicht fuer produktive ERGO-Daten. Fuer den freigegebenen Weg gibt es
das HTA-Tool mit ErgoGPT.

Versionierung: VERSION in der Inline-Header-Zeile (`Browser-KI v1.X`) hochzaehlen.
Kein ZIP, kein Build-Skript - die Datei wird direkt aus dem Repo serviert.
In `index.html` verlinkt die Tile `claude-analyse.html` (ERGO-Kategorie).
