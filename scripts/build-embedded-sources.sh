#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Bettet HTA-Quellcode als Base64 in downloads.html ein, damit die
# "Quellcode kopieren"-Buttons ohne Netzwerk-Fetch funktionieren
# (Corporate-Proxys filtern .hta-Requests weg).
#
# Aufruf:  ./scripts/build-embedded-sources.sh
#
# Ersetzt in downloads.html die Werte aller EMBEDDED_SOURCES-Eintraege
# durch die aktuelle Base64-Kodierung der jeweiligen HTA-Datei.
# Neue HTA einbetten: Datei unten in FILES ergaenzen + Key in
# EMBEDDED_SOURCES (downloads.html) anlegen.
# ---------------------------------------------------------------------------
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

TARGET="downloads.html"
FILES=(
    "un-laenderliste-excel.hta"
    "ergo-mail-statistik.hta"
)

[ -f "$TARGET" ] || { echo "FEHLER: $TARGET nicht gefunden" >&2; exit 1; }

for HTA in "${FILES[@]}"; do
    [ -f "$HTA" ] || { echo "FEHLER: $HTA nicht gefunden" >&2; exit 1; }

    # Base64 einzeilig erzeugen (GNU: -w0)
    B64="$(base64 -w0 "$HTA")"

    # Eintrag '<datei>': '<...>' ersetzen. Python fuer sicheres Ersetzen
    # (Base64 enthaelt nur [A-Za-z0-9+/=], also unkritisch).
    python3 - "$TARGET" "$HTA" "$B64" <<'PY'
import re, sys
target, hta, b64 = sys.argv[1], sys.argv[2], sys.argv[3]
src = open(target, encoding='utf-8').read()
pattern = r"('" + re.escape(hta) + r"':\s*')[^']*(')"
new, n = re.subn(pattern, lambda m: m.group(1) + b64 + m.group(2), src, count=1)
if n != 1:
    sys.exit("FEHLER: EMBEDDED_SOURCES-Eintrag fuer %s nicht gefunden (Treffer=%d)" % (hta, n))
open(target, 'w', encoding='utf-8').write(new)
print("OK: %s -> %d Bytes Base64 eingebettet." % (hta, len(b64)))
PY
done

echo "Fertig. Bitte downloads.html committen."
