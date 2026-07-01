#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Bettet HTA-Quellcode als Base64 in downloads.html ein, damit der
# "Quellcode kopieren"-Button ohne Netzwerk-Fetch funktioniert
# (Corporate-Proxys filtern .hta-Requests weg).
#
# Aufruf:  ./scripts/build-embedded-sources.sh
#
# Ersetzt in downloads.html den Wert von
#   EMBEDDED_SOURCES['un-laenderliste-excel.hta']
# durch die aktuelle Base64-Kodierung der HTA-Datei.
# ---------------------------------------------------------------------------
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

HTA="un-laenderliste-excel.hta"
TARGET="downloads.html"

[ -f "$HTA" ]    || { echo "FEHLER: $HTA nicht gefunden" >&2; exit 1; }
[ -f "$TARGET" ] || { echo "FEHLER: $TARGET nicht gefunden" >&2; exit 1; }

# Base64 einzeilig erzeugen (GNU: -w0)
B64="$(base64 -w0 "$HTA")"

# Zeile 'un-laenderliste-excel.hta': '<...>', ersetzen. Python fuer sicheres
# Escapen/Ersetzen (der Base64-String enthaelt nur [A-Za-z0-9+/=], also unkritisch).
python3 - "$TARGET" "$B64" <<'PY'
import re, sys
target, b64 = sys.argv[1], sys.argv[2]
src = open(target, encoding='utf-8').read()
pattern = r"('un-laenderliste-excel\.hta':\s*')[^']*(')"
new, n = re.subn(pattern, lambda m: m.group(1) + b64 + m.group(2), src, count=1)
if n != 1:
    sys.exit("FEHLER: EMBEDDED_SOURCES-Eintrag nicht gefunden (Treffer=%d)" % n)
open(target, 'w', encoding='utf-8').write(new)
print("OK: %d Bytes Base64 eingebettet." % len(b64))
PY

echo "Fertig. Bitte downloads.html committen."
