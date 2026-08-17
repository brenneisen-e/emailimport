#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Bettet HTA-Quellcode als Base64 in die HTML-Seiten ein, damit die
# "Quellcode kopieren"-Buttons ohne Netzwerk-Fetch funktionieren
# (Corporate-Proxys filtern .hta- und .zip-Requests weg).
#
# Aufruf:  ./scripts/build-embedded-sources.sh
#
# Ersetzt in jedem TARGET die Werte aller vorhandenen EMBEDDED_SOURCES-
# Eintraege durch die aktuelle Base64-Kodierung der jeweiligen HTA-Datei.
# Eintraege, die eine Seite nicht hat, werden uebersprungen.
# Neue HTA einbetten: Datei unten in FILES ergaenzen + Key in der jeweiligen
# Seite im Objekt EMBEDDED_SOURCES anlegen.
# ---------------------------------------------------------------------------
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

TARGETS=(
    "downloads.html"
    "index.html"
)
FILES=(
    "un-laenderliste-excel.hta"
    "ergo-mail-statistik.hta"
    "outlook-regeln-visualisierung.hta"
)

for TARGET in "${TARGETS[@]}"; do
    [ -f "$TARGET" ] || { echo "FEHLER: $TARGET nicht gefunden" >&2; exit 1; }
done

for HTA in "${FILES[@]}"; do
    [ -f "$HTA" ] || { echo "FEHLER: $HTA nicht gefunden" >&2; exit 1; }

    # Base64 einzeilig erzeugen (GNU: -w0)
    B64="$(base64 -w0 "$HTA")"

    for TARGET in "${TARGETS[@]}"; do
        # Eintrag '<datei>': '<...>' ersetzen. Python fuer sicheres Ersetzen
        # (Base64 enthaelt nur [A-Za-z0-9+/=], also unkritisch).
        python3 - "$TARGET" "$HTA" "$B64" <<'PY'
import re, sys
target, hta, b64 = sys.argv[1], sys.argv[2], sys.argv[3]
src = open(target, encoding='utf-8').read()
pattern = r"('" + re.escape(hta) + r"':\s*')[^']*(')"
new, n = re.subn(pattern, lambda m: m.group(1) + b64 + m.group(2), src, count=1)
if n == 0:
    print("   uebersprungen: %s hat keinen Eintrag fuer %s" % (target, hta))
    sys.exit(0)
open(target, 'w', encoding='utf-8').write(new)
print("OK: %s -> %s (%d Bytes Base64)" % (hta, target, len(b64)))
PY
    done
done

echo "Fertig. Bitte die geaenderten HTML-Seiten committen."
