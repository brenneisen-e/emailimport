#!/bin/bash
# build-workshop-hta.sh
# Leitet aus ergo-vorgang-analyse.hta ein eigenstaendiges "ERGO Workshop"-HTA
# ab (workshop.hta), das beim Start direkt das weisse Workshop-Frontend
# oeffnet. Baut zusaetzlich die Copy-Paste-Seite workshop-source.html
# (analog hta-source.html).

set -e
cd "$(dirname "$0")"

HTA_SRC="ergo-vorgang-analyse.hta"
HTA_OUT="workshop.hta"
OUT="workshop-source.html"

if [ ! -f "$HTA_SRC" ]; then
    echo "ERROR: $HTA_SRC nicht gefunden"
    exit 1
fi

VER=$(grep -oP 'VERSION="\K[0-9]+\.[0-9]+' "$HTA_SRC" | head -1)
echo "Building $HTA_OUT + $OUT (Workshop, Basis HTA v$VER)..."

# 1) Standalone-HTA ableiten:
#    - eigener Titel/APPLICATIONNAME (sonst SINGLEINSTANCE-Kollision mit Haupttool)
#    - nach setupDragDrop() automatisch das Workshop-Overlay oeffnen
sed -e 's|<title>ERGO Vorgang-Analyse</title>|<title>ERGO Workshop</title>|' \
    -e 's|APPLICATIONNAME="ERGO Vorgang-Analyse"|APPLICATIONNAME="ERGO Workshop"|' \
    -e 's|    setupDragDrop();|    setupDragDrop();\n    try { setTimeout(function(){ openWorkshopMode(); }, 350); } catch(e) {}|' \
    "$HTA_SRC" > "$HTA_OUT"

echo "OK - $HTA_OUT gebaut ($(wc -c < "$HTA_OUT") bytes)"

# 2) Copy-Paste-Seite bauen (Wrapper unten als Heredoc eingebettet)
cat > "$OUT" << HEADER_END
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>ERGO Workshop v${VER} - HTA Quellcode</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Tahoma, sans-serif; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #f1f5f9; min-height: 100vh; padding: 1.5rem; }
.container { max-width: 1400px; margin: 0 auto; }
.card { background: rgba(30,41,59,0.85); border-radius: 14px; padding: 1.5rem; box-shadow: 0 4px 20px rgba(0,0,0,0.4); margin-bottom: 1rem; }
h1 { font-size: 22px; margin-bottom: 0.5rem; }
.subtitle { color: #94a3b8; font-size: 13px; }
.info { background: rgba(200,16,46,0.12); border: 1px solid #c8102e; border-radius: 8px; padding: 1rem; margin: 1rem 0; font-size: 13px; color: #cbd5e1; }
.info strong { color: #fca5a5; }
.info ol { margin-left: 1.25rem; margin-top: 0.5rem; }
.info code { background: rgba(0,0,0,0.3); padding: 1px 6px; border-radius: 3px; color: #fbbf24; font-family: 'Consolas', monospace; }
.actions { display: flex; gap: 0.75rem; flex-wrap: wrap; align-items: center; margin: 1rem 0; }
button { padding: 12px 24px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.15s; display: inline-flex; align-items: center; gap: 8px; }
.btn-copy { background: linear-gradient(135deg, #10b981, #059669); color: white; }
.btn-copy:hover { background: linear-gradient(135deg, #34d399, #10b981); transform: translateY(-1px); }
.btn-copy.copied { background: linear-gradient(135deg, #3b82f6, #1d4ed8); }
.btn-download { background: linear-gradient(135deg, #c8102e, #a30d24); color: white; }
.btn-download:hover { background: linear-gradient(135deg, #e11d48, #c8102e); transform: translateY(-1px); }
.btn-icon { background: #475569; color: #f1f5f9; padding: 12px 16px; }
#status { font-size: 13px; color: #6ee7b7; font-weight: 500; margin-left: 0.5rem; }
.meta { font-size: 12px; color: #94a3b8; margin: 0.5rem 0 1rem; }
.meta strong { color: #f1f5f9; }
textarea#code { width: 100%; height: 70vh; background: #0f172a; color: #e2e8f0; border: 1px solid #334155; border-radius: 8px; padding: 1rem; font-family: 'Consolas','Courier New',monospace; font-size: 11px; line-height: 1.4; resize: vertical; white-space: pre; }
textarea#code:focus { outline: 2px solid #c8102e; }
svg.ico { width: 18px; height: 18px; vertical-align: -3px; }
</style>
</head>
<body>
<div class="container">
<div class="card">
<h1>ERGO Workshop <strong>v${VER}</strong> &mdash; HTA Quellcode</h1>
<p class="subtitle">Eigenstaendiges Workshop-Tool (startet direkt im weissen Workshop-Frontend)</p>

<div class="info">
<strong>Wie nutze ich diese Seite?</strong>
<ol>
<li>Klicke auf <strong>&quot;Code kopieren&quot;</strong> oder <strong>&quot;Als .hta herunterladen&quot;</strong></li>
<li>Beim Kopieren: Notepad / Notepad++ oeffnen und einfuegen (<code>Strg+V</code>)</li>
<li>Speichern als <code>ergo-workshop.hta</code> (Encoding <strong>UTF-8</strong>)</li>
<li>Doppelklick &rarr; HTA startet direkt im Workshop-Modus</li>
</ol>
</div>

<div class="meta">Dateigroesse: <strong id="meta-size">...</strong> | Zeilen: <strong id="meta-lines">...</strong></div>

<div class="actions">
<button class="btn-copy" id="btnCopy" onclick="copyCode()">
<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
Code kopieren
</button>
<button class="btn-download" onclick="downloadCode()">
<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
Als .hta herunterladen
</button>
<button class="btn-icon" onclick="selectAll()" title="Gesamten Text markieren">
<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 7v10a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V7"/><path d="M22 7l-10 5L2 7"/></svg>
Alles markieren
</button>
<span id="status"></span>
</div>

<textarea id="code" readonly spellcheck="false">
HEADER_END

# HTA-Inhalt anhaengen, & zu &amp; escapen (sonst Entity-Interpretation in Textarea)
sed 's/\&/\&amp;/g' "$HTA_OUT" >> "$OUT"

cat >> "$OUT" << 'FOOTER_END'
</textarea>

</div>
</div>

<script>
function copyCode() {
    var ta = document.getElementById('code');
    var status = document.getElementById('status');
    try {
        ta.focus(); ta.select(); ta.setSelectionRange(0, ta.value.length);
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(ta.value).then(showCopied).catch(function() {
                if (document.execCommand('copy')) showCopied();
                else status.innerText = 'Kopieren fehlgeschlagen - bitte manuell markieren + Strg+C';
            });
        } else {
            if (document.execCommand('copy')) showCopied();
            else status.innerText = 'Kopieren fehlgeschlagen - bitte manuell markieren + Strg+C';
        }
    } catch(e) { status.innerText = 'Kopieren fehlgeschlagen: ' + e.message; }
}
function showCopied() {
    var btn = document.getElementById('btnCopy');
    var status = document.getElementById('status');
    btn.classList.add('copied');
    btn.innerHTML = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> Kopiert!';
    status.innerText = String(document.getElementById('code').value.length) + ' Zeichen in Zwischenablage.';
    setTimeout(function() {
        btn.classList.remove('copied');
        btn.innerHTML = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Code kopieren';
        status.innerText = '';
    }, 3000);
}
function downloadCode() {
    var content = document.getElementById('code').value;
    var blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'ergo-workshop.hta';
    document.body.appendChild(a); a.click();
    setTimeout(function() { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
    document.getElementById('status').innerText = 'Download gestartet.';
    setTimeout(function() { document.getElementById('status').innerText = ''; }, 3000);
}
function selectAll() {
    var ta = document.getElementById('code');
    ta.focus(); ta.select(); ta.setSelectionRange(0, ta.value.length);
    document.getElementById('status').innerText = 'Markiert - jetzt Strg+C druecken.';
    setTimeout(function() { document.getElementById('status').innerText = ''; }, 3000);
}
(function() {
    var v = document.getElementById('code').value;
    document.getElementById('meta-size').innerText = Math.round(v.length / 1024) + ' KB';
    document.getElementById('meta-lines').innerText = v.split('\n').length;
})();
</script>
</body>
</html>
FOOTER_END

echo "OK - $OUT gebaut (Workshop v$VER, $(wc -c < "$OUT") bytes)"
