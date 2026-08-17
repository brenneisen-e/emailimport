#!/bin/bash
# build-regeln-source.sh
# Baut outlook-regeln-source.html neu auf Basis von outlook-regeln-visualisierung.hta.
#
# Zweck: Standalone-Quellcode-Seite zum Kopieren (Copy & Paste) des HTA-Codes,
# weil ZIP- und .hta-Downloads von Corporate-Proxys weggefiltert werden.
# Der komplette HTA-Quellcode steht in einer Textarea: "Code kopieren" legt ihn
# in die Zwischenablage, "Als .hta herunterladen" erzeugt die Datei per Blob.
# Die Seite ist self-contained (kein CDN, kein externes CSS/JS).
#
# Aufruf: ./build-regeln-source.sh
# Nach JEDER Aenderung an outlook-regeln-visualisierung.hta ausfuehren!

set -e

cd "$(dirname "$0")"

HTA_SRC="outlook-regeln-visualisierung.hta"
OUT="outlook-regeln-source.html"

if [ ! -f "$HTA_SRC" ]; then
    echo "ERROR: $HTA_SRC nicht gefunden"
    exit 1
fi

VER=$(grep -oP 'VERSION="\K[0-9]+\.[0-9]+' "$HTA_SRC" | head -1)
echo "Building $OUT for HTA v$VER..."

cat > "$OUT" << HEADER_END
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Outlook-Regeln visualisieren v${VER} &ndash; Quellcode</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: 'Segoe UI', Tahoma, sans-serif;
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
    color: #f1f5f9;
    min-height: 100vh;
    padding: 1.5rem;
}
.container { max-width: 1400px; margin: 0 auto; }
.card {
    background: rgba(30, 41, 59, 0.85);
    border-radius: 14px;
    padding: 1.5rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    margin-bottom: 1rem;
}
h1 { font-size: 22px; margin-bottom: 0.5rem; }
.subtitle { color: #94a3b8; font-size: 13px; }
.info {
    background: rgba(59, 130, 246, 0.1);
    border: 1px solid #3b82f6;
    border-radius: 8px;
    padding: 1rem;
    margin: 1rem 0;
    font-size: 13px;
    color: #cbd5e1;
}
.info strong { color: #93c5fd; }
.info ol, .info ul { margin-left: 1.25rem; margin-top: 0.5rem; }
.info li { margin-bottom: 3px; }
.info code {
    background: rgba(0,0,0,0.3);
    padding: 1px 6px;
    border-radius: 3px;
    color: #fbbf24;
    font-family: 'Consolas', monospace;
}
.host {
    background: rgba(16, 185, 129, 0.08);
    border: 1px solid #10b981;
    border-radius: 8px;
    padding: 1rem;
    margin: 1rem 0;
    font-size: 13px;
    color: #cbd5e1;
}
.host strong { color: #6ee7b7; }
.actions {
    display: flex;
    gap: 0.75rem;
    flex-wrap: wrap;
    align-items: center;
    margin: 1rem 0;
}
button {
    padding: 12px 24px;
    border: none;
    border-radius: 8px;
    cursor: pointer;
    font-size: 14px;
    font-weight: 600;
    transition: all 0.15s;
    display: inline-flex;
    align-items: center;
    gap: 8px;
}
.btn-copy { background: linear-gradient(135deg, #10b981, #059669); color: white; }
.btn-copy:hover { background: linear-gradient(135deg, #34d399, #10b981); transform: translateY(-1px); }
.btn-copy.copied { background: linear-gradient(135deg, #3b82f6, #1d4ed8); }
.btn-download { background: linear-gradient(135deg, #d97706, #b45309); color: white; }
.btn-download:hover { background: linear-gradient(135deg, #f59e0b, #d97706); transform: translateY(-1px); }
.btn-icon { background: #475569; color: #f1f5f9; padding: 12px 16px; }
.btn-icon:hover { background: #64748b; }
#status { font-size: 13px; color: #6ee7b7; font-weight: 500; margin-left: 0.5rem; }
#status.err { color: #fca5a5; }
.meta { font-size: 12px; color: #94a3b8; margin: 0.5rem 0 1rem; }
.meta strong { color: #f1f5f9; }
textarea#code {
    width: 100%;
    height: 70vh;
    background: #0f172a;
    color: #e2e8f0;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 1rem;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11px;
    line-height: 1.4;
    resize: vertical;
    white-space: pre;
}
textarea#code:focus { outline: 2px solid #ef4444; }
svg.ico { width: 18px; height: 18px; vertical-align: -3px; }
.back { color: #94a3b8; font-size: 12px; text-decoration: none; }
.back:hover { color: #f1f5f9; }
</style>
</head>
<body>
<div class="container">
<div class="card">
<a class="back" href="index.html">&larr; Zur Startseite</a>
<h1 style="margin-top:8px;">Outlook-Regeln visualisieren <strong>v${VER}</strong> &ndash; Quellcode</h1>
<p class="subtitle">Regeln ausgew&auml;hlter (Gruppen-)Postf&auml;cher auslesen und die Logik als Excel-Mappe darstellen &middot; kein Download-ZIP n&ouml;tig</p>

<div class="info">
<strong>So startest du das Tool (ohne ZIP-Download):</strong>
<ol>
<li>Auf <strong>&bdquo;Code kopieren&ldquo;</strong> klicken (oder <strong>&bdquo;Alles markieren&ldquo;</strong> + <code>Strg+C</code>)</li>
<li>Notepad / Notepad++ &ouml;ffnen und einf&uuml;gen (<code>Strg+V</code>)</li>
<li>Speichern als <code>outlook-regeln-visualisierung.hta</code> &ndash; Dateityp <strong>Alle Dateien</strong>, Encoding <strong>UTF-8</strong> (nicht als .txt!)</li>
<li>Outlook geoffnet lassen, Doppelklick auf die .hta &rarr; Postf&auml;cher ankreuzen &rarr; Excel</li>
</ol>
</div>

<div class="host">
<strong>Warum Copy &amp; Paste statt ZIP?</strong><br>
Corporate-Proxys filtern ZIP- und .hta-Downloads weg. Der komplette HTA-Code
steht deshalb unten im Textfeld &ndash; einmal kopieren, lokal als .hta speichern,
fertig. Nichts wird nachgeladen, die .hta l&auml;uft danach komplett offline auf dem
Arbeitsplatz (Outlook + Excel per COM).
</div>

<div class="meta">HTA-Version: <strong>v${VER}</strong> &middot; Dateigr&ouml;&szlig;e: <strong id="meta-size">&hellip;</strong> &middot; Zeilen: <strong id="meta-lines">&hellip;</strong></div>

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

# HTA-Inhalt anhaengen. & und < escapen, damit der Browser den Text in der
# Textarea unveraendert wieder herausgibt (sonst werden Entities wie &auml;
# beim Kopieren aufgeloest).
sed -e 's/&/\&amp;/g' -e 's/</\&lt;/g' "$HTA_SRC" >> "$OUT"

cat >> "$OUT" << 'FOOTER_END'
</textarea>

</div>
</div>

<script>
/* Alles lokal, keine Netzwerkzugriffe. */
function setStatus(text, isError) {
    var el = document.getElementById('status');
    el.className = isError ? 'err' : '';
    el.innerText = text || '';
}

/* Kopieren synchron im Klick: execCommand zuerst (funktioniert auch ohne
   HTTPS und ohne Clipboard-Berechtigung), Clipboard-API als Fallback. */
function copyText(text, okMessage) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.top = '-1000px';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try { ta.setSelectionRange(0, text.length); } catch (e) {}
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    if (ok) { setStatus(okMessage); return; }
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
            setStatus(okMessage);
        }).catch(function () {
            setStatus('Kopieren blockiert - bitte "Alles markieren" und Strg+C benutzen.', true);
        });
        return;
    }
    setStatus('Kopieren blockiert - bitte "Alles markieren" und Strg+C benutzen.', true);
}

function copyCode() {
    var ta = document.getElementById('code');
    var btn = document.getElementById('btnCopy');
    copyText(ta.value, ta.value.length + ' Zeichen HTA-Code in der Zwischenablage.');
    btn.classList.add('copied');
    btn.innerHTML = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> Kopiert!';
    setTimeout(function () {
        btn.classList.remove('copied');
        btn.innerHTML = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Code kopieren';
        setStatus('');
    }, 3000);
}

function downloadCode() {
    var content = document.getElementById('code').value;
    var blob = new Blob([content], { type: 'application/octet-stream' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'outlook-regeln-visualisierung.hta';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    setStatus('Datei erzeugt. Falls der Browser .hta blockiert: "Code kopieren" benutzen.');
}

function selectAll() {
    var ta = document.getElementById('code');
    ta.focus();
    ta.select();
    setStatus('Alles markiert - jetzt Strg+C druecken.');
}

(function () {
    var code = document.getElementById('code').value;
    document.getElementById('meta-size').innerText = Math.round(code.length / 1024) + ' KB';
    document.getElementById('meta-lines').innerText = code.split('\n').length;
})();
</script>
</body>
</html>
FOOTER_END

SIZE=$(stat -c%s "$OUT")
echo "OK - $OUT gebaut (HTA v$VER, $SIZE bytes)"
