#!/usr/bin/env python3
"""Baut hr-sales-dashboard-download.html aus hr-sales-dashboard/HR_Sales_Dashboard.xlsm (+ Vorlage/ERGO_Vorlage.pptx).

Warum: Direkte Downloads (.xlsm wie .zip) scheitern in Firmennetzen mit "Keine
Berechtigungen" - der Proxy bzw. die Browser-Richtlinie blockt Dateien, die als
Datei vom Server kommen. Wie bei den HTA-Quellcodeseiten steckt die Datei deshalb
als Base64 in der Seite selbst; ein Klick setzt sie im Browser zusammen und
speichert sie per Blob (der Netzwerk-Request ist nur die HTML-Seite).
Drei Knoepfe: als .xlsm, als .xlsx (gleicher Inhalt, danach von Hand in .xlsm umbenennen - fuer Netze,
die .xlsm nach Endung blocken), als .zip und als .zip mit der .xlsx darin.

Die ZIP-Varianten enthalten zusaetzlich den Ordner Vorlage mit ERGO_Vorlage.pptx (fuer „PowerPoint erzeugen“).

Aufruf:   python3 scripts/build-hr-sales-dashboard-download.py
Nach JEDEM Austausch von hr-sales-dashboard/HR_Sales_Dashboard.xlsm ausfuehren.
"""
import base64
import datetime as dt
import io
import os
import zipfile

HIER = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QUELLE = os.path.join(HIER, 'hr-sales-dashboard', 'HR_Sales_Dashboard.xlsm')
VORLAGE = os.path.join(HIER, 'hr-sales-dashboard', 'Vorlage', 'ERGO_Vorlage.pptx')
ZIEL = os.path.join(HIER, 'hr-sales-dashboard-download.html')
NAME = 'HR_Sales_Dashboard.xlsm'

daten = open(QUELLE, 'rb').read()
vorlage = open(VORLAGE, 'rb').read()


def als_zip(name):
    puffer = io.BytesIO()
    with zipfile.ZipFile(puffer, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('HR_Sales_Dashboard/' + name, daten)
        z.writestr('HR_Sales_Dashboard/Vorlage/ERGO_Vorlage.pptx', vorlage)
    return puffer.getvalue()


zip_daten = als_zip(NAME)
zip_xlsx_daten = als_zip(NAME.replace('.xlsm', '.xlsx'))        # Variante fuer Netze, die .xlsm auch im ZIP pruefen


def b64(b):
    s = base64.b64encode(b).decode('ascii')
    return '\n'.join(s[i:i + 120] for i in range(0, len(s), 120))


stand = dt.datetime.fromtimestamp(os.path.getmtime(QUELLE)).strftime('%d.%m.%Y')
groesse = ('%.1f MB' % (len(daten) / 1024 / 1024)).replace('.', ',')

seite = """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>HR Sales Dashboard – Datei herunterladen</title>
<style>
  body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; background: #f5f6f8; color: #1a1a1a; }
  .wrap { max-width: 760px; margin: 48px auto; padding: 0 16px; }
  .card { background: #fff; border: 1px solid #e1dfdd; border-radius: 12px; padding: 28px 30px; box-shadow: 0 1px 2px rgba(0,0,0,.04); }
  h1 { font-size: 22px; margin: 0 0 6px; }
  .sub { color: #6b6b6b; font-size: 13px; margin-bottom: 22px; }
  .btns { display: flex; gap: 12px; flex-wrap: wrap; margin: 18px 0 10px; }
  button { font: inherit; font-size: 14px; font-weight: 600; border: 0; border-radius: 8px; padding: 11px 18px; cursor: pointer; }
  .prim { background: #4f46e5; color: #fff; }
  .sek { background: #fff; color: #333; border: 1px solid #d0d5da; }
  #status { min-height: 20px; font-size: 13px; color: #1e7f4f; margin-top: 6px; }
  ol { padding-left: 20px; line-height: 1.6; font-size: 14px; }
  .hinweis { font-size: 12.5px; color: #6b6b6b; margin-top: 18px; line-height: 1.5; }
  a { color: #4f46e5; }
</style>
</head>
<body>
<div class="wrap">
  <p><a href="downloads.html">← Zur Tool-Bibliothek</a></p>
  <div class="card">
    <h1>HR Sales Dashboard</h1>
    <div class="sub">HR_Sales_Dashboard.xlsm · Excel mit Makros · __GROESSE__ · Stand __STAND__</div>
    <div class="btns">
      <button class="prim" onclick="speichern('xlsm')">Als .xlsm speichern</button>
      <button class="sek" onclick="speichern('xlsx')">Als .xlsx speichern</button>
      <button class="sek" onclick="speichern('zip')">Als ZIP speichern</button>
      <button class="sek" onclick="speichern('zipx')">ZIP mit .xlsx</button>
      <button class="sek" onclick="speichern('pptx')">Nur Vorlage (.pptx)</button>
    </div>
    <div class="hinweis" style="margin-top:4px">„Als .xlsx speichern“ ist derselbe Inhalt unter anderer Endung – für Netze, die .xlsm blocken.
      Danach die Datei im Explorer von <b>HR_Sales_Dashboard.xlsx</b> in <b>HR_Sales_Dashboard.xlsm</b> umbenennen
      (Dateiendungen einblenden: Explorer → Ansicht → „Dateinamenerweiterungen“). Mit der Endung .xlsx öffnet Excel die Datei nicht.
      „ZIP mit .xlsx“ enthält dieselbe .xlsx – nach dem Entpacken ebenso umbenennen.</div>
    <div id="status"></div>
    <ol>
      <li>Die ZIP entpacken: Ordner <b>HR_Sales_Dashboard</b> mit der Datei und dem Ordner <b>Vorlage</b> (ERGO_Vorlage.pptx für „PowerPoint erzeugen“).
        Wer nur die .xlsm speichert, legt die Vorlage selbst in einen Ordner „Vorlage“ neben die Datei.</li>
      <li>Datei öffnen, „Inhalt aktivieren“ klicken, auf der Startseite „Start · Werte eintragen“.</li>
      <li>BÜ-Werte: einmal „BÜ-Reporting verknüpfen“ und die BUE_Reporting_Master.xlsm wählen.</li>
    </ol>
    <div class="hinweis">Die Datei steckt vollständig in dieser Seite und wird beim Klick im Browser erzeugt – es wird keine
      Datei vom Server geladen. Zeigt Excel beim Öffnen „Makros wurden blockiert“: Datei rechts anklicken → Eigenschaften →
      „Zulassen“ ankreuzen → OK, dann erneut öffnen.</div>
  </div>
</div>
<script type="text/plain" id="daten-xlsm">
__XLSM__
</script>
<script type="text/plain" id="daten-zip">
__ZIP__
</script>
<script type="text/plain" id="daten-zipx">
__ZIPX__
</script>
<script type="text/plain" id="daten-pptx">
__PPTX__
</script>
<script>
var TYPEN = {
  xlsm: { id: 'daten-xlsm', name: 'HR_Sales_Dashboard.xlsm', mime: 'application/vnd.ms-excel.sheet.macroEnabled.12' },
  xlsx: { id: 'daten-xlsm', name: 'HR_Sales_Dashboard.xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          nachher: ' – jetzt im Explorer in HR_Sales_Dashboard.xlsm umbenennen.' },
  zip:  { id: 'daten-zip',  name: 'HR_Sales_Dashboard.zip',  mime: 'application/zip' },
  zipx: { id: 'daten-zipx', name: 'HR_Sales_Dashboard_xlsx.zip', mime: 'application/zip',
          nachher: ' – entpacken und HR_Sales_Dashboard.xlsx in HR_Sales_Dashboard.xlsm umbenennen.' },
  pptx: { id: 'daten-pptx', name: 'ERGO_Vorlage.pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          nachher: ' – in den Ordner „Vorlage“ neben der Datei legen.' }
};
function speichern(art) {
  var t = TYPEN[art], status = document.getElementById('status');
  try {
    var b64 = document.getElementById(t.id).textContent.replace(/\\s+/g, '');
    var bin = atob(b64), n = bin.length, bytes = new Uint8Array(n);
    for (var i = 0; i < n; i++) bytes[i] = bin.charCodeAt(i);
    var blob = new Blob([bytes], { type: t.mime });
    if (window.navigator && window.navigator.msSaveOrOpenBlob) { window.navigator.msSaveOrOpenBlob(blob, t.name); }
    else {
      var url = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = url; a.download = t.name; document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    }
    status.style.color = '#1e7f4f';
    status.textContent = t.name + ' erzeugt (' + Math.round(n / 1024) + ' KB)' + (t.nachher || '. Klappt es nicht, einen anderen Knopf versuchen.');
  } catch (e) {
    status.style.color = '#bf1528';
    status.textContent = 'Fehler beim Erzeugen: ' + e.message;
  }
}
</script>
</body>
</html>
"""
seite = (seite.replace('__GROESSE__', groesse).replace('__STAND__', stand)
         .replace('__XLSM__', b64(daten)).replace('__PPTX__', b64(vorlage)).replace('__ZIPX__', b64(zip_xlsx_daten)).replace('__ZIP__', b64(zip_daten)))
open(ZIEL, 'w', encoding='utf-8').write(seite)
print('geschrieben:', ZIEL, '%.1f MB' % (len(seite) / 1024 / 1024))
