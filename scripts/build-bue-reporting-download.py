#!/usr/bin/env python3
"""Baut bue-reporting-download.html aus bue-reporting/BUE_Reporting_Master.xlsm.

Warum: Direkte Downloads (.xlsm wie .zip) scheitern in Firmennetzen mit "Keine
Berechtigungen" - der Proxy bzw. die Browser-Richtlinie blockt Dateien, die als
Datei vom Server kommen. Wie bei den HTA-Quellcodeseiten steckt die Datei deshalb
als Base64 in der Seite selbst; ein Klick setzt sie im Browser zusammen und
speichert sie per Blob (der Netzwerk-Request ist nur die HTML-Seite).
Drei Knoepfe: als .xlsm, als .xlsx (gleicher Inhalt, danach von Hand in .xlsm umbenennen - fuer Netze,
die .xlsm nach Endung blocken), als .zip und als .zip mit der .xlsx darin.

Aufruf:   python3 scripts/build-bue-reporting-download.py            (HS: BUE_Reporting_Master.xlsm)
          python3 scripts/build-bue-reporting-download.py makler     (Makler: BUE_Reporting_Master_Makler.xlsm)
Nach JEDEM Austausch der Master-Datei ausfuehren.
"""
import base64
import datetime as dt
import io
import os
import sys
import zipfile

HIER = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAKLER = len(sys.argv) > 1 and sys.argv[1].lower() == 'makler'
NAME = 'BUE_Reporting_Master_Makler.xlsm' if MAKLER else 'BUE_Reporting_Master.xlsm'
QUELLE = os.path.join(HIER, 'bue-reporting', NAME)
ZIEL = os.path.join(HIER, 'bue-reporting-makler-download.html' if MAKLER else 'bue-reporting-download.html')
BASIS = NAME[:-5]                                               # Dateiname ohne .xlsm
TITEL = 'BÜ-Reporting Makler – Master-Datei' if MAKLER else 'BÜ-Reporting HS – Master-Datei'
ANDERE = ('<a href="bue-reporting-download.html">Zur Fassung HS (HSB4 K)</a>' if MAKLER
          else '<a href="bue-reporting-makler-download.html">Zur Fassung Makler (ODPB, ODPP, KBI …)</a>')

daten = open(QUELLE, 'rb').read()


def als_zip(name):
    puffer = io.BytesIO()
    with zipfile.ZipFile(puffer, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr(name, daten)                                    # nur die Master-Datei (Ordner hat der Anwender)
    return puffer.getvalue()


zip_daten = als_zip(NAME)
zip_xlsx_daten = als_zip(NAME.replace('.xlsm', '.xlsx'))        # Variante fuer Netze, die .xlsm auch im ZIP pruefen


def b64(b):
    s = base64.b64encode(b).decode('ascii')
    return '\n'.join(s[i:i + 120] for i in range(0, len(s), 120))


from zoneinfo import ZoneInfo
stand = dt.datetime.fromtimestamp(os.path.getmtime(QUELLE), ZoneInfo('Europe/Berlin')).strftime('%d.%m.%Y, %H:%M Uhr')   # deutsche Zeit
groesse = ('%.1f MB' % (len(daten) / 1024 / 1024)).replace('.', ',')

seite = """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>__TITEL__ herunterladen</title>
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
  <p><a href="downloads.html">← Zur Tool-Bibliothek</a> · __ANDERE__</p>
  <div class="card">
    <h1>__TITEL__</h1>
    <div class="sub">__BASIS__.xlsm · Excel mit Makros · __GROESSE__ · Stand __STAND__</div>
    <div class="btns">
      <button class="prim" onclick="speichern('xlsm')">Als .xlsm speichern</button>
      <button class="sek" onclick="speichern('xlsx')">Als .xlsx speichern</button>
      <button class="sek" onclick="speichern('zip')">Als ZIP speichern</button>
      <button class="sek" onclick="speichern('zipx')">ZIP mit .xlsx</button>
    </div>
    <div class="hinweis" style="margin-top:4px">„Als .xlsx speichern“ ist derselbe Inhalt unter anderer Endung – für Netze, die .xlsm blocken.
      Danach die Datei im Explorer von <b>__BASIS__.xlsx</b> in <b>__BASIS__.xlsm</b> umbenennen
      (Dateiendungen einblenden: Explorer → Ansicht → „Dateinamenerweiterungen“). Mit der Endung .xlsx öffnet Excel die Datei nicht.
      „ZIP mit .xlsx“ enthält dieselbe .xlsx – nach dem Entpacken ebenso umbenennen.</div>
    <div id="status"></div>
    <ol>
      <li>Datei in den Ordner legen, in dem „Rohdaten“ (MIS-O-Abzüge OPAG und Verbund) und „Anwesenheit“ (MAK-Liste) liegen (bei der ZIP: dorthin entpacken).</li>
      <li>Datei öffnen, „Inhalt aktivieren“ klicken, im Cockpit „Neue Daten laden“.</li>
    </ol>
    <div class="hinweis">Die Datei steckt vollständig in dieser Seite und wird beim Klick im Browser erzeugt – es wird keine
      Datei vom Server geladen. Zeigt Excel beim Öffnen „Makros wurden blockiert“: Datei rechts anklicken → Eigenschaften →
      „Zulassen“ ankreuzen → OK, dann erneut öffnen. Ausführliche Anleitung im Blatt „Anleitung“ der Datei.</div>
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
<script>
var TYPEN = {
  xlsm: { id: 'daten-xlsm', name: '__BASIS__.xlsm', mime: 'application/vnd.ms-excel.sheet.macroEnabled.12' },
  xlsx: { id: 'daten-xlsm', name: '__BASIS__.xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          nachher: ' – jetzt im Explorer in __BASIS__.xlsm umbenennen.' },
  zip:  { id: 'daten-zip',  name: '__BASIS__.zip',  mime: 'application/zip' },
  zipx: { id: 'daten-zipx', name: '__BASIS___xlsx.zip', mime: 'application/zip',
          nachher: ' – entpacken und __BASIS__.xlsx in __BASIS__.xlsm umbenennen.' }
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
seite = (seite.replace('__TITEL__', TITEL).replace('__ANDERE__', ANDERE).replace('__BASIS__', BASIS).replace('__GROESSE__', groesse).replace('__STAND__', stand)
         .replace('__XLSM__', b64(daten)).replace('__ZIPX__', b64(zip_xlsx_daten)).replace('__ZIP__', b64(zip_daten)))
open(ZIEL, 'w', encoding='utf-8').write(seite)
print('geschrieben:', ZIEL, '%.1f MB' % (len(seite) / 1024 / 1024))
