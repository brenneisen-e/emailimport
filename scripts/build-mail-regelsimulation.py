#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simuliert die von Marcus vorgeschlagenen Nummern-Regeln auf dem April-Extrakt
und baut daraus Auswertung (.xlsx) und Antwortmail (.eml).

Regeln laut Vorschlag:
  1. Nummern wie bisher von der KI auslesen
  2. Praefixe (A, V, P, ...), fuehrende Nullen und Trennzeichen entfernen
  3. alles < 7 Stellen -> Vermittlernummer
  4. alles mit 7 oder 9 Stellen -> Agenturnummer

Aufruf:  python3 scripts/build-mail-regelsimulation.py
"""

import collections
import re

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from mailbau import baue_eml, tab

QUELLE = "analyse.xlsx"
EXCEL = "Regelsimulation-Nummern.xlsx"
EML = "Regelsimulation-Nummern.eml"
BETREFF = "Nummern-Regeln am April-Extrakt durchgerechnet"

FONT = "Aptos"
BLAU = "FF1E40AF"
LILA = "FF4C1D95"
GELB = "FFFEF3C7"
ROT = "FFFEE2E2"
WEISS = "FFFFFFFF"


# ---------------------------------------------------------------------------
# Daten und Regelwerk
# ---------------------------------------------------------------------------
def lade():
    wb = openpyxl.load_workbook(QUELLE, read_only=True, data_only=True)
    zeilen = [[("" if v is None else str(v).strip()) for v in r]
              for r in wb["Analyse"].iter_rows(min_row=3, values_only=True)]
    wb.close()
    return [r for r in zeilen if r[21] in ("BÜ-Vorgang", "BUe-Vorgang")]


def kandidaten(r):
    """Alle extrahierten Nummern eines Vorgangs, je mit ihrer heutigen Spalte."""
    out = []
    for feld, heute in ((r[13], "Agentur"), (r[14], "Vermittler")):
        if not feld:
            continue
        for teil in re.split(r"[,;]| und ", feld):
            teil = teil.strip()
            if re.search(r"\d", teil):
                out.append((teil, heute))
    return out


def bereinigt(v):
    """Regel 2: nur Ziffern behalten, fuehrende Nullen streichen."""
    return re.sub(r"\D", "", v).lstrip("0")


def neue_klasse(z):
    """Regeln 3 und 4."""
    n = len(z)
    if n == 0:
        return "leer"
    if n < 7:
        return "Vermittler"
    if n in (7, 9):
        return "Agentur"
    return "ohne Zuordnung"


BUE = lade()
N = len(BUE)

wechsel = collections.Counter()
wechsler = collections.Counter()        # V -> A, mit Rohwert
ohne_zuordnung = collections.Counter()
achtsteller_mit_null = collections.Counter()
achtsteller_echt = collections.Counter()
roh_formen, norm_formen = set(), set()
vorher = collections.Counter()
nachher = collections.Counter()

for r in BUE:
    vorher["beide" if r[13] and r[14] else "nur Agentur" if r[13]
           else "nur Vermittler" if r[14] else "keine"] += 1
    hat_a = hat_v = False
    for roh, heute in kandidaten(r):
        z = bereinigt(roh)
        neu = neue_klasse(z)
        wechsel[(heute, neu)] += 1
        roh_formen.add(roh)
        norm_formen.add(z)
        if heute == "Vermittler" and neu == "Agentur":
            wechsler[(roh, z)] += 1
        if neu == "ohne Zuordnung":
            ohne_zuordnung[(roh, z, len(z))] += 1
            if len(z) == 8:
                (achtsteller_mit_null if re.sub(r"\D", "", roh).startswith("0")
                 else achtsteller_echt)[roh] += 1
        hat_a = hat_a or neu == "Agentur"
        hat_v = hat_v or neu == "Vermittler"
    nachher["beide" if hat_a and hat_v else "nur Agentur" if hat_a
            else "nur Vermittler" if hat_v else "keine"] += 1

WERTE = sum(wechsel.values())
N_WECHSEL = wechsel[("Vermittler", "Agentur")]
N_OHNE = sum(v for (h, n), v in wechsel.items() if n == "ohne Zuordnung")
N_MIT_NULL = sum(achtsteller_mit_null.values())
N_ECHT8 = sum(achtsteller_echt.values())


def pz(zaehler, nenner):
    return ("%.1f" % (100.0 * zaehler / nenner)).replace(".", ",") + " %"


def tsd(n):
    return "{:,}".format(int(n)).replace(",", ".")


def umbruch_txt(text, breite=76):
    worte, zeilen, akt = text.split(), [], ""
    for w in worte:
        if akt and len(akt) + 1 + len(w) > breite:
            zeilen.append(akt)
            akt = w
        else:
            akt = (akt + " " + w).strip()
    if akt:
        zeilen.append(akt)
    return "\n".join(zeilen)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def kopf(ws, spalten, breiten, zeile=1):
    for i, (t, b) in enumerate(zip(spalten, breiten), 1):
        z = ws.cell(row=zeile, column=i, value=t)
        z.font = Font(name=FONT, sz=10, b=True, color=WEISS)
        z.fill = PatternFill("solid", fgColor=LILA)
        z.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = b


def zeile_schreiben(ws, r, werte, fill=None):
    for i, w in enumerate(werte, 1):
        z = ws.cell(row=r, column=i, value=w)
        z.font = Font(name=FONT, sz=10)
        z.alignment = Alignment(vertical="top", wrap_text=True)
        if fill:
            z.fill = PatternFill("solid", fgColor=fill)


def titelzeile(ws, text, unter=""):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, sz=14, b=True, color=BLAU)
    if unter:
        ws["A2"] = unter
        ws["A2"].font = Font(name=FONT, sz=10, color="FF4B5563")
    ws.sheet_view.showGridLines = False


wb = openpyxl.Workbook()

ws = wb.active
ws.title = "Vergleich"
titelzeile(ws, "Nummern-Regeln: heute gegen Vorschlag",
           "Basis: %s BÜ-Vorgänge aus dem April-Extrakt mit %s extrahierten Nummern."
           % (tsd(N), tsd(WERTE)))
kopf(ws, ["Befüllung je Vorgang", "heute", "nach Vorschlag", "Veränderung"],
     [30, 14, 16, 14], zeile=4)
r = 5
for k in ("beide", "nur Agentur", "nur Vermittler", "keine"):
    zeile_schreiben(ws, r, [k, vorher[k], nachher[k], nachher[k] - vorher[k]],
                    fill=ROT if k == "keine" else None)
    r += 1
zeile_schreiben(ws, r, ["Summe", sum(vorher.values()), sum(nachher.values()), 0])
r += 3
ws.cell(row=r, column=1, value="Wechsel der Zuordnung je Nummer").font = Font(
    name=FONT, sz=11, b=True, color=BLAU)
r += 1
kopf(ws, ["heute", "nach Vorschlag", "Nummern", "Anteil"], [22, 22, 12, 12], zeile=r)
r += 1
for (h, n), c in sorted(wechsel.items(), key=lambda x: -x[1]):
    fill = GELB if h != n and n in ("Agentur", "Vermittler") else (
        ROT if n == "ohne Zuordnung" else None)
    zeile_schreiben(ws, r, [h, n, c, pz(c, WERTE)], fill=fill)
    r += 1

ws2 = wb.create_sheet("Wechsler V nach A")
titelzeile(ws2, "Nummern, die neu als Agenturnummer gelten",
           "Heute Vermittlernummer, nach den neuen Regeln 7- oder 9-stellig und damit "
           "Agenturnummer. %s Werte in %s verschiedenen Schreibweisen."
           % (tsd(N_WECHSEL), tsd(len(wechsler))))
kopf(ws2, ["Rohwert heute", "bereinigt", "Stellen", "Vorgänge", "Präfix 890/6000?"],
     [26, 20, 10, 12, 18], zeile=4)
r = 5
for (roh, z), c in wechsler.most_common():
    passt = "ja" if z.startswith(("890", "6000")) else "nein"
    zeile_schreiben(ws2, r, [roh, z, len(z), c, passt], fill=GELB if passt == "nein" else None)
    r += 1

ws3 = wb.create_sheet("Ohne Zuordnung")
titelzeile(ws3, "Nummern, die durch das Raster fallen",
           "Weder unter 7 Stellen noch genau 7 oder 9 Stellen - %s Werte." % tsd(N_OHNE))
kopf(ws3, ["Rohwert", "bereinigt", "Stellen", "Vorgänge", "war im Original 9-stellig"],
     [26, 20, 10, 12, 24], zeile=4)
r = 5
for (roh, z, laenge), c in sorted(ohne_zuordnung.items(), key=lambda x: (-x[1], x[0][2])):
    war9 = "ja" if laenge == 8 and re.sub(r"\D", "", roh).startswith("0") else ""
    zeile_schreiben(ws3, r, [roh, z, laenge, c, war9], fill=GELB if war9 else ROT)
    r += 1

wb.calculation.fullCalcOnLoad = True
wb.save(EXCEL)


# ---------------------------------------------------------------------------
# Mail
# ---------------------------------------------------------------------------
top_wechsler = ", ".join("<code>%s</code> (%dx)" % (roh, c)
                         for (roh, z), c in wechsler.most_common(6))

HTML = """<div>
<p>Hallo Marcus,</p>

<p>danke für den Vorschlag - die ersten vier Regeln habe ich über den
April-Extrakt laufen lassen. Basis sind die {N} BÜ-Vorgänge mit
{WERTE} extrahierten Nummern. Die Auswertung hängt als Excel an, mit einem
Blatt je Befund.</p>

<p>Kurzfassung: Die Regeln greifen gut und bringen bei der Agenturnummer einen
deutlichen Zugewinn. Zwei Stellen brauchen aber noch eine Entscheidung, bevor
wir das produktiv setzen.</p>

<h3>1. Befüllung je Vorgang</h3>
{TAB_BEF}
<p>Der Weg über die Agenturnummer - also euer Einstieg in EMMA - deckt statt
{AG_ALT} künftig {AG_NEU} Vorgänge ab, das sind {AG_PZ} statt {AG_PZ_ALT}. Der
Rückfallweg über die Vermittlernummer schrumpft entsprechend von 624 auf 104
Vorgänge.</p>

<p><b>Warum &bdquo;keine&ldquo; um 117 steigt:</b> Das sind genau die
{N_OHNE} Nummern, die weder unter 7 noch genau 7 oder 9 Stellen haben und
deshalb in keiner der beiden Spalten landen (Details in Abschnitt 3). Der
Punkt dabei: <b>In jedem dieser 117 Vorgänge ist das die einzige vorhandene
Nummer</b> - keiner von ihnen hat noch eine zweite, verwertbare. Jeder Ausfall
schlägt deshalb eins zu eins durch. Die häufigsten Fälle:</p>
{TAB_AUSFALL}
<p class="hint">Auffällig ist, dass diese Werte fast ausschließlich aus einer
Nummernfamilie stammen: <code>10006515</code>, <code>10006776</code>,
<code>10003083</code> und ähnliche - achtstellig, oft mit vorangestelltem
<code>A0</code> oder <code>0</code>. Genau bei dieser Gruppe müssen wir uns
also festlegen.</p>

<h3>2. Was sich je Nummer ändert</h3>
{TAB_WECHSEL}
<p>Die Agenturnummern bleiben praktisch unangetastet: {A_STABIL} von
{A_GESAMT} behalten ihre Zuordnung. Die Bewegung steckt bei den
Vermittlernummern - <b>{N_WECHSEL} Werte ({N_WECHSEL_PZ} aller Nummern) werden
neu zur Agenturnummer</b>. Häufigste Fälle: {TOP_WECHSLER}.</p>
<p>Das ist der eigentliche Unterschied zur heutigen Logik: Die bisherige Regel
&bdquo;beginnt mit 6000 oder 890&ldquo; und die neue Längenregel widersprechen
sich bei genau diesen Werten. <code>8923555</code> etwa beginnt mit 89, aber
nicht mit 890; <code>601000040</code> beginnt mit 601, nicht mit 6000. Nach der
Längenregel sind beide Agenturnummern. Ob das fachlich richtig ist, kann ich
aus den Daten nicht beurteilen - der EMMA-Abgleich würde es aber genau hier
zeigen.</p>

<h3>3. Regelungslücke bei 8 Stellen</h3>
<p>{N_OHNE} Werte fallen durch das Raster: Sie sind weder unter 7 Stellen noch
genau 7 oder 9 Stellen lang. Dadurch steigt die Zahl der Vorgänge ganz ohne
Nummer von 274 auf 391.</p>
<p><b>Der Grund ist eine Wechselwirkung zwischen Regel 2 und Regel 4:</b> Von
den {N8} achtstelligen Werten hatten <b>{N_MIT_NULL} im Original eine führende
Null</b> und waren damit neunstellig - etwa <code>010003083</code> oder
<code>01000/3451</code>. Regel 2 streicht die Null, danach greift Regel 4 nicht
mehr.</p>
<p>Ein Sonderfall noch: Bei Werten mit Textrest wie
<code>600041966 1zu1 AG</code> zieht die Bereinigung auch die Ziffern aus dem
Text mit - übrig bleibt <code>60004196611</code> mit 11 Stellen statt der
richtigen 9. Hier müsste vor dem Zusammenziehen erst der Textanteil entfernt
werden.</p>

<p>Vorschlag: <b>Reihenfolge tauschen</b> - erst Präfixe und Trennzeichen
entfernen, dann die Länge prüfen, und die führende Null erst danach streichen.
Damit landen diese {N_MIT_NULL} Werte wieder korrekt bei der Agenturnummer.
Übrig blieben {N_ECHT8} echte Achtsteller wie <code>10002866</code> oder
<code>89420798</code> - für die bräuchten wir noch eine Festlegung.</p>

<h3>4. Nebeneffekt: einheitlichere Schreibweisen</h3>
<p>Die Bereinigung löst die Dubletten aus der Formatanalyse auf:
<code>60008-1364</code> und <code>600081364</code> sind danach dieselbe Nummer,
ebenso <code>008923555</code> und <code>8923555</code>. Aus {ROH}
verschiedenen Schreibweisen werden {NORM} eindeutige Nummern.</p>

<h3>Was ich von dir bräuchte</h3>
<ol>
<li><b>Achtstellige Nummern:</b> Passt der Vorschlag, die führende Null erst
nach der Längenprüfung zu streichen? Und wohin gehören die {N_ECHT8} echten
Achtsteller - Vermittlernummer, Agenturnummer oder &bdquo;zu prüfen&ldquo;?</li>
<li><b>Die {N_WECHSEL} Wechsler:</b> Im Blatt
&bdquo;Wechsler V nach A&ldquo; stehen sie einzeln mit Häufigkeit. Magst du
stichprobenhaft draufschauen, ob die Einordnung als Agenturnummer passt? Falls
ja, ist die Längenregel der heutigen Präfix-Regel klar überlegen.</li>
<li><b>Werte über 9 Stellen:</b> {N_LANG} Werte sind länger als 9 Stellen,
meist zwei zusammengezogene Nummern. Direkt als &bdquo;zu prüfen&ldquo;
kennzeichnen?</li>
</ol>

<p>Den EMMA-Teil deines Vorschlags - Agenturnummer als führend, Vermittlernummer
ergänzen oder überschreiben, bei Verstößen &bdquo;zu prüfen&ldquo; - können wir
genau so umsetzen. Sobald die beiden Punkte oben geklärt sind, ziehe ich die
Regeln in die Nachverarbeitung ein und lasse den April-Extrakt noch einmal
komplett durchlaufen.</p>

<p>Rückfragen gerne jederzeit.</p>

<p>Viele Grüße</p>
</div>"""

HTML = HTML.replace("{TAB_BEF}", tab(
    ["Befüllung je Vorgang", "heute", "nach Vorschlag", "Veränderung"],
    [[k, tsd(vorher[k]), tsd(nachher[k]),
      ("%+d" % (nachher[k] - vorher[k])) if nachher[k] != vorher[k] else "0"]
     for k in ("beide", "nur Agentur", "nur Vermittler", "keine")]
    + [["<b>Summe</b>", "<b>%s</b>" % tsd(N), "<b>%s</b>" % tsd(N), "<b>0</b>"]],
    num_spalten=(1, 2, 3)))

HTML = HTML.replace("{TAB_AUSFALL}", tab(
    ["Rohwert heute", "bereinigt", "Stellen", "Vorgänge"],
    [[("<code>%s</code>" % roh), ("<code>%s</code>" % z), str(laenge), str(c)]
     for (roh, z, laenge), c in sorted(ohne_zuordnung.items(),
                                       key=lambda x: -x[1])[:6]],
    num_spalten=(2, 3)))

HTML = HTML.replace("{TAB_WECHSEL}", tab(
    ["heute", "nach Vorschlag", "Nummern", "Anteil"],
    [[h, n, tsd(c), pz(c, WERTE)]
     for (h, n), c in sorted(wechsel.items(), key=lambda x: -x[1])],
    num_spalten=(2, 3)))

A_STABIL = wechsel[("Agentur", "Agentur")]
A_GESAMT = sum(c for (h, _), c in wechsel.items() if h == "Agentur")
N_LANG = sum(c for (roh, z, l), c in ohne_zuordnung.items() if l > 9)

for platz, wert in (
        ("{N}", tsd(N)), ("{WERTE}", tsd(WERTE)),
        ("{AG_ALT}", tsd(vorher["beide"] + vorher["nur Agentur"])),
        ("{AG_NEU}", tsd(nachher["beide"] + nachher["nur Agentur"])),
        ("{AG_PZ}", pz(nachher["beide"] + nachher["nur Agentur"], N)),
        ("{AG_PZ_ALT}", pz(vorher["beide"] + vorher["nur Agentur"], N)),
        ("{A_STABIL}", tsd(A_STABIL)), ("{A_GESAMT}", tsd(A_GESAMT)),
        ("{N_WECHSEL_PZ}", pz(N_WECHSEL, WERTE)), ("{N_WECHSEL}", tsd(N_WECHSEL)),
        ("{TOP_WECHSLER}", top_wechsler),
        ("{N_OHNE}", tsd(N_OHNE)), ("{N8}", tsd(N_MIT_NULL + N_ECHT8)),
        ("{N_MIT_NULL}", tsd(N_MIT_NULL)), ("{N_ECHT8}", tsd(N_ECHT8)),
        ("{N_LANG}", tsd(N_LANG)),
        ("{ROH}", tsd(len(roh_formen))), ("{NORM}", tsd(len(norm_formen)))):
    HTML = HTML.replace(platz, wert)

offen = re.findall(r"\{[A-Z_0-9]+\}", HTML)
assert not offen, offen

# ---------------------------------------------------------------------------
# Klartext: HTML-Absaetze uebernehmen, Tabellen als ASCII neu setzen
# ---------------------------------------------------------------------------
def klartext(html):
    """Absaetze und Listen in Klartext wandeln, Tabellen als Marker ausschneiden."""
    t = html
    t = re.sub(r"<h3>(.*?)</h3>", lambda m: "\n@@H@@" + m.group(1) + "\n", t)
    t = re.sub(r"<li>(.*?)</li>", lambda m: "  - " + m.group(1) + "\n", t)
    t = re.sub(r"<table>.*?</table>", "@@TAB@@", t, flags=re.S)
    t = re.sub(r"</p>", "\n\n", t)
    t = re.sub(r"<[^>]+>", "", t)
    for a, b in (("&bdquo;", "„"), ("&ldquo;", "“"), ("&nbsp;", " "), ("&amp;", "&")):
        t = t.replace(a, b)
    return t


def ascii_tab(kopfzeilen, zeilen, breiten):
    aus = ["  " + "".join(("%-" + str(b) + "s") % k for k, b in zip(kopfzeilen, breiten)),
           "  " + "-" * (sum(breiten) - 2)]
    for zl in zeilen:
        aus.append("  " + "".join(("%-" + str(b) + "s") % w for w, b in zip(zl, breiten)))
    return "\n".join(aus)


AUSFALL_TOP = sorted(ohne_zuordnung.items(), key=lambda x: -x[1])[:6]

TAB_TEXT = [
    ascii_tab(["Befüllung je Vorgang", "heute", "Vorschlag", "Änderung"],
              [[k, tsd(vorher[k]), tsd(nachher[k]),
                ("%+d" % (nachher[k] - vorher[k])) if nachher[k] != vorher[k] else "0"]
               for k in ("beide", "nur Agentur", "nur Vermittler", "keine")]
              + [["Summe", tsd(N), tsd(N), "0"]], [24, 10, 12, 10]),
    ascii_tab(["Rohwert heute", "bereinigt", "Stellen", "Vorgänge"],
              [[roh, z, str(laenge), str(c)]
               for (roh, z, laenge), c in AUSFALL_TOP], [24, 16, 10, 10]),
    ascii_tab(["heute", "nach Vorschlag", "Nummern", "Anteil"],
              [[h, n, tsd(c), pz(c, WERTE)]
               for (h, n), c in sorted(wechsel.items(), key=lambda x: -x[1])],
              [14, 18, 10, 10]),
]

_t = klartext(HTML)
for tabelle in TAB_TEXT:
    _t = _t.replace("@@TAB@@", "\n" + tabelle + "\n", 1)
zeilen_txt = []
for zeile in _t.split("\n"):
    zeile = zeile.rstrip()
    if zeile.startswith("@@H@@"):
        titel = zeile[5:]
        zeilen_txt += ["", titel, "-" * len(titel)]
    elif zeile.startswith("  ") or not zeile.strip():
        zeilen_txt.append(zeile)
    else:
        zeilen_txt.append(umbruch_txt(zeile))
TEXT = re.sub(r"\n{3,}", "\n\n", "\n".join(zeilen_txt)).strip() + "\n"

baue_eml(EML, BETREFF, TEXT, HTML, [EXCEL])
print("   %s Vorgänge, %s Nummern | Wechsler V->A: %s | ohne Zuordnung: %s"
      % (tsd(N), tsd(WERTE), tsd(N_WECHSEL), tsd(N_OHNE)))
