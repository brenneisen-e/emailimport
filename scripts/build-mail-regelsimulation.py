#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simuliert die von Marcus vorgeschlagenen Nummern-Regeln auf dem April-Extrakt
und baut daraus Auswertung (.xlsx) und Antwortmail (.eml).

Drei Varianten werden gerechnet:
  A  Regeln wie vorgeschlagen: bereinigen (Praefixe, fuehrende Nullen,
     Trennzeichen), < 7 Stellen -> Vermittler, 7 oder 9 Stellen -> Agentur
  B  zusaetzlich: Felder, die zwei mit Label versehene Nummern enthalten
     ("V 85357 A 777-0503"), vorher auftrennen
  C  zusaetzlich: achtstellige Nummern mit einer fuehrenden Null auf neun
     Stellen ergaenzen und damit als Agenturnummer werten

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
GRUEN = "FFDCFCE7"
GELB = "FFFEF3C7"
ROT = "FFFEE2E2"
WEISS = "FFFFFFFF"

KATEGORIEN = ("beide", "nur Agentur", "nur Vermittler", "keine")

LABEL = re.compile(r"[A-Za-z]{1,2}\s*\d[\d\s\-\/]*")


# ---------------------------------------------------------------------------
# Regelwerk
# ---------------------------------------------------------------------------
def ziffern(v):
    return re.sub(r"\D", "", v)


def auftrennen(v):
    """Variante B: 'V 85357 A 777-0503' in zwei Nummern zerlegen.

    Nur wenn mindestens zwei mit Buchstaben-Label versehene Nummern gefunden
    werden UND diese den Wert praktisch vollstaendig abdecken. Damit bleiben
    Einzelnummern wie 'A600040124', '60002/0753' oder '00 111 40 16'
    unangetastet.
    """
    treffer = LABEL.findall(v)
    if len(treffer) < 2:
        return [v]
    abgedeckt = sum(len(re.sub(r"\s", "", t)) for t in treffer)
    gesamt = len(re.sub(r"\s", "", v)) or 1
    return [t.strip() for t in treffer] if abgedeckt / gesamt >= 0.8 else [v]


def kandidaten(r, mit_trennung):
    out = []
    for feld, heute_spalte in ((r[13], "Agentur"), (r[14], "Vermittler")):
        if not feld:
            continue
        for teil in re.split(r"[,;]| und ", feld):
            teil = teil.strip()
            if not re.search(r"\d", teil):
                continue
            for stueck in (auftrennen(teil) if mit_trennung else [teil]):
                out.append((stueck, heute_spalte))
    return out


def klasse(v, acht_auffuellen):
    """< 7 Stellen -> Vermittler, 7 oder 9 Stellen -> Agentur, sonst offen.

    acht_auffuellen: achtstellige Nummern bekommen eine fuehrende Null,
    sind damit neunstellig und gelten als Agenturnummer.
    """
    k = ziffern(v).lstrip("0")
    n = len(k)
    if n == 0:
        return None
    if acht_auffuellen and n == 8:
        return "Agentur"
    if n < 7:
        return "Vermittler"
    if n in (7, 9):
        return "Agentur"
    return "ohne Zuordnung"


def rechne(bue, mit_trennung, acht_auffuellen):
    befuellung = collections.Counter()
    offen = collections.Counter()
    wechsel = collections.Counter()
    for r in bue:
        klassen = []
        for roh, heute_spalte in kandidaten(r, mit_trennung):
            k = klasse(roh, acht_auffuellen)
            klassen.append(k)
            wechsel[(heute_spalte, k or "leer")] += 1
            if k == "ohne Zuordnung":
                rein = ziffern(roh).lstrip("0")
                offen[(roh, rein, len(rein))] += 1
        hat_a, hat_v = "Agentur" in klassen, "Vermittler" in klassen
        befuellung["beide" if hat_a and hat_v else "nur Agentur" if hat_a
                   else "nur Vermittler" if hat_v else "keine"] += 1
    return befuellung, offen, wechsel


# ---------------------------------------------------------------------------
# Daten
# ---------------------------------------------------------------------------
wb = openpyxl.load_workbook(QUELLE, read_only=True, data_only=True)
BUE = [[("" if v is None else str(v).strip()) for v in r]
       for r in wb["Analyse"].iter_rows(min_row=3, values_only=True)]
wb.close()
BUE = [r for r in BUE if r[21] in ("BÜ-Vorgang", "BUe-Vorgang")]
N = len(BUE)

heute = collections.Counter()
for r in BUE:
    heute["beide" if r[13] and r[14] else "nur Agentur" if r[13]
          else "nur Vermittler" if r[14] else "keine"] += 1

A_bef, A_offen, A_wechsel = rechne(BUE, False, False)
B_bef, B_offen, _ = rechne(BUE, True, False)
C_bef, C_offen, _ = rechne(BUE, True, True)

WERTE = sum(A_wechsel.values())
N_WECHSLER = A_wechsel[("Vermittler", "Agentur")]
wechsler_liste = collections.Counter()
for r in BUE:
    for roh, h in kandidaten(r, False):
        if h == "Vermittler" and klasse(roh, False) == "Agentur":
            wechsler_liste[(roh, ziffern(roh).lstrip("0"))] += 1

A_OFFEN, B_OFFEN, C_OFFEN = (sum(x.values()) for x in (A_offen, B_offen, C_offen))
ACHT = sum(c for (_, _, l), c in A_offen.items() if l == 8)
LANG = sum(c for (_, _, l), c in A_offen.items() if l > 9)
REST_ACHT = sum(c for (_, _, l), c in C_offen.items() if l == 8)


def pz(z, n):
    return ("%.1f" % (100.0 * z / n)).replace(".", ",") + " %"


def tsd(n):
    return "{:,}".format(int(n)).replace(",", ".")


def diff(neu, alt):
    return "%+d" % (neu - alt) if neu != alt else "0"


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


VARIANTEN = [
    ("heute (Präfix-Regel 6000/890)", heute, None, None),
    ("A - Regeln wie vorgeschlagen", A_bef, A_OFFEN, None),
    ("B - zusätzlich Felder mit zwei Nummern trennen", B_bef, B_OFFEN, GELB),
    ("C - zusätzlich achtstellige auf neun Stellen auffüllen", C_bef, C_OFFEN, GRUEN),
]

xl = openpyxl.Workbook()
ws = xl.active
ws.title = "Varianten"
titelzeile(ws, "Nummern-Regeln: drei Varianten im Vergleich",
           "Basis: %s BÜ-Vorgänge aus dem April-Extrakt mit %s extrahierten Nummern."
           % (tsd(N), tsd(WERTE)))
kopf(ws, ["Variante", "beide", "nur Agentur", "nur Vermittler", "keine",
          "Nummern ohne Zuordnung"], [50, 12, 14, 16, 12, 22], zeile=4)
r = 5
for name, bef, off, fill in VARIANTEN:
    zeile_schreiben(ws, r, [name] + [bef[k] for k in KATEGORIEN] +
                    ["" if off is None else off], fill=fill)
    r += 1

ws2 = xl.create_sheet("Wechsler V nach A")
titelzeile(ws2, "Nummern, die neu als Agenturnummer gelten",
           "Heute Vermittlernummer, nach den neuen Regeln 7- oder 9-stellig. "
           "%s Werte in %s Schreibweisen." % (tsd(N_WECHSLER), tsd(len(wechsler_liste))))
kopf(ws2, ["Rohwert heute", "bereinigt", "Stellen", "Vorgänge", "Präfix 890/6000?"],
     [26, 20, 10, 12, 18], zeile=4)
r = 5
for (roh, z), c in wechsler_liste.most_common():
    passt = "ja" if z.startswith(("890", "6000")) else "nein"
    zeile_schreiben(ws2, r, [roh, z, len(z), c, passt], fill=GELB if passt == "nein" else None)
    r += 1

ws3 = xl.create_sheet("Ohne Zuordnung")
titelzeile(ws3, "Nummern, die durch das Raster fallen",
           "Weder unter 7 Stellen noch genau 7 oder 9 Stellen. Rechte Spalte: löst "
           "Variante C den Fall?")
kopf(ws3, ["Rohwert", "bereinigt", "Stellen", "Vorgänge", "in Variante C gelöst?"],
     [28, 20, 10, 12, 22], zeile=4)
r = 5
c_offen_roh = {k[0] for k in C_offen}
for (roh, z, laenge), c in sorted(A_offen.items(), key=lambda x: (-x[1], x[0][2])):
    geloest = "nein" if roh in c_offen_roh else "ja"
    zeile_schreiben(ws3, r, [roh, z, laenge, c, geloest],
                    fill=GRUEN if geloest == "ja" else ROT)
    r += 1

xl.calculation.fullCalcOnLoad = True
xl.save(EXCEL)


# ---------------------------------------------------------------------------
# Mail
# ---------------------------------------------------------------------------
AG_HEUTE = heute["beide"] + heute["nur Agentur"]
AG_B = B_bef["beide"] + B_bef["nur Agentur"]
AG_C = C_bef["beide"] + C_bef["nur Agentur"]
acht_familien = collections.Counter()
for r in BUE:
    for roh, _ in kandidaten(r, True):
        k = ziffern(roh).lstrip("0")
        if len(k) == 8:
            acht_familien[k[:3]] += 1
N_1000 = acht_familien["100"]
N_6008 = acht_familien["600"]

top_wechsler = ", ".join("<code>%s</code> (%dx)" % (roh, c)
                         for (roh, _), c in wechsler_liste.most_common(6))

HTML = """<div>
<p>Hallo Marcus,</p>

<p>danke für den Vorschlag - ich habe die vier Regeln über den April-Extrakt
laufen lassen ({N} BÜ-Vorgänge, {WERTE} extrahierte Nummern). Kurz: Sie
funktionieren, die Abdeckung bei der Agenturnummer steigt deutlich. Eine
Entscheidung fehlt noch, dazu unten die Frage.</p>

{TAB_KERN}

<p>Der Einstieg über die Agenturnummer - für euch der Weg in EMMA - deckt
also spürbar mehr Vorgänge ab als heute.</p>

<h3>Meine Frage: Was soll mit achtstelligen Nummern passieren?</h3>
<p>Die Regeln decken zwei Längen ab: unter 7 Stellen wird Vermittlernummer,
7 oder 9 Stellen wird Agenturnummer. <b>Für 8 Stellen gibt es keine Regel</b> -
und die kommt {ACHT} mal vor. Diese Vorgänge stehen danach ganz ohne Nummer da,
weil es dort jeweils die einzige ist.</p>

<p><b>Unser Vorschlag: eine führende Null ergänzen, damit sie neunstellig sind
und als Agenturnummer gelten.</b> Die Daten stützen das - {N_1000} der {ACHT}
Werte gehören zur selben Familie, die sonst neunstellig geschrieben wird:
<code>01000-3083</code>, <code>01000/3451</code>, <code>A010006515</code>. Ohne
Null steht dort <code>10003083</code>, mit Null <code>010003083</code>, und das
passt ins Muster fünf Stellen Agentur plus vier Stellen Unternummer.</p>

<p>Mit dieser Regel bleiben von ursprünglich {A_OFFEN} nicht zuordenbaren
Nummern nur noch {C_OFFEN} übrig - Einzelfälle mit 10 und mehr Stellen, die wir
direkt als &bdquo;zu prüfen&ldquo; kennzeichnen würden.</p>

<p>Eine Einschränkung dazu: Bei {N_6008} Werten beginnt die Nummer mit
<code>6008</code>, etwa <code>60083652</code>. Dort fehlt die Null vermutlich
nicht vorn, sondern in der Mitte - <code>600083652</code> passt in die Reihe
der 6000-Agenturnummern, <code>060083652</code> nicht. Euer EMMA-Abgleich würde
diese Fälle abfangen, sie landen dann bei &bdquo;zu prüfen&ldquo;.</p>

<h3>Zwei Punkte am Rande</h3>
<ul>
<li><b>Felder mit zwei Nummern:</b> Werte wie <code>V 85357 A 777-0503</code>
enthalten zwei Nummern. Die Bereinigung zog sie zu einer zwölfstelligen
zusammen; wir trennen sie jetzt vorher auf, dann ergibt sich sauber Vermittler
85357 und Agentur 7770503. Reine Technik, an eurer Regellogik ändert sich
nichts.</li>
<li><b>{N_WECHSLER} Nummern wechseln die Spalte</b> - heute Vermittlernummer,
nach der Längenregel Agenturnummer, etwa <code>8923555</code> oder
<code>601000040</code>. Sie stehen im Excel-Blatt &bdquo;Wechsler V nach
A&ldquo;. Magst du da stichprobenhaft draufschauen, ob die Einordnung
passt?</li>
</ul>

<p>Der EMMA-Teil deines Vorschlags lässt sich genau so umsetzen. Sobald die
Frage zu den Achtstellern geklärt ist, ziehe ich die Regeln ein und lasse den
April-Extrakt neu durchlaufen.</p>

<p>Viele Grüße</p>
</div>"""

HTML = HTML.replace("{TAB_KERN}", tab(
    ["", "heute", "mit euren Regeln", "+ achtstellige als Agentur"],
    [["Vorgänge mit Agenturnummer",
      "%s (%s)" % (tsd(AG_HEUTE), pz(AG_HEUTE, N)),
      "%s (%s)" % (tsd(AG_B), pz(AG_B, N)),
      "%s (%s)" % (tsd(AG_C), pz(AG_C, N))],
     ["Vorgänge ohne jede Nummer", tsd(heute["keine"]), tsd(B_bef["keine"]),
      tsd(C_bef["keine"])],
     ["Nummern ohne Zuordnung", "-", tsd(B_OFFEN), tsd(C_OFFEN)]],
    num_spalten=(1, 2, 3)))

for platz, wert in (
        ("{N}", tsd(N)), ("{WERTE}", tsd(WERTE)),
        ("{ACHT}", tsd(ACHT)), ("{N_1000}", tsd(N_1000)), ("{N_6008}", tsd(N_6008)),
        ("{A_OFFEN}", tsd(A_OFFEN)), ("{C_OFFEN}", tsd(C_OFFEN)),
        ("{N_WECHSLER}", tsd(N_WECHSLER))):
    HTML = HTML.replace(platz, wert)

offen_platz = re.findall(r"\{[A-Z_0-9]+\}", HTML)
assert not offen_platz, offen_platz


# ---------------------------------------------------------------------------
# Klartext aus demselben HTML
# ---------------------------------------------------------------------------
def umbruch(text, breite=76):
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


def ascii_tab(kopfzeilen, zeilen, breiten):
    aus = ["  " + "".join(("%-" + str(b) + "s") % k for k, b in zip(kopfzeilen, breiten)),
           "  " + "-" * (sum(breiten) - 2)]
    for zl in zeilen:
        aus.append("  " + "".join(("%-" + str(b) + "s") % w for w, b in zip(zl, breiten)))
    return "\n".join(aus)


TAB_TEXT = [
    ascii_tab(["", "heute", "mit euren Regeln", "+ 8-stellige"],
              [["Vorgänge mit Agenturnummer", "%s (%s)" % (tsd(AG_HEUTE), pz(AG_HEUTE, N)),
                "%s (%s)" % (tsd(AG_B), pz(AG_B, N)),
                "%s (%s)" % (tsd(AG_C), pz(AG_C, N))],
               ["Vorgänge ohne jede Nummer", tsd(heute["keine"]), tsd(B_bef["keine"]),
                tsd(C_bef["keine"])],
               ["Nummern ohne Zuordnung", "-", tsd(B_OFFEN), tsd(C_OFFEN)]],
              [28, 16, 18, 16]),
]

t = HTML
t = re.sub(r"<h3>(.*?)</h3>", lambda m: "\n@@H@@" + m.group(1) + "\n", t)
t = re.sub(r"<li>(.*?)</li>",
           lambda m: "@@LI@@" + re.sub(r"\s+", " ", m.group(1)).strip() + "\n",
           t, flags=re.S)
t = re.sub(r"<table>.*?</table>", "@@TAB@@", t, flags=re.S)
t = re.sub(r"</p>", "\n\n", t)
t = re.sub(r"<[^>]+>", "", t)
for a, b in (("&bdquo;", "„"), ("&ldquo;", "“"), ("&nbsp;", " "), ("&amp;", "&")):
    t = t.replace(a, b)
for tabelle in TAB_TEXT:
    t = t.replace("@@TAB@@", "\n" + tabelle + "\n", 1)

zeilen_txt = []
for zeile in t.split("\n"):
    zeile = zeile.rstrip()
    if zeile.startswith("@@H@@"):
        titel = zeile[5:]
        zeilen_txt += ["", titel, "-" * len(titel)]
    elif zeile.startswith("@@LI@@"):
        zeilen_txt.append("  - " + umbruch(zeile[6:], 72).replace("\n", "\n    "))
    elif zeile.startswith("  ") or not zeile.strip():
        zeilen_txt.append(zeile)
    else:
        zeilen_txt.append(umbruch(zeile))
TEXT = re.sub(r"\n{3,}", "\n\n", "\n".join(zeilen_txt)).strip() + "\n"

baue_eml(EML, BETREFF, TEXT, HTML, [EXCEL])
print("   A: keine=%d offen=%d | B: keine=%d offen=%d | C: keine=%d offen=%d (heute %d)"
      % (A_bef["keine"], A_OFFEN, B_bef["keine"], B_OFFEN,
         C_bef["keine"], C_OFFEN, heute["keine"]))
