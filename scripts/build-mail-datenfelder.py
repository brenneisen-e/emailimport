#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut die Mail "Datenfelder KI-Deckblatt" als .eml (Outlook-tauglich).

Inhalt: ausschliesslich die Felder, die auf dem KI-Deckblatt stehen - im
Zielbild, also inklusive der Anpassungen aus dem Review
"260807 Anpassung Deckblatt_V1.pdf". Excel-Sonderspalten bleiben aussen vor.

Klartext- und HTML-Teil werden aus DERSELBEN Felddefinition erzeugt, damit
beide Fassungen nicht auseinanderlaufen.

Aufruf:  python3 scripts/build-mail-datenfelder.py [ziel.eml]
"""

import sys

from mailbau import baue_eml, tab

ZIEL = sys.argv[1] if len(sys.argv) > 1 else "Datenfelder-KI-Deckblatt.eml"
BETREFF = "Datenfelder KI-Deckblatt - Zielbild und offene Punkte"
EXCEL = "Datenfelder-KI-Deckblatt.xlsx"
ANHAENGE = [EXCEL, "260807 Anpassung Deckblatt_V1.pdf"]

from build_datenfelder_xlsx import baue_xlsx
from datenfelder import (ABSCHNITTE, ALLE_FELDER, BESONDERHEITEN, EINLEITUNG, OFFEN,
                         PLATZHALTER_SCHLUSS, PLATZHALTER_TEXT, PRUEFUNG, SCHLUSS,
                         rein, umbruch)

# ---------------------------------------------------------------------------
# HTML bauen
# ---------------------------------------------------------------------------
def html_tab(felder):
    return tab(("Feld", "Beschreibung", "Ausprägungen / Format", "Beispiel", "Anpassung"),
               [[f[0], f[1], f[2], f[3], f[4]] for f in felder],
               klassen=[f[5] for f in felder])


teile = ["<div>", "<p>Hallo zusammen,</p>"]
for absatz in EINLEITUNG:
    teile.append("<p>%s</p>" % absatz)
teile.append("<p>Grün hinterlegt sind neue, getrennte oder geänderte Felder, gelb die "
             "Punkte, zu denen uns noch eine fachliche Vorgabe fehlt.</p>")
teile.append("<h3>Grundsätzliche Änderung: alle Felder immer anzeigen</h3>")
teile.append("<p>%s</p>" % ALLE_FELDER)
for titel, felder in ABSCHNITTE:
    teile.append("<h3>%s</h3>" % titel)
    teile.append(html_tab(felder))
teile.append("<h3>Ausprägungen von &bdquo;Prüfung Unterlagen&ldquo;</h3>")
teile.append("<p>Das Feld wird überhaupt nur befüllt, wenn der Vorgangstyp leer oder "
             "&bdquo;Makler-Vorgang&ldquo; ist <b>und</b> die Klassifikation "
             "&bdquo;BÜ-Vorgang&ldquo; lautet.</p>")
teile.append(tab(("Ausprägung", "Wann sie gesetzt wird"),
                 [[p[0], p[1]] for p in PRUEFUNG], klassen=[p[2] for p in PRUEFUNG]))
teile.append('<p class="hint">Mehrere Mängel werden kommagetrennt ausgegeben, z.&nbsp;B. '
             "&bdquo;BÜ-Wunsch fehlt, MV unvollständig&ldquo;. Die Ausprägung "
             "&bdquo;VN-Unterschrift fehlt&ldquo; ist neu - bisher floss die Unterschrift "
             "des Kunden nicht in die Prüfung ein.</p>")
teile.append("<h3>Agentur-Nr und Personalnummer - Platzhalter</h3>")
teile.append('<div class="platzhalter">'
             '<p style="margin-top:0"><b>Diese beiden Felder lassen wir bewusst offen, bis '
             'das fachliche Regelwerk steht. Marcus, du gibst uns hier bitte die '
             'Ausprägungen mit:</b></p>'
             '<p><b>Agentur-Nr</b> &nbsp;&nbsp;&nbsp; *** von Marcus zu befüllen ***<br>'
             '<b>Personalnummer</b> &nbsp;&nbsp;&nbsp; *** von Marcus zu befüllen ***</p>'
             '</div>')
for absatz in PLATZHALTER_TEXT:
    teile.append("<p>%s</p>" % absatz)
teile.append("<ul>%s</ul>" % "".join("<li>%s</li>" % b for b in BESONDERHEITEN))
teile.append("<p>%s</p>" % PLATZHALTER_SCHLUSS)
teile.append("<h3>Offene Punkte</h3>")
teile.append("<ol>%s</ol>" % "".join("<li><b>%s:</b> %s</li>" % (a, b) for a, b in OFFEN))
teile.append("<p>%s</p>" % SCHLUSS)
teile.append("<p>Rückfragen gerne jederzeit.</p><p>Viele Grüße</p></div>")
HTML = "\n".join(teile)


# ---------------------------------------------------------------------------
# Klartext aus derselben Definition
# ---------------------------------------------------------------------------
zeilen = ["Hallo zusammen,", ""]
for absatz in EINLEITUNG:
    zeilen += [umbruch(rein(absatz)), ""]
zeilen += ["Markierungen: NEU = neues, getrenntes oder geändertes Feld, "
           "OFFEN = fachliche", "Vorgabe fehlt noch.", ""]
zeilen += ["Grundsätzliche Änderung: alle Felder immer anzeigen",
           "-" * 51, umbruch(rein(ALLE_FELDER)), ""]

for titel, felder in ABSCHNITTE:
    t = rein(titel).upper()
    zeilen += [t, "-" * len(t), ""]
    for feld, beschr, auspr, beisp, anp, mark in felder:
        marke = {"neu": "   [NEU]", "offen": "   [OFFEN]"}.get(mark, "")
        zeilen.append("  %s%s" % (rein(feld), marke))
        zeilen.append("    Beschreibung:  " + umbruch(rein(beschr), 58, " " * 19))
        zeilen.append("    Ausprägungen:  " + umbruch(rein(auspr), 58, " " * 19))
        zeilen.append("    Beispiel:      " + umbruch(rein(beisp), 58, " " * 19))
        if anp:
            zeilen.append("    Anpassung:     " + umbruch(rein(anp), 58, " " * 19))
        zeilen.append("")

zeilen += ["AUSPRÄGUNGEN VON „PRÜFUNG UNTERLAGEN“", "-" * 38, "",
           umbruch("Das Feld wird überhaupt nur befüllt, wenn der Vorgangstyp leer oder "
                   "„Makler-Vorgang“ ist UND die Klassifikation "
                   "„BÜ-Vorgang“ lautet."), ""]
for wert, wann, mark in PRUEFUNG:
    zeilen.append("  %s%s" % (rein(wert), "   [NEU]" if mark else ""))
    zeilen.append("    " + umbruch(rein(wann), 70, "    "))
    zeilen.append("")
zeilen += [umbruch("Mehrere Mängel werden kommagetrennt ausgegeben, z. B. "
                   "„BÜ-Wunsch fehlt, MV unvollständig“. Die Ausprägung "
                   "„VN-Unterschrift fehlt“ ist neu - bisher floss die "
                   "Unterschrift des Kunden nicht in die Prüfung ein."), ""]

zeilen += ["AGENTUR-NR UND PERSONALNUMMER - PLATZHALTER", "-" * 43, "",
           umbruch("Diese beiden Felder lassen wir bewusst offen, bis das fachliche "
                   "Regelwerk steht. Marcus, du gibst uns hier bitte die Ausprägungen "
                   "mit:"), "",
           "  Agentur-Nr        *** von Marcus zu befüllen ***",
           "  Personalnummer    *** von Marcus zu befüllen ***", ""]
for absatz in PLATZHALTER_TEXT:
    zeilen += [umbruch(rein(absatz)), ""]
for i, b in enumerate(BESONDERHEITEN, 1):
    zeilen.append("  %d. %s" % (i, umbruch(rein(b), 70, "     ")))
zeilen += ["", umbruch(rein(PLATZHALTER_SCHLUSS)), ""]

zeilen += ["OFFENE PUNKTE", "-" * 13, ""]
for i, (a, b) in enumerate(OFFEN, 1):
    zeilen.append("  %d. %s: %s" % (i, rein(a), umbruch(rein(b), 68, "     ")))
zeilen += ["", umbruch(SCHLUSS), "", "Rückfragen gerne jederzeit.", "", "Viele Grüße", ""]

TEXT = "\n".join(zeilen)

# Excel zuerst bauen - sie haengt an der Mail und kommt aus derselben Quelle.
baue_xlsx(EXCEL)

baue_eml(ZIEL, BETREFF, TEXT, HTML, ANHAENGE)
print("   %d Felder in %d Abschnitten, %d Ausprägungen bei 'Prüfung Unterlagen'"
      % (sum(len(f) for _, f in ABSCHNITTE), len(ABSCHNITTE), len(PRUEFUNG)))
