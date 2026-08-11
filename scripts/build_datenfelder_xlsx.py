#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut den Feldkatalog des KI-Deckblatts als Excel-Datei.

Quelle ist datenfelder.py - dasselbe Modul, aus dem auch die Mail erzeugt
wird. Drei Blaetter:
  * "Datenfelder"          - alle Felder mit Beschreibung, Ausprägungen,
                             Beispiel, Anpassung und Status
  * "Prüfung Unterlagen"   - Ausprägungen des abgeleiteten Prüffelds
  * "Offene Punkte"        - was noch fachlich zu klären ist

Aufruf:  python3 scripts/build-datenfelder-xlsx.py [ziel.xlsx]
"""

import sys

import openpyxl

from datenfelder import (ABSCHNITTE, BITTE, NUMMERN_SCHRITTE, NUMMERN_TEXT, OFFEN,
                         PRUEFUNG, SPARTEN, SPARTEN_HINWEIS, SPARTEN_TEXT, rein)
from xlsxbau import (BLAU, FONT, GELB, GRAU, GRUEN, abschluss, hinweis, kopfzeile,
                     schreibe, titel)
from openpyxl.styles import Font

STATUS_TEXT = {"neu": "NEU", "offen": "OFFEN", None: ""}
STATUS_FARBE = {"neu": GRUEN, "offen": GELB, None: None}


def baue_xlsx(ziel):
    wb = openpyxl.Workbook()

    # ---- Blatt 1: Datenfelder -------------------------------------------
    ws = wb.active
    ws.title = "Datenfelder"
    r = titel(ws, "Datenfelder KI-Deckblatt - Zielbild",
              "Alle Felder, die auf dem KI-Deckblatt erscheinen, inklusive der "
              "Anpassungen aus dem Review 260807 Anpassung Deckblatt_V1.pdf. "
              "Zusatzspalten der Excel-Auswertung sind nicht aufgeführt.",
              "Schreibweise: In allen Adressfeldern wird \"Straße\" grundsätzlich als "
              "\"Str.\" ausgegeben.",
              "Grün = neues, getrenntes oder geändertes Feld. Gelb = fachliche Vorgabe "
              "fehlt noch. Die letzte Spalte ist für eure Rückmeldung frei.")
    SP = ["Abschnitt", "Feld", "Beschreibung", "Mögliche Ausprägungen / Format",
          "Beispiel", "Anpassung", "Status", "Rückmeldung"]
    BR = [26, 24, 46, 46, 30, 40, 9, 34]
    kopf = r + 1
    r = kopfzeile(ws, SP, BR, kopf)
    for abschnitt, felder in ABSCHNITTE:
        for feld, beschr, auspr, beisp, anp, mark in felder:
            r = schreibe(ws, r, [rein(abschnitt), rein(feld), rein(beschr), rein(auspr),
                                 rein(beisp), rein(anp), STATUS_TEXT[mark], ""],
                         BR, fill=STATUS_FARBE[mark], fett_spalten=(2,))
    abschluss(ws, kopf, r - 1, len(SP), fixieren="C%d" % (kopf + 1))

    # ---- Blatt 2: Offene Punkte -----------------------------------------
    ws2 = wb.create_sheet("Rückmeldung")
    r = titel(ws2, "Rückmeldung und offene Punkte", rein(BITTE),
              "Darunter die Punkte, zu denen von unserer Seite noch eine Entscheidung "
              "fehlt.")
    SP2 = ["Nr.", "Thema", "Frage / Hintergrund"]
    BR2 = [6, 34, 86]
    kopf2 = r + 1
    r = kopfzeile(ws2, SP2, BR2, kopf2)
    for i, (thema, frage) in enumerate(OFFEN, 1):
        r = schreibe(ws2, r, [i, rein(thema), rein(frage)], BR2, fill=GELB,
                     fett_spalten=(2,))
    letzte = r - 1
    r += 1
    r += 1
    ws2.cell(row=r, column=2,
             value="So werden Agentur-Nr und Personalnummer ermittelt").font = Font(
                 name=FONT, sz=11, b=True, color=BLAU)
    r += 1
    r = schreibe(ws2, r, ["", "", rein(NUMMERN_TEXT)], BR2, fill=GRAU)
    for i, schritt in enumerate(NUMMERN_SCHRITTE, 1):
        r = schreibe(ws2, r, [i, "", rein(schritt)], BR2, fill=GRAU)
    abschluss(ws2, kopf2, letzte, len(SP2))

    # ---- Blatt 3: Sparte ------------------------------------------------
    ws3 = wb.create_sheet("Sparte")
    r = titel(ws3, "Sparte als Kürzel - unsere Annahme", rein(SPARTEN_TEXT))
    SP3 = ["Kürzel", "Sparte", "Präfix der Versicherungsnummer", "Vorgänge", "Anteil"]
    BR3 = [10, 42, 52, 12, 10]
    kopf3 = r + 1
    r = kopfzeile(ws3, SP3, BR3, kopf3)
    for kuerzel, sparte, prae, anz, ant, mark in SPARTEN:
        r = schreibe(ws3, r, [kuerzel, rein(sparte), rein(prae), anz, ant], BR3,
                     fill=STATUS_FARBE[mark], fett_spalten=(1,))
    letzte3 = r - 1
    r = hinweis(ws3, r + 1, rein(SPARTEN_HINWEIS), len(SP3))
    abschluss(ws3, kopf3, letzte3, len(SP3))

    # ---- Blatt 4: Prüfung Unterlagen ------------------------------------
    ws4 = wb.create_sheet("Prüfung Unterlagen")
    r = titel(ws4, "Ausprägungen des Feldes \"Prüfung Unterlagen\"",
              "Das Feld wird überhaupt nur befüllt, wenn der Vorgangstyp leer oder "
              "\"Makler-Vorgang\" ist UND die Klassifikation \"BÜ-Vorgang\" lautet.",
              "Mehrere Mängel werden kommagetrennt ausgegeben, z. B. "
              "\"BÜ-Wunsch fehlt, MV unvollständig\".")
    SP4 = ["Ausprägung", "Wann sie gesetzt wird", "Status"]
    BR4 = [34, 78, 9]
    kopf4 = r + 1
    r = kopfzeile(ws4, SP4, BR4, kopf4)
    for wert, wann, mark in PRUEFUNG:
        r = schreibe(ws4, r, [rein(wert), rein(wann), STATUS_TEXT[mark]], BR4,
                     fill=STATUS_FARBE[mark], fett_spalten=(1,))
    abschluss(ws4, kopf4, r - 1, len(SP4))

    wb.calculation.fullCalcOnLoad = True
    wb.save(ziel)
    return ziel


if __name__ == "__main__":
    ziel = sys.argv[1] if len(sys.argv) > 1 else "Datenfelder-KI-Deckblatt.xlsx"
    baue_xlsx(ziel)
    print("OK - %s geschrieben (%d Felder)"
          % (ziel, sum(len(f) for _, f in ABSCHNITTE)))
