#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut die Antwort auf Stefans Mail vom 11.08. als .eml.

Kernbotschaft: Variante 2 ist der richtige Weg, EMMA wird nachgeschoben,
Ziel ist der Live-Gang der ersten Version am 27.08. inklusive der
Bereinigungsregel.

Aufruf:  python3 scripts/build-mail-antwort-stefan.py [ziel.eml]
"""

import re
import sys

from mailbau import baue_eml

ZIEL = sys.argv[1] if len(sys.argv) > 1 else "Antwort-Stefan-Nummernlogik.eml"
BETREFF = "AW: Datenfelder KI-Deckblatt - Vorgehen und Termin 27.08."

HTML = """<div>
<p>Hallo Stefan,</p>

<p>danke für die Einordnung - das ist aus meiner Sicht der pragmatische und
richtige Ansatz. Wir setzen Variante 2 jetzt um und schieben den EMMA-Einbau
danach nach.</p>

<p>Warum ich das genauso sehe: Die einfache Zuordnung braucht keine
Vorarbeit und bringt schon den Großteil des Nutzens. Die Punkte, die du beim
EMMA-Abgleich aufführst - mehrere Agenturnummern je Vermittler, kein
einheitlicher Datenpunkt über die Pools hinweg, die Mehrdeutigkeit bei
siebenstelligen Nummern - sind real und lassen sich nicht nebenbei klären. Die
würden uns beim ersten Release nur aufhalten.</p>

<p><b>Ein Punkt ist mir aber wichtig:</b> Die Bereinigungsregel sollte in der
ersten Version schon drin sein. Also Präfixe, führende Nullen und Trennzeichen
entfernen, bevor zugeordnet wird, und Felder mit zwei Nummern vorher
auftrennen. Das ist reine Textverarbeitung, kostet uns keine fachliche
Klärung - und ohne sie zählen <code>60008-1364</code> und
<code>600081364</code> weiterhin als zwei verschiedene Makler. Genau diese
Dubletten wollen wir ja loswerden.</p>

<p><b>Termin:</b> Ich würde darauf hinwirken, dass wir die erste Version zum
<b>27.08.</b> live haben - Variante 2 plus Bereinigungsregel. Der EMMA-Einbau
kommt als Schritt 2 hinterher, sobald die offenen Punkte aus deiner Mail
geklärt sind. Wenn das aus eurer Sicht machbar ist, planen wir es so ein.</p>

<p>Die aktualisierte Feldliste schicke ich gleich separat - EMMA taucht dort
konsequent nicht mehr auf, sondern erst mit Schritt 2.</p>

<p>Viele Grüße</p>
</div>"""

TEXT = re.sub(r"</p>", "\n\n", HTML)
TEXT = re.sub(r"<[^>]+>", "", TEXT)
for a, b in (("&bdquo;", "„"), ("&ldquo;", "“"), ("&nbsp;", " "), ("&amp;", "&")):
    TEXT = TEXT.replace(a, b)


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


TEXT = "\n\n".join(umbruch(" ".join(absatz.split()))
                   for absatz in TEXT.split("\n\n") if absatz.strip()) + "\n"

baue_eml(ZIEL, BETREFF, TEXT, HTML)
