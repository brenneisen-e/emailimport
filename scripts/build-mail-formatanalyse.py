#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut die Mail zur Formatanalyse als .eml (Outlook-tauglich).

Inhalt: welche Schreibweisen bei Agentur-Nr und Vermittler-/Personalnummer in
den BUe-Vorgaengen des April-Laufs extrahiert wurden. analyse.xlsx haengt an.

Aufruf:  python3 scripts/build-mail-formatanalyse.py [ziel.eml]
"""

import sys

from mailbau import baue_eml, lies_text, tab

ZIEL = sys.argv[1] if len(sys.argv) > 1 else "Formatanalyse-Agentur-Personalnummern.eml"
ANHAENGE = ["analyse.xlsx"]
# Betreff bewusst kurz gehalten (<= 64 Zeichen): laengere Betreffs werden beim
# Schreiben der EML umgebrochen, was in manchen Clients doppelte Leerzeichen
# erzeugt.
BETREFF = "Formatanalyse Agentur- und Personalnummern - Bitte um Einordnung"

HTML = """<div>
<p>Hallo zusammen,</p>

<p>ich habe ausgewertet, in welchen Formaten die Agentur- und die
Vermittler-/Personalnummer bei der Extraktion der April-Vorgänge
herausgekommen sind.</p>

<p><b>Basis:</b> Von den 5.034 analysierten Mails sind <b>4.329 als
BÜ-Vorgang</b> klassifiziert (4.000 Erstvorgänge und 329 Reminder). Nur diese
sind ausgewertet. Die übrigen 705 Mails - Antrag/Änderung, Anfrage/Rücksprache,
Kein-Makler-Vorgang und Vorgänge ohne Klassifikation - sind vorher
herausgefiltert worden.</p>

<p>Die Auswertung liegt in der beigefügten <b>analyse.xlsx</b> auf zwei neuen
Blättern: <b>Format-Analyse</b> (Auswertung mit Kernaussagen oben) und
<b>Format-Details</b> (alle 5.034 Zeilen einzeln mit erkanntem Muster; die
Spalte BUe_Vorgang zeigt, welche Zeilen eingehen).</p>

<p>Das Folgende ist rein beschreibend - es zeigt, was in diesem Lauf
herausgekommen ist. Ob das den fachlichen Vorgaben entspricht, kann ich aus
den Daten heraus nicht beurteilen; dafür brauchen wir eure Einordnung.</p>

<h3>Wie die beiden Spalten zustande kommen</h3>
<p>Das passiert in zwei Schritten:</p>
<p><b>Schritt 1 - Auslesen:</b> Die KI sucht im Anschreiben nach den Nummern,
in der Regel im Briefkopf hinter &bdquo;Agentur-Nr.&ldquo; oder
&bdquo;Vermittler-Nr.&ldquo;, und übernimmt sie so, wie sie im Dokument
stehen.</p>
<p><b>Schritt 2 - Einsortieren:</b> Danach entscheidet eine fest programmierte
Regel, in welcher der beiden Spalten eine Nummer landet. Die Regel lautet:
Beginnt die Nummer mit 6000 oder 890, ist es eine Agenturnummer. Jede andere
Nummer gilt als Vermittler-/Personalnummer.</p>
<p>Wichtig dabei: Diese Regel gilt immer, auch wenn im Anschreiben eine andere
Beschriftung stand. Stand eine 890er-Nummer im Brief hinter
&bdquo;Vermittler-Nr.&ldquo;, wird sie trotzdem als Agenturnummer abgelegt.
Und weil die zweite Spalte über &bdquo;alles andere&ldquo; definiert ist,
sammelt sie auch alles ein, was gar keine Personalnummer ist - etwa
BD-Nummern, Beraterzeichen oder Pool-IDs.</p>

<h3>Befüllung: Agentur-Nr x Personalnummer (4.329 BÜ-Vorgänge)</h3>
<p>Jeder Vorgang fällt in genau eines der vier inneren Felder, die Anteile
addieren sich zu 100&nbsp;%.</p>
{TAB_MATRIX}
<p class="hint">Zum Vergleich: Eine Agentur-Nr ist in 3.431 Vorgängen (79,3 %)
vorhanden, eine Personalnummer in 2.131 (49,2 %) - diese beiden Werte
überschneiden sich und ergeben deshalb zusammen nicht 100 %.</p>

<h3>Agentur-Nr: 3.431 Werte in 13 verschiedenen Schreibweisen</h3>
{TAB_AGENTUR}
<p class="hint">Nur nach Ziffernanzahl, ohne Trennzeichen und Buchstaben:
55,5 % neun Stellen, 44,3 % sieben Stellen, 0,2 % sonstige.</p>

<h3>Vermittler-/Personalnummer: 2.131 Werte in 37 verschiedenen Schreibweisen</h3>
{TAB_VERMITTLER}
<p class="hint">Nur nach Ziffernanzahl, ohne Trennzeichen und Buchstaben:
73,2 % sechs Stellen, 11,0 % sieben, 10,6 % neun, 2,3 % acht, 1,6 % fünf,
1,3 % sonstige.</p>

<h3>Weitere Beobachtungen</h3>
<ul>
<li>Dieselbe Nummer kommt in unterschiedlichen Schreibweisen an: 12
Agenturnummern (betrifft 1.576 Vorgänge) und 6 Personalnummern (221 Vorgänge)
tauchen in mindestens zwei Varianten auf. Beispiele: <code>60008-1364</code>
(36x) neben <code>600081364</code> (12x), <code>8903235</code> (1.283x) neben
<code>890-3235</code> (5x), <code>008923555</code> (23x) neben
<code>8923555</code> (5x). Für jede nachgelagerte Auswertung sind das aktuell
zwei verschiedene Makler.</li>
<li>170 Personalnummern (8,0 %) tragen führende Nullen. Bei den Agenturnummern
kommt das in diesem Lauf nicht vor.</li>
<li>Einzelne Werte enthalten noch Wortbestandteile, z.&nbsp;B.
<code>Vermittler 891300</code> oder <code>Agentur 8887290</code>.</li>
</ul>

<h3>Führende Buchstaben - was in diesem Lauf passiert ist</h3>
<p>Im Prompt dieses Laufs war hinterlegt, dass nur die reine Nummer bzw.
Kennung ohne Text übernommen wird, führende Buchstaben also entfallen sollten.
In den Daten sieht das so aus: 97 Agentur-Werte tragen weiterhin ein
führendes &bdquo;A&ldquo; (z.&nbsp;B. <code>A600040124</code>), bei den
Personalnummern sind es 64 Werte mit A, V, P oder D. Die
Nachverarbeitung im Tool lässt genau einen führenden Buchstaben
ausdrücklich zu.</p>
<p>Sichtbare Auswirkung: <code>A600080766</code> (15 Vorgänge) und
<code>600080766</code> (1 Vorgang) werden als zwei verschiedene Schlüssel
geführt. Ob die
Buchstaben zur Nummer gehören oder hätten wegfallen sollen, kann ich nicht
beurteilen - das gehört zur Frage nach dem Sollformat.</p>

<h3>Bitte an Marcus</h3>
<p>Marcus, kannst du bitte aufschreiben, wie die Agentur- und die
Personalnummer fachlich eigentlich aufgebaut sind und welche Ausprüfungen
daraus vorstellbar wären? Hilfreich wäre alles, was sich später als Prüfregel
formulieren lässt:</p>
<ul>
<li><b>Aufbau und Länge:</b> Welche Stellenzahlen sind gültig? Haben die
Präfixe 890 und 6000 eine feste Bedeutung, und gibt es weitere?</li>
<li><b>Führende Nullen:</b> bedeutungstragend oder auffüllend? Ist
<code>008923555</code> dieselbe Nummer wie <code>8923555</code>?</li>
<li><b>Führende Buchstaben</b> (A, V, P, D): Bestandteil der Nummer,
separates Kennzeichen oder wegzulassen?</li>
<li><b>Trennzeichen:</b> Sind <code>60008-1364</code> und
<code>60002/0753</code> offizielle Darstellungen oder reine Formatierung aus
dem Anschreiben?</li>
<li><b>Prüfziffer</b> oder Systematik, gegen die sich eine Nummer validieren
lässt?</li>
<li><b>Abgrenzung:</b> Ist &bdquo;beginnt mit 6000 oder 890&ldquo; die
richtige Unterscheidung zwischen Agentur- und Personalnummer, oder gibt es
Fälle, in denen das danebenliegt?</li>
<li><b>Umgang mit Verstößen:</b> Soll ein Wert, der eine Prüfung nicht
besteht, verworfen, geleert oder als &bdquo;zu prüfen&ldquo; markiert
werden?</li>
</ul>

<p>Sobald das steht, können wir die Prüfung direkt in die Nachverarbeitung
einbauen. Die Schreibweise im Originaldokument können wir nicht steuern, die
Ablage im Tool aber schon.</p>

<p>Rückfragen gerne jederzeit.</p>

<p>Viele Grüße</p>
</div>"""

HTML = HTML.replace("{TAB_MATRIX}", tab(
    ["", "Personalnummer vorhanden", "Personalnummer fehlt", "SUMME"],
    [["<b>Agentur-Nr vorhanden</b>", "1.507 &nbsp;(34,8 %)", "1.924 &nbsp;(44,5 %)",
      "<b>3.431 &nbsp;(79,3 %)</b>"],
     ["<b>Agentur-Nr fehlt</b>", "624 &nbsp;(14,4 %)", "274 &nbsp;(6,3 %)",
      "<b>898 &nbsp;(20,7 %)</b>"],
     ["<b>SUMME</b>", "<b>2.131 &nbsp;(49,2 %)</b>", "<b>2.198 &nbsp;(50,8 %)</b>",
      "<b>4.329 &nbsp;(100,0 %)</b>"]],
    num_spalten=(1, 2, 3)))

HTML = HTML.replace("{TAB_AGENTUR}", tab(
    ["Anteil", "Format", "Beispiel", "Anzahl"],
    [["47,8 %", "9 Ziffern, beginnend mit 6000", "<code>600041633</code>", "1.640"],
     ["43,4 %", "7 Ziffern, beginnend mit 890", "<code>8903235</code>", "1.489"],
     ["4,4 %", "5 Ziffern + Bindestrich + 4 Ziffern", "<code>60008-1364</code>", "152"],
     ["2,7 %", "Buchstabe + 9 Ziffern", "<code>A600040124</code>", "92"],
     ["0,9 %", "3 Ziffern + Bindestrich + 4 Ziffern", "<code>890-3235</code>", "30"],
     ["0,4 %", "5 Ziffern + Schrägstrich + 4 Ziffern", "<code>60002/0753</code>", "13"],
     ["0,1 %", "Buchstabe + Leerzeichen + 9 Ziffern", "<code>A 600046012</code>", "4"],
     ["0,1 %", "4 Ziffern", "<code>8904</code>", "4"],
     ["0,2 %", "übrige Einzelfälle", "<code>890666</code>", "7"]],
    num_spalten=(0, 3)))

HTML = HTML.replace("{TAB_VERMITTLER}", tab(
    ["Anteil", "Format", "Beispiel", "Anzahl"],
    [["73,1 %", "6 Ziffern", "<code>811774</code>", "1.557"],
     ["5,8 %", "9 Ziffern", "<code>005556562</code>", "124"],
     ["5,2 %", "7 Ziffern", "<code>7947361</code>", "110"],
     ["5,0 %", "3 Ziffern + Bindestrich + 4 Ziffern", "<code>898-0003</code>", "106"],
     ["2,2 %", "5 Ziffern + Bindestrich + 4 Ziffern", "<code>88417-0384</code>", "47"],
     ["2,0 %", "8 Ziffern", "<code>89420798</code>", "43"],
     ["1,7 %", "Buchstabe + 9 Ziffern", "<code>A010004616</code>", "36"],
     ["1,6 %", "5 Ziffern", "<code>85357</code>", "34"],
     ["3,5 %", "übrige Einzelfälle, 29 weitere Muster",
      "<code>V 85357 A 777-0503</code>, <code>893/4055</code>, "
      "<code>00 111 40 16</code>", "74"]],
    num_spalten=(0, 3)))

baue_eml(ZIEL, BETREFF, lies_text("mail-formatanalyse.txt"), HTML, ANHAENGE)
