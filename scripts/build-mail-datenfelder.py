#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut die Mail "Datenfelder KI-Deckblatt" als .eml (Outlook-tauglich).

Inhalt: alle Datenfelder des KI-Deckblatts im Zielbild, mit Ausprägungen und
den Anpassungen aus dem Review "260807 Anpassung Deckblatt_V1.pdf".

Aufruf:  python3 scripts/build-mail-datenfelder.py [ziel.eml]
"""

import sys

from mailbau import baue_eml, lies_text, tab

ZIEL = sys.argv[1] if len(sys.argv) > 1 else "Datenfelder-KI-Deckblatt.eml"
BETREFF = "Datenfelder KI-Deckblatt - Zielbild und offene Punkte"
ANHAENGE = ["260807 Anpassung Deckblatt_V1.pdf"]

KOPF = ("Feld", "Mögliche Ausprägungen / Format", "Anpassung")


def block(zeilen):
    """zeilen: (feld, auspraegungen, anpassung, klasse)"""
    return tab(KOPF, [z[:3] for z in zeilen], klassen=[z[3] for z in zeilen])


KOPFZEILE = [
    ("Lfd-Nr", "<code>BÜ-KI-001</code>, <code>BÜ-KI-002</code>, ... - identisch zur "
               "Spalte <code>Lfd_Nr</code> in der Excel-Auswertung", "", None),
    ("Status (Ampel)", "Neue BÜ | BÜ-Reminder | Keine BÜ", "", None),
    ("Typ", "Makler | MFA", "", None),
    ("Klassifikation", "siehe Abschnitt Klassifikation &amp; Vorgang", "", None),
    ("Maklervollmacht", "siehe Abschnitt Maklervollmacht &amp; Unterschriften", "", None),
]

KUNDE = [
    ("Kunde_Nachname", "Text", "<b>NEU</b> - Trennung von &bdquo;Kunde (Name)&ldquo;", "neu"),
    ("Kunde_Vorname", "Text", "<b>NEU</b> - Trennung von &bdquo;Kunde (Name)&ldquo;", "neu"),
    ("Geburtsdatum VN", "TT.MM.JJJJ", "", None),
    ("Kunde_Strasse", "Text inkl. Hausnummer", "<b>NEU</b> - Trennung von "
                                               "&bdquo;Kunde-Adresse&ldquo;", "neu"),
    ("Kunde_PLZ", "5 Ziffern", "<b>NEU</b> - Trennung von &bdquo;Kunde-Adresse&ldquo;", "neu"),
    ("Kunde_Ort", "Text", "<b>NEU</b> - Trennung von &bdquo;Kunde-Adresse&ldquo;", "neu"),
    ("Versicherungsnummer", "Originalform inkl. Buchstaben-Präfix, führenden Nullen und "
                            "Suffixen, z.&nbsp;B. <code>SV096877374.7</code>, "
                            "<code>LF70902206</code>. Mehrere Nummern kommagetrennt.", "", None),
    ("Spartenkürzel", "2 Buchstaben, z.&nbsp;B. <code>SV</code>, <code>LF</code>, "
                      "<code>LV</code>, <code>KR</code>, <code>DA</code>, <code>KV</code>",
     "<b>NEU</b> - Liste gültiger Kürzel wird benötigt", "offen"),
    ("Sparte", "Komposit | Leben | KV | Mehrere | Unbekannt", "", None),
    ("Kunden-/Partnernummer", "Ziffern (KDNR/Partnernummer - nicht die Versicherungs- "
                              "oder Maklernummer)", "", None),
    ("MV_VN_Nachname", "Text", "<b>NEU</b> - Trennung von &bdquo;VN laut MV&ldquo;", "neu"),
    ("MV_VN_Vorname", "Text", "<b>NEU</b> - Trennung von &bdquo;VN laut MV&ldquo;", "neu"),
]

MAKLER = [
    ("Pool", "Konsolidierter Poolname (JDC / Jung, DMS &amp; Cie.; Fonds Finanz "
             "Maklerservice GmbH; blau direkt GmbH; VEMA eG; ...) oder leer", "", None),
    ("Vermittler", "Firma des angebundenen Maklerunternehmens, sonst Name der handelnden "
                   "natürlichen Person", "", None),
    ("<b>Agentur-Nr</b>", "<b>*** Platzhalter - Ausprägungen von Marcus ***</b>",
     "<b>NEU</b> - Trennung von &bdquo;Agentur-/Personalnummer&ldquo;", "offen"),
    ("<b>Personalnummer</b>", "<b>*** Platzhalter - Ausprägungen von Marcus ***</b>",
     "<b>NEU</b> - Trennung von &bdquo;Agentur-/Personalnummer&ldquo;", "offen"),
    ("Makler_E-Mail", "Mailadresse des Maklers", "<b>NEU</b> - welche Adresse?", "offen"),
    ("Makler_Adresse", "Text ohne Komma", "Trennung wie beim Kunden?", "offen"),
    ("Auftrag-Datum", "TT.MM.JJJJ - Datum des Pool-/Makler-Anschreibens, nicht der "
                      "Mail-Eingang", "", None),
]

KLASSIFIKATION = [
    ("Vorgangstyp", "Makler-Vorgang | Zustellfehler | Ergo-Outbound | System-Mail | "
                    "Werbung-Spam | Unklar", "Reminder ergänzen? siehe offene Punkte", "offen"),
    ("Klassifikation", "BUe-Vorgang | Anfrage-Ruecksprache | Antrag-Aenderung | "
                       "Schadenmeldung | Nicht-Standard | Kein-Makler-Vorgang", "", None),
    ("Geschäftstyp", "BUe einfacher Vertrag | BUe Kundenverbindung | Keine BUe",
     "bisher nur in der Excel - mit aufs Deckblatt?", "offen"),
    ("BÜ-Wunsch vorhanden", "ja | nein &mdash; gibt es einen ausformulierten Auftrag zur "
                            "Bestandsübertragung? Eine leere Mail mit nur einer "
                            "Maklervollmacht im Anhang ergibt &bdquo;nein&ldquo;.", "", None),
    ("Prüfung Unterlagen", "siehe eigene Tabelle unten", "", None),
    ("Reminder/Wiedervorlage", "ja | nein", "", None),
    ("Reminder-Quelle", "Anhang-PDF-Inhalt | Anhang-Dateiname | Mail-Betreff | Mail-Body "
                        "(kommagetrennt; leer wenn kein Reminder)", "", None),
]

MV = [
    ("MV enthalten", "ja | nein", "", None),
    ("MV-Status", "OK | MV unvollständig | MV fehlt | kein Anhang | nicht prüfbar", "", None),
    ("MV auf Makler ausgestellt", "ja | nein | nicht prüfbar - explizit auf den Makler "
                                  "ausgestellt statt blanko", "", None),
    ("MV vollumfänglich", "ja | teilweise | nein | nicht prüfbar", "", None),
    ("MV-Einschränkungen", "Freitext (nur wenn nicht vollumfänglich)", "", None),
    ("Unterschrift Kunde", "ja | nein | nicht prüfbar",
     "<b>GEÄNDERT</b> - Unterschrift des Maklers entfällt", "neu"),
    ("Untervollmacht erteilt", "ja | nein | nicht prüfbar", "", None),
    ("MV_Unterschriftsdatum", "TT.MM.JJJJ - Unterschrift durch den Kunden",
     "<b>NEU</b> auf dem Deckblatt (Feld wird bereits erhoben)", "neu"),
    ("MV_Eingangsdatum", "TT.MM.JJJJ", "<b>NEU</b> - woran ablesen?", "offen"),
]

ERGEBNIS = [
    ("Zusammenfassung", "Freitext, max. rund 280 Zeichen", "", None),
    ("Hinweis", "Freitext, max. rund 220 Zeichen - nur bei Auffälligkeiten, z.&nbsp;B. wenn "
                "Maklervollmacht und Anschreiben auf unterschiedliche Firmen lauten",
     "", None),
]

PRUEFUNG = tab(
    ("Ausprägung", "Wann sie gesetzt wird"),
    [["<i>leer</i>", "wenn Vorgangstyp nicht leer bzw. nicht &bdquo;Makler-Vorgang&ldquo; ist "
                     "oder die Klassifikation nicht &bdquo;BUe-Vorgang&ldquo; lautet"],
     ["Reminder - keine Vollst.-Prüfung", "wenn Reminder/Wiedervorlage = ja. Bei einer "
                                          "Erinnerung wird nichts neu eingereicht, deshalb "
                                          "keine Prüfung."],
     ["<b>Vollständig</b>", "kein Mangel gefunden"],
     ["BÜ-Wunsch fehlt", "&bdquo;BÜ-Wunsch vorhanden&ldquo; ist nicht &bdquo;ja&ldquo;"],
     ["MV fehlt", "MV-Status enthält &bdquo;MV fehlt&ldquo; oder &bdquo;kein Anhang&ldquo;"],
     ["MV unvollständig", "MV-Status enthält &bdquo;unvollständig&ldquo;"]])

HTML = """<div>
<p>Hallo zusammen,</p>

<p>anbei die vollständige Liste aller Datenfelder des KI-Deckblatts - im
Zielbild, also inklusive der Anpassungen aus eurem Review
<b>260807 Anpassung Deckblatt_V1.pdf</b> (der Vollständigkeit halber
angehängt). Zu jedem Feld steht, welche Ausprägungen es annehmen kann.</p>

<p>Damit haben wir eine gemeinsame Grundlage: Was wird erhoben, in welcher
Form, und wo sind noch Entscheidungen offen. Grün hinterlegt sind neue oder
getrennte Felder, gelb die Punkte, zu denen uns noch eine fachliche Vorgabe
fehlt.</p>

<h3>Grundsätzliche Änderung: alle Felder immer anzeigen</h3>
<p>Bisher blendet das Deckblatt Felder aus, die leer sind - dadurch sieht jedes
Deckblatt etwas anders aus. Nach eurem Kommentar &bdquo;Immer ALLE Felder
befüllt auf KI-Deckblatt aufführen&ldquo; stellen wir das um: Jedes Feld
erscheint immer, leere Felder mit einem Strich. Die Struktur ist damit bei
jedem Vorgang identisch. Bitte kurz bestätigen, dass es so gemeint war.</p>

<h3>Kopfzeile</h3>
{TAB_KOPF}

<h3>Kunde &amp; Vertrag</h3>
{TAB_KUNDE}

<h3>Makler</h3>
{TAB_MAKLER}

<h3>Klassifikation &amp; Vorgang</h3>
{TAB_KLASS}

<h3>Maklervollmacht &amp; Unterschriften</h3>
{TAB_MV}

<h3>Ergebnis</h3>
{TAB_ERGEBNIS}

<h3>Ausprägungen von &bdquo;Prüfung Unterlagen&ldquo;</h3>
<p>Das Feld wird überhaupt nur befüllt, wenn der Vorgangstyp leer oder
&bdquo;Makler-Vorgang&ldquo; ist <b>und</b> die Klassifikation
&bdquo;BUe-Vorgang&ldquo; lautet.</p>
{TAB_PRUEFUNG}
<p class="hint">Mehrere Mängel werden kommagetrennt ausgegeben, z.&nbsp;B.
&bdquo;BÜ-Wunsch fehlt, MV unvollständig&ldquo;.</p>

<h3>Agentur-Nr und Personalnummer - Platzhalter</h3>
<div class="platzhalter">
<p style="margin-top:0"><b>Diese beiden Felder lassen wir bewusst offen, bis
das fachliche Regelwerk steht. Marcus, du gibst uns hier bitte die Ausprägungen
mit:</b></p>
<p><b>Agentur-Nr</b> &nbsp;&nbsp;&nbsp; *** von Marcus zu befüllen ***<br>
<b>Personalnummer</b> &nbsp;&nbsp;&nbsp; *** von Marcus zu befüllen ***</p>
</div>
<p>Hintergrund: Da die Absender die Begriffe nicht einheitlich verwenden, muss
die Maschine die Unterscheidung über ein Regelwerk lernen. Aus den bisherigen
Auswertungen kennen wir diese Besonderheiten:</p>
<ul>
<li>Vermittler- und Agenturnummer sind bei der Einreichung teilweise
vertauscht.</li>
<li>Die Nummern kommen oft nur als <code>811774</code> oder
<code>8903252</code> ohne A00 bzw. P0 - Beispiel Finanzguru.</li>
<li>Manchmal wird nur eine der beiden Nummern mitgegeben - Beispiel
Fonds Finanz.</li>
<li>Schreibweisen mit Bindestrich wie <code>898-0003</code> müssten bereinigt
werden.</li>
</ul>
<p>Konkret gebraucht wird: gültige Stellenzahlen, Bedeutung der Präfixe 6000
und 890, Umgang mit führenden Nullen und führenden Buchstaben, zulässige
Trennzeichen, eine eventuelle Prüfziffer und die Frage, wie das Tool mit einem
Wert umgehen soll, der die Prüfung nicht besteht. Die Detailauswertung dazu
habt ihr in der Mail zur Formatanalyse.</p>

<h3>Offene Punkte</h3>
<ol>
<li><b>Agentur-Nr / Personalnummer:</b> Regelwerk und Ausprägungen
(Marcus).</li>
<li><b>Spartenkürzel:</b> Eine Suche nach beliebigen
Zwei-Buchstaben-Kombinationen würde zu viele Fehler liefern. Vorschlag: Wir
nehmen die ersten beiden Buchstaben der erkannten Versicherungsnummer und
prüfen sie gegen eine Liste gültiger Spartenkürzel. Diese Liste bräuchten wir
von euch.</li>
<li><b>Makler_E-Mail:</b> Welche Adresse soll hier stehen - die Absenderadresse
der Mail oder eine andere?</li>
<li><b>MV_Eingangsdatum und MV_Unterschriftsdatum:</b> Bitte definieren, woran
wir die beiden Daten jeweils ablesen sollen.</li>
<li><b>Makler_Adresse:</b> Soll sie wie die Kundenadresse in Straße, PLZ und
Ort getrennt werden?</li>
<li><b>Reminder im Vorgangstyp:</b> Ihr hattet angeregt, Reminder/Erinnerungen
mit in den Vorgangstyp aufzunehmen. Aus unserer Sicht ist das eine eigene
Dimension - eine Erinnerung an eine Bestandsübertragung bleibt fachlich ein
BÜ-Vorgang, sie ist nur nicht die erste Mail dazu. Deshalb steht sie heute in
einem eigenen Feld und zusätzlich in der Status-Ampel
(&bdquo;BÜ-Reminder&ldquo;). Wenn ihr sie trotzdem im Vorgangstyp braucht,
setzen wir das um - dann bitte kurz sagen, ob &bdquo;Reminder&ldquo; die
anderen Werte ersetzt oder zusätzlich erscheint.</li>
<li><b>Geschäftstyp:</b> Soll das Feld mit aufs Deckblatt?</li>
</ol>

<p>Sobald die offenen Punkte geklärt sind, ziehen wir Prompt, Deckblatt und
Excel-Spalten in einem Schritt nach.</p>

<p>Rückfragen gerne jederzeit.</p>

<p>Viele Grüße</p>
</div>"""

for platzhalter, inhalt in (
        ("{TAB_KOPF}", block(KOPFZEILE)),
        ("{TAB_KUNDE}", block(KUNDE)),
        ("{TAB_MAKLER}", block(MAKLER)),
        ("{TAB_KLASS}", block(KLASSIFIKATION)),
        ("{TAB_MV}", block(MV)),
        ("{TAB_ERGEBNIS}", block(ERGEBNIS)),
        ("{TAB_PRUEFUNG}", PRUEFUNG)):
    assert platzhalter in HTML, platzhalter
    HTML = HTML.replace(platzhalter, inhalt)

baue_eml(ZIEL, BETREFF, lies_text("mail-datenfelder.txt"), HTML, ANHAENGE)
