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

import re
import sys

from mailbau import baue_eml, tab

ZIEL = sys.argv[1] if len(sys.argv) > 1 else "Datenfelder-KI-Deckblatt.eml"
BETREFF = "Datenfelder KI-Deckblatt - Zielbild und offene Punkte"
ANHAENGE = ["260807 Anpassung Deckblatt_V1.pdf"]

# ---------------------------------------------------------------------------
# Felddefinition: (Feld, Beschreibung, Ausprägungen/Format, Beispiel,
#                  Anpassung, Markierung)
# Markierung: "neu" = neu oder getrennt, "offen" = fachliche Vorgabe fehlt
# Beispiele stammen aus dem Musterdeckblatt BÜ-KI-084 im Review-PDF.
# ---------------------------------------------------------------------------
ABSCHNITTE = [
 ("Kopfzeile", [
  ("Lfd-Nr", "Laufende Nummer des Vorgangs, stellt die Verbindung zur "
             "Excel-Auswertung her", "BÜ-KI-001, BÜ-KI-002, ...", "BÜ-KI-084", "", None),
  ("Status (Ampel)", "Farbige Einordnung des Vorgangs auf einen Blick",
   "Neue BÜ | BÜ-Reminder | Keine BÜ", "Neue BÜ", "", None),
  ("Typ", "Klassischer Makler mit Maklervollmacht oder Mehrfachagent mit "
          "Betreuungswechsel", "Makler | MFA", "Makler", "", None),
  ("Klassifikation", "Fachliche Einordnung des Vorgangs, als Badge wiederholt",
   "siehe Abschnitt Klassifikation &amp; Vorgang", "BUe-Vorgang", "", None),
  ("Maklervollmacht", "Status der Maklervollmacht, als Badge wiederholt",
   "siehe Abschnitt Maklervollmacht &amp; Unterschriften", "OK", "", None),
 ]),
 ("Kunde &amp; Vertrag", [
  ("Kunde_Nachname", "Nachname des Versicherungsnehmers laut Anschreiben", "Text",
   "Haug", "NEU - Trennung von &bdquo;Kunde (Name)&ldquo;", "neu"),
  ("Kunde_Vorname", "Vorname des Versicherungsnehmers laut Anschreiben", "Text",
   "Benjamin", "NEU - Trennung von &bdquo;Kunde (Name)&ldquo;", "neu"),
  ("Geburtsdatum VN", "Geburtsdatum des Versicherungsnehmers", "TT.MM.JJJJ",
   "07.07.1988", "", None),
  ("Kunde_Strasse", "Straßenname und Hausnummer des Versicherungsnehmers",
   "Text inkl. Hausnummer. &bdquo;Straße&ldquo; wird immer als "
   "<b>Str.</b> abgekürzt, auch wenn im Dokument &bdquo;Strasse&ldquo; oder "
   "&bdquo;Straße&ldquo; ausgeschrieben steht.", "Wiesenstr. 8",
   "NEU - Trennung von &bdquo;Kunde-Adresse&ldquo;", "neu"),
  ("Kunde_PLZ", "Postleitzahl des Versicherungsnehmers", "5 Ziffern", "64331",
   "NEU - Trennung von &bdquo;Kunde-Adresse&ldquo;", "neu"),
  ("Kunde_Ort", "Wohnort des Versicherungsnehmers", "Text", "Weiterstadt",
   "NEU - Trennung von &bdquo;Kunde-Adresse&ldquo;", "neu"),
  ("Versicherungsnummer", "Vertragsnummer des betroffenen Vertrags, in der Originalform "
                          "aus dem Dokument - inklusive Buchstaben-Präfix, führender "
                          "Nullen und Suffixen. Mehrere Verträge kommagetrennt.",
   "Text", "LF70902206", "", None),
  ("Sparte", "Sparte des Vertrags, als Kürzel - abgeleitet aus den ersten beiden "
             "Buchstaben der Versicherungsnummer", "2 Buchstaben", "LF",
   "GEÄNDERT - Kürzel statt der bisherigen Gruppen Komposit / Leben / KV; "
   "Liste gültiger Kürzel wird benötigt", "offen"),
  ("Kunden-/Partnernummer", "Kunden- bzw. Partnernummer des VN in den ERGO-Systemen - "
                            "nicht die Versicherungs- oder Maklernummer",
   "Ziffern", "1234567890", "", None),
  ("MV_VN_Nachname", "Nachname des Vollmachtgebers laut Maklervollmacht - dient dem "
                     "Abgleich mit dem Anschreiben", "Text", "Haug",
   "NEU - Trennung von &bdquo;VN laut MV&ldquo;", "neu"),
  ("MV_VN_Vorname", "Vorname des Vollmachtgebers laut Maklervollmacht", "Text",
   "Benjamin", "NEU - Trennung von &bdquo;VN laut MV&ldquo;", "neu"),
 ]),
 ("Makler", [
  ("Pool", "Maklerpool bzw. Abwickler, über den eingereicht wird",
   "Konsolidierter Poolname oder leer", "Fonds Finanz Maklerservice GmbH", "", None),
  ("Vermittler", "Das angebundene Maklerunternehmen unter dem Pool; ist der Makler eine "
                 "Einzelperson, steht hier deren Name",
   "Firmenname oder Personenname", "Finanzservice Limnell", "", None),
  ("Agentur-Nr", "Agenturnummer des Maklers beim Versicherer",
   "*** Platzhalter - Ausprägungen von Marcus ***", "600042497",
   "NEU - Trennung von &bdquo;Agentur-/Personalnummer&ldquo;", "offen"),
  ("Personalnummer", "Personal- bzw. Vermittlernummer des Maklers beim Versicherer",
   "*** Platzhalter - Ausprägungen von Marcus ***", "811774",
   "NEU - Trennung von &bdquo;Agentur-/Personalnummer&ldquo;", "offen"),
  ("Makler_E-Mail", "Mailadresse des Maklers", "Mailadresse",
   "info@finanzservice-limnell.de", "NEU - welche Adresse soll das sein?", "offen"),
  ("Makler_Adresse", "Anschrift des Maklers laut Anschreiben",
   "Text ohne Komma. &bdquo;Straße&ldquo; wird immer als <b>Str.</b> abgekürzt.",
   "Riesstr. 25 80992 München",
   "Trennung in Str. / PLZ / Ort wie beim Kunden?", "offen"),
  ("Auftrag-Datum", "Datum des Pool- bzw. Makler-Anschreibens - nicht der Mail-Eingang",
   "TT.MM.JJJJ", "06.04.2026", "", None),
 ]),
 ("Klassifikation &amp; Vorgang", [
  ("Vorgangstyp", "Grobe Einordnung, worum es sich bei der Mail überhaupt handelt",
   "Makler-Vorgang | Zustellfehler | Ergo-Outbound | System-Mail | Werbung-Spam | Unklar",
   "Makler-Vorgang", "Reminder ergänzen? siehe offene Punkte", "offen"),
  ("Klassifikation", "Fachliche Einordnung des Anliegens",
   "BUe-Vorgang | Anfrage-Ruecksprache | Antrag-Aenderung | Schadenmeldung | "
   "Nicht-Standard | Kein-Makler-Vorgang", "BUe-Vorgang", "", None),
  ("BÜ-Wunsch vorhanden", "Enthält die Mail einen ausformulierten Auftrag zur "
                          "Bestandsübertragung? Eine leere Mail mit nur einer "
                          "Maklervollmacht im Anhang ergibt &bdquo;nein&ldquo;.",
   "ja | nein", "ja", "", None),
  ("Prüfung Unterlagen", "Abgeleitete Vollständigkeitsprüfung der eingereichten Unterlagen",
   "siehe eigene Tabelle unten", "Vollständig",
   "Ausprägung &bdquo;VN-Unterschrift fehlt&ldquo; kommt dazu", "neu"),
  ("Reminder/Wiedervorlage", "Handelt es sich um eine Erinnerung zu einem bereits "
                             "eingereichten Vorgang?", "ja | nein", "nein", "", None),
  ("Reminder-Quelle", "Woran der Reminder erkannt wurde; leer, wenn kein Reminder",
   "Anhang-PDF-Inhalt | Anhang-Dateiname | Mail-Betreff | Mail-Body (kommagetrennt)",
   "Mail-Betreff", "", None),
 ]),
 ("Maklervollmacht &amp; Unterschriften", [
  ("MV enthalten", "Liegt der Mail eine Maklervollmacht bei?", "ja | nein", "ja", "", None),
  ("MV-Status", "Zusammenfassender Status der Maklervollmacht",
   "OK | MV unvollständig | MV fehlt | kein Anhang | nicht prüfbar", "OK", "", None),
  ("MV auf Makler ausgestellt", "Ist die Vollmacht explizit auf den Makler ausgestellt "
                                "oder blanko?", "ja | nein | nicht prüfbar", "ja", "", None),
  ("MV vollumfänglich", "Deckt die Vollmacht den gesamten Bestand ab oder ist sie "
                        "eingeschränkt?", "ja | teilweise | nein | nicht prüfbar", "ja",
   "", None),
  ("MV-Einschränkungen", "Worauf die Vollmacht eingeschränkt ist; nur befüllt, wenn nicht "
                         "vollumfänglich", "Freitext",
   "nur Komposit, ohne Leben", "", None),
  ("Unterschrift Kunde", "Hat der Versicherungsnehmer die Maklervollmacht unterschrieben?",
   "ja | nein | nicht prüfbar", "ja",
   "GEÄNDERT - Unterschrift des Maklers entfällt", "neu"),
  ("Untervollmacht erteilt", "Wurde eine Untervollmacht an einen weiteren Makler oder "
                             "Pool erteilt?", "ja | nein | nicht prüfbar", "ja", "", None),
  ("MV_Unterschriftsdatum", "Datum, an dem der Kunde die Maklervollmacht bzw. den "
                            "Betreuungswunsch unterschrieben hat", "TT.MM.JJJJ",
   "02.04.2026", "NEU auf dem Deckblatt - Feld wird bereits erhoben", "neu"),
  ("MV_Eingangsdatum", "Eingangsdatum der Maklervollmacht bzw. des Betreuungswunsches",
   "TT.MM.JJJJ", "13.04.2026", "NEU - woran ablesen?", "offen"),
 ]),
 ("Ergebnis", [
  ("Zusammenfassung", "Worum es in dem Vorgang geht, in ein bis zwei Sätzen",
   "Freitext, max. rund 280 Zeichen",
   "Fonds Finanz bittet um Bestandsübertragung des Lebensversicherungsvertrags "
   "LF70902206 von Benjamin Haug.", "", None),
  ("Hinweis", "Auffälligkeit, die eine manuelle Prüfung nahelegt; leer, wenn nichts "
              "auffällt", "Freitext, max. rund 220 Zeichen",
   "Maklervollmacht ist auf Finanzservice Limnell ausgestellt, während die "
   "Übertragung in den Bestand der Fonds Finanz erfolgen soll.", "", None),
 ]),
]

# ---------------------------------------------------------------------------
# Ausprägungen von "Prüfung Unterlagen"
# ---------------------------------------------------------------------------
PRUEFUNG = [
 ("<i>leer</i>", "wenn der Vorgangstyp nicht leer bzw. nicht &bdquo;Makler-Vorgang&ldquo; "
                 "ist oder die Klassifikation nicht &bdquo;BUe-Vorgang&ldquo; lautet - "
                 "dann wird gar nicht geprüft", None),
 ("Reminder - keine Vollst.-Prüfung", "wenn Reminder/Wiedervorlage = ja. Bei einer "
                                      "Erinnerung wird nichts neu eingereicht, deshalb "
                                      "keine Prüfung.", None),
 ("<b>Vollständig</b>", "kein Mangel gefunden", None),
 ("BÜ-Wunsch fehlt", "&bdquo;BÜ-Wunsch vorhanden&ldquo; ist nicht &bdquo;ja&ldquo;", None),
 ("MV fehlt", "MV-Status ist &bdquo;MV fehlt&ldquo; oder &bdquo;kein Anhang&ldquo;", None),
 ("MV unvollständig", "MV vollumfänglich ist &bdquo;teilweise&ldquo; oder "
                      "&bdquo;nein&ldquo;", None),
 ("VN-Unterschrift fehlt", "Unterschrift Kunde ist &bdquo;nein&ldquo; - die Vollmacht "
                           "liegt zwar vor, ist aber nicht vom Versicherungsnehmer "
                           "unterschrieben", "neu"),
]

OFFEN = [
 ("Agentur-Nr / Personalnummer", "Regelwerk und Ausprägungen (Marcus)."),
 ("Sparte als Kürzel", "Das Feld ersetzt die bisherigen Gruppen Komposit / Leben / "
                       "KV / Mehrere / Unbekannt - bitte bestätigen, dass die grobe "
                       "Einteilung nicht zusätzlich gebraucht wird. Zur Ermittlung: Eine "
                       "Suche nach beliebigen Zwei-Buchstaben-Kombinationen würde zu "
                       "viele Fehler liefern. Vorschlag: Wir nehmen die ersten beiden "
                       "Buchstaben der erkannten Versicherungsnummer und prüfen sie gegen "
                       "eine Liste gültiger Spartenkürzel. Diese Liste bräuchten wir von "
                       "euch."),
 ("Makler_E-Mail", "Welche Adresse soll hier stehen - die Absenderadresse der Mail oder "
                   "eine andere?"),
 ("MV_Eingangsdatum und MV_Unterschriftsdatum", "Bitte definieren, woran wir die beiden "
                                                "Daten jeweils ablesen sollen."),
 ("Makler_Adresse", "Soll sie wie die Kundenadresse in Str., PLZ und Ort getrennt "
                    "werden?"),
 ("VN-Unterschrift fehlt", "Wir setzen den neuen Wert bei Unterschrift Kunde = "
                           "&bdquo;nein&ldquo;. Wie soll &bdquo;nicht prüfbar&ldquo; "
                           "behandelt werden - als Mangel oder als unauffällig?"),
 ("Reminder im Vorgangstyp", "Ihr hattet angeregt, Reminder/Erinnerungen mit in den "
                             "Vorgangstyp aufzunehmen. Aus unserer Sicht ist das eine "
                             "eigene Dimension - eine Erinnerung an eine "
                             "Bestandsübertragung bleibt fachlich ein BÜ-Vorgang, sie ist "
                             "nur nicht die erste Mail dazu. Deshalb steht sie heute in "
                             "einem eigenen Feld und zusätzlich in der Status-Ampel "
                             "(&bdquo;BÜ-Reminder&ldquo;). Wenn ihr sie trotzdem im "
                             "Vorgangstyp braucht, setzen wir das um - dann bitte kurz "
                             "sagen, ob &bdquo;Reminder&ldquo; die anderen Werte ersetzt "
                             "oder zusätzlich erscheint."),
]


# ---------------------------------------------------------------------------
# Textbausteine (einmal definiert, in beide Fassungen uebernommen)
# ---------------------------------------------------------------------------
EINLEITUNG = [
 "anbei die vollständige Liste der Datenfelder des KI-Deckblatts - im Zielbild, also "
 "inklusive der Anpassungen aus eurem Review 260807 Anpassung Deckblatt_V1.pdf (der "
 "Vollständigkeit halber angehängt). Zu jedem Feld steht, was es beschreibt, welche "
 "Ausprägungen es annehmen kann und ein Beispiel.",
 "Aufgeführt sind ausschließlich die Felder, die auf dem Deckblatt erscheinen. "
 "Zusatzspalten, die es nur in der Excel-Auswertung gibt, sind bewusst nicht dabei.",
 "Eine Schreibweise vorab, weil sie mehrere Felder betrifft: In allen Adressfeldern "
 "wird „Straße“ grundsätzlich als „Str.“ ausgegeben - unabhängig davon, ob im Dokument "
 "„Straße“, „Strasse“ oder „Str.“ steht.",
 "Damit haben wir eine gemeinsame Grundlage: Was wird erhoben, in welcher Form, und wo "
 "sind noch Entscheidungen offen.",
]

ALLE_FELDER = (
 "Bisher blendet das Deckblatt Felder aus, die leer sind - dadurch sieht jedes Deckblatt "
 "etwas anders aus. Nach eurem Kommentar „Immer ALLE Felder befüllt auf KI-Deckblatt "
 "aufführen“ stellen wir das um: Jedes Feld erscheint immer, leere Felder mit einem "
 "Strich. Die Struktur ist damit bei jedem Vorgang identisch. Bitte kurz bestätigen, dass "
 "es so gemeint war.")

PLATZHALTER_TEXT = [
 "Hintergrund: Da die Absender die Begriffe nicht einheitlich verwenden, muss die Maschine "
 "die Unterscheidung über ein Regelwerk lernen. Aus den bisherigen Auswertungen kennen wir "
 "diese Besonderheiten:",
]
BESONDERHEITEN = [
 "Vermittler- und Agenturnummer sind bei der Einreichung teilweise vertauscht.",
 "Die Nummern kommen oft nur als 811774 oder 8903252 ohne A00 bzw. P0 - Beispiel "
 "Finanzguru.",
 "Manchmal wird nur eine der beiden Nummern mitgegeben - Beispiel Fonds Finanz.",
 "Schreibweisen mit Bindestrich wie 898-0003 müssten bereinigt werden.",
]
PLATZHALTER_SCHLUSS = (
 "Konkret gebraucht wird: gültige Stellenzahlen, Bedeutung der Präfixe 6000 und 890, "
 "Umgang mit führenden Nullen und führenden Buchstaben, zulässige Trennzeichen, eine "
 "eventuelle Prüfziffer und die Frage, wie das Tool mit einem Wert umgehen soll, der die "
 "Prüfung nicht besteht. Die Detailauswertung dazu habt ihr in der Mail zur "
 "Formatanalyse.")

SCHLUSS = "Sobald die offenen Punkte geklärt sind, ziehen wir Prompt, Deckblatt und " \
          "Excel-Spalten in einem Schritt nach."


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
             "&bdquo;BUe-Vorgang&ldquo; lautet.</p>")
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
def rein(s):
    """HTML-Auszeichnung und Entities fuer den Klartextteil entfernen."""
    s = re.sub(r"<[^>]+>", "", s)
    for a, b in (("&bdquo;", "„"), ("&ldquo;", "“"), ("&amp;", "&"),
                 ("&nbsp;", " "), ("&mdash;", "-")):
        s = s.replace(a, b)
    return s


def umbruch(s, breite=74, einzug=""):
    worte, zeilen, akt = s.split(), [], ""
    for w in worte:
        if akt and len(akt) + 1 + len(w) > breite:
            zeilen.append(akt)
            akt = w
        else:
            akt = (akt + " " + w).strip()
    if akt:
        zeilen.append(akt)
    return ("\n" + einzug).join(zeilen)


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
                   "„BUe-Vorgang“ lautet."), ""]
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

baue_eml(ZIEL, BETREFF, TEXT, HTML, ANHAENGE)
print("   %d Felder in %d Abschnitten, %d Ausprägungen bei 'Prüfung Unterlagen'"
      % (sum(len(f) for _, f in ABSCHNITTE), len(ABSCHNITTE), len(PRUEFUNG)))
