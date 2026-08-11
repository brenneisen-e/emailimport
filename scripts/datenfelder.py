#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Einzige Quelle fuer den Feldkatalog des KI-Deckblatts.

Wird von build-mail-datenfelder.py (Mail) und build-datenfelder-xlsx.py
(Excel) gemeinsam genutzt - so koennen beide Fassungen nicht auseinander-
laufen. Auszeichnungen sind HTML (fuer die Mail); rein() entfernt sie fuer
Klartext und Excel.
"""

import re

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
   "siehe Abschnitt Klassifikation &amp; Vorgang", "BÜ-Vorgang", "", None),
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
  ("Makler_E-Mail", "Absenderadresse der Mail, mit der der Vorgang eingereicht wurde",
   "Mailadresse", "service@fondsfinanz.de", "NEU", "neu"),
  ("Makler_Strasse", "Straßenname und Hausnummer des Maklers laut Anschreiben",
   "Text inkl. Hausnummer. &bdquo;Straße&ldquo; wird immer als <b>Str.</b> abgekürzt.",
   "Riesstr. 25", "NEU - Trennung von &bdquo;Makler-Adresse&ldquo;", "neu"),
  ("Makler_PLZ", "Postleitzahl des Maklers", "5 Ziffern", "80992",
   "NEU - Trennung von &bdquo;Makler-Adresse&ldquo;", "neu"),
  ("Makler_Ort", "Ort des Maklers", "Text", "München",
   "NEU - Trennung von &bdquo;Makler-Adresse&ldquo;", "neu"),
  ("Auftrag-Datum", "Datum des Pool- bzw. Makler-Anschreibens - nicht der Mail-Eingang",
   "TT.MM.JJJJ", "06.04.2026", "", None),
 ]),
 ("Klassifikation &amp; Vorgang", [
  ("Vorgangstyp", "Grobe Einordnung, worum es sich bei der Mail überhaupt handelt",
   "Makler-Vorgang | Zustellfehler | Ergo-Outbound | System-Mail | Werbung-Spam | Unklar",
   "Makler-Vorgang", "Reminder ergänzen? siehe offene Punkte", "offen"),
  ("Klassifikation", "Fachliche Einordnung des Anliegens",
   "BÜ-Vorgang | Anfrage-Ruecksprache | Antrag-Aenderung | Schadenmeldung | "
   "Nicht-Standard | Kein-Makler-Vorgang", "BÜ-Vorgang", "", None),
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
                 "ist oder die Klassifikation nicht &bdquo;BÜ-Vorgang&ldquo; lautet - "
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
 ("MV_Eingangsdatum und MV_Unterschriftsdatum", "Bitte definieren, woran wir die beiden "
                                                "Daten jeweils ablesen sollen."),
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
 "„Straße“, „Strasse“ oder „Str.“ steht. Kunden- und Makleradresse sind beide in Str., "
 "PLZ und Ort getrennt.",
 "Dieselbe Übersicht liegt zusätzlich als Excel bei (Datenfelder-KI-Deckblatt.xlsx) - "
 "dort lässt sie sich filtern, kommentieren und zurückspielen. Drei Blätter: Datenfelder, "
 "Prüfung Unterlagen, Offene Punkte.",
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


