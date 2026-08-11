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
  ("Sparte", "Sparte des Vertrags als Kürzel, abgeleitet aus dem Buchstaben-Präfix der "
             "Versicherungsnummer", "KO | KV | LF | KR - siehe eigene Tabelle unten", "LF",
   "GEÄNDERT - Kürzel statt der bisherigen Gruppen Komposit / Leben / KV; Zuordnung ist "
   "eine Annahme und braucht eure Bestätigung", "offen"),
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
  ("Agentur-Nr", "Agenturnummer des Maklers beim Versicherer. Sie ist das führende Feld "
                 "und der Einstieg in den EMMA-Abgleich. Fehlt sie im Anschreiben, wird "
                 "sie über die Personalnummer aus EMMA ergänzt.",
   "7 bis 9 Ziffern nach der Bereinigung. Achtstellige Werte sind die neuen Agenturen "
   "nach ZAV und ohne führende Null gültig.", "600042497",
   "NEU - Trennung von &bdquo;Agentur-/Personalnummer&ldquo;, Zuordnung über die "
   "Stellenzahl", "neu"),
  ("Personalnummer", "Personal- bzw. Vermittlernummer des Maklers beim Versicherer. "
                     "<b>Der Wert kann aus EMMA stammen:</b> Er wird gegen den EMMA-Abzug "
                     "plausibilisiert, bei Abweichung mit der dort hinterlegten Nummer "
                     "überschrieben und bei fehlender Angabe ergänzt.",
   "Weniger als 7 Ziffern nach der Bereinigung, in der Regel 5 oder 6.", "811774",
   "NEU - Trennung von &bdquo;Agentur-/Personalnummer&ldquo;, Zuordnung über die "
   "Stellenzahl", "neu"),
  ("Nummern-Status", "Ergebnis des Abgleichs mit dem EMMA-Abzug. Zeigt zugleich an, ob "
                     "eine Nummer aus EMMA übernommen wurde, weil die Angabe aus der Mail "
                     "nicht passte oder fehlte - der Sachbearbeiter sieht damit sofort, "
                     "wo er hinschauen muss.",
   "OK - Makler mit EMMA validiert<br>"
   "zu prüfen - Agenturnummer aus EMMA übernommen<br>"
   "zu prüfen - Vermittlernummer aus EMMA übernommen<br>"
   "zu prüfen - keine Nummer gefunden",
   "OK - Makler mit EMMA validiert",
   "NEU - ergibt sich aus dem EMMA-Abgleich", "neu"),
  ("Makler_E-Mail", "Absenderadresse der Mail, mit der der Vorgang eingereicht wurde",
   "Mailadresse", "service@fondsfinanz.de", "NEU", "neu"),
  ("Makler_Strasse", "Straßenname und Hausnummer des Maklers laut Anschreiben",
   "Text inkl. Hausnummer. &bdquo;Straße&ldquo; wird immer als <b>Str.</b> abgekürzt.",
   "Riesstr. 25", "NEU - Trennung von &bdquo;Makler-Adresse&ldquo;", "neu"),
  ("Makler_PLZ", "Postleitzahl des Maklers", "5 Ziffern", "80992",
   "NEU - Trennung von &bdquo;Makler-Adresse&ldquo;", "neu"),
  ("Makler_Ort", "Ort des Maklers", "Text", "München",
   "NEU - Trennung von &bdquo;Makler-Adresse&ldquo;", "neu"),
 ]),
 ("Klassifikation &amp; Vorgang", [
  ("Vorgangstyp", "Grobe Einordnung, worum es sich bei der Mail überhaupt handelt",
   "Makler-Vorgang | Zustellfehler | Ergo-Outbound | System-Mail | Werbung-Spam | Unklar",
   "Makler-Vorgang", "", None),
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
  ("Unterschrift Kunde", "Hat der Versicherungsnehmer die Maklervollmacht unterschrieben?",
   "ja | nein | nicht prüfbar", "ja",
   "GEÄNDERT - Unterschrift des Maklers entfällt", "neu"),
  ("Untervollmacht erteilt", "Wurde eine Untervollmacht an einen weiteren Makler oder "
                             "Pool erteilt?", "ja | nein | nicht prüfbar", "ja", "", None),
  ("MV_Unterschriftsdatum", "Datum, an dem der Kunde die Maklervollmacht bzw. den "
                            "Betreuungswunsch unterschrieben hat", "TT.MM.JJJJ",
   "02.04.2026", "NEU auf dem Deckblatt - Feld wird bereits erhoben", "neu"),
  ("MV_Eingangsdatum", "Eingangsdatum der Maklervollmacht bzw. des Betreuungswunsches - "
                       "das ist das Datum der Mail, mit der eingereicht wurde",
   "TT.MM.JJJJ", "13.04.2026", "NEU", "neu"),
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
# Sparte: vorgeschlagene Kürzel und ihre Herleitung aus der Versicherungsnummer
# Mengen aus dem April-Lauf, 4.286 Versicherungsnummern in BÜ-Vorgängen.
# ---------------------------------------------------------------------------
SPARTEN = [
 ("KO", "Komposit", "SV, DA, UP, RS, HP, HA, HR, AS, WG, SB, UV, ER, GU, HG, V",
  "1.813", "42,3 %", None),
 ("KV", "Krankenversicherung", "KV", "1.054", "24,6 %", None),
 ("LF", "Leben", "LF, LV, R, T, FL", "656", "15,3 %", None),
 ("KR", "Kraftfahrt", "KR", "307", "7,2 %", None),
 ("?", "nicht ableitbar - Versicherungsnummer ohne Buchstaben-Präfix",
  "z. B. 328124276, 7125170230214", "386", "9,0 %", "offen"),
 ("?", "nicht ableitbar - Präfix in keiner der vier Gruppen",
  "M, SA, HV, BG, TB, IT, RE, GL und weitere Einzelfälle", "70", "1,6 %", "offen"),
]

SPARTEN_TEXT = (
 "Ihr habt KR = Kraftfahrt sowie KV, KO und LF als Kürzel genannt. Daraus leiten wir "
 "folgende Annahme ab: Die Sparte hat vier Ausprägungen - KO (Komposit), KV "
 "(Krankenversicherung), LF (Leben) und KR (Kraftfahrt). Ermittelt wird sie über das "
 "Buchstaben-Präfix der Versicherungsnummer. Die Zuordnung unten ist aus dem April-Lauf "
 "abgeleitet: Wir haben je Präfix nachgesehen, welcher Sparte die Vorgänge heute "
 "zugeordnet sind. KR-Verträge laufen heute unter Komposit, was zur Lesart Kraftfahrt "
 "passt.")

SPARTEN_HINWEIS = (
 "Wichtig: Bei 386 Vorgängen (9,0 %) hat die Versicherungsnummer gar kein "
 "Buchstaben-Präfix, bei weiteren 70 (1,6 %) passt das Präfix in keine der vier Gruppen. "
 "Für rund jeden zehnten Vorgang brauchen wir also eine Regel, was dann passieren soll - "
 "Feld leer lassen oder die Sparte aus dem Mailtext ableiten.")


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
 ("VN-Unterschrift fehlt", "Unterschrift Kunde ist &bdquo;nein&ldquo; oder &bdquo;nicht "
                           "prüfbar&ldquo; - die Vollmacht liegt zwar vor, aber die "
                           "Unterschrift des Versicherungsnehmers ist nicht belegt. Beide "
                           "Fälle gelten als Mangel bei der Vollständigkeit.", "neu"),
]

OFFEN = [
 ("Ursprungswert bei Überschreibung", "Wenn eine Nummer aus EMMA übernommen wird, steht "
                                     "die vom Makler genannte nicht mehr auf dem "
                                     "Deckblatt. Reicht der Nummern-Status als Hinweis, "
                                     "oder soll der ursprüngliche Wert zusätzlich "
                                     "angezeigt werden - etwa für den Fall, dass sich der "
                                     "Makler auf seine Angabe beruft?"),
 ("Sparte als Kürzel", "Bitte die angenommene Zuordnung in der Spartentabelle "
                       "bestätigen oder korrigieren - insbesondere, ob KO wirklich alles "
                       "außer Kranken, Leben und Kraftfahrt umfasst und ob die "
                       "Präfix-Listen vollständig sind."),
 ("Sparte ohne Präfix", "Bei rund jedem zehnten Vorgang lässt sich die Sparte nicht aus "
                        "der Versicherungsnummer ableiten (kein Buchstaben-Präfix oder "
                        "unbekanntes Präfix). Feld leer lassen oder aus dem Mailtext "
                        "ableiten?"),
]


# ---------------------------------------------------------------------------
# Textbausteine (einmal definiert, in beide Fassungen uebernommen)
# ---------------------------------------------------------------------------
EINLEITUNG = [
 "anbei findet ihr einen Vorschlag für die vollständige Liste der Datenfelder des "
 "KI-Deckblatts - im Zielbild, also inklusive der Anpassungen aus eurem Review 260807 "
 "Anpassung Deckblatt_V1.pdf und der von Marcus vorgeschlagenen Regeln für Agentur- und "
 "Personalnummer. Zu jedem Feld steht, was es beschreibt, welche Ausprägungen es annehmen "
 "kann und ein Beispiel. Grundlage ist genau das Muster-Deckblatt aus eurem Review "
 "(BÜ-KI-084): aufgeführt sind die Felder, die dort stehen, plus die neu hinzugekommenen.",
 "Dieselbe Übersicht liegt zusätzlich als Excel bei (Datenfelder-KI-Deckblatt.xlsx) - "
 "dort lässt sie sich filtern und direkt kommentieren. Vier Blätter: Datenfelder, "
 "Rückmeldung, Sparte, Prüfung Unterlagen.",
 "Grün hinterlegt sind neue, getrennte oder geänderte Felder, gelb die Punkte, zu denen "
 "noch eine fachliche Vorgabe fehlt.",
]

BITTE = (
 "Bitte schaut die Liste einmal durch und meldet bestmöglich bis morgen EOB zurück, "
 "wenn ein Feld fehlt, anders "
 "aussehen soll oder für die manuelle bzw. automatische Prüfung gar nicht gebraucht "
 "wird. In der Excel gibt es dafür eine leere Spalte „Rückmeldung“, die ihr direkt "
 "ausfüllen könnt.")

ALLE_FELDER = (
 "Bisher blendet das Deckblatt Felder aus, die leer sind - dadurch sieht jedes Deckblatt "
 "etwas anders aus. Nach eurem Kommentar „Immer ALLE Felder befüllt auf KI-Deckblatt "
 "aufführen“ stellen wir das um: Jedes Feld erscheint immer, leere Felder mit einem "
 "Strich. Die Struktur ist damit bei jedem Vorgang identisch. Bitte kurz bestätigen, dass "
 "es so gemeint war.")

NUMMERN_TEXT = (
 "Weil die Absender die Begriffe nicht einheitlich verwenden, entscheidet nicht die "
 "Beschriftung im Anschreiben, sondern die Stellenzahl. So läuft es künftig ab:")

NUMMERN_SCHRITTE = [  # noqa: E501
 "Die KI liest die Nummern aus dem Anschreiben aus, so wie sie dort stehen.",
 "Felder, die zwei Nummern enthalten, werden getrennt - aus „V 85357 A 777-0503“ werden "
 "85357 und 777-0503.",
 "Buchstaben-Präfixe, führende Nullen und Trennzeichen werden entfernt.",
 "Weniger als 7 Stellen ergibt die Personalnummer, 7 bis 9 Stellen die Agenturnummer. "
 "Achtstellige Werte gehören dazu - das sind die neuen Agenturen nach ZAV.",
 "Die Agenturnummer ist führend und der Einstieg in den EMMA-Abgleich: Existiert sie und "
 "gehört sie zum Vermittler? Die Personalnummer wird dabei ergänzt oder bei Abweichung "
 "überschrieben.",
 "Fehlt die Agenturnummer, wird über die Personalnummer eingestiegen und die zugehörige "
 "Agentur ergänzt, sofern die Beziehung eindeutig ist.",
 "Bleibt etwas unstimmig, wird der Vorgang im Feld Nummern-Status als „zu prüfen“ "
 "gekennzeichnet.",
 "Das Feld Nummern-Status hält fest, was dabei herausgekommen ist: mit EMMA validiert, "
 "oder eine der drei Prüf-Ausprägungen, wenn eine Nummer aus EMMA übernommen wurde oder "
 "gar keine gefunden wurde.",
]

NUMMERN_SCHLUSS = (
 "Am April-Extrakt gerechnet haben damit 91,1 % der Vorgänge eine Agenturnummer, "
 "gegenüber 79,3 % heute. Offen bleiben nur 7 Werte mit mehr als 9 Stellen, meist zwei "
 "zusammengezogene Nummern - die bekommen den Status „zu prüfen“.")

SCHLUSS = ("Sobald eure Rückmeldung da ist, ziehen wir Prompt, Deckblatt und Excel-Spalten "
           "in einem Schritt nach.")

def rein(s):
    """HTML-Auszeichnung und Entities fuer Klartext und Excel entfernen.

    <br> wird zum Zeilenumbruch - in der Excel bleibt die Aufzaehlung damit
    lesbar, im Klartext setzt umbruch() sie zeilenweise.
    """
    s = re.sub(r"<br\s*/?>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    for a, b in (("&bdquo;", "„"), ("&ldquo;", "“"), ("&amp;", "&"),
                 ("&nbsp;", " "), ("&mdash;", "-")):
        s = s.replace(a, b)
    return s


def umbruch(s, breite=74, einzug=""):
    if "\n" in s:                       # vorhandene Umbrueche erhalten
        return ("\n" + einzug).join(umbruch(teil, breite, einzug)
                                    for teil in s.split("\n"))
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


