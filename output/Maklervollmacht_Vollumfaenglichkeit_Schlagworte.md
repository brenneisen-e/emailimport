# Maklervollmacht — Vollumfänglichkeit (Schlagwort-Liste)

Quellen: AAW-Allg (`Bestandsübertragung Makler allgemeine Bearbeitung AAW unabhängig von ZAV.docx`, 2.1.1 + 2.1.2), AAW-GES (Master-AAW Gesundheit, Kapitel 5), PVC2D-Vorgabe der Rechtsabteilung, sowie Hospitationen T2 (Inge Lorusso, KV) und T3 (Alexandra Mai, KOM/LEB).

Diese Liste ist die maßgebliche Prüfgrundlage für `Vorgaenge_Analysieren` ab v2.8 — die GPT-Klassifikation orientiert sich daran, NICHT an einer generischen Form-Felder-Heuristik.

---

## A) Juristische Grundlage

§§ 133, 157 BGB: Auslegung nach dem **geäußerten Willen** (wie der Vertragspartner sie verstehen darf), nicht nach dem inneren Willen des Vollmachtgebers.

**Inhaltlicher Minimalstandard** (PVC2D / AAW-Allg 2.1.2):

> „Die uneingeschränkte aktive und passive Vertretung des Kunden gegenüber den jeweiligen Versicherungsgesellschaften einschließlich der Abgabe aller die Vertragsangelegenheiten betreffenden Willenserklärungen und Anzeigen, insbesondere die Kündigung bestehender und den Abschluss neuer Versicherungsverträge."

---

## B) Anerkannte Schlagwort-Varianten — eine reicht

Eine der folgenden Formulierungen muss in der MV stehen, damit sie als vollumfänglich gilt:

| # | Schlagwort / Formulierung | Quelle |
|---|---|---|
| 1 | „abschließen, ändern und kündigen" (in Bezug auf Versicherungsverträge des VN) | T2 (Lorusso) |
| 2 | „Willenserklärungen abgeben und entgegen[nehmen]" / „aktiv und passiv vertreten" | T2, T3, AAW-Allg |
| 3 | „bevollmächtigt zu vertreten" (uneingeschränkt) | T2, AAW-Allg |
| 4 | „Abgabe und Entgegennahme von Willenserklärungen" + Änderung, Kündigung, Abschluss eines Folgevertrages | T3 (am Check 24-Layout erläutert) |
| 5 | „uneingeschränkte aktive und passive Vertretung … einschließlich Abgabe aller … Willenserklärungen und Anzeigen, insbesondere Kündigung bestehender und Abschluss neuer Versicherungsverträge" | AAW-Allg / PVC2D |

Wörtliche Belege:

- T2 (Lorusso, KV): „Maklervollmachten für die Bestandsübertragung müssen vollumfänglich sein. Das heißt, der Makler muss für den Kunden alles dürfen. Das sind die Punkte, die hier genannt sind: **Abschließen, Ändern und Kündigen**. Das beinhaltet diese Vollumfänglichkeit."
- T2: „… **Willenserklärung abgeben und entgegen** ist auch vollumfänglich. […] **„es bevollmächtigt zu vertreten"** würde auch vollumfänglich bedeuten."
- T3 (Mai, KOM/LEB): „Die **Abgabe und Entgegennahme von den Willenserklärungen**. Die **Änderung, Kündigung** und des bestehenden Versicherungsvertrags und auch den **Abschluss eines Folgevertrages**, also das sind so die wesentlichen Dinge […] dafür, dass die Maklervollmacht vollumfänglich ist."

---

## C) Einschränkungen, die die Vollumfänglichkeit ausschließen (Ablehnungsgründe)

Aus AAW-GES Kapitel 5:

- **„MV eingeschränkt"** — die MV beinhaltet nicht alle oben genannten Voraussetzungen.
- **„KV ausgeschlossen"** — die MV gilt explizit nicht für Krankenversicherungsverträge.
- **„Schriftliche Zustimmung VN"** — in der MV steht, dass der VN allen Änderungen schriftlich zustimmen muss → keine uneingeschränkte Vollmacht.
- Fehlt eines der Schlagwörter aus B) komplett oder ist es eingeschränkt → **Ablehnungsgrund**, Notiz **2013** (Ablehnungs-Bot, GeVo „MV fehlt" oder „MV eingeschränkt").

---

## D) Flankierende Pflicht-Bedingungen (zusätzlich zur Vollumfänglichkeit)

Lt. AAW-Allg 2.1.1 „Wann ist eine MV gültig?":

1. Sie ist auf den **VN** ausgestellt (nicht auf eine versicherte Person).
2. Der **Makler / Maklerpool** ist namentlich genannt.
3. Sie ist **unterschrieben** (auch digital zulässig). **VN-Unterschrift Pflicht, Maklerunterschrift optional.**
4. Der Aussteller muss tatsächlich Makler sein (in KV: VVR-Prüfung).

Pool-spezifische Ausnahmen bei Punkt 3:

- **Check 24 digital**: ohne Unterschrift zulässig.
- **Verifox / Fonds Finanz / Impuls / Watson**: Unterschrift im separaten Datenschutz-Anhang.

---

## E) Mapping auf das Tool (Spalten der Excel-Analyse, ab v2.8)

| Spalte | Header | Inhalt |
|---|---|---|
| P | Maklervollmacht_Enthalten | ja / nein |
| Q | MV_Vollumfaenglich | ja / nein / teilweise / nicht_pruefbar (Schlagwort gefunden?) |
| R | MV_Einschraenkungen | Klartext: gefundene Einschränkung oder fehlendes Schlagwort |
| U | Unterschrift_Kunde | ja / nein / nicht_pruefbar (Pflicht — VN-Unterschrift) |
| V | Unterschrift_Makler | ja / nein / nicht_pruefbar (optional) |
| W | Auf_VN_Ausgestellt | ja / nein / nicht_pruefbar |
| X | Makler_Namentlich | ja / nein / nicht_pruefbar |

GPT füllt diese Felder anhand der Schlagwort-Suche und der flankierenden Pflichten aus den oben genannten Quellen.
