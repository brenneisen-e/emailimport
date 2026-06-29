# KI-Deckblatt — aktueller Inhalt

> Stand: automatisch erzeugt aus der Funktion `wsBuildNotizHtml(data, lfdNr)`
> in `ergo-vorgang-analyse.hta` (ab Zeile ~10044, identisch zu
> `ergo-deckblatt-msg-code.txt`). Das Deckblatt ist **kein Prompt**, sondern
> wird lokal je Vorgang aus den von der KI gelieferten JSON-Feldern als
> 1-seitiges HTML gerendert und anschließend per Word zu PDF gemacht und in
> die Original-`.msg` eingebettet.
>
> Diese Datei beschreibt **genau das, was der Code aktuell tatsächlich
> ausgibt** (Reihenfolge, Labels, Bedingungen). Felder mit dem Hinweis
> *(nur wenn befüllt)* werden nur gerendert, wenn der jeweilige Wert nicht
> leer/„k.A."/„-" ist (siehe `clean()`).

---

## Kopfzeile

| Element | Inhalt |
|---|---|
| **Titel** | `KI-Vorgangs-Analyse` |
| **Lfd-Nr (Badge oben rechts)** | rotes Kästchen `Lfd-Nr <Wert>` (nur wenn `lfdNr` gesetzt) |
| **Untertitel** | Verweis: identisch zur Spalte `Lfd_Nr` in der Excel-Auswertung (Mapping Excel ↔ Vorgang), „Farbgruppen wie in der Excel" |

### Badge-Leiste (direkt unter dem Titel)

In dieser Reihenfolge:

1. **BÜ-Ampel-Badge**
   - `Neue BÜ` (grün) — Makler-Vorgang + Klassifikation `BUe-Vorgang`, kein Reminder
   - `BÜ-Reminder` (gelb) — wie oben, aber Reminder
   - `Keine BÜ` (rot) — kein Makler-Vorgang bzw. keine BÜ-Klassifikation
2. **Typ-Badge** — `Makler` (blau) oder `MFA` (gelb, Mehrfachagent); nur wenn ableitbar
3. **Klassifikations-Badge** (blau) — Wert aus `klassifikation` bzw. `(leer)`
4. **Reminder-Badge** (gelb) — `Reminder`, nur wenn `ist_reminder = ja`
5. **Maklervollmacht-Badge** — `Maklervollmacht: <Status>`
   - grün bei `OK` / `OK (MFA)`
   - rot sonst (z.B. `MV unvollständig`, `MV fehlt`, `MFA – VN-Unterschrift prüfen`)
   - Bei MFA zählt die **Kundenunterschrift**: `unterschrift_kunde = ja` → `OK (MFA)`, sonst `MFA - VN-Unterschrift pruefen`

---

## Sektionen (in dieser Reihenfolge auf dem Deckblatt)

### 1. Kunde & Vertrag — blau

| Label | Quelle / Hinweis |
|---|---|
| Kunde (Name) | Vor- + Nachname VN bzw. `kunde_name` |
| Geburtsdatum VN | *(nur wenn befüllt)* |
| Kunde-Adresse | *(nur wenn befüllt)* |
| Versicherungsnummer | `versicherungsnummer` |
| Sparte | `sparte` |
| Kunden-/Partnernummer | *(nur wenn befüllt)* — KDNR, **nicht** Versicherungsnummer |
| VN laut MV | *(nur wenn befüllt)* — Name aus Maklervollmacht (Quervergleich) |

### 2. Makler — teal

| Label | Quelle / Hinweis |
|---|---|
| Pool | Maklerpool (kanonisch) |
| Vermittler | *(nur wenn befüllt)* — Untervermittler, sonst Makler-Person |
| **Agentur-/Personalnummer** | **hervorgehoben (orange)** — Agentur-Nr. und/oder Personalnummer |
| Makler-Adresse | *(nur wenn befüllt)* |
| Auftrag-Datum | *(nur wenn befüllt)* — Datum des Anschreibens |

### 3. Klassifikation & Vorgang — rot

| Label | Quelle / Hinweis |
|---|---|
| Vorgangstyp | z.B. Makler-Vorgang / Zustellfehler / Ergo-Outbound / System-Mail / Werbung-Spam / Unklar |
| Klassifikation | z.B. BUe-Vorgang / Schadenmeldung / Antrag-Änderung / Anfrage-Rücksprache / Nicht-Standard / Kein-Makler-Vorgang |
| BÜ-Wunsch vorhanden | `bue_wunsch_vorhanden` |
| Prüfung Unterlagen | *(nur bei echten BÜ-Vorgängen)* — abgeleitet: `Vollständig`, oder fehlende Punkte (`BÜ-Wunsch fehlt`, `MV fehlt`, `MV unvollständig`), bei Reminder: `Reminder - keine Vollst.-Prüfung` |
| Reminder/Wiedervorlage | `ja` / `nein` |
| Reminder-Quelle | *(nur wenn Reminder = ja)* — Mail-Betreff / Mail-Body / Anhang-Dateiname / Anhang-PDF-Inhalt |

### 4. Maklervollmacht & Unterschriften — lila

| Label | Quelle / Hinweis |
|---|---|
| MV enthalten / Status | `enthaelt_maklervollmacht` + `(mv_status)` |
| MV auf Makler ausgestellt | *(nur wenn befüllt)* |
| MV vollumfänglich | ja / teilweise / nein / nicht prüfbar |
| MV-Einschränkungen | *(nur wenn befüllt)* — Klartext, gekürzt auf max. 140 Zeichen |
| Unterschrift Kunde / Makler | `unterschrift_kunde` / `unterschrift_makler` |
| Untervollmacht erteilt | *(nur wenn befüllt und ≠ „nein")* |

### 5. Ergebnis — grün

| Box | Inhalt |
|---|---|
| **Zusammenfassung** (grüne Box) | `mail_zusammenfassung`, gekürzt auf max. 280 Zeichen — 1–2 Sätze, neutral, was die Mail macht |
| **Hinweis** (gelbe Box) | `hinweis`, gekürzt auf max. 220 Zeichen — *(nur wenn befüllt)*, Auffälligkeiten |

---

## Fußzeile

```
Erstellt durch ERGO Vorgang-Analyse · <Zeitstempel>
```

---

## Technik

| Aspekt | Wert |
|---|---|
| Format | A4 hoch, 10 mm Rand, Schrift „Segoe UI" ~10,5 pt, eine Seite |
| Erzeugung | HTML → MS Word (COM) → PDF → eingebettet als Anlage `KI-Notiz_<Lfd-Nr>.pdf` in die Original-`.msg`; ohne Word Fallback als HTML-Anlage |
| Speicherung | Original-`.msg` wird unter `Vorgang_<Lfd-Nr>_<Betreff>.msg` neu gespeichert (Original bleibt führend), optional automatischer Forward an ein Zielpostfach |
| Encoding | Alle Nicht-ASCII-Zeichen werden am Ende als numerische HTML-Entities (`&#228;` usw.) kodiert, damit Word/Outlook charset-unabhängig korrekt rendern |
| Datenbasis | strukturierte KI-Felder aus `buildVorgangPrompt` (ki_prompt.txt) |
