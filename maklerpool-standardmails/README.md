# Maklerpool-Standardmails

Sammelstelle für **echte Beispiel-Mails** von Maklerpools, die sich von Pool zu Pool extrem ähneln (Briefkopf-Template, identische Formulierungen, gleiche Anhang-Struktur etc.).

## Warum?

Bei den ersten 18-100 analysierten Vorgängen wird sichtbar: ~60-80 % der Pool-Mails folgen einer Handvoll Templates. Aktuell wird **jede** Mail durch den vollen ErgoGPT-Lauf geschickt (Klassifikation + Triage + MV-Prüfung + 28 Felder, ~20 s pro Vorgang).

Wenn das Tool ein Template **vorher** erkennt (per Subject-Pattern, Absender-Domain, Anhang-Namensmuster, Body-Schlüsselwörter), kann es:

- **Felder pre-fillen** mit den schon bekannten Werten (Pool-Name, Vorgangstyp, Klassifikation, etc.)
- **Nur die variablen Felder** an GPT schicken (oder ganz darauf verzichten, wenn alles im Template steht)
- **Datenschutz-Vorteil**: weniger Mailinhalte verlassen das Haus

Geschätzter Speedup: 5-10x bei Standard-Pool-Mails (sub-sekunde statt 20 s).

## Struktur

Pro Pool ein Unterordner mit echten Beispielen:

```
maklerpool-standardmails/
├── DEMA/                       # DEMA Deutsche Versicherungsmakler AG
│   ├── BUe-Anschreiben-1.msg
│   ├── Mahnung-Anschreiben-1.msg
│   └── Pattern.md              # Beschreibung der Erkennungsmerkmale
├── BCA/                        # BCA AG
├── Fonds-Finanz/               # Fonds Finanz Maklerservice GmbH
├── JDC/                        # JDC Group AG
├── Maxpool/                    # Maxpool Maklerkooperation GmbH
└── sonstige/                   # Pools mit nur 1-2 Beispielen
```

## Erkennungsmerkmale (pro Pool in `Pattern.md`)

Für jeden Pool dokumentieren:

- **Absender-Domain** (z.B. `@dema-ag.de`, `@fondsfinanz.de`)
- **Subject-Pattern** (Regex oder Substring, z.B. `^\[EXTERN\] Dokumente von der DEMA Maklerabteilung`)
- **Anhang-Namen-Pattern** (z.B. `PDF_Maklerauftrag_(Anschr|Mahnung)_Gesell_\d{7}.*\.pdf`)
- **Body-Marker** (typische erste Zeilen, Briefkopf-Texte)
- **Default-Werte** für die 28 GPT-Felder, soweit ableitbar (Maklerpool, Geschäfts-Typ, Sparte etc.)

## Wichtig — Datenschutz

- **Keine echten Kundendaten** in das Repo. Vor dem Ablegen:
  - Kundennamen → `Max Mustermann`
  - Kundennummern → `XXX-XXXXXX`
  - Versicherungsnummern → `XX000000000`
  - Geburtsdaten → `01.01.1970`
  - Anschriften → `Musterweg 1, 12345 Musterstadt`
- Anonymisierung am besten **direkt in Outlook** vor dem `Save as .msg`
- Wer eine echte Mail anonymisiert hochlädt, fügt sich oben in einer Zeile als „Beigetragen: <Initialen> <Datum>" zu

## Nächste Schritte (Roadmap)

1. **Phase 1 (aktuell):** Sammeln. 5-10 Beispiele pro Pool, anonymisiert.
2. **Phase 2:** Template-Fingerprints generieren (subject regex, attachment regex, body keywords).
3. **Phase 3:** HTA + Excel-Tool prüfen Fingerprint **vor** dem GPT-Call; bei Treffer → instant fill, kein GPT-Call (oder nur Validierung).
4. **Phase 4:** Pool-Stats — wie viele Vorgänge pro Pool, welche Templates am häufigsten, wo lohnt sich Template-Engineering vs. LLM.
