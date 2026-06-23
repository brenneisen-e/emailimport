# ERGO Vorgang-Analyse — Zusammenfassung

> Was das ERGO-Tool kann und was wir mit der Vollerhebung des BÜ-Posteingangs
> herausgefunden haben. Stand: HTA v2.53 · Auswertung KW 13.–20.04.2026.

---

## 1. Worum geht es?

Im ERGO-Maklergeschäft landen **Bestandsübertragungs-Anfragen (BÜ)** von
Maklern und Maklerpools heute als **unstrukturierte E-Mails** in einem
Sammel-Postfach. Die Daten in Mail und Anhängen (Versicherungsnummer, Sparte,
Pool, Maklervollmacht-Status) stehen an wechselnden Stellen und bleiben
ungenutzt; zusammengehörige Vorgänge sind im Postkorb nicht auffindbar.

Das **ERGO Vorgang-Analyse-Tool** ist ein KI-gestützter Proof of Concept, der
jede Eingangsmail **strukturiert**: Es liest Mail + PDF-Anhänge, schickt sie an
**ErgoGPT** (interner ERGO-Dienst, freigegeben), extrahiert **42 Datenfelder**
pro Mail und schreibt das Ergebnis als **50-spaltige Excel-Tabelle** plus
optionalem **KI-Deckblatt (PDF)** pro Vorgang. Damit werden drei Dinge möglich,
die heute fehlen: korrektes **Routing**, eine **Cluster-Sicht** (welche Mails
gehören zum selben Fall) und eine **Vorqualifizierung** (ist die Maklervollmacht
da, unterschrieben, vollumfänglich?).

Einordnung im Projekt: Das Tool ist die **Interim-Lösung „PoC / ERGO-GPT"**
(halbautomatisch). Die Ziel­architektur ist ein vollautomatischer
**Strukturierungs-Bot** vor dem Postkorb.

---

## 2. Die Werkzeuge

Das ERGO-Toolset besteht aus mehreren zusammenspielenden Teilen:

| Datei | Typ | Zweck |
|-------|-----|-------|
| **`ergo-vorgang-analyse.hta`** | HTML Application (Windows) | Hauptwerkzeug mit UI: Outlook-Mails sammeln, an ErgoGPT analysieren, Excel + KI-Deckblatt erzeugen. **v2.53** |
| **`ergo-vorgang-analyse.bas`** | Excel-VBA-Modul | Variante als Excel-Makro: scannt einen `.msg`-Ordner, gleiche Analyse, mit **Parallel-Modus** (v2.13) |
| **`ergo-email-batch.hta` / `.bas`** | HTA / VBA | „Einsammeln": Outlook-Ordner → Tabelle/`.msg`/`.eml`/JSON. Read-only, keine KI. **v1.1** |
| **`hta-source.html`** | Quellcode-Seite | Liefert den HTA-Quellcode inline (Textarea + Copy/Download), umgeht Corporate-Proxy-Filter. Ersetzt ZIP/`.hta.txt`-Auslieferung |
| **`claude-analyse.html`** | Browser-KI (Standalone) | Alternative ohne Outlook: `.msg` per Drag-Drop, Analyse direkt gegen `api.anthropic.com` (nur für nicht-produktive Testdaten) |

Begleitende Doku im Repo:
`ANLEITUNG-ERGO-VORGANG-ANALYSE.txt`, `ANLEITUNG-ERGO-EXCEL.txt`,
`Prompt-1-KI-Analyse.txt` (Analyse-Prompt), `Prompt-2-KI-Deckblatt.txt`
(Deckblatt-Schema), `Prompt-3-Excel.txt` (Spalten-Schema).

---

## 3. Was das Tool kann (Fähigkeiten)

### 3.1 Workflow

```
.msg-Ordner / Outlook-Postfach
        │
        ▼  pro Mail (ein setTimeout-Tick, kein Einfrieren)
  Outlook öffnet die Mail  →  Anhänge extrahieren
        │                     (max. 3 PDFs, je < 20 MB)
        ▼
  Mail-Body (auf 6.000 Zeichen gekappt) + PDFs  →  HOCHLADEN an ErgoGPT
        │
        ▼
  JSON-Antwort (42 Felder)  →  1 Zeile in Sheet 'Analyse'
        │                       (Auto-Save alle 3 Vorgänge)
        ▼
  optional: KI-Deckblatt (HTML → Word → PDF) in die Original-.msg einbetten
```

- Verarbeitet werden **nur `.msg`** (verlustfrei, mit Anhängen). `.eml` nicht,
  da Anhang-Extraktion ohne Outlook-MAPI unzuverlässig ist.
- **Anzahl-Limit** beim Start wählbar (z. B. 10 zum Testen, leer = alle).
- **Fehlertoleranz**: Schlägt eine Mail fehl, bekommt die Zeile einen
  `[FEHLER]`-Hinweis und das Tool läuft weiter.

### 3.2 ErgoGPT-Integration (technisch)

- HTTP über `MSXML2.ServerXMLHTTP` bzw. `WinHttp`, Endpunkt **`gpt.ergo.com`**.
- **Cookie-Auth**: Cookie aus Datei (`F:\ExcelGPT-Cookie\Cookie.txt`),
  Sheet-Zelle oder Start-Dialog; XSRF/CSRF-Token werden abgeleitet.
  Egal welche Quelle — der Cookie wird immer an den Standard-Pfad gespiegelt.
- **PDF-Upload**: Konversation anlegen → bis zu 3 PDFs hochladen → Prompt mit
  Anhängen abschicken → NDJSON-Antwort parsen → Konversation wieder löschen.
- **Default-Modell**: `gpt-5.1` (in Sheet `GPT!A6`, Temp 0). ErgoGPT setzt
  unter der Haube **Claude Opus** (4.7) ein.

### 3.3 Die 42 KI-Felder (Auszug)

Der Prompt macht den Sachbearbeiter im Innendienst nach. Zentrale Logik:

- **Schritt 0 – Triage** (`vorgangstyp`): Ist es überhaupt ein echter
  Makler-Vorgang? Zustellfehler / Ergo-Outbound / System-Mail / Werbung →
  alle anderen Felder bleiben leer. **Es wird nie geraten.**
- **Quellentrennung** (entscheidend für Genauigkeit): Anschreiben/BÜ-Wunsch und
  Maklervollmacht werden **getrennt** ausgewertet und nicht vermischt — so
  bleiben „Makler laut Anschreiben" und „Makler laut MV" vergleichbar
  (Falschrouting-/Plausibilitäts-Check).
- **Klassifikation**: BÜ-Vorgang / Schadenmeldung / Antrag-Änderung /
  Anfrage-Rücksprache / Nicht-Standard.
- **Geschäfts-Typ**: BÜ einfacher Vertrag vs. BÜ Kundenverbindung (mehrere
  Verträge / „gesamte Kundenverbindung").
- **Sparte** mit VNR-Präfix als Indikator — **Achtung: `KR` = Kraftfahrt
  (Komposit), NICHT Kranken; `KV` = Krankenversicherung.**
- **Maklerpool** wird auf **kanonische Schreibweisen** konsolidiert (Map mit
  ~35 Varianten, z. B. „Jung DMS / JDC" → `JDC / Jung, DMS & Cie.`).
- **Pool vs. Untervermittler** getrennt geführt (z. B. Pool = JDC,
  Untervermittler = Finanzguru).
- **Maklernummer-Priorisierung**: Agentur-Nr. (beginnt oft mit `6000`/`890`)
  vor BD-/Bestandsnummer vor Pool-internen IDs.
- **Maklervollmacht-Prüfung**: enthalten? vollumfänglich (Schlagwort-basiert)?
  Einschränkungen? auf VN ausgestellt? Unterschrift Kunde/Makler sichtbar?
- **Reminder-Erkennung** (`ist_reminder`, cross-cutting) inkl. Quelle
  (Betreff / Body / Anhang-Dateiname / Anhang-PDF-Inhalt).

### 3.4 Excel-Output (50 Spalten, Sheet „Analyse")

Pro Mail eine Zeile. Spaltengruppen u. a.: `Cluster_ID`, `Status` (Ampel:
Neue BÜ / BÜ-Reminder / Keine BÜ), `Typ` (Makler / MFA), Versicherungsnummer,
Sparte, Kunde, Maklerpool/Untervermittler, Agentur-/Vermittler-Nr.,
Geschäfts-Typ, Klassifikation, MV-Status, Unterschriften, Hinweis,
Mail-Zusammenfassung, **abgeleitete** Felder `Segment`, `Auto_Cluster`,
`Vorgang_Alter_Tage`, `Makler_Key` sowie die Roh-Maildaten.

**Abgeleitete Logik (lokal, regelbasiert — kein KI-Output):**

- **Cluster_ID** (`001a/001b…` + Pastellfarbe): **gleicher Makler**
  (Agentur-/Vermittlernr.-Union, sonst Pool + Name) **UND gleicher Kunde**
  (Kundennr., sonst Name) → gemeinsame ID. Macht zusammengehörige Mails sichtbar.
- **MV_Status** (konsolidiert): OK / MV unvollständig / MV fehlt / kein Anhang /
  nicht prüfbar / Reminder (MV n. e.) — mit MFA-Sonderfällen.
- **Auto_Cluster** (Automatisierbarkeit):
  - **vollautomatisierbar**: BÜ-Vorgang + einfacher Vertrag + MV OK + Text-PDF
  - **teilautomatisierbar**: Mixed-PDF, MV unvollständig, Reminder,
    Kundenverbindung, Antrag-Änderung
  - **manuell**: Scan-PDF, MV fehlt, Sonderfall (Flotte/Sondertarif),
    Nicht-Standard, Nicht-Makler-Vorgang, KI-Fehler
- **Segment**: BU-Erst-Einzeln / BU-Reminder-Kundenverb / Schaden / Antrag /
  Anfrage / Nicht-Standard / Zustellfehler / Nicht-Makler.

### 3.5 KI-Deckblatt (PDF)

Pro Vorgang wird aus den 42 Feldern lokal eine **1-seitige A4-Notiz** gerendert
(HTML → Word → PDF) und in die Original-`.msg` eingebettet. Sektionen:
Klassifikation, Makler, Kunde & Vertrag, Maklervollmacht & Unterschriften,
Ergebnis (Zusammenfassung + Hinweis). Status- und MV-Status-Badges oben.

### 3.6 Parallel-Modus (`.bas` ab v2.13)

Bis zu **6 ErgoGPT-Antwort-Calls gleichzeitig** (Default 3, empfohlen). Die
Vorbereitung (Outlook/COM) bleibt sequentiell — nur der langsame Antwort-Call
(15–25 s) läuft parallel. Speedup ~2,5× bei N=3; bei 30 Vorgängen ~10 min
seriell → ~3–4 min. Bei N ≥ 5 häufiger 403/429 (Rate-Limit).

### 3.7 Datenschutz

Mail-Inhalte und PDFs gehen an **ErgoGPT (intern, freigegeben)** — zulässig für
produktive ERGO-Daten. Die Browser-Variante `claude-analyse.html` schickt
dagegen an Anthropic (USA) und ist **nur für Testdaten** gedacht.

---

## 4. Was wir damit herausgefunden haben (Mengengerüst)

> **Datenbasis:** Vollerhebung des BÜ-Posteingangs einer Kalenderwoche
> **Mo 13.04.–Mo 20.04.2026 = 5.034 Mails**. 94 % über ein Sammel-Postfach
> („Service Team Makleraufträge"), 5 % „Bestandsübertragungen", 12 weitere < 1 %.
> KI-strukturiert mit **Claude Opus 4.7**, 50 Felder/Mail; 99 Mails mit
> Analyse-Fehler/Timeout. Hochrechnung **~260.000 Mails/Jahr**.
> Definition: *Vorgang* = Cluster aus Mails zum selben Makler + selben Kunden.

### 4.1 Status-Split der 5.034 Mails/Woche

| Status | Mails/Wo | Anteil | Bedeutung |
|--------|---------:|-------:|-----------|
| Keine BÜ (falscher Kanal) | 705 | 14,0 % | Schaden/Antrag/Anfrage — gehört nicht ins Postfach |
| BÜ-Reminder | 329 | 6,5 % | Sachstandsanfrage zu Vorgang aus früherer Woche |
| **Neue BÜ (echte BÜ-Anfragen)** | **4.000** | **79,5 %** | echte Bestandsübertragungs-Mails |
| ⮑ davon Nachreichungen | 1.505 | 29,9 % | Folge-Mails zu schon in dieser Woche geöffnetem Vorgang |
| ⮑ davon **echte Neuanlagen** | **2.495** | **49,6 %** | zu eröffnende neue Fälle |

**Kernbefund:** Nur **jede 2. Mail ist eine echte Neuanlage**; **jede 3. Mail
(30 %) ist Nachreichung** zu einem schon offenen Vorgang.

### 4.2 Mail-Perspektive ≠ Vorgangs-Perspektive

Die **4.000 BÜ-Anfrage-Mails** verteilen sich auf nur **2.495 Vorgänge**:

- **1.883 Einzelmails (47 %)** — je ein eigener Vorgang
- **2.117 Mails (53 %)** zu nur **612 Vorgängen** — derselbe Makler schreibt
  mehrfach zum selben Kunden (**Cluster**)
- Ø **1,60 Mails/Vorgang** (Securess Spitze mit 4,4)

→ **25 % der Vorgänge (die 612 Cluster) erzeugen 100 % der Nachreichungen**
(im Schnitt 3,4 Mails je Cluster-Vorgang). Größte Cluster: 16 Mails
(Cluster 005), 15 (Cluster 509), 14 (163 + 211), 13 (484 + 585).

### 4.3 Sparten-Split

| Sparte | BÜ-Anfragen (Mail) | davon in Cluster |
|--------|-------------------:|-----------------:|
| Komposit | 2.089 | 1.405 (67 %) |
| KV | 1.147 | 354 (31 %) |
| Leben | 662 | 325 (49 %) |
| Sonstige | 102 | 33 (32 %) |
| **Gesamt** | **4.000** | **2.117 (53 %)** |

**Befund:** Auf **Mail-Ebene** dominiert Komposit (~52 %), weil Komposit-Vorgänge
oft Multi-Mail-Cluster sind. Auf **Vorgangs-Ebene** ist **KV (38 %) fast
gleichauf mit Komposit (41 %)** — KV-Vorgänge werden meist in 1 Mail erledigt.
→ **KV ist unterschätzt, wenn man nur Mails zählt** (KV hat kein VNR-Mapping,
daher jede Mail = eigener Vorgang).

### 4.4 Maklerpool-Konzentration

| Maklerpool | BÜ-Anfragen/Wo | Vorgänge | Ø Mails/Vorgang |
|------------|---------------:|---------:|----------------:|
| JDC / Jung, DMS & Cie. | 1.538 | 913 | 1,68 |
| Fonds Finanz | 612 | 291 | 2,10 |
| blau direkt | 255 | 180 | 1,42 |
| DEMA | 158 | 124 | 1,27 |
| vfm Service | 116 | 77 | 1,51 |
| **Top-5 (65 %)** | **2.594** | **1.519** | **1,71** |
| Sonstige (Long-Tail) | 1.406 | 982 | 1,43 |
| **Gesamt** | **4.000** | **2.495** | **1,60** |

**Befund:** **65 % der BÜ-Anfragen kommen von 5 Maklern**, die Top-3
(JDC + Fonds Finanz + blau direkt) machen **60 %** aus — **JDC allein 38 %**.
→ Eine Automatisierungs-/Anbindungsstrecke für die Top-Pools ist der größte
Hebel auf einen Schlag.

### 4.5 Die zwei Hebel

**Hebel 1 — Routing (Fehl-Routing durch „MV → BÜ").**
Die Eingangsregel „Maklervollmacht im Anhang → BÜ" greift unabhängig vom Status
der Mail. Dadurch landen **240 Mails/Woche fälschlich im BÜ-Postkorb**
(**210 Reminder + 30 fachfremd**, jeweils mit MV im Anhang) → Fehlanlagen,
Dubletten, Korrekturaufwand. **Die KI liest den echten Status und routet
korrekt.**

**Hebel 2 — Clustering & Arbeitsvorrats-Steuerung.**
4.000 BÜ-Mails bündeln sich zu 2.495 Vorgängen; 65 % von Top-5-Maklern, 53 % in
Clustern — heute im Postkorb **nicht filterbar**. Mit strukturierten Daten +
Excel können Sachbearbeiter gezielt nach **Prio-Makler** oder **Cluster**
filtern und gebündelt abarbeiten. Mail-Zuordnung zum offenen Vorgang spart
**1.505 Nachreichungs-Eröffnungen/Woche**.

### 4.6 Vier Herausforderungen im Eingang (heute messbar)

1. **Routing**: Regel „MV → BÜ" statusunabhängig → 240 Fehl-Routings/Wo.
2. **Unstrukturierte Daten**: VNR/Pool/Status/MV an wechselnden Stellen →
   manuelle Suche je Vorgang.
3. **Keine Cluster-Bildung**: 612 Cluster mit 2.117 Mails werden einzeln
   bearbeitet; Top-5 = 65 %, aber nicht filterbar.
4. **Keine Vorqualifizierung**: MV vorhanden/unterschrieben/vollumfänglich? wird
   erst in der Sachbearbeitung geklärt.

Die KI-Strukturierung adressiert **alle vier zugleich**.

---

## 5. Lösungsoptionen & Empfehlung (Zwischenbericht Phase 1)

| Option | Einordnung | Kern |
|--------|-----------|------|
| Makler-Formular | **nicht empfohlen** | Web-Formular statt Mail — zu hoher Change-Aufwand bei Maklern |
| **PoC / ERGO-GPT** | **Interim** | *Dieses Tool*: KI-Vorqualifizierung halbautomatisch, sofort einsetzbar, Excel + Deckblatt akzeptiert |
| **Strukturierungs-Bot** | **EMPFEHLUNG (Zielarchitektur, Q4)** | Vollautomatischer KI-Bot mit Beiblatt vor/nach Postkorb (BluePrism-Anbindung); Grundlage für Reporting + Prüfbot |
| Celonis | TBD | Process Mining auf ERGO-Standard, aber frühestens 2028 verfügbar |

**Architektur — ein Bot, zwei Andockpunkte:**
- **Option 1 (Präferenz):** Bot **vor** dem Postkorb über neues Gruppenpostfach
  → greift zusätzlich ins **Routing** ein (240 Fehl-Routings/Wo fallen weg),
  Postkorb wird sauberer.
- **Option 2 (Fallback):** Bot **im** Postkorb → Cluster-Sicht und
  Vorqualifizierung trotzdem verfügbar, aber Routing bleibt wie heute.

**Praxis-Skizze — eine Mail, drei Outputs:** (1) korrektes Postfach (Routing),
(2) aktualisierter Excel-Cluster (Steuerung), (3) 50 maschinenlesbare Felder als
Basis für einen nachgelagerten **Prüfbot** (BiPRO/Dunkelverarbeitung).

---

## 6. Technische Eckdaten & Stand

- **Versionen:** HTA `ergo-vorgang-analyse.hta` = **v2.53** (Entschlackung:
  KPI-Sheets aus, Prompt gekürzt, Standalone-Workshop entfernt) ·
  VBA `.bas` = **v2.13** (Parallel-Modus) · `ergo-email-batch` = **v1.1**.
- **Pool-Erkennung:** ~17 hinterlegte Pool-Regeln (Domain → Betreff → Signatur)
  plus ~35er Konsolidierungs-Map im Prompt; Pool-Extraktion ist seit v2.10
  „ergebnisoffen" (keine Whitelist-Filterung).
- **KPI-Sheets** (Cluster, Pool-Konzentration, Eingangskanal, Business-Case/ROI)
  sind im Code vorhanden, aber ab **v2.53 per Flag deaktiviert**
  (`BUILD_KPI_SHEETS = false`, Anwender-Wunsch — schlankere Excel).
- **Auslieferung HTA:** über `hta-source.html` (Inline-Quellcode, umgeht
  Proxy-Filter) statt ZIP/`.hta.txt`.

## 7. Limits / offene Punkte

- Pro Vorgang max. **3 PDFs**, je **< 20 MB**; Body auf **6.000 Zeichen** gekappt.
- **Scan-PDFs ohne Textebene** werden nicht analysiert (manuell prüfen).
- Nur **`.msg`**, kein `.eml` (Anhang-Extraktion).
- **KV-Cluster**: ohne VNR-Mapping wird jede KV-Mail ein eigener Vorgang.
- Bei Parallel-Modus N ≥ 5 vermehrt Rate-Limit-Fehler (403/429).
- 99 von 5.034 Mails der Vollerhebung mit Analyse-Fehler/Timeout (~2 %).

---

*Datenbasis der Erkenntnisse: KI-Auswertung Posteingang „Service Team
Makleraufträge", 13.–20.04.2026 (5.034 Mails, eine Kalenderwoche), Analyse-Stand
27.05.2026, Claude Opus 4.7. Quellen im Repo: `ERGO-Mengengeruest-2Slides*.pptx`,
`Prompt-1-KI-Analyse.txt`, `Prompt-2-KI-Deckblatt.txt`, `Prompt-3-Excel.txt`,
`ANLEITUNG-ERGO-VORGANG-ANALYSE.txt`, `ergo-vorgang-analyse.hta/.bas`.*
