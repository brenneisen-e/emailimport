# 🚀 BGAV Hypercare - Standalone HTA Version

## Was ist die HTA-Version?

Die **HTA (HTML Application)** Version ist eine **echte All-in-One-Lösung**, die:

✅ **Direkt auf Outlook zugreift** - Kein separates Export-Script nötig
✅ **Direkt nach Excel exportiert** - Kein manueller Download
✅ **Als Windows-App läuft** - Kein Browser erforderlich
✅ **Eine einzige Datei** - Vollständig standalone
✅ **Komplett offline** - Keine Internetverbindung nötig

## 🆚 Unterschied zur HTML-Version

| Feature | HTML-Version | HTA-Version |
|---------|--------------|-------------|
| **Outlook-Zugriff** | ❌ Nein (benötigt PowerShell/VBA) | ✅ Ja, direkt |
| **Excel-Export** | ⚠️ Ja (XLSX-Download) | ✅ Ja, direkt in Excel |
| **Plattform** | Alle Browser | Nur Windows |
| **Läuft in** | Webbrowser | Windows-App |
| **Setup** | Keine | Keine |

## 📋 Voraussetzungen

- ✅ Windows (7, 8, 10, 11)
- ✅ Microsoft Outlook (installiert und mindestens einmal gestartet)
- ✅ Microsoft Excel (für direkten Export)
- ✅ Keine Admin-Rechte erforderlich

## 🎯 Installation

**Keine Installation erforderlich!**

1. **Datei herunterladen**: `bgav-hypercare-standalone.hta`
2. **Doppelklick** auf die Datei
3. **Fertig!** Die App startet als Windows-Anwendung

### ⚠️ Sicherheitswarnung beim ersten Start

Beim ersten Ausführen erscheint möglicherweise eine Windows-Sicherheitswarnung:

```
"Möchten Sie diese Datei ausführen?"
```

**Lösung:** Klicke auf **"Ausführen"** oder **"Ja"**

Das ist normal, da HTA-Dateien erweiterte Rechte haben (um auf Outlook/Excel zuzugreifen).

## 🚀 Verwendung

### Schritt 1: App starten

Doppelklick auf `bgav-hypercare-standalone.hta`

Die App öffnet sich als Windows-Fenster (nicht im Browser).

### Schritt 2: Emails laden

Du hast **zwei Optionen**:

#### Option A: Ausgewählte Emails laden

1. **In Outlook**: Wähle eine oder mehrere Emails aus
   - Einzelne Email: Klick
   - Mehrere Emails: Strg + Klick
   - Bereich: Shift + Klick

2. **In der HTA-App**: Klicke auf **"📥 Ausgewählte Emails laden"**

3. **Fertig!** Die Emails werden automatisch geladen und kategorisiert

#### Option B: Ganzen Ordner laden

1. **In der HTA-App**: Klicke auf **"📂 Ordner laden"**

2. **Ordner auswählen**: Ein Dialog zeigt alle deine Outlook-Ordner
   - Wähle den gewünschten Ordner aus
   - Die Anzahl der Emails wird angezeigt

3. **"Ordner laden" klicken**

4. **Warten**: Bei großen Ordnern (>1000 Emails) kann das einige Minuten dauern

### Schritt 3: Review & Kategorisierung

Die App kategorisiert alle Emails automatisch:

- **Kategorie**: Incident oder Fachliche Rückfrage
- **Cluster**: SHUK, LV, KV, Provisionierung, etc.
- **Agentur**: Aus Absender-Name extrahiert
- **Status**: Neu (Standard)

**Änderungen vornehmen:**

1. **Einzelne Email**: Checkbox neben Feld aktivieren → Wert ändern
2. **Mehrere Emails**:
   - Emails auswählen (Checkboxen)
   - "✏️ Mehrere bearbeiten" klicken
   - Gewünschte Änderungen vornehmen

**Filter & Suche:**

- **Suchfeld**: Suche in Betreff und Text
- **Kategorie-Filter**: Nur Incidents oder Rückfragen
- **Cluster-Filter**: Nur bestimmte Cluster

### Schritt 4: Nach Excel exportieren

1. **Emails auswählen**: Wähle welche Emails exportiert werden sollen
   - "✓ Alle auswählen" für alle
   - Oder manuell einzelne Emails

2. **"💾 Nach Excel exportieren" klicken**

3. **Excel öffnet sich automatisch** mit den Daten im Hypercare-Format:
   - Alle Spalten korrekt befüllt
   - Formatierung angewendet
   - Bereit zum Speichern

4. **In Excel speichern**:
   - Datei → Speichern unter
   - Name: `Tracking_Hypercare_YYYY-MM-DD.xlsx`
   - Fertig!

## 💡 Tipps & Tricks

### Performance

**Kleine Mengen (<100 Emails):**
- Instant-Load, keine Wartezeit

**Mittlere Mengen (100-1000 Emails):**
- Lädt in wenigen Sekunden
- Fortschrittsanzeige wird angezeigt

**Große Mengen (>1000 Emails):**
- Kann 1-5 Minuten dauern
- App fragt um Bestätigung
- Besser: Mehrere kleinere Exports

### Empfohlener Workflow

```
1. Hypercare-Ordner in Outlook anlegen
   ↓
2. Emails in Ordner sammeln (täglich/wöchentlich)
   ↓
3. HTA-App öffnen
   ↓
4. "📂 Ordner laden" → Hypercare-Ordner auswählen
   ↓
5. Review & ggf. Anpassungen
   ↓
6. "💾 Nach Excel exportieren"
   ↓
7. Excel-Datei speichern
   ↓
8. Fertig!
```

### Batch-Kategorisierung

Wenn viele Emails die gleiche Kategorie/Cluster haben:

1. **Alle auswählen**: "✓ Alle auswählen"
2. **Mehrere bearbeiten**: "✏️ Mehrere bearbeiten"
3. **Einstellung ändern**: z.B. Cluster → SHUK
4. **Fertig!** Alle ausgewählten Emails werden aktualisiert

### Outlook-Ordner-Struktur

Empfohlen:

```
📁 Posteingang
  └─ 📁 Hypercare
      ├─ 📁 2025-11
      ├─ 📁 2025-10
      └─ 📁 Archiv
```

So kannst du gezielt einzelne Monate laden.

## ⚙️ Erweiterte Einstellungen

### HTA-Datei auf Desktop legen

1. Rechtsklick auf `bgav-hypercare-standalone.hta`
2. "Verknüpfung erstellen"
3. Verknüpfung auf Desktop ziehen
4. Optional: Umbenennen zu "Hypercare Email Tool"

Jetzt kannst du die App mit einem Doppelklick starten!

### Automatischer Start

Um die App beim Windows-Start automatisch zu öffnen:

1. Verknüpfung erstellen (siehe oben)
2. Verknüpfung kopieren
3. Windows-Taste + R → `shell:startup`
4. Verknüpfung in den Autostart-Ordner einfügen

**Achtung:** Die App öffnet sich dann bei jedem Windows-Start!

## 🔧 Troubleshooting

### "Outlook konnte nicht gestartet werden"

**Ursache:** Outlook ist nicht installiert oder wurde noch nie gestartet

**Lösung:**
1. Stelle sicher dass Outlook installiert ist
2. Starte Outlook mindestens einmal manuell
3. HTA-App neu starten

### "Excel konnte nicht gestartet werden"

**Ursache:** Excel ist nicht installiert

**Lösung:**
- Installiere Microsoft Excel
- Oder nutze die HTML-Version (exportiert als XLSX-Download)

### "Fehler beim Laden der Emails"

**Mögliche Ursachen:**
- Keine Emails ausgewählt → Wähle Emails aus
- Ordner ist leer → Wähle anderen Ordner
- Outlook ist geschlossen → Starte Outlook

**Lösung:**
- Fehlermeldung genau lesen
- Outlook neu starten
- HTA-App neu starten

### App öffnet sich nicht / Leeres Fenster

**Ursache:** Internet Explorer Kompatibilitätsmodus

**Lösung:**
1. Rechtsklick auf HTA-Datei → Eigenschaften
2. Unter "Sicherheit": "Zulassen" aktivieren
3. Übernehmen → OK
4. HTA-Datei neu öffnen

### Performance-Probleme bei vielen Emails

**Symptom:** App wird langsam bei >2000 Emails

**Lösung:**
- Lade kleinere Batches (z.B. pro Monat)
- Oder nutze Filter um Emails zu reduzieren
- Oder nutze PowerShell-Script für Massenexport

### "Skriptfehler" beim Start

**Ursache:** HTA wurde als normale HTML-Datei geöffnet

**Lösung:**
- Stelle sicher dass Dateiendung `.hta` ist (nicht `.html`)
- Öffne die Datei durch Doppelklick (nicht über Browser)

## 🔒 Sicherheit & Datenschutz

### Was macht die HTA-App?

Die App:
- ✅ Greift nur auf Outlook-Emails zu (mit deiner Erlaubnis)
- ✅ Verarbeitet Emails lokal auf deinem PC
- ✅ Erstellt Excel-Dateien lokal
- ❌ Sendet **KEINE** Daten ins Internet
- ❌ Speichert **KEINE** Daten auf Servern
- ❌ Verbindet sich **NICHT** mit externen Diensten

### Warum braucht die HTA erweiterte Rechte?

HTA-Dateien laufen außerhalb der Browser-Sandbox um:
- Auf COM-Objekte (Outlook, Excel) zuzugreifen
- Lokale Dateien zu erstellen
- Mit Windows-Anwendungen zu interagieren

**Das ist sicher**, solange du die Datei aus vertrauenswürdiger Quelle hast (dieses Repository).

### Kann ich den Code überprüfen?

**Ja!** Die HTA-Datei ist eine Textdatei:

1. Rechtsklick auf `bgav-hypercare-standalone.hta`
2. "Bearbeiten mit" → Notepad/Editor
3. Du siehst den gesamten HTML/JavaScript-Code
4. Alles ist transparent und überprüfbar

## 📊 Vergleich: Wann welche Version?

### Nutze die **HTA-Version** wenn:

✅ Du auf Windows arbeitest
✅ Du regelmäßig Outlook-Emails verarbeiten musst
✅ Du direkt in Excel exportieren möchtest
✅ Du eine One-Click-Lösung bevorzugst

### Nutze die **HTML-Version** wenn:

✅ Du auf Mac/Linux arbeitest
✅ Du verschiedene Email-Quellen hast (nicht nur Outlook)
✅ Du mehr Kontrolle über den Export-Prozess möchtest
✅ Du die Daten erst prüfen willst bevor du exportierst

### Nutze **PowerShell + HTML** wenn:

✅ Du sehr große Datenmengen hast (>5000 Emails)
✅ Du automatisierte Exports einrichten möchtest
✅ Du den Export-Prozess scripten möchtest

## 🎓 Häufig gestellte Fragen (FAQ)

### Kann ich die HTA-Datei umbenennen?

**Ja!** Aber behalte die `.hta` Endung:
- ✅ `Hypercare-Tool.hta`
- ✅ `Email-Review.hta`
- ❌ `bgav-hypercare.html` (funktioniert nicht!)

### Kann ich mehrere Ordner gleichzeitig laden?

**Nein**, aber du kannst:
1. Ersten Ordner laden
2. Kategorisieren
3. Zweiten Ordner laden (wird hinzugefügt)
4. Zusammen exportieren

### Werden die Emails aus Outlook gelöscht?

**Nein!** Die App:
- Liest nur Emails
- Ändert nichts in Outlook
- Verschiebt nichts
- Löscht nichts

### Kann ich die Daten speichern und später weitermachen?

**Nein**, die HTA-App speichert nichts automatisch.

**Workflow:**
1. Emails laden
2. Review durchführen
3. Sofort nach Excel exportieren
4. Excel-Datei speichern

Wenn du die App schließt, sind die Daten weg.

### Kann ich die Excel-Vorlage anpassen?

**Ja!** Im Code (Zeile ~1100+):

1. Öffne HTA-Datei mit Editor
2. Suche nach `function exportToExcel()`
3. Ändere `headers` Array für andere Spalten
4. Ändere Zuordnungen für andere Daten
5. Speichern & HTA neu starten

## 📞 Support

Bei Problemen:

1. **Lies diese Anleitung** - Meistens ist die Lösung hier
2. **Prüfe Fehlermeldung** - Oft steht dort die Lösung
3. **Outlook/Excel neu starten** - Hilft in 80% der Fälle
4. **HTA-App neu starten** - Ein frischer Start hilft oft

## 🎉 Zusammenfassung

Die **HTA-Version** ist die **einfachste und schnellste** Lösung für den Outlook → Excel Workflow:

```
📧 Outlook (Emails auswählen)
        ↓
🖥️ HTA-App (Laden & Review)
        ↓
📊 Excel (Direkter Export)
        ↓
✅ Fertig!
```

**Alles in einer Datei. Alles offline. Alles lokal.**

---

**Viel Erfolg mit dem BGAV Hypercare Email Review Tool! 🚀**
