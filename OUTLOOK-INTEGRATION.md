# Outlook Integration - Anleitung

Diese Anleitung zeigt dir, wie du Emails direkt aus Outlook exportieren und in die BGAV Hypercare Email Review App laden kannst.

## 📋 Übersicht

Es gibt **zwei Methoden** zum Exportieren von Outlook-Emails:

1. **PowerShell-Script** (empfohlen) - Einfach, schnell, keine Outlook-Konfiguration nötig
2. **VBA-Makro** - Direkt in Outlook integriert, kann als Button hinzugefügt werden

---

## 🚀 Methode 1: PowerShell-Script (Empfohlen)

### Voraussetzungen
- Windows mit PowerShell (bereits vorinstalliert)
- Microsoft Outlook installiert
- Keine Admin-Rechte erforderlich

### Installation

Keine Installation erforderlich! Script einfach ausführen.

### Verwendung

#### Variante A: Ausgewählte Emails exportieren

1. **In Outlook**: Wähle eine oder mehrere Emails aus (Strg+Klick für mehrere)
2. **PowerShell öffnen**: Rechtsklick auf `Export-OutlookEmails.ps1` → "Mit PowerShell ausführen"
   - Falls das nicht funktioniert: PowerShell manuell öffnen und eingeben:
   ```powershell
   cd "C:\Pfad\zum\Script"
   .\Export-OutlookEmails.ps1
   ```
3. **Speicherort wählen**: Ein Dialog öffnet sich - wähle wo die JSON-Datei gespeichert werden soll
4. **Fertig!** Die JSON-Datei kann jetzt in die HTML-App geladen werden

#### Variante B: Ganzen Ordner exportieren

```powershell
.\Export-OutlookEmails.ps1 -FolderPath "Posteingang\Hypercare"
```

Exportiert alle Emails aus dem angegebenen Ordner.

#### Variante C: Direkt zu Datei exportieren

```powershell
.\Export-OutlookEmails.ps1 -OutputPath "C:\Temp\emails.json"
```

Exportiert ausgewählte Emails direkt zur angegebenen Datei.

### Troubleshooting PowerShell

#### "Ausführung von Skripts ist auf diesem System deaktiviert"

Lösung 1 (Temporär, nur für diese Sitzung):
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Export-OutlookEmails.ps1
```

Lösung 2 (Permanent, für aktuellen Benutzer):
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Dann kannst du das Script normal ausführen.

#### "Outlook kann nicht gestartet werden"

- Stelle sicher, dass Outlook installiert ist
- Starte Outlook einmal manuell
- Führe das Script erneut aus

---

## 🎯 Methode 2: VBA-Makro

### Voraussetzungen
- Microsoft Outlook
- Makros müssen in Outlook aktiviert sein

### Installation

1. **Outlook öffnen**

2. **VBA-Editor öffnen**:
   - Drücke `ALT + F11`
   - Oder: Datei → Optionen → Menüband anpassen → Entwicklertools aktivieren → Entwicklertools → Visual Basic

3. **Neues Modul erstellen**:
   - Im VBA-Editor: Menü → Einfügen → Modul
   - Ein neues Modul-Fenster öffnet sich

4. **Code einfügen**:
   - Öffne die Datei `OutlookExportMacro.vba` in einem Texteditor
   - Kopiere den gesamten Code
   - Füge ihn in das Modul-Fenster ein

5. **Speichern**:
   - Drücke `Strg + S` zum Speichern
   - Schließe den VBA-Editor

### Verwendung

#### Variante A: Mit Tastenkombination

1. **Emails auswählen**: Wähle in Outlook eine oder mehrere Emails aus
2. **Makro ausführen**: Drücke `ALT + F8`
3. **Makro wählen**: Wähle "ExportSelectedEmailsToJSON" und klicke "Ausführen"
4. **Speicherort eingeben**: Gib den Pfad für die JSON-Datei ein
5. **Fertig!**

#### Variante B: Button im Ribbon hinzufügen (Optional)

1. **Outlook-Optionen öffnen**: Datei → Optionen → Menüband anpassen
2. **Neue Gruppe erstellen**:
   - Wähle "Start (E-Mail)" in der rechten Liste
   - Klicke "Neue Gruppe" → Benenne sie "Hypercare"
3. **Makro hinzufügen**:
   - Wähle links "Makros" aus dem Dropdown
   - Wähle "ExportSelectedEmailsToJSON"
   - Klicke "Hinzufügen >>"
4. **Speichern und schließen**

Jetzt hast du einen Button im Outlook-Ribbon!

### Zwei verfügbare Makros

1. **ExportSelectedEmailsToJSON**
   - Exportiert ausgewählte Emails
   - Flexibler, du wählst was exportiert wird

2. **ExportFolderToJSON**
   - Exportiert ALLE Emails aus dem aktuell geöffneten Ordner
   - Gut für Massenexporte

### Troubleshooting VBA

#### "Makros wurden deaktiviert"

Lösung:
1. Outlook → Datei → Optionen → Trust Center → Einstellungen für das Trust Center
2. Makrosicherheit → "Benachrichtigung für alle Makros" auswählen
3. Outlook neu starten

#### "Laufzeitfehler bei großen Exporten"

Bei sehr vielen Emails (>1000):
- Nutze das PowerShell-Script stattdessen (stabiler bei großen Datenmengen)
- Oder: Exportiere mehrere kleinere Batches

---

## 📊 Vergleich der Methoden

| Feature | PowerShell-Script | VBA-Makro |
|---------|------------------|-----------|
| **Installation** | Keine | VBA-Editor Setup |
| **Benutzerfreundlichkeit** | Extern ausführen | Direkt in Outlook |
| **Performance** | Sehr gut | Gut |
| **Große Datenmengen** | Exzellent (>10.000 Emails) | OK (< 1.000 Emails) |
| **Ordner-Export** | ✅ Ja | ✅ Ja |
| **Fortschrittsanzeige** | ✅ Detailliert | ✅ Statusbar |
| **Fehlerbehandlung** | ✅ Ausführlich | ✅ Basis |
| **Ribbon-Integration** | ❌ Nein | ✅ Ja (optional) |

**Empfehlung**: Nutze das **PowerShell-Script** für regelmäßige Exports oder große Datenmengen. Nutze das **VBA-Makro** wenn du es direkt in Outlook integriert haben möchtest.

---

## 🔄 Kompletter Workflow

### Von Outlook bis zur Excel-Datei

1. **Export aus Outlook** (eine der beiden Methoden):
   ```
   Outlook Emails → Export-Script → JSON-Datei
   ```

2. **Import in HTML-App**:
   - Öffne `bgav-hypercare-email-review.html` im Browser
   - Klicke "📤 Emails laden"
   - Wähle die JSON-Datei aus

3. **Review & Kategorisierung**:
   - Prüfe die automatische Kategorisierung
   - Passe bei Bedarf an (Kategorie, Cluster, Status)
   - Nutze Filter und Suche

4. **Export nach Excel**:
   - Klicke "💾 Nach Excel exportieren"
   - Die Excel-Datei wird heruntergeladen
   - Öffne sie in Excel oder importiere sie weiter

---

## 💡 Tipps & Best Practices

### Für PowerShell-Export

- **Große Exports**: Bei >5.000 Emails lieber mehrere kleinere Exports machen
- **Automatisierung**: Du kannst das Script in Task Scheduler einbinden für regelmäßige Exports
- **Ordnerstruktur**: Nutze aussagekräftige Ordnernamen in Outlook (z.B. "Hypercare_2025_11")

### Für VBA-Export

- **Backup**: Sichere dein VBA-Code (Datei → Exportieren im VBA-Editor)
- **Outlook-Updates**: Nach Outlook-Updates ggf. Makrosicherheit erneut prüfen
- **Testen**: Teste erst mit wenigen Emails (z.B. 10) bevor du große Mengen exportierst

### Für die HTML-App

- **Browser**: Chrome, Firefox oder Edge empfohlen
- **Große JSON-Dateien**: Browser kann bei sehr großen Dateien (>50 MB) langsam werden
- **Speichern**: Die HTML-App speichert nichts automatisch - exportiere regelmäßig nach Excel

---

## 🆘 Häufige Fehler

### "Keine Emails ausgewählt"
➡️ Lösung: Wähle in Outlook mindestens eine Email aus

### "COM-Objekt kann nicht erstellt werden"
➡️ Lösung: Outlook muss installiert sein und wurde mindestens einmal gestartet

### "JSON-Format ungültig"
➡️ Lösung: Prüfe ob die Export-Datei vollständig ist (sollte mit `]` enden)

### "Export dauert sehr lange"
➡️ Lösung: Exportiere kleinere Batches oder nutze das PowerShell-Script

### "Umlaute werden falsch dargestellt"
➡️ Lösung: Datei wird als UTF-8 gespeichert, Browser sollte das korrekt lesen

---

## 📞 Weitere Hilfe

Falls Probleme auftreten:

1. **Prüfe die Fehlermeldung** - oft steht dort schon die Lösung
2. **Test mit wenigen Emails** - erst 2-3 Emails zum Testen
3. **Outlook neu starten** - manchmal hilft das
4. **PowerShell als Admin** - falls Berechtigungsprobleme auftreten

---

## 🔐 Sicherheit & Datenschutz

- ✅ Alle Daten bleiben lokal auf deinem Computer
- ✅ Keine Cloud-Verbindungen
- ✅ Keine Telemetrie oder Tracking
- ✅ Open Source - du kannst den Code jederzeit überprüfen

Die JSON-Dateien enthalten die gleichen Daten wie deine Outlook-Emails - behandle sie entsprechend vertraulich!
