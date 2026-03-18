package de.bagoit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * BGAV Batch Upload Tool
 *
 * <p>Liest die vom HTA-Tool (bgav-testmail-extraktion.hta) erzeugten Testfaelle
 * und laedt sie ueber die SST-Workflow-API (WorkflowAdapter) hoch.</p>
 *
 * <p><b>Verwendung:</b></p>
 * <pre>
 *   java -jar bgav-batch-upload.jar &lt;pfad-zum-output-ordner&gt;
 *   java -jar bgav-batch-upload.jar &lt;pfad-zum-output-ordner&gt; --dry-run
 * </pre>
 *
 * <p><b>Erwartet diese Ordnerstruktur (erzeugt vom HTA-Tool):</b></p>
 * <pre>
 *   output/
 *   ├── eml/                        *.eml Dateien (Original-Mails)
 *   ├── pdf/                        *.pdf Dateien (Arbeitsdokumente)
 *   └── testfaelle_metadaten.csv    Metadaten mit BD-Nummern, Klammerbegriff etc.
 * </pre>
 *
 * <p><b>Ablauf:</b></p>
 * <ol>
 *   <li>CSV laden und parsen (unterstuetzt UTF-8 und UTF-16LE vom HTA)</li>
 *   <li>Pruefen ob alle referenzierten EML/PDF-Dateien vorhanden sind</li>
 *   <li>Pro Testfall: WorkflowAdapter.legeMailAb() aufrufen</li>
 *   <li>Zusammenfassung ausgeben (Erfolg/Fehler)</li>
 * </ol>
 */
public class BgavBatchUpload {

    private static final Logger log = LoggerFactory.getLogger(BgavBatchUpload.class);

    /** CSV-Trennzeichen - Semikolon, da vom HTA-Tool so erzeugt */
    private static final String CSV_SEPARATOR = ";";

    /**
     * Einstiegspunkt. Erwartet als erstes Argument den Pfad zum Output-Ordner.
     * Optionales zweites Argument: --dry-run (nur pruefen, kein API-Upload).
     */
    public static void main(String[] args) {
        if (args.length < 1) {
            System.err.println("Verwendung: java -jar bgav-batch-upload.jar <pfad-zum-output-ordner>");
            System.err.println();
            System.err.println("Beispiel:");
            System.err.println("  java -jar bgav-batch-upload.jar C:\\BGAV_Testmails");
            System.err.println("  java -jar bgav-batch-upload.jar C:\\BGAV_Testmails --dry-run");
            System.exit(1);
        }

        final Path baseDir = Paths.get(args[0]);
        final boolean dryRun = args.length > 1 && "--dry-run".equals(args[1]);

        if (dryRun) {
            log.info("=== DRY-RUN Modus - keine API-Aufrufe ===");
        }

        try {
            new BgavBatchUpload().run(baseDir, dryRun);
        } catch (Exception e) {
            log.error("Fataler Fehler: {}", e.getMessage(), e);
            System.exit(1);
        }
    }

    /**
     * Hauptlogik: CSV laden, Dateien pruefen, Upload durchfuehren.
     *
     * @param baseDir  Pfad zum Output-Ordner (mit eml/, pdf/, testfaelle_metadaten.csv)
     * @param dryRun   true = nur pruefen und loggen, kein tatsaechlicher API-Aufruf
     * @throws IOException wenn CSV nicht gefunden oder Dateien fehlen
     */
    public void run(Path baseDir, boolean dryRun) throws IOException {
        log.info("============================================");
        log.info("BGAV Batch Upload");
        log.info("Basis-Ordner: {}", baseDir);
        log.info("============================================");

        // ── Schritt 1: CSV laden ──
        final Path csvPath = baseDir.resolve("testfaelle_metadaten.csv");
        if (!Files.exists(csvPath)) {
            throw new IOException("CSV nicht gefunden: " + csvPath);
        }

        final List<TestfallRecord> records = readCsv(csvPath);
        log.info("{} Testfaelle geladen", records.size());

        // ── Schritt 2: Pruefen ob alle EML/PDF-Dateien vorhanden sind ──
        // Bricht komplett ab wenn Dateien fehlen, damit keine Teiluploads entstehen
        final Path emlDir = baseDir.resolve("eml");
        final Path pdfDir = baseDir.resolve("pdf");

        int missingFiles = 0;
        for (TestfallRecord rec : records) {
            Path emlFile = emlDir.resolve(rec.dateinameEml);
            Path pdfFile = pdfDir.resolve(rec.datenamePdf);

            if (!Files.exists(emlFile)) {
                log.error("EML nicht gefunden: {}", emlFile);
                missingFiles++;
            }
            if (!Files.exists(pdfFile)) {
                log.error("PDF nicht gefunden: {}", pdfFile);
                missingFiles++;
            }
        }

        if (missingFiles > 0) {
            throw new IOException(missingFiles + " Dateien fehlen. Abbruch.");
        }

        // ── Schritt 3: Upload pro Testfall ──
        // Fehler bei einzelnen Testfaellen fuehren NICHT zum Gesamtabbruch -
        // es wird weitergemacht und am Ende eine Zusammenfassung ausgegeben
        final WorkflowAdapter adapter = new WorkflowAdapter();
        int success = 0;
        int failed = 0;

        for (int i = 0; i < records.size(); i++) {
            final TestfallRecord rec = records.get(i);
            log.info("");
            log.info("--- Testfall {}/{}: BD {} | {} ---",
                    i + 1, records.size(), rec.bdNummer, rec.empfaengerEmail);

            // Ohne BD-Nummer kann kein Vorgang in der Akte angelegt werden
            if (rec.bdNummer == null || rec.bdNummer.isEmpty()) {
                log.warn("Ueberspringe: Keine BD-Nummer vorhanden");
                failed++;
                continue;
            }

            try {
                final byte[] emlBytes = Files.readAllBytes(emlDir.resolve(rec.dateinameEml));
                final byte[] pdfBytes = Files.readAllBytes(pdfDir.resolve(rec.datenamePdf));

                log.info("  EML: {} ({} bytes)", rec.dateinameEml, emlBytes.length);
                log.info("  PDF: {} ({} bytes)", rec.datenamePdf, pdfBytes.length);
                log.info("  BD:  {}", rec.bdNummer);
                log.info("  Klammer: {}", rec.klammerbegriff);

                if (dryRun) {
                    log.info("  [DRY-RUN] Wuerde legeMailAb() aufrufen...");
                } else {
                    // Ruft die SST-Workflow-API auf:
                    // 1. EML archivieren  2. PDF + Vorgang anlegen  3. Vorgang schliessen  4. Klammern
                    adapter.legeMailAb(emlBytes, pdfBytes, rec.bdNummer, rec.klammerbegriff);
                }

                success++;
                log.info("  -> OK");

            } catch (Exception e) {
                // Einzelfehler loggen aber nicht abbrechen
                failed++;
                log.error("  -> FEHLER: {}", e.getMessage(), e);
            }
        }

        // ── Schritt 4: Zusammenfassung ──
        log.info("");
        log.info("============================================");
        log.info("ERGEBNIS");
        log.info("============================================");
        log.info("Gesamt:      {}", records.size());
        log.info("Erfolgreich: {}", success);
        log.info("Fehler:      {}", failed);

        if (failed > 0) {
            log.warn("Es gab {} Fehler. Bitte Log pruefen.", failed);
        } else {
            log.info("Alle Testfaelle erfolgreich hochgeladen!");
        }
    }

    // ═══════════════════════════════════════════════════════════
    // CSV-Verarbeitung
    // ═══════════════════════════════════════════════════════════

    /**
     * Liest die testfaelle_metadaten.csv und gibt eine Liste von TestfallRecords zurueck.
     *
     * <p>Unterstuetzt zwei Encodings, da das HTA-Tool je nach Windows-Konfiguration
     * unterschiedlich schreibt:</p>
     * <ul>
     *   <li>UTF-8 (Standard)</li>
     *   <li>UTF-16LE mit BOM (FF FE) - typisch fuer HTA FileSystemObject mit unicode=true</li>
     * </ul>
     *
     * <p><b>CSV-Spalten (Semikolon-getrennt):</b></p>
     * <pre>
     * Index  Spalte              Beispiel
     * ─────  ──────              ────────
     * 0      Nr                  1
     * 1      Dateiname_EML       00010091_BGAV_Titelbezeichnung_001.eml
     * 2      Dateiname_PDF       00010091_BGAV_Titelbezeichnung_001.pdf
     * 3      Empfaenger_Email    name@firma.de
     * 4      BD_Nummer           00010091
     * 5      Klammerbegriff      ADM Vertrag
     * 6      Original_Dateiname  original.eml (optional)
     * </pre>
     *
     * @param csvPath Pfad zur CSV-Datei
     * @return Liste der gelesenen Testfall-Datensaetze
     * @throws IOException bei Lesefehlern oder leerer CSV
     */
    private List<TestfallRecord> readCsv(Path csvPath) throws IOException {
        // Encoding-Erkennung: UTF-16LE beginnt mit BOM-Bytes FF FE
        final byte[] rawBytes = Files.readAllBytes(csvPath);
        Charset charset = StandardCharsets.UTF_8;
        if (rawBytes.length >= 2 && rawBytes[0] == (byte) 0xFF && rawBytes[1] == (byte) 0xFE) {
            charset = StandardCharsets.UTF_16LE;
        }

        final List<TestfallRecord> records = new ArrayList<>();

        try (BufferedReader reader = Files.newBufferedReader(csvPath, charset)) {
            String headerLine = reader.readLine();
            if (headerLine == null) {
                throw new IOException("CSV ist leer");
            }

            // BOM-Zeichen entfernen (kann bei UTF-8 mit BOM vorkommen)
            headerLine = headerLine.replace("\uFEFF", "").trim();

            // Plausibilitaetspruefung: Header muss mit "Nr;" beginnen
            if (!headerLine.startsWith("Nr" + CSV_SEPARATOR)) {
                log.warn("Unerwarteter CSV-Header: {}", headerLine);
            }

            String line;
            int lineNr = 1;
            while ((line = reader.readLine()) != null) {
                lineNr++;
                line = line.trim();
                if (line.isEmpty()) continue;

                final String[] parts = line.split(CSV_SEPARATOR, -1);
                // Mindestens 6 Spalten erwartet (Nr bis Klammerbegriff)
                if (parts.length < 6) {
                    log.warn("Zeile {} hat nur {} Spalten (erwartet min. 6), ueberspringe", lineNr, parts.length);
                    continue;
                }

                TestfallRecord rec = new TestfallRecord();
                rec.nr = parseInt(parts[0]);
                rec.dateinameEml = parts[1].trim();
                rec.datenamePdf = parts[2].trim();
                rec.empfaengerEmail = parts[3].trim();
                rec.bdNummer = parts[4].trim();
                rec.klammerbegriff = parts[5].trim();

                records.add(rec);
            }
        }

        return records;
    }

    /** Sichere parseInt-Variante - gibt 0 zurueck bei ungueltigem Format. */
    private int parseInt(String s) {
        try {
            return Integer.parseInt(s.trim());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    // ═══════════════════════════════════════════════════════════
    // Datenmodell
    // ═══════════════════════════════════════════════════════════

    /**
     * Ein Datensatz aus der testfaelle_metadaten.csv.
     * Entspricht einer Zeile = ein Empfaenger einer BGAV-Mail.
     */
    static class TestfallRecord {
        /** Laufende Nummer aus der CSV */
        int nr;
        /** Dateiname der EML-Datei (z.B. "00010091_BGAV_Titelbezeichnung_001.eml") */
        String dateinameEml;
        /** Dateiname der PDF-Datei (z.B. "00010091_BGAV_Titelbezeichnung_001.pdf") */
        String datenamePdf;
        /** E-Mail-Adresse des Empfaengers */
        String empfaengerEmail;
        /** BD-Vermittlernummer (z.B. "00010091") - Pflichtfeld fuer den Upload */
        String bdNummer;
        /** Klammerbegriff fuer die SST-API (z.B. "ADM Vertrag") */
        String klammerbegriff;

        @Override
        public String toString() {
            return String.format("TestfallRecord{nr=%d, bd=%s, email=%s, eml=%s}",
                    nr, bdNummer, empfaengerEmail, dateinameEml);
        }
    }
}
