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
 * PDF Batch Upload Tool
 *
 * <p>Liest die vom HTA-Tool (pdf-massenupload.hta) erzeugte CSV
 * und laedt die Dashboard-PDFs ueber die SST-Workflow-API hoch.</p>
 *
 * <p>Verwendet {@link WorkflowAdapter#legeDashboardAb(byte[], String)}
 * (nur PDF, kein EML-Originaldokument).</p>
 *
 * <p><b>Verwendung:</b></p>
 * <pre>
 *   java -jar pdf-batch-upload.jar &lt;pfad-zum-ordner&gt;
 *   java -jar pdf-batch-upload.jar &lt;pfad-zum-ordner&gt; --dry-run
 * </pre>
 *
 * <p><b>Erwartet diese Ordnerstruktur (erzeugt vom HTA-Tool):</b></p>
 * <pre>
 *   ordner/
 *   ├── Dashboard_00630595_29102025.pdf   (PDF-Dateien)
 *   ├── Dashboard_00630593_29102025.pdf
 *   ├── ...
 *   └── pdf_upload_metadaten.csv          (Metadaten vom HTA-Tool)
 * </pre>
 *
 * <p><b>CSV-Spalten (Semikolon-getrennt):</b></p>
 * <pre>
 *   Nr;Dateiname_PDF;BD_Nummer;Klammerbegriff
 * </pre>
 */
public class PdfBatchUpload {

    private static final Logger log = LoggerFactory.getLogger(PdfBatchUpload.class);

    private static final String CSV_SEPARATOR = ";";
    private static final String CSV_FILENAME = "pdf_upload_metadaten.csv";

    public static void main(String[] args) {
        if (args.length < 1) {
            System.err.println("Verwendung: java -jar pdf-batch-upload.jar <pfad-zum-ordner>");
            System.err.println();
            System.err.println("Der Ordner muss die Dashboard-PDFs und die vom HTA-Tool");
            System.err.println("erzeugte " + CSV_FILENAME + " enthalten.");
            System.err.println();
            System.err.println("Beispiel:");
            System.err.println("  java -jar pdf-batch-upload.jar C:\\Dashboard_PDFs");
            System.err.println("  java -jar pdf-batch-upload.jar C:\\Dashboard_PDFs --dry-run");
            System.exit(1);
        }

        final Path baseDir = Paths.get(args[0]);
        final boolean dryRun = args.length > 1 && "--dry-run".equals(args[1]);

        if (dryRun) {
            log.info("=== DRY-RUN Modus - keine API-Aufrufe ===");
        }

        try {
            new PdfBatchUpload().run(baseDir, dryRun);
        } catch (Exception e) {
            log.error("Fataler Fehler: {}", e.getMessage(), e);
            System.exit(1);
        }
    }

    /**
     * Hauptlogik: CSV laden, PDFs pruefen, Upload durchfuehren.
     *
     * @param baseDir  Pfad zum Ordner (mit PDFs und pdf_upload_metadaten.csv)
     * @param dryRun   true = nur pruefen und loggen, kein API-Aufruf
     * @throws IOException wenn CSV nicht gefunden oder Dateien fehlen
     */
    public void run(Path baseDir, boolean dryRun) throws IOException {
        log.info("============================================");
        log.info("PDF Batch Upload (Dashboard)");
        log.info("Basis-Ordner: {}", baseDir);
        log.info("============================================");

        // ── Schritt 1: CSV laden ──
        final Path csvPath = baseDir.resolve(CSV_FILENAME);
        if (!Files.exists(csvPath)) {
            throw new IOException("CSV nicht gefunden: " + csvPath +
                    "\nBitte zuerst das HTA-Tool (pdf-massenupload.hta) ausfuehren.");
        }

        final List<PdfRecord> records = readCsv(csvPath);
        log.info("{} PDF-Eintraege geladen", records.size());

        // ── Schritt 2: Pruefen ob alle PDF-Dateien vorhanden sind ──
        int missingFiles = 0;
        for (PdfRecord rec : records) {
            Path pdfFile = baseDir.resolve(rec.datenamePdf);
            if (!Files.exists(pdfFile)) {
                log.error("PDF nicht gefunden: {}", pdfFile);
                missingFiles++;
            }
        }

        if (missingFiles > 0) {
            throw new IOException(missingFiles + " PDF-Dateien fehlen. Abbruch.");
        }

        log.info("Alle {} PDF-Dateien vorhanden", records.size());

        // ── Schritt 3: Upload pro PDF ──
        final WorkflowAdapter adapter = new WorkflowAdapter();
        int success = 0;
        int skipped = 0;
        int failed = 0;

        for (int i = 0; i < records.size(); i++) {
            final PdfRecord rec = records.get(i);
            log.info("");
            log.info("--- PDF {}/{}: BD {} | {} ---",
                    i + 1, records.size(), rec.bdNummer, rec.datenamePdf);

            // Ohne BD-Nummer kann kein Vorgang in der Akte angelegt werden
            if (rec.bdNummer == null || rec.bdNummer.isEmpty()) {
                log.warn("Ueberspringe: Keine BD-Nummer vorhanden");
                skipped++;
                continue;
            }

            try {
                final byte[] pdfBytes = Files.readAllBytes(baseDir.resolve(rec.datenamePdf));

                log.info("  PDF:    {} ({} bytes)", rec.datenamePdf, pdfBytes.length);
                log.info("  BD:     {}", rec.bdNummer);
                log.info("  Klammer: {}", rec.klammerbegriff);

                if (dryRun) {
                    log.info("  [DRY-RUN] Wuerde legeDashboardAb() aufrufen...");
                } else {
                    // Ruft die SST-Workflow-API auf (nur PDF, kein EML):
                    // 1. PDF + Vorgang anlegen  2. Vorgang schliessen
                    adapter.legeDashboardAb(pdfBytes, rec.bdNummer);
                }

                success++;
                log.info("  -> OK");

            } catch (Exception e) {
                failed++;
                log.error("  -> FEHLER: {}", e.getMessage(), e);
            }
        }

        // ── Schritt 4: Zusammenfassung ──
        log.info("");
        log.info("============================================");
        log.info("ERGEBNIS");
        log.info("============================================");
        log.info("Gesamt:       {}", records.size());
        log.info("Erfolgreich:  {}", success);
        log.info("Uebersprungen: {}", skipped);
        log.info("Fehler:       {}", failed);

        if (failed > 0) {
            log.warn("Es gab {} Fehler. Bitte Log pruefen.", failed);
        } else if (skipped > 0) {
            log.info("Alle verarbeitbaren PDFs erfolgreich hochgeladen ({} ohne BD-Nummer uebersprungen).", skipped);
        } else {
            log.info("Alle PDFs erfolgreich hochgeladen!");
        }
    }

    // ═══════════════════════════════════════════════════════════
    // CSV-Verarbeitung
    // ═══════════════════════════════════════════════════════════

    /**
     * Liest die pdf_upload_metadaten.csv und gibt eine Liste von PdfRecords zurueck.
     *
     * <p>Unterstuetzt UTF-8 und UTF-16LE (mit BOM), da das HTA-Tool
     * je nach Windows-Konfiguration unterschiedlich schreibt.</p>
     *
     * <p><b>CSV-Spalten (Semikolon-getrennt):</b></p>
     * <pre>
     * Index  Spalte            Beispiel
     * ─────  ──────            ────────
     * 0      Nr                1
     * 1      Dateiname_PDF     Dashboard_00630595_29102025.pdf
     * 2      BD_Nummer         00630595
     * 3      Klammerbegriff    ADM Vertrag
     * </pre>
     */
    private List<PdfRecord> readCsv(Path csvPath) throws IOException {
        // Encoding-Erkennung: UTF-16LE beginnt mit BOM-Bytes FF FE
        final byte[] rawBytes = Files.readAllBytes(csvPath);
        Charset charset = StandardCharsets.UTF_8;
        if (rawBytes.length >= 2 && rawBytes[0] == (byte) 0xFF && rawBytes[1] == (byte) 0xFE) {
            charset = StandardCharsets.UTF_16LE;
        }

        final List<PdfRecord> records = new ArrayList<>();

        try (BufferedReader reader = Files.newBufferedReader(csvPath, charset)) {
            String headerLine = reader.readLine();
            if (headerLine == null) {
                throw new IOException("CSV ist leer");
            }

            // BOM-Zeichen entfernen
            headerLine = headerLine.replace("\uFEFF", "").trim();

            // Plausibilitaetspruefung
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
                // Mindestens 4 Spalten erwartet (Nr, Dateiname_PDF, BD_Nummer, Klammerbegriff)
                if (parts.length < 4) {
                    log.warn("Zeile {} hat nur {} Spalten (erwartet min. 4), ueberspringe", lineNr, parts.length);
                    continue;
                }

                PdfRecord rec = new PdfRecord();
                rec.nr = parseInt(parts[0]);
                rec.datenamePdf = parts[1].trim();
                rec.bdNummer = parts[2].trim();
                rec.klammerbegriff = parts[3].trim();

                records.add(rec);
            }
        }

        return records;
    }

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
     * Ein Datensatz aus der pdf_upload_metadaten.csv.
     * Entspricht einer Zeile = ein Dashboard-PDF mit BD-Nummer.
     */
    static class PdfRecord {
        /** Laufende Nummer aus der CSV */
        int nr;
        /** Dateiname der PDF-Datei (z.B. "Dashboard_00630595_29102025.pdf") */
        String datenamePdf;
        /** BD-Vermittlernummer (z.B. "00630595") - Pflichtfeld fuer den Upload */
        String bdNummer;
        /** Klammerbegriff fuer die SST-API (z.B. "ADM Vertrag") */
        String klammerbegriff;

        @Override
        public String toString() {
            return String.format("PdfRecord{nr=%d, bd=%s, pdf=%s}",
                    nr, bdNummer, datenamePdf);
        }
    }
}
