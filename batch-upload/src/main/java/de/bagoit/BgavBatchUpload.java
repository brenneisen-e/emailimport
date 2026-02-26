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
 * Liest die vom HTA-Tool erzeugten Testfaelle und laedt sie
 * ueber die SST-Workflow-API hoch.
 *
 * <pre>
 * Verwendung:
 *   java -jar bgav-batch-upload.jar &lt;pfad-zum-output-ordner&gt;
 *
 * Erwartet diese Struktur:
 *   output/
 *   ├── eml/  (*.eml Dateien)
 *   ├── pdf/  (*.pdf Dateien)
 *   └── testfaelle_metadaten.csv
 * </pre>
 */
public class BgavBatchUpload {

    private static final Logger log = LoggerFactory.getLogger(BgavBatchUpload.class);
    private static final String CSV_SEPARATOR = ";";

    public static void main(String[] args) {
        if (args.length < 1) {
            System.err.println("Verwendung: java -jar bgav-batch-upload.jar <pfad-zum-output-ordner>");
            System.err.println();
            System.err.println("Beispiel:");
            System.err.println("  java -jar bgav-batch-upload.jar C:\\BGAV_Testmails");
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

    public void run(Path baseDir, boolean dryRun) throws IOException {
        log.info("============================================");
        log.info("BGAV Batch Upload");
        log.info("Basis-Ordner: {}", baseDir);
        log.info("============================================");

        // 1. CSV laden
        final Path csvPath = baseDir.resolve("testfaelle_metadaten.csv");
        if (!Files.exists(csvPath)) {
            throw new IOException("CSV nicht gefunden: " + csvPath);
        }

        final List<TestfallRecord> records = readCsv(csvPath);
        log.info("{} Testfaelle geladen", records.size());

        // 2. Dateien pruefen
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

        // 3. Upload pro Testfall
        final WorkflowAdapter adapter = new WorkflowAdapter();
        int success = 0;
        int failed = 0;

        for (int i = 0; i < records.size(); i++) {
            final TestfallRecord rec = records.get(i);
            log.info("");
            log.info("--- Testfall {}/{}: BD {} | {} ---",
                    i + 1, records.size(), rec.bdNummer, rec.empfaengerEmail);

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
                if (rec.konversationsId != null && !rec.konversationsId.isEmpty()) {
                    log.info("  ConversationID: {}", rec.konversationsId);
                }

                if (dryRun) {
                    log.info("  [DRY-RUN] Wuerde legeMailAb() aufrufen...");
                } else {
                    adapter.legeMailAb(emlBytes, pdfBytes, rec.bdNummer, rec.klammerbegriff);
                }

                success++;
                log.info("  -> OK");

            } catch (Exception e) {
                failed++;
                log.error("  -> FEHLER: {}", e.getMessage(), e);
            }
        }

        // 4. Zusammenfassung
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

    /**
     * Liest die testfaelle_metadaten.csv.
     * Unterstuetzt sowohl UTF-8 als auch UTF-16LE (HTA-Output).
     *
     * CSV-Header:
     * Nr;Dateiname_EML;Dateiname_PDF;Empfaenger_Email;Vorname;Nachname;BD_Nummer;Klammerbegriff;Konversations_ID
     */
    private List<TestfallRecord> readCsv(Path csvPath) throws IOException {
        // Erkennung: UTF-16LE beginnt mit BOM FF FE
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

            // BOM entfernen falls vorhanden
            headerLine = headerLine.replace("\uFEFF", "").trim();

            // Header validieren
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
                if (parts.length < 8) {
                    log.warn("Zeile {} hat nur {} Spalten (erwartet min. 8), ueberspringe", lineNr, parts.length);
                    continue;
                }

                TestfallRecord rec = new TestfallRecord();
                rec.nr = parseInt(parts[0]);
                rec.dateinameEml = parts[1].trim();
                rec.datenamePdf = parts[2].trim();
                rec.empfaengerEmail = parts[3].trim();
                rec.vorname = parts[4].trim();
                rec.nachname = parts[5].trim();
                rec.bdNummer = parts[6].trim();
                rec.klammerbegriff = parts[7].trim();
                rec.konversationsId = parts.length > 8 ? parts[8].trim() : "";

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

    /** Ein Datensatz aus der testfaelle_metadaten.csv */
    static class TestfallRecord {
        int nr;
        String dateinameEml;
        String datenamePdf;
        String empfaengerEmail;
        String vorname;
        String nachname;
        String bdNummer;
        String klammerbegriff;
        String konversationsId;

        @Override
        public String toString() {
            return String.format("TestfallRecord{nr=%d, bd=%s, email=%s, eml=%s}",
                    nr, bdNummer, empfaengerEmail, dateinameEml);
        }
    }
}
