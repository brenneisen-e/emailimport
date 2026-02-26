package de.bagoit;

import de.barmenia.sst.workflow2.api.userintent.MDCWrapper;
import de.barmenia.sst.workflow2.api.userintent.UserIntentService;
import de.barmenia.sst.workflow2.api.userintent.UserIntentServiceException;
import de.barmenia.sst.workflow2.api.userintent.UserIntentServiceFactory;
import de.barmenia.sst.workflow2.api.userintent.descriptors.AkteDescriptor;
import de.barmenia.sst.workflow2.api.userintent.descriptors.ArbeitsdokumentDescriptor;
import de.barmenia.sst.workflow2.api.userintent.descriptors.VorgangDescriptor;
import de.barmenia.sst.workflow2.api.userintent.entities.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.slf4j.MDC;

import java.util.List;
import java.util.Optional;
import java.util.UUID;

/**
 * Adapter fuer die SST-Workflow-API.
 * Legt eine Mail (EML + PDF) in der Akte eines BD-Vermittlers ab.
 *
 * <p><b>TODO:</b> Vor Produktiveinsatz muessen die Konstanten
 * VORGANGS_TYP, DOKUMENT_TYP, HINWEIS_TEXT, ERSTELLER und
 * FREMDSCHLUESSEL_SYSTEM mit den richtigen Werten befuellt werden.</p>
 */
public class WorkflowAdapter {

    private static final Logger log = LoggerFactory.getLogger(WorkflowAdapter.class);

    // ───────────────────────────────────────────────────────────
    // TODO: Diese Konstanten mit den richtigen Werten befuellen!
    // ───────────────────────────────────────────────────────────
    /** Vorgangstyp-ID aus dem SST-System */
    private static final long VORGANGS_TYP = -1;  // TODO: richtigen Wert eintragen

    /** Dokumenttyp-ID fuer die BGAV-Titelbezeichnung-Mails */
    private static final long DOKUMENT_TYP = -1;  // TODO: richtigen Wert eintragen

    /** Hinweistext der am Dokument hinterlegt wird */
    private static final String HINWEIS_TEXT = "Serienbrief BGAV Kommunikation Titel Außendarstellung";

    /** Erstellender User (1 + 6-stellige PersNr) */
    private static final String ERSTELLER = "";  // TODO: z.B. "1234567"

    /** Fremdschluesselsystem-Kennung */
    private static final String FREMDSCHLUESSEL_SYSTEM = "";  // TODO: z.B. Anwendungskuerzel
    // ───────────────────────────────────────────────────────────

    /**
     * Legt eine E-Mail (EML-Original + PDF-Arbeitsdokument) in der
     * BD-Vermittler-Akte ab und klammert das Dokument.
     *
     * @param emailEml           EML-Datei als byte[]
     * @param emailPdf           PDF-Datei als byte[]
     * @param bdVermittlerNummer  BD-Nummer des Vermittlers (z.B. "00550179")
     * @param klammerBegriff     Klammerbegriff (z.B. "BGAV Titelbezeichnung")
     */
    public void legeMailAb(
            final byte[] emailEml,
            final byte[] emailPdf,
            final String bdVermittlerNummer,
            final String klammerBegriff
    ) throws UserIntentServiceException {

        final String correlationId = Optional.ofNullable(MDC.get("correlationId"))
                .orElse(UUID.randomUUID().toString());

        log.info("Starte Ablage fuer BD {} (correlationId: {})", bdVermittlerNummer, correlationId);

        final MDCWrapper mdcWrapper = MDCWrapper.Builder.create()
                .withCorrelationId(correlationId)
                .build();

        final UserIntentService userIntentService = UserIntentServiceFactory.newInstance();

        // 1. Originaldokument (EML) archivieren
        log.info("  1/5 Archiviere Originaldokument (EML, {} bytes)...", emailEml.length);
        final String docnumberOriginal = userIntentService.archiviereOriginaldokument(
                mdcWrapper,
                Originaldokument.Builder.create()
                        .withContent(emailEml)
                        .withDokumentFormat(DokumentFormat.EML)
                        .withQuelle(Quelle.API)
                        .withFremdschluessel(
                                Fremdschluessel.Builder.create()
                                        .withFremdschluesselId(correlationId + "_original")
                                        .withFremdschluesselSystem(FremdschluesselSystem.WORKFLOW)
                                        .build()
                        )
                        .build()
        );
        log.info("  -> Originaldokument archiviert: docnumber={}", docnumberOriginal);

        // 2. Arbeitsdokument (PDF) erzeugen + Vorgang anlegen
        log.info("  2/5 Erzeuge Arbeitsdokument (PDF, {} bytes)...", emailPdf.length);
        final ArbeitsdokumentDescriptor arbeitsdokumentDescriptor = ArbeitsdokumentDescriptor.Builder.create()
                .withBytes(emailPdf)
                .withOriginalDocNumber(Long.parseLong(docnumberOriginal))
                .withIdDoktyp(DOKUMENT_TYP)
                .withDokumentFormat(DokumentFormat.PDF)
                .withQuelle(Quelle.API)
                .withExterneSystemId(FREMDSCHLUESSEL_SYSTEM.isEmpty() ? "WORKFLOW" : FREMDSCHLUESSEL_SYSTEM)
                .withExterneId(correlationId + "_arbeit")
                .build();

        arbeitsdokumentDescriptor.setSeitenzahl(-1);
        arbeitsdokumentDescriptor.setHinweis(HINWEIS_TEXT);

        userIntentService.erzeugeDokument(
                mdcWrapper,
                arbeitsdokumentDescriptor,
                VorgangDescriptor.Builder.create()
                        .withIdVorgtyp(VORGANGS_TYP)
                        .withPostverteilart(Postverteilart.OHNE)
                        .withFremdschluesselId(correlationId)
                        .withFremdschluesselSystem(FREMDSCHLUESSEL_SYSTEM)
                        .withErsteller(ERSTELLER)
                        .build(),
                AkteDescriptor.Builder.create()
                        .withGeschnotyp(Geschnotyp.BD_VERMITTLER_VERTRAG)
                        .withGeschno(bdVermittlerNummer)
                        .build()
        );
        log.info("  -> Arbeitsdokument + Vorgang erzeugt");

        // 3. Vorgang abschliessen
        log.info("  3/5 Schliesse Vorgang ab...");
        final Vorgang vorgang = userIntentService.gibVorgangZuFremdschluessel(
                mdcWrapper,
                Fremdschluessel.Builder.create()
                        .withFremdschluesselSystem(FremdschluesselSystem.valueOf(
                                FREMDSCHLUESSEL_SYSTEM.isEmpty() ? "WORKFLOW" : FREMDSCHLUESSEL_SYSTEM))
                        .withFremdschluesselId(correlationId)
                        .build()
        );

        userIntentService.schliesseVorgangAb(
                mdcWrapper,
                VorgangDescriptor.Builder.create()
                        .withIdVorgang(Long.parseLong(vorgang.getId()))
                        .build()
        );
        log.info("  -> Vorgang {} abgeschlossen", vorgang.getId());

        // 4. Dokument klammern
        log.info("  4/5 Klammere Dokument mit '{}'...", klammerBegriff);
        final List<ArbeitsdokumentDescriptor> dokumente =
                userIntentService.gibDokumenteZuVorgang(mdcWrapper, Long.parseLong(vorgang.getId()));

        userIntentService.klammerDokument(mdcWrapper, dokumente.getFirst().getIdDokument(), klammerBegriff);
        log.info("  -> Dokument geklammert");

        log.info("  5/5 Ablage fuer BD {} erfolgreich abgeschlossen.", bdVermittlerNummer);
    }
}
