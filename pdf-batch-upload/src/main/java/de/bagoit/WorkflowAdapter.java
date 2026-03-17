package de.bagoit;

import de.barmenia.sst.workflow2.api.userintent.MDCWrapper;
import de.barmenia.sst.workflow2.api.userintent.UserIntentService;
import de.barmenia.sst.workflow2.api.userintent.UserIntentServiceException;
import de.barmenia.sst.workflow2.api.userintent.UserIntentServiceFactory;
import de.barmenia.sst.workflow2.api.userintent.descriptors.AkteDescriptor;
import de.barmenia.sst.workflow2.api.userintent.descriptors.ArbeitsdokumentDescriptor;
import de.barmenia.sst.workflow2.api.userintent.descriptors.DokumentBytesDescriptor;
import de.barmenia.sst.workflow2.api.userintent.descriptors.VorgangDescriptor;
import de.barmenia.sst.workflow2.api.userintent.entities.*;
import org.slf4j.MDC;

import java.util.List;
import java.util.Optional;
import java.util.UUID;

/**
 * Referenz-Code vom SST-Team.
 * Adapter fuer die SST-Workflow-API - wird von BgavBatchUpload aufgerufen.
 *
 * Konfigurationswerte (BGAV-Projekt):
 *   Klammer:              ADM Vertrag (wird als Parameter uebergeben)
 *   Dokumentenart:         VSW = Schriftwechsel (DOKUMENT_TYP - ID noch offen)
 *   Vorgangsart:           Sonstige (VORGANGS_TYP - ID noch offen)
 *   Hinweis:               Titel ab 01.01.26
 *   Fremdschluesselsystem: EV_Technischer_Vertragsnachtrag
 *   Technischer User:      TBD (ERSTELLER - echte PersNr noch offen)
 */
public class WorkflowAdapter {
    private static final long VORGANGS_TYP = -1;   // TODO: ID fuer "Sonstige"
    private static final long DOKUMENT_TYP = -1;   // TODO: ID fuer "VSW" (Schriftwechsel)
    private static final String HINWEIS_TEXT = "Titel ab 01.01.26";
    private static final String FREMDSCHLUESSEL_SYSTEM = "EV_Technischer_Vertragsnachtrag";
    private static final String ERSTELLER = "TBD";  // TODO: echte PersNr

    public void legeMailAb(
            final byte[] emailEml,
            final byte[] emailPdf,
            final String bdVermittlerNummer,
            final String klammerBegriff
    ) throws UserIntentServiceException {
        final String correlationId = Optional.ofNullable(MDC.get("correlationId")).orElse(UUID.randomUUID().toString());

        final MDCWrapper mdcWrapper = MDCWrapper.Builder.create()
                .withCorrelationId(correlationId)
                .build();

        final UserIntentService userIntentService = UserIntentServiceFactory.newInstance();

        final String docnumberOriginal = userIntentService.archiviereOriginaldokument(
                mdcWrapper,
                Originaldokument.Builder.create()
                        .withContent(emailEml)
                        .withDokumentFormat(DokumentFormat.EML)
                        .withQuelle(Quelle.API)
                        .withFremdschluessel(
                                Fremdschluessel.Builder.create()
                                        .withFremdschluesselId(correlationId + "_original")
                                        //Hier wird idealerweise noch ein eigener Eintrag bei uns in der DB gemacht und das "richtige" system eingetragen
                                        .withFremdschluesselSystem(FremdschluesselSystem.WORKFLOW)
                                        .build()
                        )
                        .build()
        );

        final ArbeitsdokumentDescriptor arbeitsdokumentDescriptor = ArbeitsdokumentDescriptor.Builder.create()
                .withBytes(emailPdf)
                .withOriginalDocNumber(Long.parseLong(docnumberOriginal))
                .withIdDoktyp(DOKUMENT_TYP)
                .withDokumentFormat(DokumentFormat.PDF)
                .withQuelle(Quelle.API)
                //Hier wird idealerweise noch ein eigener Eintrag bei uns in der DB gemacht und das "richtige" system eingetragen
                .withExterneSystemId(FREMDSCHLUESSEL_SYSTEM)
                .withExterneId(correlationId + "_arbeit")
                .build();

        //Seitenzahl ermitteln und eintragen
        arbeitsdokumentDescriptor.setSeitenzahl(-1);
        arbeitsdokumentDescriptor.setHinweis(HINWEIS_TEXT);

        userIntentService.erzeugeDokument(
                mdcWrapper,
                arbeitsdokumentDescriptor,
                //Hier wird auch ein Vorgang erstellt, sollte dieser nicht existieren
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

        final Vorgang vorgang = userIntentService.gibVorgangZuFremdschluessel(mdcWrapper, Fremdschluessel.Builder.create()
                .withFremdschluesselSystem(FremdschluesselSystem.valueOf(FREMDSCHLUESSEL_SYSTEM))
                .withFremdschluesselId(correlationId)
                .build());

        userIntentService.schliesseVorgangAb(mdcWrapper, VorgangDescriptor.Builder.create()
                .withIdVorgang(Long.parseLong(vorgang.getId()))
                .build());

        final List<ArbeitsdokumentDescriptor> dokumente = userIntentService.gibDokumenteZuVorgang(mdcWrapper, Long.parseLong(vorgang.getId()));

        userIntentService.klammerDokument(mdcWrapper, dokumente.getFirst().getIdDokument(), klammerBegriff);
    }

    public void legeDashboardAb(
            final byte[] emailPdf,
            final String bdVermittlerNummer
    ) throws UserIntentServiceException {
        //Wird als eindeutiger Identifier verwendet! Entweder eindeutig oder wenn mail und dashboard abgelegt werden mit suffix arbeiten
        final String correlationId = Optional.ofNullable(MDC.get("correlationId")).orElse(UUID.randomUUID().toString());

        final MDCWrapper mdcWrapper = MDCWrapper.Builder.create()
                .withCorrelationId(correlationId)
                .build();

        final UserIntentService userIntentService = UserIntentServiceFactory.newInstance();

        final DokumentBytesDescriptor arbeitsdokumentDescriptor = DokumentBytesDescriptor.Builder.create()
                .withBytes(emailPdf)
                .withIdDoktyp(DOKUMENT_TYP)
                .withDokumentFormat(DokumentFormat.PDF)
                .withQuelle(Quelle.API)
                //Hier wird idealerweise noch ein eigener Eintrag bei uns in der DB gemacht und das "richtige" system eingetragen
                .withExterneSystemId(FREMDSCHLUESSEL_SYSTEM)
                .withExterneId(correlationId + "_arbeit")
                .withSeitenzahl(-1)
                .withHinweis(HINWEIS_TEXT)
                .build();

        userIntentService.erzeugeDokument(
                mdcWrapper,
                arbeitsdokumentDescriptor,
                //Hier wird auch ein Vorgang erstellt, sollte dieser nicht existieren
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

        final Vorgang vorgang = userIntentService.gibVorgangZuFremdschluessel(mdcWrapper, Fremdschluessel.Builder.create()
                .withFremdschluesselSystem(FremdschluesselSystem.valueOf(FREMDSCHLUESSEL_SYSTEM))
                .withFremdschluesselId(correlationId)
                .build());

        userIntentService.schliesseVorgangAb(mdcWrapper, VorgangDescriptor.Builder.create()
                .withIdVorgang(Long.parseLong(vorgang.getId()))
                .build());
    }
}
