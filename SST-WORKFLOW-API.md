# SST-Workflow-API - Schnittstellen-Referenz

Referenz-Code fuer die SST-Workflow-API (`de.barmenia.sst.workflow2.api.userintent`), bereitgestellt vom SST-Team.

## Verwendete API-Klassen

| Klasse | Paket | Beschreibung |
|--------|-------|-------------|
| `UserIntentService` | `api.userintent` | Haupt-Service fuer alle Workflow-Operationen |
| `UserIntentServiceFactory` | `api.userintent` | Factory zum Erzeugen einer `UserIntentService`-Instanz |
| `MDCWrapper` | `api.userintent` | Wrapper fuer Logging-Kontext (correlationId) |
| `AkteDescriptor` | `descriptors` | Beschreibt die Ziel-Akte (Geschnotyp + Geschno) |
| `ArbeitsdokumentDescriptor` | `descriptors` | Beschreibt ein Arbeitsdokument (PDF mit Referenz auf Original) |
| `DokumentBytesDescriptor` | `descriptors` | Beschreibt ein Dokument ohne Original-Referenz (z.B. Dashboard) |
| `VorgangDescriptor` | `descriptors` | Beschreibt einen Vorgang (Typ, Ersteller, Fremdschluessel) |

## Verwendete Entities

| Entity | Beschreibung |
|--------|-------------|
| `Originaldokument` | EML-Original das archiviert wird |
| `DokumentFormat` | Enum: `EML`, `PDF`, etc. |
| `Quelle` | Enum: `API`, etc. |
| `Fremdschluessel` | Eindeutiger externer Schluessel (System + ID) |
| `FremdschluesselSystem` | Enum: `WORKFLOW`, etc. |
| `Postverteilart` | Enum: `OHNE`, etc. |
| `Geschnotyp` | Enum: `BD_VERMITTLER_VERTRAG`, etc. |
| `Vorgang` | Repraesentiert einen SST-Vorgang |

## Ablauf: Mail ablegen (`legeMailAb`)

```
1. archiviereOriginaldokument()    -> docnumber (EML ins Archiv)
2. erzeugeDokument()               -> Arbeitsdokument (PDF) + Vorgang anlegen
3. gibVorgangZuFremdschluessel()   -> Vorgang-Objekt abrufen
4. schliesseVorgangAb()            -> Vorgang abschliessen
5. gibDokumenteZuVorgang()         -> Dokument-Liste abrufen
6. klammerDokument()               -> Dokument mit Klammerbegriff versehen
```

## Ablauf: Dashboard ablegen (`legeDashboardAb`)

```
1. erzeugeDokument()               -> Dokument (PDF) + Vorgang anlegen (kein EML-Original)
2. gibVorgangZuFremdschluessel()   -> Vorgang-Objekt abrufen
3. schliesseVorgangAb()            -> Vorgang abschliessen
```

## Referenz-Code (vom SST-Team)

```java
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

public class WorkflowAdapter {
    private static final long VORGANGS_TYP = -1;
    private static final long DOKUMENT_TYP = -1;
    private static final String HINWEIS_TEXT = "";

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
                                        //Hier wird idealerweise noch ein eigener Eintrag bei uns in der DB gemacht
                                        //und das "richtige" system eingetragen
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
                //Hier wird idealerweise noch ein eigener Eintrag bei uns in der DB gemacht
                //und das "richtige" system eingetragen
                .withExterneSystemId("WORKFLOW")
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
                        //Hier z.B. Anwendungskuerzel
                        .withFremdschluesselSystem("")
                        //Erstellender User (1 + 6 stellige PersNr)
                        .withErsteller("")
                        .build(),
                AkteDescriptor.Builder.create()
                        .withGeschnotyp(Geschnotyp.BD_VERMITTLER_VERTRAG)
                        .withGeschno(bdVermittlerNummer)
                        .build()
        );

        final Vorgang vorgang = userIntentService.gibVorgangZuFremdschluessel(mdcWrapper,
                Fremdschluessel.Builder.create()
                        .withFremdschluesselSystem(FremdschluesselSystem.valueOf(""))
                        .withFremdschluesselId(correlationId)
                        .build());

        userIntentService.schliesseVorgangAb(mdcWrapper, VorgangDescriptor.Builder.create()
                .withIdVorgang(Long.parseLong(vorgang.getId()))
                .build());

        final List<ArbeitsdokumentDescriptor> dokumente =
                userIntentService.gibDokumenteZuVorgang(mdcWrapper, Long.parseLong(vorgang.getId()));

        userIntentService.klammerDokument(mdcWrapper, dokumente.getFirst().getIdDokument(), klammerBegriff);
    }

    public void legeDashboardAb(
            final byte[] emailPdf,
            final String bdVermittlerNummer
    ) throws UserIntentServiceException {
        //Wird als eindeutiger Identifier verwendet!
        //Entweder eindeutig oder wenn mail und dashboard abgelegt werden mit suffix arbeiten
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
                //Hier wird idealerweise noch ein eigener Eintrag bei uns in der DB gemacht
                //und das "richtige" system eingetragen
                .withExterneSystemId("WORKFLOW")
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
                        //Hier z.B. Anwendungskuerzel
                        .withFremdschluesselSystem("")
                        //Erstellender User (1 + 6 stellige PersNr)
                        .withErsteller("")
                        .build(),
                AkteDescriptor.Builder.create()
                        .withGeschnotyp(Geschnotyp.BD_VERMITTLER_VERTRAG)
                        .withGeschno(bdVermittlerNummer)
                        .build()
        );

        final Vorgang vorgang = userIntentService.gibVorgangZuFremdschluessel(mdcWrapper,
                Fremdschluessel.Builder.create()
                        .withFremdschluesselSystem(FremdschluesselSystem.valueOf(""))
                        .withFremdschluesselId(correlationId)
                        .build());

        userIntentService.schliesseVorgangAb(mdcWrapper, VorgangDescriptor.Builder.create()
                .withIdVorgang(Long.parseLong(vorgang.getId()))
                .build());
    }
}
```

## Aktuelle Konfiguration (BGAV-Projekt)

| Parameter | Wert | Status |
|-----------|------|--------|
| Klammer | `ADM Vertrag` | gesetzt |
| Dokumentenart | VSW (Schriftwechsel) | TODO: numerische ID fehlt |
| Vorgangsart | Sonstige | TODO: numerische ID fehlt |
| Hinweis | `Titel ab 01.01.26` | gesetzt |
| Fremdschluesselsystem | `EV_Technischer_Vertragsnachtrag` | gesetzt |
| Technischer User | `TBD` | TODO: echte PersNr eintragen |
| Geschnotyp | `BD_VERMITTLER_VERTRAG` | gesetzt |
