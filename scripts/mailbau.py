#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gemeinsame Bausteine fuer die .eml-Generatoren in diesem Ordner.

Wichtig: Serialisierung ueber policy.SMTP (CRLF). Mit blossen LF-Zeilenenden
verschluckt Outlook an den Quoted-Printable-Umbruechen Zeichen.
"""

import mimetypes
import os
from email import policy
from email.message import EmailMessage

CSS = """
  body { font-family: Aptos, 'Segoe UI', Calibri, sans-serif; font-size: 11pt;
         color: #111827; line-height: 1.45; }
  h3 { font-size: 11.5pt; margin: 18px 0 6px 0; color: #1E40AF; }
  p { margin: 8px 0; }
  table { border-collapse: collapse; margin: 8px 0 14px 0; }
  th, td { border: 1px solid #D1D5DB; padding: 4px 9px; font-size: 10.5pt;
           text-align: left; vertical-align: top; }
  th { background: #1E40AF; color: #FFFFFF; font-weight: 600; }
  td.n { text-align: right; white-space: nowrap; }
  tr.neu td { background: #DCFCE7; }
  tr.offen td { background: #FEF3C7; }
  code { font-family: Consolas, 'Courier New', monospace; font-size: 10pt; }
  ul { margin: 6px 0 12px 0; padding-left: 22px; }
  li { margin: 4px 0; }
  .hint { color: #4B5563; }
  .platzhalter { border: 2px dashed #D97706; background: #FFFBEB; padding: 10px 14px;
                 margin: 10px 0 14px 0; }
"""


def tab(kopf, zeilen, num_spalten=(), klassen=None):
    """HTML-Tabelle. zeilen = Liste von Zellwert-Listen.
    klassen = optionale Liste mit Zeilenklassen ('neu', 'offen' oder None)."""
    h = "".join("<th>%s</th>" % k for k in kopf)
    body = ""
    for i, zl in enumerate(zeilen):
        kl = (klassen[i] if klassen and i < len(klassen) else None)
        tds = ""
        for j, wert in enumerate(zl):
            tds += "<td%s>%s</td>" % (' class="n"' if j in num_spalten else "", wert)
        body += "<tr%s>%s</tr>" % (' class="%s"' % kl if kl else "", tds)
    return "<table><tr>%s</tr>%s</table>" % (h, body)


def lies_text(dateiname):
    """Klartextteil laden; eine fuehrende 'Betreff:'-Zeile wird entfernt."""
    pfad = os.path.join(os.path.dirname(os.path.abspath(__file__)), dateiname)
    text = open(pfad, encoding="utf-8").read()
    if text.startswith("Betreff:"):
        text = text.split("\n", 1)[1].lstrip("\n")
    return text


def baue_eml(ziel, betreff, text, html, anhaenge=()):
    """Schreibt eine .eml mit Klartext- und HTML-Teil sowie Anhaengen.

    Der Betreff sollte hoechstens 64 Zeichen haben - laengere werden beim
    Schreiben umgebrochen, was in manchen Clients doppelte Leerzeichen erzeugt.
    """
    msg = EmailMessage()
    msg["Subject"] = betreff
    msg["To"] = ""
    msg["X-Unsent"] = "1"          # Outlook oeffnet die Datei als Entwurf
    msg.set_content(text)
    msg.add_alternative(
        '<html><head><meta charset="utf-8"><style>%s</style></head><body>%s</body></html>'
        % (CSS, html), subtype="html")

    for pfad in anhaenge:
        if not os.path.exists(pfad):
            print("Hinweis: %s nicht gefunden - wird nicht angehaengt." % pfad)
            continue
        typ, _ = mimetypes.guess_type(pfad)
        haupt, unter = (typ or "application/octet-stream").split("/", 1)
        with open(pfad, "rb") as f:
            msg.add_attachment(f.read(), maintype=haupt, subtype=unter,
                               filename=os.path.basename(pfad))

    with open(ziel, "wb") as f:
        f.write(msg.as_bytes(policy=policy.SMTP))
    print("OK - %s geschrieben (%.1f MB)" % (ziel, os.path.getsize(ziel) / 1048576.0))
