"""
Generate Auswertung-BUe-Stichwoche.pptx: 1 Slide mit Funnel + Phasen + Top-10-Pools.
Layout: 16:9 (13.33 x 7.5 inch).
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy

# Farbpalette - ERGO-naehe
ERGO_RED    = RGBColor(0xC8, 0x10, 0x2E)
ERGO_DARK   = RGBColor(0x1F, 0x29, 0x37)
PHASE1_BG   = RGBColor(0x16, 0x65, 0x34)   # gruen - Auto-Kandidat
PHASE2_BG   = RGBColor(0xCA, 0x8A, 0x04)   # gelb - Klaerung
PHASE3_BG   = RGBColor(0x6B, 0x72, 0x80)   # grau - out-of-scope
GREY_BG     = RGBColor(0xE5, 0xE7, 0xEB)
GREY_BORDER = RGBColor(0x9C, 0xA3, 0xAF)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT   = RGBColor(0x1F, 0x29, 0x37)
ACCENT_BLUE = RGBColor(0x1E, 0x40, 0xAF)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

slide = prs.slides.add_slide(prs.slide_layouts[6])   # blank

# ---------- HEADER ----------
def add_textbox(left, top, width, height, text, font_size=10, bold=False,
                color=DARK_TEXT, fill=None, border=None, align=PP_ALIGN.LEFT,
                anchor=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    if fill:
        tb.fill.solid(); tb.fill.fore_color.rgb = fill
    else:
        tb.fill.background()
    if border:
        tb.line.color.rgb = border
        tb.line.width = Pt(0.75)
    else:
        tb.line.fill.background()
    tf = tb.text_frame
    tf.margin_left = Pt(6); tf.margin_right = Pt(6)
    tf.margin_top = Pt(2);  tf.margin_bottom = Pt(2)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "Calibri"
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb

def add_multiline(left, top, width, height, lines, font_size=10, font_bold=False,
                  fill=None, border=None, color=DARK_TEXT, align=PP_ALIGN.LEFT,
                  line_spacing=1.05):
    """lines: list of either str or (str, dict) where dict has size/bold/color/align overrides."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    if fill:
        tb.fill.solid(); tb.fill.fore_color.rgb = fill
    else:
        tb.fill.background()
    if border:
        tb.line.color.rgb = border; tb.line.width = Pt(0.75)
    else:
        tb.line.fill.background()
    tf = tb.text_frame
    tf.margin_left = Pt(6); tf.margin_right = Pt(6)
    tf.margin_top = Pt(3);  tf.margin_bottom = Pt(3)
    tf.word_wrap = True
    for i, item in enumerate(lines):
        opts = {}
        if isinstance(item, tuple):
            txt, opts = item
        else:
            txt = item
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = opts.get("align", align)
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = txt
        r.font.name = opts.get("font", "Calibri")
        r.font.size = Pt(opts.get("size", font_size))
        r.font.bold = opts.get("bold", font_bold)
        r.font.italic = opts.get("italic", False)
        r.font.color.rgb = opts.get("color", color)
    return tb

# Title bar
title_h = Inches(0.55)
title_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, title_h)
title_bar.fill.solid(); title_bar.fill.fore_color.rgb = ERGO_DARK
title_bar.line.fill.background()
title_bar.text_frame.margin_left = Inches(0.4)
title_bar.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
p = title_bar.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "Automatisierungs-Potenzial Bestandsuebertragung ERGO   |   Stichwoche April 2026   |   ca. 5.000 Mails analysiert"
r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = WHITE

# Kernaussage banner (under title)
kern_top = title_h + Inches(0.08)
kern_h = Inches(0.45)
kern = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.3), kern_top, prs.slide_width - Inches(0.6), kern_h)
kern.fill.solid(); kern.fill.fore_color.rgb = PHASE1_BG
kern.line.fill.background()
kern.text_frame.margin_left = Inches(0.2); kern.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
p = kern.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "▶  Kernaussage: 65 % der Mails sind regelbasiert automatisierbare BUe-Neuanfragen → rund 170.000 Vorgaenge p.a."
r.font.name = "Calibri"; r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = WHITE

# ---------- FUNNEL (LINKS) ----------
funnel_top  = kern_top + kern_h + Inches(0.18)
funnel_left = Inches(0.3)
funnel_w    = Inches(4.7)

add_textbox(funnel_left, funnel_top, funnel_w, Inches(0.32),
            "FUNNEL: Vom Eingang zum Automatisierungs-Kandidat",
            font_size=12, bold=True, color=ERGO_DARK)

# Funnel-Stufen
stages = [
    ("ca. 5.000",   "Eingang gesamt",            "100 %", WHITE, ERGO_DARK),
    ("– 6 %",       "Rauschen (Bounce / System / Werbung)", "", GREY_BG, DARK_TEXT),
    ("ca. 4.700",   "Echte Makler-Vorgaenge",     "94 %",  WHITE, ERGO_DARK),
    ("– 4 %",       "Andere Klassifikation (Anfrage / Antrag)", "", GREY_BG, DARK_TEXT),
    ("ca. 4.500",   "BUe-Vorgaenge",              "90 %",  WHITE, ERGO_DARK),
    ("– 7 %",       "Reminder / Wiedervorlage",   "",      GREY_BG, DARK_TEXT),
    ("ca. 4.150",   "BUe-Erstvorgaenge",          "83 %",  WHITE, ERGO_DARK),
    ("– 18 %",      "MV-Klaerungsbedarf",         "",      GREY_BG, DARK_TEXT),
    ("ca. 3.250",   "AUTOMATISIERUNGS-KANDIDATEN", "65 %", PHASE1_BG, WHITE),
]

y = funnel_top + Inches(0.32)
row_h = Inches(0.36)
for absol, label, pct, bg, fg in stages:
    is_loss = absol.startswith("–")
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, funnel_left, y, funnel_w, row_h)
    box.fill.solid(); box.fill.fore_color.rgb = bg
    if is_loss:
        box.line.color.rgb = GREY_BORDER; box.line.width = Pt(0.5)
    else:
        box.line.color.rgb = ERGO_DARK if fg == WHITE else GREY_BORDER
        box.line.width = Pt(0.5)
    box.text_frame.margin_left = Pt(8); box.text_frame.margin_right = Pt(8)
    box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    # Left: absolut
    r1 = p.add_run(); r1.text = absol + "   "
    r1.font.name = "Calibri"; r1.font.size = Pt(11 if not is_loss else 9)
    r1.font.bold = True; r1.font.italic = is_loss; r1.font.color.rgb = fg
    # Middle: label
    r2 = p.add_run(); r2.text = label
    r2.font.name = "Calibri"; r2.font.size = Pt(10 if not is_loss else 9)
    r2.font.italic = is_loss; r2.font.color.rgb = fg
    # Right: pct (in separate paragraph alignment? - use trailing tab)
    if pct:
        r3 = p.add_run(); r3.text = "    " + pct
        r3.font.name = "Calibri"; r3.font.size = Pt(11)
        r3.font.bold = True; r3.font.color.rgb = fg
    y = y + row_h + Inches(0.02)

# Anmerkung unter dem Funnel
y_end = y + Inches(0.04)
add_multiline(funnel_left, y_end, funnel_w, Inches(0.4),
              [("Definition Automatisierungs-Kandidat: vollstaendige Maklervollmacht + Erstvorgang + BUe-Klassifikation.", {"size":8, "italic":True, "color":RGBColor(0x6b,0x72,0x80)})])

# ---------- MITTLERE SPALTE: Phasen ----------
mid_left = funnel_left + funnel_w + Inches(0.25)
mid_w    = Inches(4.0)

add_textbox(mid_left, funnel_top, mid_w, Inches(0.32),
            "Drei Phasen der Bearbeitung",
            font_size=12, bold=True, color=ERGO_DARK)

# Phase 1 Box
p1_top = funnel_top + Inches(0.34)
p1_h   = Inches(1.65)
add_multiline(mid_left, p1_top, mid_w, p1_h, [
    ("PHASE 1   Sofort automatisierbar   65 %", {"size":11, "bold":True, "color":WHITE}),
    ("Sparten: Komposit ~ 60 %  |  KV ~ 30 %  |  Leben ~ 20 %", {"size":9, "color":WHITE}),
    ("Geschaeftstyp: Einfacher Vertrag ~ 85 %  |  Kundenverbindung ~ 20 %", {"size":9, "color":WHITE}),
    ("Sonderfaelle: < 1 %  (manuell ausgenommen)", {"size":9, "color":WHITE}),
    ("Pool-Konzentration: Top 3 Pools decken ~ 65 % ab", {"size":9, "color":WHITE, "bold":True}),
    ("→ Schnittstellen-Priorisierung", {"size":9, "color":WHITE, "italic":True}),
], fill=PHASE1_BG, border=PHASE1_BG)

# Phase 2 Box
p2_top = p1_top + p1_h + Inches(0.12)
p2_h   = Inches(1.55)
add_multiline(mid_left, p2_top, mid_w, p2_h, [
    ("PHASE 2   Klaerungsbedarf   ~ 18 %", {"size":11, "bold":True, "color":WHITE}),
    ("~ 1/3   MV unvollstaendig    →  Klaerungs-Template", {"size":9, "color":WHITE}),
    ("~ 1/3   kein Anhang          →  Outbound: MV nachfordern", {"size":9, "color":WHITE}),
    ("~ 1/3   MV fehlt im Anhang   →  Outbound: MV nachfordern", {"size":9, "color":WHITE}),
    ("< 5 %   nicht pruefbar       →  manuelle Sichtung", {"size":9, "color":WHITE}),
    ("→ Halbautomatischer Workflow mit Standard-Anschreiben", {"size":9, "color":WHITE, "italic":True}),
], fill=PHASE2_BG, border=PHASE2_BG)

# Phase 3 Box
p3_top = p2_top + p2_h + Inches(0.12)
p3_h   = Inches(1.35)
add_multiline(mid_left, p3_top, mid_w, p3_h, [
    ("Ausserhalb Scope   ~ 17 %", {"size":11, "bold":True, "color":WHITE}),
    ("~ 7 %   Reminder / Wiedervorlage   →  eigener SLA-Use-Case", {"size":9, "color":WHITE}),
    ("~ 4 %   Nicht-BUe                  →  andere Pipeline", {"size":9, "color":WHITE}),
    ("~ 6 %   Rauschen                   →  Spam-Filter / NDR", {"size":9, "color":WHITE}),
    ("< 1 %   Sonderfaelle                →  Konsortien manuell", {"size":9, "color":WHITE}),
], fill=PHASE3_BG, border=PHASE3_BG)

# ---------- RECHTE SPALTE: Top Pools + Hochrechnung ----------
right_left = mid_left + mid_w + Inches(0.25)
right_w    = prs.slide_width - right_left - Inches(0.3)

add_textbox(right_left, funnel_top, right_w, Inches(0.32),
            "Top 10 Pools in Automatisierungs-Kandidaten",
            font_size=12, bold=True, color=ERGO_DARK)

pools = [
    ("1.  JDC / Jung, DMS & Cie.",       40, True),   # konsolidiert
    ("2.  Fonds Finanz Maklerservice",   18, True),
    ("3.  blau direkt",                   7, True),
    ("4.  DEMA",                          3, False),
    ("5.  Securess Versicherungsmakler",  2, False),
    ("6.  1:1 Assekuranzservice",         2, False),
    ("7.  [pma:] Finanz-Service",         1, False),
    ("8.  Qualitypool",                   1, False),
    ("9.  Invers GmbH",                   1, False),
    ("10. Apella AG",                     1, False),
]

pool_y = funnel_top + Inches(0.34)
pool_row_h = Inches(0.32)
pool_label_w = Inches(2.1)
pool_bar_max_w = right_w - pool_label_w - Inches(0.55)
max_pct = 40
for label, pct, is_top3 in pools:
    # Label
    lbl = slide.shapes.add_textbox(right_left, pool_y, pool_label_w, pool_row_h)
    lbl.fill.background(); lbl.line.fill.background()
    lbl.text_frame.margin_left = Pt(2); lbl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = lbl.text_frame.paragraphs[0]; pp.alignment = PP_ALIGN.LEFT
    rr = pp.add_run(); rr.text = label
    rr.font.name = "Calibri"; rr.font.size = Pt(9.5)
    rr.font.bold = is_top3; rr.font.color.rgb = DARK_TEXT
    # Bar background (full)
    bar_bg_left = right_left + pool_label_w
    bg_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_bg_left,
                                     pool_y + Inches(0.07), pool_bar_max_w, Inches(0.18))
    bg_bar.fill.solid(); bg_bar.fill.fore_color.rgb = GREY_BG
    bg_bar.line.fill.background()
    # Bar value
    val_w = Emu(int(pool_bar_max_w * pct / max_pct))
    if val_w > 0:
        val_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_bg_left,
                                          pool_y + Inches(0.07), val_w, Inches(0.18))
        val_bar.fill.solid()
        val_bar.fill.fore_color.rgb = ACCENT_BLUE if is_top3 else RGBColor(0x60, 0xa5, 0xfa)
        val_bar.line.fill.background()
    # Percent label
    pct_lbl = slide.shapes.add_textbox(bar_bg_left + pool_bar_max_w + Inches(0.05),
                                        pool_y, Inches(0.55), pool_row_h)
    pct_lbl.fill.background(); pct_lbl.line.fill.background()
    pct_lbl.text_frame.margin_left = Pt(2); pct_lbl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp2 = pct_lbl.text_frame.paragraphs[0]; pp2.alignment = PP_ALIGN.LEFT
    rr2 = pp2.add_run(); rr2.text = "ca. " + str(pct) + " %"
    rr2.font.name = "Calibri"; rr2.font.size = Pt(9.5)
    rr2.font.bold = is_top3; rr2.font.color.rgb = ACCENT_BLUE if is_top3 else DARK_TEXT
    pool_y = pool_y + pool_row_h

# Pool-Summary
sum_y = pool_y + Inches(0.05)
add_multiline(right_left, sum_y, right_w, Inches(0.4), [
    ("Top 10 zusammen ca. 75 %  |  Long Tail (100+ Pools) ca. 25 %", {"size":9, "italic":True, "color":RGBColor(0x4b,0x55,0x63)})
])

# ---------- UNTEN: Jahres-Hochrechnung + Empfehlung ----------
bottom_top = Inches(6.05)
bottom_h   = Inches(1.4)

# Jahres-Hochrechnung links
hr_w = Inches(6.5)
add_multiline(Inches(0.3), bottom_top, hr_w, bottom_h, [
    ("Hochrechnung auf das Jahr (annahmegestuetzt)", {"size":11, "bold":True, "color":ERGO_DARK}),
    ("ca. 170.000   Phase 1   Automatisierungs-Potenzial", {"size":10, "color":DARK_TEXT, "bold":True}),
    ("ca.  45.000   Phase 2   Halbautomatische Klaerung", {"size":10, "color":DARK_TEXT}),
    ("ca.  18.000   Phase 3   Reminder- / SLA-Logik", {"size":10, "color":DARK_TEXT}),
    ("ca.  10.000   Manuell   Sonstige + Eskalationen", {"size":10, "color":DARK_TEXT}),
    ("ca.  15.000   Rauschen  (NDR / Spam-Filter)", {"size":10, "color":DARK_TEXT}),
    ("ca. 260.000   GESAMT-Mails p.a. im SHUK-Innendienst", {"size":10, "color":ERGO_DARK, "bold":True}),
], fill=RGBColor(0xF3, 0xF4, 0xF6), border=GREY_BORDER)

# Empfehlung rechts
emp_left = Inches(0.3) + hr_w + Inches(0.25)
emp_w = prs.slide_width - emp_left - Inches(0.3)
add_multiline(emp_left, bottom_top, emp_w, bottom_h, [
    ("Empfehlung", {"size":11, "bold":True, "color":WHITE}),
    ("Phase-1-Pilot mit Top 3 Pools (JDC, Fonds Finanz, blau direkt) deckt mit nur DREI Schnittstellen rund zwei Drittel des Automatisierungs-Potenzials ab.", {"size":11, "color":WHITE}),
    ("", {"size":4}),
    ("Folge-Phasen: MV-Klaerungs-Templates fuer Top 3 Klaerungsgruende → +18 % Volumen mit halbautomatischem Workflow.", {"size":10, "color":WHITE, "italic":True}),
], fill=ERGO_RED, border=ERGO_RED)

# Footer / Datenquelle
add_textbox(Inches(0.3), Inches(7.25), Inches(13.0), Inches(0.22),
            "Datenquelle: ergo-vorgang-analyse.hta v1.62 | KI-gestuetzte Klassifikation Mail + PDF-Anhang | Pool-Schreibweisen konsolidiert | Eine Stichwoche, keine Saisonalitaet beruecksichtigt",
            font_size=7.5, italic=True, color=RGBColor(0x6b, 0x72, 0x80))

out = "/home/user/emailimport/Auswertung-BUe-Stichwoche-Vorschlag.pptx"
prs.save(out)
import os
print("Saved:", out, "size:", os.path.getsize(out), "bytes")
