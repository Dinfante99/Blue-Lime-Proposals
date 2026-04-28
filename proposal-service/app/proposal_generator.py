"""
Blue Lime proposal generator.

This is the layout/rendering engine. It takes a `ProposalData` dictionary
(produced by `excel_parser.py`) and writes a polished, branded PDF to the
given output path.

The visual layout is identical to what we built and approved on the
Haven at Keith Harrow proposal — only the data flows in dynamically.
"""
from __future__ import annotations

import os
from typing import Any

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether, Flowable, NextPageTemplate, Image,
)

# -----------------------------------------------------------------------------
# BRAND ASSETS  — paths resolve relative to the project root in the container
# -----------------------------------------------------------------------------
HERE        = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(HERE)            # one level up from /app
BRAND_DIR   = os.path.join(PROJECT_ROOT, "brand")
LOGO_PATH      = os.path.join(BRAND_DIR, "logo.png")
WATERMARK_PATH = os.path.join(BRAND_DIR, "watermark.png")
COVER_BG_PATH  = os.path.join(BRAND_DIR, "cover_bg.png")
HEADSHOT_DIR   = os.path.join(BRAND_DIR, "headshots")

# -----------------------------------------------------------------------------
# COLORS
# -----------------------------------------------------------------------------
BL_BLUE      = HexColor("#3BB5E8")
BL_BLUE_DK   = HexColor("#1E90C4")
BL_BLUE_LT   = HexColor("#D8F0FB")
BL_BLUE_XL   = HexColor("#EEF8FD")
BL_NAVY      = HexColor("#0F2A3D")
BL_TEXT      = HexColor("#22313F")
BL_MUTED     = HexColor("#6B7A86")
BL_RULE      = HexColor("#D9E3EC")
BL_PANEL     = HexColor("#F4F8FB")

PAGE_W, PAGE_H = LETTER
LEFT, RIGHT, TOP, BOTTOM = 0.7 * inch, 0.7 * inch, 0.7 * inch, 0.8 * inch
CONTENT_W = PAGE_W - LEFT - RIGHT


def esc(s: Any) -> str:
    """Escape ampersands so the ReportLab Paragraph parser doesn't choke."""
    return str(s).replace("&", "&amp;")


# =============================================================================
# STYLES (built once, module-level — these never change)
# =============================================================================
styles = {
    "H1":     ParagraphStyle("H1", fontName="Helvetica-Bold", fontSize=22,
                             leading=26, textColor=BL_NAVY, spaceAfter=6),
    "H2":     ParagraphStyle("H2", fontName="Helvetica-Bold", fontSize=14,
                             leading=18, textColor=BL_NAVY, spaceBefore=6, spaceAfter=4),
    "H3":     ParagraphStyle("H3", fontName="Helvetica-Bold", fontSize=11.5,
                             leading=14, textColor=BL_BLUE_DK, spaceBefore=2, spaceAfter=2),
    "Body":   ParagraphStyle("Body", fontName="Helvetica", fontSize=10,
                             leading=14, textColor=BL_TEXT, alignment=TA_LEFT),
    "Body-J": ParagraphStyle("Body-J", fontName="Helvetica", fontSize=10,
                             leading=14, textColor=BL_TEXT, alignment=TA_JUSTIFY),
    "Lede":   ParagraphStyle("Lede", fontName="Helvetica", fontSize=11,
                             leading=16, textColor=BL_TEXT, alignment=TA_LEFT, spaceAfter=6),
    "Caption": ParagraphStyle("Caption", fontName="Helvetica", fontSize=8.5,
                              leading=11, textColor=BL_MUTED),
    "Label":   ParagraphStyle("Label", fontName="Helvetica-Bold", fontSize=8.5,
                              leading=11, textColor=BL_MUTED),
    "Value":   ParagraphStyle("Value", fontName="Helvetica-Bold", fontSize=11,
                              leading=14, textColor=BL_NAVY),
    "Small":   ParagraphStyle("Small", fontName="Helvetica", fontSize=9,
                              leading=12, textColor=BL_TEXT),
    "SmallJ":  ParagraphStyle("SmallJ", fontName="Helvetica", fontSize=9,
                              leading=12, textColor=BL_TEXT, alignment=TA_JUSTIFY),
    "TH":      ParagraphStyle("TH", fontName="Helvetica-Bold", fontSize=9.5,
                              leading=12, textColor=white),
    "TD":      ParagraphStyle("TD", fontName="Helvetica", fontSize=9.5,
                              leading=12, textColor=BL_TEXT),
    "TDb":     ParagraphStyle("TDb", fontName="Helvetica-Bold", fontSize=9.5,
                              leading=12, textColor=BL_NAVY),
    "Hero":    ParagraphStyle("Hero", fontName="Helvetica-Bold", fontSize=22,
                              leading=26, textColor=BL_NAVY),
    "HeroBlue": ParagraphStyle("HeroBlue", fontName="Helvetica-Bold", fontSize=22,
                               leading=26, textColor=BL_BLUE_DK),
    "BigName": ParagraphStyle("BigName", fontName="Helvetica-Bold", fontSize=18,
                              leading=22, textColor=BL_NAVY),
    "SubName": ParagraphStyle("SubName", fontName="Helvetica-Bold", fontSize=11,
                              leading=14, textColor=BL_BLUE_DK),
}


def p(text: str, style: str = "Body") -> Paragraph:
    return Paragraph(text, styles[style])


# =============================================================================
# PAGE FURNITURE  (cover + interior chrome)
# =============================================================================
def _make_cover_drawer(data):
    """Returns a closure that draws the cover page using the data."""
    def draw_cover(canv, doc):
        client = data["client"]
        am = data["account_manager"]

        canv.drawImage(COVER_BG_PATH, 0, 0, width=PAGE_W, height=PAGE_H,
                       preserveAspectRatio=False, mask='auto')

        # Real Blue Lime logo
        logo_w = 3.6 * inch
        logo_h = logo_w / (600 / 203)
        canv.drawImage(LOGO_PATH,
                       (PAGE_W - logo_w) / 2,
                       PAGE_H - 1.5 * inch - logo_h,
                       width=logo_w, height=logo_h,
                       mask='auto', preserveAspectRatio=True)

        cx = PAGE_W / 2
        canv.setFillColor(white)
        canv.setFont("Helvetica-Bold", 34)
        canv.drawCentredString(cx, PAGE_H / 2 - 10, "Insurance Proposal")

        canv.setFont("Helvetica", 14)
        canv.setFillColor(HexColor("#EAF7FD"))
        canv.drawCentredString(cx, PAGE_H / 2 - 38, "Prepared exclusively for")

        canv.setFillColor(white)
        canv.setFont("Helvetica-Bold", 20)
        # If the client name is long, split it across two lines on the cover
        name = client["name"]
        if " Owners Association" in name:
            short, suffix = name.split(" Owners Association", 1)
            canv.drawCentredString(cx, PAGE_H / 2 - 68, short.strip())
            canv.setFont("Helvetica-Bold", 16)
            canv.drawCentredString(cx, PAGE_H / 2 - 90, "Owners Association" + suffix)
        else:
            canv.drawCentredString(cx, PAGE_H / 2 - 68, name)

        canv.setStrokeColor(white)
        canv.setLineWidth(1)
        canv.line(cx - 70, PAGE_H / 2 - 110, cx + 70, PAGE_H / 2 - 110)

        canv.setFont("Helvetica", 12)
        canv.setFillColor(HexColor("#EAF7FD"))
        canv.drawCentredString(cx, PAGE_H / 2 - 130, "Policy Term")
        canv.setFont("Helvetica-Bold", 14)
        canv.setFillColor(white)
        canv.drawCentredString(cx, PAGE_H / 2 - 150, client["policy_term"])

        # Prepared-by card
        box_x, box_y, box_w, box_h = LEFT, BOTTOM, CONTENT_W, 1.15 * inch
        canv.setFillColor(white)
        canv.setStrokeColor(white)
        canv.roundRect(box_x, box_y, box_w, box_h, 8, stroke=0, fill=1)

        canv.setFillColor(BL_MUTED)
        canv.setFont("Helvetica", 9)
        canv.drawString(box_x + 20, box_y + box_h - 22, "PREPARED BY")

        canv.setFillColor(BL_NAVY)
        canv.setFont("Helvetica-Bold", 13)
        canv.drawString(box_x + 20, box_y + box_h - 42,
                        f"{am['name']} \u00b7 {am['title']}")

        canv.setFillColor(BL_TEXT)
        canv.setFont("Helvetica", 10)
        canv.drawString(box_x + 20, box_y + box_h - 58, am["email"])
        canv.drawString(box_x + 20, box_y + box_h - 72, am["phone"])

        canv.setFillColor(BL_MUTED)
        canv.setFont("Helvetica", 8.5)
        canv.drawRightString(box_x + box_w - 20, box_y + 14,
                             "Blue Lime Insurance Group  \u00b7  www.bluelimeins.com  \u00b7  contact@bluelimeins.com")
    return draw_cover


def _make_interior_drawer(data):
    def draw_interior(canv, doc):
        canv.saveState()
        band_h = 0.42 * inch
        canv.drawImage(WATERMARK_PATH, 0, PAGE_H - band_h, width=PAGE_W, height=band_h,
                       preserveAspectRatio=False, mask='auto')

        logo_h = 0.32 * inch
        logo_w = logo_h * (600 / 203)
        pad = 4
        canv.setFillColor(white)
        canv.roundRect(LEFT - pad, PAGE_H - band_h + (band_h - logo_h) / 2 - pad,
                       logo_w + 2 * pad, logo_h + 2 * pad, 4, stroke=0, fill=1)
        canv.drawImage(LOGO_PATH, LEFT, PAGE_H - band_h + (band_h - logo_h) / 2,
                       width=logo_w, height=logo_h, mask='auto', preserveAspectRatio=True)

        client_short = data["client"]["short_name"]
        canv.setFillColor(white)
        canv.setFont("Helvetica-Bold", 9.5)
        canv.drawRightString(PAGE_W - RIGHT, PAGE_H - band_h / 2 - 3.5,
                             f"Insurance Proposal  \u00b7  {client_short}")

        canv.setStrokeColor(BL_RULE)
        canv.setLineWidth(0.5)
        canv.line(LEFT, BOTTOM - 0.25 * inch, PAGE_W - RIGHT, BOTTOM - 0.25 * inch)

        canv.setFont("Helvetica", 8.5)
        canv.setFillColor(BL_MUTED)
        canv.drawString(LEFT, BOTTOM - 0.42 * inch,
                        "Blue Lime Insurance Group  \u00b7  San Antonio, TX  \u00b7  contact@bluelimeins.com")
        canv.drawRightString(PAGE_W - RIGHT, BOTTOM - 0.42 * inch,
                             f"Page {canv.getPageNumber() - 1}")
        canv.restoreState()
    return draw_interior


# =============================================================================
# FLOWABLES & HELPERS
# =============================================================================
class SectionBanner(Flowable):
    def __init__(self, text, width=None, height=30, fill=BL_BLUE):
        super().__init__()
        self.text = text
        self.width = width or CONTENT_W
        self.height = height
        self.fill = fill

    def draw(self):
        c = self.canv
        c.setFillColor(self.fill)
        c.roundRect(0, 0, self.width, self.height, 6, stroke=0, fill=1)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(14, self.height / 2 - 4.5, self.text)


def section(title):
    return [SectionBanner(title), Spacer(1, 10)]


# =============================================================================
# PAGE BUILDERS
# =============================================================================
def _summary_page(data):
    client = data["client"]
    am = data["account_manager"]
    premium = data["premium"]

    story = [
        p("Proposal Overview", "H1"), Spacer(1, 2),
        p("A tailored insurance program for your community association.", "Lede"),
        Spacer(1, 6),
    ]

    info = [
        [p("ASSOCIATION", "Label"),       p("POLICY TERM", "Label"),       p("ACCOUNT MANAGER", "Label")],
        [p(f"<b>{esc(client['name'])}</b>", "Small"),
         p(f"<b>{esc(client['policy_term'])}</b>", "Small"),
         p(f"<b>{esc(am['name'])}</b>", "Small")],
        [p(esc(client['address']), "Small"),
         p(esc(client.get('homes_summary', '')), "Small"),
         p(f"{esc(am['email'])}<br/>{esc(am['phone'])}", "Small")],
    ]
    t = Table(info, colWidths=[CONTENT_W * 0.36, CONTENT_W * 0.32, CONTENT_W * 0.32])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BL_BLUE_XL),
        ("BOX",           (0, 0), (-1, -1), 0.5, BL_RULE),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LINEAFTER",     (0, 0), (0, -1), 0.4, BL_RULE),
        ("LINEAFTER",     (1, 0), (1, -1), 0.4, BL_RULE),
    ]))
    story.append(t)
    story.append(Spacer(1, 18))

    cols = [
        [p("What's Inside", "H3"),
         p("A premium comparison against expiring coverage, a line-by-line "
           "breakdown of each coverage part, your Statement of Values, an "
           "authorization page to bind coverage, and disclosures.", "Small")],
        [p("Why Blue Lime", "H3"),
         p("We exclusively serve community associations. Over 125 combined "
           "years of expertise and long-standing carrier relationships let "
           "us deliver the right coverage at the right price.", "Small")],
        [p("Next Steps", "H3"),
         p("Review the coverage detail and premium comparison. When you're "
           "ready to bind, sign the authorization page and return to "
           "contact@bluelimeins.com. We'll handle the rest.", "Small")],
    ]
    t = Table([cols], colWidths=[CONTENT_W / 3] * 3)
    t.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LINEABOVE",     (0, 0), (-1, 0), 2.5, BL_BLUE),
        ("BACKGROUND",    (0, 0), (-1, -1), BL_PANEL),
        ("LINEAFTER",     (0, 0), (0, -1), 0.4, BL_RULE),
        ("LINEAFTER",     (1, 0), (1, -1), 0.4, BL_RULE),
    ]))
    story.append(t)
    story.append(Spacer(1, 18))
    story.extend(section("Premium at a Glance"))

    expiring = [
        p("EXPIRING PREMIUM", "Label"), Spacer(1, 4),
        p(esc(premium["expiring_total_str"]), "Hero"), Spacer(1, 2),
        p("Annual, prior term", "Caption"),
    ]
    proposed = [
        p("PROPOSED PREMIUM", "Label"), Spacer(1, 4),
        p(esc(premium["proposed_total_str"]), "HeroBlue"), Spacer(1, 2),
        p("Annual, includes agency fee", "Caption"),
    ]
    change_amt = premium["proposed_total"] - premium["expiring_total"]
    sign = "+" if change_amt >= 0 else "\u2212"
    change = [
        p("CHANGE", "Label"), Spacer(1, 4),
        p(f"{sign}${abs(change_amt):,.2f}", "Hero"), Spacer(1, 2),
        p(esc(premium.get("change_note", "")), "Caption"),
    ]
    t = Table([[expiring, proposed, change]], colWidths=[CONTENT_W / 3] * 3)
    t.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
        ("TOPPADDING",    (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("BACKGROUND",    (0, 0), (0, 0), BL_PANEL),
        ("BACKGROUND",    (1, 0), (1, 0), BL_BLUE_XL),
        ("BACKGROUND",    (2, 0), (2, 0), BL_PANEL),
        ("BOX",           (0, 0), (-1, -1), 0.5, BL_RULE),
        ("LINEAFTER",     (0, 0), (0, -1), 0.5, BL_RULE),
        ("LINEAFTER",     (1, 0), (1, -1), 0.5, BL_RULE),
    ]))
    story.append(t)
    return story


def _premium_comparison_page(data):
    premium = data["premium"]
    story = section("Premium Summary Comparison")
    story.append(p("Side-by-side comparison of your expiring and proposed programs across every coverage line.", "Body"))
    story.append(Spacer(1, 10))

    expiring_carrier = premium.get("expiring_carrier", "Expiring")
    proposed_carrier = premium.get("proposed_carrier", "Proposed")

    header = [
        p("Coverage (Limits)", "TH"),
        p(f"Expiring<br/><font size=8>{esc(expiring_carrier)}</font>", "TH"),
        p(f"Proposed<br/><font size=8>{esc(proposed_carrier)}</font>", "TH"),
    ]
    rows = [header]
    for line in premium["comparison_lines"]:
        rows.append([
            p(esc(line["label"]), "TDb"),
            p(esc(line["expiring"]), "TD"),
            p(esc(line["proposed"]), "TD"),
        ])

    white_para = ParagraphStyle("white", fontName="Helvetica-Bold",
                                fontSize=10.5, textColor=white, leading=13)
    rows.append([
        Paragraph("Total Annual Premium", white_para),
        Paragraph(esc(premium["expiring_total_str"]), white_para),
        Paragraph(esc(premium["proposed_total_str"]), white_para),
    ])

    col_w = [CONTENT_W * 0.42, CONTENT_W * 0.29, CONTENT_W * 0.29]
    t = Table(rows, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0), BL_BLUE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [white, BL_BLUE_XL]),
        ("BACKGROUND",    (0, -1), (-1, -1), BL_NAVY),
        ("TOPPADDING",    (0, -1), (-1, -1), 10),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 10),
        ("BOX",           (0, 0), (-1, -1), 0.5, BL_RULE),
        ("LINEBEFORE",    (1, 0), (1, -1), 0.4, BL_RULE),
        ("LINEBEFORE",    (2, 0), (2, -1), 0.4, BL_RULE),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(p(
        "Expiring premium reflects the prior-term program. Proposed premium "
        "includes all coverage parts shown plus any applicable agency fee. "
        "Limits and deductibles shown are summary; policy forms govern.",
        "Caption"))
    return story


def _coverage_panel(pairs):
    cells = []
    for label, value in pairs:
        cells.append([p(esc(label).upper(), "Label"), Spacer(1, 2),
                      p(f"<b>{esc(value)}</b>", "Small")])
    col_w = CONTENT_W / max(len(pairs), 1)
    t = Table([cells], colWidths=[col_w] * len(pairs))
    t.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("BACKGROUND",    (0, 0), (-1, -1), BL_PANEL),
        ("LINEAFTER",     (0, 0), (-2, -1), 0.4, BL_RULE),
        ("BOX",           (0, 0), (-1, -1), 0.4, BL_RULE),
    ]))
    return t


def _coverage_block(title, carrier, description, pair_rows, carrier_label="Carrier"):
    flows = []
    left  = p(f"<font color='#0F2A3D' size=13><b>{esc(title)}</b></font>", "Body")
    right = Paragraph(
        f"<font color='#6B7A86' size=8>{esc(carrier_label).upper()}</font><br/>"
        f"<font color='#1E90C4' size=10.5><b>{esc(carrier)}</b></font>",
        styles["Body"])
    hdr = Table([[left, right]], colWidths=[CONTENT_W * 0.62, CONTENT_W * 0.38])
    hdr.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LINEBELOW",     (0, 0), (-1, 0), 2, BL_BLUE),
    ]))
    flows.append(hdr)
    flows.append(Spacer(1, 6))
    flows.append(p(esc(description), "SmallJ"))
    flows.append(Spacer(1, 8))
    for pairs in pair_rows:
        flows.append(_coverage_panel(pairs))
        flows.append(Spacer(1, 6))
    flows.append(Spacer(1, 10))
    return KeepTogether(flows)


def _coverage_pages(data):
    story = section("Coverage Detail")
    story.append(p("Each line of coverage included in your proposed program, with key limits and deductibles.", "Body"))
    story.append(Spacer(1, 10))

    for cov in data["coverages"]:
        carrier_label = "Status" if cov.get("not_included") else "Carrier"
        story.append(_coverage_block(
            cov["title"],
            cov["carrier"],
            cov["description"],
            cov["panel_rows"],
            carrier_label=carrier_label,
        ))
    return story


def _sov_page(data):
    sov = data["sov"]
    client = data["client"]
    story = section("Statement of Values")
    story.append(p("Every insured asset, documented. Values are provided by the association and should be reviewed annually to ensure appropriate property coverage remains in place.", "Body"))
    story.append(Spacer(1, 10))

    strip = Table([[
        [p("ASSOCIATION", "Label"), p(f"<b>{esc(client['name'])}</b>", "Small")],
        [p("HOMES / FINAL BUILDOUT", "Label"), p(f"<b>{esc(client.get('homes', ''))}</b>", "Small")],
        [p("PROPERTY ADDRESS", "Label"), p(f"<b>{esc(client['address'])}</b>", "Small")],
    ]], colWidths=[CONTENT_W * 0.35, CONTENT_W * 0.25, CONTENT_W * 0.40])
    strip.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BL_BLUE_XL),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LINEAFTER",     (0, 0), (0, -1), 0.4, BL_RULE),
        ("LINEAFTER",     (1, 0), (1, -1), 0.4, BL_RULE),
        ("BOX",           (0, 0), (-1, -1), 0.5, BL_RULE),
    ]))
    story.append(strip)
    story.append(Spacer(1, 14))

    hdr = [p("Coverage", "TH"), p("# of Units", "TH"), p("Area (sq ft)", "TH"), p("Value", "TH")]
    rows = [hdr]
    section_bg_indices = []
    subtotal_above_indices = []
    row_idx = 1

    for sec in sov["sections"]:
        rows.append([p(f"<b>{esc(sec['name'])}</b>", "TDb"),
                     p("", "TD"), p("", "TD"), p("", "TD")])
        section_bg_indices.append(row_idx)
        row_idx += 1

        for item in sec["items"]:
            rows.append([
                p(f"  {esc(item['name'])}", "TD"),
                p(esc(item.get('units', '\u2014')), "TD"),
                p(esc(item.get('area', '\u2014')), "TD"),
                p(f"${item['value']:,.0f}" if isinstance(item['value'], (int, float)) else esc(item['value']), "TD"),
            ])
            row_idx += 1

        if "subtotal_label" in sec:
            rows.append([
                p(f"<b>{esc(sec['subtotal_label'])}</b>", "TDb"),
                p("", "TD"), p("", "TD"),
                p(f"<b>${sec['subtotal']:,.0f}</b>", "TDb"),
            ])
            subtotal_above_indices.append(row_idx)
            row_idx += 1

    white_para = ParagraphStyle("white", fontName="Helvetica-Bold",
                                fontSize=10.5, textColor=white, leading=13)
    rows.append([
        Paragraph("Total Property Coverage", white_para),
        Paragraph("", white_para),
        Paragraph("", white_para),
        Paragraph(f"${sov['total']:,.0f}", white_para),
    ])

    col_w = [CONTENT_W * 0.45, CONTENT_W * 0.17, CONTENT_W * 0.18, CONTENT_W * 0.20]
    t = Table(rows, colWidths=col_w, repeatRows=1)
    base_style = [
        ("BACKGROUND",    (0, 0), (-1, 0), BL_BLUE),
        ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0), (0, 0), "LEFT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("BACKGROUND",    (0, -1), (-1, -1), BL_NAVY),
        ("TOPPADDING",    (0, -1), (-1, -1), 10),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 10),
        ("BOX",           (0, 0), (-1, -1), 0.5, BL_RULE),
        ("LINEBELOW",     (0, 0), (-1, 0), 0.5, BL_BLUE_DK),
        ("LINEBEFORE",    (1, 0), (1, -2), 0.3, BL_RULE),
        ("LINEBEFORE",    (2, 0), (2, -2), 0.3, BL_RULE),
        ("LINEBEFORE",    (3, 0), (3, -2), 0.3, BL_RULE),
    ]
    for idx in section_bg_indices:
        base_style.append(("BACKGROUND", (0, idx), (-1, idx), BL_BLUE_LT))
    for idx in subtotal_above_indices:
        base_style.append(("LINEABOVE", (0, idx), (-1, idx), 0.8, BL_BLUE_DK))
    t.setStyle(TableStyle(base_style))
    story.append(t)

    story.append(Spacer(1, 14))
    sig = Table([[
        [p("SIGNED &amp; ACCEPTED", "Label"), Spacer(1, 30),
         Paragraph("_____________________________", styles["Small"]),
         p("Authorized Association Representative", "Caption")],
        [p("DATE", "Label"), Spacer(1, 30),
         Paragraph("_____________________________", styles["Small"]),
         p("MM / DD / YYYY", "Caption")],
    ]], colWidths=[CONTENT_W * 0.58, CONTENT_W * 0.42])
    sig.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("BACKGROUND",    (0, 0), (-1, -1), BL_PANEL),
        ("LINEAFTER",     (0, 0), (0, -1), 0.4, BL_RULE),
        ("BOX",           (0, 0), (-1, -1), 0.4, BL_RULE),
    ]))
    story.append(sig)
    return story


def _authorization_page(data):
    auth = data["authorization"]
    client = data["client"]
    story = section("Authorization to Bind")
    story.append(p(
        f"After careful consideration and review, the authorized "
        f"representative for <b>{esc(client['name'])}</b> hereby accepts the "
        "proposal of insurance provided by Blue Lime Insurance Group and "
        "agrees to the payment and/or financing of the premiums outlined "
        "below. This acceptance serves as authorization to bind coverage.",
        "SmallJ"))
    story.append(Spacer(1, 12))

    hdr = [p("Policy Type", "TH"), p("Proposed Premium", "TH"),
           p("Option 1", "TH"), p("Option 2", "TH")]
    rows = [hdr]
    for line in auth["policy_lines"]:
        rows.append([
            p(esc(line["label"]), "TDb"),
            p(esc(line["proposed"]), "TD"),
            p(esc(line.get("option1", "\u2014")), "TD"),
            p(esc(line.get("option2", "\u2014")), "TD"),
        ])

    white_para = ParagraphStyle("white", fontName="Helvetica-Bold",
                                fontSize=10.5, textColor=white, leading=13)
    rows.append([
        Paragraph("Total Premium With Fees", white_para),
        Paragraph(esc(auth["total_str"]), white_para),
        Paragraph("\u2014", white_para),
        Paragraph("\u2014", white_para),
    ])

    col_w = [CONTENT_W * 0.43, CONTENT_W * 0.21, CONTENT_W * 0.18, CONTENT_W * 0.18]
    t = Table(rows, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0), BL_BLUE),
        ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [white, BL_BLUE_XL]),
        ("BACKGROUND",    (0, -1), (-1, -1), BL_NAVY),
        ("TOPPADDING",    (0, -1), (-1, -1), 10),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 10),
        ("BOX",           (0, 0), (-1, -1), 0.5, BL_RULE),
        ("LINEBEFORE",    (1, 0), (1, -1), 0.4, BL_RULE),
        ("LINEBEFORE",    (2, 0), (2, -1), 0.4, BL_RULE),
        ("LINEBEFORE",    (3, 0), (3, -1), 0.4, BL_RULE),
    ]))
    story.append(t)

    story.append(Spacer(1, 16))
    story.append(p("Payment Options", "H3"))
    story.append(Spacer(1, 4))

    pay_in_full = auth.get("pay_in_full_str", auth["total_str"])
    plan_amt = ParagraphStyle("pa", fontName="Helvetica-Bold", fontSize=14,
                              leading=18, textColor=BL_NAVY)

    pif = [p("PAY IN FULL", "Label"), Spacer(1, 6),
           p(pay_in_full, "Hero"), Spacer(1, 4),
           p("One-time payment. No financing fees.", "Caption")]

    if "down_payment_str" in auth:
        plan = [p("PAYMENT PLAN", "Label"), Spacer(1, 6),
                Paragraph(f"<b>{esc(auth['down_payment_str'])}</b> <font size=10 color='#6B7A86'>down</font>", plan_amt),
                Paragraph(f"<b>+ {auth['installments_count']} &times; {esc(auth['installment_amount_str'])}</b> "
                          f"<font size=10 color='#6B7A86'>monthly installments</font>", plan_amt),
                Spacer(1, 4),
                p("Premium financing available via our partner &mdash; no lengthy bank application.", "Caption")]
    else:
        plan = [p("PAYMENT PLAN", "Label"), Spacer(1, 6),
                p("Available on request", "Hero"), Spacer(1, 4),
                p("Premium financing options available &mdash; contact your account manager.", "Caption")]

    payt = Table([[pif, plan]], colWidths=[CONTENT_W / 2] * 2)
    payt.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
        ("TOPPADDING",    (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("BACKGROUND",    (0, 0), (0, 0), BL_PANEL),
        ("BACKGROUND",    (1, 0), (1, 0), BL_BLUE_XL),
        ("BOX",           (0, 0), (-1, -1), 0.4, BL_RULE),
        ("LINEAFTER",     (0, 0), (0, -1), 0.4, BL_RULE),
    ]))
    story.append(payt)

    story.append(Spacer(1, 20))
    story.append(p("Authorized Association Representative", "H3"))
    story.append(Spacer(1, 10))

    for row in [
        [("PRINTED NAME", ""), ("TITLE / ROLE", "")],
        [("SIGNATURE", ""),    ("DATE", "")],
    ]:
        cells = [[p(label, "Label"), Spacer(1, 28),
                  Paragraph("_____________________________", styles["Small"])]
                 for label, _ in row]
        tbl = Table([cells], colWidths=[CONTENT_W / 2] * 2)
        tbl.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 14),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("LINEAFTER",     (0, 0), (0, -1), 0.4, BL_RULE),
            ("BOX",           (0, 0), (-1, -1), 0.4, BL_RULE),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 6))

    story.append(Spacer(1, 10))
    story.append(p("Please return the completed authorization to Blue Lime Insurance Group at "
                   "<b>contact@bluelimeins.com</b>.", "Small"))
    return story


def _disclosures_page():
    story = section("Disclosures")
    paragraphs = [
        "The coverage options provided within this proposal were based on the information submitted by the client in the original application. For convenience we have summarized the coverage options by including general policy limits, deductibles, policy term, and premiums; however, this summary proposal does not include all of the terms, coverages, exclusions, limitations, definitions, or conditions of the actual policies. Please read the actual policies for full coverage detail.",
        "Property value limits within the proposal were provided by the client and will be used by the insurer in the event of a loss. These limits should be carefully reviewed to ensure appropriate property coverage is in place for all buildings and structures. We recommend a professional appraisal every three to four years, or any time major construction or renovation occurs.",
        "Premiums charged by insurance companies are determined by the amount of exposure or risk of loss. These exposures were based on the information provided to our agency during the application process. Additional exposures often arise during the policy period — including new structures, property, or owners/occupants. To ensure an appropriate amount of coverage remains in place, please inform us of any changes in the community. Changes to deductibles and/or limits of liability can affect the total premium.",
        "Blue Lime Insurance Group is paid through commissions issued from the insurance companies we represent when placing your insurance. The amount of commission varies between companies but is generally set by the insurance company and paid as a percentage of the premium you pay. From time to time intermediary wholesale brokers may be used to access certain markets in order to locate the best coverage option, including non-admitted (surplus lines) carriers. When intermediaries are used, our commission or fee may be allocated from a portion of the compensation the intermediary receives. Our agency is also entitled to charge a service fee outside of issued commissions. This fee is not set by law and may be in lieu of or in addition to any commissions.",
        "We strive to deliver exceptional service throughout your policy term. Please do not hesitate to reach out if you have any questions about your coverage.",
    ]
    for para in paragraphs:
        story.append(p(para, "SmallJ"))
        story.append(Spacer(1, 8))
    return story


def _headshot_cell(slug, name, title, phone, email):
    img_path = os.path.join(HEADSHOT_DIR, f"{slug}.png")
    img = Image(img_path, width=0.85 * inch, height=0.85 * inch)

    name_style = ParagraphStyle("n", fontName="Helvetica-Bold", fontSize=10.5,
                                 textColor=BL_NAVY, alignment=TA_CENTER, leading=13)
    title_style = ParagraphStyle("t", fontName="Helvetica", fontSize=9,
                                 textColor=BL_TEXT, alignment=TA_CENTER, leading=11)
    contact_style = ParagraphStyle("c", fontName="Helvetica", fontSize=8,
                                   textColor=BL_MUTED, alignment=TA_CENTER, leading=10)
    contact_bold = ParagraphStyle("cb", fontName="Helvetica-Bold", fontSize=8,
                                  textColor=BL_BLUE_DK, alignment=TA_CENTER, leading=10)
    return [
        img, Spacer(1, 4),
        Paragraph(esc(name), name_style),
        Paragraph(esc(title), title_style),
        Spacer(1, 2),
        Paragraph(esc(phone), contact_style),
        Paragraph(esc(email), contact_bold),
    ]


def _team_page(data):
    am = data["account_manager"]
    am_slug = am.get("slug", "david")  # default fallback
    story = section("Your Blue Lime Team")
    story.append(p("Your community's success is our top priority. Our dedicated team is here to help at every step.", "Body"))
    story.append(Spacer(1, 10))

    # Full team — David's slug is the AM, badge appears above his card.
    team_rows = [
        [("briana",  "Briana Howard",  "Account Manager",              "210-955-6369", "bhoward@bluelimeins.com"),
         ("david",   "David Ritualo",  "Account Manager",              "210-507-0262", "dritualo@bluelimeins.com")],
        [("carol",   "Carol Marquez",  "Account Manager",              "210-951-8704", "cmarquez@bluelimeins.com"),
         ("valerie", "Valerie Cordes", "Claims Specialist",            "210-951-8705", "vcordes@bluelimeins.com")],
        [("steven",  "Steven Melgosa", "Account Manager",              "210-951-8702", "smelgosa@bluelimeins.com"),
         ("daniel",  "Daniel Infante", "Head of Insurance Operations", "210-483-8105", "dinfante@bluelimeins.com")],
        [("susan",   "Susan Finke",    "Account Manager",              "210-955-6372", "sfinke@bluelimeins.com"),
         None],
    ]

    badge_style = ParagraphStyle("badge", fontName="Helvetica-Bold", fontSize=7.5,
                                 textColor=BL_BLUE_DK, alignment=TA_CENTER, leading=10)

    col_w = CONTENT_W / 2
    grid_rows = []
    for row in team_rows:
        cells = []
        for person in row:
            if person is None:
                cells.append("")
            else:
                slug, n, t, ph, em = person
                cell = _headshot_cell(slug, n, t, ph, em)
                if slug == am_slug:
                    cell = [Paragraph("YOUR ACCOUNT MANAGER", badge_style),
                            Spacer(1, 3)] + cell
                cells.append(cell)
        grid_rows.append(cells)

    grid = Table(grid_rows, colWidths=[col_w, col_w])
    grid.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(grid)

    story.append(Spacer(1, 10))
    white_small = ParagraphStyle("ws", fontName="Helvetica-Bold", fontSize=11,
                                 textColor=white, alignment=TA_CENTER, leading=14)
    white_cap = ParagraphStyle("wc", fontName="Helvetica", fontSize=9,
                               textColor=HexColor("#B4D9EC"), alignment=TA_CENTER, leading=12)
    white_link = ParagraphStyle("wl", fontName="Helvetica-Bold", fontSize=10.5,
                                textColor=HexColor("#5DC3EC"), alignment=TA_CENTER, leading=14)
    brand_cells = [[[
        Paragraph("Blue Lime Insurance Group", white_small),
        Paragraph("San Antonio, Texas", white_cap),
        Spacer(1, 4),
        Paragraph("contact@bluelimeins.com  \u00b7  www.bluelimeins.com", white_link),
    ]]]
    brand = Table(brand_cells, colWidths=[CONTENT_W])
    brand.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BL_NAVY),
        ("LEFTPADDING",   (0, 0), (-1, -1), 20),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 20),
        ("TOPPADDING",    (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(brand)
    return story


# =============================================================================
# PUBLIC API
# =============================================================================
def build_proposal(data: dict, output_path: str) -> str:
    """Render a proposal PDF.

    Args:
        data: ProposalData dict produced by `excel_parser.parse_excel()`.
        output_path: Where to write the PDF.

    Returns:
        The output path.
    """
    doc = BaseDocTemplate(
        output_path,
        pagesize=LETTER,
        leftMargin=LEFT, rightMargin=RIGHT,
        topMargin=TOP + 0.25 * inch, bottomMargin=BOTTOM + 0.25 * inch,
        title=f"{data['client']['short_name']} \u2013 Insurance Proposal",
        author="Blue Lime Insurance Group",
        subject="Insurance Proposal",
    )

    cover_frame = Frame(0, 0, PAGE_W, PAGE_H, leftPadding=0, rightPadding=0,
                        topPadding=0, bottomPadding=0, id="cover")
    content_frame = Frame(LEFT, BOTTOM, CONTENT_W, PAGE_H - TOP - BOTTOM - 0.2 * inch,
                          leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                          id="content")

    doc.addPageTemplates([
        PageTemplate(id="Cover",    frames=[cover_frame],   onPage=_make_cover_drawer(data)),
        PageTemplate(id="Interior", frames=[content_frame], onPage=_make_interior_drawer(data)),
    ])

    story = [NextPageTemplate("Interior"), PageBreak()]
    story += _summary_page(data); story.append(PageBreak())
    story += _premium_comparison_page(data); story.append(PageBreak())
    story += _coverage_pages(data); story.append(PageBreak())
    story += _sov_page(data); story.append(PageBreak())
    story += _authorization_page(data); story.append(PageBreak())
    story += _disclosures_page(); story.append(PageBreak())
    story += _team_page(data)

    doc.build(story)
    return output_path
