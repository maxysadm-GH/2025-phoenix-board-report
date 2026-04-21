"""Generate a 3-slide PowerPoint deck from the Phoenix Board Report content.
Produces editorial.pptx and brand.pptx (identical content, different naming).
Board members can copy-paste slides directly into their own decks."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pathlib import Path
import copy

# Paths
OUT = Path(__file__).parent

# Brand palette
PURPLE = RGBColor(0x38, 0x16, 0x5F)
PLUM = RGBColor(0x3D, 0x11, 0x52)
LAVENDER = RGBColor(0xEA, 0xE1, 0xF5)
CHARCOAL = RGBColor(0x24, 0x24, 0x23)
MUTED = RGBColor(0xA5, 0xA5, 0xA5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WARM = RGBColor(0xF6, 0xF6, 0xF4)
IVORY = RGBColor(0xFA, 0xF7, 0xF1)

# Slide dimensions (widescreen 16:9 = 13.333 x 7.5)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _add_text(slide, left, top, width, height, text, font_size=11,
              color=CHARCOAL, bold=False, italic=False, font_name='Calibri',
              alignment=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """Add a text box with a single run."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf.vertical_anchor = anchor
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    return tf


def _add_rich_text(slide, left, top, width, height, parts, font_size=11,
                   alignment=PP_ALIGN.LEFT, line_spacing=1.15):
    """Add text with mixed formatting. parts = [(text, {font attrs}), ...]"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alignment
    if line_spacing:
        p.line_spacing = Pt(font_size * line_spacing)
    for text, attrs in parts:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(attrs.get('size', font_size))
        run.font.color.rgb = attrs.get('color', CHARCOAL)
        run.font.bold = attrs.get('bold', False)
        run.font.italic = attrs.get('italic', False)
        run.font.name = attrs.get('font', 'Calibri')
    return tf


def _add_rect(slide, left, top, width, height, fill_color=None, border_color=None):
    """Add a rectangle shape."""
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def _stat_block(slide, left, top, number, label, num_color=PURPLE, num_size=28):
    """Render a stat: big number + small label underneath."""
    _add_text(slide, left, top, Inches(2), Inches(0.5), number,
              font_size=num_size, color=num_color, font_name='Georgia')
    _add_text(slide, left, top + Inches(0.45), Inches(2), Inches(0.4), label,
              font_size=8, color=MUTED)


def _table_row(slide, left, top, key, value, width=Inches(5.2), val_bold_prefix=None):
    """Single capability-grid row."""
    # Key
    _add_text(slide, left, top, Inches(1.1), Inches(0.26), key.upper(),
              font_size=8, color=MUTED, bold=True)
    # Value
    if val_bold_prefix:
        parts = [(val_bold_prefix, {'bold': True, 'color': PURPLE, 'font': 'Georgia', 'size': 10.5}),
                 (value, {'size': 10.5})]
        _add_rich_text(slide, left + Inches(1.15), top, width - Inches(1.15), Inches(0.26),
                       parts, font_size=10.5)
    else:
        _add_text(slide, left + Inches(1.15), top, width - Inches(1.15), Inches(0.26), value,
                  font_size=10.5)


# ============ SLIDE 1 — THEN / NOW ============
def build_slide1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Background
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = IVORY

    # Eyebrow
    _add_text(slide, Inches(0.6), Inches(0.3), Inches(10), Inches(0.3),
              'VOSGES HAUT-CHOCOLAT  \u00b7  A CONVERSATION ABOUT CAPABILITY',
              font_size=9, color=PLUM, bold=True)

    # Headline
    parts = [
        ('Same walls. Same brand. ', {'size': 32, 'font': 'Georgia', 'color': CHARCOAL}),
        ('A different company underneath.', {'size': 32, 'font': 'Georgia', 'color': PLUM, 'italic': True}),
    ]
    _add_rich_text(slide, Inches(0.6), Inches(0.55), Inches(11), Inches(0.6), parts, font_size=32)

    # Sub
    _add_text(slide, Inches(0.6), Inches(1.1), Inches(10), Inches(0.3),
              'What the prior era built. What the new era built. And what we kept on purpose.',
              font_size=12, color=CHARCOAL)

    # Rule line
    _add_rect(slide, Inches(0.6), Inches(1.45), Inches(1), Pt(1), fill_color=PURPLE)

    # ---- THEN column ----
    col_l = Inches(0.6)
    col_t = Inches(1.65)
    _add_rect(slide, col_l, col_t, Inches(5.6), Inches(5.3), border_color=MUTED)

    _add_text(slide, col_l + Inches(0.2), col_t + Inches(0.1), Inches(3), Inches(0.2),
              'THEN \u00b7 PRIOR OWNERSHIP', font_size=9, color=MUTED, bold=True)

    parts = [('A brand ahead of its ', {'font': 'Georgia', 'size': 18, 'color': CHARCOAL}),
             ('engine.', {'font': 'Georgia', 'size': 18, 'color': CHARCOAL, 'italic': True})]
    _add_rich_text(slide, col_l + Inches(0.2), col_t + Inches(0.35), Inches(5), Inches(0.35), parts, 18)

    _add_text(slide, col_l + Inches(0.2), col_t + Inches(0.65), Inches(5), Inches(0.35),
              'World-class creative and innovation no one else could touch \u2014\nbut the operation was never built to carry the brand at scale.',
              font_size=10.5, color=RGBColor(0x5A, 0x4E, 0x65), italic=True)

    rows_then = [
        ('Brand \u00b7 R&D', 'Iconic. Truffles, bars, bacon bar, PB-PB cup.'),
        ('Systems', '15-year-old CMS \u00b7 Nutracoster recipe tool, 20+ yrs \u00b7 islands of data.'),
        ('Floor tools', 'Paper pick lists. Paper counts. Rekeying everywhere.'),
        ('QA', 'Manual inspection. Errors caught downstream \u2014 often after shipment.'),
        ('Capacity', '~1,675 orders/day peak \u00b7 manual-heavy \u00b7 single carrier.'),
        ('Cost / ship', '$34.47 avg \u00b7 legacy pricing \u00b7 no carrier optionality.'),
        ('Service', 'Customer replies averaged 24\u201348 hours.'),
    ]
    ry = col_t + Inches(1.1)
    for k, v in rows_then:
        _table_row(slide, col_l + Inches(0.2), ry, k, v, width=Inches(5.2))
        ry += Inches(0.26)

    # PB-PB quote
    _add_text(slide, col_l + Inches(0.2), ry + Inches(0.1), Inches(5.2), Inches(0.4),
              '\u201cThe PB-PB cup never left a kitchen recipe.\u201d The creative was there. The chassis to ship it wasn\u2019t.',
              font_size=10, color=RGBColor(0x5A, 0x4E, 0x65), italic=True)

    # Bottom stats THEN
    stats_y = col_t + Inches(4.15)
    _add_rect(slide, col_l + Inches(0.2), stats_y, Inches(5.2), Pt(1), fill_color=MUTED)
    _stat_block(slide, col_l + Inches(0.3), stats_y + Inches(0.12), '1,675', 'DAILY CAPACITY\nPEAK (2023)', num_color=MUTED)
    _stat_block(slide, col_l + Inches(2.1), stats_y + Inches(0.12), '$4.01', 'DIRECT FULFILLMENT\nLABOR COST', num_color=MUTED)
    _stat_block(slide, col_l + Inches(3.9), stats_y + Inches(0.12), '$34.47', 'COST PER\nSHIPMENT (AVG)', num_color=MUTED)

    # ---- NOW column ----
    col_r = Inches(6.6)
    _add_rect(slide, col_r, col_t, Inches(6.1), Inches(5.3), border_color=PURPLE)

    _add_text(slide, col_r + Inches(0.2), col_t + Inches(0.1), Inches(3), Inches(0.2),
              'NOW \u00b7 CURRENT OWNERSHIP', font_size=9, color=PURPLE, bold=True)

    parts = [('The brand, ', {'font': 'Georgia', 'size': 18, 'color': PURPLE}),
             ('plus the engine it deserves.', {'font': 'Georgia', 'size': 18, 'color': PLUM, 'italic': True})]
    _add_rich_text(slide, col_r + Inches(0.2), col_t + Inches(0.35), Inches(5.5), Inches(0.35), parts, 18)

    _add_text(slide, col_r + Inches(0.2), col_t + Inches(0.65), Inches(5.5), Inches(0.35),
              'Same building. Same team DNA. A modernized operational spine \u2014\nconnected systems, instrumented floors, simpler choices everywhere.',
              font_size=10.5, color=CHARCOAL, italic=True)

    rows_now = [
        ('Brand \u00b7 R&D', 'Kept intact. Innovation calendar protected \u2014 the reason you buy Vosges.', 'Kept intact. '),
        ('Systems', 'retires the 15-yr CMS \u00b7 Flavor Studio AI replaces Nutracoster.', 'Current-gen MRP '),
        ('Floor tools', 'on every station. Paper cut 50%.', 'Tablets & scanners '),
        ('QA', '(Gorgias FY25) \u00b7 checkweigher + scan-verified.', '98.2% pick accuracy '),
        ('Capacity', '', '~3,120/day sustainable.'),
        ('Cost / ship', '(FY25 DTC) \u2014 carrier simplification + negotiation.', '$24.07 avg (FY24 DTC) | $18.38 avg '),
        ('Service', 'with Gorgias + 24/7 virtual agents.', 'Under 1-hour first-response '),
    ]
    ry = col_t + Inches(1.1)
    for k, v, bp in rows_now:
        _table_row(slide, col_r + Inches(0.2), ry, k, v, width=Inches(5.6), val_bold_prefix=bp)
        ry += Inches(0.26)

    # Keep callout
    _add_rect(slide, col_r + Inches(0.2), ry + Inches(0.08), Inches(5.6), Inches(0.4),
              fill_color=LAVENDER)
    _add_rect(slide, col_r + Inches(0.2), ry + Inches(0.08), Pt(3), Inches(0.4),
              fill_color=PURPLE)
    parts = [('What we kept on purpose: ', {'bold': True, 'color': PURPLE, 'size': 10}),
             ('the creative, the flavor R&D, the brand voice. None of the upgrade touched the product.', {'size': 10})]
    _add_rich_text(slide, col_r + Inches(0.32), ry + Inches(0.12), Inches(5.4), Inches(0.35), parts, 10)

    # Bottom stats NOW
    stats_y = col_t + Inches(4.15)
    _add_rect(slide, col_r + Inches(0.2), stats_y, Inches(5.6), Pt(1), fill_color=PURPLE)
    _stat_block(slide, col_r + Inches(0.3), stats_y + Inches(0.12), '3,120', 'CAPACITY\n(SAME SPACE)')
    _stat_block(slide, col_r + Inches(2.2), stats_y + Inches(0.12), '$1.38', 'DIRECT FULFILLMENT\nLABOR COST')
    _stat_block(slide, col_r + Inches(4.1), stats_y + Inches(0.12), '$18.38', 'COST PER\nSHIPMENT (AVG)')


# ============ SLIDE 2 — THE DIVIDEND ============
def build_slide2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = IVORY

    # Eyebrow
    _add_text(slide, Inches(0.6), Inches(0.3), Inches(10), Inches(0.3),
              'THE DIVIDEND OF MODERNIZATION \u00b7 AND WHAT HELD DURING THE REBUILD',
              font_size=9, color=PLUM, bold=True)

    # Headline
    parts = [
        ('A footprint sized for December. ', {'size': 32, 'font': 'Georgia', 'color': CHARCOAL}),
        ('Running in April.', {'size': 32, 'font': 'Georgia', 'color': CHARCOAL, 'italic': True}),
    ]
    _add_rich_text(slide, Inches(0.6), Inches(0.55), Inches(11), Inches(0.6), parts, font_size=32)

    _add_text(slide, Inches(0.6), Inches(1.1), Inches(11), Inches(0.3),
              'Fewer systems, better ones. More throughput, same walls. Faster answers at every door \u2014 even as the line was being rebuilt underneath.',
              font_size=12, color=CHARCOAL, italic=True, font_name='Georgia')

    # Rule
    _add_rect(slide, Inches(0.6), Inches(1.45), Inches(1), Pt(1), fill_color=PURPLE)

    # KPI row
    kpis = [
        ('~3,120', 'SUSTAINABLE DAILY\nSHIPPING CAPACITY'),
        ('5,268/d', 'SINGLE-DAY PEAK\nPROVEN \u00b7 DEC 2025'),
        ('\u221250%', 'TECHNOLOGY COST\nYEAR OVER YEAR'),
        ('\u221275%', 'OLD SOFTWARE RISK\nELIMINATED'),
    ]
    kpi_x = Inches(0.6)
    for num, lbl in kpis:
        _add_rect(slide, kpi_x, Inches(1.65), Inches(2.9), Pt(2), fill_color=PURPLE)
        _add_text(slide, kpi_x, Inches(1.75), Inches(2.9), Inches(0.5), num,
                  font_size=28, color=PURPLE, font_name='Georgia')
        _add_text(slide, kpi_x, Inches(2.2), Inches(2.9), Inches(0.4), lbl,
                  font_size=8, color=MUTED, bold=True)
        kpi_x += Inches(3.1)

    # Modernization list
    _add_text(slide, Inches(0.6), Inches(2.85), Inches(5), Inches(0.25),
              'WHAT THE MODERNIZATION DELIVERED', font_size=9, color=PLUM, bold=True)

    items = [
        ('MRP', 'Retired the 15-year CMS; current-gen MRP in its place.', '1 stack'),
        ('RECIPE', 'Flavor Studio replaces Nutracoster (20+ yrs) \u2014 AI at the bench.', 'Next-gen'),
        ('FLOOR', 'Tablets & scanners at every station. Paper \u221250%.', 'Every station'),
        ('SHIPPING', 'Carrier simplification. Ice & unboxing standardized.', '\u221247% $/ship'),
        ('SERVICE', 'Gorgias + 24/7 agents \u2014 answered under the hour.', '24\u201348h \u2192 1h'),
        ('VISIBILITY', 'Live dashboards. 24/7 automated audits. Data-driven decisions.', 'Always-on'),
        ('INTEGRATION', 'Every system talks to every other system.', 'One graph'),
    ]
    iy = Inches(3.15)
    for k, v, s in items:
        _add_text(slide, Inches(0.6), iy, Inches(1), Inches(0.22), k,
                  font_size=8, color=MUTED, bold=True)
        _add_text(slide, Inches(1.7), iy, Inches(5), Inches(0.22), v,
                  font_size=10.5)
        _add_text(slide, Inches(7.0), iy, Inches(1.2), Inches(0.22), s,
                  font_size=10.5, color=PURPLE, italic=True, font_name='Georgia',
                  alignment=PP_ALIGN.RIGHT)
        iy += Inches(0.28)

    # Pick accuracy hero
    _add_text(slide, Inches(0.6), Inches(5.3), Inches(5), Inches(0.2),
              'PICK ACCURACY \u00b7 GORGIAS FY25', font_size=9, color=MUTED, bold=True)

    parts = [('98.2', {'size': 48, 'color': PURPLE, 'font': 'Georgia'}),
             ('%', {'size': 24, 'color': MUTED, 'font': 'Georgia'})]
    _add_rich_text(slide, Inches(0.6), Inches(5.5), Inches(3), Inches(0.7), parts, 48)

    _add_text(slide, Inches(2.8), Inches(5.65), Inches(4), Inches(0.3),
              'OF SHIPPED ORDERS \u2014 NO CUSTOMER-REPORTED ISSUE',
              font_size=9, color=MUTED, bold=True)

    # Callout
    _add_rect(slide, Inches(0.6), Inches(6.4), Pt(2), Inches(0.6), fill_color=PURPLE)
    parts = [('Held the line during the rebuild. ', {'bold': True, 'size': 10.5}),
             ('Q4 2025 \u2014 the quarter we re-sequenced the pick line \u2014 shipped record holiday volume (Dec 2025: 61,503 shipments, peak 5,268/day) without degrading the accuracy floor.',
              {'size': 10.5, 'italic': True, 'font': 'Georgia'})]
    _add_rich_text(slide, Inches(0.75), Inches(6.42), Inches(11.5), Inches(0.6), parts, 10.5)


# ============ SLIDE 3 — WHAT LIES AHEAD ============
def build_slide3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = IVORY

    # Eyebrow
    _add_text(slide, Inches(0.6), Inches(0.3), Inches(10), Inches(0.3),
              'WHAT LIES AHEAD \u00b7 PHOENIX PHASE TWO \u00b7 AND BEYOND',
              font_size=9, color=PLUM, bold=True)

    parts = [
        ('Faster. Simpler. ', {'size': 32, 'font': 'Georgia', 'color': CHARCOAL}),
        ('And then \u2014 portable.', {'size': 32, 'font': 'Georgia', 'color': CHARCOAL, 'italic': True}),
    ]
    _add_rich_text(slide, Inches(0.6), Inches(0.55), Inches(11), Inches(0.6), parts, font_size=32)

    _add_text(slide, Inches(0.6), Inches(1.1), Inches(11), Inches(0.3),
              'An $80,000 investment that pays itself back in twelve weeks \u2014 and teaches us how to wire the next warehouse in a week, not a year.',
              font_size=12, color=CHARCOAL, italic=True, font_name='Georgia')

    _add_rect(slide, Inches(0.6), Inches(1.45), Inches(1), Pt(1), fill_color=PURPLE)

    # KPI row (6 tiles)
    kpis3 = [
        ('$80K', 'PHASE 2\nINVEST'),
        ('$7K/wk', 'ONGOING\nSAVINGS'),
        ('12 wks', 'PAYBACK\nPERIOD'),
        ('$364K', 'YEAR-1\nANNUALIZED'),
        ('$1.1M', 'THREE-YEAR\nCUMULATIVE'),
        ('25\u201335%', 'OVERHEAD CUT\nON THE LINE'),
    ]
    kx = Inches(0.6)
    for num, lbl in kpis3:
        _add_rect(slide, kx, Inches(1.65), Inches(1.85), Pt(2), fill_color=PURPLE)
        _add_text(slide, kx, Inches(1.75), Inches(1.85), Inches(0.45), num,
                  font_size=22, color=PURPLE, font_name='Georgia')
        _add_text(slide, kx, Inches(2.15), Inches(1.85), Inches(0.35), lbl,
                  font_size=8, color=MUTED, bold=True)
        kx += Inches(2.05)

    # --- Phase Two column ---
    _add_text(slide, Inches(0.6), Inches(2.7), Inches(4), Inches(0.2),
              'PHASE TWO \u00b7 ON THE LINE', font_size=9, color=PLUM, bold=True)

    parts = [('The things that still look ', {'font': 'Georgia', 'size': 16, 'color': CHARCOAL}),
             ('manual.', {'font': 'Georgia', 'size': 16, 'color': PURPLE, 'italic': True})]
    _add_rich_text(slide, Inches(0.6), Inches(2.95), Inches(4), Inches(0.3), parts, 16)

    phase2 = [
        ('Gift orders', 'Dedicated box-maker station at the line\u2019s head. Gift-note printed on scan \u2014 never missed again.'),
        ('QA stop', 'QA & Organizer merge into one role. One person owns the box before it moves.'),
        ('Sealing', 'Auto-sealer replaces manual tape & fold. Seasonal rental \u2014 paid back in one peak week.'),
        ('Labeling', 'Three-tap ShipStation Mobile retired. Auto peel-and-apply labeler takes over.'),
        ('Boxes', 'Sizes re-aligned to FedEx One Rate tiers. New SKUs require box-team approval.'),
        ('Bedding', 'Pre-cut, known-weight. The checkweigher stops guessing; shipping math becomes exact.'),
    ]
    py = Inches(3.35)
    for tag, desc in phase2:
        _add_text(slide, Inches(0.6), py, Inches(0.9), Inches(0.25), tag,
                  font_size=10.5, color=PURPLE, italic=True, font_name='Georgia')
        _add_text(slide, Inches(1.6), py, Inches(3), Inches(0.25), desc, font_size=10.5)
        py += Inches(0.32)

    # --- The math column ---
    math_l = Inches(5.0)
    _add_text(slide, math_l, Inches(2.7), Inches(4), Inches(0.2),
              'THE MATH \u00b7 CUMULATIVE $ THROUGH THE DOOR', font_size=9, color=PLUM, bold=True)

    parts = [('$80K out. ', {'font': 'Georgia', 'size': 16, 'color': CHARCOAL}),
             ('$1.1M back.', {'font': 'Georgia', 'size': 16, 'color': PURPLE, 'italic': True})]
    _add_rich_text(slide, math_l, Inches(2.95), Inches(4), Inches(0.3), parts, 16)

    # Math header
    mh = [('Wk 12', 'BREAK-EVEN'), ('Wk 52', '+$284K NET'), ('Yr 3', '+$1M NET')]
    mx = math_l
    for num, lbl in mh:
        _add_rect(slide, mx, Inches(3.35), Inches(1.2), Pt(1), fill_color=CHARCOAL)
        _add_text(slide, mx, Inches(3.42), Inches(1.2), Inches(0.35), num,
                  font_size=20, color=PURPLE, font_name='Georgia')
        _add_text(slide, mx, Inches(3.72), Inches(1.2), Inches(0.2), lbl,
                  font_size=7.5, color=MUTED, bold=True)
        mx += Inches(1.35)

    # Payback narrative (since Chart.js can't render in PPTX, describe the curve)
    _add_rect(slide, math_l, Inches(4.1), Pt(2), Inches(0.6), fill_color=PURPLE)
    parts = [('Read it this way \u2014 ', {'bold': True, 'size': 10.5}),
             ('every line of cost stays flat; every line of savings stacks. The curves cross at week twelve. Everything after is structural margin on the P&L.',
              {'size': 10.5, 'italic': True, 'font': 'Georgia'})]
    _add_rich_text(slide, math_l + Inches(0.12), Inches(4.12), Inches(3.8), Inches(0.6), parts, 10.5)

    # --- Vision column ---
    vis_l = Inches(9.2)
    _add_rect(slide, vis_l, Inches(2.7), Inches(3.9), Inches(4.4), fill_color=PLUM)

    _add_text(slide, vis_l + Inches(0.2), Inches(2.8), Inches(3.5), Inches(0.2),
              'AND THEN \u2014 THE REAL UNLOCK', font_size=9, color=LAVENDER, bold=True)

    parts = [('A template that ', {'font': 'Georgia', 'size': 16, 'color': WHITE}),
             ('travels.', {'font': 'Georgia', 'size': 16, 'color': LAVENDER, 'italic': True})]
    _add_rich_text(slide, vis_l + Inches(0.2), Inches(3.05), Inches(3.5), Inches(0.3), parts, 16)

    # Node diagram (text-based)
    nodes_txt = 'Chicago (Hub \u00b7 Full SKU)  \u2500\u2500\u2500  Node II (Subset)  \u2500\u2500\u2500  Node III (Subset)'
    _add_text(slide, vis_l + Inches(0.2), Inches(3.5), Inches(3.5), Inches(0.3),
              nodes_txt, font_size=10, color=WHITE, font_name='Georgia')

    # Vision text
    vision_lines = [
        'Every simplification on this line is a step toward a standard \u2014 something we can wire somewhere else.',
        'A portable warehouse template. Light up a second location to fulfill a subset of products \u2014 closer to the customer, faster to deliver, cheaper to ship.',
        'Not a bigger building. A smarter network.',
    ]
    vy = Inches(4.0)
    for line in vision_lines:
        _add_text(slide, vis_l + Inches(0.2), vy, Inches(3.5), Inches(0.5),
                  line, font_size=11, color=WHITE, font_name='Georgia')
        vy += Inches(0.55)

    # Close
    _add_rect(slide, vis_l + Inches(0.2), Inches(5.7), Inches(3.5), Pt(1), fill_color=LAVENDER)
    parts = [('Simplify first. Replicate second. ', {'font': 'Georgia', 'size': 11, 'color': LAVENDER, 'italic': True, 'bold': True}),
             ('The next warehouse is a configuration, not a construction project.', {'font': 'Georgia', 'size': 11, 'color': WHITE, 'italic': True})]
    _add_rich_text(slide, vis_l + Inches(0.2), Inches(5.8), Inches(3.5), Inches(0.5), parts, 11)


def build(name):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    build_slide1(prs)
    build_slide2(prs)
    build_slide3(prs)
    out = OUT / f"{name}.pptx"
    prs.save(str(out))
    print(f"[{name}] -> {out}")


if __name__ == '__main__':
    build('editorial')
    build('brand')
    print("Done.")
