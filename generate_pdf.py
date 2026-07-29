"""Generate the DQE product overview PDF using Duetto brand guidelines."""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable

# ── Brand colours ────────────────────────────────────────────────────────────
GREEN   = colors.HexColor('#C4FF45')
DARK    = colors.HexColor('#0E2124')
WHITE   = colors.white
LIGHT   = colors.HexColor('#F4F7F4')
MUTED   = colors.HexColor('#6B8080')
BORDER  = colors.HexColor('#D6E4DC')
GREY_BG = colors.HexColor('#F0F4F0')
GREY_BD = colors.HexColor('#CBD5CB')
RED     = colors.HexColor('#DC2626')
AMBER   = colors.HexColor('#D97706')
TEAL    = colors.HexColor('#0D9488')
MON_RED = colors.HexColor('#F04E5E')
MON_YEL = colors.HexColor('#FFCB00')
MON_GRN = colors.HexColor('#00CA72')

W, H = A4
MARGIN = 18 * mm


# ── Header / footer callback ─────────────────────────────────────────────────
def draw_chrome(canvas, doc):
    canvas.saveState()

    # Header
    canvas.setFillColor(DARK)
    canvas.rect(0, H - 14*mm, W, 14*mm, fill=1, stroke=0)
    canvas.setFillColor(WHITE)
    canvas.setFont('Helvetica-Bold', 13)
    canvas.drawString(MARGIN, H - 9.5*mm, 'Duetto')
    canvas.setFillColor(GREEN)
    canvas.circle(MARGIN + 39*mm, H - 7*mm, 2.2*mm, fill=1, stroke=0)
    canvas.setFillColor(GREEN)
    canvas.roundRect(MARGIN + 43*mm, H - 11*mm, 14*mm, 7*mm, 1.5*mm, fill=1, stroke=0)
    canvas.setFillColor(DARK)
    canvas.setFont('Helvetica-Bold', 8)
    canvas.drawCentredString(MARGIN + 50*mm, H - 7.2*mm, 'DQE')

    # Footer
    canvas.setFillColor(DARK)
    canvas.rect(0, 0, W, 8*mm, fill=1, stroke=0)
    canvas.setFillColor(MUTED)
    canvas.setFont('Helvetica', 7)
    canvas.drawString(MARGIN, 2.8*mm, 'Duetto Data Quality Engine  ·  Internal Use Only')
    canvas.setFillColor(GREEN)
    canvas.drawRightString(W - MARGIN, 2.8*mm, 'duettocloud.com')

    canvas.restoreState()


# ── Styles ───────────────────────────────────────────────────────────────────
def make_styles():
    s = {}
    s['h2'] = ParagraphStyle('h2',
        fontName='Helvetica-Bold', fontSize=13, leading=17,
        textColor=DARK, spaceBefore=10, spaceAfter=4)
    s['h3'] = ParagraphStyle('h3',
        fontName='Helvetica-Bold', fontSize=10, leading=14,
        textColor=DARK, spaceBefore=6, spaceAfter=2)
    s['body'] = ParagraphStyle('body',
        fontName='Helvetica', fontSize=9.5, leading=15,
        textColor=colors.HexColor('#1A2E2E'), spaceAfter=5)
    s['bullet'] = ParagraphStyle('bullet',
        fontName='Helvetica', fontSize=9.5, leading=14,
        textColor=colors.HexColor('#1A2E2E'),
        leftIndent=14, firstLineIndent=-10, spaceAfter=3)
    s['caption'] = ParagraphStyle('caption',
        fontName='Helvetica-Oblique', fontSize=8, leading=11,
        textColor=MUTED, spaceAfter=4)
    s['img_caption'] = ParagraphStyle('img_caption',
        fontName='Helvetica-Oblique', fontSize=8, leading=11,
        textColor=MUTED, spaceAfter=6, alignment=TA_CENTER)
    return s


def bullet(text, s):
    return Paragraph(f'<bullet>&bull;</bullet> {text}', s['bullet'])


def section_rule(story):
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width='100%', thickness=0.8, color=BORDER))
    story.append(Spacer(1, 3*mm))


def _cell(text, bold=False, color=DARK, size=8.5):
    style = ParagraphStyle('cell',
        fontName='Helvetica-Bold' if bold else 'Helvetica',
        fontSize=size, leading=size * 1.35,
        textColor=color, spaceBefore=0, spaceAfter=0)
    return Paragraph(str(text), style)


def green_box(rows, col_widths, story):
    # Convert all cells to Paragraph so text wraps within column width
    para_rows = []
    for ri, row in enumerate(rows):
        para_row = []
        for cell in row:
            if ri == 0:
                para_row.append(_cell(cell, bold=True, color=GREEN, size=8.5))
            else:
                para_row.append(_cell(cell, bold=False, color=DARK, size=8.5))
        para_rows.append(para_row)

    t = Table(para_rows, colWidths=col_widths, repeatRows=1, splitByRow=True)
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, 0), DARK),
        ('TOPPADDING',    (0, 0), (-1, 0), 5),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('LEFTPADDING',   (0, 0), (-1, -1), 8),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 8),
        ('BACKGROUND',    (0, 1), (-1, -1), WHITE),
        ('ROWBACKGROUNDS',(0, 1), (-1, -1), [WHITE, LIGHT]),
        ('TOPPADDING',    (0, 1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('GRID',          (0, 0), (-1, -1), 0.4, BORDER),
        ('BOX',           (0, 0), (-1, -1), 0.8, DARK),
    ]))
    story.append(t)


# ── Cover flowable ───────────────────────────────────────────────────────────
class CoverFlowable(Flowable):
    def __init__(self, width, height):
        super().__init__()
        self.width  = width
        self.height = height

    def draw(self):
        c = self.canv
        c.saveState()

        # Dark panel
        c.setFillColor(DARK)
        c.roundRect(0, 0, self.width, self.height, 6, fill=1, stroke=0)

        # Green accent left stripe
        c.setFillColor(GREEN)
        c.rect(0, 0, 5, self.height, fill=1, stroke=0)

        # DQE badge
        c.setFillColor(GREEN)
        c.roundRect(16, self.height - 38, 52, 22, 4, fill=1, stroke=0)
        c.setFillColor(DARK)
        c.setFont('Helvetica-Bold', 11)
        c.drawCentredString(42, self.height - 29, 'DQE')

        # Product name
        c.setFillColor(WHITE)
        c.setFont('Helvetica-Bold', 26)
        c.drawString(16, self.height - 72, 'Data Quality Engine')

        # Tagline
        c.setFillColor(GREEN)
        c.setFont('Helvetica', 11)
        c.drawString(16, self.height - 92, 'Automated PMS–Duetto discrepancy detection and root-cause analysis')

        # Divider line
        c.setStrokeColor(colors.HexColor('#1E3A3F'))
        c.setLineWidth(0.5)
        c.line(16, self.height - 104, self.width - 16, self.height - 104)

        # Meta row
        c.setFillColor(MUTED)
        c.setFont('Helvetica', 8.5)
        c.drawString(16, self.height - 120, 'For DVA-supported integrations  ·  Internal Use — Duetto Employees')
        c.drawRightString(self.width - 16, self.height - 120, 'Version 1.0  ·  June 2026')

        c.restoreState()


# ── Monday modal mockup ───────────────────────────────────────────────────────
class MondayModalMockup(Flowable):
    """Draws a stylised representation of the Submit to Monday modal."""
    W_MOD = 130*mm
    H_MOD = 148*mm

    def __init__(self, page_width):
        super().__init__()
        self.width  = page_width
        self.height = self.H_MOD + 6*mm

    def _field(self, c, x, y, w, h, text, placeholder=False, filled=False):
        c.setFillColor(GREY_BG if filled else WHITE)
        c.roundRect(x, y, w, h, 2, fill=1, stroke=1)
        c.setStrokeColor(GREY_BD)
        c.roundRect(x, y, w, h, 2, fill=0, stroke=1)
        c.setFillColor(MUTED if placeholder else DARK)
        c.setFont('Helvetica', 7)
        c.drawString(x + 5, y + h/2 - 3.5, text)

    def _label(self, c, x, y, text):
        c.setFillColor(DARK)
        c.setFont('Helvetica-Bold', 6.5)
        c.drawString(x, y, text)

    def draw(self):
        c = self.canv
        c.saveState()

        # Centre the modal
        mx = (self.width - self.W_MOD) / 2
        my = 3*mm

        # Modal shadow
        c.setFillColor(colors.HexColor('#00000018'))
        c.roundRect(mx + 1.5, my - 1.5, self.W_MOD, self.H_MOD, 5, fill=1, stroke=0)

        # Modal background
        c.setFillColor(WHITE)
        c.setStrokeColor(GREY_BD)
        c.setLineWidth(0.5)
        c.roundRect(mx, my, self.W_MOD, self.H_MOD, 5, fill=1, stroke=1)

        fw = self.W_MOD - 12*mm  # field width
        fx = mx + 6*mm           # field x

        # Title bar
        c.setFillColor(DARK)
        c.setFont('Helvetica-Bold', 10)
        c.drawString(fx, my + self.H_MOD - 10*mm, 'Submit to Monday.com')
        c.setFillColor(MUTED)
        c.setFont('Helvetica', 7)
        c.drawString(fx, my + self.H_MOD - 14*mm,
                     'Add hotel details and feedback — results and Excel report attached automatically.')

        # Thin divider
        c.setStrokeColor(BORDER)
        c.setLineWidth(0.4)
        c.line(fx, my + self.H_MOD - 16*mm, fx + fw, my + self.H_MOD - 16*mm)

        # Fields (top → bottom)
        fields = [
            ('SUBMITTED BY',              'Tyler Linton',               False, True),
            ('HOTEL NAME',                'Sixty DC',                   False, False),
            ('HOTEL ID',                  'ho667099',                   False, False),
            ('STAY DATE RANGE (AUTO-FILLED)', '2025-03-20 – 2026-09-08', True,  True),
            ('ROOMS ACCURACY % (AUTO-FILLED)', '79.0%',                 True,  True),
            ('REVENUE ACCURACY % (AUTO-FILLED)', '21.1%',               True,  True),
            ('FAILING DATES (AUTO-FILLED)', '2025-03-20, 2025-05-27, 2025-05-30…', True, True),
        ]

        fh = 8*mm
        gap = 2.5*mm
        ty = my + self.H_MOD - 19*mm

        for label, value, placeholder, filled in fields:
            ty -= (gap + fh)
            self._label(c, fx, ty + fh + 1, label)
            self._field(c, fx, ty, fw, fh, value, placeholder=placeholder, filled=filled)

        # Feedback textarea
        ty -= (gap + 14*mm + 3)
        self._label(c, fx, ty + 14*mm + 5, 'FEEDBACK / NOTES')
        c.setFillColor(WHITE)
        c.setStrokeColor(GREY_BD)
        c.setLineWidth(0.5)
        c.roundRect(fx, ty, fw, 14*mm, 2, fill=1, stroke=1)
        c.setFillColor(MUTED)
        c.setFont('Helvetica-Oblique', 7)
        c.drawString(fx + 5, ty + 14*mm - 8, 'Any observations, follow-up actions, or notes…')

        # Buttons
        btn_y = my + 5*mm
        # Cancel
        c.setFillColor(WHITE)
        c.setStrokeColor(GREY_BD)
        c.setLineWidth(0.8)
        c.roundRect(fx + fw - 52*mm, btn_y, 24*mm, 8*mm, 2, fill=1, stroke=1)
        c.setFillColor(DARK)
        c.setFont('Helvetica', 8)
        c.drawCentredString(fx + fw - 40*mm, btn_y + 3, 'Cancel')

        # Submit
        c.setFillColor(DARK)
        c.roundRect(fx + fw - 26*mm, btn_y, 26*mm, 8*mm, 2, fill=1, stroke=0)
        c.setFillColor(GREEN)
        c.setFont('Helvetica-Bold', 8)
        c.drawCentredString(fx + fw - 13*mm, btn_y + 3, 'Submit')

        c.restoreState()


# ── Monday board mockup ───────────────────────────────────────────────────────
class MondayBoardMockup(Flowable):
    """Draws a stylised representation of the Monday.com DQE board."""

    def __init__(self, page_width, height=60*mm):
        super().__init__()
        self.width  = page_width
        self.height = height

    def draw(self):
        c = self.canv
        c.saveState()
        pw = self.width
        ph = self.height

        # Outer card
        c.setFillColor(WHITE)
        c.setStrokeColor(GREY_BD)
        c.setLineWidth(0.5)
        c.roundRect(0, 0, pw, ph, 4, fill=1, stroke=1)

        # Top bar (Monday nav simulation)
        c.setFillColor(colors.HexColor('#F5F6F8'))
        c.rect(0, ph - 9*mm, pw, 9*mm, fill=1, stroke=0)
        c.setFillColor(DARK)
        c.setFont('Helvetica-Bold', 9)
        c.drawString(6*mm, ph - 6*mm, 'DQE Feedback')
        c.setFillColor(MUTED)
        c.setFont('Helvetica', 7.5)
        c.drawString(6*mm, ph - 9.5*mm + 1, 'Main table')

        # Column header row
        col_headers = ['Project', 'Files Used', 'Excel Report', 'Feedback', 'Submitted By', 'DVA File', 'Bookings File']
        col_xs      = [0, 52, 80, 100, 130, 155, 175]  # in mm, relative
        col_xs_pts  = [x*mm for x in col_xs]
        hdr_y = ph - 18*mm
        c.setFillColor(colors.HexColor('#F0F1F3'))
        c.rect(0, hdr_y, pw, 8*mm, fill=1, stroke=0)
        c.setStrokeColor(GREY_BD)
        c.setLineWidth(0.3)
        c.line(0, hdr_y, pw, hdr_y)
        c.line(0, hdr_y + 8*mm, pw, hdr_y + 8*mm)

        c.setFillColor(MUTED)
        c.setFont('Helvetica-Bold', 6.5)
        for i, hdr in enumerate(col_headers):
            cx = col_xs_pts[i] + 3
            if cx + 3 > pw:
                break
            c.drawString(cx, hdr_y + 2.5*mm, hdr)

        # Data rows
        rows = [
            ('Chesterfield Hotel & Suites — 2026-0…', 'DVA: Chesterfield…', '📗', 'Testing 123', '👤', '📗', '📘'),
            ('Sixty DC — 2026-06-30',                  'DVA: Sixty+DC_h…',   '📗', 'Testing 1234','👤', '📗', '📘'),
        ]
        row_h = 8*mm
        for ri, row in enumerate(rows):
            ry = hdr_y - (ri + 1) * row_h
            # Alternating bg
            bg = WHITE if ri % 2 == 0 else colors.HexColor('#FAFBFA')
            c.setFillColor(bg)
            c.rect(0, ry, pw, row_h, fill=1, stroke=0)
            c.setStrokeColor(GREY_BD)
            c.setLineWidth(0.3)
            c.line(0, ry, pw, ry)

            # Blue project name accent
            c.setFillColor(colors.HexColor('#4A90D9'))
            c.rect(0, ry, 2.5, row_h, fill=1, stroke=0)

            c.setFillColor(DARK)
            c.setFont('Helvetica', 7)
            for ci, cell in enumerate(row):
                cx = col_xs_pts[ci] + 4
                if cx + 3 > pw:
                    break
                c.drawString(cx, ry + 2.5*mm, str(cell))

        # Bottom: + Add project row
        add_y = hdr_y - len(rows) * row_h
        c.setFillColor(WHITE)
        c.rect(0, add_y, pw, row_h, fill=1, stroke=0)
        c.setFillColor(MUTED)
        c.setFont('Helvetica', 7)
        c.drawString(6*mm, add_y + 2.5*mm, '+ Add project')

        c.restoreState()


# ── Main build ───────────────────────────────────────────────────────────────
def build():
    out = '/Users/tylerlinton/DQE/DQE_Product_Overview.pdf'
    doc = SimpleDocTemplate(
        out,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=18*mm, bottomMargin=12*mm,
        title='Duetto DQE — Product Overview',
        author='Duetto',
    )

    s = make_styles()
    usable_w = W - 2 * MARGIN
    story = []

    # ── Cover ────────────────────────────────────────────────────────────────
    story.append(CoverFlowable(usable_w, 52*mm))
    story.append(Spacer(1, 5*mm))

    # ── 1. What is the DQE? ──────────────────────────────────────────────────
    story.append(Paragraph('What is the Data Quality Engine?', s['h2']))
    story.append(Paragraph(
        'The <b>Data Quality Engine (DQE)</b> is a Duetto-built web application that automates '
        'the identification and root-cause analysis of discrepancies between a hotel\'s '
        '<b>Property Management System (PMS)</b> and the data held inside <b>Duetto</b>. '
        'It is designed for use by <b>Duetto Quality Analysts</b> with use cases for '
        '<b>Duetto Support Agents</b>.', s['body']))
    story.append(Paragraph(
        'Rather than manually comparing exports line by line, the DQE ingests multiple source '
        'files, applies a library of classification rules, computes accuracy metrics, and '
        'generates a structured report — in seconds.', s['body']))

    section_rule(story)

    # ── 2. DVA-Supported Integrations ────────────────────────────────────────
    story.append(Paragraph('DVA-Supported Integrations', s['h2']))
    story.append(Paragraph(
        'The DQE is built around the <b>Daily Variance Analysis (DVA)</b> workbook, the '
        'standard output of Duetto\'s PMS integration layer. Any integration that produces '
        'a DVA file is compatible with the engine. The following integrations have been '
        'tested and validated:', s['body']))

    integ_data = [
        ['PMS Platform', 'Integration Type', 'Notes'],
        ['Opera / Opera Cloud', 'XML / API', 'Primary target; Arrival Details Report supported'],
        ['Infor HMS', 'File-based', 'Validated; DVA export required'],
        ['Any DVA-producing PMS', '—', 'Compatible if standard DVA columns are present'],
    ]
    col_w = [usable_w * 0.34, usable_w * 0.22, usable_w * 0.44]
    green_box(integ_data, col_w, story)
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        'The DVA workbook is the only <b>required</b> input. All other source files are optional '
        'and progressively unlock additional analysis capabilities.', s['caption']))

    section_rule(story)

    # ── 3. Source file inputs ─────────────────────────────────────────────────
    story.append(KeepTogether([
        Paragraph('Source File Inputs', s['h2']),
        Paragraph(
            'The DQE accepts up to five source files through a drag-and-drop upload interface. '
            'Each file unlocks an additional layer of diagnostic capability:', s['body']),
    ]))

    files_data = [
        ['File', 'Format', 'Required', 'Capability Unlocked'],
        ['DVA (Daily Variance Analysis)', '.xlsx', '✓ Yes',
         'Core rooms & revenue comparison; hotel name extraction'],
        ['Bookings Stay Date Report', '.tsv / .csv', 'Optional',
         'Reservation-level root-cause analysis; missing/phantom booking detection'],
        ['Blocks Stay Date Report', '.tsv / .csv', 'Optional',
         'Block reservation context for room discrepancies'],
        ['Flat Folio Report', '.tsv / .csv', 'Optional',
         'Folio-level revenue analysis; no-show fees, adjustments, refunds'],
        ['Arrival Details Report (Opera)', '.xml / .pdf / .txt', 'Optional',
         'PMS-to-Duetto cross-reference; sync-gap detection & rate variance'],
    ]
    col_w2 = [usable_w*0.30, usable_w*0.12, usable_w*0.10, usable_w*0.48]
    green_box(files_data, col_w2, story)

    section_rule(story)

    # ── 4. Analysis engine ────────────────────────────────────────────────────
    story.append(Paragraph('Analysis Engine & Discrepancy Classification', s['h2']))
    story.append(Paragraph(
        'For every stay date in scope, the DQE computes room and revenue accuracy against '
        'PMS-reported values from the DVA. Failing dates are classified using a library of '
        'named diagnostic codes that identify the most likely root cause:', s['body']))

    codes_data = [
        ['Code', 'Category', 'Description'],
        ['HO-O-10', 'Historic – Rooms Over',    'Stale RESERVED-status bookings on actualized dates'],
        ['HO-O-11', 'Historic – Rooms Over',    'Phantom active bookings — cancellations not synced'],
        ['HO-O-13', 'Historic – Rooms Over',    'Orphaned share reservations in Duetto'],
        ['HO-U-05', 'Historic – Rooms Under',   'Leg-perm / booking leg configuration issue'],
        ['HO-U-06', 'Historic – Rooms Under',   'Missing bookings — integration publisher gap'],
        ['FO-O-45', 'Future – Rooms Over',      'Phantom future bookings — cancellation not received'],
        ['FO-O-46', 'Future – Rooms Over',      'Stale reservations — resync required'],
        ['FO-U-40', 'Future – Rooms Under',     'Missing future bookings — integration setup issue'],
        ['HR-U-14', 'Historic – Revenue Under', 'No-show fees captured in folio but not by Duetto'],
        ['HR-U-21', 'Historic – Revenue Under', 'Active bookings carry $0 rate in Duetto'],
        ['HR-O-24', 'Historic – Revenue Over',  'Folio refund/cancellation fee — PMS negative adjustment'],
        ['FR-U-51', 'Future – Revenue Under',   'Future bookings with $0 or suppressed rate in Duetto'],
        ['FR-O-53', 'Future – Revenue Over',    'Future folio credit — Duetto does not capture adjustments'],
        ['SYNC-GAP', 'PMS Integration Gap',     'Opera Arrivals & Duetto agree, but DVA PMS commit is lower — OHIP sync feed incomplete'],
    ]
    col_w3 = [usable_w*0.14, usable_w*0.26, usable_w*0.60]
    green_box(codes_data, col_w3, story)
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        'Each discrepancy record includes the diagnostic code, a plain-English explanation, '
        'and a list of contributing bookings (confirmation numbers, statuses, rates) to '
        'support immediate action.', s['caption']))

    section_rule(story)

    # ── 5. Accuracy metrics ───────────────────────────────────────────────────
    story.append(KeepTogether([
        Paragraph('Accuracy Metrics', s['h2']),
        Paragraph(
            'The DQE calculates two headline accuracy percentages across the scoped date range:',
            s['body']),
        bullet('<b>Room Accuracy</b> — percentage of stay dates where Duetto\'s committed room count matches the PMS exactly (zero-tolerance threshold).', s),
        bullet('<b>Revenue Accuracy</b> — percentage of stay dates where Duetto\'s committed revenue is within the acceptable variance of the PMS value.', s),
        Spacer(1, 2*mm),
        Paragraph(
            'Dates are scoped automatically to the union of dates present in the uploaded '
            'source files — allowing short-range targeted analysis without reprocessing the '
            'full calendar.', s['body']),
    ]))

    section_rule(story)

    # ── 6. Arrival Details cross-reference ────────────────────────────────────
    story.append(KeepTogether([
        Paragraph('Arrival Details Report — Stay Date Reconciliation', s['h2']),
        Paragraph(
            'When an <b>Opera Arrival Details Report</b> is uploaded alongside the Duetto '
            'Bookings Stay Date Report, the DQE reconciles three independent sources of room '
            'count for the stay date(s) in scope. The report is accepted in <b>.xml, .pdf, or '
            '.txt (tab-delimited)</b> format — the engine auto-detects the format and parses '
            'accordingly.', s['body']),
        Paragraph('The reconciliation surfaces a three-way comparison:', s['body']),
        bullet('<b>Opera Arrivals (PMS)</b> — the true reservation count exported directly from Opera.', s),
        bullet('<b>Duetto Bookings</b> — the reservation count held in Duetto, matched by Opera confirmation number (ALTERNATE_SOURCE_ID).', s),
        bullet('<b>DVA PMS Commit</b> — the PMS room commit figure recorded in the Daily Variance Analysis, fed by the OHIP sync.', s),
        Spacer(1, 2*mm),
        Paragraph(
            '<b>Sync-gap detection.</b> The most valuable pattern this exposes is when Opera '
            'Arrivals and Duetto Bookings <i>agree</i> (e.g. both 71) but the DVA PMS commit is '
            'lower (e.g. 42). This proves the discrepancy is not a Duetto accuracy problem — it '
            'is an incomplete OHIP sync feed under-reporting to the DVA. The affected stay date '
            'is re-classified as <b>SYNC-GAP</b> (amber) rather than a Duetto "Over" failure, '
            'and the exact room gap is called out on-screen and in the export.', s['body']),
        bullet('<b>Rate mismatches</b> — flags reservations where the Opera effective rate differs from the Duetto rate by more than $1.00, surfacing ADR variance root causes.', s),
        Spacer(1, 2*mm),
        Paragraph(
            'Results appear in a dedicated <b>Stay Date Reconciliation</b> panel and the '
            'reframed accuracy tile in the on-screen report, and in a dedicated <b>Arrivals '
            'Reconciliation</b> sheet in the Excel export (three-source comparison, plain-English '
            'summary, and a full rate-mismatch detail table). The original source file is also '
            'uploaded to the Monday.com board for audit purposes.', s['body']),
    ]))

    section_rule(story)

    # ── 7. Report outputs ─────────────────────────────────────────────────────
    story.append(Paragraph('Report Outputs', s['h2']))
    output_data = [
        ['Output', 'Description'],
        ['On-screen Report',
         'Interactive accuracy dashboard with discrepancy table, folio analysis, arrival '
         'cross-reference, and remediation recommendations — rendered immediately in the browser.'],
        ['Excel Workbook (.xlsx)',
         'Structured export with a Discrepancy Report sheet (rooms & revenue, root-cause codes, '
         'recommendations) and — when an Arrival Details Report is supplied — an Arrivals '
         'Reconciliation sheet showing the three-source room comparison, sync-gap summary, and '
         'rate-mismatch detail. Branded with Duetto colours.'],
        ['Monday.com Board Item',
         'Automatic submission to the DQE Feedback board — includes hotel metadata, accuracy scores, '
         'failing dates, all source files, and a notes field for the submitting employee.'],
    ]
    col_w4 = [usable_w*0.24, usable_w*0.76]
    green_box(output_data, col_w4, story)

    section_rule(story)

    # ── 8. Monday.com integration ─────────────────────────────────────────────
    story.append(Paragraph('Monday.com Integration', s['h2']))
    story.append(Paragraph(
        'After running an analysis, results can be submitted directly to the shared '
        '<b>DQE Feedback board on Monday.com</b> via a single-click modal. '
        'The following data is captured on each board item:', s['body']))
    story.append(bullet('Hotel Name and Hotel ID (auto-filled from the DVA filename)', s))
    story.append(bullet('Analysis date, stay date range, rooms accuracy %, revenue accuracy %', s))
    story.append(bullet('Failing dates summary', s))
    story.append(bullet('Files used (list of all uploaded source file names)', s))
    story.append(bullet('Submitted By — workspace member dropdown populated from Monday.com', s))
    story.append(bullet('Free-text Feedback / Notes field', s))
    story.append(bullet('All source files attached in dedicated columns (DVA, Bookings, Blocks, Folio, Arrivals, Excel Report)', s))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        'The board provides a centralised audit trail of all DQE submissions across the team, '
        'with attached files retained for review and escalation.', s['body']))
    story.append(Paragraph(
        'Board link: <a href="https://duettoresearch.monday.com/boards/18419945182" color="#0E2124">'
        '<u>https://duettoresearch.monday.com/boards/18419945182</u></a>', s['body']))

    section_rule(story)

    # ── 9. Recommendations ───────────────────────────────────────────────────
    story.append(KeepTogether([
        Paragraph('Recommendations Engine', s['h2']),
        Paragraph(
            'Based on the frequency and type of diagnostic codes identified, the DQE '
            'automatically generates a prioritised set of recommended remediation actions '
            'for Duetto employees to review and act on. Examples include:', s['body']),
        bullet('Request a historical reservation resync to clear stale RESERVED-status bookings.', s),
        bullet('Review integration XML publisher settings for missing booking codes.', s),
        bullet('Investigate leg-perm configuration with the integration manager.', s),
        bullet('Manually cancel phantom future bookings; verify the integration receives delete messages.', s),
        bullet('Review folio-level integration options where no-show fee revenue is systematically absent.', s),
        Spacer(1, 2*mm),
        Paragraph(
            'Recommendations are included in both the on-screen report and the Excel export, '
            'and grouped by diagnostic code to support efficient action planning by Duetto employees.', s['body']),
    ]))

    section_rule(story)

    # ── 10. Running the tool ──────────────────────────────────────────────────
    story.append(Paragraph('Running the DQE', s['h2']))
    story.append(Paragraph('The DQE runs as a local Flask web application:', s['body']))
    steps_data = [
        ['Step', 'Action'],
        ['1', 'Open Terminal and navigate to the DQE folder'],
        ['2', 'Run:  python app.py'],
        ['3', 'Open a browser and go to:  http://localhost:5055'],
        ['4', 'Upload source files using the drag-and-drop zones'],
        ['5', 'Click Run Analysis and review the on-screen report'],
        ['6', 'Download the Excel report and/or Submit to Monday.com'],
    ]
    col_w5 = [usable_w*0.06, usable_w*0.94]
    green_box(steps_data, col_w5, story)
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        'Requires Python 3.9+ with dependencies installed via pip. No internet connection is '
        'needed for the analysis itself; Monday.com submission requires network access.', s['caption']))

    section_rule(story)

    # ── 11. Key benefits ──────────────────────────────────────────────────────
    story.append(Paragraph('Key Benefits', s['h2']))
    benefits = [
        ('Speed',         'Reduces manual DVA review from hours to seconds. A full analysis across a 365-day date range completes in under 10 seconds.'),
        ('Consistency',   'Every discrepancy is classified against the same ruleset — eliminating analyst-to-analyst variance in root-cause attribution.'),
        ('Depth',         'Booking-level, block-level, folio-level, and PMS-level evidence is surfaced automatically — no manual cross-referencing required.'),
        ('Traceability',  'Every submission creates an auditable Monday.com record with source files attached, supporting QA, escalation, and hand-off workflows.'),
        ('Scalability',   'Distributable as a local app — any Duetto employee can run it independently without a shared server or cloud dependency.'),
    ]
    for title, desc in benefits:
        story.append(KeepTogether([
            Paragraph(f'<b>{title}</b>', s['h3']),
            Paragraph(desc, s['body']),
        ]))

    section_rule(story)

    # ── Closing note ──────────────────────────────────────────────────────────
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        'The Data Quality Engine is an internal Duetto tool. For questions, issues, or feature '
        'requests, contact the DQE team via the Monday.com DQE Feedback board or Slack.',
        s['caption']))

    # ── Build ─────────────────────────────────────────────────────────────────
    doc.build(story, onFirstPage=draw_chrome, onLaterPages=draw_chrome)
    print(f'PDF written to {out}')


if __name__ == '__main__':
    build()
