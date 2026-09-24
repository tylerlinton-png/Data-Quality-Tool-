"""
Generates docs/dqe-user-guide.pdf — a 2-page branded user guide for the
Duetto Data Quality Engine (DQE) app.

Colors, typography choices, and copy are pulled directly from the running
app (templates/index.html CSS variables, app.py routes/behavior, and the
project's actual git remote / requirements.txt / launch config) — nothing
here is invented. Sora (the app's UI font) isn't a built-in PDF font, so
Helvetica is used as the closest standard substitute; everything else
(colors, layout language, button/card shapes) mirrors the live app.

Regenerate after any UI/branding or setup-process change with:
    python3 docs/generate_user_guide.py
"""

import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    NextPageTemplate, PageBreak, KeepTogether, Image as RLImage,
)
from reportlab.pdfbase.pdfmetrics import stringWidth

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, 'dqe-user-guide.pdf')
LOGO_SRC = os.path.join(HERE, '..', 'static', 'duetto-logo.png')
LOGO_WHITE = os.path.join(HERE, 'logo_white.png')


def _ensure_white_logo():
    """
    static/duetto-logo.png is black-on-transparent — the app itself displays
    it via CSS `filter: brightness(0) invert(1)` to show it white on the dark
    nav bar (see .nav-logo-img in templates/index.html). Reportlab can't apply
    CSS filters to an embedded image, so this reproduces the same effect once
    and caches the result alongside this script.
    """
    if os.path.exists(LOGO_WHITE) and os.path.getmtime(LOGO_WHITE) >= os.path.getmtime(LOGO_SRC):
        return
    from PIL import Image
    im = Image.open(LOGO_SRC).convert('RGBA')
    alpha = im.split()[3]
    white = Image.new('RGBA', im.size, (255, 255, 255, 0))
    white.putalpha(alpha)
    white.save(LOGO_WHITE)

# ── Brand colors, lifted directly from templates/index.html :root ──────────
DARK_HEX, GREEN_HEX, GREEN_D_HEX = '#0E2124', '#C4FF45', '#8FBF1F'
AMBER_HEX = '#B45309'

DARK    = HexColor(DARK_HEX)    # --duetto-dark (nav / header band)
GREEN   = HexColor(GREEN_HEX)   # --duetto-green / --accent
GREEN_D = HexColor(GREEN_D_HEX) # darker green for text-on-white (accessible)
BG      = HexColor('#F0F2F6')   # --bg
CARD    = HexColor('#FFFFFF')   # --card
BORDER  = HexColor('#E2E8F0')   # --border
TEXT    = HexColor('#0F172A')   # --text
TEXT_MID= HexColor('#334155')   # --text-mid
MUTED   = HexColor('#64748A')   # darker than --muted (#94A3B8) for print legibility
AMBER   = HexColor(AMBER_HEX)   # darker than --amber for print legibility on light bg
CODE_BG = HexColor('#0E2124')   # code blocks styled like the app's dark surfaces

PAGE_W, PAGE_H = letter
MARGIN = 0.55 * inch

styles = {
    'H1White': ParagraphStyle('H1White', fontName='Helvetica-Bold', fontSize=20,
                               leading=24, textColor=CARD),
    'SubWhite': ParagraphStyle('SubWhite', fontName='Helvetica', fontSize=11,
                                leading=14, textColor=GREEN),
    'Intro': ParagraphStyle('Intro', fontName='Helvetica', fontSize=9.3,
                             leading=13.5, textColor=TEXT_MID, spaceAfter=6),
    'SectionHead': ParagraphStyle('SectionHead', fontName='Helvetica-Bold', fontSize=13,
                                   leading=16, textColor=DARK, spaceBefore=4, spaceAfter=6),
    'StepTitle': ParagraphStyle('StepTitle', fontName='Helvetica-Bold', fontSize=9.6,
                                 leading=12.5, textColor=TEXT),
    'StepBody': ParagraphStyle('StepBody', fontName='Helvetica', fontSize=8.7,
                                leading=12, textColor=TEXT_MID),
    'Code': ParagraphStyle('Code', fontName='Courier', fontSize=8.6,
                            leading=11.5, textColor=GREEN),
    'TroubleHead': ParagraphStyle('TroubleHead', fontName='Helvetica-Bold', fontSize=8.7,
                                   leading=12, textColor=TEXT),
    'TroubleBody': ParagraphStyle('TroubleBody', fontName='Helvetica', fontSize=8.4,
                                   leading=11.5, textColor=TEXT_MID),
    'FeatureName': ParagraphStyle('FeatureName', fontName='Helvetica-Bold', fontSize=11.3,
                                   leading=14, textColor=DARK),
    'FeatureLabel': ParagraphStyle('FeatureLabel', fontName='Helvetica-Bold', fontSize=8.3,
                                    leading=11.5, textColor=GREEN_D),
    'FeatureBody': ParagraphStyle('FeatureBody', fontName='Helvetica', fontSize=8.5,
                                   leading=11.8, textColor=TEXT_MID),
    'PageHead': ParagraphStyle('PageHead', fontName='Helvetica-Bold', fontSize=16,
                                leading=20, textColor=DARK),
}


def _header_band(canvas, doc, height=1.15 * inch, title=None, subtitle=None, show_logo=True):
    canvas.saveState()
    canvas.setFillColor(DARK)
    canvas.rect(0, PAGE_H - height, PAGE_W, height, stroke=0, fill=1)
    # thin green accent line under the band, matching the app's nav bottom border
    canvas.setFillColor(GREEN)
    canvas.rect(0, PAGE_H - height - 2, PAGE_W, 2, stroke=0, fill=1)

    x = MARGIN
    if show_logo and os.path.exists(LOGO_WHITE):
        logo_h = 0.30 * inch
        logo_w = logo_h * (1005 / 226.0)
        canvas.drawImage(LOGO_WHITE, x, PAGE_H - height / 2 - logo_h / 2 - 6,
                          width=logo_w, height=logo_h, mask='auto')
        x += logo_w + 16

    if title:
        canvas.setFillColor(CARD)
        canvas.setFont('Helvetica-Bold', 19)
        canvas.drawString(x, PAGE_H - height / 2 + 6, title)
    if subtitle:
        canvas.setFillColor(GREEN)
        canvas.setFont('Helvetica', 10.5)
        canvas.drawString(x, PAGE_H - height / 2 - 12, subtitle)
    canvas.restoreState()


def _page1_bg(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG)
    canvas.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
    canvas.restoreState()
    _header_band(canvas, doc, height=1.0 * inch,
                  title='Data Quality Engine', subtitle='User Guide')
    canvas.saveState()
    canvas.setFillColor(MUTED)
    canvas.setFont('Helvetica', 7.3)
    canvas.drawRightString(PAGE_W - MARGIN, 0.35 * inch, 'DQE User Guide — internal tool, Duetto')
    canvas.drawString(MARGIN, 0.35 * inch, 'Page 1 of 2')
    canvas.restoreState()


def _page2_bg(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG)
    canvas.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
    canvas.restoreState()
    _header_band(canvas, doc, height=0.75 * inch, title='Application Features', show_logo=False)
    canvas.saveState()
    canvas.setFillColor(MUTED)
    canvas.setFont('Helvetica', 7.3)
    canvas.drawRightString(PAGE_W - MARGIN, 0.35 * inch, 'DQE User Guide — internal tool, Duetto')
    canvas.drawString(MARGIN, 0.35 * inch, 'Page 2 of 2')
    canvas.restoreState()


def code_block(lines):
    """A dark, monospace command block styled like the app's own code/monospace surfaces."""
    para = Paragraph('<br/>'.join(lines), styles['Code'])
    t = Table([[para]], colWidths=[PAGE_W - 2 * MARGIN - 0.35 * inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), CODE_BG),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('ROUNDEDCORNERS', [4, 4, 4, 4]),
    ]))
    return t


def step_row(number, title, body, code_lines=None):
    badge = Paragraph(
        f'<para alignment="center"><font color="{DARK_HEX}" size="9.5">'
        f'<b>{number}</b></font></para>', styles['StepTitle'])
    badge_cell = Table([[badge]], colWidths=[0.26 * inch], rowHeights=[0.26 * inch])
    badge_cell.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), GREEN),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 1),
        ('RIGHTPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('ROUNDEDCORNERS', [13, 13, 13, 13]),
    ]))
    content = [Paragraph(title, styles['StepTitle']), Paragraph(body, styles['StepBody'])]
    if code_lines:
        content.append(Spacer(1, 3))
        content.append(code_block(code_lines))
    inner = Table([[badge_cell, content]], colWidths=[0.4 * inch, None])
    inner.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    return inner


def trouble_row(title, body):
    dot = Paragraph(f'<font color="{AMBER_HEX}">●</font>', styles['TroubleHead'])
    dot_cell = Table([[dot]], colWidths=[0.18 * inch])
    dot_cell.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    content = [Paragraph(title, styles['TroubleHead']), Paragraph(body, styles['TroubleBody'])]
    t = Table([[dot_cell, content]], colWidths=[0.18 * inch, None])
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return t


def feature_card(name, what, how, why):
    body = [
        Paragraph(name, styles['FeatureName']),
        Spacer(1, 2),
        Paragraph('WHAT IT DOES', styles['FeatureLabel']),
        Paragraph(what, styles['FeatureBody']),
        Spacer(1, 2),
        Paragraph('HOW TO USE IT', styles['FeatureLabel']),
        Paragraph(how, styles['FeatureBody']),
        Spacer(1, 2),
        Paragraph('WHY IT MATTERS', styles['FeatureLabel']),
        Paragraph(why, styles['FeatureBody']),
    ]
    t = Table([[body]], colWidths=[(PAGE_W - 2 * MARGIN - 0.16 * inch) / 2])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), CARD),
        ('BOX', (0, 0), (-1, -1), 0.75, BORDER),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('ROUNDEDCORNERS', [8, 8, 8, 8]),
    ]))
    return t


def build():
    _ensure_white_logo()
    doc = BaseDocTemplate(OUT_PATH, pagesize=letter,
                           leftMargin=MARGIN, rightMargin=MARGIN,
                           topMargin=1.25 * inch, bottomMargin=0.55 * inch,
                           title='DQE User Guide')

    frame1 = Frame(MARGIN, 0.55 * inch, PAGE_W - 2 * MARGIN, PAGE_H - 1.25 * inch - 0.55 * inch,
                    id='f1')
    frame2 = Frame(MARGIN, 0.55 * inch, PAGE_W - 2 * MARGIN, PAGE_H - 1.0 * inch - 0.55 * inch,
                    id='f2')

    doc.addPageTemplates([
        PageTemplate(id='Page1', frames=[frame1], onPage=_page1_bg),
        PageTemplate(id='Page2', frames=[frame2], onPage=_page2_bg),
    ])

    story = []

    # ── Page 1 ──────────────────────────────────────────────────────────────
    story.append(Paragraph(
        'Get the application running locally in a few simple steps.',
        ParagraphStyle('Sub', fontName='Helvetica-Oblique', fontSize=9.5,
                        textColor=MUTED, spaceAfter=8)))
    story.append(Paragraph(
        'DQE (Data Quality Engine) pulls live data directly from Opera Cloud / OHIP and compares '
        'it against Duetto\'s own reports — replacing manual Postman calls and spreadsheet diffing '
        'with one page. It\'s built for the DQE team (root-causing Duetto-vs-PMS discrepancies) and '
        'the Deployment/Migrations team (pre-go-live configuration checks).',
        styles['Intro']))

    story.append(Paragraph('Getting Started', styles['SectionHead']))

    story.append(step_row(1, 'Access GitHub',
        'The app lives in a private GitHub repository. You\'ll need to be added as a collaborator '
        'before you can clone it — ask Tyler Linton for access if you don\'t have it yet.',
        ['github.com/tylerlinton-png/Data-Quality-Tool-']))
    story.append(step_row(2, 'Clone the repository',
        'Open a terminal and run:',
        ['git clone https://github.com/tylerlinton-png/Data-Quality-Tool-.git']))
    story.append(step_row(3, 'Open the project folder',
        None if False else 'Move into the project directory:',
        ['cd Data-Quality-Tool-']))
    story.append(step_row(4, 'Install dependencies',
        'Requires Python 3.9 or later. Install the required packages:',
        ['pip3 install -r requirements.txt']))
    story.append(step_row(5, 'Configuration',
        'No .env file or API keys are needed to install and run the app. Your Oracle OHIP '
        'credentials are entered directly in the app itself (Step 1 on the main page) and are '
        'saved locally on your machine — never in the codebase.'))
    story.append(step_row(6, 'Start the application',
        'Run the app, then open it in your browser:',
        ['python3 app.py', '→ Open http://localhost:5055']))

    story.append(Spacer(1, 4))
    story.append(Paragraph('Quick Troubleshooting', styles['SectionHead']))
    story.append(trouble_row('“command not found: pip3” or “python3”',
        'Install Python 3.9+ from python.org, or via Homebrew (brew install python3) on Mac.'))
    story.append(trouble_row('“Address already in use” on port 5055',
        'Another copy of the app is already running. Find and stop it, or close the old terminal '
        'window, then restart.'))
    story.append(trouble_row('Permission denied cloning the repo',
        'You haven\'t been added as a GitHub collaborator yet — request access before cloning.'))
    story.append(trouble_row('“Missing required field(s)” when validating credentials',
        'Double-check Hostname, Hotel ID, App Key, and Client ID/Secret are all filled in — all '
        'four are required to connect to OHIP.'))
    story.append(trouble_row('A pull or comparison comes back empty',
        'Confirm the date range actually has data for that hotel, and that the External System '
        'Code (Step 1) matches what\'s configured for that property in Opera.'))

    story.append(NextPageTemplate('Page2'))
    story.append(PageBreak())

    # ── Page 2 ──────────────────────────────────────────────────────────────
    story.append(Spacer(1, 4))

    f1 = feature_card('Pull Bookings / Blocks / Folio',
        'Fetches raw reservations, blocks, or folio data straight from Opera Cloud for a date '
        'range — no Postman collection required.',
        'Pick a data type and date range in the Pull card, click Fetch Data, then search, filter, '
        'or download the results as CSV/TSV.',
        'Removes the need for manual Postman calls for everyday data checks.')
    f2 = feature_card('Upload DVA &amp; Compare',
        'The core analysis engine — compares Duetto’s validation export, bookings, and '
        'folio data against Opera to score accuracy and root-cause every discrepancy.',
        'Drop in whichever files you have (all are optional), click Run Analysis, then review the '
        'Summary, Discrepancies, and Recommendations tabs.',
        'Turns a multi-hour manual spreadsheet comparison into a one-click report with an actual '
        'root cause, not just a mismatch.')
    row1 = Table([[f1, f2]], colWidths=[(PAGE_W - 2 * MARGIN - 0.16 * inch) / 2] * 2)
    row1.setStyle(TableStyle([('LEFTPADDING', (1, 0), (1, 0), 8),
                               ('RIGHTPADDING', (0, 0), (0, 0), 8)]))
    story.append(row1)
    story.append(Spacer(1, 8))

    f3 = feature_card('Pull Deployment Info (Config Snapshot)',
        'Runs six pre-go-live checks in one click — Opera Cloud Version, Block Status Codes, '
        'Transaction Codes, Room Types, Hotel Info, and Rate Plans.',
        'Click Run Config Snapshot, review each result card, and use Download All for a single '
        'Excel workbook with everything.',
        'Gives Deployment/Migrations a fast, complete sanity check before a hotel goes live.')
    f4 = feature_card('Hotel Stats',
        'Pulls Revenue Inventory Statistics — rooms sold, revenue, arrivals/departures/'
        'cancellations — directly from Opera for a date range.',
        'Set a date range and click Run Hotel Stats. Results take a little longer since it’s '
        'a multi-step live job, not an instant call.',
        'Answers PMS-side inventory/revenue questions without needing a separate Opera report.')
    row2 = Table([[f3, f4]], colWidths=[(PAGE_W - 2 * MARGIN - 0.16 * inch) / 2] * 2)
    row2.setStyle(TableStyle([('LEFTPADDING', (1, 0), (1, 0), 8),
                               ('RIGHTPADDING', (0, 0), (0, 0), 8)]))
    story.append(row2)
    story.append(Spacer(1, 8))

    f5 = feature_card('Live API Comparisons',
        'Compares a live OHIP pull directly against Duetto’s own Folio or Bookings report — '
        'side by side, including the ones that match.',
        'Pull live data (or reuse a Browse Data pull), attach the matching Duetto report, then run '
        'the analysis — results appear as their own tab.',
        'Answers "can I trust what Duetto is showing" without waiting on a fresh export.')
    f6 = feature_card('Send Feedback',
        'A direct line to report a bug or suggest an improvement while using the app.',
        'Click Feedback in the top navigation, describe the issue or idea, and submit.',
        'Feedback goes straight to the team building the tool — nothing gets lost in chat.')
    row3 = Table([[f5, f6]], colWidths=[(PAGE_W - 2 * MARGIN - 0.16 * inch) / 2] * 2)
    row3.setStyle(TableStyle([('LEFTPADDING', (1, 0), (1, 0), 8),
                               ('RIGHTPADDING', (0, 0), (0, 0), 8)]))
    story.append(row3)

    doc.build(story)


if __name__ == '__main__':
    build()
    print(f'Wrote {OUT_PATH}')
