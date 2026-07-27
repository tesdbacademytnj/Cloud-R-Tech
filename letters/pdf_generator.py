"""
Cloud R tech — PDF Letter Generator  (v4 – pixel-perfect match)

All measurements derived by PyMuPDF bbox analysis of the reference PDF:
  Page size:     595.56 x 842.52 pt  (A4)
  Left margin:   42.5 pt  (15.0 mm)   ← letterhead content origin
  Right margin:  79.6 pt              ← text ends ≈ 516 pt
  Top margin:    152 pt               ← first content (heading)
  Body font:     TimesNewRoman  12 pt  leading≈17 pt
  Title font:    TimesNewRoman-BoldItalic  12 pt  centred
  Table numbers: TimesNewRoman  12.48 pt
"""

import os
import math
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from django.conf import settings
from reportlab.platypus import Flowable
from reportlab.lib.utils import ImageReader


class _SignatureSeal(Flowable):
    """
    Exact measurements from DOCX XML anchor analysis:
      Seal stamp : x=60.7pt, y_top=11.8pt, 82.8x83.4pt  (image4.png — original RGBA)
      Signature  : x=2.9pt,  y_top=17.8pt, 80.2x51.8pt  (image5.png — shifted 15pt down)
    Flowable height = 104.2 pt.
    sig_src: file path string or BytesIO of custom signature (None = default)
    """
    HEIGHT = 104.2

    def __init__(self, sig_src=None):
        super().__init__()
        self.width   = 200
        self.height  = self.HEIGHT
        self.sig_src = sig_src  # None => use default SIGNATURE path

    def draw(self):
        c = self.canv
        h = self.HEIGHT
        seal_y = h - 11.8 - 83.4
        sig_y  = h - 40.0 - 51.8
        sig = self.sig_src if self.sig_src is not None else SIGNATURE
        try:
            if sig:
                c.drawImage(ImageReader(sig), 0.0, sig_y,  width=80.2, height=51.8, mask='auto')
            if SEAL_STAMP:
                c.drawImage(ImageReader(SEAL_STAMP), 60.7, seal_y, width=82.8, height=83.4, mask='auto')
        except Exception:
            pass

def _img_path(filename):
    """Return absolute path to an image file; None if the file does not exist."""
    path = os.path.join(settings.BASE_DIR, 'static', 'letters', 'img', filename)
    return path if os.path.isfile(path) else None

LETTERHEAD = _img_path('letterhead.png')
SIGNATURE  = _img_path('signature.png')
SEAL_STAMP = _img_path('seal_stamp.png')

# ── Exact colours from pixel analysis ────────────────────────────────────────
BODY_FG = colors.black          # #000000 — verified by PyMuPDF span colour

# ── Font sizes — measured directly from template spans ───────────────────────
FS        = 12      # body / address / signature  (template uses 12 pt)
FS_ADDR   = 11.52   # "To" block address lines    (template: 11.52 pt)
FS_TBL    = 12.48   # salary table numbers        (template: 12.48 pt)
FS_TITLE  = 12      # heading (same 12 pt, just BoldItalic + centred)


# ── Salutation helper ─────────────────────────────────────────────────────────
def _salutation(gender: str) -> str:
    return 'Mr.' if str(gender).strip().lower() in ('male', 'mr', 'mr.', 'm') else 'Ms.'


# ── Paragraph styles — tuned to template measurements ────────────────────────
def _styles():
    """
    Leading values:
      Body text rows Y-gap ≈ 17.5 pt  → leading=17.5
      Address lines Y-gap ≈ 15.7 pt   → leading=15.7
    spaceAfter=0 everywhere; spacing is injected as explicit Spacer() objects
    so we control every gap to match the template exactly.
    """
    kw = dict(textColor=BODY_FG, spaceAfter=0, spaceBefore=0)

    normal = ParagraphStyle('CRTNormal',
                            fontName='Times-Roman', fontSize=FS,
                            leading=17.5, **kw)
    bold   = ParagraphStyle('CRTBold',
                            fontName='Times-Bold', fontSize=FS,
                            leading=17.5, **kw)
    bold_italic = ParagraphStyle('CRTBoldItalic',
                                 fontName='Times-BoldItalic', fontSize=FS,
                                 leading=17.5, **kw)
    italic = ParagraphStyle('CRTItalic',
                            fontName='Times-Italic', fontSize=FS,
                            leading=17.5, **kw)
    # Heading: centred, BoldItalic, 12 pt — exactly as in template
    title  = ParagraphStyle('CRTTitle',
                            fontName='Times-BoldItalic', fontSize=FS_TITLE,
                            leading=17.5, alignment=TA_CENTER, **kw)
    # Justified body paragraphs
    justify = ParagraphStyle('CRTJustify',
                             fontName='Times-Roman', fontSize=FS,
                             leading=17.5, alignment=TA_JUSTIFY, **kw)
    # Address block — slightly smaller (11.52 pt) matching template
    addr   = ParagraphStyle('CRTAddr',
                            fontName='Times-Roman', fontSize=FS_ADDR,
                            leading=15.7, **kw)
    addr_bold = ParagraphStyle('CRTAddrBold',
                               fontName='Times-Bold', fontSize=FS_ADDR,
                               leading=15.7, **kw)
    # Right-aligned (used for Offer Letter date)
    right  = ParagraphStyle('CRTRight',
                            fontName='Times-Bold', fontSize=FS,
                            leading=17.5, alignment=TA_RIGHT, **kw)
    return dict(normal=normal, bold=bold, bold_italic=bold_italic,
                italic=italic, title=title, justify=justify,
                addr=addr, addr_bold=addr_bold, right=right)


# ── Letterhead drawn on every page ───────────────────────────────────────────
def _draw_letterhead(c, doc):
    if not LETTERHEAD:
        return
    c.saveState()
    c.drawImage(LETTERHEAD, 0, 0, width=A4[0], height=A4[1],
                preserveAspectRatio=False, mask='auto')
    c.restoreState()


# ── Document factory ──────────────────────────────────────────────────────────
# Template measurements (in points):
#   Content starts at Y=152.6 from top  → topMargin = 842.52 − 152.6 = 689.9 ... but
#   ReportLab topMargin = space above the first flowable.
#   First element (heading) is at y=152.6 from PDF top = 689.9 pt from bottom.
#   Frame top = A4[1] − topMargin.  We want frame top = ~690 pt from bottom.
#   → topMargin = 842.52 − 690 = 152.5 pt
#
#   Left margin:  42.5 pt  (template left edge of all text)
#   Right margin: 79.6 pt  (template: text ends ~516 pt, page=595.56)
#   Bottom margin: ~35 mm = 99 pt  (footer area in letterhead image)
def _make_doc(buffer):
    doc = BaseDocTemplate(
        buffer, pagesize=A4,
        leftMargin=42.5,
        rightMargin=79.6,
        topMargin=152.5,
        bottomMargin=99,
    )
    frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        doc.width, doc.height,
        id='main',
        leftPadding=0, rightPadding=0,
        topPadding=0,  bottomPadding=0,
    )
    doc.addPageTemplates([PageTemplate(
        id='letterhead', frames=[frame], onPage=_draw_letterhead
    )])
    return doc


# ── Annexure info block ───────────────────────────────────────────────────────
# Template measurements (page 2):
#   Label col: x=66.3, width=141.1 pt
#   Colon col: x=207.4, width=36 pt  (colon at 207.4, value at 243.4)
#   Value col: x=243.4
#   Row gap:   ≈24.6 pt  (label Y spacing 217→241→266)
def _ann_info_table(rows_data):
    """rows_data: list of [label, value]"""
    N = ParagraphStyle('AnnN', fontName='Times-Roman', fontSize=FS,
                       leading=17.5, textColor=BODY_FG, spaceAfter=0)
    rows = [[Paragraph(r[0], N), Paragraph(':', N), Paragraph(r[1], N)]
            for r in rows_data]
    # col widths: label=141 pt, colon+gap=36 pt, value=rest
    tbl = Table(rows, colWidths=[141, 36, 200], hAlign='LEFT')
    tbl.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), 'Times-Roman'),
        ('FONTSIZE',      (0, 0), (-1, -1), FS),
        ('LEADING',       (0, 0), (-1, -1), 17.5),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7.5),   # 24.6 pt gap ≈ 17 leading + 7.5 pad
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
    ]))
    return tbl


# ── Salary table ──────────────────────────────────────────────────────────────
# Template table (page 2) measurements:
#   Total table width: from x≈59.5 to x≈370  → 310.5 pt
#   Component col: 59.5→~225  = 165 pt
#   Monthly col:   ~225→~285  =  60 pt
#   Annual col:    ~285→~370  =  85 pt  (total 310 pt)
#   Row height:    ≈17.5 pt  (matches body leading)
#   Header row extra height: ≈23 pt
#   Table numbers: 12.48 pt  Times-Roman (regular, right-aligned)
#   Header labels: 12.48 pt  Times-BoldItalic, centred
#   Component header: 12 pt  Times-BoldItalic, centred
def _salary_table(monthly_ctc):
    m      = float(monthly_ctc)
    basic  = round(m * 0.40)
    hra    = round(m * 0.20)
    med    = round(m * 0.10)
    perf   = round(m * 0.15)
    conv   = round(m * 0.10)
    mgmt   = int(m) - (basic + hra + med + perf + conv)
    ctc    = int(m)
    deduct = round(basic / 12) + 250
    net    = ctc - deduct

    def f(v):
        """Monthly: western comma format  e.g. 18,000"""
        return f"{v:,.0f}"

    def fa(v):
        """Annual: Indian comma format  e.g. 2,16,000"""
        s = f"{v:.0f}"
        if len(s) <= 3:
            return s
        result = s[-3:]
        s = s[:-3]
        while s:
            result = s[-2:] + "," + result
            s = s[:-2]
        return result.lstrip(",")

    # Exact font styles matching template
    H = ParagraphStyle('TH', fontName='Times-BoldItalic', fontSize=FS_TBL,
                       textColor=BODY_FG, alignment=TA_CENTER,
                       leading=FS_TBL * 1.2, spaceAfter=0, spaceBefore=0)
    HC = ParagraphStyle('THC', fontName='Times-BoldItalic', fontSize=FS,
                        textColor=BODY_FG, alignment=TA_CENTER,
                        leading=FS * 1.2, spaceAfter=0, spaceBefore=0)
    N  = ParagraphStyle('TN',  fontName='Times-Roman', fontSize=FS,
                        textColor=BODY_FG, leading=17.5,
                        spaceAfter=0, spaceBefore=0)
    B  = ParagraphStyle('TB',  fontName='Times-BoldItalic', fontSize=FS,
                        textColor=BODY_FG, leading=17.5,
                        spaceAfter=0, spaceBefore=0)

    rows = [
        # Header row — Component centred at 12pt, Monthly/Annual centred at 12.48pt
        [Paragraph('Component', HC),
         Paragraph('Monthly',   H),
         Paragraph('Annual',    H)],
        [Paragraph('Basic Salary',                  N),  f(basic),   fa(basic  * 12)],
        [Paragraph('House Rent Allowance (HRA)',    N),  f(hra),     fa(hra    * 12)],
        [Paragraph('Medical Allowance',             N),  f(med),     fa(med    * 12)],
        [Paragraph('Performance Bonus',             N),  f(perf),    fa(perf   * 12)],
        [Paragraph('Conveyance',                    N),  f(conv),    fa(conv   * 12)],
        [Paragraph('Allowance Management',          N),  f(mgmt),    fa(mgmt   * 12)],
        [Paragraph('Total Compensation (CTC)',      B),  f(ctc),     fa(ctc    * 12)],
        [Paragraph('Deduction',                     N),  f(deduct),  fa(deduct * 12)],
        [Paragraph('Net Salary',                    B),  f(net),     fa(net    * 12)],
    ]

    # Col widths: 165 + 60 + 85 = 310 pt (matches template ~310.5 pt)
    tbl = Table(rows, colWidths=[165, 85, 85], hAlign='LEFT',
                rowHeights=None)   # let ReportLab size rows naturally

    tbl.setStyle(TableStyle([
        # Grid
        ('BOX',           (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID',     (0, 0), (-1, -1), 0.5, colors.black),
        # No header background — matches template (white/transparent)
        ('BACKGROUND',    (0, 0), (-1,  0), colors.white),
        # All text top-middle aligned
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        # Numbers right-aligned in cols 1 & 2
        ('ALIGN',         (1, 0), (-1, -1), 'RIGHT'),
        # Padding — matches template row height ≈17.5 pt with 12 pt font
        ('TOPPADDING',    (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING',   (0, 0), (0, -1), 5),   # label col indent
        ('RIGHTPADDING',  (1, 0), (-1, -1), 5),  # number col right gap
        ('LEFTPADDING',   (1, 0), (-1, -1), 3),
        # Font overrides for number cells (rows 1-9, cols 1-2)
        ('FONTNAME',      (1, 1), (-1, -1), 'Times-Roman'),
        ('FONTSIZE',      (1, 1), (-1, -1), FS_TBL),
        # Bold rows
        ('FONTNAME',      (1, 7), (-1, 7),  'Times-Roman'),
        ('FONTNAME',      (1, 9), (-1, 9),  'Times-Roman'),
    ]))
    return tbl


# ── Signature block ───────────────────────────────────────────────────────────
DEFAULT_HR_NAME = 'Raj Padmanaban'

def _signature_block(story, s, hr_name=None, sig_src=None):
    """
    hr_name : display name under signature (default = DEFAULT_HR_NAME)
    sig_src : file path or BytesIO of custom PNG (None = default signature.png)
    """
    name = hr_name if hr_name else DEFAULT_HR_NAME
    story.append(Paragraph('For <b><i>M/s Cloud R tech</i></b>', s['normal']))
    story.append(_SignatureSeal(sig_src=sig_src))
    story.append(Paragraph(f'<b><i>{name}</i></b>', s['bold_italic']))
    story.append(Paragraph('<b><i>HR Manager</i></b>', s['bold_italic']))


# ═══════════════════════════════════════════════════════════════════════════════
# 1.  APPRAISAL LETTER  — pixel-perfect v4
# ═══════════════════════════════════════════════════════════════════════════════
def generate_appraisal(data: dict) -> bytes:
    """
    Exact spacing derived from PyMuPDF bbox analysis of reference PDF:

    Heading Y:         152.6 pt from top
    'To' Y:            199.4 pt  → gap from heading bottom (165.9) = 33.5 pt
    Name Y:            226.0 pt  → gap from To = 26.6 pt
    Address line 1 Y:  241.8 pt  → line gap = 15.8 pt  (addr leading)
    Pincode Y:         257.4 pt  → line gap = 15.6 pt
    Dear Y:            326.7 pt  → gap from pincode bottom (270.1) = 56.6 pt
    Body Y:            366.5 pt  → gap from Dear = 39.8 pt
    Para 2 Y:          432.7 pt  → gap from body end (415.8) = 16.9 pt
    Signature Y:       552.7 pt  → gap from para 2 end (445.9) = 106.8 pt
    HR Manager Y:      675.3 pt  → gap from signature line = 122.6 pt
    """
    s   = _styles()
    buf = BytesIO()
    doc = _make_doc(buf)
    story = []

    name        = data['name']
    sal         = _salutation(data.get('gender', 'male'))
    address     = data['address']
    pincode     = data['pincode']
    date_str    = data['date']
    current     = float(data['current_monthly'])
    new_monthly = float(data['new_monthly'])
    designation = data['designation']
    ctc_lpa     = data.get('ctc_lpa', '').replace('LPA', '').replace('lpa', '').strip()

    # ── HEADING  (Y=152.6) ───────────────────────────────────────────────────
    story.append(Paragraph('<b><i>APPRAISAL LETTER</i></b>', s['title']))
    # Gap: heading bottom ≈165.9, 'To' at 199.4 → spacer = 199.4 − 165.9 = 33.5 pt
    story.append(Spacer(1, 33.5))

    # ── TO BLOCK  (Y=199.4) ──────────────────────────────────────────────────
    story.append(Paragraph('<b>To</b>', s['addr_bold']))
    # Gap: 'To' bottom ≈212.1, Name at 226.0 → spacer = 226.0 − 212.1 = 13.9 pt
    story.append(Spacer(1, 13.9))

    story.append(Paragraph(f'{sal} {name},', s['addr']))
    story.append(Paragraph(address, s['addr']))
    story.append(Paragraph(f'Pincode - {pincode}', s['addr']))
    # Gap: pincode bottom ≈270.1, Dear at 326.7 → spacer = 56.6 pt
    story.append(Spacer(1, 56.6))

    # ── DEAR LINE  (Y=326.7) ─────────────────────────────────────────────────
    story.append(Paragraph(f'Dear {sal}{name},', s['normal']))
    # Gap: Dear bottom ≈340.0, Body at 366.5 → spacer = 26.5 pt
    story.append(Spacer(1, 26.5))

    # ── BODY PARAGRAPH 1  (Y=366.5) ──────────────────────────────────────────
    body = (
        f'This Appraisal Letter is made on <b><i>{date_str}</i></b>, to appreciate your '
        f'performance. We revise the remuneration from INR {current:,.0f}/- per month to '
        f'INR {new_monthly:,.0f}/- per month with effect from beginning of the next month. '
        f'A detailed salary structure is given in the annexure.'
    )
    story.append(Paragraph(body, s['justify']))
    # Gap: body end ≈415.8, para2 at 432.7 → spacer = 16.9 pt
    story.append(Spacer(1, 16.9))

    # ── BODY PARAGRAPH 2 ─────────────────────────────────────────────────────
    story.append(Paragraph(
        'All the other terms and conditions of the original contract remain unchanged.',
        s['normal']
    ))
    # Gap: para2 end ≈445.9, For M/s at 552.7 → spacer = 85.6 pt (adjusted for frame offset)
    story.append(Spacer(1, 85.6))

    # ── SIGNATURE ────────────────────────────────────────────────────────────
    _signature_block(story, s, hr_name=data.get('hr_name') or None, sig_src=data.get('sig_src'))

    # ── PAGE 2: ANNEXURE ─────────────────────────────────────────────────────
    story.append(PageBreak())

    # Annexure heading Y=192.1 → same topMargin positions it correctly
    story.append(Paragraph('<b><i>Annexure - (Salary Details)</i></b>', s['bold_italic']))
    # Gap: heading bottom ≈205.4, Name row at 217.1 → spacer = 11.7 pt
    story.append(Spacer(1, 11.7))

    story.append(_ann_info_table([
        ['Name',                f'{sal} {name}'],
        ['Designation',         designation],
        ['Cost to the company', f'{ctc_lpa} LPA'],
    ]))
    # Gap: CTC row bottom ≈279.7, Table at 344.9 → spacer = 65.2 pt
    story.append(Spacer(1, 65.2))
    story.append(_salary_table(new_monthly))

    doc.build(story)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# 2.  EXPERIENCE / RELIEVING LETTER
# ═══════════════════════════════════════════════════════════════════════════════
def generate_experience(data: dict) -> bytes:
    s   = _styles()
    buf = BytesIO()
    doc = _make_doc(buf)
    story = []

    name              = data['name']
    sal               = _salutation(data.get('gender', 'male'))
    ref               = data['ref']
    closing_date      = data['closing_date']
    designation       = data['designation']
    date_of_joining   = data['date_of_joining']
    date_of_relieving = data['date_of_relieving']
    reason            = data['reason']
    conduct           = data['conduct']

    def _super_ordinal(text: str) -> str:
        import re
        return re.sub(r'(\d+)(st|nd|rd|th)', r'\1<super>\2</super>', text)

    story.append(Paragraph(f'Ref : <b><i>{ref}</i></b>', s['normal']))
    story.append(Spacer(1, 14))
    story.append(Paragraph('<b><i>Sub: Relieving / Experience Letter</i></b>', s['title']))
    story.append(Spacer(1, 14))
    story.append(Paragraph(f'<b>Dear {sal}{name} ,</b>', s['bold']))
    story.append(Spacer(1, 14))

    body = (
        'This is with reference to your resignation from the services of this Organization. '
        'Your resignation has been accepted and you are relieved from the services of the '
        f'Organization at the closing of business hours on <b><i>{closing_date}</i></b>.'
    )
    story.append(Paragraph(body, s['justify']))
    story.append(Spacer(1, 18))

    N = ParagraphStyle('DtN', fontName='Times-Roman', fontSize=FS,
                       leading=17.5, textColor=BODY_FG, spaceAfter=0)
    det_rows = [
        [Paragraph('Designation',          N), Paragraph(':', N), Paragraph(designation, N)],
        [Paragraph('Date of Joining',      N), Paragraph(':', N),
         Paragraph(_super_ordinal(date_of_joining), N)],
        [Paragraph('Date of Relieving',    N), Paragraph(':', N),
         Paragraph(_super_ordinal(date_of_relieving), N)],
        [Paragraph('Reason for Relieving', N), Paragraph(':', N), Paragraph(reason, N)],
        [Paragraph('Conduct',              N), Paragraph(':', N),
         Paragraph(f'<i>{conduct}</i>', N)],
    ]
    det_tbl = Table(det_rows, colWidths=[141, 36, 200], hAlign='LEFT')
    det_tbl.setStyle(TableStyle([
        ('FONTSIZE',      (0, 0), (-1, -1), FS),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7.5),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
    ]))
    story.append(det_tbl)
    story.append(Spacer(1, 18))
    story.append(Paragraph('We wish you all the best in your future endeavors.', s['normal']))
    story.append(Spacer(1, 14))
    _signature_block(story, s, hr_name=data.get('hr_name') or None, sig_src=data.get('sig_src'))

    doc.build(story)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# 3.  OFFER LETTER
# ═══════════════════════════════════════════════════════════════════════════════
def generate_offer(data: dict) -> bytes:
    s   = _styles()
    buf = BytesIO()
    doc = _make_doc(buf)
    story = []

    date_str    = data['date']
    name        = data['name']
    gender      = data.get('gender', 'male')
    sal         = _salutation(gender)
    address     = data['address']
    pincode     = data['pincode']
    designation = data['designation']
    joining_date= data['joining_date']
    ctc_monthly = data['ctc_monthly']
    ctc_lpa     = data['ctc_lpa'].replace('LPA', '').replace('lpa', '').strip()

    story.append(Paragraph(f'<b>Date : {date_str}</b>', s['right']))
    story.append(Spacer(1, 10))
    story.append(Paragraph('<b>To</b>', s['bold']))
    story.append(Spacer(1, 5))
    story.append(Paragraph(f'{sal} {name} ,', s['normal']))
    story.append(Paragraph(address, s['normal']))
    story.append(Paragraph(f'Pincode - {pincode}', s['normal']))
    story.append(Spacer(1, 14))
    story.append(Paragraph(f'Dear <b>{sal} {name} ,</b>', s['normal']))
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        f'We are pleased to appoint you to the position of <b>{designation}</b> '
        f'with <b>M/s Cloud R tech</b> and you will be working with us with the following '
        f'terms and conditions',
        s['justify']
    ))
    story.append(Spacer(1, 10))

    all_clauses = [
        f'Your Joining period will commence on <b>{joining_date}</b> on which you will be '
        f'reporting to HR person at 10am in the previously communicated address',
        f'Your annual total compensation will be INR <b>{ctc_lpa}</b> LPA. Apart from this, '
        f'you are eligible for shift allowance depending on the project cost. Please find the '
        f'salary structure given in the annexure.',
        'Your present place of work will be at Chennai, but during the course of the above '
        'assignment, you shall be liable to be posted / transferred anywhere to serve any of the '
        "Company's Projects or any other establishment in India or outside, at the sole discretion "
        'of the Management.',
        'You will be eligible for 12 working days of vacation and public holidays as notified '
        'by the company.',
        'This appointment is subject to you, having been found medically (physically and mentally) '
        'fit by the authorized Medical Practitioner.',
        "You will not (except in the normal course of the Company's business) publish any article "
        "or statement, deliver any lecture or broadcast or make any communication to the press, "
        "including magazine publication relating to the Company's products or to any matter with "
        "which the Company may be concerned, unless you have previously applied to and obtained.",
        'Any of our technical or other important information which might come into your possession '
        'during the continuance of your assignment with us shall not be disclosed, divulged or made '
        'public by you even thereafter.',
        'If at any time in our opinion, which is final in this matter you are found non-performer '
        'or guilty of fraud, dishonest, disobedience, disorderly behavior, negligence, '
        'indiscipline, absence from duty without permission or any other conduct considered by us '
        'deterrent to our interest or of violation of one or more terms of this letter, your '
        'services may be terminated without notice and on account of reason of any of the acts or '
        'omission, the company shall be entitled to recover the damages from you.',
        'All appointments are based on the information furnished by you in your employment '
        'application and all further declarations and undertakings. Hence, any false statement or '
        'information furnished as above will lead to your dismissal without notice.',
        'You hereby warrant that you are not in breach of any contract with any third party or '
        'restricted in any way in your ability to undertake or perform the duties of your '
        'employment. During your employment with the Company you will agree to work on any project '
        'that you are assigned to, irrespective of technical platforms / skills and nature of the '
        'project. If necessary, you may also be required to work in shifts.',
        'Regardless of any other M/s Cloud R tech entities or where you may be required to '
        'work overseas for any such M/s Cloud R tech entities for an extensive period, you '
        'shall at all times remain an employee of the Company exclusively and shall not be entitled '
        'to any such foreign salary or benefits (including medical insurance, green card '
        'sponsorship, etc.) payable or applicable to employees of such other M/s Cloud R '
        'tech entities other than the salary and benefits specified in this offer letter.',
        'You also agree that if you breach any of the terms and conditions stipulated in this '
        'Agreement, you will be liable for any loss or damage suffered directly or indirectly by '
        'the Company as a result of your action.',
        'You will be responsible for safekeeping and return in good condition and order of all '
        'company property, which may be in your use, custody or charge.',
        "The Company shall be at liberty to amend the whole or any part of this Agreement after "
        "its execution if it considers it necessary and reasonable but only after two (2) weeks' "
        "notice is duly given to you informing you of the proposed amendment(s). The terms and "
        "conditions of your proposed employment with the company in this letter supersede any "
        "contrary verbal representations concerning conditions of employment.",
    ]
    for i, clause in enumerate(all_clauses, 1):
        story.append(Paragraph(f'{i}.&nbsp;&nbsp;{clause}', s['justify']))
        story.append(Spacer(1, 6))

    story.append(Spacer(1, 10))
    story.append(Paragraph(
        'We would appreciate you, confirming your acceptance by signing on the space provided '
        'and returning this letter to HR department, indicating your proposed start date. '
        'Let me close by reaffirming our belief that the skill and background you bring to '
        '<b>M/s Cloud R tech</b> will be instrumental to the future success of the Company. '
        'Without hesitation, the single most important factor in our success has been our people. '
        'We look forward to working with you very soon.',
        s['justify']
    ))
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        'We welcome you to <b>M/s Cloud R tech</b> family and look forward to a '
        'fruitful collaboration. With best wishes,',
        s['normal']
    ))
    story.append(Spacer(1, 12))
    _signature_block(story, s, hr_name=data.get('hr_name') or None, sig_src=data.get('sig_src'))

    story.append(PageBreak())
    story.append(Paragraph('<b><i>Annexure - (Salary Details)</i></b>', s['bold_italic']))
    story.append(Spacer(1, 11.7))
    story.append(_ann_info_table([
        ['Name',                f'{sal} {name}'],
        ['Designation',         designation],
        ['Cost to the company', f'{ctc_lpa} LPA'],
    ]))
    story.append(Spacer(1, 65.2))
    story.append(_salary_table(ctc_monthly))

    doc.build(story)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# 4.  CONTRACT EXTENSION
# ═══════════════════════════════════════════════════════════════════════════════
def generate_contract(data: dict) -> bytes:
    s   = _styles()
    buf = BytesIO()
    doc = _make_doc(buf)
    story = []

    ref               = data['ref']
    name              = data['name']
    gender            = data.get('gender', 'male')
    sal               = _salutation(gender)
    address           = data['address']
    pincode           = data['pincode']
    extended_date     = data['extended_date']
    current_ctc_lpa   = data['current_ctc_lpa'].replace('LPA', '').replace('lpa', '').strip()
    increment_ctc_lpa = data['increment_ctc_lpa'].replace('LPA', '').replace('lpa', '').strip()
    effective_month   = data['effective_month']
    designation       = data['designation']
    ctc_monthly       = data['ctc_monthly']

    story.append(Paragraph(f'<b><i>Ref : {ref}</i></b>', s['bold_italic']))
    story.append(Spacer(1, 10))
    story.append(Paragraph('<b><i>CONTRACT EXTENSION</i></b>', s['title']))
    story.append(Spacer(1, 8))
    story.append(Paragraph('<b><i>To:</i></b>', s['bold_italic']))
    story.append(Paragraph(f'{sal} {name},', s['normal']))
    story.append(Paragraph(address, s['normal']))
    story.append(Paragraph(f'{pincode}.', s['normal']))
    story.append(Spacer(1, 14))
    story.append(Paragraph(f'Dear {sal} {name},', s['normal']))
    story.append(Spacer(1, 12))

    body = (
        f'We are glad to inform you that your Contract period is extended till '
        f'<b>{extended_date}</b> . '
        f'Your remuneration has been revised from '
        f'<b>INR {current_ctc_lpa} LPA to INR {increment_ctc_lpa} LPA</b> '
        f'with effect from the beginning of {effective_month}. '
        f'A detailed of salary structure is given in the annexure'
    )
    story.append(Paragraph(body, s['justify']))
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        'All the other terms and conditions of the original contract remain unchanged.',
        s['normal']
    ))
    story.append(Spacer(1, 24))
    _signature_block(story, s, hr_name=data.get('hr_name') or None, sig_src=data.get('sig_src'))

    story.append(PageBreak())
    story.append(Paragraph('<b><i>Annexure - (Salary Details)</i></b>', s['bold_italic']))
    story.append(Spacer(1, 11.7))
    story.append(_ann_info_table([
        ['Name',                f'{sal} {name}'],
        ['Designation',         designation],
        ['Cost to the company', f'{increment_ctc_lpa} LPA'],
    ]))
    story.append(Spacer(1, 65.2))
    story.append(_salary_table(ctc_monthly))

    doc.build(story)
    return buf.getvalue()



# ─────────────────────────────────────────────────────────────────────────────
# Payslip PDF generator — pixel-perfect match to DOCX reference
# Measured from word/media group shapes (EMU→pt) in Payslip_Feb__2026.docx
# ─────────────────────────────────────────────────────────────────────────────

# Page constants (Letter: 612×792 pt)
_PS_PW   = 612.0
_PS_PH   = 792.0
# Group width = 532.82pt → centered: (612 - 532.82) / 2 = 39.59pt each side
_PS_ML   = 39.59    # left margin — centers the 532.82pt table on the page
_PS_MT   = 83.0     # top margin

def _ps_img():
    """Absolute path to the payslip header image."""
    return _img_path('payslip_header.jpg')


def _rl_y(group_y, shape_h):
    """Convert group-local Y (pt, down from group top) to ReportLab Y (pt, up from page bottom)."""
    return (_PS_PH - _PS_MT) - group_y - shape_h


def _rl_x(group_x):
    """Convert group-local X to page X."""
    return _PS_ML + group_x


def _draw_text_in_box(c, x, y, w, h, text, font='Times-Roman', size=12,
                       bold=False, align='left', color=(0, 0, 0)):
    """Draw text clipped to a box. x,y,w,h all in RL page coords."""
    if not text:
        return
    fname = font
    if bold and 'Bold' not in font:
        fname = font.replace('-Roman', '-Bold').replace('-Italic', '-BoldItalic')
        if fname == font:
            fname = font + '-Bold'
    try:
        c.setFont(fname, size)
    except Exception:
        c.setFont('Times-Bold' if bold else 'Times-Roman', size)
    c.setFillColorRGB(*color)
    if align == 'right':
        c.drawRightString(x + w, y + (h - size) / 2 + 1.5, text)
    elif align == 'center':
        c.drawCentredString(x + w / 2, y + (h - size) / 2 + 1.5, text)
    else:
        c.drawString(x, y + (h - size) / 2 + 1.5, text)


def generate_payslip(data: dict) -> bytes:
    """
    Generate a payslip PDF that is pixel-perfect to the Cloud R tech DOCX template.

    Coordinates reverse-engineered from Payslip_Feb__2026.docx (EMU→pt).
    All drawing uses ReportLab canvas directly — no Platypus, no added colour.
    """
    buf = BytesIO()
    c   = canvas.Canvas(buf, pagesize=(_PS_PW, _PS_PH))
    c.setTitle('Pay Slip')

    # ── Salary calculation — percentages verified from Payslip_April__2026.docx ──
    # Gross 37500 → Basic 15000(40%) HRA 7500(20%) Med 3750(10%)
    #               Perf 5625(15%)  Conv 3750(10%) Mgmt 1875(5%)  Total=37500
    # Other Deduction = Basic / 12  (EPF employee share — confirmed 1250.00 exact)
    # Professional Tax = 250.00 fixed
    gross     = float(data.get('gross_salary', 0))
    basic     = round(gross * 0.40)
    hra       = round(gross * 0.20)
    med       = round(gross * 0.10)
    perf      = round(gross * 0.15)
    conv      = round(gross * 0.10)
    mgmt      = int(gross) - (basic + hra + med + perf + conv)
    total_earn = int(gross)
    other_ded   = round(basic / 12)   # EPF = Basic ÷ 12 (8.33…%)
    pt_tax      = 250                 # Professional Tax fixed
    total_ded   = other_ded + pt_tax
    net_pay     = total_earn - total_ded
    rounding_adj = 0                  # absorbs any residual rounding difference

    def rs(v):
        return f"Rs.{v:,.0f}"

    wdays  = str(data.get('working_days', ''))
    phol   = str(data.get('paid_holiday', '0'))

    # ── Header image — exact position from DOCX (pic:pic xfrm, EMU→pt) ─────────
    # off x=4457/12700=0.35pt  y=0  →  group-local top-left
    # ext cx=6761848/12700=532.43pt  cy=611759/12700=48.17pt
    # Group origin: page_x=_PS_ML=18pt, group_top=_PS_PH-_PS_MT=709pt (RL)
    img_path = _ps_img()
    if img_path:
        IMG_X = _PS_ML           # match table left edge
        IMG_W = 532.82           # match table width (left+right edges)
        # Anchor: bottom edge of the banner stays fixed at the original
        # position (= top of the table below it), regardless of height,
        # so a taller banner grows UPWARD into the top margin instead of
        # overlapping the table.
        IMG_BOTTOM = (_PS_PH - _PS_MT) - 48.17   # 660.83pt from page bottom (fixed anchor)
        # Size the banner to its true aspect ratio (no stretching), capped
        # so it never eats into the top page margin.
        try:
            from PIL import Image as _PILImage
            with _PILImage.open(img_path) as _im:
                _iw, _ih = _im.size
            IMG_H = IMG_W * (_ih / _iw)
        except Exception:
            IMG_H = 88.5
        IMG_H = min(IMG_H, _PS_PH - IMG_BOTTOM - 12)   # leave >=12pt top margin
        IMG_Y = IMG_BOTTOM
        c.drawImage(img_path, IMG_X, IMG_Y, width=IMG_W, height=IMG_H,
                    preserveAspectRatio=False, mask='auto')
        # Banner outline — matches the table border style
        c.saveState()
        c.setStrokeColorRGB(0, 0, 0)
        c.setLineWidth(0.5)
        c.rect(IMG_X, IMG_Y, IMG_W, IMG_H, stroke=1, fill=0)
        c.restoreState()

    # ─────────────────────────────────────────────────────────────────────────
    # All shape coordinates below are (group_x, group_y) from the DOCX measurement.
    # Convert with _rl_x() / _rl_y() before drawing.
    # Black = (0,0,0). No extra colours.
    # ─────────────────────────────────────────────────────────────────────────

    BLACK = (0, 0, 0)
    TN    = 'Times-Roman'
    TNB   = 'Times-Bold'

    def line(gx, gy, gw, gh):
        """Draw a 0.1pt rule (the DOCX uses pairs of 1pt+0.1pt — we use the thinner one)."""
        px = _rl_x(gx)
        py = _rl_y(gy, gh)
        c.saveState()
        c.setStrokeColorRGB(0, 0, 0)
        c.setLineWidth(0.5)
        if gw > gh:          # horizontal
            c.line(px, py + gh / 2, px + gw, py + gh / 2)
        else:                # vertical
            c.line(px + gw / 2, py, px + gw / 2, py + gh)
        c.restoreState()

    def txt(gx, gy, gw, gh, text, font=TN, size=12, bold=False, align='left'):
        _draw_text_in_box(c,
                          _rl_x(gx), _rl_y(gy, gh), gw, gh,
                          text, font=font, size=size, bold=bold,
                          align=align, color=BLACK)

    # ════════════════════════════════════════════════════════════════════════
    # COLUMN X-POSITIONS (group-local, from DOCX EMU measurement)
    # ════════════════════════════════════════════════════════════════════════
    # Emp block — 4 original DOCX cols:
    #   EmpNo  : x=  1.92 → 71.42   (w= 69.50)
    #   Name   : x= 71.895→238.045  (w=166.15)
    #   Desig  : x=238.005→454.905  (w=216.90)
    #   Category: x=454.895→532.295 (w= 77.40)
    # Row-2 of emp block:
    #   Sex    : x=238.005 (below Designation)
    #   DOJ    : x=321.065 (original — has inner divider at x=320.59)
    #
    # STEP CHANGES:
    #   Step 1 — add vertical after Designation    → x=454.42  row1 only (y=70.08 h=29.04)
    #   Step 2 — remove box after Category         → no vertical at x=454.42 in row2
    #   Step 3 — Sex straight down Designation     → Sex keeps x=238; remove x=71.42/x=237.53 from emp block
    #   Step 4 — DOJ one box, no inner divider     → remove x=320.59 from row2; DOJ spans x=238..532
    #   Step 5 — vertical line closing net pay box → x=531.82 y=258.87 h=29.05
    #   Step 6 — NOTE left-aligned bold
    #   Step 7 — consistent PAD=6, centered title, clean spacing
    #
    # Resulting emp block layout:
    #   Row 1 (y=70.08→99.12): EmpNo | Name | Designation | Category
    #     Verticals: x=71.42 (after EmpNo), x=238.00 (after Name), x=454.42 (after Designation=Step1)
    #     NO vertical after Category (Step 2 — outer box is right edge)
    #   Row 2 (y=99.12→113.64): [EmpNo+Name span] | Sex | Date of Joining (one wide box)
    #     Verticals: x=238.00 only (left of Sex; Step 3 — Sex below Designation)
    #     NO x=320.59 (Step 4 — DOJ is one box from x=238 right of Sex to outer edge)
    #     NO x=454.42 (Step 2)
    # ════════════════════════════════════════════════════════════════════════

    PAD   = 8.0    # inner left padding — all cells
    FSZ   = 10     # body text size
    LFSZ  = 9      # label size

    # ── Outer box — full rectangle, corners closed ──────────────────────────
    # Use canvas rect() for pixel-perfect connected corners
    # Start box at y=49.50 (title bar top) so the top edge doesn't cut through the header banner
    BOX_X, BOX_Y, BOX_W, BOX_H = 0, 49.50, 532.82, 224.37
    c.saveState()
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.5)
    c.rect(_rl_x(BOX_X), _rl_y(BOX_Y, BOX_H), BOX_W, BOX_H, stroke=1, fill=0)
    c.restoreState()
    # Title bar divider at y=49.50 — below the header image (bottom at y=48.17)
    line(0.00, 49.50, 532.82, 0.01)

    # ── Horizontal lines — full width from x=0 to x=532.82 ──────────────────
    FULL_W = 532.82
    for gy in [
         70.08,   # emp block top
         99.12,   # emp row1 / row2
        113.64,   # emp block bottom
        128.16,   # below sub-headers
        142.68,   # Basic
        157.20,   # HRA
        171.75,   # Medical
        186.27,   # Perf Bonus
        200.79,   # Conveyance
        215.31,   # Management
        229.83,   # Rounding Adjustment
        244.35,   # Totals top
        258.87,   # Totals bottom / Net Pay top
        273.87,   # Net Pay bottom — table final edge
    ]:
        line(0.00, gy, FULL_W, 1.00)

    # ── Verticals — emp block ─────────────────────────────────────────────────
    line( 71.42,  70.08,  1.00,  29.04)   # after Employee No   (row1 only)
    line(238.00,  70.08,  1.00,  29.04)   # after Name          (row1 only)
    line(454.42,  70.08,  1.00,  29.04)   # after Designation   (row1 only)
    line(238.00,  99.12,  1.00,  14.52)   # left of Sex         (row2)
    line(360.00,  99.12,  1.00,  14.52)   # after Sex | DOJ     (row2)

    # ── Verticals — salary table ──────────────────────────────────────────────
    line(103.70, 114.61,  1.00, 159.77)   # Attendance | Earnings
    line( 71.42, 129.13,  1.00, 145.22)   # att value sub-col
    line(237.53, 114.61,  1.00, 159.79)   # Earnings | Amount
    line(320.59, 114.61,  1.00, 159.79)   # Amount | Deductions
    line(450.39, 128.16,  1.00, 130.71)   # Deductions Particulars | Amount

    # ── Verticals — Net Pay box ───────────────────────────────────────────────
    line(  0.50, 258.87,  1.00,  15.00)   # left (aligns with grid x=0)
    line(532.32, 258.87,  1.00,  15.00)   # right (aligns with grid x=532.82)

    # ── Title ────────────────────────────────────────────────────────────────
    month_year = data.get('month_year', '')
    txt(1.92, 50.65, 530.90, 14.65,
        f'Pay Slip for the month of {month_year}',
        font=TNB, size=12, bold=True, align='center')

    # ── Emp block row 1 (y=70.08 → 99.12, h=29.04) ───────────────────────────
    # Split into: label band top 13pt, value band bottom 16pt
    LY, LH = 70.08, 13.00   # label band
    VY, VH = 83.08, 16.04   # value band

    txt( 1.92 + PAD, LY,  61.50, LH, 'Employee No',               font=TNB, size=LFSZ, bold=True)
    txt( 1.92 + PAD, VY,  61.50, VH, data.get('emp_no', ''),       font=TN,  size=FSZ)

    txt(71.42 + PAD, LY, 158.58, LH, 'Name',                       font=TNB, size=LFSZ, bold=True)
    txt(71.42 + PAD, VY, 158.58, VH, data.get('name', ''),          font=TN,  size=FSZ)

    txt(238.00 + PAD, LY, 208.42, LH, 'Designation',               font=TNB, size=LFSZ, bold=True)
    txt(238.00 + PAD, VY, 208.42, VH, data.get('designation', ''),  font=TN,  size=FSZ)

    txt(454.42 + PAD, LY,  69.40, LH, 'Category',                  font=TNB, size=LFSZ, bold=True)
    txt(454.42 + PAD, VY,  69.40, VH, data.get('category', ''),     font=TN,  size=FSZ)

    # ── Emp block row 2 (y=99.12 → 113.64, h=14.52) ──────────────────────────
    # Single-line row — text centred vertically by _draw_text_in_box
    R2Y, R2H = 99.12, 14.52

    # Sex cell (x=238..360, w=122)
    txt(238.00 + PAD, R2Y,  22.00, R2H, 'Sex:',                     font=TNB, size=LFSZ, bold=True)
    txt(238.00 + 32,  R2Y,  80.00, R2H, data.get('sex', ''),        font=TN,  size=FSZ)

    # DOJ cell (x=360..531.82, w=171.82) — label 74pt, value fills rest
    txt(360.00 + PAD, R2Y,  74.00, R2H, 'Date of Joining:',          font=TNB, size=LFSZ, bold=True)
    txt(360.00 + PAD + 76, R2Y, 87.82, R2H, data.get('date_of_joining', ''), font=TN, size=FSZ)

    # ── Salary header row (y=113.64 → 128.16, h=14.52) ───────────────────────
    ROW_H = 14.52
    txt(  1.92 + PAD, 113.64, 100.00, ROW_H, 'Working Days', font=TNB, size=FSZ, bold=True)
    txt(103.70,        113.64, 133.83, ROW_H, 'Earnings',     font=TNB, size=FSZ, bold=True, align='center')
    txt(320.59,        113.64, 211.24, ROW_H, 'Deductions',   font=TNB, size=FSZ, bold=True, align='center')

    # ── Sub-header row (y=128.16 → 142.68, h=14.52) ──────────────────────────
    txt(  1.92 + PAD, 128.16,  63.50, ROW_H, 'Attendance',  font=TNB, size=FSZ, bold=True)
    txt(103.70 + PAD, 128.16, 125.83, ROW_H, 'Particulars', font=TNB, size=FSZ, bold=True)
    txt(237.53,        128.16,  82.50, ROW_H, 'Amount',      font=TNB, size=FSZ, bold=True, align='right')
    txt(320.59 + PAD, 128.16, 125.83, ROW_H, 'Particulars', font=TNB, size=FSZ, bold=True)
    txt(450.39,        128.16,  81.43, ROW_H, 'Amount',      font=TNB, size=FSZ, bold=True, align='right')

    # ── Data rows ────────────────────────────────────────────────────────────
    rounding_amt = rs(rounding_adj) if rounding_adj else ''
    rows = [
        (142.68, wdays,  'Basic',                rs(basic),  'Other Deduction',   rs(other_ded)),
        (157.20, phol,   'HRA',                  rs(hra),    'Professional Tax',  rs(pt_tax)),
        (171.75, '',     'Medical Allowance',    rs(med),    '',                  ''),
        (186.27, '',     'Performance Bonus',    rs(perf),   '',                  ''),
        (200.79, '',     'Conveyance',           rs(conv),   '',                  ''),
        (215.31, '',     'Management Allowance', rs(mgmt),   '',                  ''),
        (229.83, '',     'Rounding Adjustment',  rounding_amt, '',                ''),
    ]
    att_labels = {142.68: 'Working Days', 157.20: 'Paid Holiday'}

    for gy, att, elabel, eamt, dlabel, damt in rows:
        if gy in att_labels:
            txt(1.92 + PAD,   gy,  63.50, ROW_H, att_labels[gy], font=TN, size=FSZ)
        if att:
            txt(71.42,         gy,  30.28, ROW_H, att,            font=TN, size=FSZ, align='right')
        txt(103.70 + PAD,     gy, 125.83, ROW_H, elabel,          font=TN, size=FSZ)
        if eamt:
            txt(237.53,        gy,  82.50, ROW_H, eamt,           font=TN, size=FSZ, align='right')
        if dlabel:
            txt(320.59 + PAD,  gy, 125.83, ROW_H, dlabel,         font=TN, size=FSZ)
        if damt:
            txt(450.39,        gy,  81.43, ROW_H, damt,           font=TN, size=FSZ, align='right')

    # ── Totals row (y=244.35 → 258.87, h=14.52) ──────────────────────────────
    txt(1.92 + PAD,   244.35,  63.50, ROW_H, 'Total Days',        font=TNB, size=FSZ, bold=True)
    txt(71.42,         244.35,  30.28, ROW_H, wdays,               font=TNB, size=FSZ, bold=True, align='right')
    txt(103.70 + PAD,  244.35, 125.83, ROW_H, 'Total Earnings',   font=TNB, size=FSZ, bold=True)
    txt(237.53,        244.35,  82.50, ROW_H, rs(total_earn),      font=TNB, size=FSZ, bold=True, align='right')
    txt(320.59 + PAD,  244.35, 125.83, ROW_H, 'Total Deductions', font=TNB, size=FSZ, bold=True)
    txt(450.39,        244.35,  81.43, ROW_H, rs(total_ded),       font=TNB, size=FSZ, bold=True, align='right')

    # ── Net Pay row (y=258.87 → 273.87, h=15.00) ─────────────────────────────
    NET_H = 15.00
    txt(1.92 + PAD,  258.87, 444.49, NET_H, 'Net Pay:',      font=TNB, size=FSZ, bold=True, align='right')
    txt(450.39,       258.87,  81.43, NET_H, rs(net_pay),    font=TNB, size=FSZ, bold=True, align='right')

    # ── NOTE — left-aligned, bold ─────────────────────────────────────────────
    c.setFont(TNB, 9)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(_rl_x(1.92), _rl_y(280, 9),
        'NOTE: This is a computer generated copy and it does not need any seal or signature.')

    c.showPage()
    c.save()
    return buf.getvalue()