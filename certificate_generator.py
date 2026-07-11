#!/usr/bin/env python3
"""
HumanX Tech'26 × University of Moratuwa
Certificate Generator — Dark Edition

Usage:
    pip install reportlab pandas openpyxl Pillow
    python certificate_generator.py participants.xlsx
    python certificate_generator.py participants.csv

Outputs PDF files into the  certificates/  folder.

Optional TTF fonts (place next to this script for best results):
    Montserrat-Regular.ttf, Montserrat-SemiBold.ttf, Montserrat-Bold.ttf
    CormorantGaramond-Regular.ttf, CormorantGaramond-Bold.ttf,
    CormorantGaramond-Italic.ttf
Falls back to Helvetica / Times-Roman if TTFs are absent.
"""

import os, sys, math, hashlib, csv
from datetime import datetime
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── Page ──────────────────────────────────────────────────────────────────────
W, H = landscape(A4)   # 841.89 × 595.28 pt

# ── Config ────────────────────────────────────────────────────────────────────
EVENT_NAME  = "HumanX Tech'26"
CONF_NAME   = "HR & Tech Conference 2026"
EVENT_DATE  = "12 July 2026"
EVENT_LOC   = "Colombo, Sri Lanka"
CEO_TITLE   = "Chief Executive Officer"
HOD_TITLE   = "Head of the Department"
HOD_DEPT    = "Dept. of Computer Science & Engineering"
COORD_TITLE = "MBA in IT Coordinator"
VERIFY_URL  = "humanxtech.cse.uom.lk"
ORG_LINES   = [
    "Department of Computer Science & Engineering",
    "Faculty of Engineering",
    "University of Moratuwa",
]
ISSUED      = datetime.now()
UOM_LOGO    = "cse-logo.jpeg"    # must sit next to this script (CSE dept logo)
HOD_SIGNATURE = "hod-signature.png"  # HoD signature image, sits next to this script
OUTPUT_DIR  = "certificates"

CERT_LABELS = {
    "participation": "CERTIFICATE OF PARTICIPATION",
    "speaker":       "CERTIFICATE OF RECOGNITION",
    "organiser":     "CERTIFICATE OF APPRECIATION",
    "winner":        "CERTIFICATE OF EXCELLENCE",
}

# ── Colours ───────────────────────────────────────────────────────────────────
def rgba(h, a=1.0):
    c = colors.HexColor(h)
    return colors.Color(c.red, c.green, c.blue, alpha=a)

C_BG_1      = colors.HexColor('#7E4AD8')   # lighter purple glow centre
C_BG_2      = colors.HexColor('#4C1E8F')   # rich purple mid
C_BG_3      = colors.HexColor('#250642')   # deep violet edge
C_PURPLE    = rgba('#8D72FF')
C_PLT       = rgba('#A891FF')    # purple-light
C_PDK       = rgba('#5A31D6')    # purple-dark
C_OFFWHITE  = rgba('#BBAEFF')
C_WHITE     = colors.white
C_TEXT_SOFT = rgba('#DCD2FF')    # light lavender body text
C_TEXT_DIM  = rgba('#BDB0EC')    # dimmer lavender for meta text

def tracking(text, gap=' '):
    return gap.join(str(text))

# ── Font registration ─────────────────────────────────────────────────────────
def _try_register(alias, filename):
    if os.path.exists(filename):
        try:
            pdfmetrics.registerFont(TTFont(alias, filename))
            return True
        except Exception:
            pass
    return False

def register_fonts():
    F = {}
    F['sans']        = 'Helvetica'        if not _try_register('Mont',    'Montserrat-Regular.ttf')   else 'Mont'
    F['sans-semi']   = 'Helvetica-Bold'   if not _try_register('MontSemi','Montserrat-SemiBold.ttf')  else 'MontSemi'
    F['sans-bold']   = 'Helvetica-Bold'   if not _try_register('MontBold','Montserrat-Bold.ttf')      else 'MontBold'
    F['serif']       = 'Times-Roman'      if not _try_register('Corm',    'CormorantGaramond-Regular.ttf') else 'Corm'
    F['serif-bold']  = 'Times-Bold'       if not _try_register('CormBold','CormorantGaramond-Bold.ttf')   else 'CormBold'
    F['serif-ital']  = 'Times-Italic'     if not _try_register('CormItal','CormorantGaramond-Italic.ttf') else 'CormItal'
    return F

# ── Network watermark ─────────────────────────────────────────────────────────
# Positions in 1122×793 SVG space → scaled to A4 landscape
SVG_W, SVG_H = 1122, 793
SX, SY = W / SVG_W, H / SVG_H

def _sx(x): return x * SX
def _sy(y): return H - y * SY

_NODES = [
    (90,55),(300,85),(530,52),(760,78),(990,58),(1090,155),
    (170,165),(390,145),(610,175),(840,155),(1090,155),
    (85,290),(290,265),(490,305),(715,278),(945,292),(1085,270),
    (195,405),(415,385),(645,425),(865,400),(1068,418),
    (75,525),(305,548),(535,515),(770,542),(998,524),
    (155,658),(375,645),(595,672),(808,658),(1048,648),
    (88,752),(295,742),(515,762),(738,748),(958,760),(1088,742),
]

_EDGES = [
    (0,1),(1,2),(2,3),(3,4),(4,5),
    (6,1),(6,7),(7,2),(7,8),(8,3),(8,9),(9,4),(9,5),
    (0,6),(0,11),
    (11,6),(11,12),(12,7),(12,13),(13,8),(13,14),(14,9),(14,15),(15,5),(15,16),
    (11,17),(12,17),(17,18),(18,13),(18,19),(19,14),(19,20),(20,15),(20,21),
    (22,17),(22,23),(23,18),(23,24),(24,19),(24,25),(25,20),(25,26),
    (22,27),(23,28),(28,24),(28,29),(29,25),(29,30),(30,26),(30,31),
    (27,28),(27,32),(32,33),(33,28),(33,34),(34,29),(34,35),(35,30),(35,36),(36,26),(36,37),
    (2,13),(3,14),(5,16),(16,21),(21,26),(26,31),(31,37),
]

def draw_watermark(c):
    c.saveState()
    c.setStrokeColor(colors.Color(0.70, 0.61, 1.0, alpha=0.025))
    c.setLineWidth(0.45)
    for a, b in _EDGES:
        if a < len(_NODES) and b < len(_NODES):
            ax, ay = _NODES[a]; bx, by = _NODES[b]
            c.line(_sx(ax), _sy(ay), _sx(bx), _sy(by))
    c.setStrokeColor(colors.Color(0,0,0,0))
    for nx, ny in _NODES:
        c.setFillColor(colors.Color(0.78, 0.71, 1.0, alpha=0.07))
        c.circle(_sx(nx), _sy(ny), 1.35, fill=1, stroke=0)
    c.restoreState()

# ── Background ────────────────────────────────────────────────────────────────
def _lerp_color(c1, c2, t):
    t = max(0.0, min(1.0, t))
    return colors.Color(
        c1.red   + (c2.red   - c1.red)   * t,
        c1.green + (c2.green - c1.green) * t,
        c1.blue  + (c2.blue  - c1.blue)  * t,
    )

def draw_background(c):
    """Approximate CSS radial-gradient(45% 30%, C1 0%, C2 42%, C3 100%)
    with concentric rings drawn edge-color-first."""
    cx, cy = W * 0.48, H * 0.64
    max_r  = math.hypot(max(cx, W-cx), max(cy, H-cy))

    c.setFillColor(C_BG_3)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    steps = 42
    for i in range(steps, 0, -1):
        t = i / steps
        r = max_r * t
        if t > 0.42:
            local = (t - 0.42) / (1 - 0.42)
            col = _lerp_color(C_BG_2, C_BG_3, local)
        else:
            local = t / 0.42
            col = _lerp_color(C_BG_1, C_BG_2, local)
        c.setFillColor(col)
        c.ellipse(cx - r, cy - r, cx + r, cy + r, fill=1, stroke=0)

    c.setFillColor(colors.Color(0.02, 0.00, 0.12, alpha=0.16))
    c.rect(0, 0, W, H, fill=1, stroke=0)

# ── Borders ───────────────────────────────────────────────────────────────────
def draw_borders(c):
    i1, i2 = 14*mm, 28*mm
    c.setStrokeColor(colors.Color(0.72, 0.64, 1.0, alpha=0.72))
    c.setLineWidth(0.75)
    c.roundRect(i1, i1, W-2*i1, H-2*i1, 8, fill=0, stroke=1)
    c.setStrokeColor(colors.Color(0.72, 0.64, 1.0, alpha=0.18))
    c.setLineWidth(0.4)
    c.line(i2, H-i2, W-i2, H-i2)

# ── Corner ornaments ──────────────────────────────────────────────────────────
def draw_corners(c):
    arm  = 13 * mm
    ins  = 9  * mm
    tick = arm * 0.45

    def corner(ox, oy, dx, dy):
        c.setStrokeColor(C_PURPLE); c.setLineWidth(1.2)
        c.line(ox, oy, ox + arm*dx, oy)
        c.line(ox, oy, ox, oy + arm*dy)
        c.setStrokeColor(colors.Color(0.482,0.424,0.961,alpha=0.38))
        c.setLineWidth(0.5)
        c.line(ox, oy + tick*dy*1.4, ox, oy + (arm - tick*0.25)*dy)
        c.line(ox + tick*dx*1.4, oy, ox + (arm - tick*0.25)*dx, oy)

    corner(ins,   H-ins, +1, -1)  # TL
    corner(W-ins, H-ins, -1, -1)  # TR
    corner(ins,   ins,   +1, +1)  # BL
    corner(W-ins, ins,   -1, +1)  # BR

# ── Diamond divider ───────────────────────────────────────────────────────────
def draw_divider(c, cx, y, width, alpha=0.35):
    s = 3.2
    c.setStrokeColor(colors.Color(0.482, 0.424, 0.961, alpha=alpha))
    c.setLineWidth(0.55)
    c.line(cx - width/2, y, cx - s*2, y)
    c.line(cx + s*2, y,    cx + width/2, y)
    p = c.beginPath()
    p.moveTo(cx, y+s); p.lineTo(cx+s, y); p.lineTo(cx, y-s); p.lineTo(cx-s, y)
    p.close()
    c.setFillColor(colors.Color(0.482, 0.424, 0.961, alpha=alpha+0.15))
    c.drawPath(p, fill=1, stroke=0)

def draw_meta_icon(c, x, y, kind):
    c.saveState()
    c.setStrokeColor(colors.Color(0.73, 0.66, 1.0, alpha=0.76))
    c.setFillColor(colors.Color(0.73, 0.66, 1.0, alpha=0.0))
    c.setLineWidth(0.7)
    if kind == 'calendar':
        c.roundRect(x, y, 8, 8, 1.2, fill=0, stroke=1)
        c.line(x, y + 5.7, x + 8, y + 5.7)
        c.line(x + 2.1, y + 9.4, x + 2.1, y + 6.8)
        c.line(x + 5.9, y + 9.4, x + 5.9, y + 6.8)
    elif kind == 'pin':
        c.circle(x + 4, y + 5.6, 3.3, fill=0, stroke=1)
        c.circle(x + 4, y + 5.6, 1.0, fill=0, stroke=1)
        c.line(x + 1.8, y + 3.1, x + 4, y - 0.7)
        c.line(x + 6.2, y + 3.1, x + 4, y - 0.7)
    c.restoreState()

def draw_cse_logo(c, cx, top_y, height):
    if not os.path.exists(UOM_LOGO):
        return

    img = ImageReader(UOM_LOGO)
    iw, ih = img.getSize()
    width = iw * (height / ih)
    x = cx - width / 2
    y = top_y - height
    pad = 3

    c.saveState()
    c.setFillColor(colors.Color(0.015, 0.010, 0.045, alpha=1.0))
    c.roundRect(x - pad, y - pad, width + pad * 2, height + pad * 2, 1.5, fill=1, stroke=0)
    c.drawImage(img, x, y, width, height, mask='auto')
    c.restoreState()

# ── Seal ──────────────────────────────────────────────────────────────────────
def draw_seal(c, cx, cy, F):
    r = 30
    c.setFillColor(colors.Color(0.482, 0.424, 0.961, alpha=0.10))
    c.setStrokeColor(colors.Color(0.659, 0.612, 0.973, alpha=0.62))
    c.setLineWidth(0.7)
    c.circle(cx, cy, r, fill=1, stroke=1)
    c.setStrokeColor(colors.Color(0.482, 0.424, 0.961, alpha=0.26))
    c.setLineWidth(0.35)
    c.circle(cx, cy, r*0.83, fill=0, stroke=1)

    # Cardinal dots
    for ang in range(0, 360, 45):
        rad = math.radians(ang)
        c.setFillColor(colors.Color(0.659, 0.612, 0.973, alpha=0.30))
        c.circle(cx + r*math.cos(rad), cy + r*math.sin(rad), 1.1, fill=1, stroke=0)

    # X mark
    s = r * 0.33
    c.setLineCap(1)
    c.setLineWidth(2.2)
    c.setStrokeColor(C_PLT)
    c.line(cx-s, cy+s, cx, cy); c.line(cx, cy, cx-s, cy-s)
    c.setStrokeColor(C_PDK)
    c.line(cx+s, cy+s, cx, cy); c.line(cx, cy, cx+s, cy-s)
    c.setLineCap(0)

    # Arc text (approximate with straight text at top)
    c.setFont(F['sans'], 5)
    c.setFillColor(colors.Color(0.659, 0.612, 0.973, alpha=0.75))
    text = "HUMANX TECH '26"
    c.saveState()
    c.translate(cx, cy)
    # Draw text along arc by rotating per character
    chars = list(text)
    n = len(chars)
    arc_span = math.radians(160)
    start_ang = math.radians(90) + arc_span/2
    char_w = r * 0.83 * arc_span / n
    for i, ch in enumerate(chars):
        ang = start_ang - i * (arc_span / (n-1))
        cr = r * 0.83
        c.saveState()
        c.rotate(math.degrees(ang) - 90)
        c.drawCentredString(0, cr - 1, ch)
        c.restoreState()
    c.restoreState()

    # "OFFICIAL SEAL" text
    c.setFont(F['sans'], 5)
    c.setFillColor(colors.Color(0.659, 0.612, 0.973, alpha=0.35))
    c.drawCentredString(cx, cy - r*0.45, "OFFICIAL SEAL")

# ── Certificate ID box ────────────────────────────────────────────────────────
def draw_cert_id_box(c, x, y, cid, F):
    c.setFont(F['sans-semi'], 5.5)
    c.setFillColor(colors.Color(0.76, 0.69, 1.0, alpha=0.74))
    c.drawString(x, y + 16, "CERTIFICATE ID")
    c.setFont(F['sans'], 8)
    c.setFillColor(C_WHITE)
    c.drawString(x, y + 6, cid)
    c.setFont(F['sans'], 5)
    c.setFillColor(C_TEXT_DIM)
    c.drawString(x, y - 5, f"verify at {VERIFY_URL}")

# ── Gradient underline ────────────────────────────────────────────────────────
def draw_underline(c, cx, y, width=90*mm):
    steps = 40
    half  = width / 2
    for i in range(steps):
        t  = i / steps
        t2 = (i+1) / steps
        xA = cx - half + t  * width
        xB = cx - half + t2 * width
        # alpha: 0 at edges, 1 at centre
        a = math.sin(math.pi * (t + t2) / 2)
        c.setStrokeColor(colors.Color(0.70, 0.40, 1.0, alpha=a * 0.72))
        c.setLineWidth(0.9)
        c.line(xA, y, xB, y)

    c.setStrokeColor(colors.Color(1.0, 0.90, 1.0, alpha=0.85))
    c.setLineWidth(1.1)
    c.line(cx - 8*mm, y, cx + 8*mm, y)

# ── Main certificate draw ─────────────────────────────────────────────────────
def draw_certificate(c, row, F):
    name      = str(row.get('name',         'Participant')).strip()
    job_title = str(row.get('job_title',    '')).strip()
    org       = str(row.get('organization', '')).strip()
    cert_type = str(row.get('cert_type',    'participation')).strip().lower()
    c_label   = CERT_LABELS.get(cert_type, CERT_LABELS['participation'])
    detail    = '   ·   '.join(x for x in [job_title, org] if x and x.lower() != 'nan')
    h_cid     = hashlib.md5(f"{name}{ISSUED.date()}".encode()).hexdigest().upper()
    cid       = f"CERT-{ISSUED.strftime('%Y%m%d')}-{h_cid[:8]}"

    draw_background(c)
    draw_watermark(c)
    draw_borders(c)
    draw_corners(c)

    # ── HEADER ────────────────────────────────────────────────────────────────
    hdr_top = H - 36*mm

    # CSE dept logo (organiser)
    logo_h_pt = 22 * mm
    draw_cse_logo(c, W/2, hdr_top, logo_h_pt)

    # 3-line organisational hierarchy under the logo
    org_y = hdr_top - logo_h_pt - 16
    c.setFont(F['sans-semi'], 5.9)
    c.setFillColor(colors.Color(0.86, 0.82, 1.0, alpha=0.95))
    for line in ORG_LINES:
        c.drawCentredString(W/2, org_y, line.upper())
        org_y -= 7.0

    # "Presents"
    pres_y = org_y - 8
    c.setFont(F['serif-ital'], 9)
    c.setFillColor(colors.Color(0.9, 0.784, 0.980, alpha=0.9))
    c.drawCentredString(W/2, pres_y, 'Presents')

    # HumanX Tech'26 wordmark
    hx_y  = pres_y - 20
    c.setFont(F['sans-bold'], 16)
    c.setFillColor(C_WHITE)
    c.drawCentredString(W/2, hx_y, EVENT_NAME)

    # Certificate type label
    c.setFont(F['sans-semi'], 8)
    c.setFillColor(colors.Color(0.8, 0.70, 1.0, alpha=0.92))
    c.drawCentredString(W/2, hx_y - 16, tracking(c_label, '  '))

    # ── BODY ──────────────────────────────────────────────────────────────────
    body_cy = H/2 - 20     # approx vertical centre of body zone

    # "This is to certify that"
    c.setFont(F['sans'], 9)
    c.setFillColor(C_TEXT_SOFT)
    c.drawCentredString(W/2, body_cy + 52, 'This certificate is proudly presented to')

    # Recipient name — auto-size
    name_sz = 41 if len(name) > 28 else (37 if len(name) > 22 else 45)
    c.setFont(F['serif-bold'], name_sz)
    c.setFillColor(C_WHITE)
    c.drawCentredString(W/2, body_cy + 12, name)

    # Gradient underline
    draw_underline(c, W/2, body_cy + 3, width=102*mm)

    # Job title · Organisation
    if detail:
        c.setFont(F['sans-semi'], 6.4)
        c.setFillColor(C_TEXT_SOFT)
        c.drawCentredString(W/2, body_cy - 20, tracking(detail.upper(), ' '))

    # "has successfully participated in"
    c.setFont(F['sans'], 8.2)
    c.setFillColor(C_TEXT_SOFT)
    c.drawCentredString(W/2, body_cy - 42, 'has successfully participated in')

    # Event pill badge
    pill_w = c.stringWidth(CONF_NAME, F['sans-semi'], 8) + 50
    pill_h = 18; pill_y = body_cy - 69; pill_x = W/2 - pill_w/2
    for offset, alpha in [(3.0, 0.08), (1.6, 0.14)]:
        c.setFillColor(colors.Color(0.72, 0.47, 1.0, alpha=alpha))
        c.roundRect(pill_x - offset, pill_y - offset, pill_w + offset*2, pill_h + offset*2, pill_h/2 + offset, fill=1, stroke=0)
    c.setFillColor(colors.Color(0.35, 0.16, 0.82, alpha=0.96))
    c.setStrokeColor(colors.Color(0.67, 0.54, 1.0, alpha=0.52))
    c.setLineWidth(0.6)
    c.roundRect(pill_x, pill_y, pill_w, pill_h, pill_h/2, fill=1, stroke=1)
    draw_meta_icon(c, pill_x + 17, pill_y + 4.2, 'calendar')
    c.setFont(F['sans-semi'], 8)
    c.setFillColor(C_WHITE)
    c.drawCentredString(W/2 + 8, pill_y + 5.6, CONF_NAME)

    # Date · Location
    meta_y = body_cy - 90
    c.setFont(F['sans'], 5.8)
    c.setFillColor(C_TEXT_DIM)
    date_text = EVENT_DATE.upper()
    loc_text = EVENT_LOC.upper()
    date_w = c.stringWidth(date_text, F['sans'], 5.8)
    loc_w = c.stringWidth(loc_text, F['sans'], 5.8)
    group_w = 8 + 6 + date_w + 17 + 8 + 6 + loc_w
    gx = W/2 - group_w/2
    draw_meta_icon(c, gx, meta_y - 2.2, 'calendar')
    c.drawString(gx + 14, meta_y, date_text)
    c.setStrokeColor(colors.Color(0.73, 0.66, 1.0, alpha=0.5))
    c.line(gx + 14 + date_w + 8, meta_y - 2.2, gx + 14 + date_w + 8, meta_y + 6.2)
    draw_meta_icon(c, gx + 14 + date_w + 18, meta_y - 2.0, 'pin')
    c.drawString(gx + 14 + date_w + 32, meta_y, loc_text)

    # ── FOOTER ────────────────────────────────────────────────────────────────
    foot_top = 35*mm

    # Cert ID (bottom-left)
    draw_cert_id_box(c, 24*mm, 23*mm, cid, F)

    # Three signatures (bottom-centre), evenly spaced across the footer.
    sig_y_line = 32*mm
    sig_gap    = 52*mm
    sig_centres = [W/2 - sig_gap, W/2, W/2 + sig_gap]

    # HoD signature image (if present) sits just above the signature line
    # if os.path.exists(HOD_SIGNATURE):
    #     sig_img = ImageReader(HOD_SIGNATURE)
    #     siw, sih = sig_img.getSize()
    #     sig_h_pt = 16*mm
    #     sig_w_pt = siw * (sig_h_pt / sih)
    #     c.drawImage(sig_img, sig_centres[0] - sig_w_pt/2, sig_y_line + 1,
    #                 sig_w_pt, sig_h_pt, mask='auto')

    for cx, title, subtitle in [
        (sig_centres[0], "Prof. Shantha Fernando", "Head of the Department"),
        (sig_centres[1], "Dr. Adeesha Wijayasiri", "MBA in IT Coordinator"),
        (sig_centres[2], "Janaka Kumarasinghe", "Conference Organizer"),
    ]:
        c.setStrokeColor(colors.Color(0.482, 0.424, 0.961, alpha=0.45))
        c.setLineWidth(0.5)
        c.line(cx - 16*mm, sig_y_line, cx + 16*mm, sig_y_line)
        c.setFont(F['serif-bold'], 8.2)
        c.setFillColor(C_WHITE)
        c.drawCentredString(cx, sig_y_line - 9.5, title)
        c.setFont(F['sans'], 5.5)
        c.setFillColor(C_TEXT_SOFT)
        c.drawCentredString(cx, sig_y_line - 17, subtitle)

    # Seal (bottom-right)
    draw_seal(c, W - 36*mm, 33*mm, F)

    # Re-draw the organiser logo last so watermark/border effects never sit above it.
    draw_cse_logo(c, W/2, hdr_top, logo_h_pt)

    # Bottom strip
    c.setFont(F['sans'], 5)
    c.setFillColor(colors.Color(0.81, 0.77, 0.95, alpha=0.72))
    c.drawCentredString(W/2, 10*mm,
        f"{EVENT_NAME}   ·   {VERIFY_URL}   ·   "
        "This certificate is issued as an official recognition of achievement".upper())

# ── Read input file ───────────────────────────────────────────────────────────
def _normalise(row: dict) -> dict:
    """Lower-case + underscore column names; fill missing optional cols."""
    out = {k.strip().lower().replace(' ', '_'): (v or '') for k, v in row.items()}
    for col, default in [('job_title',''), ('organization',''), ('cert_type','participation')]:
        out.setdefault(col, default)
    return out

def read_participants(filepath: str) -> list:
    ext = os.path.splitext(filepath)[1].lower()
    rows = []

    if ext == '.csv':
        with open(filepath, newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(_normalise(row))

    elif ext in ('.xlsx', '.xls'):
        try:
            import openpyxl
        except ImportError:
            sys.exit("openpyxl not found. Run: pip install openpyxl")
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        headers = None
        for r in ws.iter_rows(values_only=True):
            if headers is None:
                headers = [str(c).strip() if c is not None else '' for c in r]
            else:
                row = {headers[i]: (str(v).strip() if v is not None else '') for i, v in enumerate(r)}
                rows.append(_normalise(row))
        wb.close()

    else:
        sys.exit(f"Unsupported format: {ext}. Use .csv or .xlsx")

    if not rows:
        sys.exit("No data rows found in the file.")
    if 'name' not in rows[0]:
        sys.exit(f"Missing 'name' column. Found: {list(rows[0].keys())}")
    return rows

# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("Usage: python certificate_generator.py participants.xlsx")
        print("       python certificate_generator.py participants.csv")
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        sys.exit(f"File not found: {filepath}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    F = register_fonts()
    rows = read_participants(filepath)

    print(f"Generating {len(rows)} certificate(s)...")
    for i, row in enumerate(rows, 1):
        name     = str(row.get('name', f'participant_{i}')).strip()
        safe     = "".join(ch if ch.isalnum() or ch in ' -_' else '' for ch in name).strip()
        out_path = os.path.join(OUTPUT_DIR, f"{safe}.pdf")

        c = pdf_canvas.Canvas(out_path, pagesize=landscape(A4))
        draw_certificate(c, row, F)
        c.save()
        print(f"  [{i}/{len(rows)}] {out_path}")

    print(f"\nDone! Certificates saved to ./{OUTPUT_DIR}/")

if __name__ == '__main__':
    main()
