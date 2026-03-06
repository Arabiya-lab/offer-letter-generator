"""
Reintenspark Technology – Offer Letter PDF Generator
Exactly matches the uploaded template.
"""
import io as _io
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPM
from PIL import Image as PILImage

ASSETS_DIR = Path(__file__).parent / 'static' / 'assets'
W, H = A4

GREEN      = colors.HexColor('#5CB85C')
BLACK      = colors.black
WHITE      = colors.white
DARK_GRAY  = colors.HexColor('#333333')
LIGHT_GRAY = colors.HexColor('#888888')

def _pt(v): return v * mm

def _gen_qr_ir(url):
    """Generate QR code as ImageReader using ReportLab's built-in QR encoder."""
    qr = QrCodeWidget(url)
    b  = qr.getBounds()
    bw, bh = b[2]-b[0], b[3]-b[1]
    d = Drawing(bw, bh, transform=[1,0,0,1,-b[0],-b[1]])
    d.add(qr)
    png_bytes = renderPM.drawToString(d, fmt='PNG', dpi=150)
    return ImageReader(_io.BytesIO(png_bytes))

def _ir(path_or_pil):
    """Get ImageReader from file path or PIL Image."""
    if isinstance(path_or_pil, (str, Path)):
        return ImageReader(str(path_or_pil))
    buf = _io.BytesIO()
    path_or_pil.save(buf, format='PNG')
    buf.seek(0)
    return ImageReader(buf)

def _para(c, text, x, y, max_w, font='Helvetica', size=10.5, color=DARK_GRAY, lh=None):
    if lh is None: lh = size * 1.4
    c.setFont(font, size); c.setFillColor(color)
    words = text.split(); line = ''; cy = y
    for word in words:
        test = (line+' '+word).strip()
        if c.stringWidth(test, font, size) <= max_w: line = test
        else:
            if line: c.drawString(x, cy, line); cy -= lh
            line = word
    if line: c.drawString(x, cy, line); cy -= lh
    return cy

def generate_reintenspark_letter(intern, include_qr=True, base_url='http://localhost:5000'):
    buf = _io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"Internship Offer Letter - {intern.get('Name','')}")

    ML = _pt(15); MR = W - _pt(15); CW = MR - ML
    name     = intern.get('Name','')
    iid      = intern.get('InternID','')
    domain   = intern.get('Domain', intern.get('InternshipDomain','Software Development'))
    mode     = intern.get('Mode','')
    start    = intern.get('StartDate','')
    end      = intern.get('EndDate','')
    dur      = intern.get('InternshipDuration','')
    date_str = intern.get('GeneratedDate','')

    # ── 1. Green top banner ──────────────────────────────────────────────────
    BNR = _pt(10)
    c.setFillColor(GREEN); c.rect(0, H-BNR, W, BNR, fill=1, stroke=0)

    # ── 2. Header ────────────────────────────────────────────────────────────
    LY = H - BNR - _pt(13)

    # Company name (left)
    c.setFont('Helvetica-Bold', 17); c.setFillColor(GREEN)
    rw = c.stringWidth('REINTENSPARK ', 'Helvetica-Bold', 17)
    c.drawString(ML, LY, 'REINTENSPARK ')
    c.setFillColor(BLACK); c.drawString(ML+rw, LY, 'TECHNOLOGY')
    c.setFont('Helvetica-Bold', 9); c.drawString(ML, LY-_pt(7), 'Private Limited')

    # Address block (below "Private Limited", centred between logo and QR)
    addr_top = LY - _pt(7) - _pt(5)
    ax = ML; ay = addr_top
    c.setFont('Helvetica', 7.5); c.setFillColor(DARK_GRAY)
    for ln in ['Bengaluru', 'www.reintenspark.com',
               'support@reintenspark.com', '+91 8296969260  +91 8762719260']:
        c.drawString(ax, ay, ln); ay -= _pt(4.2)

    # QR top-right
    QSZ = _pt(24); QX = MR - QSZ; QY = H - BNR - _pt(4) - QSZ
    if include_qr:
        try:
            qr_ir = _gen_qr_ir(f"{base_url}/verify/{iid}")
            c.drawImage(qr_ir, QX, QY, width=QSZ, height=QSZ, preserveAspectRatio=True)
        except Exception as e:
            c.setStrokeColor(DARK_GRAY); c.setLineWidth(0.7)
            c.rect(QX, QY, QSZ, QSZ, fill=0, stroke=1)
            c.setFont('Helvetica', 7); c.setFillColor(LIGHT_GRAY)
            c.drawCentredString(QX+QSZ/2, QY+QSZ/2-3, 'QR CODE')

    # Date label below QR
    c.setFont('Helvetica-Bold', 8.5); c.setFillColor(BLACK)
    c.drawCentredString(QX+QSZ/2, QY - _pt(5.5), f'Date: {date_str}')

    # ── 3. Watermark ─────────────────────────────────────────────────────────
    wm = ASSETS_DIR / 'watermark.png'
    if wm.exists():
        wsz = _pt(130)
        c.saveState(); c.setFillAlpha(0.06)
        c.drawImage(_ir(wm), (W-wsz)/2, (H-wsz)/2,
                    width=wsz, height=wsz, mask='auto', preserveAspectRatio=True)
        c.restoreState()

    # ── 4. Title ─────────────────────────────────────────────────────────────
    # Title sits below the header block. Header block ends around H-BNR-37mm
    TY = H - BNR - _pt(38)
    c.setFont('Helvetica-Bold', 13); c.setFillColor(BLACK)
    c.drawCentredString(W/2, TY, 'INTERNSHIP OFFER LETTER')
    RY = TY - _pt(3)
    c.setStrokeColor(DARK_GRAY); c.setLineWidth(0.6)
    c.line(ML, RY, MR, RY)

    # ── 5. Body ───────────────────────────────────────────────────────────────
    BX = ML; BY = RY - _pt(7); LH = _pt(5.6); BW = CW - _pt(2); FS = 10.5

    c.setFont('Helvetica', FS); c.setFillColor(DARK_GRAY)
    c.drawString(BX, BY, f'Dear {name},'); BY -= _pt(7)

    BY = _para(c,
        f'We are pleased to offer you an opportunity to join Reintenspark Technology Private Limited as an '
        f'Intern with the Intern ID {iid} in the domain of {domain}. Your internship will be '
        f'conducted {mode} and will commence from {start} for a duration of {dur} month.',
        BX, BY, BW, size=FS, lh=LH)
    BY -= _pt(4)

    c.setFont('Helvetica-Bold', FS); c.setFillColor(DARK_GRAY)
    c.drawString(BX, BY, 'Internship Details'); BY -= LH * 1.3

    BUX = BX + _pt(5)
    c.setFont('Helvetica', FS); c.setFillColor(DARK_GRAY)
    for item in [f'Mode: {mode} .', f'Duration: From {start} To {end} .']:
        c.circle(BUX-_pt(3), BY+_pt(1.8), _pt(1.3), fill=1, stroke=0)
        c.drawString(BUX, BY, item); BY -= LH * 1.3
    BY -= _pt(3)

    c.setFont('Helvetica-Bold', FS); c.drawString(BX, BY, 'Benefits of this Internship'); BY -= LH
    c.setFont('Helvetica', FS)
    c.drawString(BX, BY, 'As an intern with us, you will gain access to the following benefits:'); BY -= LH * 1.3

    for btxt in [
        'Internship Completion Certificate \u2013 official recognition of your successful completion.',
        'Project-based Experience Letter \u2013 highlighting your contributions and skills.',
        'Appreciation Certificate \u2013 for top-performing interns.',
        'Placement Assistance & Direct Hiring \u2013 for outstanding interns demonstrating exceptional Performance.',
    ]:
        c.circle(BUX-_pt(3), BY+_pt(1.8), _pt(1.3), fill=1, stroke=0)
        BY = _para(c, btxt, BUX, BY, BW-_pt(5), size=FS, lh=LH); BY -= _pt(1)
    BY -= _pt(3)

    BY = _para(c,
        'We believe this internship will help you bridge the gap between academics and industry, enhance '
        'your technical and professional skills, and provide you with real-world exposure. We are excited to '
        'have you onboard and look forward to your contributions. Please confirm your acceptance of this '
        'offer by replying to this email within 5 days.',
        BX, BY, BW, size=FS, lh=LH)
    BY -= _pt(8)

    # ── 6. Signature + Badges ─────────────────────────────────────────────────
    SY = BY
    c.setFont('Helvetica', FS); c.setFillColor(DARK_GRAY)
    for ln in ['Best Regards,', 'Guru basha', 'HRM', 'ReintensparkTechnology', '+91 8762719260']:
        c.drawString(BX, SY, ln); SY -= LH * 1.2

    # Badges (right-aligned, same row as signature)
    BDH = _pt(14); BDY = BY - BDH; bx = MR
    for bp in reversed([ASSETS_DIR/'badge_iso.png', ASSETS_DIR/'badge_aicte.png', ASSETS_DIR/'badge_msme.png']):
        if bp.exists():
            try:
                bi = PILImage.open(bp)
                bw = BDH * (bi.width / bi.height)
                bx -= bw
                c.drawImage(_ir(bi), bx, BDY, width=bw, height=BDH,
                            mask='auto', preserveAspectRatio=True)
                bx -= _pt(3)
            except Exception:
                pass

    # ── 7. Footer ─────────────────────────────────────────────────────────────
    FH = _pt(13)
    c.setFillColor(DARK_GRAY); c.rect(0, 0, W, FH, fill=1, stroke=0)
    c.setFillColor(GREEN);     c.rect(0, FH, W, _pt(1.5), fill=1, stroke=0)

    fy = _pt(3.5)
    c.setFillColor(WHITE)
    c.setFont('Helvetica-Bold', 7.5); c.drawString(ML, fy+_pt(3), 'Bengaluru')
    c.setFont('Helvetica',       7.5); c.drawString(ML, fy-_pt(0.5), 'www.reintenspark.com')
    c.setStrokeColor(LIGHT_GRAY); c.setLineWidth(0.5)
    c.line(W/2, fy-_pt(1), W/2, fy+FH-_pt(3))
    c.setFont('Helvetica-Bold', 7.5); c.drawString(W/2+_pt(5), fy+_pt(3), 'support@reintenspark.com')
    c.setFont('Helvetica',       7.5); c.drawString(W/2+_pt(5), fy-_pt(0.5), '+91 8296969260    +91 8762719260')

    # ── 8. Page border ────────────────────────────────────────────────────────
    c.setStrokeColor(colors.HexColor('#c0c0c0')); c.setLineWidth(0.5)
    c.rect(_pt(5), FH+_pt(1), W-_pt(10), H-BNR-FH-_pt(6), fill=0, stroke=1)

    c.save(); buf.seek(0)
    return buf.read()
