"""Runtime font compatibility for ReportLab PDF generation on the production server."""
import os

try:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    normal = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    if os.path.exists(normal):
        try:
            pdfmetrics.registerFont(TTFont("DejaVu", normal))
        except Exception:
            pass
    if os.path.exists(bold):
        try:
            pdfmetrics.registerFont(TTFont("DejaVuBold", bold))
        except Exception:
            pass
    # ReportLab's Paragraph parser expects a registered font family when a
    # style uses the bold variant name. Without this mapping, PDF generation
    # can fail with: Can't map determine family/bold/italic for dejavubold.
    pdfmetrics.registerFontFamily(
        "DejaVu", normal="DejaVu", bold="DejaVuBold", italic="DejaVu", boldItalic="DejaVuBold"
    )
except Exception:
    # Never prevent the Flask application from starting if the optional
    # compatibility registration cannot be performed.
    pass
