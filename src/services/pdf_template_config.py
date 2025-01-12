from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

LOGO_PATH = 'path/to/your/logo.png'

FONTS = {
    'regular': 'Arial',
    'bold': 'Arial-Bold',
}

STYLES = {
    'CustomNormal': {
        'fontName': 'Arial',
        'fontSize': 10,
        'leading': 12,
    },
    'CustomHeading': {
        'fontName': 'Arial-Bold',
        'fontSize': 12,
        'leading': 14,
        'alignment': 1,
    },
}

MARGINS = {
    'left': 2*cm,
    'right': 2*cm,
    'top': 5*cm,
    'bottom': 2*cm,
}

PAGE_SIZE = A4

COLORS = {
    'header_bg': colors.HexColor('#283C5C'),
    'header_text': colors.white,
    'row_alternate_bg': colors.HexColor('#F8F8F8'),
}
