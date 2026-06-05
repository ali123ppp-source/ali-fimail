# word_formatter.py
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_direction(cell, direction='btLr'):
    """تدوير النص داخل الخلية (90 درجة عمودياً)"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), direction)
    tcPr.append(textDirection)

def set_cell_width(cell, width_cm):
    """تحديد عرض الخلية بالسنتيمتر بدقة"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567))) # 1 سم يعادل 567 twips
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def apply_rtl(doc):
    """تطبيق اتجاه النص من اليمين لليسار (RTL) لكامل الوثيقة"""
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.rtl = True