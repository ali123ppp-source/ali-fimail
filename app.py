import streamlit as st
import pandas as pd
from io import BytesIO
import docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# -----------------------------------------------------------------------------
# إعدادات الواجهة (CSS & RTL)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام فرز وترتيب العوائل", layout="wide", page_icon="👨‍👩‍👧‍👦")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px; padding: 10px;}
    div.stButton > button:hover { background-color: #1A252F; color: #F1C40F;}
    .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
    .stat-title { font-size: 16px; color: #7F8C8D; font-weight: bold; }
    .stat-value { font-size: 24px; color: #2C3E50; font-weight: bold; margin-top: 5px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام فرز العوائل الذكي 👨‍👩‍👧‍👦📄</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>يقوم بفرز البيانات أبجدياً وربط الأبناء بآبائهم آلياً، مع تصدير ملف Word بتنسيق رسمي ومقاسات دقيقة.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# دوال المساعدة للوورد (Word Helpers)
# -----------------------------------------------------------------------------
def set_cell_direction(cell, direction='btLr'):
    """تدوير النص داخل الخلية 90 درجة"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), direction)
    tcPr.append(textDirection)

def set_cell_width(cell, width_cm):
    """تحديد عرض الخلية بالسنتيمتر"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567))) # 1 cm ~ 567 twips
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def apply_rtl(doc):
    """تطبيق اتجاه النص من اليمين لليسار"""
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.rtl = True

def preprocess_name(name):
    """معالجة الأسماء المركبة"""
    return name.replace('عبد ', 'عبد_').replace('ابو ', 'ابو_').replace('ام ', 'ام_').replace('امة ', 'امة_')

def postprocess_name(name):
    """إرجاع الأسماء المركبة لطبيعتها"""
    return name.replace('_', ' ')

# -----------------------------------------------------------------------------
# محرك استخراج البيانات والفرز
# -----------------------------------------------------------------------------
def process_family_data(file_obj):
    doc_in = Document(file_obj)
    lines = [para.text.strip() for para in doc_in.paragraphs if para.text.strip()]
    
    records = []
    for line in lines:
        parts = line.split(',')
        if len(parts) == 7 and parts[0].isdigit() and parts[-1].isdigit():
            records.append({
                'withheld': parts[0].strip(),
                'eligible': parts[1].strip(),
                'total': parts[2].strip(),
                'name': parts[3].strip(),
                'old_card': parts[4].strip(),
                'card': parts[5].strip(),
                'seq': parts[6].strip()
            })
            
    names_set = {r['name'] for r in records}
    children_map = {name: [] for name in names_set}
    parents_map = {}
    
    for name in names_set:
        processed_name = preprocess_name(name)
        words = processed_name.split()
        for i in range(1, len(words) - 1):
            potential_parent_processed = " ".join(words[i:])
            potential_parent = postprocess_name(potential_parent_processed)
            if potential_parent in names_set:
                children_map[potential_parent].append(name)
                parents_map[name] = potential_parent
                break
                
    roots = [name for name in names_set if name not in parents_map]
    roots.sort()
    
    record_by_name = {r['name']: r for r in records}
    sorted_records = []
    
    def add_family(person):
        sorted_records.append(record_by_name[person])
        children_map[person].sort()
        for child in children_map[person]:
            add_family(child)
            
    for root in roots:
        add_family(root)
        
    # إحصائيات سريعة للواجهة
    stats = {
        "total_records": len(records),
        "families": len(roots),
        "children": len(records) - len(roots)
    }
        
    return sorted_records, stats

# -----------------------------------------------------------------------------
# محرك إنشاء ملف Word النهائي
# -----------------------------------------------------------------------------
def create_sorted_word_report(sorted_records):
    doc_out = Document()
    apply_rtl(doc_out)
    
    headers = ['ت', 'رقم البطاقة', 'رقم البطاقة القديم', 'اسم رب الاسرة', 'الافراد الكلية', 'الافراد المستحقة', 'الافراد المحجوبين']
    table = doc_out.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    table.autofit = False
    
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in hdr_cells[i].paragraphs[0].runs:
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.name = 'Arial'
            
    # الدوران 90 درجة للحقول الثلاثة
    set_cell_direction(hdr_cells[4], 'btLr')
    set_cell_direction(hdr_cells[5], 'btLr')
    set_cell_direction(hdr_cells[6], 'btLr')
    
    for idx, rec in enumerate(sorted_records):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx + 1)
        row_cells[1].text = rec['card']
        row_cells[2].text = rec['old_card']
        row_cells[3].text = rec['name']
        row_cells[4].text = rec['total']
        row_cells[5].text = rec['eligible']
        row_cells[6].text = rec['withheld']
        
        for i in range(7):
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in row_cells[i].paragraphs[0].runs:
                run.font.size = Pt(16)
                run.font.name = 'Arial'
                
    # المقاسات المطلوبة (الاسم 5 سم)
    widths = [1.2, 3.5, 3.5, 5.0, 0.8, 0.8, 0.8]
    for row in table.rows:
        for idx, width in enumerate(widths):
            set_cell_width(row.cells[idx], width)
            
    buffer = BytesIO()
    doc_out.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة الاستخدام والتفاعل
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع ملف البيانات</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("قم برفع ملف الوورد (المراد ترتيبه)", type=['docx'], label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("⚙️ بدء الفرز والترتيب الأبجدي"):
    if uploaded_file:
        with st.spinner('جاري تحليل الأسماء وبناء شجرة العائلة...'):
            try:
                # 1. المعالجة
                sorted_records, stats = process_family_data(uploaded_file)
                
                # 2. عرض الإحصائيات
                st.markdown("<h3 style='text-align: right; margin-top: 20px;'>📊 إحصائية الفرز</h3>", unsafe_allow_html=True)
                c1, c2, c3 = st.columns(3)
                with c1: 
                    st.markdown(f"""<div class='report-box'>
                        <div class='stat-title'>إجمالي القيود (الأفراد)</div>
                        <div class='stat-value'>{stats['total_records']} قيد</div>
                    </div>""", unsafe_allow_html=True)
                with c2: 
                    st.markdown(f"""<div class='report-box'>
                        <div class='stat-title'>الآباء (العوائل الأساسية)</div>
                        <div class='stat-value'>{stats['families']} أب</div>
                    </div>""", unsafe_allow_html=True)
                with c3: 
                    st.markdown(f"""<div class='report-box'>
                        <div class='stat-title'>الأبناء (المتفرعين)</div>
                        <div class='stat-value'>{stats['children']} ابن</div>
                    </div>""", unsafe_allow_html=True)
                
                # 3. عرض جدول معاينة سريع
                st.markdown("<h3 style='text-align: right; color: #2C3E50;'>📋 معاينة سريعة للترتيب الجديد</h3>", unsafe_allow_html=True)
                df_preview = pd.DataFrame(sorted_records)
                # إعادة تسمية الأعمدة للمعاينة فقط
                df_preview = df_preview.rename(columns={
                    'name': 'اسم رب الاسرة', 'card': 'رقم البطاقة', 
                    'total': 'الكلية', 'eligible': 'المستحقة', 'withheld': 'المحجوبين'
                })[['اسم رب الاسرة', 'رقم البطاقة', 'الكلية', 'المستحقة', 'المحجوبين']]
                
                st.dataframe(df_preview.head(15), use_container_width=True, hide_index=True)
                st.caption("يعرض أول 15 قيد للتأكد من الترتيب (الأب ثم أبناؤه).")
                
                # 4. توليد وتحميل الوورد
                word_buffer = create_sorted_word_report(sorted_records)
                original_name = uploaded_file.name.rsplit('.', 1)[0]
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 تحميل ملف الوورد النهائي (منسق وجاهز للطباعة)",
                    data=word_buffer,
                    file_name=f"مرتب_{original_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                st.success("🎉 تمت عملية الفرز والتنسيق بنجاح! يمكنك تحميل الملف الآن.")
                
            except Exception as e:
                st.error(f"❌ حدث خطأ أثناء المعالجة: تأكد من أن الملف المدخل مطابق للصيغة المطلوبة.\n{e}")
    else:
        st.warning("⚠️ يرجى رفع الملف أولاً قبل بدء الفرز.")