import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from collections import defaultdict

# -----------------------------------------------------------------------------
# إعدادات الواجهة (CSS & RTL)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام فرز وترتيب العوائل المتقدم", layout="wide", page_icon="👨‍👩‍👧‍👦")
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

st.markdown("<h1 style='text-align: right;'>نظام فرز العوائل الاحترافي 👨‍👩‍👧‍👦📄</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>الترتيب المعتمد: الأبجدية العامة ⬅️ حجم العائلة تنازلياً ⬅️ أسماء الأفراد داخلياً أبجدياً.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# دوال المساعدة للوورد (Word Helpers)
# -----------------------------------------------------------------------------
def set_cell_direction(cell, direction='btLr'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), direction)
    tcPr.append(textDirection)

def set_cell_width(cell, width_cm):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567))) # 1 cm ~ 567 twips
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def apply_rtl(doc):
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.rtl = True

def preprocess_name(name):
    """معالجة الأسماء المركبة لضمان دقة التجميع الفاميلي"""
    return name.replace('عبد ', 'عبد_').replace('ابو ', 'ابو_').replace('ام ', 'ام_').replace('امة ', 'امة_').replace('الله', 'الله_')

# -----------------------------------------------------------------------------
# محرك استخراج البيانات والفرز المتعدد المستويات
# -----------------------------------------------------------------------------
def process_family_data(file_obj):
    doc_in = Document(file_obj)
    records = []
    
    # 1. قراءة البيانات من الجداول
    for table in doc_in.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
                continue
            
            name_idx = -1
            max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len:
                        max_len = len(c)
                        name_idx = i
            
            if name_idx == -1: continue
            
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices: continue
            
            card_num = cells[card_indices[1]] if len(card_indices) >= 2 else cells[card_indices[0]]
            old_card = cells[card_indices[0]] if len(card_indices) >= 2 else ""
                
            seq = "-"
            for i in range(len(cells)-1, card_indices[-1], -1):
                if cells[i].isdigit():
                    seq = cells[i]
                    break
            
            digit_cells = [cells[i] for i in range(name_idx) if cells[i].isdigit()]
            if len(digit_cells) >= 3:
                withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
            elif len(digit_cells) == 2:
                withheld, eligible, total = "0", digit_cells[0], digit_cells[1]
            else:
                withheld, eligible, total = "0", "0", "0"
                
            records.append({
                'withheld': str(withheld), 'eligible': str(eligible), 'total': str(total),
                'name': cells[name_idx], 'old_card': old_card, 'card': card_num, 'seq': seq
            })

    # 2. قراءة البيانات من النصوص (في حال عدم وجود جداول)
    if not records:
        lines = [para.text.strip() for para in doc_in.paragraphs if para.text.strip()]
        for line in lines:
            parts = line.split(',')
            if len(parts) >= 7 and parts[0].isdigit():
                records.append({
                    'withheld': parts[0].strip(), 'eligible': parts[1].strip(), 'total': parts[2].strip(),
                    'name': parts[3].strip(), 'old_card': parts[4].strip(), 'card': parts[5].strip(), 'seq': parts[-1].strip()
                })
                
    if not records:
        return [], None

    # ---------------------------------------------------------
    # خوارزمية الفرز والترتيب الهرمي المحدثة بالكامل
    # ---------------------------------------------------------
    
    # أ- تجميع الأفراد في عوائل بناءً على الجذر (آخر كلمتين)
    family_groups = defaultdict(list)
    for rec in records:
        proc_name = preprocess_name(rec['name'])
        words = proc_name.split()
        if len(words) >= 3:
            root = " ".join(words[-2:])
        elif len(words) == 2:
            root = words[-1]
        else:
            root = words[0]
        family_groups[root].append(rec)

    # ب- معالجة كل عائلة (ترتيب أفرادها داخلياً أبجدياً وتحديد الحرف الأول للجروب)
    family_list = []
    for root, members in family_groups.items():
        # ترتيب أفراد العائلة داخلياً أبجدياً
        members_sorted = sorted(members, key=lambda m: m['name'])
        leading_name = members_sorted[0]['name'].strip()
        first_char = leading_name[0] if leading_name else "أ"
        
        # توحيد كافة أشكال الألف (أ، إ، آ، ا) لكي لا تتشتت العوائل
        if first_char in ['أ', 'إ', 'آ', 'ا']:
            letter_key = 'أ'
        else:
            letter_key = first_char
            
        family_list.append({
            'root': root,
            'members': members_sorted,
            'size': len(members_sorted),
            'leading_name': leading_name,
            'letter_key': letter_key
        })

    # ج- تجميع العوائل داخل مجموعات الأحرف الكبيرة
    letter_groups = defaultdict(list)
    for fam in family_list:
        letter_groups[fam['letter_key']].append(fam)

    # د- ترتيب الأحرف أبجدياً (من الألف إلى الياء) ثم تطبيق شروط الفرز الداخلي
    sorted_letters = sorted(letter_groups.keys())
    sorted_records = []
    
    for letter in sorted_letters:
        fams_in_letter = letter_groups[letter]
        # فرز العوائل داخل الحرف الواحد: 1- حسب حجم العائلة تنازلياً. 2- حسب الاسم القائد أبجدياً
        fams_sorted = sorted(fams_in_letter, key=lambda x: (-x['size'], x['leading_name']))
        
        for fam in fams_sorted:
            sorted_records.extend(fam['members'])
        
    stats = {
        "total_records": len(records),
        "families": len(family_groups),
        "max_family_size": max([len(m) for m in family_groups.values()]) if family_groups else 0
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
            
    set_cell_direction(hdr_cells[4], 'btLr')
    set_cell_direction(hdr_cells[5], 'btLr')
    set_cell_direction(hdr_cells[6], 'btLr')
    
    for idx, rec in enumerate(sorted_records):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx + 1) # ترقيم تسلسلي تلقائي صحيح ومستمر
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

if st.button("⚙️ بدء الفرز والترتيب الهرمي المتقدم"):
    if uploaded_file:
        with st.spinner('جاري ترتيب الملف أبجدياً وفرز العوائل حسب الشروط...'):
            try:
                sorted_records, stats = process_family_data(uploaded_file)
                
                if not sorted_records:
                    st.error("❌ لم يتم العثور على جداول أو بيانات صحيحة في هذا الملف!")
                else:
                    st.markdown("<h3 style='text-align: right; margin-top: 20px;'>📊 إحصائية الفرز</h3>", unsafe_allow_html=True)
                    c1, c2, c3 = st.columns(3)
                    with c1: 
                        st.markdown(f"""<div class='report-box'>
                            <div class='stat-title'>إجمالي القيود (الأفراد)</div>
                            <div class='stat-value'>{stats['total_records']} قيد</div>
                        </div>""", unsafe_allow_html=True)
                    with c2: 
                        st.markdown(f"""<div class='report-box'>
                            <div class='stat-title'>إجمالي العوائل المستقلة</div>
                            <div class='stat-value'>{stats['families']} عائلة</div>
                        </div>""", unsafe_allow_html=True)
                    with c3: 
                        st.markdown(f"""<div class='report-box'>
                            <div class='stat-title'>أكبر عائلة مجمعة بالملف</div>
                            <div class='stat-value'>{stats['max_family_size']} أفراد</div>
                        </div>""", unsafe_allow_html=True)
                    
                    st.markdown("<h3 style='text-align: right; color: #2C3E50;'>📋 معاينة سريعة للترتيب الجديد هرمياً</h3>", unsafe_allow_html=True)
                    df_preview = pd.DataFrame(sorted_records)
                    df_preview = df_preview.rename(columns={
                        'name': 'اسم رب الاسرة', 'card': 'رقم البطاقة', 
                        'total': 'الكلية', 'eligible': 'المستحقة', 'withheld': 'المحجوبين'
                    })[['اسم رب الاسرة', 'رقم البطاقة', 'الكلية', 'المستحقة', 'المحجوبين']]
                    
                    # معاينة بـ 40 اسم لتوضيح دقة الفرز الجديد
                    st.dataframe(df_preview.head(40), use_container_width=True, hide_index=True)
                    st.caption("يعرض أول 40 قيد للمعاينة (الترتيب: الأبجدية العامة ⬅️ العوائل الأكبر أولاً ⬅️ فرز الأسماء داخلياً).")
                    
                    word_buffer = create_sorted_word_report(sorted_records)
                    original_name = uploaded_file.name.rsplit('.', 1)[0]
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.download_button(
                        label="📥 تحميل ملف الوورد النهائي الجاهز للطباعة",
                        data=word_buffer,
                        file_name=f"مرتب_هرميا_{original_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    st.success("🎉 تمت عملية الفرز والتنسيق الهرمي بنجاح تام! يمكنك تحميل ملفك الآن.")
                
            except Exception as e:
                st.error(f"❌ حدث خطأ غير متوقع: {e}")
    else:
        st.warning("⚠️ يرجى رفع الملف أولاً قبل بدء الفرز.")
