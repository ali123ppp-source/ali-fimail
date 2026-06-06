import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from collections import defaultdict
import re

# -----------------------------------------------------------------------------
# إعدادات الواجهة الرسومية والتنسيق (CSS & RTL)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام فرز وترتيب العوائل الذكي", layout="wide", page_icon="👨‍👩‍👧‍👦")
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

st.markdown("<h1 style='text-align: right;'>نظام الفرز الهرمي وحماية الألقاب 👨‍👩‍👧‍👦📄</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>تم التحديث: عزل كامل للعوائل المتشابهة بالأسماء الثلاثية بناءً على اللقب والاسم الرابع، مع دمج الحالات الخاصة (مثل خالد) فقط في حال تطابقها غير اللامشروط مع إخوتها.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# دوال المساعدة لملفات الوورد (Word Helpers)
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
    tcW.set(qn('w:w'), str(int(width_cm * 567))) 
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def apply_rtl(doc):
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.rtl = True

def preprocess_name(name):
    """ربط الأسماء المركبة لضمان دقة فصل واستخراج الألقاب وسلسلة الآباء"""
    name = name.strip()
    name = re.sub(r'\s+', ' ', name)
    name = name.replace('عبد ', 'عبد_')
    name = name.replace('ابو ', 'ابو_')
    name = name.replace('أبو ', 'أبو_')
    name = name.replace('أم ', 'أم_')
    name = name.replace('ام ', 'ام_')
    name = name.replace('امة ', 'امة_')
    name = name.replace('آل ', 'آل_')
    return name

# -----------------------------------------------------------------------------
# محرك استخراج البيانات والفرز المتقدم لحماية الألقاب
# -----------------------------------------------------------------------------
def process_family_data(file_obj):
    doc_in = Document(file_obj)
    records = []
    
    # 1. استخراج البيانات من الجداول
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

    # 2. استخراج البيانات من النصوص الحرة (في حال عدم وجود جداول)
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
    # خوارزمية الفرز المحدثة: حماية الألقاب ومنع التداخل العشوائي
    # ---------------------------------------------------------
    
    # أ- بناء سلسلة الآباء لكل قيد (الاسم كاملاً بدون الكلمة الأولى)
    for rec in records:
        proc_name = preprocess_name(rec['name'])
        words = proc_name.split()
        if len(words) > 1:
            rec['father_string'] = " ".join(words[1:])
        else:
            rec['father_string'] = proc_name

    unique_fathers = list(set(r['father_string'] for r in records))
    
    # ب- خوارزمية التقييم والربط الآمن لحماية الألقاب (Ambiguity Guard)
    family_map = {}
    for f in unique_fathers:
        # البحث عن عوائل كاملة تبدأ بنفس السلسلة وتختلف عنها (امتداد لها باللقب)
        matches = [other for other in unique_fathers if other.startswith(f + " ") and other != f]
        
        # لا ندمج إلا إذا كان الاسم المختصر ينتمي لعائلة واحدة فريدة ومحددة في الملف (مثل حالة خالد)
        if len(matches) == 1:
            family_map[f] = matches[0]
        else:
            # إذا لم يوجد امتداد، أو وُجد أكثر من امتداد بلقبين مختلفين (مثال: الخفاجي والجبوري)، نمنع الدمج تماماً لحماية اللقب
            family_map[f] = f
            
    # ج- حل السلاسل التتابعية للوصول للجذر النهائي الموحد
    final_family_map = {}
    for f in unique_fathers:
        current = f
        for _ in range(5):
            target = family_map.get(current, current)
            if target == current:
                break
            current = target
        final_family_map[f] = current

    # د- توزيع القيود على مجموعات العوائل النهائية الآمنة والمفصولة بالألقاب
    final_family_groups = defaultdict(list)
    for rec in records:
        assigned_root = final_family_map[rec['father_string']]
        final_family_groups[assigned_root].append(rec)

    # هـ- ترتيب أفراد كل عائلة داخلياً أبجدياً وتحديد الحرف العام للقائد
    family_list = []
    for root, members in final_family_groups.items():
        members_sorted = sorted(members, key=lambda m: m['name'])
        leading_name = members_sorted[0]['name'].strip()
        first_char = leading_name[0] if leading_name else "أ"
        
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

    # و- التجميع النهائي حسب القواعد (الحرف الأبجدي أولاً -> حجم العائلة تنازلياً -> العوائل المتساوية أبجدياً)
    letter_groups = defaultdict(list)
    for fam in family_list:
        letter_groups[fam['letter_key']].append(fam)

    sorted_letters = sorted(letter_groups.keys())
    sorted_records = []
    
    for letter in sorted_letters:
        fams_in_letter = letter_groups[letter]
        # فرز العوائل داخل الحرف: الأكبر حجماً أولاً، وعند التساوي يتم الاعتماد على أبجدية الاسم القائد
        fams_sorted = sorted(fams_in_letter, key=lambda x: (-x['size'], x['leading_name']))
        
        for fam in fams_sorted:
            # صب أفراد العائلة الواحدة ككتلة متكاملة ومستمرة تماماً دون أي انقطاع
            sorted_records.extend(fam['members'])
        
    stats = {
        "total_records": len(records),
        "families": len(final_family_groups),
        "max_family_size": max([len(m) for m in final_family_groups.values()]) if final_family_groups else 0
    }
        
    return sorted_records, stats

# -----------------------------------------------------------------------------
# محرك بناء تقرير ملف Word النهائي
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
                
    widths = [1.2, 3.5, 3.5, 5.0, 0.8, 0.8, 0.8]
    for row in table.rows:
        for idx, width in enumerate(widths):
            set_cell_width(row.cells[idx], width)
            
    buffer = BytesIO()
    doc_out.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة الاستخدام Streamlit
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع ملف البيانات الأصلي</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("قم برفع ملف الوورد المراد فرزه وترتيبه", type=['docx'], label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("⚙️ بدء المعالجة والفرز الآمن بحماية الألقاب"):
    if uploaded_file:
        with st.spinner('جاري حماية الألقاب وفصل العوائل المتشابهة ثلاثياً وتجميع الكتل...'):
            try:
                sorted_records, stats = process_family_data(uploaded_file)
                
                if not sorted_records:
                    st.error("❌ لم يتم العثور على بيانات أو جداول صالحة للمعالجة في هذا الملف!")
                else:
                    st.markdown("<h3 style='text-align: right; margin-top: 20px;'>📊 إحصائيات دقة الفرز</h3>", unsafe_allow_html=True)
                    c1, c2, c3 = st.columns(3)
                    with c1: 
                        st.markdown(f"""<div class='report-box'>
                            <div class='stat-title'>إجمالي الأفراد المعالجين</div>
                            <div class='stat-value'>{stats['total_records']} فرد</div>
                        </div>""", unsafe_allow_html=True)
                    with c2: 
                        st.markdown(f"""<div class='report-box'>
                            <div class='stat-title'>إجمالي العوائل المستقلة (المفصولة باللقب)</div>
                            <div class='stat-value'>{stats['families']} عائلة</div>
                        </div>""", unsafe_allow_html=True)
                    with c3: 
                        st.markdown(f"""<div class='report-box'>
                            <div class='stat-title'>أكبر كتلة عائلية متتابعة</div>
                            <div class='stat-value'>{stats['max_family_size']} أفراد</div>
                        </div>""", unsafe_allow_html=True)
                    
                    st.markdown("<h3 style='text-align: right; color: #2C3E50;'>📋 معاينة مباشرة للجداول الناتجة</h3>", unsafe_allow_html=True)
                    df_preview = pd.DataFrame(sorted_records)
                    df_preview = df_preview.rename(columns={
                        'name': 'اسم رب الاسرة', 'card': 'رقم البطاقة', 
                        'total': 'الكلية', 'eligible': 'المستحقة', 'withheld': 'المحجوبين'
                    })[['اسم رب الاسرة', 'رقم البطاقة', 'الكلية', 'المستحقة', 'المحجوبين']]
                    
                    st.dataframe(df_preview.head(50), use_container_width=True, hide_index=True)
                    st.caption("توضح المعاينة أعلاه ثبات كل عائلة ككتلة صلبة، مع بقاء عوائل الألقاب المختلفة منفصلة تماماً حتى وإن تطابقت أسماؤها الثلاثية.")
                    
                    word_buffer = create_sorted_word_report(sorted_records)
                    original_name = uploaded_file.name.rsplit('.', 1)[0]
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.download_button(
                        label="📥 تحميل ملف الوورد النهائي المنسق والمحمي",
                        data=word_buffer,
                        file_name=f"فرز_آمن_{original_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    st.success("🎉 اكتملت عملية الفرز وتأمين الألقاب والكتل المتتابعة بنجاح تام! الملف جاهز للتحميل والطباعة فوراً.")
                
            except Exception as e:
                st.error(f"❌ حدث خطأ غير متوقع أثناء الفرز: {e}")
    else:
        st.warning("⚠️ يرجى رفع ملف الوورد أولاً لبدء العمل.")
