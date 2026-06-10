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
# إعدادات الواجهة الرسومية والتنسيق الفاخر
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام فرز وتوزيع العوائل الذكي", layout="wide", page_icon="⚡")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2E7D32; color: white; width: 100%; font-weight: bold; border-radius: 8px; padding: 12px; font-size: 16px;}
    div.stButton > button:hover { background-color: #1B5E20; color: #F1C40F;}
    .report-box { background-color: #F8F9FA; padding: 12px; border-radius: 8px; border-right: 5px solid #2E7D32; text-align: right; margin-bottom: 10px; box-shadow: 1px 1px 5px rgba(0,0,0,0.05);}
    .stat-title { font-size: 13px; color: #7F8C8D; font-weight: bold; }
    .stat-value { font-size: 16px; color: #2C3E50; font-weight: bold; margin-top: 3px;}
    /* تحسين اتجاه جداول Streamlit للعربية */
    div[data-testid="stDataEditor"] { direction: rtl !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>اللوحة الذكية لفرز وتوزيع العوائل (7 مناطق) 🗺️⚡</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>قم بتوزيع العوائل ككتل كاملة بسرعة فائقة من خلال الجدول التفاعلي أدناه، مع ميزة التقديم التلقائي للآباء وعزل النساء.</p>", unsafe_allow_html=True)

LIST_OF_ZONES = ["المنطقة الأولى", "المنطقة الثانية", "المنطقة الثالثة", "المنطقة الرابعة", "المنطقة الخامسة", "المنطقة السادسة", "المنطقة السابعة"]

# -----------------------------------------------------------------------------
# دوال معالجة مستندات Word والأسماء
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
        if hasattr(style, 'font'): style.font.rtl = True

def preprocess_name(name):
    name = re.sub(r'\s+', ' ', name.strip())
    for prefix in ['عبد ', 'ابو ', 'أبو ', 'أم ', 'ام ', 'امة ', 'أمة ', 'آل ']:
        name = name.replace(prefix, prefix.strip() + '_')
    return name

def is_female(name):
    first_word = name.strip().split()[0]
    if first_word in ['ام', 'أم', 'امة', 'أمة', 'آمال']: return True
    male_exceptions = ['حمزة', 'اسامة', 'أسامة', 'حذيفة', 'قتادة', 'طلحة', 'خليفة', 'معاوية', 'عطية', 'حارثة', 'عروة', 'عبيدة', 'ميسرة', 'سلامة']
    if first_word.endswith('ة') and first_word not in male_exceptions: return True
    female_names = [
        'زينب', 'مريم', 'شهد', 'نور', 'هدى', 'ندى', 'ليلى', 'سها', 'مها', 'ريم', 'سعاد', 'هند',
        'ايمان', 'إيمان', 'رحاب', 'سحر', 'سمر', 'كوثر', 'غدير', 'حنان', 'منى', 'اسماء', 'أسماء',
        'اسراء', 'إسراء', 'شيماء', 'دعاء', 'وفاء', 'رجاء', 'سناء', 'لقاء', 'نجلاء', 'حوراء',
        'عذراء', 'زهراء', 'بيداء', 'رفل', 'رسل', 'ضحى', 'شمس', 'قمر', 'فرح', 'مرح', 'ابتسام',
        'إبتسام', 'امال', 'ابتهال', 'انتصار', 'إنتصار', 'انعام', 'أنعام', 'بشرى', 'ذكرى', 'رؤى', 
        'سرى', 'سروة', 'سهام', 'شروق', 'صبا', 'عواطف', 'فاتن', 'لمياء', 'لينا', 'محاسن', 'مروة', 
        'نوال', 'هاجر', 'وجدان', 'ياسمين', 'يسرى', 'بنين', 'غفران', 'فردوس', 'افراح', 'أفراح', 
        'اخلاص', 'إخلاص', 'ازهار', 'أزهار', 'الحان', 'ألحان', 'انوار', 'أنوار', 'اشواق', 'أشواق', 
        'براء', 'تبارك', 'تغريد', 'جمانة', 'جنان', 'جواهر', 'حلا', 'خلود', 'داليا', 'دلال', 'رغد', 
        'رنا', 'رنين', 'روان', 'سالي', 'سجى', 'سوزان', 'صابرين', 'عالية', 'عبير', 'عفاف', 'غادة', 
        'ليال', 'مي', 'ميسون', 'نادية', 'نجاة', 'نغم', 'نهى', 'هديل', 'هناء', 'هيفاء', 'ورود', 
        'وسن', 'وعد', 'يقين', 'إلهام', 'الهام', 'تهاني', 'سلوى', 'رشا', 'سهير', 'منال', 'آية', 
        'اية', 'غسق', 'شيرين', 'نسرين', 'جيهان', 'إيناس', 'ايناس', 'رواء'
    ]
    if first_word in female_names: return True
    return False

# -----------------------------------------------------------------------------
# دالة سحب وتجميع البيانات العائلية الهرمية
# -----------------------------------------------------------------------------
def extract_and_group_data(file_obj):
    doc_in = Document(file_obj)
    records = []
    
    for table in doc_in.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells): continue
            
            name_idx = -1; max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len: max_len = len(c); name_idx = i
            if name_idx == -1: continue
            
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices: continue
            card_num = cells[card_indices[1]] if len(card_indices) >= 2 else cells[card_indices[0]]
            old_card = cells[card_indices[0]] if len(card_indices) >= 2 else ""
            seq = next((cells[i] for i in range(len(cells)-1, card_indices[-1], -1) if cells[i].isdigit()), "-")
            digit_cells = [cells[i] for i in range(name_idx) if cells[i].isdigit()]
            
            if len(digit_cells) >= 3: withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
            elif len(digit_cells) == 2: withheld, eligible, total = "0", digit_cells[0], digit_cells[1]
            else: withheld, eligible, total = "0", "0", "0"
                
            records.append({
                'withheld': str(withheld), 'eligible': str(eligible), 'total': str(total),
                'name': cells[name_idx].strip(), 'old_card': old_card, 'card': card_num, 'seq': seq
            })

    if not records: return None

    male_records = []
    female_records = []
    
    for rec in records:
        rec['processed_name'] = preprocess_name(rec['name'])
        rec['is_female'] = is_female(rec['name'])
        if rec['is_female']:
            female_records.append(rec)
        else:
            words = rec['processed_name'].split()
            rec['father_string'] = " ".join(words[1:]) if len(words) > 1 else rec['processed_name']
            male_records.append(rec)

    male_full_names = {r['processed_name']: r for r in male_records}
    unique_fathers = list(set(r['father_string'] for r in male_records))
    
    family_map = {}
    for f in unique_fathers:
        if f in male_full_names:
            family_map[f] = f
        else:
            matches = [other for other in unique_fathers if other.startswith(f + " ") and other != f]
            family_map[f] = matches[0] if len(matches) == 1 else f

    final_family_map = {}
    for f in unique_fathers:
        current = f
        for _ in range(5):
            target = family_map.get(current, current)
            if target == current: break
            current = target
        final_family_map[f] = current

    final_family_groups = defaultdict(list)
    for rec in male_records:
        if rec['processed_name'] in unique_fathers:
            assigned_root = final_family_map[rec['processed_name']]
        else:
            assigned_root = final_family_map[rec['father_string']]
        rec['root_id'] = assigned_root
        final_family_groups[assigned_root].append(rec)

    families_list = []
    for root, members in final_family_groups.items():
        def sort_key(m):
            is_father = 0 if m['processed_name'] == root else 1
            return (is_father, m['name'])
            
        members_sorted = sorted(members, key=sort_key)
        families_list.append({
            'root_id': root,
            'father_display': members_sorted[0]['processed_name'].replace('_', ' '),
            'leading_name': members_sorted[0]['name'],
            'size': len(members_sorted),
            'members': members_sorted,
            'is_female_group': False
        })

    for rec in female_records:
        fam_id = 'female_' + str(id(rec))
        rec['root_id'] = fam_id
        families_list.append({
            'root_id': fam_id,
            'father_display': rec['name'] + ' (بيانات نساء)',
            'leading_name': rec['name'],
            'size': 1,
            'members': [rec],
            'is_female_group': True
        })
        
    return families_list

def create_word_file(sorted_records):
    doc_out = Document()
    apply_rtl(doc_out)
    headers = ['ت', 'رقم البطاقة', 'رقم البطاقة القديم', 'اسم رب الاسرة', 'الافراد الكلية', 'الافراد المستحقة', 'الافراد المحجوبين']
    table = doc_out.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in hdr_cells[i].paragraphs[0].runs: run.font.size, run.font.bold, run.font.name = Pt(13), True, 'Arial'
            
    set_cell_direction(hdr_cells[4], 'btLr'); set_cell_direction(hdr_cells[5], 'btLr'); set_cell_direction(hdr_cells[6], 'btLr')
    
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
            for run in row_cells[i].paragraphs[0].runs: run.font.size, run.font.name = Pt(15), 'Arial'
                
    widths = [1.2, 3.5, 3.5, 5.0, 0.8, 0.8, 0.8]
    for row in table.rows:
        for idx, width in enumerate(widths): set_cell_width(row.cells[idx], width)
            
    buffer = BytesIO()
    doc_out.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# سير العمل والواجهة التفاعلية الفائقة
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 خطوة 1: ارفع ملف الوورد الأساسي</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=['docx'], label_visibility="collapsed")

if uploaded_file:
    # تهيئة البيانات وبنائها عند الرفع لأول مرة
    if 'file_key' not in st.session_state or st.session_state.file_key != uploaded_file.name:
        families = extract_and_group_data(uploaded_file)
        if families:
            st.session_state.families = families
            st.session_state.file_key = uploaded_file.name
            
            # بناء جدول تعديل البيانات الذكي
            rows_list = []
            for fam in families:
                rows_list.append({
                    'id': fam['root_id'],
                    'اسم العائلة / الأب المرجعي': fam['father_display'],
                    'عدد الأفراد': fam['size'],
                    'معاينة أفراد العائلة والأبناء داخل الكتلة': ", ".join([m['name'] for m in fam['members']]),
                    'تحديد المنطقة': LIST_OF_ZONES[0] # الافتراضية
                })
            st.session_state.df_editable = pd.DataFrame(rows_list)

    if 'df_editable' in st.session_state:
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: right;'>🛠️ خطوة 2: جدول التوزيع الفوري (تعديل مباشر وسريع جداً)</h3>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: right; color:#7F8C8D;'>💡 نصيحة: استخدم خانة البحث في أعلى يمين الجدول لتصفية عائلة معينة باللقب أو الاسم وتغيير منطقتها فوراً.</p>", unsafe_allow_html=True)

        # عرض الجدول التفاعلي الاحترافي
        edited_df = st.data_editor(
            st.session_state.df_editable,
            column_config={
                "id": None, # إخفاء المعرف الرقمي للخلفية
                "اسم العائلة / الأب المرجعي": st.column_config.TextColumn("👨‍👦 اسم العائلة المرجعي", disabled=True, width="medium"),
                "عدد الأفراد": st.column_config.NumberColumn("🔢 العدد", disabled=True, width="small"),
                "معاينة أفراد العائلة والأبناء داخل الكتلة": st.column_config.TextColumn("👥 الأفراد المدموجين تلقائياً", disabled=True, width="large"),
                "تحديد المنطقة": st.column_config.SelectboxColumn(
                    "🏢 اختر المنطقة السكنية",
                    options=LIST_OF_ZONES,
                    required=True,
                    width="medium"
                )
            },
            hide_index=True,
            use_container_width=True
        )
        
        # حفظ التعديلات بداخل الـ session لحين معالجتها
        st.session_state.df_editable = edited_df

        # ---------------------------------------------------------------------
        # توليد المستندات النهائية للمناطق السبعة بناءً على جدول التعديل المباشر
        # ---------------------------------------------------------------------
        st.markdown("<br><hr>", unsafe_allow_html=True)
        if st.button("⚙️ معالجة نهائية وإنتاج مستندات الـ 7 مناطق"):
            
            # خلق خريطة سريعة لربط المعرف بالمنطقة المختارة من قبل المستخدم
            zone_mapping = dict(zip(edited_df['id'], edited_df['تحديد المنطقة']))
            
            zone_fams = {zone: [] for zone in LIST_OF_ZONES}
            for fam in st.session_state.families:
                assigned_zone = zone_mapping.get(fam['root_id'], LIST_OF_ZONES[0])
                zone_fams[assigned_zone].append(fam)
            
            def finalize_zone(families_in_zone):
                males = [f for f in families_in_zone if not f['is_female_group']]
                females = [f for f in families_in_zone if f['is_female_group']]
                
                # ترتيب هرمي للرجال: أبجدياً، ثم الأب في البداية تلقائياً
                sorted_males = sorted(males, key=lambda x: (x['leading_name'][0] if x['leading_name'] else 'أ', -x['size'], x['leading_name']))
                # ترتيب النساء أبجدياً وعزلهن كاملاً في أسفل القائمة
                sorted_females = sorted(females, key=lambda x: x['leading_name'])

                final_list = []
                for f in sorted_males: final_list.extend(f['members'])
                for f in sorted_females: final_list.extend(f['members'])
                return final_list

            final_buffers = {}
            final_counts = {}
            for zone in LIST_OF_ZONES:
                zone_records = finalize_zone(zone_fams[zone])
                final_counts[zone] = len(zone_records)
                final_buffers[zone] = create_word_file(zone_records) if zone_records else None

            st.markdown("<h3 style='text-align: right;'>📥 خطوة 3: روابط تحميل ملفات المناطق الجاهزة للطباعة</h3>", unsafe_allow_html=True)
            
            st.markdown("<h4 style='text-align: right; color:#1B5E20;'>📍 الجزء الأول:</h4>", unsafe_allow_html=True)
            col_r1 = st.columns(4)
            for i in range(4):
                z_name = LIST_OF_ZONES[i]
                with col_r1[i]:
                    st.markdown(f"<div class='report-box'><div class='stat-title'>{z_name}</div><div class='stat-value'>{final_counts[z_name]} قيد عائلي</div></div>", unsafe_allow_html=True)
                    if final_buffers[z_name]:
                        st.download_button(f"📥 مستند {z_name}", data=final_buffers[z_name], file_name=f"{z_name.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                    else: st.caption("فارغة")

            st.markdown("<br><h4 style='text-align: right; color:#1B5E20;'>📍 الجزء الثاني:</h4>", unsafe_allow_html=True)
            col_r2 = st.columns(3)
            for i in range(3):
                z_name = LIST_OF_ZONES[4 + i]
                with col_r2[i]:
                    st.markdown(f"<div class='report-box'><div class='stat-title'>{z_name}</div><div class='stat-value'>{final_counts[z_name]} قيد عائلي</div></div>", unsafe_allow_html=True)
                    if final_buffers[z_name]:
                        st.download_button(f"📥 مستند {z_name}", data=final_buffers[z_name], file_name=f"{z_name.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                    else: st.caption("فارغة")
                    
            st.success("🎉 تمت عملية تصفية الفرز والدمج بنجاح مذهل! العوائل والنساء تم تنظيمهن بدقة حاسوبية كاملة.")
else:
    st.info("👋 مرحباً بك! يرجى رفع ملف الـ Word الأصلي للبدء بالفرز السريع وفصل المناطق السبعة فوراً.")
