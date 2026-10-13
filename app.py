import os
import io
import sys
import subprocess
import tempfile
import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn

# إعداد الصفحة
st.set_page_config(page_title="نظام تعديل نماذج KYC الذكي", page_icon="📝", layout="centered")

# ==========================================
# 🔑 إعدادات الحسابات وربطها بالمجلدات
# ==========================================
USER_CREDENTIALS = {
    "yousif": {
        "password": "Asdf1asdf",        # كلمة المرور ليوسف
        "folder": "يوسف"          # اسم المجلد الخاص به
    },
    "saif": {
        "password": "saif",        # كلمة المرور لسيف
        "folder": "سيف"            # اسم المجلد الخاص به
    }
}

# تحديد مسار المجلد الأساسي
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# ==========================================
# 📊 فصل الأنشطة الإنجليزية المدمجة وتنسيق العرض (مع دعم المطابقة غير الحساسة للأحرف)
# ==========================================
@st.cache_data
def load_mcc_database():
    mcc_file_path = os.path.join(BASE_DIR, "MCC NEW.xlsx")
    mcc_dict = {}
    display_map = {} 
    dropdown_labels = set()
    
    if os.path.exists(mcc_file_path):
        try:
            df = pd.read_excel(mcc_file_path, sheet_name="Sheet1")
            valid_df = df.dropna(subset=['MCC'])
            for _, row in valid_df.iterrows():
                seg_ar = str(row['Segment']).strip()
                seg_en = str(row['Segment.1']).strip()
                mcc_val = str(row['MCC']).strip()
                if mcc_val.endswith('.0'):
                    mcc_val = mcc_val[:-2]
                
                if seg_en and seg_en.lower() != 'nan':
                    # فصل الأنشطة الإنجليزية عند وجود علامة الزائد أو الفاصلة أو المائلة
                    clean_seg_en = seg_en.replace('(', '').replace(')', '').strip()
                    parts_en = [p.strip() for p in clean_seg_en.replace('+', ',').replace('/', ',').split(',')]
                    
                    clean_seg_ar = seg_ar.replace('(', '').replace(')', '').strip() if seg_ar and seg_ar.lower() != 'nan' else ""
                    
                    for part_en in parts_en:
                        if part_en:
                            # تخزين القيم بصيغ متعددة لتجنب مشاكل تطابق الحروف الكبيرة والصغيرة
                            mcc_dict[part_en] = mcc_val
                            mcc_dict[part_en.upper()] = mcc_val
                            mcc_dict[part_en.lower()] = mcc_val
                            
                            # تنسيق العرض: الإنجليزي بأحرف كبيره واضح وبجانبه العربي الكامل
                            label = f"{part_en.upper()} — ({clean_seg_ar})"
                            display_map[label] = part_en
                            dropdown_labels.add(label)
        except Exception as e:
            st.error(f"خطأ في قراءة ملف الـ MCC: {e}")
            
    return mcc_dict, display_map, sorted(list(dropdown_labels))

MCC_MAPPING, DISPLAY_MAP, AVAILABLE_SEGMENTS = load_mcc_database()

# إدارة حالة تسجيل الدخول عبر Session State
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "logged_user" not in st.session_state:
    st.session_state["logged_user"] = None

# مفتاح تصفير النموذج
if "form_reset_key" not in st.session_state:
    st.session_state["form_reset_key"] = 0


# ==========================================
# 1. شاشة تسجيل الدخول
# ==========================================
if not st.session_state["authenticated"]:
    st.title("🔐 تسجيل الدخول للنظام")
    st.write("يرجى إدخال اسم المستخدم ورمز الدخول للوصول إلى النماذج الخاصّة بك.")

    with st.form("login_form"):
        username_input = st.text_input("اسم المستخدم (yousif / saif)").strip().lower()
        password_input = st.text_input("رمز الدخول (كلمة المرور)", type="password").strip()
        submit_button = st.form_submit_button("تسجيل الدخول", type="primary")

        if submit_button:
            if username_input in USER_CREDENTIALS and USER_CREDENTIALS[username_input]["password"] == password_input:
                st.session_state["authenticated"] = True
                st.session_state["logged_user"] = username_input
                st.success(f"مرحباً بك {username_input}! جاري تحميل لوحة التحكم...")
                st.rerun()
            else:
                st.error("اسم المستخدم أو رمز الدخول غير صحيح!")

else:
    # ==========================================
    # 2. الواجهة الرئيسية بعد تسجيل الدخول
    # ==========================================
    logged_user_key = st.session_state["logged_user"]
    user_folder_name = USER_CREDENTIALS[logged_user_key]["folder"]

    col_user, col_logout = st.columns([3, 1])
    with col_user:
        st.write(f"👤 **المستخدم الحالي:** `{logged_user_key}` ({user_folder_name})")
    with col_logout:
        if st.button("تسجيل الخروج"):
            st.session_state["authenticated"] = False
            st.session_state["logged_user"] = None
            st.rerun()

    st.title("نظام معالجة وتحديث نماذج KYC الذكي")
    st.write("قم بتحديد النموذج المطلوب وإدخال بيانات المحافظة والنشاط التجاري بكل دقة.")

    def get_user_templates(folder_name):
        user_path = os.path.join(TEMPLATES_DIR, folder_name)
        if os.path.exists(user_path):
            return [f for f in os.listdir(user_path) if f.endswith(".docx") and not f.startswith("~$")]
        return []

    template_files = get_user_templates(user_folder_name)

    if not template_files:
        st.warning(f"لا توجد ملفات Word متوفرة في مجلد المستخدم ({user_folder_name}).")
    else:
        template_options = ["-- Select Template --"] + template_files
        
        selected_template = st.selectbox(
            "اختر النموذج المطلوب معالجته:", 
            template_options, 
            key=f"template_select_{st.session_state['form_reset_key']}"
        )

        if selected_template == "-- Select Template --":
            st.info("👆 يرجى اختيار النموذج من القائمة أعلاه لبدء إدخال البيانات.")
            st.stop()

        st.markdown("---")
        st.subheader("إدخال البيانات المتغيرة والنشاط التجاري")

        kyc_status = st.text_input("KYC STATUS (مثال: Approved / Pending)")
        region = st.text_input("Region (أدخل اسم المحافظة يدوياً)")

        office_name_en = st.text_input("Office Name EN")
        office_name_ar = st.text_input("Office Name Ar")
        
        merchant_id_val = st.number_input("Merchant_ID", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        merchant_id = str(merchant_id_val) if merchant_id_val is not None else ""

        # قائمة الأنشطة المفصولة إنجليزياً ومعروضة مع العربي الكامل
        segment_options = AVAILABLE_SEGMENTS + ["Other"]
        selected_dropdown_label = st.selectbox("Select Business Activity (Segment):", segment_options)

        if selected_dropdown_label == "Other":
            nature_of_activity = st.text_input("Enter Custom Activity (أدخل النشاط يدوياً):")
            mcc_val_input = st.text_input("Enter MCC Code (أدخل رمز MCC المكون من 4 أرقام):", value="")
            mcc = mcc_val_input
        else:
            # استخراج النص الإنجليزي النقي والمفرد للملف النهائي مع مطابقة مرنة
            raw_activity = DISPLAY_MAP.get(selected_dropdown_label, selected_dropdown_label)
            nature_of_activity = raw_activity
            
            # بحث مرن عن رمز MCC بغض النظر عن حالة الأحرف
            auto_mcc = MCC_MAPPING.get(nature_of_activity, MCC_MAPPING.get(nature_of_activity.upper(), MCC_MAPPING.get(nature_of_activity.lower(), "")))
            
            mcc_val_input = st.text_input("MCC (Auto-fetched for Selected Activity)", value=auto_mcc)
            mcc = mcc_val_input

        def validate_inputs(status, region_val, name_en, name_ar, m_id, mcc_code):
            errors = []
            if not status.strip():
                errors.append("حقل KYC STATUS إلزامي.")
            if not region_val.strip():
                errors.append("حقل Region (المحافظة) إلزامي.")
            if not name_en.strip() or not name_ar.strip():
                errors.append("يجب إدخال اسم المكتب باللغتين العربية والإنجليزية بدقة.")
            if not m_id.strip() or len(m_id) < 5:
                errors.append("رقم الـ Merchant ID غير صحيح أو ناقص.")
            if not mcc_code.strip() or len(mcc_code) != 4:
                errors.append("رمز الـ MCC يجب أن يتكون من 4 أرقام صحيحة.")
            return errors

        data = {
            "KYC_STATUS": kyc_status,
            "KYC_SUBJECT": kyc_status,
            "REGION": region,
            "OFFICE_NAME_EN": office_name_en,
            "CLIENT_NAME_HEADER": office_name_en,
            "OFFICE_NAME_AR": office_name_ar,
            "CLIENT_NAME_CELL": office_name_ar,
            "MERCHANT_ID": merchant_id,
            "CLIENT_ID": merchant_id,
            "NATURE_OF_ACTIVITY": nature_of_activity,
            "MCC": mcc
        }

        st.markdown("---")

        def strip_form_fields(doc):
            body = doc.element.body
            for target_tag in ['fldSimple', 'ffData', 'checkBox']:
                for elem in body.xpath(f'//w:{target_tag}'):
                    parent = elem.getparent()
                    if parent is not None:
                        parent.remove(elem)
            
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        xml_str = cell._element.xml.lower()
                        is_yellow = any(col in xml_str for col in ['yellow', 'ffff00', 'ffff99', 'fff2cc'])
                        if cell.text.strip() in ['1', '0']:
                            if not is_yellow:
                                cell.text = ""

        def replace_in_paragraph(paragraph, replacements):
            full_text = paragraph.text
            has_match = False
            for placeholder, value in replacements.items():
                if placeholder in full_text:
                    full_text = full_text.replace(placeholder, value)
                    has_match = True

            if has_match:
                if paragraph.runs:
                    paragraph.runs[0].text = full_text
                    for r in paragraph.runs[1:]:
                        r.text = ""
                else:
                    paragraph.text = full_text

        def replace_placeholders(doc, data):
            strip_form_fields(doc)
            replacements = {f"{{{{{key}}}}}" : val for key, val in data.items() if val}
            if not replacements:
                return
                
            for p in doc.paragraphs:
                replace_in_paragraph(p, replacements)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_in_paragraph(p, replacements)

            for section in doc.sections:
                for p in section.header.paragraphs:
                    replace_in_paragraph(p, replacements)
                for p in section.footer.paragraphs:
                    replace_in_paragraph(p, replacements)

        def convert_to_pdf(docx_path, output_pdf_path):
            abs_docx = os.path.abspath(docx_path)
            abs_pdf = os.path.abspath(output_pdf_path)

            try:
                from docx2pdf import convert
                convert(abs_docx, abs_pdf)
                if os.path.exists(abs_pdf):
                    return True
            except Exception:
                pass

            if sys.platform == "win32":
                try:
                    import pythoncom
                    import win32com.client
                    pythoncom.CoInitialize()
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(abs_docx)
                    doc.Fields.Unlink()
                    doc.SaveAs(abs_pdf, FileFormat=17)
                    doc.Close(False)
                    word.Quit()
                    if os.path.exists(abs_pdf):
                        return True
                except Exception:
                    pass

            try:
                cmd = f'libreoffice --headless --convert-to pdf "{abs_docx}" --outdir "{os.path.dirname(abs_pdf)}"'
                subprocess.run(cmd, shell=True, check=True)
                if os.path.exists(abs_pdf):
                    return True
            except Exception:
                pass

            return False

        if st.button("معالجة وتجهيز المستند (PDF)", type="primary"):
            validation_errors = validate_inputs(kyc_status, region, office_name_en, office_name_ar, merchant_id, mcc)
            
            if validation_errors:
                for err in validation_errors:
                    st.error(err)
                st.warning("⚠️ يرجى تصحيح الأخطاء أعلاه لتجنب رفض المعاملة.")
            else:
                template_path = os.path.join(TEMPLATES_DIR, user_folder_name, selected_template)
                try:
                    doc = Document(template_path)
                    replace_placeholders(doc, data)

                    with tempfile.TemporaryDirectory() as temp_dir:
                        temp_docx = os.path.join(temp_dir, "temp_processed.docx")
                        temp_pdf = os.path.join(temp_dir, "temp_processed.pdf")

                        doc.save(temp_docx)

                        with st.spinner("جاري معالجة المستند وتحويله إلى PDF بدقة عالية... يرجى الانتظار"):
                            success = convert_to_pdf(temp_docx, temp_pdf)

                        if success and os.path.exists(temp_pdf):
                            with open(temp_pdf, "rb") as f:
                                pdf_bytes = f.read()

                            st.success("🎉 تم تحديث المستند وتحويله بصيغة PDF بنجاح تام!")
                            
                            st.download_button(
                                label="📥 تنزيل المستند بصيغة PDF النهائي",
                                data=pdf_bytes,
                                file_name=f"{logged_user_key}_{os.path.splitext(selected_template)[0]}.pdf",
                                mime="application/pdf",
                                on_click=lambda: (st.session_state.update({"form_reset_key": st.session_state["form_reset_key"] + 1}), st.rerun())
                            )
                        else:
                            st.error("❌ تعذر تحويل الملف إلى PDF لعدم توفر محرك تحويل على البيئة الحالية.")

                except Exception as e:
                    st.error(f"حدث خطأ أثناء معالجة المستند: {e}")
