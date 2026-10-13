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
        "folder": "سيف"           # اسم المجلد الخاص به
    }
}

# تحديد مسار المجلد الأساسي
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# ==========================================
# 📊 تحميل قواميس الـ MCC من ملف الإكسل المعتمد
# ==========================================
@st.cache_data
def load_mcc_database():
    mcc_file_path = os.path.join(BASE_DIR, "MCC NEW.xlsx")
    mcc_dict = {}
    segments_list = []
    
    if os.path.exists(mcc_file_path):
        try:
            df = pd.read_excel(mcc_file_path, sheet_name="Sheet1")
            valid_df = df.dropna(subset=['MCC'])
            for _, row in valid_df.iterrows():
                seg_ar = str(row['Segment']).strip()
                mcc_val = str(row['MCC']).strip()
                # تنظيف الرقم إذا كان بصيغة عشرية
                if mcc_val.endswith('.0'):
                    mcc_val = mcc_val[:-2]
                mcc_dict[seg_ar] = mcc_val
                segments_list.append(seg_ar)
        except Exception as e:
            st.error(f"خطأ في قراءة ملف الـ MCC: {e}")
            
    return mcc_dict, sorted(list(set(segments_list)))

MCC_MAPPING, AVAILABLE_SEGMENTS = load_mcc_database()

# إدارة حالة تسجيل الدخول عبر Session State
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "logged_user" not in st.session_state:
    st.session_state["logged_user"] = None


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

    # شريط علوي لإظهار المستخدم الحالي وزر الخروج
    col_user, col_logout = st.columns([3, 1])
    with col_user:
        st.write(f"👤 **المستخدم الحالي:** `{logged_user_key}` ({user_folder_name})")
    with col_logout:
        if st.button("تسجيل الخروج"):
            st.session_state["authenticated"] = False
            st.session_state["logged_user"] = None
            st.rerun()

    st.title("نظام معالجة وتحديث نماذج KYC الذكي")
    st.write("قم بتحديد النموذج المطلوب والنشاط التجاري المعتمد لضمان دقة رموز الـ MCC وعدم الرفض.")

    # جلب النماذج من مجلد المستخدم
    def get_user_templates(folder_name):
        user_path = os.path.join(TEMPLATES_DIR, folder_name)
        if os.path.exists(user_path):
            return [f for f in os.listdir(user_path) if f.endswith(".docx") and not f.startswith("~$")]
        return []

    template_files = get_user_templates(user_folder_name)

    if not template_files:
        st.warning(f"لا توجد ملفات Word متوفرة في مجلد المستخدم ({user_folder_name}).")
    else:
        selected_template = st.selectbox("اختر النموذج المطلوب معالجته:", template_files)

        st.markdown("---")
        st.subheader("إدخال البيانات المتغيرة والنشاط التجاري")

        kyc_status = st.text_input("KYC STATUS (مثال: Approved / Pending)")
        region = st.text_input("Region (مثال: Baghdad / Karbala)")
        office_name_en = st.text_input("Office Name EN")
        office_name_ar = st.text_input("Office Name Ar")
        
        merchant_id_val = st.number_input("Merchant_ID", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        merchant_id = str(merchant_id_val) if merchant_id_val is not None else ""

        # اختيار النشاط التجاري من القائمة المنسدلة المربوطة بملف الإكسل
        if AVAILABLE_SEGMENTS:
            selected_segment = st.selectbox("اختر النشاط التجاري (Segment):", AVAILABLE_SEGMENTS)
            nature_of_activity = selected_segment
            # استخراج الرمز التلقائي المطابق للنشاط المختار
            auto_mcc = MCC_MAPPING.get(selected_segment, "")
        else:
            nature_of_activity = st.text_input("Nature of Activity")
            auto_mcc = ""

        # عرض رمز الـ MCC التلقائي المستخرج من الملف مع إمكانية التأكيد
        mcc = st.text_input("MCC (تم تعبئته تلقائياً من الملف المعتمد)", value=auto_mcc)

        # ==========================================
        # 🛡️ طبقة التحقق الذكي (Smart Validation Layer)
        # ==========================================
        def validate_inputs(status, region_val, name_en, name_ar, m_id, mcc_code):
            errors = []
            if not status.strip():
                errors.append("حقل KYC STATUS إلزامي.")
            if not region_val.strip():
                errors.append("حقل Region إلزامي.")
            if not name_en.strip() or not name_ar.strip():
                errors.append("يجب إدخال اسم المكتب باللغتين العربية والإنجليزية بدقة.")
            if not m_id.strip() or len(m_id) < 5:
                errors.append("رقم الـ Merchant ID غير صحيح أو ناقص.")
            if not mcc_code.strip() or len(mcc_code) != 4:
                errors.append("رمز الـ MCC يجب أن يتكون من 4 أرقام صحيحة (تم جلبها من الملف).")
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

        def convert_to_pdf(docx_path, output_pdf_path, temp_dir):
            abs_docx = os.path.abspath(docx_path)
            abs_pdf = os.path.abspath(output_pdf_path)

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
                cmd = f'libreoffice --headless --convert-to "pdf:writer_pdf_Export:{{\"ExportFormFields\":{{\"type\":\"boolean\",\"value\":\"false\"}}}}" "{abs_docx}" --outdir "{temp_dir}"'
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

                        with st.spinner("جاري معالجة المستند وتحويله إلى PDF... يرجى الانتظار"):
                            success = convert_to_pdf(temp_docx, temp_pdf, temp_dir)

                        if success and os.path.exists(temp_pdf):
                            with open(temp_pdf, "rb") as f:
                                pdf_bytes = f.read()

                            st.success("🎉 تم تحديث المستند وتحويله إلى PDF بالرمز الصحيح وبدون أخطاء!")
                            st.download_button(
                                label="تنزيل المستند بصيغة PDF",
                                data=pdf_bytes,
                                file_name=f"{logged_user_key}_{os.path.splitext(selected_template)[0]}.pdf",
                                mime="application/pdf"
                            )
                        else:
                            st.warning("تعذر تحويل الملف إلى PDF تلقائياً، يمكنك تنزيله بصيغة DOCX:")
                            bio = io.BytesIO()
                            doc.save(bio)
                            bio.seek(0)
                            st.download_button(
                                label="تنزيل المستند بصيغة DOCX",
                                data=bio,
                                file_name=f"{logged_user_key}_{selected_template}",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

                except Exception as e:
                    st.error(f"حدث خطأ أثناء معالجة المستند: {e}")
