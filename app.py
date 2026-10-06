import os
import io
import sys
import subprocess
import tempfile
import streamlit as st
from docx import Document

# إعداد الصفحة
st.set_page_config(page_title="نظام تعديل نماذج KYC", page_icon="📝", layout="centered")

st.title("نظام معالجة وتحديث نماذج KYC")
st.write("قم باختيار المستخدم، ثم تحديد النموذج المطلوب وإدخال البيانات المعنية لتجهيز المستند.")

# تحديد مسار المجلد الأساسي
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# 1. جلب قائمة المستخدمين (المجلدات الفرعية داخل templates)
def get_users():
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
    return [d for d in os.listdir(TEMPLATES_DIR) if os.path.isdir(os.path.join(TEMPLATES_DIR, d))]

# 2. جلب النماذج المتاحة للمستخدم المحدد فقط
def get_user_templates(user_folder):
    user_path = os.path.join(TEMPLATES_DIR, user_folder)
    if os.path.exists(user_path):
        return [f for f in os.listdir(user_path) if f.endswith(".docx") and not f.startswith("~$")]
    return []

users = get_users()

if not users:
    st.error("لم يتم العثور على أي مجلدات فرعية للمستخدمين داخل مجلد templates! يرجى إنشاء مجلد باسم كل مستخدم ووضع قوالبه بداخله.")
else:
    # القائمة المنسدلة لاختيار اسم المستخدم
    selected_user = st.selectbox("اختر اسم المستخدم:", users)

    # القائمة المنسدلة لاختيار النموذج المخصص للمستخدم المحدد
    template_files = get_user_templates(selected_user)

    if not template_files:
        st.warning(f"لا توجد ملفات Word في مجلد المستخدم ({selected_user}).")
    else:
        selected_template = st.selectbox("اختر النموذج المطلوب معالجته:", template_files)

        st.markdown("---")
        st.subheader("إدخال البيانات المتغيرة")

        # حقول إدخال البيانات بالتسميات الجديدة
        kyc_status = st.text_input("KYC STATUS")
        region = st.text_input("Region")
        office_name_en = st.text_input("Office Name EN")
        office_name_ar = st.text_input("Office Name Ar")
        
        merchant_id_val = st.number_input("Merchant_ID", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        merchant_id = str(merchant_id_val) if merchant_id_val is not None else ""

        nature_of_activity = st.text_input("Nature of Activity")
        
        mcc_val = st.number_input("MCC", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        mcc = str(mcc_val) if mcc_val is not None else ""

        # قاموس البيانات المستبدلة (يشمل مفاتيح القوالب القديمة والجديدة لضمان التوافق)
        data = {
            "KYC_STATUS": kyc_status,
            "KYC_SUBJECT": kyc_status,  # للتوافق إذا كان القالب يستخدم الاسم القديم
            "REGION": region,
            "OFFICE_NAME_EN": office_name_en,
            "CLIENT_NAME_HEADER": office_name_en,  # للتوافق
            "OFFICE_NAME_AR": office_name_ar,
            "CLIENT_NAME_CELL": office_name_ar,  # للتوافق
            "MERCHANT_ID": merchant_id,
            "CLIENT_ID": merchant_id,  # للتوافق
            "NATURE_OF_ACTIVITY": nature_of_activity,
            "MCC": mcc
        }

        st.markdown("---")

        def replace_in_paragraph(paragraph, replacements):
            # تجميع النص الكامل للفقرة للتعامل مع النصوص المقسمة مثل {{REGION}}
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
            replacements = {f"{{{{{key}}}}}" : val for key, val in data.items() if val}
            if not replacements:
                return
                
            # 1. الاستبدال في الفقرات العامة
            for p in doc.paragraphs:
                replace_in_paragraph(p, replacements)

            # 2. الاستبدال داخل الجداول
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_in_paragraph(p, replacements)

            # 3. الاستبدال داخل الهيدر والفوتر
            for section in doc.sections:
                for p in section.header.paragraphs:
                    replace_in_paragraph(p, replacements)
                for p in section.footer.paragraphs:
                    replace_in_paragraph(p, replacements)

        def convert_to_pdf(docx_path, output_pdf_path):
            abs_docx = os.path.abspath(docx_path)
            abs_pdf = os.path.abspath(output_pdf_path)

            # 1. بيئة Windows (تشغيل محلي عبر MS Word)
            if sys.platform == "win32":
                try:
                    import pythoncom
                    import win32com.client
                    pythoncom.CoInitialize()
                    
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    
                    doc = word.Documents.Open(abs_docx)
                    
                    # فك ارتباط حقول النموذج لمنع تحولها إلى أرقام
                    doc.Fields.Unlink()
                    
                    doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
                    doc.Close(False)
                    word.Quit()
                    
                    if os.path.exists(abs_pdf):
                        return True
                except Exception:
                    pass

                try:
                    from docx2pdf import convert
                    convert(abs_docx, abs_pdf)
                    if os.path.exists(abs_pdf):
                        return True
                except Exception:
                    pass

            # 2. بيئة Streamlit Cloud (Linux / LibreOffice)
            try:
                out_dir = os.path.dirname(abs_pdf)
                # فلتر منع تحويل حقول Form Fields إلى أرقام تفاعلية 0 و 1
                cmd = f'libreoffice --headless --convert-to "pdf:writer_pdf_Export:{{\"ExportFormFields\":{{\"type\":\"boolean\",\"value\":\"false\"}}}}" "{abs_docx}" --outdir "{out_dir}"'
                subprocess.run(cmd, shell=True, check=True)
                if os.path.exists(abs_pdf):
                    return True
            except Exception:
                # محاولة تحويل احتياطية
                try:
                    out_dir = os.path.dirname(abs_pdf)
                    cmd_fallback = f'libreoffice --headless --convert-to pdf "{abs_docx}" --outdir "{out_dir}"'
                    subprocess.run(cmd_fallback, shell=True, check=True)
                    if os.path.exists(abs_pdf):
                        return True
                except Exception:
                    pass

            return False

        if st.button("معالجة وتجهيز المستند (PDF)", type="primary"):
            template_path = os.path.join(TEMPLATES_DIR, selected_user, selected_template)
            try:
                doc = Document(template_path)
                replace_placeholders(doc, data)

                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_docx = os.path.join(temp_dir, "temp_processed.docx")
                    temp_pdf = os.path.join(temp_dir, "temp_processed.pdf")

                    doc.save(temp_docx)

                    with st.spinner("جاري معالجة المستند وتحويله إلى PDF... يرجى الانتظار"):
                        success = convert_to_pdf(temp_docx, temp_pdf)

                    if success and os.path.exists(temp_pdf):
                        with open(temp_pdf, "rb") as f:
                            pdf_bytes = f.read()

                        st.success("تم تحديث المستند وتحويله إلى PDF بنجاح!")
                        st.download_button(
                            label="تنزيل المستند بصيغة PDF",
                            data=pdf_bytes,
                            file_name=f"{selected_user}_{os.path.splitext(selected_template)[0]}.pdf",
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
                            file_name=f"{selected_user}_{selected_template}",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة المستند: {e}")
