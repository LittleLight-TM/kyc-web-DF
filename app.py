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
st.write("قم باختيار النموذج وإدخال البيانات المطلوبة ثم تحميل المستند الناتج بصيغة PDF مباشرة.")

# تحديد مسار المجلد سواء عند التشغيل العادي أو داخل PyInstaller
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

def get_templates():
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
    return [f for f in os.listdir(TEMPLATES_DIR) if f.endswith(".docx") and not f.startswith("~$")]

template_files = get_templates()

if not template_files:
    st.error("لم يتم العثور على أي ملفات Word داخل مجلد templates!")
else:
    selected_template = st.selectbox("اختر النموذج المطلوب معالجته:", template_files)

    st.markdown("---")
    st.subheader("إدخال البيانات المتغيرة")

    # حقول البيانات بترتيب متتالي لضمان التنقل السلس عبر زر Tab
    kyc_subject = st.text_input("KYC Subject / Main Dealer")
    region = st.text_input("Region")
    client_name_english = st.text_input("Client Name (ENGLISH)")
    client_name_arabic = st.text_input("Client Name (ARABIC)")
    
    # حقل Client ID مقيد بالأرقام فقط
    client_id_val = st.number_input("Client ID", min_value=0, step=1, value=None, placeholder="أرقام فقط")
    client_id = str(client_id_val) if client_id_val is not None else ""

    nature_of_activity = st.text_input("Nature of Activity")
    
    # حقل MCC مقيد بالأرقام فقط
    mcc_val = st.number_input("MCC", min_value=0, step=1, value=None, placeholder="أرقام فقط")
    mcc = str(mcc_val) if mcc_val is not None else ""

    # قاموس البيانات المستبدلة
    data = {
        "KYC_SUBJECT": kyc_subject,
        "REGION": region,
        "CLIENT_NAME_HEADER": client_name_english,
        "CLIENT_NAME_CELL": client_name_arabic,
        "CLIENT_ID": client_id,
        "NATURE_OF_ACTIVITY": nature_of_activity,
        "MCC": mcc
    }

    st.markdown("---")

    def replace_in_paragraphs(paragraphs, replacements):
        for p in paragraphs:
            for placeholder, value in replacements.items():
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, value)

    def replace_placeholders(doc, data):
        replacements = {f"{{{{{key}}}}}" : val for key, val in data.items() if val}
        if not replacements:
            return
            
        # 1. استبدال في الفقرات العادية
        replace_in_paragraphs(doc.paragraphs, replacements)

        # 2. استبدال داخل الجداول
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs, replacements)

        # 3. استبدال داخل الهيدر والفوتر (Header & Footer)
        for section in doc.sections:
            replace_in_paragraphs(section.header.paragraphs, replacements)
            replace_in_paragraphs(section.footer.paragraphs, replacements)

    def convert_to_pdf(docx_path, output_pdf_path):
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(output_pdf_path)

        # 1. التحويل السريع المباشر عبر MS Word (Windows)
        if sys.platform == "win32":
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                
                doc = word.Documents.Open(abs_docx)
                doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
                doc.Close()
                word.Quit()
                
                if os.path.exists(abs_pdf):
                    return True
            except Exception:
                pass

            # تجربة احتياطية عبر docx2pdf
            try:
                from docx2pdf import convert
                convert(abs_docx, abs_pdf)
                if os.path.exists(abs_pdf):
                    return True
            except Exception:
                pass

        # 2. تجربة التحويل عبر LibreOffice
        try:
            out_dir = os.path.dirname(abs_pdf)
            cmd = f'libreoffice --headless --convert-to pdf "{abs_docx}" --outdir "{out_dir}"'
            subprocess.run(cmd, shell=True, check=True)
            if os.path.exists(abs_pdf):
                return True
        except Exception:
            pass

        return False

    if st.button("معالجة وتجهيز المستند (PDF)", type="primary"):
        template_path = os.path.join(TEMPLATES_DIR, selected_template)
        try:
            doc = Document(template_path)
            replace_placeholders(doc, data)

            with tempfile.TemporaryDirectory() as temp_dir:
                temp_docx = os.path.join(temp_dir, "temp_processed.docx")
                temp_pdf = os.path.join(temp_dir, "temp_processed.pdf")

                doc.save(temp_docx)

                # إضافة مؤشر الانتظار التفاعلي أثناء المعالجة
                with st.spinner("جاري معالجة المستند وتحويله إلى PDF... يرجى الانتظار"):
                    success = convert_to_pdf(temp_docx, temp_pdf)

                if success and os.path.exists(temp_pdf):
                    with open(temp_pdf, "rb") as f:
                        pdf_bytes = f.read()

                    st.success("تم تحديث المستند وتحويله إلى PDF بنجاح!")
                    st.download_button(
                        label="تنزيل المستند بصيغة PDF",
                        data=pdf_bytes,
                        file_name=f"Updated_{os.path.splitext(selected_template)[0]}.pdf",
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
                        file_name=f"Updated_{selected_template}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة المستند: {e}")