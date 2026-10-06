import os
import io
import sys
import subprocess
import tempfile
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

# 1. جلب قائمة المستخدمين
def get_users():
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
    return [d for d in os.listdir(TEMPLATES_DIR) if os.path.isdir(os.path.join(TEMPLATES_DIR, d))]

# 2. جلب النماذج المتاحة للمستخدم المحدد
def get_user_templates(user_folder):
    user_path = os.path.join(TEMPLATES_DIR, user_folder)
    if os.path.exists(user_path):
        return [f for f in os.listdir(user_path) if f.endswith(".docx") and not f.startswith("~$")]
    return []

users = get_users()

if not users:
    st.error("لم يتم العثور على أي مجلدات فرعية للمستخدمين داخل مجلد templates!")
else:
    selected_user = st.selectbox("اختر اسم المستخدم:", users)
    template_files = get_user_templates(selected_user)

    if not template_files:
        st.warning(f"لا توجد ملفات Word في مجلد المستخدم ({selected_user}).")
    else:
        selected_template = st.selectbox("اختر النموذج المطلوب معالجته:", template_files)

        st.markdown("---")
        st.subheader("إدخال البيانات المتغيرة")

        kyc_status = st.text_input("KYC STATUS")
        region = st.text_input("Region")
        office_name_en = st.text_input("Office Name EN")
        office_name_ar = st.text_input("Office Name Ar")
        
        merchant_id_val = st.number_input("Merchant_ID", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        merchant_id = str(merchant_id_val) if merchant_id_val is not None else ""

        nature_of_activity = st.text_input("Nature of Activity")
        
        mcc_val = st.number_input("MCC", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        mcc = str(mcc_val) if mcc_val is not None else ""

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

        def convert_checkboxes_to_symbols(doc):
            """تحويل جميع حقول Checkbox و Form Fields التفاعلية إلى رموز نصية ثابته أو تفريغها من الأرقام تلقائياً"""
            body_element = doc.element.body
            
            # 1. البحث عن أوسام الحقول التفاعلية في XML واستبدالها بنص ثابت
            for elem in list(body_element.iter()):
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                
                # المعالجة الخاصة بـ Form Fields
                if tag in ['ffData', 'checkBox', 'fldSimple']:
                    parent = elem.getparent()
                    if parent is not None:
                        # التأكد مما إذا كان الحقل مفعلاً (Checked) أو غير مفعل
                        is_checked = False
                        for child in elem.iter():
                            child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                            if child_tag == 'checked':
                                val = child.get(qn('w:val'), '1')
                                if val in ['1', 'true', 'on']:
                                    is_checked = True
                        
                        # إنشاء عنصر نصي بديل يحتوي على الرمز المناسب
                        new_run = OxmlElement('w:r')
                        new_text = OxmlElement('w:t')
                        new_text.text = "☑" if is_checked else "☐"
                        new_run.append(new_text)
                        
                        parent.replace(elem, new_run)

            # 2. تنظيف نصوص الأرقام 0 و 1 المعلقة داخل خلايا الجداول المظللة
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text.strip()
                        if cell_text in ['0', '1']:
                            # إذا كانت الخلية تحتوي فقط على 0 أو 1، تحويل 1 إلى رمز المربع المظلل وإخفاء 0
                            for p in cell.paragraphs:
                                if p.text.strip() == '1':
                                    p.text = "☑"
                                elif p.text.strip() == '0':
                                    p.text = "☐"

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
            # تحويل الحقول التفاعلية والأرقام برمجياً أولاً
            convert_checkboxes_to_symbols(doc)

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

            # التحويل في بيئة Streamlit Cloud
            try:
                cmd = f'libreoffice --headless --convert-to pdf "{abs_docx}" --outdir "{temp_dir}"'
                subprocess.run(cmd, shell=True, check=True)
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
                        success = convert_to_pdf(temp_docx, temp_pdf, temp_dir)

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
