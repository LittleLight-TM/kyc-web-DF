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

# قائمة تصنيفات المخاطر
RISK_CATEGORIES = ["cash in", "cash out", "pos"]

# تحديد مسار المجلد الأساسي
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# ==========================================
# 📊 فصل الأنشطة الإنجليزية المدمجة وتنسيق العرض
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
                seg_ar = str(row['Segment']).strip() if pd.notna(row['Segment']) else ""
                seg_en = str(row['Segment.1']).strip() if pd.notna(row['Segment.1']) else ""
                mcc_val = str(row['MCC']).strip()
                if mcc_val.endswith('.0'):
                    mcc_val = mcc_val[:-2]
                
                clean_seg_ar = seg_ar.replace('(', '').replace(')', '').strip() if seg_ar and seg_ar.lower() != 'nan' else ""
                
                if seg_en and seg_en.lower() != 'nan':
                    clean_seg_en = seg_en.replace('(', '').replace(')', '').strip()
                    parts_en = [p.strip() for p in clean_seg_en.replace('+', ',').replace('/', ',').split(',')]
                    
                    for part_en in parts_en:
                        if part_en:
                            mcc_dict[part_en] = mcc_val
                            mcc_dict[part_en.upper()] = mcc_val
                            mcc_dict[part_en.lower()] = mcc_val
                            
                            if clean_seg_ar:
                                label = f"{part_en.upper()} — ({clean_seg_ar})"
                            else:
                                label = f"{part_en.upper()}"
                                
                            display_map[label] = part_en
                            dropdown_labels.add(label)
                elif clean_seg_ar:
                    mcc_dict[clean_seg_ar] = mcc_val
                    label = f"{clean_seg_ar}"
                    display_map[label] = clean_seg_ar
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
    st.write("قم بتحديد درجة المخاطرة والنموذج المطلوب، ثم أدخل بيانات المحافظة والنشاط التجاري بدقة.")

    def get_user_templates(folder_name):
        user_path = os.path.join(TEMPLATES_DIR, folder_name)
        if os.path.exists(user_path):
            return [f for f in os.listdir(user_path) if f.endswith(".docx") and not f.startswith("~$")]
        return []

    template_files = get_user_templates(user_folder_name)

    if not template_files:
        st.warning(f"لا توجد ملفات Word متوفرة في مجلد المستخدم ({user_folder_name}).")
    else:
        # ==========================================
        # 🎯 خيار Risk Select وتصفية النماذج بدقة
        # ==========================================
        risk_category = st.selectbox(
            "اختر نوع المخاطرة (Risk Select):",
            ["-- Select Risk Category --"] + RISK_CATEGORIES,
            key=f"risk_select_{st.session_state['form_reset_key']}"
        )

        if risk_category == "-- Select Risk Category --":
            st.info("👆 يرجى اختيار نوع المخاطرة (Risk Select) لعرض النماذج الخاصة بها.")
            st.stop()

        # تصفية النماذج بناءً على الكلمات المفتاحية لكل تصنيف مخاطر
        filtered_templates = []
        for f in template_files:
            f_lower = f.lower()
            if risk_category == "cash in":
                if "cash in" in f_lower or "cashin" in f_lower:
                    filtered_templates.append(f)
            elif risk_category == "cash out":
                if "cashout" in f_lower or "cash out" in f_lower:
                    filtered_templates.append(f)
            elif risk_category == "pos":
                if ("kyc" in f_lower or "pos" in f_lower) and "cash in" not in f_lower and "cashin" not in f_lower and "cashout" not in f_lower and "cash out" not in f_lower:
                    filtered_templates.append(f)

        if not filtered_templates:
            st.warning("معدل المخاطرة المحدد لا يحتوي على ملفات مطابقة في مجلدك.")
            st.stop()

        selected_template = st.selectbox(
            f"اختر النموذج المطلوب لـ ({risk_category}):", 
            ["-- Select Template --"] + filtered_templates, 
            key=f"template_select_{st.session_state['form_reset_key']}"
        )

        if selected_template == "-- Select Template --":
            st.info("👆 يرجى اختيار النموذج المحدد لبدء إدخال البيانات.")
            st.stop()

        st.markdown("---")
        st.subheader("إدخال البيانات المتغيرة والنشاط التجاري")

        kyc_status = st.text_input("KYC STATUS (مثال: Approved / Pending)")
        region = st.text_input("Region (أدخل اسم المحافظة يدوياً)")

        office_name_en = st.text_input("Office Name EN")
        office_name_ar = st.text_input("Office Name Ar")
        
        merchant_id_val = st.number_input("Merchant_ID", min_value=0, step=1, value=None, placeholder="أرقام فقط")
        merchant_id = str(merchant_id_val) if merchant_id_val is not None else ""

        segment_options = AVAILABLE_SEGMENTS + ["Other"]
        selected_dropdown_label = st.selectbox("Select Business Activity (Segment):", segment_options)

        if selected_dropdown_label == "Other":
            nature_of_activity = st.text_input("Enter Custom Activity (أدخل النشاط يدوياً):")
            mcc_val_input = st.text_input("Enter MCC Code (أدخل رمز MCC المكون من 4 أرقام):", value="")
            mcc = mcc_val_input
        else:
            raw_activity = DISPLAY_MAP.get(selected_dropdown_label, selected_dropdown_label)
            nature_of_activity = raw_activity
            
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
            "OFFICE_NAME_AR": office
