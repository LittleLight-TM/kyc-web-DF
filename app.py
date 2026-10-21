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
                if "cash in"
