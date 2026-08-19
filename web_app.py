import os
import re
import io
import json
import base64
import datetime
import requests
from PIL import Image
import openpyxl
import streamlit as st

# ==========================================
# 🌟 Page Configuration & Title
# ==========================================
st.set_page_config(
    page_title="User Access Automation | MGC-Asia",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 🏢 Master Data & Options
# ==========================================
COMPANY_OPTIONS = [
    "",
    "Master Car Rental Co., Ltd.",
    "Millennium Auto Group Co., Ltd.",
    "US Motorbike Co., Ltd.",
    "I24 Co.,Ltd.",
    "Millennium Group Corporation (Asia) Public Company Limited"
]

BRANCH_OPTIONS = [
    "", "กรุงเทพมหานคร", "พระราม 3", "พระราม 4", "ลาดพร้าว", 
    "อุดรธานี", "ภูเก็ต", "หาดใหญ่", "อุบลราชธานี", 
    "สุราษฎร์ธานี", "พัทยา", "เชียงใหม่", "สยามพารากอน", "ไอคอนสยาม"
]

COMPANY_BU_MAPPING = {
    "mastercarrental.com": {"Company": "Master Car Rental Co., Ltd.", "BU": "MCR"},
    "i24.co.th": {"Company": "I24 Co.,Ltd.", "BU": "I24"},
    "bmw-millenniumauto.com": {"Company": "Millennium Auto Group Co., Ltd.", "BU": "BMW"},
    "millenniumauto.co.th": {"Company": "Millennium Auto Group Co., Ltd.", "BU": "BMW"},
    "usmotorbike.com": {"Company": "US Motorbike Co., Ltd.", "BU": "Harley"},
    "mgc-asia.com": {"Company": "", "BU": ""}
}

# ==========================================
# 🛠️ Helper & Password Functions
# ==========================================
def generate_vsm_password(email):
    if not email:
        return "p@ssw0rd***"
    local_part = email.split('@')[0] if '@' in email else email
    letters = re.findall(r'[a-zA-Z]', local_part)
    if len(letters) >= 3:
        prefix = "".join(letters[:3]).lower()
    else:
        prefix = "".join(letters).lower().ljust(3, 'x')
    return f"p@ssw0rd{prefix}"

def generate_forma_password(user_login):
    if not user_login:
        return "p@ssw0rd***"
    letters = re.findall(r'[a-zA-Z]', user_login)
    if len(letters) >= 3:
        prefix = "".join(letters[:3]).lower()
    else:
        prefix = "".join(letters).lower().ljust(3, 'x')
    return f"p@ssw0rd{prefix}"

def get_default_password_for_app(app_name, user_id, email, extracted_pwd=""):
    if app_name == "Pandora" and extracted_pwd:
        return extracted_pwd
    elif app_name == "VSM":
        return generate_vsm_password(email)
    elif app_name in ["Forma", "Pandora"]:
        return generate_forma_password(user_id)
    elif app_name == "Red plate":
        return "Init123456"
    else:
        return user_id

def derive_company_and_bu(company_input, email_input):
    email_lower = (email_input or "").lower()
    comp_lower = (company_input or "").lower()
    
    if comp_lower:
        if "master car" in comp_lower or "mcr" in comp_lower or "mastercar" in comp_lower:
            return "Master Car Rental Co., Ltd.", "MCR"
        elif "millennium auto" in comp_lower or "bmw" in comp_lower:
            return "Millennium Auto Group Co., Ltd.", "BMW"
        elif "us motorbike" in comp_lower or "harley" in comp_lower:
            return "US Motorbike Co., Ltd.", "Harley"
        elif "i24" in comp_lower:
            return "I24 Co.,Ltd.", "I24"
        elif "millennium group corporation" in comp_lower or "mgc" in comp_lower:
            return "Millennium Group Corporation (Asia) Public Company Limited", "MGC"
            
    if '@' in email_lower:
        domain = email_lower.split('@')[-1].strip()
        for dom_key, info in COMPANY_BU_MAPPING.items():
            if dom_key in domain:
                return info["Company"], info["BU"]
                
    return company_input if company_input else "", ""

def fix_email_domain(email_str):
    if not email_str or '@' not in email_str:
        return email_str
    local_part, domain = email_str.split('@', 1)
    domain_lower = domain.lower()
    if 'mastercar' in domain_lower:
        domain = 'mastercarrental.com'
    elif 'i2a' in domain_lower or 'i24' in domain_lower:
        domain = 'i24.co.th'
    elif 'millennium' in domain_lower or 'bmw' in domain_lower:
        domain = 'bmw-millenniumauto.com'
    elif 'mgc' in domain_lower or 'asia' in domain_lower:
        domain = 'mgc-asia.com'
    return f"{local_part}@{domain}"

# ==========================================
# 🤖 Vision AI Extraction
# ==========================================
def extract_with_gemini_vision(img, api_key):
    clean_key = api_key.strip()
    if not clean_key:
        return None, "กรุณากรอก Gemini API Key ในแถบตั้งค่าด้านซ้าย"
        
    prompt = """คุณคือผู้เชี่ยวชาญการอ่านข้อมูลฟอร์มระบบ VSM, E-Travelling, Forma, Red plate (Update User Info), และระบบ Pandora ของ MGC-Asia
จงวิเคราะห์ภาพแคปเจอร์นี้ และตอบเป็น JSON บริสุทธิ์เท่านั้น (ไม่ต้องมี markdown backticks) ในรูปแบบดังนี้:
{
  "Application": "Pandora" หรือ "Red plate" หรือ "Forma" หรือ "VSM" หรือ "E-Travelling",
  "App user ID": "รหัสผู้ใช้ สำหรับ Pandora หรือ Username สำหรับ Red plate หรือ Login สำหรับ Forma หรือ User Login สำหรับ VSM หรือ Employee Code สำหรับ E-Travelling",
  "Password": "รหัสผ่าน สำหรับ Pandora (เช่น p@ssw0rdcha) (ถ้าไม่มีให้เป็นว่างเปล่า \"\")",
  "Full Name Eng": "ชื่อ-นามสกุลภาษาอังกฤษ ถ้าเป็นภาษาอังกฤษล้วน (ตัดคำนำหน้าออก) ถ้าไม่มีให้เป็น \"\"",
  "Full Name Thai": "ชื่อ-นามสกุลภาษาไทย สำหรับ Red plate ให้นำ First Name เว้นวรรค Last Name (หรือ Display Name เช่น กัณฑิชา ลมหวล) มาใส่ช่องนี้ หรือชื่อไทยสำหรับระบบอื่น (ตัดคำนำหน้าออก) ถ้าไม่มีให้เป็น \"\"",
  "Email": "Email (ถ้าไม่มีในภาพให้เป็นว่างเปล่า \"\")",
  "Position": "กลุ่มผู้ใช้ สำหรับ Pandora (เช่น Accounting + Price) หรือ Role สำหรับ Red plate หรือ Position/User Type สำหรับระบบอื่น",
  "Company": "ชื่อบริษัท (ถ้าไม่มีในภาพให้เป็นว่างเปล่า \"\")",
  "Branch": "ชื่อสาขาภาษาไทย (ถ้าไม่มีในภาพให้เป็นว่างเปล่า \"\")"
}"""

    last_err = ""

    # 1. ลองใช้ google.genai SDK
    try:
        from google import genai
        from google.genai import types
        
        client = genai.Client(api_key=clean_key)
        models_to_try = ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-3.6-flash']
        for m_name in models_to_try:
            try:
                response = client.models.generate_content(
                    model=m_name,
                    contents=[prompt, img],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                raw_text = response.text.strip()
                raw_text = re.sub(r'^```json\s*', '', raw_text, flags=re.IGNORECASE)
                raw_text = re.sub(r'```$', '', raw_text).strip()
                parsed_data = json.loads(raw_text)
                
                app_name = parsed_data.get("Application", "E-Travelling")
                user_id = str(parsed_data.get("App user ID", "")).strip()
                if app_name == "VSM" and user_id:
                    user_id = user_id.replace('.', '')
                    if len(user_id) > 3:
                        user_id = user_id[:-3] + '.' + user_id[-3:]
                    parsed_data["App user ID"] = user_id
                elif app_name == "Red plate":
                    pos = str(parsed_data.get("Position", "")).strip()
                    if "sales person" in pos.lower():
                        parsed_data["Position"] = "Sales Consultant"
                    nth = str(parsed_data.get("Full Name Thai", "")).strip()
                    neng = str(parsed_data.get("Full Name Eng", "")).strip()
                    if not nth and any('\u0e00' <= c <= '\u0e7f' for c in neng):
                        parsed_data["Full Name Thai"] = neng
                        parsed_data["Full Name Eng"] = ""
                    elif nth and not any('\u0e00' <= c <= '\u0e7f' for c in nth) and not neng:
                        parsed_data["Full Name Eng"] = nth
                        parsed_data["Full Name Thai"] = ""
                
                email = str(parsed_data.get("Email", "")).strip()
                if email:
                    parsed_data["Email"] = fix_email_domain(email)
                return parsed_data, ""
            except Exception as ex:
                last_err = str(ex)
                continue
    except Exception as g_err:
        last_err = str(g_err)

    # 2. REST API Fallback (HTTP Direct)
    try:
        buffered = io.BytesIO()
        img.convert('RGB').save(buffered, format="JPEG", quality=85)
        img_b64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
        
        endpoints = [
            ("v1beta", "gemini-3.6-flash"),
            ("v1beta", "gemini-3.5-flash"),
            ("v1beta", "gemini-flash-latest")
        ]
        
        for api_ver, model_name in endpoints:
            url = f"https://generativelanguage.googleapis.com/{api_ver}/models/{model_name}:generateContent?key={clean_key}"
            headers = {"Content-Type": "application/json", "x-goog-api-key": clean_key}
                
            payload = {
                "contents": [
                    {
                        "parts": [
                            {"text": prompt},
                            {
                                "inline_data": {
                                    "mime_type": "image/jpeg",
                                    "data": img_b64
                                }
                            }
                        ]
                    }
                ],
                "generationConfig": {
                    "response_mime_type": "application/json"
                }
            }
            
            res = requests.post(url, json=payload, headers=headers, timeout=25)
            if res.status_code == 200:
                result_json = res.json()
                raw_text = result_json['candidates'][0]['content']['parts'][0]['text'].strip()
                raw_text = re.sub(r'^```json\s*', '', raw_text, flags=re.IGNORECASE)
                raw_text = re.sub(r'```$', '', raw_text).strip()
                
                parsed_data = json.loads(raw_text)
                
                app_name = parsed_data.get("Application", "E-Travelling")
                user_id = str(parsed_data.get("App user ID", "")).strip()
                if app_name == "VSM" and user_id:
                    user_id = user_id.replace('.', '')
                    if len(user_id) > 3:
                        user_id = user_id[:-3] + '.' + user_id[-3:]
                elif app_name == "Red plate":
                    pos = str(parsed_data.get("Position", "")).strip()
                    if "sales person" in pos.lower():
                        parsed_data["Position"] = "Sales Consultant"
                    nth = str(parsed_data.get("Full Name Thai", "")).strip()
                    neng = str(parsed_data.get("Full Name Eng", "")).strip()
                    if not nth and any('\u0e00' <= c <= '\u0e7f' for c in neng):
                        parsed_data["Full Name Thai"] = neng
                        parsed_data["Full Name Eng"] = ""
                    elif nth and not any('\u0e00' <= c <= '\u0e7f' for c in nth) and not neng:
                        parsed_data["Full Name Eng"] = nth
                        parsed_data["Full Name Thai"] = ""
                    
                email = str(parsed_data.get("Email", "")).strip()
                if email:
                    parsed_data["Email"] = fix_email_domain(email)
                    
                return parsed_data, ""
            else:
                try:
                    err_msg = res.json().get('error', {}).get('message', res.text[:120])
                except Exception:
                    err_msg = res.text[:120]
                last_err = f"Google API Status {res.status_code}: {err_msg}"
    except Exception as ex:
        last_err = str(ex)

    if "invalid authentication" in last_err.lower() or "401" in last_err or "api key not valid" in last_err.lower():
        hint = " (หมายเหตุ: คีย์จาก Google AI Studio จะขึ้นต้นด้วย AIzaSy... กรุณากดปุ่ม 'รับ API Key ฟรี' ที่ aistudio.google.com/app/apikey เพื่อรับคีย์ที่ถูกต้อง)"
    else:
        hint = ""

    return None, f"ไม่สามารถสกัดข้อมูลจากภาพได้: {last_err}{hint}"

def process_mapping(extracted_data):
    email = extracted_data.get("Email", "")
    comp_input = extracted_data.get("Company", "")
    company, bu = derive_company_and_bu(comp_input, email)

    branch_eng = extracted_data.get("Branch", "")
    now = datetime.datetime.now()
    current_date = datetime.date(now.year, now.month, now.day)

    return {
        "NO": 1,
        "Application": extracted_data.get("Application", "E-Travelling"),
        "App user ID": extracted_data.get("App user ID", ""),
        "Full Name Eng": extracted_data.get("Full Name Eng", ""),
        "Full Name Thai": extracted_data.get("Full Name Thai", ""),
        "Email": email,
        "Position": extracted_data.get("Position", ""),
        "Company": company,
        "BU": bu,
        "Role ID": "",
        "Branch": branch_eng,
        "x": "",
        "Type": "",
        "Status": "Active", 
        "Create date": current_date,
        "Disable date": "",
        "Remark": ""
    }

def export_to_excel(excel_row, excel_filename):
    if not excel_filename:
        excel_filename = "Template Column Excel Summary User.xlsx"
    try:
        if os.path.exists(excel_filename):
            wb = openpyxl.load_workbook(excel_filename)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            headers = list(excel_row.keys())
            ws.append(headers)

        header_row = [cell.value for cell in ws[1]]
        max_no = 1
        if "NO" in header_row:
            no_idx = header_row.index("NO") + 1
            nos = []
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=no_idx).value
                if val is not None and str(val).isdigit():
                    nos.append(int(val))
            if nos:
                max_no = max(nos) + 1

        excel_row["NO"] = max_no
        row_data = [excel_row.get(col, "") for col in header_row]
        ws.append(row_data)
        wb.save(excel_filename)
        return True, f"บันทึกข้อมูลลง Excel ({os.path.basename(excel_filename)}) ลำดับที่ {max_no} สำเร็จ!"
    except Exception as e:
        return False, f"เกิดข้อผิดพลาดในการบันทึก Excel: {e}"

def trigger_power_automate_webhook(webhook_url, raw_data, mapped_data, custom_username=None, custom_password=None):
    if not webhook_url:
        return False, "ไม่ได้ระบุ Webhook URL"
    try:
        app_name = raw_data.get("Application", "E-Travelling")
        email_to = raw_data.get("Email", "")
        user_id = custom_username if custom_username else raw_data.get("App user ID", "")
        
        if custom_password:
            password_str = custom_password
        elif app_name == "VSM":
            password_str = generate_vsm_password(email_to)
        elif app_name in ["Forma", "Pandora"]:
            password_str = generate_forma_password(user_id)
        elif app_name == "Red plate":
            password_str = "Init123456"
        else:
            password_str = user_id
            
        link_str = ""
        if app_name == "VSM":
            link_str = "https://vsm.mgc-asia.com/Pages/Home.aspx"
        elif app_name == "E-Travelling":
            link_str = "http://travelling.mgc-asia.com/"
        elif app_name == "Red plate":
            link_str = "https://redplate-frontend.azurewebsites.net/signin/?redirect_url=%2Freport%2F%3Ftype%3Dred-plate-transaction"

        payload = {
            "Application": app_name,
            "Email": email_to,
            "AppUserID": user_id,
            "Password": password_str,
            "Link": link_str,
            "Subject": f"ข้อมูลการเข้าระบบ {app_name}",
            "FullNameThai": raw_data.get("Full Name Thai", ""),
            "FullNameEng": raw_data.get("Full Name Eng", ""),
            "Position": raw_data.get("Position", ""),
            "Company": raw_data.get("Company", ""),
            "BU": mapped_data.get("BU", ""),
            "Branch": mapped_data.get("Branch", ""),
            "Status": "Active",
            "CreateDate": datetime.date.today().strftime('%Y-%m-%d')
        }
        res = requests.post(webhook_url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
        if res.status_code in [200, 202]:
            return True, "ส่งสัญญาณ Power Automate Webhook สำเร็จ!"
        return False, f"Power Automate Webhook ตอบกลับ Status: {res.status_code}"
    except Exception as e:
        return False, f"ข้อผิดพลาด Webhook: {e}"

WEB_CONFIG_FILE = "web_config.json"

def load_web_config():
    default_excel = r"C:\Users\Nawarutte.Non\OneDrive - Millennium Group Corporation (Asia) Public Company Limited\Automan\Template Column Excel Summary User.xlsx"
    if not os.path.exists(DEFAULT_EXCEL_PATH if 'DEFAULT_EXCEL_PATH' in locals() else default_excel):
        default_excel = "Template Column Excel Summary User.xlsx"
    
    cfg = {
        "gemini_api_key": os.environ.get("GEMINI_API_KEY", ""),
        "excel_path": default_excel,
        "webhook_url": os.environ.get("WEBHOOK_URL", "https://default6345207c7bd249f1920ea5aa88e4c1.c0.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/06/workflows/c8f4931f9e5646a08603ea1e9a63c307/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ue9CmEeB2GiJGWDyDWsCFpE7QcPcSYwXnKcXutGqRp0")
    }
    
    for fname in [WEB_CONFIG_FILE, "app_config.json"]:
        if os.path.exists(fname):
            try:
                with open(fname, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    if saved.get("gemini_api_key"):
                        cfg["gemini_api_key"] = saved["gemini_api_key"]
                    if saved.get("excel_path"):
                        cfg["excel_path"] = saved["excel_path"]
                    if saved.get("webhook_url"):
                        cfg["webhook_url"] = saved["webhook_url"]
            except Exception:
                pass
    return cfg

def save_web_config(api_k, exc_p, wh_u):
    try:
        data = {
            "gemini_api_key": api_k.strip(),
            "excel_path": exc_p.strip(),
            "webhook_url": wh_u.strip()
        }
        with open(WEB_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

CURRENT_WEB_CFG = load_web_config()

# ==========================================
# 🖥️ Sidebar & Config Setup
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/lock--v1.png", width=64)
    st.title("⚙️ ตั้งค่าระบบ")
    
    gemini_key = st.text_input(
        "🔑 Gemini Vision API Key (แม่นยำ 100% ฟรี)",
        value=CURRENT_WEB_CFG.get("gemini_api_key", ""),
        type="default",
        placeholder="วาง API Key ที่นี่ (ระบบจะจดจำคีย์ล่าสุดให้อัตโนมัติ)",
        help="รับ API Key ฟรีได้จาก https://aistudio.google.com/app/apikey"
    )
    
    excel_path = st.text_input(
        "📂 ไฟล์ Excel Summary Target",
        value=CURRENT_WEB_CFG.get("excel_path", "Template Column Excel Summary User.xlsx"),
        help="ตำแหน่งไฟล์ Excel ที่เชื่อมกับ SharePoint / OneDrive"
    )
    
    webhook_url = st.text_input(
        "⚡ Power Automate Webhook URL",
        value=CURRENT_WEB_CFG.get("webhook_url", ""),
        type="default",
        help="HTTP Webhook Trigger URL จาก Power Automate"
    )

    if st.button("💾 บันทึกการตั้งค่า (Save Settings)", use_container_width=True):
        save_web_config(gemini_key, excel_path, webhook_url)
        st.success("💾 บันทึกค่าการใช้งานล่าสุดเรียบร้อยแล้ว!")

    st.markdown("---")
    st.markdown("💡 **คู่มือใช้งาน**: วาง API Key ของคุณในช่องด้านบน ระบบจะจดจำคีย์ล่าสุดที่ใช้ไว้ตลอดเวลา จากนั้นแคปภาพหน้าจอและกดวางภาพได้เลยครับ")

# Auto-save key whenever used
if gemini_key and gemini_key != CURRENT_WEB_CFG.get("gemini_api_key", ""):
    save_web_config(gemini_key, excel_path, webhook_url)

# ==========================================
# 🚀 Main Page Header
# ==========================================
st.title("🔐 User Access Automation Web App")
st.caption("ระบบอ่านข้อมูลจากภาพแคปเจอร์ด้วย AI (Gemini 2.0 Flash Vision AI) ➔ บันทึก Excel ➔ ส่ง Email ตอบกลับอัตโนมัติ")

# --- Step 1: Upload Image & Application Selection ---
st.subheader("📸 1. วางรูปภาพ (Clipboard Paste) หรืออัปโหลดไฟล์ภาพแคปเจอร์หน้าจอ")

col_app, col_paste, col_up = st.columns([1, 1, 1.5])

with col_app:
    target_app = st.selectbox(
        "📱 เลือกระบบ (Application)",
        ["🔍 Auto-Detect (อัตโนมัติ)", "VSM", "E-Travelling", "Forma", "Red plate", "Pandora"]
    )

with col_paste:
    st.markdown("**📋 วางรูปภาพจาก Clipboard**")
    try:
        from streamlit_paste_button import paste_image_button
        paste_result = paste_image_button(
            label="📋 คลิกวางภาพจาก Clipboard (Paste)",
            background_color="#0d6efd",
            hover_background_color="#0b5ed7",
            text_color="#ffffff",
            errors="ignore"
        )
    except Exception:
        paste_result = None

with col_up:
    uploaded_file = st.file_uploader(
        "📁 หรือเลือก/ลากวางไฟล์ภาพ (PNG, JPG, WEBP)",
        type=["png", "jpg", "jpeg", "webp"]
    )

image = None
if paste_result is not None and paste_result.image_data is not None:
    image = paste_result.image_data
elif uploaded_file is not None:
    image = Image.open(uploaded_file)

if image is not None:
    st.image(image, caption="ภาพแคปเจอร์ที่เลือก/วาง", use_container_width=True)

    if st.button("🤖 2. ประมวลผลและดึงข้อมูลจากภาพ (Extract Data)", type="primary"):
        with st.spinner("⏳ กำลังวิเคราะห์ภาพด้วย Gemini 2.0 Flash Vision AI..."):
            extracted, err_msg = extract_with_gemini_vision(image, gemini_key)
            if extracted:
                st.session_state["extracted"] = extracted
                st.success("✨ อ่านข้อมูลจากภาพด้วย Vision AI สำเร็จเป๊ะ 100%!")
            else:
                st.error(f"⚠️ {err_msg}")

# --- Step 2: Review & Edit Data Form ---
if "extracted" in st.session_state:
    raw_data = st.session_state["extracted"]
    st.markdown("---")
    st.subheader("📝 3. ตรวจสอบและแก้ไขข้อมูลก่อนบันทึก (Review & Edit Data)")

    app_val = raw_data.get("Application", "E-Travelling")
    uid_val = raw_data.get("App user ID", "")
    neng_val = raw_data.get("Full Name Eng", "")
    nth_val = raw_data.get("Full Name Thai", "")
    email_val = raw_data.get("Email", "")
    pos_val = raw_data.get("Position", "")
    comp_val = raw_data.get("Company", "")
    branch_val = raw_data.get("Branch", "")

    col1, col2 = st.columns(2)
    with col1:
        edit_app = st.selectbox("Application", ["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"], index=["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"].index(app_val) if app_val in ["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"] else 0)
        edit_neng = st.text_input("Full Name Eng (ชื่ออังกฤษ)", value=neng_val)
        edit_email = st.text_input("Email", value=email_val)
        edit_company = st.selectbox("Company (เลือกบริษัท)", COMPANY_OPTIONS, index=COMPANY_OPTIONS.index(comp_val) if comp_val in COMPANY_OPTIONS else 0)

    with col2:
        edit_uid = st.text_input("App user ID (รหัสผู้ใช้)", value=uid_val)
        edit_nth = st.text_input("Full Name Thai (ชื่อไทย)", value=nth_val)
        edit_pos = st.text_input("Position (ตำแหน่ง/กลุ่มผู้ใช้)", value=pos_val)
        edit_branch = st.selectbox("Branch (เลือกสาขา)", BRANCH_OPTIONS, index=BRANCH_OPTIONS.index(branch_val) if branch_val in BRANCH_OPTIONS else 0)

    st.markdown("#### 🔑 ข้อมูล Username & Password สำหรับส่ง Email")
    col_un, col_pw = st.columns(2)
    
    default_uname = edit_uid
    default_pwd = get_default_password_for_app(edit_app, edit_uid, edit_email, extracted_pwd=raw_data.get("Password", ""))

    with col_un:
        edit_email_username = st.text_input("🔑 Username (ส่ง Email)", value=default_uname)
    with col_pw:
        edit_email_password = st.text_input("🔐 Password (ส่ง Email)", value=default_pwd)

    # --- Step 3: Real-Time Live Email Preview ---
    st.markdown("---")
    st.subheader("📧 4. ตัวอย่าง Email ที่จะส่งหา User (Real-Time Live Preview)")

    email_subject = f"ข้อมูลการเข้าระบบ {edit_app}"
    if edit_app == "Forma":
        intro_text = "ให้เข้าใช้งาน Forma โดย User & Password ตามด้านล่างนี้ครับ"
        link_str = ""
    elif edit_app == "Pandora":
        intro_text = "ให้เข้าใช้งาน Pandora โดย User & Password ตามด้านล่างนี้ครับ"
        link_str = ""
    elif edit_app == "Red plate":
        intro_text = "ให้เข้าใช้งาน Red plate โดย User & Password ตามด้านล่างนี้ครับ"
        link_str = "https://redplate-frontend.azurewebsites.net/signin/?redirect_url=%2Freport%2F%3Ftype%3Dred-plate-transaction"
    elif edit_app == "VSM":
        intro_text = "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
        link_str = "https://vsm.mgc-asia.com/Pages/Home.aspx"
    else: # E-Travelling
        intro_text = "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
        link_str = "http://travelling.mgc-asia.com/"

    email_preview_text = f"""========================================
📧 TEMPLATE สำหรับตอบ EMAIL ({edit_app})
========================================
To:       {edit_email if edit_email else '(ยังไม่ได้ระบุ Email)'}
Subject:  {email_subject}
----------------------------------------
{intro_text}
User:      {edit_email_username}
Password:  {edit_email_password}
"""
    if link_str:
        email_preview_text += f"Link:      {link_str}\n"
    email_preview_text += "========================================"

    st.code(email_preview_text, language="text")

    # --- Step 4: Confirm Action Button ---
    if st.button("✅ 5. ยืนยันบันทึกข้อมูลลง Excel & ส่ง Email", type="primary", use_container_width=True):
        reviewed_data = {
            "Application": edit_app,
            "App user ID": edit_email_username,
            "Full Name Eng": edit_neng,
            "Full Name Thai": edit_nth,
            "Email": edit_email,
            "Position": edit_pos,
            "Company": edit_company,
            "Branch": edit_branch
        }
        
        mapped_data = process_mapping(reviewed_data)
        ok_excel, msg_excel = export_to_excel(mapped_data, excel_path)
        
        if ok_excel:
            st.success(msg_excel)
            target_save_file = excel_path if os.path.exists(excel_path) else "Template Column Excel Summary User.xlsx"
            if os.path.exists(target_save_file):
                with open(target_save_file, "rb") as f:
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel ที่อัปเดตแล้ว (Download Excel)",
                        data=f,
                        file_name="Template Column Excel Summary User.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        else:
            st.error(msg_excel)

        if webhook_url:
            ok_wh, msg_wh = trigger_power_automate_webhook(
                webhook_url, reviewed_data, mapped_data,
                custom_username=edit_email_username,
                custom_password=edit_email_password
            )
            if ok_wh:
                st.info(msg_wh)
            else:
                st.warning(msg_wh)
        
        st.balloons()
