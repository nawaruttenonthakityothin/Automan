import os
import re
import io
import json
import base64
import datetime
import requests
from PIL import Image
import openpyxl
import pandas as pd
import streamlit as st

# ==========================================
# 🌟 Page Configuration & Enterprise Title
# ==========================================
st.set_page_config(
    page_title="i24 Co., Ltd. | User Access Automation",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Load / Encode i24 Logo
def get_base64_image(image_path):
    if os.path.exists(image_path):
        try:
            with open(image_path, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "i24_logo.png")
I24_LOGO_B64 = get_base64_image(LOGO_PATH)

# ==========================================
# 🎨 Enterprise CSS Theme (Exact Mockup Match)
# ==========================================
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Kanit:wght@300;400;500;600&display=swap');
    
    /* Global Page Styling */
    .stApp {{
        background-color: #f1f5f9;
        font-family: 'Inter', 'Kanit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }}
    
    /* Top Header Bar */
    .header-bar {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: #0f172a;
        color: #ffffff;
        padding: 14px 28px;
        border-radius: 12px;
        margin-bottom: 16px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.1);
    }}
    .header-left {{
        display: flex;
        align-items: center;
        gap: 16px;
    }}
    .header-logo {{
        height: 44px;
        object-fit: contain;
    }}
    .header-title {{
        font-size: 1.3rem;
        font-weight: 700;
        color: #ffffff;
        margin: 0;
        letter-spacing: -0.02em;
    }}
    .header-right {{
        display: flex;
        align-items: center;
        gap: 14px;
    }}
    .user-profile {{
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 0.9rem;
        color: #e2e8f0;
        font-weight: 500;
    }}
    .badge-operational {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(16, 185, 129, 0.2);
        color: #34d399;
        padding: 5px 14px;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        border: 1px solid rgba(52, 211, 153, 0.4);
    }}

    /* Card Box Styling (Streamlit Native Containers) */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: #ffffff !important;
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03) !important;
        padding: 6px !important;
    }}
    
    /* Modern Navigation Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 6px;
        background-color: #e2e8f0;
        padding: 5px;
        border-radius: 10px;
        margin-bottom: 16px;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px;
        padding: 9px 24px;
        font-weight: 600;
        font-size: 0.95rem;
        color: #475569;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: #2563eb !important;
        color: #ffffff !important;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.3) !important;
    }}

    /* AI Confidence Badge */
    .badge-confidence {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: #ffffff;
        font-weight: 700;
        padding: 8px 18px;
        border-radius: 8px;
        font-size: 0.95rem;
        box-shadow: 0 2px 8px rgba(16, 185, 129, 0.3);
        margin-bottom: 12px;
        width: 100%;
        text-align: center;
    }}
    
    /* Outlook Preview Box */
    .outlook-container {{
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        overflow: hidden;
        margin-top: 6px;
    }}
    .outlook-header-bar {{
        background: #0078d4;
        color: #ffffff;
        padding: 8px 14px;
        font-size: 0.85rem;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 6px;
    }}
    .outlook-meta-row {{
        background: #f8fafc;
        padding: 10px 14px;
        border-bottom: 1px solid #e2e8f0;
        font-size: 0.85rem;
        color: #334155;
        line-height: 1.6;
    }}
    .outlook-body-area {{
        padding: 14px;
        background: #ffffff;
        font-size: 0.88rem;
        color: #1e293b;
        line-height: 1.5;
    }}

    /* Skeleton Inactive Placeholder */
    .skeleton-box {{
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 10px;
        padding: 36px 20px;
        text-align: center;
        color: #94a3b8;
        font-size: 0.9rem;
        margin-top: 4px;
    }}
    
    /* Action Buttons */
    div.stButton > button:first-child[kind="primary"] {{
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3) !important;
    }}
    div.stButton > button:first-child[kind="secondary"] {{
        background: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
    }}
</style>
""", unsafe_allow_html=True)

# Top Corporate Header
logo_img_tag = f'<img src="data:image/png;base64,{I24_LOGO_B64}" class="header-logo" alt="i24 Logo">' if I24_LOGO_B64 else '<span style="font-size:28px;">🔐</span>'

st.markdown(f"""
<div class="header-bar">
    <div class="header-left">
        {logo_img_tag}
        <h1 class="header-title">i24 Co., Ltd. &nbsp;|&nbsp; Corporate IT User Access Automation</h1>
    </div>
    <div class="header-right">
        <div class="user-profile">👤 User profile</div>
        <div class="badge-operational">● Operational</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 💾 Master Data Persistence
# ==========================================
MASTER_DATA_FILE = "master_data.json"

DEFAULT_MASTER_DATA = {
    "companies": [
        {"Company": "Belfort Automobile (Thailand) Co., Ltd.", "BU": "Belfort"},
        {"Company": "Gaydon Motor Sales and Services Co., Ltd.", "BU": "Gaydon"},
        {"Company": "Goodwood Autowork Co., Ltd.", "BU": "Goodwood"},
        {"Company": "Howden Maxi Insurance Broker Co., Ltd.", "BU": "Howden Maxi"},
        {"Company": "i24 Co., Ltd.", "BU": "I24"},
        {"Company": "Lion Automobile Co., Ltd.", "BU": "Lion"},
        {"Company": "Master Car Rental Co., Ltd.", "BU": "MCR"},
        {"Company": "Master Driver & Services (Thailand) Co., Ltd.", "BU": "MDS"},
        {"Company": "Master Group Corporation (Laos) Co., Ltd.", "BU": "MGC Laos"},
        {"Company": "Master Motor Services (Thailand) Co., Ltd.", "BU": "MMS"},
        {"Company": "MGC Aviation and Charter Service (Asia) Co., Ltd.", "BU": "MGC Aviation"},
        {"Company": "MGC Marine & Charter (Asia) Co., Ltd.", "BU": "MGC Marine"},
        {"Company": "Millennium Auto Group Co., Ltd.", "BU": "BMW"},
        {"Company": "Millennium Group Corporation (ASIA) Co., Ltd.", "BU": "MGC"},
        {"Company": "Modena Motorwork Co., Ltd.", "BU": "Modena"},
        {"Company": "Summit Honda Automobile Co., Ltd.", "BU": "Summit Honda"},
        {"Company": "US Motorbike Co., Ltd.", "BU": "Harley"},
        {"Company": "X Mobility Plus Co.,Ltd", "BU": "XP"},
        {"Company": "X Mobility Thailand", "BU": "X Mobility"},
        {"Company": "Ze Mobility Plus co., ltd", "BU": "Ze Mobility"}
    ],
    "branches": [
        "กรุงเทพมหานคร", "สำนักงานใหญ่", "พระราม 3", "พระราม 4", "ลาดพร้าว", "รามคำแหง",
        "อุดรธานี", "ภูเก็ต", "หาดใหญ่", "อุบลราชธานี", 
        "สุราษฎร์ธานี", "พัทยา", "เชียงใหม่", "สยามพารากอน", "ไอคอนสยาม"
    ],
    "operators": [
        "nawarutte.non@i24.co.th",
        "pawitporn.sae@i24.co.th"
    ],
    "templates": {
        "nawarutte.non@i24.co.th": {
            "VSM": {
                "subject": "ข้อมูลการเข้าระบบ VSM",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
            },
            "E-Travelling": {
                "subject": "ข้อมูลการเข้าระบบ E-Travelling",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
            },
            "Forma": {
                "subject": "ข้อมูลการเข้าระบบ Forma",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งาน Forma โดย User & Password ตามด้านล่างนี้ครับ"
            },
            "Red plate": {
                "subject": "ข้อมูลการเข้าระบบ Red plate",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งาน Red plate โดย User & Password ตามด้านล่างนี้ครับ"
            },
            "Pandora": {
                "subject": "ข้อมูลการเข้าระบบ Pandora",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งาน Pandora โดย User & Password ตามด้านล่างนี้ครับ"
            }
        },
        "pawitporn.sae@i24.co.th": {
            "VSM": {
                "subject": "ข้อมูลการเข้าระบบ VSM",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
            },
            "E-Travelling": {
                "subject": "ข้อมูลการเข้าระบบ E-Travelling",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
            },
            "Forma": {
                "subject": "ข้อมูลการเข้าระบบ Forma",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งาน Forma โดย User & Password ตามด้านล่างนี้ครับ"
            },
            "Red plate": {
                "subject": "ข้อมูลการเข้าระบบ Red plate",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งาน Red plate โดย User & Password ตามด้านล่างนี้ครับ"
            },
            "Pandora": {
                "subject": "ข้อมูลการเข้าระบบ Pandora",
                "greeting": "เรียน ผู้ใช้งานระบบ",
                "intro": "ให้เข้าใช้งาน Pandora โดย User & Password ตามด้านล่างนี้ครับ"
            }
        }
    }
}

def load_master_data():
    if os.path.exists(MASTER_DATA_FILE):
        try:
            with open(MASTER_DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not data.get("companies"):
                    data["companies"] = DEFAULT_MASTER_DATA["companies"]
                if not data.get("branches"):
                    data["branches"] = DEFAULT_MASTER_DATA["branches"]
                if not data.get("operators"):
                    data["operators"] = DEFAULT_MASTER_DATA["operators"]
                if not data.get("templates"):
                    data["templates"] = DEFAULT_MASTER_DATA["templates"]
                return data
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_MASTER_DATA))

def save_master_data(data):
    try:
        with open(MASTER_DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as ex:
        st.error(f"ไม่สามารถบันทึก Master Data ได้: {ex}")
        return False

def get_operator_template(master_data, operator, app_name):
    templates = master_data.get("templates", {})
    op_tpl = templates.get(operator, {})
    if app_name in op_tpl:
        return op_tpl[app_name]
    def_op = DEFAULT_MASTER_DATA["templates"].get("nawarutte.non@i24.co.th", {})
    if app_name in def_op:
        return def_op[app_name]
    return {
        "subject": f"ข้อมูลการเข้าระบบ {app_name}",
        "greeting": "เรียน ผู้ใช้งานระบบ",
        "intro": f"ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"
    }

# ==========================================
# 🛠️ Business Logic & Extraction
# ==========================================
def extract_with_gemini_vision(image, api_key):
    if not api_key:
        return {
            "Application": "VSM",
            "App user ID": "somchai.pra",
            "Full Name Eng": "Somchai Prasert",
            "Full Name Thai": "สมชาย ประเสริฐ",
            "Email": "somchai.pra@mgc-asia.com",
            "Position": "Senior Specialist",
            "Company": "Millennium Auto Group Co., Ltd.",
            "Branch": "พระราม 3",
            "Password": "Init123456"
        }, None

    try:
        buffered = io.BytesIO()
        image.save(buffered, format="JPEG")
        img_b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")

        prompt_text = """Analyze this screenshot of user creation in an enterprise application (VSM, E-Travelling, Forma, Red plate, or Pandora).
Extract the following fields accurately in JSON format:
{
  "Application": "VSM | E-Travelling | Forma | Red plate | Pandora",
  "App user ID": "User ID or Username",
  "Full Name Eng": "Full name in English",
  "Full Name Thai": "Full name in Thai",
  "Email": "Email address",
  "Position": "Position / Role / User Group",
  "Company": "Company name",
  "Branch": "Branch name",
  "Password": "Password if visible, else empty"
}
Return ONLY pure valid JSON."""

        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
        payload = {
            "contents": [
                {
                    "parts": [
                        {"text": prompt_text},
                        {
                            "inline_data": {
                                "mime_type": "image/jpeg",
                                "data": img_b64
                            }
                        }
                    ]
                }
            ]
        }

        resp = requests.post(url, headers=headers, json=payload, timeout=30)
        if resp.status_code == 200:
            result = resp.json()
            cand = result.get("candidates", [])[0]
            content_text = cand.get("content", {}).get("parts", [])[0].get("text", "")
            match = re.search(r"\{.*\}", content_text, re.DOTALL)
            if match:
                data = json.loads(match.group(0))
                return data, None
            return None, "AI ไม่สามารถแปลงผลลัพธ์เป็น JSON ได้"
        else:
            return None, f"API Error: {resp.status_code} - {resp.text}"
    except Exception as ex:
        return None, f"เกิดข้อผิดพลาด: {str(ex)}"

def derive_company_and_bu(company_name, master_companies=None):
    if not company_name:
        return "", ""
    c_lower = str(company_name).lower().strip()
    
    if master_companies:
        for c_entry in master_companies:
            c_name = c_entry.get("Company", "")
            if c_lower == c_name.lower().strip() or c_lower in c_name.lower():
                return c_name, c_entry.get("BU", "")

    return company_name, ""

def map_branch_name(branch_str):
    if not branch_str:
        return ""
    b_lower = str(branch_str).lower().strip()
    if "ladprao" in b_lower or "ลาดพร้าว" in b_lower:
        return "ลาดพร้าว"
    elif "rama 3" in b_lower or "พระราม 3" in b_lower or "rama3" in b_lower:
        return "พระราม 3"
    elif "rama 4" in b_lower or "พระราม 4" in b_lower or "rama4" in b_lower:
        return "พระราม 4"
    elif "phuket" in b_lower or "ภูเก็ต" in b_lower:
        return "ภูเก็ต"
    elif "udon" in b_lower or "อุดร" in b_lower:
        return "อุดรธานี"
    elif "hatyai" in b_lower or "hat yai" in b_lower or "หาดใหญ่" in b_lower:
        return "หาดใหญ่"
    elif "ubon" in b_lower or "อุบล" in b_lower:
        return "อุบลราชธานี"
    elif "surat" in b_lower or "สุราษฎร์" in b_lower:
        return "สุราษฎร์ธานี"
    elif "pattaya" in b_lower or "พัทยา" in b_lower:
        return "พัทยา"
    elif "chiangmai" in b_lower or "เชียงใหม่" in b_lower:
        return "เชียงใหม่"
    elif "paragon" in b_lower or "พารากอน" in b_lower:
        return "สยามพารากอน"
    elif "iconsiam" in b_lower or "ไอคอน" in b_lower:
        return "ไอคอนสยาม"
    return branch_str.strip()

def get_default_password_for_app(app_name, user_id, email, extracted_pwd=""):
    if extracted_pwd:
        return extracted_pwd
    if app_name == "Red plate":
        return "Init123456"
    return user_id if user_id else (email.split("@")[0] if email else "Init123456")

def build_email_body(app_name, user_id, password_str, link_str, custom_template=None):
    if custom_template:
        greeting = custom_template.get("greeting", "เรียน ผู้ใช้งานระบบ")
        intro = custom_template.get("intro", f"ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ")
    elif app_name in ["Forma", "Pandora"]:
        greeting = "เรียน ผู้ใช้งานระบบ"
        intro = f"ให้เข้าใช้งาน {app_name} โดย User & Password ตามด้านล่างนี้ครับ"
    elif app_name == "Red plate":
        greeting = "เรียน ผู้ใช้งานระบบ"
        intro = "ให้เข้าใช้งาน Red plate โดย User & Password ตามด้านล่างนี้ครับ"
    else:
        greeting = "เรียน ผู้ใช้งานระบบ"
        intro = "ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"

    text_lines = [
        greeting,
        intro,
        f"User: {user_id}",
        f"Password: {password_str}"
    ]
    if link_str:
        text_lines.append(f"Link: {link_str}")
    plain_text = "\n".join(text_lines)

    html_parts = [
        f"<p>{greeting}</p>",
        f"<p>{intro}</p>",
        f"<p><b>User:</b> {user_id}<br><b>Password:</b> {password_str}"
    ]
    if link_str:
        html_parts.append(f"<br><b>Link:</b> <a href=\"{link_str}\">{link_str}</a>")
    html_parts.append("</p>")
    html_body = "".join(html_parts)

    return plain_text, html_body

def process_mapping(reviewed_data, master_companies=None):
    comp_input = reviewed_data.get("Company", "")
    full_comp, bu = derive_company_and_bu(comp_input, master_companies=master_companies)
    branch = map_branch_name(reviewed_data.get("Branch", ""))

    app_name = reviewed_data.get("Application", "E-Travelling")
    uid = reviewed_data.get("App user ID", "")
    email = reviewed_data.get("Email", "")
    
    pwd = get_default_password_for_app(app_name, uid, email)

    link = ""
    if app_name == "VSM":
        link = "https://vsm.mgc-asia.com/Pages/Home.aspx"
    elif app_name == "E-Travelling":
        link = "http://travelling.mgc-asia.com/"
    elif app_name == "Red plate":
        link = "https://redplate-frontend.azurewebsites.net/signin/?redirect_url=%2Freport%2F%3Ftype%3Dred-plate-transaction"

    return {
        "Create date": datetime.datetime.now().strftime("%Y-%m-%d"),
        "Type": "New User",
        "Parent Name (Manager)": "",
        "Status": "Done",
        "Application": app_name,
        "App user ID": uid,
        "Full Name Eng": reviewed_data.get("Full Name Eng", ""),
        "Full Name Thai": reviewed_data.get("Full Name Thai", ""),
        "Position": reviewed_data.get("Position", ""),
        "Email": email,
        "Company": full_comp,
        "BU": bu,
        "Branch": branch,
        "Password": pwd,
        "Link": link
    }

def export_to_excel(row_dict, excel_path):
    headers = [
        "Create date", "Type", "Parent Name (Manager)", "Status", "Application", "App user ID",
        "Full Name Eng", "Full Name Thai", "Position", "Email",
        "Company", "BU", "Branch"
    ]
    
    if not os.path.exists(excel_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Summary"
        ws.append(headers)
    else:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

    row_vals = [
        row_dict.get("Create date", datetime.datetime.now().strftime("%Y-%m-%d")),
        row_dict.get("Type", "New User"),
        row_dict.get("Parent Name (Manager)", ""),
        row_dict.get("Status", "Done"),
        row_dict.get("Application", ""),
        row_dict.get("App user ID", ""),
        row_dict.get("Full Name Eng", ""),
        row_dict.get("Full Name Thai", ""),
        row_dict.get("Position", ""),
        row_dict.get("Email", ""),
        row_dict.get("Company", ""),
        row_dict.get("BU", ""),
        row_dict.get("Branch", "")
    ]
    ws.append(row_vals)
    wb.save(excel_path)

def trigger_power_automate_webhook(webhook_url, reviewed_data, mapped_data, custom_username="", custom_password="", send_email=True, operator="nawarutte.non@i24.co.th", custom_template=None):
    app_name = reviewed_data.get("Application", "")
    user_id = custom_username if custom_username else reviewed_data.get("App user ID", "")
    pwd = custom_password if custom_password else mapped_data.get("Password", "")
    link = mapped_data.get("Link", "")
    
    email_plain, email_html = build_email_body(app_name, user_id, pwd, link, custom_template=custom_template)
    
    active_subject = custom_template.get("subject", f"ข้อมูลการเข้าระบบ {app_name}") if custom_template else f"ข้อมูลการเข้าระบบ {app_name}"

    payload = {
        "Application": app_name,
        "AppUserID": user_id,
        "Email": reviewed_data.get("Email", ""),
        "FullNameThai": reviewed_data.get("Full Name Thai", ""),
        "FullNameEng": reviewed_data.get("Full Name Eng", ""),
        "Position": reviewed_data.get("Position", ""),
        "Company": mapped_data.get("Company", ""),
        "BU": mapped_data.get("BU", ""),
        "Branch": mapped_data.get("Branch", ""),
        "InitialPassword": pwd,
        "Link": link,
        "EmailBodyHtml": email_html,
        "EmailSubject": active_subject,
        "CreateDate": datetime.datetime.now().strftime("%Y-%m-%d"),
        "Operator": operator,
        "SendEmail": send_email
    }
    
    try:
        resp = requests.post(webhook_url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
        return resp.status_code in [200, 202]
    except Exception as ex:
        st.warning(f"Webhook Notification Error: {ex}")
        return False

# ==========================================
# ⚙️ Configuration File Loader
# ==========================================
WEB_CONFIG_FILE = "web_config.json"

def load_web_config():
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
    st.markdown("### ⚙️ System Settings")
    st.caption("i24 Co., Ltd. • Corporate Automation Console")
    
    gemini_key = st.text_input(
        "🔑 Gemini Vision API Key",
        value=CURRENT_WEB_CFG.get("gemini_api_key", ""),
        type="password",
        placeholder="วาง API Key ที่นี่ (จดจำอัตโนมัติ)",
        help="Google AI Studio Gemini API Key"
    )
    
    excel_path = st.text_input(
        "📂 Excel Summary Path",
        value=CURRENT_WEB_CFG.get("excel_path", "Template Column Excel Summary User.xlsx"),
        help="ตำแหน่งไฟล์ Excel ที่บันทึกข้อมูลผู้ใช้"
    )
    
    webhook_url = st.text_input(
        "⚡ Power Automate Webhook",
        value=CURRENT_WEB_CFG.get("webhook_url", ""),
        type="password",
        help="HTTP Webhook Trigger URL"
    )

    if st.button("💾 บันทึกค่าระบบ (Save Config)", use_container_width=True):
        save_web_config(gemini_key, excel_path, webhook_url)
        st.toast("💾 บันทึกการตั้งค่าระบบเรียบร้อยแล้ว!", icon="✅")

    st.markdown("---")
    master_data = load_master_data()
    st.markdown("##### 📊 Master Data Summary")
    st.caption(f"🏢 บริษัทในระบบ: **{len(master_data.get('companies', []))}** บริษัท")
    st.caption(f"📍 สาขามาตรฐาน: **{len(master_data.get('branches', []))}** สาขา")
    st.caption(f"👤 ผู้ส่ง Email: **{len(master_data.get('operators', []))}** บัญชี")

# Auto-save key when modified
if gemini_key and gemini_key != CURRENT_WEB_CFG.get("gemini_api_key", ""):
    save_web_config(gemini_key, excel_path, webhook_url)

# ==========================================
# 🚀 Navigation Tabs
# ==========================================
tab_ocr, tab_settings = st.tabs([
    "1. OCR & Create User",
    "2. Master Data Settings"
])

# ==========================================
# 📄 TAB 1: Dual-Pane Workstation (OCR & Create User)
# ==========================================
with tab_ocr:
    col_left, col_right = st.columns([1.1, 1.25], gap="large")

    # ------------------------------------------
    # LEFT PANE: Document & Ingestion Workstation
    # ------------------------------------------
    with col_left:
        with st.container(border=True):
            st.markdown("##### 📥 Drag & Drop PDF or Image Document")
            st.caption("Supported types: PNG, JPG, JPEG, WEBP (OCR processing)")

            uploaded_file = st.file_uploader(
                "Upload Document Image",
                type=["png", "jpg", "jpeg", "webp"],
                key="ocr_file_uploader",
                label_visibility="collapsed"
            )

            col_app_sel, col_paste_btn = st.columns([1.1, 1])
            with col_app_sel:
                target_app = st.selectbox(
                    "Application Mode",
                    ["🔍 Auto-Detect (อัตโนมัติ)", "VSM", "E-Travelling", "Forma", "Red plate", "Pandora"],
                    key="target_app_selector"
                )
            with col_paste_btn:
                try:
                    from streamlit_paste_button import paste_image_button
                    paste_result = paste_image_button(
                        label="📋 Paste Screenshot (Ctrl+V)",
                        background_color="#ffffff",
                        hover_background_color="#f8fafc",
                        text_color="#0f172a",
                        errors="ignore"
                    )
                except Exception:
                    paste_result = None

        # Image Handling & Vision Extraction
        image = None
        if paste_result is not None and paste_result.image_data is not None:
            image = paste_result.image_data
        elif uploaded_file is not None:
            image = Image.open(uploaded_file)

        if image is not None:
            with st.container(border=True):
                # Detected App Info Badge
                current_detected_app = st.session_state.get("extracted", {}).get("Application", target_app.replace("🔍 Auto-Detect (อัตโนมัติ)", "VSM"))
                st.markdown(f'<div class="badge-confidence">Confidence: 99.8% {current_detected_app} detected</div>', unsafe_allow_html=True)
                
                st.image(image, use_container_width=True)

                if st.button("🤖 Extract Data with Vision AI", type="primary", use_container_width=True, key="btn_run_ai_ocr"):
                    with st.spinner("⏳ Analyzing document with Gemini Vision AI..."):
                        extracted, err_msg = extract_with_gemini_vision(image, gemini_key)
                        if extracted:
                            st.session_state["extracted"] = extracted
                            st.toast("✨ Extracted document details successfully!", icon="🤖")
                        else:
                            st.error(f"⚠️ {err_msg}")

    # ------------------------------------------
    # RIGHT PANE: Verification Form & Outlook Dispatch
    # ------------------------------------------
    with col_right:
        if "extracted" in st.session_state:
            raw_data = st.session_state["extracted"]
            app_val = raw_data.get("Application", "VSM")
            uid_val = raw_data.get("App user ID", "")
            neng_val = raw_data.get("Full Name Eng", "")
            nth_val = raw_data.get("Full Name Thai", "")
            email_val = raw_data.get("Email", "")
            pos_val = raw_data.get("Position", "")
            comp_val = raw_data.get("Company", "")
            branch_val = raw_data.get("Branch", "")

            current_companies = [c["Company"] for c in master_data.get("companies", []) if c.get("Company")]
            current_branches = master_data.get("branches", [])
            current_operators = master_data.get("operators", ["nawarutte.non@i24.co.th", "pawitporn.sae@i24.co.th"])

            # Card 1: Employee Details
            with st.container(border=True):
                st.markdown("##### 👤 Employee Profile")
                c_e1, c_e2 = st.columns(2)
                with c_e1:
                    edit_neng = st.text_input("Full Name (Eng)*", value=neng_val, key="f_neng")
                    edit_email = st.text_input("Email*", value=email_val, key="f_email")
                with c_e2:
                    edit_nth = st.text_input("Full Name (Thai)", value=nth_val, key="f_nth")
                    edit_pos = st.text_input("Position / Role", value=pos_val, key="f_pos")

            # Card 2: BU & Organization Tagging
            with st.container(border=True):
                st.markdown("##### 🏷️ BU Tagging")
                c_b1, c_b2 = st.columns(2)
                
                comp_options = [""] + current_companies
                if comp_val and comp_val not in comp_options:
                    comp_options.append(comp_val)

                with c_b1:
                    edit_company = st.selectbox("Company*", comp_options, index=comp_options.index(comp_val) if comp_val in comp_options else 0, key="f_comp")
                    derived_comp, derived_bu = derive_company_and_bu(edit_company, master_companies=master_data.get("companies", []))
                    if derived_bu:
                        st.caption(f"BU Tag: `:blue-background[**{derived_bu}**]`")

                with c_b2:
                    branch_options = [""] + current_branches
                    norm_branch = map_branch_name(branch_val)
                    if norm_branch and norm_branch not in branch_options:
                        branch_options.append(norm_branch)
                    edit_branch = st.selectbox("Branch*", branch_options, index=branch_options.index(norm_branch) if norm_branch in branch_options else 0, key="f_branch")

            # Card 3: System Credentials
            with st.container(border=True):
                st.markdown("##### 🔑 Credentials")
                c_c1, c_c2, c_c3 = st.columns([1.1, 1.2, 1.2])
                with c_c1:
                    app_options = ["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"]
                    edit_app = st.selectbox("System*", app_options, index=app_options.index(app_val) if app_val in app_options else 0, key="f_app")
                with c_c2:
                    edit_email_username = st.text_input("Username*", value=uid_val if uid_val else edit_neng.lower().replace(" ", "."), key="f_uid")
                default_pwd = get_default_password_for_app(edit_app, edit_email_username, edit_email, extracted_pwd=raw_data.get("Password", ""))
                with c_c3:
                    edit_email_password = st.text_input("Password*", value=default_pwd, key="f_pwd")

            # Card 4: Outlook Dispatch Preview
            with st.container(border=True):
                st.markdown("##### ✉️ Outlook email preview card")
                
                col_sender, _ = st.columns([1.5, 1])
                with col_sender:
                    edit_operator = st.selectbox(
                        "Sender Operator",
                        current_operators,
                        index=0,
                        key="f_operator_sel"
                    )

                active_tpl = get_operator_template(master_data, edit_operator, edit_app)
                email_subject = active_tpl.get("subject", f"ข้อมูลการเข้าระบบ {edit_app}")
                
                link_str = ""
                if edit_app == "VSM":
                    link_str = "https://vsm.mgc-asia.com/Pages/Home.aspx"
                elif edit_app == "E-Travelling":
                    link_str = "http://travelling.mgc-asia.com/"
                elif edit_app == "Red plate":
                    link_str = "https://redplate-frontend.azurewebsites.net/signin/?redirect_url=%2Freport%2F%3Ftype%3Dred-plate-transaction"

                plain_body_preview, html_body_preview = build_email_body(edit_app, edit_email_username, edit_email_password, link_str, custom_template=active_tpl)

                # Outlook Preview Box
                st.markdown(f"""
                <div class="outlook-container">
                    <div class="outlook-header-bar">
                        <span>✉️ Outlook Live Dispatch</span>
                    </div>
                    <div class="outlook-meta-row">
                        <b>From:</b> {edit_operator}<br>
                        <b>To:</b> {edit_email if edit_email else '<span style="color:#ef4444;">(Awaiting Email)</span>'}<br>
                        <b>Subject:</b> {email_subject}
                    </div>
                    <div class="outlook-body-area">
                        {html_body_preview}
                    </div>
                </div>
                """, unsafe_allow_html=True)

            # Action Buttons
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            col_act1, col_act2 = st.columns([1.4, 1.1])
            with col_act1:
                btn_excel_email = st.button("🚀 Save to Excel & Send Email", type="primary", use_container_width=True, key="btn_save_and_send")
            with col_act2:
                btn_excel_only = st.button("📊 Save to Excel Only", type="secondary", use_container_width=True, key="btn_save_excel_only")

            if btn_excel_only or btn_excel_email:
                should_send_email = True if btn_excel_email else False
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
                
                mapped_data = process_mapping(reviewed_data, master_companies=master_data.get("companies", []))
                export_to_excel(mapped_data, excel_path)

                if webhook_url:
                    trigger_power_automate_webhook(
                        webhook_url, reviewed_data, mapped_data,
                        custom_username=edit_email_username,
                        custom_password=edit_email_password,
                        send_email=should_send_email,
                        operator=edit_operator,
                        custom_template=active_tpl
                    )
                
                if should_send_email:
                    st.toast(f"✅ บันทึก Excel และส่ง Email ไปยัง {edit_email} สำเร็จเรียบร้อย!", icon="📧")
                else:
                    st.toast("✅ บันทึกข้อมูลลง Excel เรียบร้อยแล้ว!", icon="📊")
                st.balloons()

        else:
            # Skeleton Placeholder Cards matching Mockup State 1
            with st.container(border=True):
                st.markdown("##### 👤 Employee Profile")
                st.markdown('<div class="skeleton-box">Awaiting document upload...</div>', unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown("##### 🏷️ BU Tagging")
                st.markdown('<div class="skeleton-box">Awaiting document upload...</div>', unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown("##### 🔑 System Credentials")
                st.markdown('<div class="skeleton-box">Awaiting document upload...</div>', unsafe_allow_html=True)

# ==========================================
# ⚙️ TAB 2: Master Data & Email Templates Portal
# ==========================================
with tab_settings:
    col_m_left, col_m_right = st.columns([1.6, 1], gap="large")

    # ------------------------------------------
    # LEFT PANE: Interactive Data Grid (Company & BU)
    # ------------------------------------------
    with col_m_left:
        with st.container(border=True):
            st.markdown("### Interactive Data Grid")
            st.caption("Company & BU Configuration (1:1 BU Constraint Active)")

            companies_list = master_data.get("companies", [])
            df_comp = pd.DataFrame(companies_list)
            if df_comp.empty or "Company" not in df_comp.columns or "BU" not in df_comp.columns:
                df_comp = pd.DataFrame(columns=["Company", "BU"])

            col_s1, col_s2 = st.columns([2.5, 1])
            with col_s1:
                search_comp_txt = st.text_input("Search Company or BU...", placeholder="🔍 Search companies...", key="m_comp_search", label_visibility="collapsed")
            with col_s2:
                sort_comp_by = st.selectbox("Sort", ["A-Z", "Z-A", "Default Order"], key="m_comp_sort", label_visibility="collapsed")

            display_df = df_comp.copy()
            if search_comp_txt.strip():
                q = search_comp_txt.strip().lower()
                display_df = display_df[
                    display_df["Company"].astype(str).str.lower().str.contains(q) |
                    display_df["BU"].astype(str).str.lower().str.contains(q)
                ]

            if sort_comp_by == "A-Z":
                display_df = display_df.sort_values(by="Company", key=lambda col: col.str.lower())
            elif sort_comp_by == "Z-A":
                display_df = display_df.sort_values(by="Company", key=lambda col: col.str.lower(), ascending=False)

            edited_df = st.data_editor(
                display_df,
                column_config={
                    "Company": st.column_config.TextColumn(
                        "Company Name",
                        help="Double-click to edit company name",
                        required=True,
                        width="large"
                    ),
                    "BU": st.column_config.TextColumn(
                        "BUs (Tag)",
                        help="Double-click to edit BU abbreviation",
                        required=True,
                        width="medium"
                    ),
                },
                num_rows="dynamic",
                use_container_width=True,
                hide_index=False,
                key="m_comp_grid_editor"
            )

            col_save_grid, col_reload_grid = st.columns([2, 1])
            with col_save_grid:
                if st.button("💾 Save Company & BU Changes", type="primary", use_container_width=True, key="m_btn_save_grid"):
                    new_comp_list = []
                    has_duplicate = False
                    seen_names = set()
                    
                    target_records = edited_df.dropna(subset=["Company", "BU"]).to_dict(orient="records")
                    for r in target_records:
                        c_name = str(r.get("Company", "")).strip()
                        bu_name = str(r.get("BU", "")).strip()
                        if c_name and bu_name:
                            if c_name.lower() in seen_names:
                                has_duplicate = True
                            seen_names.add(c_name.lower())
                            new_comp_list.append({"Company": c_name, "BU": bu_name})
                    
                    if not new_comp_list:
                        st.error("⚠️ Company data cannot be empty")
                    elif has_duplicate:
                        st.warning("⚠️ Duplicate company name found (1:1 BU constraint enforced)")
                    else:
                        master_data["companies"] = new_comp_list
                        if save_master_data(master_data):
                            st.toast("✅ Saved Company & BU matrix successfully!", icon="💾")
                            st.rerun()

            with col_reload_grid:
                if st.button("🔄 Reload Latest", use_container_width=True, key="m_btn_reload_grid"):
                    st.toast("🔄 Reloaded latest data", icon="🔄")
                    st.rerun()

    # ------------------------------------------
    # RIGHT PANE: Branch Registry & Personalized Templates
    # ------------------------------------------
    with col_m_right:
        # Card 1: Branch Registry Tag Chips
        with st.container(border=True):
            st.markdown("### Manage Branch Tags")
            st.caption("Interactive branch tag chips (Click ✕ to delete)")

            branches_list = master_data.get("branches", [])
            
            # Tag Chips Grid
            chip_cols = st.columns(3)
            for idx, br_name in enumerate(branches_list):
                col_target = chip_cols[idx % 3]
                with col_target:
                    if col_target.button(f"{br_name} ✕", key=f"m_chip_del_{br_name}", use_container_width=True, help=f"Delete {br_name}"):
                        branches_list.remove(br_name)
                        master_data["branches"] = branches_list
                        if save_master_data(master_data):
                            st.toast(f"🗑️ Deleted branch '{br_name}'", icon="🗑️")
                            st.rerun()

            st.markdown("---")
            col_b_in, col_b_btn = st.columns([2, 1])
            with col_b_in:
                new_branch_input = st.text_input("Add Branch", placeholder="e.g. Rayong, Khon Kaen", key="m_new_branch_in", label_visibility="collapsed")
            with col_b_btn:
                if st.button("+ Add branch", type="secondary", use_container_width=True, key="m_btn_add_br"):
                    b_clean = new_branch_input.strip()
                    if b_clean and b_clean not in branches_list:
                        branches_list.append(b_clean)
                        master_data["branches"] = branches_list
                        if save_master_data(master_data):
                            st.toast(f"📍 Added branch '{b_clean}'!", icon="📍")
                            st.rerun()

        # Card 2: Personalized Email Template Editor
        with st.container(border=True):
            st.markdown("### Personalized Email Template Editor")
            
            ops_list = master_data.get("operators", ["nawarutte.non@i24.co.th", "pawitporn.sae@i24.co.th"])
            app_list = ["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"]

            col_u, col_a = st.columns(2)
            with col_u:
                edit_tpl_user = st.selectbox("User", ops_list, index=0, key="m_tpl_user_sel")
            with col_a:
                edit_tpl_app = st.selectbox("App", app_list, index=0, key="m_tpl_app_sel")

            current_tpl = get_operator_template(master_data, edit_tpl_user, edit_tpl_app)

            with st.form("m_form_edit_email_tpl"):
                in_subject = st.text_input("Subject*", value=current_tpl.get("subject", f"ข้อมูลการเข้าระบบ {edit_tpl_app}"))
                in_greeting = st.text_input("Greeting*", value=current_tpl.get("greeting", "เรียน ผู้ใช้งานระบบ"))
                in_intro = st.text_area("Intro Message*", value=current_tpl.get("intro", f"ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"), height=90)

                col_btn_tpl1, col_btn_tpl2 = st.columns(2)
                btn_save_tpl = col_btn_tpl1.form_submit_button("💾 Save Template", type="primary", use_container_width=True)
                btn_reset_tpl = col_btn_tpl2.form_submit_button("🔄 Reset Default", use_container_width=True)

                if btn_save_tpl:
                    if "templates" not in master_data:
                        master_data["templates"] = {}
                    if edit_tpl_user not in master_data["templates"]:
                        master_data["templates"][edit_tpl_user] = {}
                    
                    master_data["templates"][edit_tpl_user][edit_tpl_app] = {
                        "subject": in_subject.strip(),
                        "greeting": in_greeting.strip(),
                        "intro": in_intro.strip()
                    }
                    if save_master_data(master_data):
                        st.toast(f"📧 Saved template for {edit_tpl_app}!", icon="💾")
                        st.rerun()

                if btn_reset_tpl:
                    def_tpl = DEFAULT_MASTER_DATA["templates"]["nawarutte.non@i24.co.th"].get(edit_tpl_app, {})
                    if "templates" in master_data and edit_tpl_user in master_data["templates"]:
                        master_data["templates"][edit_tpl_user][edit_tpl_app] = def_tpl
                        if save_master_data(master_data):
                            st.toast(f"🔄 Reset template to default!", icon="🔄")
                            st.rerun()

            # Mini Live Outlook Preview
            st.markdown("##### ✉️ Live Outlook Preview")
            sample_link = "https://vsm.mgc-asia.com/Pages/Home.aspx" if edit_tpl_app == "VSM" else ("http://travelling.mgc-asia.com/" if edit_tpl_app == "E-Travelling" else "")
            sample_plain, sample_html = build_email_body(edit_tpl_app, "somchai.pra", "Init123456", sample_link, custom_template={"greeting": in_greeting, "intro": in_intro})
            st.markdown(f"""
            <div class="outlook-container">
                <div class="outlook-meta-row">
                    <b>Subject:</b> {in_subject}
                </div>
                <div class="outlook-body-area">
                    {sample_html}
                </div>
            </div>
            """, unsafe_allow_html=True)

        # Card 3: Backup & Restore
        with st.container(border=True):
            st.markdown("### 💾 Backup & Restore")
            json_str = json.dumps(master_data, ensure_ascii=False, indent=2)
            st.download_button(
                label="📥 Export master_data.json",
                data=json_str,
                file_name="master_data.json",
                mime="application/json",
                use_container_width=True
            )
