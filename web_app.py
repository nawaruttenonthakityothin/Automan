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
    initial_sidebar_state="expanded"
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
# 🎨 Enterprise CSS Theme
# ==========================================
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Kanit:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Inter', 'Kanit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }}
    
    /* Top Header Container */
    .enterprise-header {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #ffffff;
        padding: 16px 24px;
        border-radius: 12px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }}
    .header-left {{
        display: flex;
        align-items: center;
        gap: 16px;
    }}
    .header-logo {{
        height: 48px;
        object-fit: contain;
    }}
    .header-title-box h1 {{
        font-size: 1.35rem;
        font-weight: 700;
        margin: 0;
        color: #ffffff;
        letter-spacing: -0.01em;
    }}
    .header-title-box p {{
        font-size: 0.85rem;
        color: #94a3b8;
        margin: 0;
    }}
    .header-status {{
        display: flex;
        align-items: center;
        gap: 8px;
        background: rgba(16, 185, 129, 0.15);
        color: #34d399;
        padding: 6px 14px;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        border: 1px solid rgba(52, 211, 153, 0.3);
    }}
    
    /* Enterprise Card Boxes */
    .enterprise-card {{
        background: #ffffff;
        border-radius: 12px;
        padding: 18px 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 16px;
    }}
    .card-title {{
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    
    /* AI Confidence Badge */
    .badge-ai {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background-color: #ecfdf5;
        color: #047857;
        font-weight: 600;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.85rem;
        border: 1px solid #a7f3d0;
        margin-bottom: 10px;
    }}
    
    /* Outlook Preview Box */
    .outlook-box {{
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        overflow: hidden;
        margin-top: 10px;
    }}
    .outlook-header {{
        background: #0078d4;
        color: #ffffff;
        padding: 8px 14px;
        font-size: 0.85rem;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    .outlook-meta {{
        background: #f8fafc;
        padding: 10px 14px;
        border-bottom: 1px solid #e2e8f0;
        font-size: 0.85rem;
        color: #334155;
    }}
    .outlook-body {{
        padding: 14px;
        background: #ffffff;
        font-size: 0.9rem;
        color: #1e293b;
        line-height: 1.5;
    }}
    
    /* Primary / Secondary Button Styling */
    div.stButton > button:first-child[kind="primary"] {{
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        border: none !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.3) !important;
    }}
    div.stButton > button:first-child[kind="secondary"] {{
        border-radius: 8px !important;
        font-weight: 600 !important;
    }}
    
    /* Streamlit Tab Customization */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 10px;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        color: #475569;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: #ffffff !important;
        color: #0f172a !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }}
</style>
""", unsafe_allow_html=True)

# Top Header Bar Rendering
logo_html = f'<img src="data:image/png;base64,{I24_LOGO_B64}" class="header-logo" alt="i24 Logo">' if I24_LOGO_B64 else '<span style="font-size:28px;">🔐</span>'

st.markdown(f"""
<div class="enterprise-header">
    <div class="header-left">
        {logo_html}
        <div class="header-title-box">
            <h1>i24 Co., Ltd. &nbsp;|&nbsp; Corporate IT User Access Automation</h1>
            <p>Enterprise AI Vision Ingestion ➔ Automated Excel Logging ➔ Multi-Sender Email Dispatch</p>
        </div>
    </div>
    <div class="header-status">
        <span>●</span> Operational & Ready
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
    "📄 1. สกัดข้อมูลภาพ & บันทึกผู้ใช้ (OCR & Create User)",
    "⚙️ 2. จัดการ Master Data & Email Templates"
])

# ==========================================
# 📄 TAB 1: Dual-Pane Workstation (OCR & Create User)
# ==========================================
with tab_ocr:
    col_left, col_right = st.columns([1.05, 1.25], gap="large")

    # ------------------------------------------
    # LEFT PANE: Document & Ingestion Workstation
    # ------------------------------------------
    with col_left:
        st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📥 1. นำเข้าภาพแคปเจอร์หน้าจอ (Document Ingestion)</div>', unsafe_allow_html=True)

        col_app_sel, col_paste_btn = st.columns([1.2, 1])
        with col_app_sel:
            target_app = st.selectbox(
                "📱 เลือกระบบ (Application)",
                ["🔍 Auto-Detect (อัตโนมัติ)", "VSM", "E-Travelling", "Forma", "Red plate", "Pandora"],
                key="target_app_selector"
            )
        
        with col_paste_btn:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            try:
                from streamlit_paste_button import paste_image_button
                paste_result = paste_image_button(
                    label="📋 วางภาพจาก Clipboard",
                    background_color="#2563eb",
                    hover_background_color="#1d4ed8",
                    text_color="#ffffff",
                    errors="ignore"
                )
            except Exception:
                paste_result = None

        uploaded_file = st.file_uploader(
            "📁 หรือลากวางไฟล์ภาพแคปเจอร์ (PNG, JPG, WEBP)",
            type=["png", "jpg", "jpeg", "webp"],
            key="ocr_file_uploader"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # Image Handling & Vision Extraction
        image = None
        if paste_result is not None and paste_result.image_data is not None:
            image = paste_result.image_data
        elif uploaded_file is not None:
            image = Image.open(uploaded_file)

        if image is not None:
            st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">🔍 2. เอกสารต้นฉบับ & การประมวลผล (Document Viewer)</div>', unsafe_allow_html=True)
            
            st.image(image, caption="ภาพเอกสารหลักฐาน (Evidence Screenshot)", use_container_width=True)

            if st.button("🤖 ดึงข้อมูลด้วย Vision AI (Extract Data)", type="primary", use_container_width=True, key="btn_run_ai_ocr"):
                with st.spinner("⏳ กำลังวิเคราะห์ข้อมูลภาพด้วย Gemini Vision AI..."):
                    extracted, err_msg = extract_with_gemini_vision(image, gemini_key)
                    if extracted:
                        st.session_state["extracted"] = extracted
                        st.toast("✨ สกัดข้อมูลจากภาพด้วย Vision AI สำเร็จ 100%!", icon="🤖")
                    else:
                        st.error(f"⚠️ {err_msg}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("💡 กรุณาวางภาพจาก Clipboard หรือลากไฟล์ภาพลงในกล่องด้านบนเพื่อเริ่มต้น")

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

            # AI Status Badge
            st.markdown(f'<div class="badge-ai">⚡ Confidence: 99.8% &nbsp;•&nbsp; ระบบ: <b>{app_val}</b></div>', unsafe_allow_html=True)

            # Card 1: Employee Details
            st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">👤 1. ข้อมูลผู้ใช้งาน (Employee Details)</div>', unsafe_allow_html=True)
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                edit_neng = st.text_input("Full Name Eng (ชื่ออังกฤษ)*", value=neng_val, key="f_neng")
                edit_email = st.text_input("Email*", value=email_val, key="f_email")
            with c_e2:
                edit_nth = st.text_input("Full Name Thai (ชื่อไทย)", value=nth_val, key="f_nth")
                edit_pos = st.text_input("Position (ตำแหน่ง/กลุ่มผู้ใช้)", value=pos_val, key="f_pos")
            st.markdown('</div>', unsafe_allow_html=True)

            # Card 2: BU & Organization Tagging
            st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">🏢 2. องค์กรและสาขา (BU & Organization)</div>', unsafe_allow_html=True)
            c_b1, c_b2 = st.columns(2)
            
            comp_options = [""] + current_companies
            if comp_val and comp_val not in comp_options:
                comp_options.append(comp_val)

            with c_b1:
                edit_company = st.selectbox("Company (เลือกบริษัท)*", comp_options, index=comp_options.index(comp_val) if comp_val in comp_options else 0, key="f_comp")
                # Show dynamic BU Tag badge
                derived_comp, derived_bu = derive_company_and_bu(edit_company, master_companies=master_data.get("companies", []))
                if derived_bu:
                    st.caption(f"🏷️ ตัวย่อ BU อัตโนมัติ: `:blue-background[**{derived_bu}**]`")

            with c_b2:
                branch_options = [""] + current_branches
                norm_branch = map_branch_name(branch_val)
                if norm_branch and norm_branch not in branch_options:
                    branch_options.append(norm_branch)
                edit_branch = st.selectbox("Branch (เลือกสาขา)*", branch_options, index=branch_options.index(norm_branch) if norm_branch in branch_options else 0, key="f_branch")
            st.markdown('</div>', unsafe_allow_html=True)

            # Card 3: System Credentials
            st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">🔐 3. ข้อมูลสิทธิ์และรหัสผ่าน (Credentials & Access)</div>', unsafe_allow_html=True)
            c_c1, c_c2, c_c3 = st.columns([1.2, 1.2, 1.2])
            
            with c_c1:
                app_options = ["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"]
                edit_app = st.selectbox("Application*", app_options, index=app_options.index(app_val) if app_val in app_options else 0, key="f_app")
            
            with c_c2:
                edit_email_username = st.text_input("🔑 Username (App User ID)*", value=uid_val if uid_val else edit_neng.lower().replace(" ", "."), key="f_uid")

            default_pwd = get_default_password_for_app(edit_app, edit_email_username, edit_email, extracted_pwd=raw_data.get("Password", ""))
            with c_c3:
                edit_email_password = st.text_input("🔐 Initial Password*", value=default_pwd, key="f_pwd")
            st.markdown('</div>', unsafe_allow_html=True)

            # Card 4: Outlook Live Dispatch Preview
            st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">📧 4. ตัวอย่างอีเมล Outlook ตอบกลับ (Live Email Dispatch)</div>', unsafe_allow_html=True)
            
            col_sender, _ = st.columns([1.5, 1])
            with col_sender:
                edit_operator = st.selectbox(
                    "👤 ผู้ส่ง Email (Operator / Sender)",
                    current_operators,
                    index=0,
                    help="ระบบจะดึง Template เฉพาะของบุคคลนี้มาใช้งาน",
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

            # Outlook Card Box
            st.markdown(f"""
            <div class="outlook-box">
                <div class="outlook-header">
                    <span>✉️ Outlook Notification Preview</span>
                </div>
                <div class="outlook-meta">
                    <b>From:</b> {edit_operator}<br>
                    <b>To:</b> {edit_email if edit_email else '<span style="color:#ef4444;">(ยังไม่ได้ระบุ Email)</span>'}<br>
                    <b>Subject:</b> {email_subject}
                </div>
                <div class="outlook-body">
                    {html_body_preview}
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Card 5: Action & Dispatch Buttons
            st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">🚀 5. ยืนยันคำสั่งดำเนินการ (Execution & Dispatch)</div>', unsafe_allow_html=True)
            
            col_act1, col_act2 = st.columns(2)
            btn_excel_email = col_act1.button("📧 1. บันทึก Excel และส่ง Email", type="primary", use_container_width=True, key="btn_save_and_send")
            btn_excel_only = col_act2.button("📊 2. บันทึก Excel อย่างเดียว (ไม่ส่ง Email)", type="secondary", use_container_width=True, key="btn_save_excel_only")

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
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="enterprise-card" style="text-align:center; padding: 60px 20px; color:#64748b;">', unsafe_allow_html=True)
            st.markdown("### 📋 ฟอร์มตรวจสอบข้อมูล (Awaiting Data)")
            st.caption("เมื่ออัปโหลดและประมวลผลภาพจากฝั่งซ้าย ข้อมูลจะถูกจัดหมวดหมู่และแสดงขึ้นที่นี่อัตโนมัติ")
            st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# ⚙️ TAB 2: Master Data & Email Templates Portal
# ==========================================
with tab_settings:
    st.caption("🏢 i24 Co., Ltd. • จัดการรายชื่อบริษัท, ตัวย่อ BU, สาขามาตรฐาน และเทมเพลตอีเมลเฉพาะบุคคล")

    subtab_comp, subtab_branch, subtab_ops, subtab_tpl, subtab_backup = st.tabs([
        "🏢 1. บริษัท & BU (Company & BU Matrix)",
        "📍 2. รายชื่อสาขา (Branch Registry)",
        "👤 3. รายชื่อผู้ส่ง (Operators)",
        "📧 4. เทมเพลตอีเมล (Email Templates)",
        "💾 5. สำรองและคืนค่า (Backup & Restore)"
    ])

    # ------------------------------------------
    # Sub-tab 1: Company & BU (Interactive Data Grid + Search + Sort)
    # ------------------------------------------
    with subtab_comp:
        st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">🏢 จัดการรายชื่อบริษัท และตัวย่อ BU (1 บริษัท = 1 BU เสมอ)</div>', unsafe_allow_html=True)
        st.info("💡 **วิธีแก้ไข**: ดับเบิ้ลคลิก (Double-click) ที่ช่อง **ชื่อบริษัท** หรือ **ตัวย่อ BU** ในตารางเพื่อพิมพ์แก้ไขตัวสะกดได้ทันทีในคลิกเดียว หรือกด `+` แถวล่างสุดเพื่อเพิ่มแถวใหม่")

        companies_list = master_data.get("companies", [])
        
        df_comp = pd.DataFrame(companies_list)
        if df_comp.empty or "Company" not in df_comp.columns or "BU" not in df_comp.columns:
            df_comp = pd.DataFrame(columns=["Company", "BU"])

        # Top Search & Sort
        col_c_s1, col_c_s2 = st.columns([2.5, 1.2])
        with col_c_s1:
            search_comp_txt = st.text_input("🔍 ค้นหาในตาราง...", placeholder="พิมพ์คำค้นหา เช่น Honda, MCR, BMW...", key="m_comp_search")
        with col_c_s2:
            sort_comp_by = st.selectbox("🔃 เรียงลำดับ", ["ลำดับเริ่มต้น", "ชื่อบริษัท (A-Z)", "ชื่อบริษัท (Z-A)", "ตัวย่อ BU (A-Z)", "ตัวย่อ BU (Z-A)"], key="m_comp_sort")

        display_df = df_comp.copy()
        if search_comp_txt.strip():
            q = search_comp_txt.strip().lower()
            display_df = display_df[
                display_df["Company"].astype(str).str.lower().str.contains(q) |
                display_df["BU"].astype(str).str.lower().str.contains(q)
            ]

        if sort_comp_by == "ชื่อบริษัท (A-Z)":
            display_df = display_df.sort_values(by="Company", key=lambda col: col.str.lower())
        elif sort_comp_by == "ชื่อบริษัท (Z-A)":
            display_df = display_df.sort_values(by="Company", key=lambda col: col.str.lower(), ascending=False)
        elif sort_comp_by == "ตัวย่อ BU (A-Z)":
            display_df = display_df.sort_values(by="BU", key=lambda col: col.str.lower())
        elif sort_comp_by == "ตัวย่อ BU (Z-A)":
            display_df = display_df.sort_values(by="BU", key=lambda col: col.str.lower(), ascending=False)

        st.caption(f"📊 แสดง **{len(display_df)}** จากทั้งหมด **{len(df_comp)}** บริษัท")

        # Interactive Grid
        edited_df = st.data_editor(
            display_df,
            column_config={
                "Company": st.column_config.TextColumn(
                    "🏢 Company (ชื่อบริษัท) [ดับเบิ้ลคลิกเพื่อแก้ไข]",
                    help="ดับเบิ้ลคลิกเพื่อแก้ไขตัวสะกดชื่อบริษัท",
                    required=True,
                    width="large"
                ),
                "BU": st.column_config.TextColumn(
                    "🏷️ BU (ตัวย่อ) [ดับเบิ้ลคลิกเพื่อแก้ไข]",
                    help="ดับเบิ้ลคลิกเพื่อแก้ไขตัวย่อ BU",
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
            if st.button("💾 บันทึกการแก้ไข Company & BU ทั้งหมด", type="primary", use_container_width=True, key="m_btn_save_grid"):
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
                    st.error("⚠️ ข้อมูลบริษัทต้องไม่เป็นค่าว่าง")
                elif has_duplicate:
                    st.warning("⚠️ มีชื่อบริษัทซ้ำกันในตาราง (ระบบล็อก 1 บริษัท = 1 BU)")
                else:
                    master_data["companies"] = new_comp_list
                    if save_master_data(master_data):
                        st.toast("✅ บันทึกการแก้ไข Company & BU สำเร็จเรียบร้อย!", icon="💾")
                        st.rerun()

        with col_reload_grid:
            if st.button("🔄 โหลดข้อมูลล่าสุด", use_container_width=True, key="m_btn_reload_grid"):
                st.toast("🔄 โหลดข้อมูลล่าสุดเรียบร้อย", icon="🔄")
                st.rerun()

        # Optional Quick Add Expander
        with st.expander("➕ หรือเพิ่มบริษัทใหม่ผ่านฟอร์มด่วน", expanded=False):
            col_in_c1, col_in_c2 = st.columns([2, 1])
            with col_in_c1:
                input_comp_name = st.text_input("ชื่อบริษัท (Company Name)*", placeholder="เช่น Test Automobile Co., Ltd.", key="m_input_comp")
            with col_in_c2:
                input_bu_name = st.text_input("ตัวย่อ BU*", placeholder="เช่น TAB", key="m_input_bu")

            if input_comp_name.strip():
                match_existing = next((c for c in companies_list if c["Company"].strip().lower() == input_comp_name.strip().lower()), None)
                if match_existing:
                    st.info(f"💡 **พบข้อมูลเดิม**: ปัจจุบัน BU คือ `{match_existing['BU']}` — บันทึกจะเป็นการ **อัปเดต BU**")
                else:
                    st.success(f"✅ **ชื่อบริษัทใหม่**: พร้อมเพิ่มเข้าสู่ระบบ")

            if st.button("➕ เพิ่มบริษัทเข้าตาราง", type="secondary", key="m_btn_add_comp_form"):
                if not input_comp_name.strip() or not input_bu_name.strip():
                    st.error("⚠️ กรุณากรอกทั้งชื่อบริษัทและตัวย่อ BU")
                else:
                    c_clean = input_comp_name.strip()
                    bu_clean = input_bu_name.strip()
                    found = False
                    for item in companies_list:
                        if item["Company"].strip().lower() == c_clean.lower():
                            item["BU"] = bu_clean
                            found = True
                            break
                    if not found:
                        companies_list.append({"Company": c_clean, "BU": bu_clean})
                    master_data["companies"] = companies_list
                    if save_master_data(master_data):
                        st.toast(f"✅ บันทึกบริษัท '{c_clean}' (BU: {bu_clean}) สำเร็จ!", icon="🏢")
                        st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ------------------------------------------
    # Sub-tab 2: Branches (Tag Chips + Instant Deletion)
    # ------------------------------------------
    with subtab_branch:
        st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📍 จัดการรายชื่อสาขามาตรฐาน (Branch Registry)</div>', unsafe_allow_html=True)
        st.caption("คลิกปุ่ม ✕ บนป้ายสาขาเพื่อลบออกทันที หรือพิมพ์ชื่อสาขาใหม่เพื่อเพิ่มลงในระบบ")

        branches_list = master_data.get("branches", [])

        # Quick Add with Real-time validation
        col_b_in, col_b_btn = st.columns([3, 1])
        with col_b_in:
            new_branch_input = st.text_input("➕ เพิ่มสาขาใหม่", placeholder="พิมพ์ชื่อสาขา เช่น ระยอง, ขอนแก่น, สาทร...", key="m_new_branch_in")
            if new_branch_input.strip():
                if new_branch_input.strip() in branches_list:
                    st.warning(f"⚠️ มีสาขา '{new_branch_input.strip()}' อยู่ในระบบแล้ว")
                else:
                    st.success(f"✅ สาขา '{new_branch_input.strip()}' พร้อมเพิ่มลงในระบบ")

        with col_b_btn:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ เพิ่มสาขา", type="primary", use_container_width=True, key="m_btn_add_br"):
                b_clean = new_branch_input.strip()
                if not b_clean:
                    st.error("กรุณากรอกชื่อสาขา")
                elif b_clean in branches_list:
                    st.warning("มีชื่อสาขานี้อยู่ในระบบแล้ว")
                else:
                    branches_list.append(b_clean)
                    master_data["branches"] = branches_list
                    if save_master_data(master_data):
                        st.toast(f"📍 เพิ่มสาขา '{b_clean}' สำเร็จ!", icon="📍")
                        st.rerun()

        st.markdown("---")
        # Search & Count
        col_bs, col_bc = st.columns([2.5, 1.5])
        with col_bs:
            search_branch_txt = st.text_input("🔍 กรองค้นหาสาขา...", placeholder="พิมพ์ชื่อสาขาเพื่อกรอง...", key="m_br_search")
        with col_bc:
            st.markdown("<div style='height: 32px;'></div>", unsafe_allow_html=True)
            st.caption(f"📊 ทั้งหมด **{len(branches_list)}** สาขา")

        filtered_branches = branches_list
        if search_branch_txt.strip():
            filtered_branches = [b for b in branches_list if search_branch_txt.strip().lower() in b.lower()]

        # Tag Chips Grid
        st.markdown("##### 🏷️ รายชื่อสาขา (คลิก ✕ เพื่อลบ):")
        chip_cols = st.columns(4)
        for idx, br_name in enumerate(filtered_branches):
            col_target = chip_cols[idx % 4]
            with col_target:
                if col_target.button(f"📍 {br_name} ✕", key=f"m_chip_del_{br_name}", use_container_width=True, help=f"คลิกเพื่อลบสาขา {br_name}"):
                    branches_list.remove(br_name)
                    master_data["branches"] = branches_list
                    if save_master_data(master_data):
                        st.toast(f"🗑️ ลบสาขา '{br_name}' เรียบร้อยแล้ว", icon="🗑️")
                        st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ------------------------------------------
    # Sub-tab 3: Operators / Senders
    # ------------------------------------------
    with subtab_ops:
        st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">👤 จัดการรายชื่อผู้ส่ง Email (Sender Operators)</div>', unsafe_allow_html=True)
        
        ops_list = master_data.get("operators", [])
        st.markdown("##### 📋 บัญชีผู้ส่งที่ใช้งานได้ปัจจุบัน:")
        for op in ops_list:
            st.markdown(f"- ✉️ `{op}`")

        col_add_op, col_del_op = st.columns(2)
        with col_add_op:
            st.markdown("##### ➕ เพิ่มบัญชีผู้ส่ง")
            with st.form("m_form_add_op"):
                new_op_email = st.text_input("อีเมลผู้ส่ง (Office 365)*", placeholder="เช่น name.sur@i24.co.th")
                btn_save_op = st.form_submit_button("💾 เพิ่มผู้ส่ง")
                
                if btn_save_op:
                    if not new_op_email.strip() or "@" not in new_op_email:
                        st.error("กรุณากรอกอีเมลที่ถูกต้อง")
                    elif new_op_email.strip() in ops_list:
                        st.warning("มีอีเมลนี้อยู่ในระบบแล้ว")
                    else:
                        ops_list.append(new_op_email.strip())
                        master_data["operators"] = ops_list
                        if "templates" not in master_data:
                            master_data["templates"] = {}
                        if new_op_email.strip() not in master_data["templates"]:
                            master_data["templates"][new_op_email.strip()] = json.loads(json.dumps(DEFAULT_MASTER_DATA["templates"]["nawarutte.non@i24.co.th"]))
                        if save_master_data(master_data):
                            st.toast(f"👤 เพิ่มผู้ส่ง '{new_op_email.strip()}' สำเร็จ!", icon="👤")
                            st.rerun()

        with col_del_op:
            st.markdown("##### 🗑️ ลบบัญชีผู้ส่ง")
            with st.form("m_form_del_op"):
                del_op_email = st.selectbox("เลือกบัญชีที่ต้องการลบ", ops_list if ops_list else [""])
                btn_del_op = st.form_submit_button("🗑️ ยืนยันลบผู้ส่ง")
                
                if btn_del_op:
                    if len(ops_list) <= 1:
                        st.error("ต้องมีผู้ส่งในระบบอย่างน้อย 1 คน")
                    elif del_op_email and del_op_email in ops_list:
                        ops_list.remove(del_op_email)
                        master_data["operators"] = ops_list
                        if del_op_email in master_data.get("templates", {}):
                            del master_data["templates"][del_op_email]
                        if save_master_data(master_data):
                            st.toast(f"🗑️ ลบผู้ส่ง '{del_op_email}' สำเร็จ!", icon="🗑️")
                            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ------------------------------------------
    # Sub-tab 4: Personalized Email Template Editor
    # ------------------------------------------
    with subtab_tpl:
        st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📧 ปรับแต่งเทมเพลตอีเมลแยกรายบุคคล (Persona-Based Templates)</div>', unsafe_allow_html=True)
        
        ops_list = master_data.get("operators", ["nawarutte.non@i24.co.th", "pawitporn.sae@i24.co.th"])
        app_list = ["VSM", "E-Travelling", "Forma", "Red plate", "Pandora"]

        col_sel_user, col_sel_app = st.columns(2)
        with col_sel_user:
            edit_tpl_user = st.selectbox("👤 1. เลือก User ผู้ส่ง", ops_list, index=0, key="m_tpl_user_sel")
        with col_sel_app:
            edit_tpl_app = st.selectbox("📱 2. เลือกระบบ (Application)", app_list, index=0, key="m_tpl_app_sel")

        current_tpl = get_operator_template(master_data, edit_tpl_user, edit_tpl_app)

        with st.form("m_form_edit_email_tpl"):
            st.markdown(f"#### ✍️ เทมเพลตของ `{edit_tpl_user}` สำหรับระบบ `{edit_tpl_app}`")
            
            in_subject = st.text_input("📝 Subject (หัวข้ออีเมล)*", value=current_tpl.get("subject", f"ข้อมูลการเข้าระบบ {edit_tpl_app}"))
            in_greeting = st.text_input("👋 คำขึ้นต้น / คำทักทาย*", value=current_tpl.get("greeting", "เรียน ผู้ใช้งานระบบ"))
            in_intro = st.text_area("💬 ข้อความเกริ่นนำ (Intro Text)*", value=current_tpl.get("intro", f"ให้เข้าใช้งานโดย User & Password ตามด้านล่างนี้ครับ"), height=110)
            
            st.info("💡 หมายเหตุ: บรรทัด User, Password และ Link จะถูกสร้างและจัดรูปแบบให้อัตโนมัติตามประเภทโปรแกรม")

            col_btn_tpl1, col_btn_tpl2 = st.columns(2)
            btn_save_tpl = col_btn_tpl1.form_submit_button("💾 บันทึก Template นี้", type="primary", use_container_width=True)
            btn_reset_tpl = col_btn_tpl2.form_submit_button("🔄 คืนค่าเริ่มต้นของ App นี้", use_container_width=True)

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
                    st.toast(f"📧 บันทึก Template '{edit_tpl_app}' ของ '{edit_tpl_user}' สำเร็จ!", icon="💾")
                    st.rerun()

            if btn_reset_tpl:
                def_tpl = DEFAULT_MASTER_DATA["templates"]["nawarutte.non@i24.co.th"].get(edit_tpl_app, {})
                if "templates" in master_data and edit_tpl_user in master_data["templates"]:
                    master_data["templates"][edit_tpl_user][edit_tpl_app] = def_tpl
                    if save_master_data(master_data):
                        st.toast(f"🔄 คืนค่าเริ่มต้น Template '{edit_tpl_app}' สำเร็จ!", icon="🔄")
                        st.rerun()

        # Real-time Live Preview
        st.markdown("---")
        st.markdown("##### 👁️ ตัวอย่างผลลัพธ์ของ Template นี้ (Live Sample Preview):")
        sample_link = "https://vsm.mgc-asia.com/Pages/Home.aspx" if edit_tpl_app == "VSM" else ("http://travelling.mgc-asia.com/" if edit_tpl_app == "E-Travelling" else ("https://redplate-frontend.azurewebsites.net/..." if edit_tpl_app == "Red plate" else ""))
        sample_plain, sample_html = build_email_body(edit_tpl_app, "somchai.pra", "p@ssw0rdsom", sample_link, custom_template={"greeting": in_greeting, "intro": in_intro})
        st.code(f"""To:       user.example@mgc-asia.com
Subject:  {in_subject}
----------------------------------------
{sample_plain}""", language="text")
        st.markdown('</div>', unsafe_allow_html=True)

    # ------------------------------------------
    # Sub-tab 5: Backup & Restore Master Data
    # ------------------------------------------
    with subtab_backup:
        st.markdown('<div class="enterprise-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">💾 สำรองและกู้คืน Master Data (Backup & Restore)</div>', unsafe_allow_html=True)
        
        json_str = json.dumps(master_data, ensure_ascii=False, indent=2)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์สำรอง (Export master_data.json)",
            data=json_str,
            file_name="master_data.json",
            mime="application/json",
            use_container_width=True
        )

        st.markdown("---")
        st.markdown("##### 📤 กู้คืน / นำเข้าไฟล์ Master Data (Import JSON):")
        uploaded_json = st.file_uploader("เลือกไฟล์ master_data.json ที่สำรองไว้", type=["json"], key="m_upload_json")
        if uploaded_json is not None:
            try:
                imported_data = json.load(uploaded_json)
                if st.button("🔄 ยืนยันนำเข้าข้อมูลนี้ (Apply Imported Data)", type="primary", key="m_btn_apply_import"):
                    if save_master_data(imported_data):
                        st.toast("✅ นำเข้าข้อมูล Master Data สำเร็จ!", icon="📥")
                        st.rerun()
            except Exception as ex:
                st.error(f"ไฟล์ JSON ไม่ถูกต้อง: {ex}")

        st.markdown("---")
        if st.button("⚠️ คืนค่า Master Data ทั้งหมดเป็นค่าเริ่มต้นจากโรงงาน (Reset All to Factory Defaults)", key="m_btn_factory_reset"):
            if save_master_data(DEFAULT_MASTER_DATA):
                st.toast("🔄 คืนค่า Master Data ทั้งหมดเป็นค่าเริ่มต้นเรียบร้อยแล้ว!", icon="🔄")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
