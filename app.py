import os
import io
import time
import json
import urllib.request
from datetime import datetime
import streamlit as st

# ==========================================
# 1. טעינת משתני סביבה מ-Secrets ל-os.environ
# ==========================================
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", ""))
CHROMA_KEY = st.secrets.get("CHROMA_OPENAI_API_KEY", OPENAI_API_KEY)
GOOGLE_DRIVE_FOLDER_ID = st.secrets.get("GOOGLE_DRIVE_FOLDER_ID", os.getenv("GOOGLE_DRIVE_FOLDER_ID", "1Q_isBiLXre0PGEFsc7Y9CS3EU8W1DzB2"))

if OPENAI_API_KEY:
    os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY
    os.environ["CHROMA_OPENAI_API_KEY"] = CHROMA_KEY

from openai import OpenAI
import chromadb
from chromadb.utils import embedding_functions

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials

LOGO_URL = "https://upload.wikimedia.org/wikipedia/commons/e/e2/Israel_Defense_Ministry_Symbol.png"
SCOPES = ['https://www.googleapis.com/auth/drive']
CHROMA_DB_PATH = "./chroma_db"
SYNC_LOG_FILE = "sync_log.json"

st.set_page_config(
    page_title="מערכת ניסוח הסכמים - יחידת מתוו\"ה",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 2. עיצוב CSS לניגודיות וקריאות מקסימלית
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Rubik:wght@300;400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Rubik', sans-serif;
        direction: rtl;
        text-align: right;
    }
    
    .stApp {
        background-color: #f8fafc !important;
        background-image: linear-gradient(rgba(248, 250, 252, 0.94), rgba(248, 250, 252, 0.94)), 
                          url('https://upload.wikimedia.org/wikipedia/commons/e/e2/Israel_Defense_Ministry_Symbol.png');
        background-repeat: no-repeat;
        background-position: center center;
        background-size: 500px auto;
        background-attachment: fixed;
    }

    h1, h2, h3, h4, h5, h6, p, label, span, div {
        color: #0f172a !important;
    }

    .stTextInput > label, .stSelectbox > label, .stTextArea > label {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
    }

    .stTextInput input, .stTextArea textarea, div[data-baseweb="select"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 2px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }

    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 2rem 2.5rem;
        border-radius: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2);
        margin-bottom: 2rem;
        border-bottom: 4px solid #3b82f6;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    .main-header h1, .main-header p {
        color: #ffffff !important;
    }

    .css-card {
        background: rgba(255, 255, 255, 0.92);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(226, 232, 240, 0.9);
        border-radius: 16px;
        padding: 1.8rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        margin-bottom: 1.5rem;
    }

    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        border-radius: 12px !important;
        padding: 0.75rem 2rem !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3) !important;
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. פונקציות עזר להורדת לוגו ויצירת קבצים
# ==========================================
def get_logo_bytes():
    try:
        req = urllib.request.Request(LOGO_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            return response.read()
    except Exception:
        return None

class PDFWithHeaderFooter(FPDF):
    def header(self):
        self.set_font("Arial", 'B', 10)
        self.cell(0, 8, "בלתי מסווג", align='C', ln=True)
        self.set_font("Arial", '', 9)
        self.cell(0, 5, f"תאריך: {datetime.now().strftime('%d/%m/%Y')}", align='R', ln=True)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", 'I', 8)
        self.cell(0, 10, f"עמוד {self.page_no()}", align='C')

def create_docx(text_content, category_name, party_a, party_b):
    doc = Document()
    
    # הוספת סיווג בכותרת עליונה
    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hrun = hp.add_run("בלתי מסווג")
    hrun.font.bold = True
    hrun.font.size = Pt(10)
    
    # הוספת לוגו משרד הביטחון
    logo_bytes = get_logo_bytes()
    if logo_bytes:
        logo_stream = io.BytesIO(logo_bytes)
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_stream, width=Inches(1.2))

    # כותרת ותאריך
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title_p.add_run(f"משרד הביטחון - יחידת מתוו\"ה\n{category_name}")
    t_run.font.bold = True
    t_run.font.size = Pt(16)
    
    details_p = doc.add_paragraph()
    details_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d_run = details_p.add_run(f"תאריך ניסוח: {datetime.now().strftime('%d/%m/%Y')}\nבין: {party_a} לבין: {party_b}\n")
    d_run.font.size = Pt(10)
    
    # תוכן המסמך
    for paragraph in text_content.split('\n'):
        if paragraph.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(paragraph)
            
    # כותרת תחתית - סיווג
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = fp.add_run("בלתי מסווג")
    frun.font.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_pdf(text_content, category_name, party_a, party_b):
    pdf = PDFWithHeaderFooter()
    pdf.add_page()
    
    pdf.set_font("Arial", 'B', 14)
    pdf.multi_cell(0, 8, f"משרד הביטחון - יחידת מתוו\"ה\n{category_name}".encode('latin-1', 'replace').decode('latin-1'), align='C')
    pdf.ln(3)
    
    pdf.set_font("Arial", '', 10)
    pdf.multi_cell(0, 6, f"בין: {party_a} לבין: {party_b}".encode('latin-1', 'replace').decode('latin-1'), align='R')
    pdf.ln(5)
    
    lines = text_content.split('\n')
    for line in lines:
        cleaned_line = line.encode('latin-1', 'replace').decode('latin-1')
        pdf.multi_cell(0, 6, cleaned_line, align='R')
        
    return pdf.output(dest='S').encode('latin-1', 'replace')

@st.cache_resource
def get_openai_client():
    if not OPENAI_API_KEY:
        st.error("אנא הגדר את OPENAI_API_KEY ב-Secrets של Streamlit.")
        return None
    return OpenAI(api_key=OPENAI_API_KEY)

@st.cache_resource
def get_vector_db():
    client = chromadb.PersistentClient(path=CHROMA_DB_PATH)
    openai_ef = embedding_functions.OpenAIEmbeddingFunction(
        api_key=OPENAI_API_KEY,
        model_name="text-embedding-3-small"
    )
    return client.get_or_create_collection(
        name="legal_agreements",
        embedding_function=openai_ef
    )

def get_drive_service():
    if "GOOGLE_TOKEN_JSON" in st.secrets:
        try:
            token_raw = st.secrets["GOOGLE_TOKEN_JSON"]
            token_info = json.loads(token_raw) if isinstance(token_raw, str) else dict(token_raw)
            creds = UserCredentials.from_authorized_user_info(token_info, SCOPES)
            if creds and creds.valid:
                return build('drive', 'v3', credentials=creds)
            elif creds and creds.expired and creds.refresh_token:
                from google.auth.transport.requests import Request
                creds.refresh(Request())
                return build('drive', 'v3', credentials=creds)
        except Exception as e:
            st.error(f"שגיאה בפענוח GOOGLE_TOKEN_JSON: {e}")

    if "GCP_SERVICE_ACCOUNT" in st.secrets:
        try:
            creds_raw = st.secrets["GCP_SERVICE_ACCOUNT"]
            creds_dict = json.loads(creds_raw) if isinstance(creds_raw, str) else dict(creds_raw)
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
            return build('drive', 'v3', credentials=creds)
        except Exception as e:
            st.error(f"שגיאה בחיבור דרך Service Account: {e}")

    if os.path.exists('token.json'):
        try:
            creds = UserCredentials.from_authorized_user_file('token.json', SCOPES)
            if creds and creds.valid:
                return build('drive', 'v3', credentials=creds)
        except Exception:
            pass

    return None

def upload_text_as_file_to_drive(file_title, text_content, party_a, party_b):
    drive_service = get_drive_service()
    if not drive_service:
        return False, "לא נמצאו הרשאות חיבור תקפות ב-Streamlit Secrets (נא לוודא את תקינות GOOGLE_TOKEN_JSON)."

    try:
        timestamp_str = datetime.now().strftime('%Y%m%d_%H%M')
        formatted_filename = f"{file_title}_בין_{party_a}_לבין_{party_b}_{timestamp_str}.txt"
        
        file_metadata = {
            'name': formatted_filename,
            'parents': [GOOGLE_DRIVE_FOLDER_ID]
        }
        media = MediaIoBaseUpload(io.BytesIO(text_content.encode('utf-8')), mimetype='text/plain')
        uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True, uploaded_file.get('id')
    except Exception as e:
        return False, str(e)

def load_sync_log():
    if os.path.exists(SYNC_LOG_FILE):
        with open(SYNC_LOG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"processed_files": [], "last_sync": None, "agreement_types": []}

def generate_strict_contract_proposal(selected_category, party_a, party_b, special_terms, collection, client):
    search_query = f"{selected_category} {special_terms}"
    results = collection.query(query_texts=[search_query], n_results=15)

    retrieved_context = ""
    sources_used = set()

    if results and results.get("documents"):
        for docs, metass in zip(results["documents"], results["metadatas"]):
            for doc, meta in zip(docs, metass):
                retrieved_context += f"\n[מתוך: {meta.get('filename')} | סוג: {meta.get('category_type', 'כללי')}]\n{doc}\n"
                sources_used.add(meta.get('filename'))

    prompt = f"""
אתה מנוע משפטי מתקדם של יחידת מתוו"ה במשרד הביטחון.
עליך לנסח הצעה להסכם מסוג: {selected_category}

ההסכם חייב להתבסס **אך ורק** על הסגנון, הטרמינולוגיה, המבנה והסעיפים המשפטיים מתוך קובצי ה-Drive שנלמדו.

פרטי העסקה:
- סוג ההסכם המבוקש: {selected_category}
- צד א': {party_a}
- צד ב': {party_b}
- תנאים מיוחדים/שינויים: {special_terms}

סעיפים רלוונטיים שנשלפו מתוך ארכיון ה-Drive:
=========================================
{retrieved_context if retrieved_context else "לא נמצאו סעיפים ספציפיים."}
=========================================

הנחיות לניסוח:
1. ניסוח משפטי מחייב, מדויק, כולל פתיח (הואיל ו...), סעיפי התחייבות, ערבויות וסמכות שיפוט.
2. שלב באופן טבעי את שמות הצדדים ({party_a} ו-{party_b}) והדגשים המיוחדים ({special_terms}).
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2
    )

    return response.choices[0].message.content, retrieved_context, list(sources_used)

# --- UI ---
st.markdown("""
    <div class="main-header">
        <div>
            <h1>🏛️ משרד הביטחון - יחידת מתוו"ה</h1>
            <p>מערכת ארגונית לניסוח, ניתוח ולמידת הסכמים מול Google Drive המשרדי</p>
        </div>
        <div style="font-size: 2.5rem;">📜</div>
    </div>
""", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📝 ניסוח הסכם מבוסס Drive", "📤 העלאת קבצים לדרייב", "🔄 סטטוס וסנכרון"])

client = get_openai_client()
collection = get_vector_db()
log_data = load_sync_log()
agreement_types = log_data.get("agreement_types", [])

with tab1:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("📝 בחירת סוג הסכם מתוך מאגר הידע המשרדי ב-Drive")
    
    col1, col2 = st.columns(2)
    with col1:
        if agreement_types:
            selected_category = st.selectbox("🎯 בחר סוג הסכם להתבססות:", agreement_types)
        else:
            selected_category = st.text_input("סוג ההסכם", value="לדוגמא: אספקת עוגות לתחום GIS")
        
        party_a = st.text_input("שם צד א' (למשל: משרד הביטחון / מזמין)")

    with col2:
        party_b = st.text_input("שם צד ב' (למשל: הקבלן / הספק)")

    special_terms = st.text_area("✍️ תנאים מיוחדים ודגשים למסמך", placeholder="פרט תאריכים, סכומים, ערבויות או דגשים מיוחדים...")

    if st.button("🚀 נסח הצעה להסכם"):
        if not party_a or not party_b:
            st.warning("אנא מלא את שמות הצדדים.")
        else:
            with st.spinner("שולף סעיפים תואמים מתוך הדרייב המשרדי ומנסח מסמך..."):
                proposal, context_used, sources = generate_strict_contract_proposal(
                    selected_category, party_a, party_b, special_terms, collection, client
                )
                
                st.session_state["current_proposal"] = proposal
                st.session_state["edited_proposal"] = proposal
                st.session_state["current_category"] = selected_category
                st.session_state["party_a"] = party_a
                st.session_state["party_b"] = party_b
                st.session_state["show_edit_box"] = False
                
                st.success("המסמך נוסח בהצלחה!")
                if sources:
                    st.info(f"📚 מסמכי המקור בדרייב המשרדי ששימשו בסיס לניסוח: {', '.join(sources)}")

    if "current_proposal" in st.session_state:
        st.markdown("### 📄 טיוטת המסמך המוצעת:")
        
        current_text = st.session_state.get("edited_proposal", st.session_state["current_proposal"])
        st.text_area("תוכן ההסכם (תצוגה)", value=current_text, height=350, key="display_area")

        btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)

        cat = st.session_state.get('current_category', 'הסכם')
        pa = st.session_state.get('party_a', 'צד א')
        pb = st.session_state.get('party_b', 'צד ב')

        with btn_col1:
            if st.button("✏️ ערוך את ההסכם המוצע"):
                st.session_state["show_edit_box"] = not st.session_state.get("show_edit_box", False)

        with btn_col2:
            docx_data = create_docx(current_text, cat, pa, pb)
            st.download_button(
                label="📄 יצא ל-WORD",
                data=docx_data,
                file_name=f"{cat}_בין_{pa}_לבין_{pb}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        with btn_col3:
            pdf_data = create_pdf(current_text, cat, pa, pb)
            st.download_button(
                label="Picture/PDF יצא ל-PDF",
                data=pdf_data,
                file_name=f"{cat}_בין_{pa}_לבין_{pb}.pdf",
                mime="application/pdf"
            )

        with btn_col4:
            if st.button("💾 שמור ב-Drive"):
                with st.spinner("מעלה את המסמך המנוסח לתיקייה בדרייב..."):
                    success, res_msg = upload_text_as_file_to_drive(
                        cat,
                        current_text,
                        pa,
                        pb
                    )
                    if success:
                        st.success("🎉 המסמך נשמר בהצלחה ב-Google Drive!")
                    else:
                        st.error(f"שגיאה בעת שמירת המסמך לדרייב: {res_msg}")

        if st.session_state.get("show_edit_box", False):
            st.markdown("---")
            st.subheader("✏️ עריכת הטקסט המוצע:")
            edited_text = st.text_area("בצע את השינויים הרצויים כאן:", value=current_text, height=400, key="edit_area")
            if st.button("✅ אישור ועדכון השינויים"):
                st.session_state["edited_proposal"] = edited_text
                st.session_state["show_edit_box"] = False
                st.success("הטקסט עודכן בהצלחה!")
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("📤 העלאת קובצי PDF סרוקים ישירות ל-Google Drive המשרדי")
    uploaded_files = st.file_uploader("בחר קובצי PDF להעלאה", type=["pdf"], accept_multiple_files=True)
    if st.button("עלה ל-Drive"):
        st.info("הקבצים יעלו לתיקייה המשרדית ויעובדו בסנכרון הבא.")
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("🔄 סטטוס הדרייב המשרדי והסכמים שנלמדו")
    st.write(f"📂 מזהה התיקייה בדרייב: **`{GOOGLE_DRIVE_FOLDER_ID}`**")
    st.write(f"📅 סנכרון אחרון: **{log_data.get('last_sync', 'טרם בוצע')}**")
    st.write(f"📄 סך הכל קבצים שעובדו: **{len(log_data.get('processed_files', []))}**")
    
    if agreement_types:
        st.markdown("### 🏷️ סוגי ההסכמים הזמינים במערכת:")
        for t in agreement_types:
            st.markdown(f"- **`{t}`**")
    st.markdown('</div>', unsafe_allow_html=True)
