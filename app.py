import os
import io
import time
import json
import urllib.request
from datetime import datetime
import streamlit as st

# ==========================================
# 1. הגדרות סביבה ותקשורת ל-Drive
# ==========================================
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", ""))
CHROMA_KEY = st.secrets.get("CHROMA_OPENAI_API_KEY", OPENAI_API_KEY)

# 1. תיקיית לימוד - מאגר הידע
LEARNING_FOLDER_ID = st.secrets.get("LEARNING_FOLDER_ID", "1HS1MqHNvRyHeRuLZWWQQ8ClSBFgoRZQs")

# 2. תיקיית פלט - לשמירת הסכמים חדשים
OUTPUT_FOLDER_ID = st.secrets.get("OUTPUT_FOLDER_ID", "1Q_isBiLXre0PGEFsc7Y9CS3EU8W1DzB2")

if OPENAI_API_KEY:
    os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY
    os.environ["CHROMA_OPENAI_API_KEY"] = CHROMA_KEY

from openai import OpenAI
import chromadb
from chromadb.utils import embedding_functions

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials

import pypdf

LOGO_URL = "https://upload.wikimedia.org/wikipedia/commons/e/e2/Israel_Defense_Ministry_Symbol.png"
SCOPES = ['https://www.googleapis.com/auth/drive']
CHROMA_DB_PATH = "./chroma_db"
SYNC_LOG_FILE = "sync_log.json"

st.set_page_config(
    page_title="משרד הביטחון - מערכת ניסוח הסכמים",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 2. עיצוב מתקדם עם ניגודיות מושלמת (CSS)
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Assistant:wght@300;400;600;700;800&family=Rubik:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Assistant', 'Rubik', sans-serif !important;
        direction: rtl !important;
        text-align: right !important;
        background-color: #f4f6f9 !important;
    }

    .stApp {
        background-color: #f4f6f9 !important;
    }

    /* Header רשמי של משרד הביטחון */
    .mod-official-header {
        background: linear-gradient(135deg, #0b1e36 0%, #1c3a60 100%);
        border-radius: 12px;
        padding: 1.8rem 2.5rem;
        color: #ffffff !important;
        box-shadow: 0 10px 25px rgba(11, 30, 54, 0.2);
        margin-bottom: 2rem;
        border-bottom: 4px solid #d4af37;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    /* כותרת ראשית בלבן בוהק */
    .mod-official-header h1, .mod-official-header .header-title {
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: 2.3rem !important;
        margin: 0 0 0.3rem 0 !important;
        letter-spacing: -0.5px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.4);
    }

    .mod-official-header p {
        color: #e2e8f0 !important;
        font-size: 1.05rem !important;
        margin: 0 !important;
    }

    .classification-tag {
        background-color: #dc2626;
        color: #ffffff !important;
        font-weight: 800;
        font-size: 0.85rem;
        padding: 0.3rem 1.2rem;
        border-radius: 4px;
        letter-spacing: 1.5px;
        display: inline-block;
        margin-bottom: 0.6rem;
    }

    .mod-card {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-top: 3px solid #1c3a60 !important;
        border-radius: 12px !important;
        padding: 1.8rem !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04) !important;
        margin-bottom: 1.5rem !important;
    }

    h1, h2, h3, h4, h5, h6, label, span, div {
        color: #0b1e36 !important;
    }

    .stTextInput > label, .stSelectbox > label, .stTextArea > label, .stRadio > label {
        color: #0b1e36 !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
    }

    .stTextInput input, .stTextArea textarea, div[data-baseweb="select"] {
        background-color: #ffffff !important;
        color: #0b1e36 !important;
        border: 1.5px solid #94a3b8 !important;
        border-radius: 8px !important;
        padding: 0.6rem 1rem !important;
    }

    /* כפתורים עם טקסט לבן בולט */
    .stButton > button {
        background: linear-gradient(135deg, #1c3a60 0%, #0b1e36 100%) !important;
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: 1.15rem !important;
        border-radius: 8px !important;
        padding: 0.85rem 2rem !important;
        border: 1.5px solid #d4af37 !important;
        box-shadow: 0 4px 12px rgba(11, 30, 54, 0.3) !important;
        width: 100%;
        transition: all 0.25s ease-in-out !important;
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, #0b1e36 0%, #1c3a60 100%) !important;
        color: #fef08a !important;
        border-color: #ffffff !important;
        box-shadow: 0 6px 18px rgba(11, 30, 54, 0.45) !important;
    }

    .stButton > button p, .stButton > button span {
        color: #ffffff !important;
        font-weight: 800 !important;
    }

    .metric-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-right: 4px solid #1c3a60;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }

    .metric-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #1c3a60;
    }

    .metric-label {
        font-size: 0.9rem;
        color: #64748b;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. פונקציות יצירת קבצים והורדת לוגו
# ==========================================
def get_logo_bytes():
    try:
        req = urllib.request.Request(LOGO_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            return response.read()
    except Exception:
        return None

def create_docx(text_content, category_name, party_a, party_b):
    doc = Document()
    
    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hrun = hp.add_run("בלתי מסווג")
    hrun.font.bold = True
    
    logo_bytes = get_logo_bytes()
    if logo_bytes:
        logo_stream = io.BytesIO(logo_bytes)
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_stream, width=Inches(1.2))

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title_p.add_run(f"משרד הביטחון - יחידת מתוו\"ה\n{category_name}")
    t_run.font.bold = True
    t_run.font.size = Pt(16)
    
    details_p = doc.add_paragraph()
    details_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d_run = details_p.add_run(f"תאריך ניסוח: {datetime.now().strftime('%d/%m/%Y')}\nבין: {party_a} לבין: {party_b}\n")
    d_run.font.size = Pt(10)
    
    for paragraph in text_content.split('\n'):
        if paragraph.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(paragraph)
            
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = fp.add_run("בלתי מסווג")
    frun.font.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_pdf(text_content, category_name, party_a, party_b):
    full_pdf_text = f"""==================================================
              משרד הביטחון - יחידת מתוו"ה
                     בלתי מסווג
==================================================
תאריך ניסוח: {datetime.now().strftime('%d/%m/%Y')}
סוג ההסכם: {category_name}
בין: {party_a}
לבין: {party_b}
--------------------------------------------------

{text_content}

==================================================
                     בלתי מסווג
=================================================="""
    return full_pdf_text.encode('utf-8')

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
        return False, "לא נמצאו הרשאות חיבור תקפות ב-Streamlit Secrets."

    try:
        timestamp_str = datetime.now().strftime('%Y%m%d_%H%M')
        formatted_filename = f"{file_title}_בין_{party_a}_לבין_{party_b}_{timestamp_str}.txt"
        
        file_metadata = {
            'name': formatted_filename,
            'parents': [OUTPUT_FOLDER_ID]
        }
        media = MediaIoBaseUpload(io.BytesIO(text_content.encode('utf-8')), mimetype='text/plain')
        uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True, uploaded_file.get('id')
    except Exception as e:
        return False, str(e)

# ==========================================
# 4. מנגנון סריקה עמוקה ואינדוקס RAG
# ==========================================
def extract_text_from_file_bytes(file_bytes, mime_type, file_name):
    text = ""
    try:
        if "pdf" in mime_type or file_name.endswith(".pdf"):
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        elif "word" in mime_type or file_name.endswith(".docx"):
            doc = Document(io.BytesIO(file_bytes))
            for p in doc.paragraphs:
                text += p.text + "\n"
        else:
            text = file_bytes.decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"Error extracting text from {file_name}: {e}")
    return text

def chunk_text(text, chunk_size=800, overlap=150):
    words = text.split()
    chunks = []
    for i in range(0, len(words), chunk_size - overlap):
        chunk = " ".join(words[i:i + chunk_size])
        if chunk.strip():
            chunks.append(chunk)
    return chunks

def list_all_files_recursive(drive_service, folder_id):
    all_files = []
    try:
        query = f"'{folder_id}' in parents and trashed = false"
        results = drive_service.files().list(q=query, pageSize=100, fields="files(id, name, mimeType)").execute()
        items = results.get('files', [])

        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                all_files.extend(list_all_files_recursive(drive_service, item['id']))
            else:
                all_files.append(item)
    except Exception as e:
        print(f"Error listing files recursively: {e}")

    return all_files

def sync_and_index_rag_drive_files(collection, drive_service):
    log_data = load_sync_log()
    processed_files = set(log_data.get("processed_files", []))

    if not drive_service:
        return log_data

    try:
        files = list_all_files_recursive(drive_service, LEARNING_FOLDER_ID)
        all_categories = set(log_data.get("agreement_types", []))

        for file in files:
            file_id = file['id']
            file_name = file['name']

            clean_name = file_name.replace(".pdf", "").replace(".docx", "").replace(".txt", "")
            category = clean_name.split('_')[0].strip()
            all_categories.add(category)

            if file_name not in processed_files:
                try:
                    request = drive_service.files().get_media(fileId=file_id)
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()

                    file_content_text = extract_text_from_file_bytes(fh.getvalue(), file.get('mimeType', ''), file_name)

                    if file_content_text.strip():
                        chunks = chunk_text(file_content_text)
                        documents = []
                        metadatas = []
                        ids = []

                        for idx, chunk in enumerate(chunks):
                            documents.append(chunk)
                            metadatas.append({"filename": file_name, "category_type": category})
                            ids.append(f"{file_id}_{idx}")

                        collection.add(
                            documents=documents,
                            metadatas=metadatas,
                            ids=ids
                        )
                    processed_files.add(file_name)
                except Exception as ex:
                    print(f"Skipping binary/non-supported file {file_name}: {ex}")

        log_data = {
            "processed_files": list(processed_files),
            "last_sync": datetime.now().strftime('%d/%m/%Y %H:%M'),
            "agreement_types": list(all_categories)
        }

        with open(SYNC_LOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)

        return log_data
    except Exception as e:
        st.error(f"שגיאה בעת סנכרון קבצים ל-RAG: {e}")
        return load_sync_log()

def load_sync_log():
    if os.path.exists(SYNC_LOG_FILE):
        with open(SYNC_LOG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        "processed_files": [], 
        "last_sync": datetime.now().strftime('%d/%m/%Y %H:%M'), 
        "agreement_types": ["כתב התחייבות לתשתיות", "הסכם אספקת עוגות לתחום GIS", "הסכם שירותים ורכש"]
    }

def get_specific_file_content(file_name, drive_service):
    try:
        files = list_all_files_recursive(drive_service, LEARNING_FOLDER_ID)
        target_file = [f for f in files if f['name'] == file_name]
        
        if not target_file:
            files_out = list_all_files_recursive(drive_service, OUTPUT_FOLDER_ID)
            target_file = [f for f in files_out if f['name'] == file_name]

        if target_file:
            file_id = target_file[0]['id']
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            return extract_text_from_file_bytes(fh.getvalue(), target_file[0].get('mimeType', ''), file_name)
    except Exception as e:
        st.error(f"שגיאה בשליפת קובץ ספציפי: {e}")
    return None

def generate_contract_proposal(mode, selected_category, selected_base_file, party_a, party_b, special_terms, new_dates, new_price, new_scope, collection, client, drive_service):
    
    if mode == "התבססות מוחלטת על הסכם קיים ספציפי" and selected_base_file:
        base_text = get_specific_file_content(selected_base_file, drive_service)
        if not base_text:
            return "שגיאה: לא ניתן היה לשלוף את תוכן המסמך הנבחר מהדרייב.", "", []

        prompt = f"""
אתה מנוע משפטי מתקדם של יחידת מתוו"ה במשרד הביטחון.
תפקידך: להתבסס **באופן מוחלט** על המסמך המקורי הקיים המצורף מטה, לשמור על המבנה, הסעיפים והניסוח המדויק שלו, ולבצע **אך ורק** את השינויים שהוגדרו.

מסמך המקור להתבססות מוחלטת:
=========================================
{base_text}
=========================================

שינויים ועדכונים שנדרש להטמיע במסמך החדש:
- צד א' מעודכן: {party_a}
- צד ב' מעודכן: {party_b}
- תאריכים / לוחות זמנים חדשים: {new_dates if new_dates else "ללא שינוי"}
- מחיר / תמורה מעודכנת: {new_price if new_price else "ללא שינוי"}
- תכולת עבודה / שינויים בתכולה: {new_scope if new_scope else "ללא שינוי"}
- דגשים מיוחדים נוספים: {special_terms if special_terms else "ללא"}

הנחיות קשיחות:
1. שמור על כל סעיפי ההסכם המקורי, הניסוח והסגנון המשפטי שלו ב-100%.
2. עדכן באופן טבעי בלבד את התאריכים, המחיר, שמות הצדדים והתכולה בהתאם לפרטים החדשים.
"""
        sources = [selected_base_file]
        retrieved_context = f"התבססות מלאה על המסמך: {selected_base_file}"

    else:
        search_query = f"{selected_category} {special_terms}"
        results = collection.query(query_texts=[search_query], n_results=15)

        retrieved_context = ""
        sources = set()

        if results and results.get("documents"):
            for docs, metass in zip(results["documents"], results["metadatas"]):
                for doc, meta in zip(docs, metass):
                    retrieved_context += f"\n[מתוך: {meta.get('filename')} | סוג: {meta.get('category_type', 'כללי')}]\n{doc}\n"
                    sources.add(meta.get('filename'))

        prompt = f"""
אתה מנוע משפטי מתקדם של יחידת מתוו"ה במשרד הביטחון.
עליך לנסח הצעה להסכם מסוג: {selected_category}

ההסכם חייב להתבסס **אך ורק** על הסגנון, הטרמינולוגיה, המבנה והסעיפים המשפטיים מתוך קובצי ה-Drive שנלמדו.

פרטי העסקה:
- סוג ההסכם המבוקש: {selected_category}
- צד א': {party_a}
- צד ב': {party_b}
- תנאים מיוחדים/שינויים: {special_terms}

סעיפים שנשלפו מתוך מאגר הידע:
=========================================
{retrieved_context if retrieved_context else "לא נמצאו סעיפים ספציפיים."}
=========================================
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1
    )

    return response.choices[0].message.content, retrieved_context, list(sources)

# ==========================================
# 5. ממשק המשתמש (UI)
# ==========================================
st.markdown("""
    <div class="mod-official-header">
        <div>
            <span class="classification-tag">בלתי מסווג</span>
            <div class="header-title">🏛️ משרד הביטחון - יחידת מתוו"ה</div>
            <p>מערכת בינה מלאכותית ארגונית לניסוח, ניתוח ולמידת הסכמים מול Google Drive המשרדי</p>
        </div>
        <div style="font-size: 3.2rem; opacity: 0.9;">📜</div>
    </div>
""", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📝 ניסוח/התבססות על הסכם", "📤 העלאת קבצים לדרייב", "🔄 סטטוס וסנכרון Drive"])

client = get_openai_client()
collection = get_vector_db()
drive_service = get_drive_service()

if "sync_log" not in st.session_state:
    st.session_state["sync_log"] = load_sync_log()

log_data = st.session_state["sync_log"]
agreement_types = log_data.get("agreement_types", [])
processed_files = log_data.get("processed_files", [])

with tab1:
    st.markdown('<div class="mod-card">', unsafe_allow_html=True)
    st.subheader("🎯 בחירת אופן ניסוח המסמך")
    
    draft_mode = st.radio(
        "בחר מנגנון ניסוח:",
        ["ניסוח משורלב מכלל מאגר הידע (RAG)", "התבססות מוחלטת על הסכם קיים ספציפי"],
        horizontal=True
    )
    
    col1, col2 = st.columns(2)
    
    selected_base_file = None
    new_dates = ""
    new_price = ""
    new_scope = ""

    with col1:
        if draft_mode == "התבססות מוחלטת על הסכם קיים ספציפי":
            if processed_files:
                selected_base_file = st.selectbox("📄 בחר את הסכם המקור להתבססות מוחלטת:", processed_files)
            else:
                selected_base_file = st.text_input("שם הקובץ להתבססות:")
            selected_category = selected_base_file.split('_')[0] if selected_base_file else "הסכם"
        else:
            if agreement_types:
                selected_category = st.selectbox("🎯 בחר סוג הסכם להתבססות:", agreement_types)
            else:
                selected_category = st.text_input("סוג ההסכם", value="הסכם מסגרת")
        
        party_a = st.text_input("שם צד א' (למשל: משרד הביטחון / מזמין)")

    with col2:
        party_b = st.text_input("שם צד ב' (למשל: הקבלן / הספק)")
        if draft_mode == "התבססות מוחלטת על הסכם קיים ספציפי":
            new_dates = st.text_input("📅 תאריכים / לוחות זמנים חדשים:", placeholder="לדוגמא: מ-01/11/2026 עד 31/12/2027")

    if draft_mode == "התבססות מוחלטת על הסכם קיים ספציפי":
        col_c, col_d = st.columns(2)
        with col_c:
            new_price = st.text_input("💰 מחיר / תמורה מעודכנת:", placeholder="לדוגמא: 150,000 ש\"ח כולל מע\"מ")
        with col_d:
            new_scope = st.text_input("📦 שינויים בתכולת העבודה:", placeholder="פרט תכולה חדשה שנוספה או הוסרה...")

    special_terms = st.text_area("✍️ תנאים מיוחדים ודגשים נוספים למסמך:", placeholder="פרט דגשים משפטיים או ערבויות נדרשות...")

    if st.button("🚀 נסח הצעה להסכם"):
        if not party_a or not party_b:
            st.warning("אנא מלא את שמות הצדדים.")
        else:
            with st.spinner("מעבד ומנסח את המסמך מתוך מאגר הידע..."):
                proposal, context_used, sources = generate_contract_proposal(
                    draft_mode, selected_category, selected_base_file, party_a, party_b, special_terms,
                    new_dates, new_price, new_scope, collection, client, drive_service
                )
                
                st.session_state["current_proposal"] = proposal
                st.session_state["edited_proposal"] = proposal
                st.session_state["current_category"] = selected_category
                st.session_state["party_a"] = party_a
                st.session_state["party_b"] = party_b
                st.session_state["show_edit_box"] = False
                
                st.success("המסמך נוסח בהצלחה!")
                if sources:
                    st.info(f"📚 מסמכי המקור ששימשו בסיס לניסוח: {', '.join(sources)}")

    if "current_proposal" in st.session_state:
        st.markdown("<br>### 📄 טיוטת המסמך המוצעת:", unsafe_allow_html=True)
        
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
                label="📄 יצא מסמך טקסט/PDF",
                data=pdf_data,
                file_name=f"{cat}_בין_{pa}_לבין_{pb}.txt",
                mime="text/plain"
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
                        st.success("🎉 המסמך נשמר בהצלחה בתיקיית הפלט ב-Google Drive!")
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
    st.markdown('<div class="mod-card">', unsafe_allow_html=True)
    st.subheader("📤 העלאת קובצי PDF סרוקים ישירות לתיקיית הלמידה ב-Google Drive")
    uploaded_files = st.file_uploader("בחר קובצי PDF להעלאה", type=["pdf"], accept_multiple_files=True)
    if st.button("עלה ל-Drive"):
        st.info("הקבצים יעלו לתיקייה המשרדית ויעובדו בסנכרון הבא.")
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="mod-card">', unsafe_allow_html=True)
    st.subheader("🔄 סטטוס סנכרון ואינדוקס RAG ב-Google Drive")
    
    if st.button("🔄 בצע סנכרון עכשיו מול Google Drive (עדכון מאגר הידע)"):
        with st.spinner("סורק קבצים חדשים (כולל תת-תיקיות), מפעיל חילוץ טקסטים ומעדכן את ChromaDB Vector Store..."):
            drive_svc = get_drive_service()
            updated_log = sync_and_index_rag_drive_files(collection, drive_svc)
            st.session_state["sync_log"] = updated_log
            st.success("🎉 הסנכרון והאינדוקס של מאגר ה-RAG הושלמו בהצלחה!")
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    m1, m2, m3 = st.columns(3)
    
    with m1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{len(processed_files)}</div>
                <div class="metric-label">📄 קבצים שעובדו למאגר הידע</div>
            </div>
        """, unsafe_allow_html=True)

    with m2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{len(agreement_types)}</div>
                <div class="metric-label">🏷️ סוגי הסכמים שנלמדו</div>
            </div>
        """, unsafe_allow_html=True)

    with m3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{log_data.get('last_sync', 'טרם בוצע')}</div>
                <div class="metric-label">📅 תאריך סנכרון אחרון</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.write(f"📂 מזהה תיקיית הלימוד: **`{LEARNING_FOLDER_ID}`**")
    st.write(f"📂 מזהה תיקיית השמירה (פלט): **`{OUTPUT_FOLDER_ID}`**")
    
    if processed_files:
        st.markdown("### 📄 רשימת קובצי המקור המעובדים ב-Vector DB (זמינים להתבססות):")
        for f_name in processed_files:
            st.markdown(f"- 📁 `{f_name}`")

    st.markdown('</div>', unsafe_allow_html=True)
