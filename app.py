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

    /* תיקון צבע הכותרת הראשית ללבן בוהק */
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
    """סריקה עמוקה של קבצים בתיקייה ובתת-תיקיות"""
    all_files = []
    query = f"'{folder_id}' in parents and trashed = false"
    results = drive_service.files().list(q=query, pageSize=100, fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])

    for item in items:
        if item['mimeType'] == 'application/vnd.google-apps.folder':
            all_files.extend(list_all_files_recursive(drive_service, item['id']))
        else:
            all_files.append(item)

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
אתה מנוע משפטי מתקדם של יחידת מתוו"ה במשרד הביטח
