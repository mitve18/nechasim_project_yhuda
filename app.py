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

{text_
