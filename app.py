import os
import io
import time
import json
from datetime import datetime
import streamlit as st

from openai import OpenAI
import chromadb
from chromadb.utils import embedding_functions

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", ""))
GOOGLE_DRIVE_FOLDER_ID = st.secrets.get("GOOGLE_DRIVE_FOLDER_ID", os.getenv("GOOGLE_DRIVE_FOLDER_ID", "1Q_isBiLXre0PGEFsc7Y9CS3EU8W1DzB2"))
GOOGLE_CREDENTIALS_FILE = "credentials.json"
SCOPES = ['https://www.googleapis.com/auth/drive']
CHROMA_DB_PATH = "./chroma_db"
SYNC_LOG_FILE = "sync_log.json"

st.set_page_config(
    page_title="מערכת ניסוח הסכמים - יחידת מתוו\"ה",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Rubik:wght@300;400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Rubik', sans-serif;
        direction: rtl;
        text-align: right;
        background-color: #f8fafc;
    }
    
    .stApp {
        background-image: linear-gradient(rgba(248, 250, 252, 0.94), rgba(248, 250, 252, 0.94)), 
                          url('https://upload.wikimedia.org/wikipedia/commons/e/e2/Israel_Defense_Ministry_Symbol.png');
        background-repeat: no-repeat;
        background-position: center center;
        background-size: 500px auto;
        background-attachment: fixed;
    }

    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #ffffff;
        padding: 2rem 2.5rem;
        border-radius: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2);
        margin-bottom: 2rem;
        border-bottom: 4px solid #3b82f6;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    .main-header h1 {
        margin: 0;
        font-weight: 700;
        font-size: 2.1rem;
        color: #f8fafc;
    }

    .css-card {
        background: rgba(255, 255, 255, 0.88);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(226, 232, 240, 0.8);
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
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
            return None
        flow = InstalledAppFlow.from_client_secrets_file(GOOGLE_CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def upload_text_as_file_to_drive(file_title, text_content):
    drive_service = get_drive_service()
    if not drive_service:
        return False, "שגיאה בחיבור ל-Google Drive."

    file_metadata = {
        'name': f"{file_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        'parents': [GOOGLE_DRIVE_FOLDER_ID]
    }
    media = MediaIoBaseUpload(io.BytesIO(text_content.encode('utf-8')), mimetype='text/plain')
    uploaded_file = drive_service.files().create(body=file_metadata, mediaBody=media, fields='id').execute()
    return True, uploaded_file.get('id')

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
            selected_category = st.text_input("סוג ההסכם", value="כתב התחייבות לתשתיות")
        
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
                st.session_state["current_category"] = selected_category
                
                st.success("המסמך נוסח בהצלחה!")
                if sources:
                    st.info(f"📚 מסמכי המקור בדרייב המשרדי ששימשו בסיס לניסוח: {', '.join(sources)}")
                
                st.markdown("### 📄 טיוטת המסמך המוצעת:")
                st.text_area("תוכן ההסכם", value=proposal, height=400)

    if "current_proposal" in st.session_state:
        st.markdown("---")
        if st.button("💾 שמור הסכם זה בחזרה ישירות ל-Google Drive המשרדי"):
            with st.spinner("מעלה את המסמך המנוסח לתיקייה בדרייב..."):
                success, file_id = upload_text_as_file_to_drive(
                    st.session_state["current_category"],
                    st.session_state["current_proposal"]
                )
                if success:
                    st.success("🎉 המסמך נשמר בהצלחה בתיקיית ה-Google Drive המשרדית!")
                else:
                    st.error("שגיאה בעת שמירת המסמך לדרייב.")

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
