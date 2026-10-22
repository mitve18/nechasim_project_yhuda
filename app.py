import os
import io
import time
import base64
import json
import fitz
from datetime import datetime
import streamlit as st

from openai import OpenAI
import chromadb
from chromadb.utils import embedding_functions

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
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
        st.error("אנא הגדר את OPENAI_API_KEY במשתני הסביבה.")
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

# --- ממשק משתמש ---
st.markdown("""
    <div class="main-header">
        <div>
            <h1>🏛️ משרד הביטחון - יחידת מתוו"ה</h1>
            <p>מערכת ארגונית לניסוח, ניתוח ולמידת הסכמים</p>
        </div>
        <div style="font-size: 2.5rem;">📜</div>
    </div>
""", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📝 ניסוח הסכם לפי סוג", "📤 העלאת קבצים", "🔄 ניהול וסנכרון"])

client = get_openai_client()
collection = get_vector_db()
log_data = load_sync_log()
agreement_types = log_data.get("agreement_types", [])

with tab1:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("📝 בחירת סוג הסכם לניסוח מתוך ידע המערכת")
    
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
            with st.spinner("שולף סעיפים תואמים מתוך סוג ההסכם הנבחר ומנסח מסמך..."):
                proposal, context_used, sources = generate_strict_contract_proposal(
                    selected_category, party_a, party_b, special_terms, collection, client
                )
                
                st.success("המסמך נוסח בהצלחה!")
                if sources:
                    st.info(f"📚 מסמכי המקור בדרייב ששימשו בסיס לניסוח: {', '.join(sources)}")
                
                st.markdown("### 📄 טיוטת המסמך המוצעת:")
                st.text_area("תוכן ההסכם", value=proposal, height=450)

                with st.expander("🔍 הצג את הקטעים המשפטיים שנשלפו (Context)"):
                    st.text(context_used)
    st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("📤 העלאת קובצי PDF סרוקים ל-Google Drive")
    st.info("העלה קבצים חדשים לתיקיית הדרייב והרץ סנכרון כדי שהמערכת תסווג אותם אוטומטית.")
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="css-card">', unsafe_allow_html=True)
    st.subheader("🔄 סטטוס סוגי ההסכמים שנלמדו")
    st.write(f"📅 סנכרון אחרון: **{log_data.get('last_sync', 'טרם בוצע')}**")
    st.write(f"📂 סך הכל קבצים שעובדו: **{len(log_data.get('processed_files', []))}**")
    
    if agreement_types:
        st.markdown("### 🏷️ סוגי ההסכמים הזמינים במערכת:")
        for t in agreement_types:
            st.markdown(f"- **`{t}`**")
    st.markdown('</div>', unsafe_allow_html=True)