import os
import io
import time
import base64
import json
import fitz  # PyMuPDF
from datetime import datetime

from openai import OpenAI
import chromadb
from chromadb.utils import embedding_functions

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "1HS1MqHNvRyHeRuLZWWQQ8ClSBFgoRZQs")
GOOGLE_CREDENTIALS_FILE = "credentials.json"
SCOPES = ['https://www.googleapis.com/auth/drive']
CHROMA_DB_PATH = "./chroma_db"
SYNC_LOG_FILE = "sync_log.json"

def get_drive_service():
    if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
        print(f"❌ שגיאה: הקובץ '{GOOGLE_CREDENTIALS_FILE}' לא נמצא בתיקייה!")
        return None
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(GOOGLE_CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def run_sync():
    print("="*60)
    print("🚀 מתחיל סנכרון ישיר מול Google Drive עם סיווג קטגוריות...")
    print("="*60)

    if not OPENAI_API_KEY:
        print("❌ שגיאה: OPENAI_API_KEY אינו מוגדר.")
        return

    drive_service = get_drive_service()
    if not drive_service:
        return

    openai_client = OpenAI(api_key=OPENAI_API_KEY)
    chroma_client = chromadb.PersistentClient(path=CHROMA_DB_PATH)
    
    openai_ef = embedding_functions.OpenAIEmbeddingFunction(
        api_key=OPENAI_API_KEY,
        model_name="text-embedding-3-small"
    )
    collection = chroma_client.get_or_create_collection(
        name="legal_agreements",
        embedding_function=openai_ef
    )

    log_data = {"processed_files": [], "last_sync": None, "agreement_types": []}
    if os.path.exists(SYNC_LOG_FILE):
        with open(SYNC_LOG_FILE, 'r', encoding='utf-8') as f:
            log_data = json.load(f)

    processed_files = set(log_data.get("processed_files", []))
    agreement_types = set(log_data.get("agreement_types", []))

    query = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and mimeType='application/pdf' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    print(f"\n🔍 נמצאו {len(files)} קובצי PDF בתיקייה.")

    for file in files:
        file_id = file['id']
        file_name = file['name']

        if file_id in processed_files:
            print(f"⏭️  '{file_name}' כבר עובד בעבר.")
            continue

        print(f"\n📥 מוריד ומעבד קובץ: {file_name}...")
        request = drive_service.files().get_media(fileId=file_id)
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        doc = fitz.open(stream=file_stream.getvalue(), filetype="pdf")
        extracted_text = ""
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150)
            base64_image = base64.b64encode(pix.tobytes("png")).decode('utf-8')
            resp = openai_client.chat.completions.create(
                model="gpt-4o",
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"חלץ את הטקסט המלא מעמוד {i+1} בהסכם:"},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                    ]
                }],
                max_tokens=3000
            )
            extracted_text += f"\n--- עמוד {i+1} ---\n" + resp.choices[0].message.content

        # סיווג אוטומטי לסוג הסכם כללי
        type_resp = openai_client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": f"על סמך הטקסט הבא, הגדר ב-2 עד 4 מילים בלבד את **סוג ההסכם / הקטגוריה** (למשל: 'כתב התחייבות תשתיות', 'הסכם קידוחים וביצוע', 'הסכם שכירות מסחרי'). החזר רק את שם הקטגוריה:\n\n{extracted_text[:1500]}"
            }]
        )
        cat_type = type_resp.choices[0].message.content.strip().replace('"', '')
        agreement_types.add(cat_type)

        # שמירה ב-Vector DB עם metadata של הקטגוריה
        paragraphs = [p.strip() for p in extracted_text.split("\n\n") if len(p.strip()) > 20]
        docs, metas, ids = [], [], []
        for idx, para in enumerate(paragraphs):
            docs.append(para)
            metas.append({
                "filename": file_name,
                "category_type": cat_type,
                "chunk_index": idx,
                "indexed_at": str(datetime.now())
            })
            ids.append(f"{file_name}_chunk_{idx}_{int(time.time())}")

        if docs:
            collection.add(documents=docs, metadatas=metas, ids=ids)

        processed_files.add(file_id)
        print(f"✅ '{file_name}' סווג בהצלחה תחת הקטגוריה: [{cat_type}]")

    log_data["processed_files"] = list(processed_files)
    log_data["agreement_types"] = list(agreement_types)
    log_data["last_sync"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(SYNC_LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(log_data, f, ensure_ascii=False, indent=2)

    print("\n🎉 הסנכרון והסיווג הושלמו בהצלחה!")

if __name__ == "__main__":
    run_sync()