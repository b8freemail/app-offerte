import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

PASSWORD_APP = "offerta2026"
FOLDER_ID = "1wp0Vz2jeKf8_N2If9injU5BbwObGjWYD"  # <--- Inserisci il tuo ID

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.markdown("### 🔒 Accesso Riservato")
        pwd = st.text_input("Inserisci la password aziendale:", type="password")
        if st.button("Accedi"):
            if pwd == PASSWORD_APP:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("😕 Password errata")
        return False
    return True

if not check_password():
    st.stop()

st.title("🔍 Diagnosi Collegamento Google Drive")

try:
    creds_json = st.secrets["gcp_service_account"]
    robot_email = creds_json.get("client_email", "Email non trovata nei secrets")
    st.write(f"🤖 **Email dell'account di servizio:** `{robot_email}`")
    st.write(f"📁 **ID Cartella impostato:** `{FOLDER_ID}`")
    
    credentials = service_account.Credentials.from_service_account_info(creds_json)
    service = build('drive', 'v3', credentials=credentials)
    
    # 1. Test accesso alla cartella principale
    try:
        folder_info = service.files().get(fileId=FOLDER_ID, fields="id, name").execute()
        st.success(f"✅ La cartella principale è stata trovata! Nome: **{folder_info.get('name')}**")
    except Exception as e:
        st.error(f"❌ Impossibile accedere alla cartella principale. Errore: {e}")
        st.warning("⚠️ Verificare di aver condiviso la cartella su Google Drive con l'email del robot indicata sopra!")
        st.stop()
        
    # 2. Test scansione elementi contenuti
    q = f"'{FOLDER_ID}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id, name, mimeType)", pageSize=100).execute()
    items = res.get('files', [])
    
    if not items:
        st.warning("⚠️ La cartella principale è visibile, ma risulta vuota per il robot.")
    else:
        st.write(f"📋 **Elementi trovati nella cartella principale ({len(items)}):**")
        for item in items:
            st.write(f"- **{item['name']}** (Tipo: `{item['mimeType']}`)")

except Exception as e:
    st.error(f"Errore generale: {e}")
