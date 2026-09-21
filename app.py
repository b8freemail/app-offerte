import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from docx import Document
from docxcompose.composer import Composer
import io

# --- 1. CONFIGURAZIONE PASSWORD E ID ---
PASSWORD_APP = "offerte2026"  # Puoi cambiare la password qui
FOLDER_ID = "1qIIS1byFpfXnQvmAl1gZjYTrAgh2qbiw"  # <--- INCOLLA QUI IL CODICE DELLA CARTELLA DRIVE

# --- 2. SCHERMATA DI LOGIN ---
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

# --- 3. CONNESSIONE A GOOGLE DRIVE ---
st.title("📄 Compositore Offerte Automatico")

@st.cache_resource
def get_drive_service():
    # Prende la chiave segreta (che configureremo su Streamlit a breve)
    creds_json = st.secrets["gcp_service_account"]
    credentials = service_account.Credentials.from_service_account_info(creds_json)
    service = build('drive', 'v3', credentials=credentials)
    return service

def get_files_from_folder(service, folder_id):
    query = f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'"
    results = service.files().list(q=query, fields="nextPageToken, files(id, name)", pageSize=1000).execute()
    items = results.get('files', [])
    items.sort(key=lambda x: x['name'].lower())
    return items

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_io = io.BytesIO()
    downloader = request.execute()
    file_io.write(downloader)
    file_io.seek(0)
    return file_io

service = get_drive_service()

# --- 4. INTERFACCIA E CREAZIONE OFFERTA ---
with st.spinner("Sincronizzazione con Google Drive in corso..."):
    try:
        files = get_files_from_folder(service, FOLDER_ID)
    except Exception as e:
        st.error("Errore di collegamento con Google Drive. Verifica i permessi.")
        st.stop()

if not files:
    st.warning("Nessun file Word trovato nella cartella Google Drive.")
    st.stop()

st.write("✅ **Seleziona i moduli da includere nell'offerta:**")

selections = {}
for f in files:
    nome_pulito = f['name'].replace('.docx', '')
    selections[f['id']] = st.checkbox(nome_pulito, key=f['id'])

if st.button("🚀 Genera Offerta Word", type="primary"):
    selected_ids = [fid for fid, is_selected in selections.items() if is_selected]
    
    if not selected_ids:
        st.error("Seleziona almeno un modulo prima di generare!")
    else:
        with st.spinner("Creazione offerta in corso... (richiede qualche decina di secondi)"):
            try:
                # Usa il primo file come Master
                master_io = download_file(service, selected_ids[0])
                master_doc = Document(master_io)
                composer = Composer(master_doc)
                
                # Unisce gli altri documenti selezionati
                for fid in selected_ids[1:]:
                    master_doc.add_page_break() # Inserisce un salto pagina tra una scheda e l'altra
                    doc_io = download_file(service, fid)
                    doc_to_append = Document(doc_io)
                    composer.append(doc_to_append)
                
                # Prepara il file finale da scaricare
                output_io = io.BytesIO()
                composer.save(output_io)
                output_io.seek(0)
                
                st.success("🎉 Offerta generata con successo! Clicca qui sotto per salvarla sul tuo computer.")
                st.download_button(
                    label="⬇️ SCARICA IL FILE WORD FINALE",
                    data=output_io,
                    file_name="Nuova_Offerta_Composta.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Errore durante l'unione dei file: {e}")
