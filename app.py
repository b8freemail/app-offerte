import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from docx import Document
from docxcompose.composer import Composer
import io

# --- 1. CONFIGURAZIONE PASSWORD E ID ---
PASSWORD_APP = "offerte2026"  # Puoi cambiare la password qui
FOLDER_ID = "1qIIS1byFpfXnQvmAl1gZjYTrAgh2qbiw"  # <--- INCOLLA QUI IL CODICE DELLA TUA CARTELLA DRIVE

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

# Mappa dei nomi dei file agli ID di Google Drive
file_map = {f['name'].replace('.docx', ''): f['id'] for f in files}
file_options = list(file_map.keys())

st.write("✅ **Seleziona i moduli nell'ordine in cui desideri unirli:**")
st.caption("💡 *Puoi trascinare le etichette con il mouse dentro la casella per cambiarne l'ordine finale.*")

selected_names = st.multiselect(
    "Scegli i moduli dall'elenco:",
    options=file_options,
    default=[]
)

if st.button("🚀 Genera Offerta Word", type="primary"):
    if not selected_names:
        st.error("Seleziona almeno un modulo prima di generare!")
    else:
        # Prende gli ID rispettando l'ordine esatto scelto dall'utente
        selected_ids = [file_map[name] for name in selected_names]
        
        with st.spinner("Creazione offerta in corso... (richiede qualche decina di secondi)"):
            try:
                # Usa il primo file selezionato come Master
                master_io = download_file(service, selected_ids[0])
                master_doc = Document(master_io)
                composer = Composer(master_doc)
                
                # Unisce gli altri documenti nell'ordine esatto dell'elenco
                for fid in selected_ids[1:]:
                    master_doc.add_page_break() # Salto pagina tra un modulo e l'altro
                    doc_io = download_file(service, fid)
                    doc_to_append = Document(doc_io)
                    composer.append(doc_to_append)
                
                # Prepara il file finale da scaricare
                output_io = io.BytesIO()
                composer.save(output_io)
                output_io.seek(0)
                
                st.success("🎉 Offerta generata con successo!")
                st.download_button(
                    label="⬇️ SCARICA IL FILE WORD FINALE",
                    data=output_io,
                    file_name="Nuova_Offerta_Composta.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Errore durante l'unione dei file: {e}")
