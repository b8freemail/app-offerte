import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from docx import Document
from docxcompose.composer import Composer
import io

# --- 1. CONFIGURAZIONE PASSWORD E ID ---
PASSWORD_APP = "offerta2026"  # Password di accesso
FOLDER_ID = "1wp0Vz2jeKf8_N2If9injU5BbwObGjWYD"  # <--- Incolla l'ID della cartella principale Schede_Word

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
    return build('drive', 'v3', credentials=credentials)

def get_subfolders_and_files(service, main_folder_id):
    # Recupera le sotto-cartelle (es. Presentazioni, Schede servizio, Prezzi)
    q_folders = f"'{main_folder_id}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'"
    res_folders = service.files().list(q=q_folders, fields="files(id, name)", pageSize=100).execute()
    subfolders = res_folders.get('files', [])
    subfolders.sort(key=lambda x: x['name'].lower())

    all_files = []

    # 1. Scansione dei file Word nelle sotto-cartelle
    for sf in subfolders:
        sf_id = sf['id']
        sf_name = sf['name']
        q_files = f"'{sf_id}' in parents and trashed=false and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'"
        res_files = service.files().list(q=q_files, fields="files(id, name)", pageSize=1000).execute()
        files = res_files.get('files', [])
        files.sort(key=lambda x: x['name'].lower())
        for f in files:
            nome_pulito = f['name'].replace('.docx', '')
            display_name = f"📁 [{sf_name}] {nome_pulito}"
            all_files.append((f['id'], display_name))

    # 2. Scansione di eventuali file messi direttamente nella radice
    q_root = f"'{main_folder_id}' in parents and trashed=false and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'"
    res_root = service.files().list(q=q_root, fields="files(id, name)", pageSize=1000).execute()
    root_files = res_root.get('files', [])
    root_files.sort(key=lambda x: x['name'].lower())
    for f in root_files:
        nome_pulito = f['name'].replace('.docx', '')
        display_name = f"📄 {nome_pulito}"
        all_files.append((f['id'], display_name))

    return all_files

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_io = io.BytesIO()
    downloader = request.execute()
    file_io.write(downloader)
    file_io.seek(0)
    return file_io

service = get_drive_service()

# --- 4. INTERFACCIA E CREAZIONE OFFERTA ---
with st.spinner("Scansione sotto-cartelle su Google Drive in corso..."):
    try:
        files_list = get_subfolders_and_files(service, FOLDER_ID)
    except Exception as e:
        st.error("Errore di collegamento con Google Drive. Verifica i permessi.")
        st.stop()

if not files_list:
    st.warning("Nessun file Word trovato nella cartella principale o nelle sotto-cartelle.")
    st.stop()

file_map = {display_name: fid for fid, display_name in files_list}
file_options = list(file_map.keys())

st.write("✅ **Seleziona le schede suddivise per tipologia:**")
st.caption("💡 *Le schede indicano la categoria `📁 [Nome Cartella]`. L'ordine di unione finale seguirà esattamente la sequenza con cui le selezioni.*")

selected_names = st.multiselect(
    "Scegli i moduli nell'ordine desiderato:",
    options=file_options,
    default=[]
)

if st.button("🚀 Genera Offerta Word", type="primary"):
    if not selected_names:
        st.error("Seleziona almeno un modulo prima di generare!")
    else:
        selected_ids = [file_map[name] for name in selected_names]
        
        with st.spinner("Creazione offerta in corso..."):
            try:
                master_io = download_file(service, selected_ids[0])
                master_doc = Document(master_io)
                composer = Composer(master_doc)
                
                for fid in selected_ids[1:]:
                    master_doc.add_page_break()
                    doc_io = download_file(service, fid)
                    doc_to_append = Document(doc_io)
                    composer.append(doc_to_append)
                
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
