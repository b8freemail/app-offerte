import io
from docx import Document
from docxcompose.composer import Composer
from google.oauth2 import service_account
from googleapiclient.discovery import build
import streamlit as st

# --- 1. CONFIGURAZIONE PASSWORD E ID ---
PASSWORD_APP = "offerta2026"
FOLDER_ID = "1wp0Vz2jeKf8_N2If9injU5BbwObGjWYD"

# --- STILE PERSONALIZZATO PER AUMENTARE LA LEGGIBILITÀ ---
st.set_page_config(page_title="Compositore Offerte", layout="wide")
st.markdown(
    """
    <style>
    /* Rende le etichette del multiselect più ampie e leggibili */
    span[data-baseweb="tag"] {
        max-width: 100% !important;
        white-space: normal !important;
        height: auto !important;
        padding: 6px 10px !important;
        margin: 3px !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# --- 2. SCHERMATA DI LOGIN (SUPPORTO TASTO INVIO) ---
def check_password():
  if "password_correct" not in st.session_state:
    st.session_state["password_correct"] = False

  if not st.session_state["password_correct"]:
    st.markdown("### 🔒 Accesso Riservato")

    # L'uso di st.form abilita l'invio sia tramite click che tramite il tasto INVIO
    with st.form("login_form"):
      pwd = st.text_input("Inserisci la password aziendale:", type="password")
      submit_button = st.form_submit_button("Accedi")

      if submit_button:
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
  credentials = service_account.Credentials.from_service_account_info(
      creds_json
  )
  return build("drive", "v3", credentials=credentials)


def get_subfolders_and_files(service, main_folder_id):
  all_files = []

  q_main = f"'{main_folder_id}' in parents and trashed=false"
  res_main = (
      service.files()
      .list(q=q_main, fields="files(id, name, mimeType)", pageSize=1000)
      .execute()
  )
  items_main = res_main.get("files", [])

  subfolders = [
      item
      for item in items_main
      if item["mimeType"] == "application/vnd.google-apps.folder"
  ]
  subfolders.sort(key=lambda x: x["name"].lower())

  root_files = [
      item
      for item in items_main
      if item["mimeType"] != "application/vnd.google-apps.folder"
  ]
  root_files.sort(key=lambda x: x["name"].lower())

  # 1. File dentro le sotto-cartelle
  for sf in subfolders:
    sf_id = sf["id"]
    sf_name = sf["name"]
    q_sub = f"'{sf_id}' in parents and trashed=false"
    res_sub = (
        service.files()
        .list(q=q_sub, fields="files(id, name, mimeType)", pageSize=1000)
        .execute()
    )
    sub_items = res_sub.get("files", [])
    sub_items.sort(key=lambda x: x["name"].lower())

    for f in sub_items:
      if f["mimeType"] != "application/vnd.google-apps.folder":
        nome_pulito = f["name"].replace(".docx", "").replace(".DOCX", "")
        display_name = f"📁 [{sf_name}]  ➔  {nome_pulito}"
        all_files.append((f["id"], display_name))

  # 2. File nella cartella radice
  for f in root_files:
    nome_pulito = f["name"].replace(".docx", "").replace(".DOCX", "")
    display_name = f"📄 {nome_pulito}"
    all_files.append((f["id"], display_name))

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
with st.spinner("Scansione cartelle su Google Drive in corso..."):
  try:
    files_list = get_subfolders_and_files(service, FOLDER_ID)
  except Exception as e:
    st.error(f"Errore di collegamento con Google Drive: {e}")
    st.stop()

if not files_list:
  st.warning("Nessun file trovato nelle sotto-cartelle.")
  st.stop()

file_map = {display_name: fid for fid, display_name in files_list}
file_options = list(file_map.keys())

st.write("### 1️⃣ Seleziona le schede da includere:")

selected_names = st.multiselect(
    "Scegli i moduli dall'elenco:",
    options=file_options,
    default=[],
    help="Clicca per aggiungere i moduli nell'ordine desiderato.",
)

# --- ANTEPRIMA ORDINATA PER RIGA ---
if selected_names:
  st.markdown("---")
  st.write("### 2️⃣ Sequenza di unione dell'offerta (Un file per riga):")

  for idx, name in enumerate(selected_names, start=1):
    st.info(f"**Posizione {idx}:** {name}")

  st.markdown("---")

  if st.button("🚀 Genera Offerta Word", type="primary", use_container_width=True):
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
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            use_container_width=True,
        )
      except Exception as e:
        st.error(f"Errore durante l'unione dei file: {e}")
