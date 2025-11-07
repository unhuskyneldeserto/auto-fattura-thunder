import streamlit as st
import pandas as pd
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
import re

# ==========================
# CONFIGURAZIONE INIZIALE
# ==========================

# Inserisci qui gli ID dei tuoi documenti e cartella
TEMPLATE_RICEVUTO_ID = "18aubXYaQzS1J7bI9NPzAZxHZNbiynQz-"
TEMPLATE_VERSATO_ID = "1-8qsZmw-lhK04evJb-5X9UnPcAbaglq9"
CARTELLA_DRIVE_ID = "19hjBzK9dDx7A9D8fEUSDEsUPW_1O09Js"

# ID del foglio Google con la rubrica (modifica qui!)
SHEET_ID = "https://docs.google.com/spreadsheets/d/1dYPtQg-C3HlZzAMF4IpN0JFCiWeZ1byp973LV2XfraM/edit?usp=sharing"
SHEET_NAME = "Foglio1"

# ==========================
# AUTENTICAZIONE GOOGLE
# ==========================

@st.cache_resource
def get_gdrive_services():
    creds = service_account.Credentials.from_service_account_file(
        "credentials.json",
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/documents",
            "https://www.googleapis.com/auth/spreadsheets.readonly"
        ],
    )
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return docs_service, drive_service, sheets_service

docs_service, drive_service, sheets_service = get_gdrive_services()

# ==========================
# LETTURA RUBRICA
# ==========================

@st.cache_data
def carica_rubrica():
    range_name = f"{SHEET_NAME}!A:Z"
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=range_name
    ).execute()
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()
    header, *rows = values
    return pd.DataFrame(rows, columns=header)

rubrica = carica_rubrica()

# ==========================
# INTERFACCIA WEB
# ==========================

st.title("📄 Generatore Ricevute / Dichiarazioni")

# Ricerca contatto
if not rubrica.empty:
    query = st.text_input("🔍 Cerca per nome o email:")
    risultati = rubrica[rubrica.apply(lambda r: query.lower() in str(r.values).lower(), axis=1)] if query else pd.DataFrame()

    if not risultati.empty:
        opzione = st.selectbox("Seleziona la persona:", risultati["Nome"].tolist())
        persona = risultati[risultati["Nome"] == opzione].iloc[0]
    else:
        st.info("Digita almeno 3 lettere per cercare un nominativo.")
        persona = None
else:
    st.error("Rubrica vuota o non trovata.")
    persona = None

if persona is not None:
    st.subheader("📋 Dati del documento")
    
    tipo_doc = st.radio("Tipo di documento:", ["Ricevuta di contributo ricevuto", "Dichiarazione di contributo versato"])
    importo = st.number_input("Importo (€):", min_value=0.0, step=10.0)
    luogo = st.text_input("Luogo dell'evento:")
    evento = st.text_input("Evento (opzionale):")
    numero_ricevuta = st.text_input("Numero ricevuta/fattura:")
    data_evento = st.date_input("Data dell'evento:", datetime.date.today())
    data_ricevuta = st.date_input("Data della ricevuta:", datetime.date.today())

    if st.button("📄 Genera documento"):
        st.info("Generazione in corso...")

        # Scelta template
        template_id = TEMPLATE_RICEVUTO_ID if tipo_doc == "Ricevuta di contributo ricevuto" else TEMPLATE_VERSATO_ID

        # Duplica il template
        copia = drive_service.files().copy(
            fileId=template_id,
            body={
                "name": f"{tipo_doc.replace(' ', '_')}_{persona['Nome']}_{data_ricevuta}",
                "parents": [CARTELLA_DRIVE_ID],
            },
        ).execute()
        nuovo_doc_id = copia["id"]

        # Campi da sostituire
        sostituzioni = {
            "{NOME}": persona.get("Nome", ""),
            "{INDIRIZZO}": persona.get("Indirizzo", ""),
            "{CF}": persona.get("Codice Fiscale", ""),
            "{EMAIL}": persona.get("Email", ""),
            "{PEC}": persona.get("PEC", ""),
            "{IMPORTO}": f"{importo:.2f}",
            "{EVENTO}": evento,
            "{LUOGO}": luogo,
            "{DATA_EVENTO}": data_evento.strftime("%d/%m/%Y"),
            "{DATA_RICEVUTA}": data_ricevuta.strftime("%d/%m/%Y"),
            "{NUMERO}": numero_ricevuta,
        }

        # Aggiorna testo nel documento
        requests = [
            {"replaceAllText": {"containsText": {"text": k, "matchCase": True}, "replaceText": v}}
            for k, v in sostituzioni.items()
        ]
        docs_service.documents().batchUpdate(documentId=nuovo_doc_id, body={"requests": requests}).execute()

        link = f"https://docs.google.com/document/d/{nuovo_doc_id}/edit"
        st.success(f"✅ Documento creato con successo: [Apri il documento]({link})")
