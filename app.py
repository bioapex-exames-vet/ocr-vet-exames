import streamlit as st
import time
import io
import json
import datetime
import re
import smtplib

from docx import Document
from reportlab.pdfgen import canvas
from email.message import EmailMessage
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

st.set_page_config(layout="wide")
INACTIVITY_LIMIT = 600

marcadores_hemograma = [
    "WBC","LYM%","MON%","GRA%","LYM#","MON#","GRA#",
    "RBC","HGB","HCT","MCV","MCH","MCHC",
    "RDW_CV","RDW_SD","PLT","PCT","MPV","PDW","P_LCR","P_LCC"
]

if "logado" not in st.session_state:
    st.session_state.logado = False
if "last_active" not in st.session_state:
    st.session_state.last_active = time.time()

if time.time() - st.session_state.last_active > INACTIVITY_LIMIT:
    st.session_state.clear()
    st.warning("⏰ Sessão expirada. Recarregue a página.")
    st.stop()

if not st.session_state.logado:
    try:
        st.image("logo_Bioapex.png", use_column_width=True)
    except:
        st.write("Bioapex - Exames Veterinários")

    st.title("🔐 Login")

    with st.form("login"):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")

    if entrar:
        if usuario == st.secrets["USUARIO1"] and senha == st.secrets["SENHA1"]:
            st.session_state.logado = True
            st.session_state.last_active = time.time()
            st.rerun()
        else:
            st.error("Credenciais inválidas")

st.session_state.last_active = time.time()

service_account_info = json.loads(st.secrets["GDRIVE_JSON"])
credentials = Credentials.from_service_account_info(
    service_account_info,
    scopes=['https://www.googleapis.com/auth/drive']
)

drive_service = build('drive', 'v3', credentials=credentials)
PARENT_FOLDER_ID = st.secrets["GDRIVE_FOLDER_ID"]

def extrair_dados(texto):

    dados = {}

    match = re.search(r'Propriet[áa]rio:\s*(.+)', texto, re.I)
    dados["Proprietario"] = match.group(1).strip() if match else ""

    match = re.search(r'(?:Nome\s*de\s*paciente|de paciente):\s*(.+)', texto, re.I)
    dados["Paciente"] = match.group(1).strip() if match else ""

    match = re.search(r'ID da anostra:\s*(\d+)', texto, re.I)
    dados["ID_amostra"] = match.group(1).strip() if match else ""

    match = re.search(r'Esp[ée]cie:\s*(.+)', texto, re.I)
    dados["Especie"] = match.group(1).strip() if match else ""

    texto = re.sub(r'RDW[\s\r\n]+CV', 'RDW_CV', texto, flags=re.I)
    texto = re.sub(r'RDW[\s\r\n]+SD', 'RDW_SD', texto, flags=re.I)
    texto = re.sub(r'P[\s\r\n]+LCR', 'P_LCR', texto, flags=re.I)
    texto = re.sub(r'P[\s\r\n]+LCC', 'P_LCC', texto, flags=re.I)

    tokens = re.split(r'\s+|[\r\n]+', texto)
    hemograma = {}

    for i, token in enumerate(tokens):
        token = token.upper()

        if token in marcadores_hemograma:

            for j in range(i+1, len(tokens)):

                t = re.sub(r'^(L\s*|None\s*)', '', tokens[j])
                match = re.search(r'\d+[.,]?\d*', t)

                if match:
                    hemograma[token] = float(match.group(0).replace(",", "."))
                    break

    dados["hemograma"] = hemograma

    return dados

def salvar_drive(file_bytes, nome, mime):

    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime)

    drive_service.files().create(
        body={'name': nome, 'parents':[PARENT_FOLDER_ID]},
        media_body=media,
        fields='id'
    ).execute()

def preencher_template(nome, numero, texto, data_exame):

    template_id = st.secrets["TEMPLATE_DOCX_ID"]

    fh = io.BytesIO()

    try:
        request = drive_service.files().get_media(fileId=template_id)
    except HttpError as e:
        st.error(f"Erro Google Drive: {e}")
        st.stop()

    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    fh.seek(0)

    doc = Document(fh)

    for p in doc.paragraphs:
        p.text = p.text.replace("{{NOME}}", nome)
        p.text = p.text.replace("{{DOCUMENTO}}", numero)
        p.text = p.text.replace("{{DATA}}", str(data_exame))
        p.text = p.text.replace("{{TEXTO}}", texto)

    nome_doc = f"{nome}_{datetime.datetime.now():%Y%m%d%H%M%S}.docx"

    buffer = io.BytesIO()
    doc.save(buffer)

    salvar_drive(
        buffer.getvalue(),
        nome_doc,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    return nome_doc

def gerar_pdf(texto, nome):

    pdf = io.BytesIO()
    c = canvas.Canvas(pdf)

    y = 800
    for linha in texto.split("\n"):
        c.drawString(30, y, linha)
        y -= 15

    c.save()
    pdf.seek(0)

    salvar_drive(pdf.read(), f"{nome}.pdf", "application/pdf")

def enviar_email(destino, arquivo):

    msg = EmailMessage()
    msg["Subject"] = "Documento Processado"
    msg["From"] = st.secrets["EMAIL"]
    msg["To"] = destino
    msg.set_content("Segue documento em anexo.")

    with open(arquivo, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=arquivo
        )

    with smtplib.SMTP_SSL("smtp.gmail.com",465) as smtp:
        smtp.login(st.secrets["EMAIL"], st.secrets["SENHA_EMAIL"])
        smtp.send_message(msg)

if st.session_state.logado:

    col1,col2 = st.columns([8,1])

    with col2:
        if st.button("🚪 Sair"):
            st.session_state.clear()
            st.rerun()

    try:
        st.image("logo_Bioapex.png", use_column_width=True)
    except:
        pass

    st.title("Bioapex - Exames Veterinários")

    texto = st.text_area("Cole o texto do OCR", height=200)
    email_destino = st.text_input("Enviar PDF para email")

    if texto.strip():

        dados = extrair_dados(texto)

        st.subheader("Dados do paciente")

        col1,col2 = st.columns(2)

        with col1:
            proprietario = st.text_input("Proprietário", dados["Proprietario"])
            paciente = st.text_input("Paciente", dados["Paciente"])

        with col2:
            id_amostra = st.text_input("ID Amostra", dados["ID_amostra"])
            especie = st.text_input("Espécie", dados["Especie"])

        data_exame = st.date_input("Data do exame")

        st.subheader("Hemograma")

        hemograma_editado = {}
        cols = st.columns(3)

        for i, marcador in enumerate(marcadores_hemograma):

            valor = dados["hemograma"].get(marcador,0.0)

            with cols[i%3]:

                if valor == 0:
                    st.markdown(f":red[**{marcador}**]")
                else:
                    st.markdown(f"**{marcador}**")

                hemograma_editado[marcador] = st.number_input(
                    "",
                    value=float(valor),
                    step=0.01,
                    format="%.2f",
                    key=marcador
                )

        if st.button("Gerar Documento"):

            nome_doc = preencher_template(
                paciente,
                id_amostra,
                texto,
                data_exame
            )

            if nome_doc:

                gerar_pdf(texto, nome_doc.replace(".docx",""))

                if email_destino:
                    enviar_email(
                        email_destino,
                        nome_doc.replace(".docx",".pdf")
                    )

                st.success("Documento gerado com sucesso!")
