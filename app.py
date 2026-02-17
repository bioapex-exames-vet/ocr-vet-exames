import streamlit as st
from PIL import Image
import time
import easyocr
from docx import Document
from reportlab.pdfgen import canvas
import io
import json
import smtplib
from email.message import EmailMessage
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import datetime
import re

# =======================
# CONFIGURAÇÕES GERAIS
# =======================

st.set_page_config(layout="wide")
INACTIVITY_LIMIT = 30

# =======================
# SESSION STATE
# =======================
if "logado" not in st.session_state:
    st.session_state["logado"] = False
if "last_active" not in st.session_state:
    st.session_state["last_active"] = time.time()

# =======================
# TIMEOUT DE SESSÃO
# =======================
now = time.time()
if st.session_state.get("last_active") and (now - st.session_state["last_active"] > INACTIVITY_LIMIT):
    st.session_state.clear()
    st.warning("⏰ Sessão expirada. Recarregue a página e faça login novamente.")
    st.stop()

# =======================
# LOGIN
# =======================
if not st.session_state.get("logado"):
    try:
        logo = Image.open("logo_Bioapex.png")
        st.image(logo, use_column_width=True)
    except:
        st.write("🔹 Bioapex - Exames Veterinários")
    st.title("🔐 Login")
    with st.form(key="login_form"):
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            submit_btn = st.form_submit_button("Entrar")
    if submit_btn:
        if usuario == st.secrets["USUARIO1"] and senha == st.secrets["SENHA1"]:
            st.session_state["logado"] = True
            st.session_state["last_active"] = time.time()
            st.rerun()
        else:
            st.error("Credenciais inválidas")

# =======================
# ATUALIZA TEMPO DE ATIVIDADE
# =======================
st.session_state["last_active"] = time.time()

# =======================
# Configura Google Drive
# =======================
service_account_info = json.loads(st.secrets["GDRIVE_JSON"])
SCOPES = ['https://www.googleapis.com/auth/drive']
credentials = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)
PARENT_FOLDER_ID = st.secrets["GDRIVE_FOLDER_ID"]

# =======================
# Funções
# =======================
reader = easyocr.Reader(['pt'])

def realizar_ocr(imagem):
    img = Image.open(imagem).convert('RGB')
    result = reader.readtext(img)
    texto = "\n".join([t[1] for t in result])
    return texto

def extrair_dados(texto):
    cpf = re.findall(r'\d{3}\.\d{3}\.\d{3}-\d{2}', texto)
    data = re.findall(r'\d{2}/\d{2}/\d{4}', texto)
    return {
        "cpf": cpf[0] if cpf else "",
        "data": data[0] if data else ""
    }

def salvar_no_drive(file_bytes, nome_arquivo, mime_type):
    file_metadata = {'name': nome_arquivo, 'parents': [PARENT_FOLDER_ID]}
    media = io.BytesIO(file_bytes)
    from googleapiclient.http import MediaIoBaseUpload
    media_upload = MediaIoBaseUpload(media, mimetype=mime_type)
    drive_service.files().create(body=file_metadata, media_body=media_upload, fields='id').execute()

def preencher_template(nome, numero, texto, dados):
    # Baixa template do Drive
    results = drive_service.files().list(q=f"'{PARENT_FOLDER_ID}' in parents and name='modelo_padrao.docx'",
                                         fields="files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        st.error("Template não encontrado")
        return None
    template_id = items[0]['id']

    from googleapiclient.http import MediaIoBaseDownload
    fh = io.BytesIO()
    request = drive_service.files().get_media(fileId=template_id)
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)

    doc = Document(fh)
    for p in doc.paragraphs:
        p.text = p.text.replace("{{NOME}}", nome)
        p.text = p.text.replace("{{DOCUMENTO}}", numero)
        p.text = p.text.replace("{{CPF}}", dados["cpf"])
        p.text = p.text.replace("{{DATA}}", dados["data"])
        p.text = p.text.replace("{{TEXTO}}", texto)

    nome_arquivo = f"{nome}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    salvar_no_drive(doc_bytes.read(), nome_arquivo,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    return nome_arquivo

def gerar_pdf(texto, nome_base):
    pdf_bytes = io.BytesIO()
    c = canvas.Canvas(pdf_bytes)
    y = 800
    for linha in texto.split("\n"):
        c.drawString(30, y, linha)
        y -= 15
    c.save()
    pdf_bytes.seek(0)
    salvar_no_drive(pdf_bytes.read(), f"{nome_base}.pdf", "application/pdf")
    return f"{nome_base}.pdf"

def enviar_email(destino, anexo):
    msg = EmailMessage()
    msg["Subject"] = "Documento Processado"
    msg["From"] = st.secrets["EMAIL"]
    msg["To"] = destino
    msg.set_content("Segue documento em anexo.")
    with open(anexo, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=anexo)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(st.secrets["EMAIL"], st.secrets["SENHA_EMAIL"])
        smtp.send_message(msg)

# =======================
# Interface
# =======================
if st.session_state["logado"]:
    col1, col2 = st.columns([8,1])
    with col2:
        if st.button("🚪 Sair"):
            st.session_state.clear()
            st.rerun()
    st.session_state["last_active"] = time.time()
    logo = Image.open("logo_Bioapex.png")
    st.image(logo, use_column_width=True)
    st.title("Bioapex - Exames Veterinários")
    imagem = st.file_uploader("Envie a imagem do exame", type=["jpg","png","jpeg"])
    nome = st.text_input("Nome do paciente")
    tutor = st.text_input("Nome do tutor")
    data = st.date_input("Data do exame")
    email_destino = st.text_input("Enviar para email")

    if st.button("Processar"):
        if imagem:
            texto = realizar_ocr(imagem)
            dados = extrair_dados(texto)
            nome_docx = preencher_template(nome, numero, texto, dados)
            gerar_pdf(texto, nome_docx.replace(".docx",""))
            enviar_email(email_destino, nome_docx.replace(".docx",".pdf"))
            st.success("Processamento concluído! DOCX e PDF salvos no Drive.")
