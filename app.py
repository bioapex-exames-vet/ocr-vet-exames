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
import numpy as np
import cv2

# =======================
# CONFIGURAÇÕES GERAIS
# =======================

st.set_page_config(layout="wide")
INACTIVITY_LIMIT = 600

marcadores_hemograma = [
    "WBC", "LYM%", "MON%", "GRA%", "LYM#", "MON#", "GRA#",
    "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC",
    "RDW_CV", "RDW_SD", "PLT", "PCT", "MPV", "PDW", "P_LCR", "P_LCC"
]

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

def realizar_ocr(image_bytes):
    file_bytes = np.asarray(bytearray(image_bytes), dtype=np.uint8)
    img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError("Imagem inválida ou corrompida.")
    result = reader.readtext(img)
    texto = "\n".join([t[1] for t in result])
    return texto

def extrair_dados(texto, marcadores_hemograma):
    dados = {}
    match = re.search(r'Propriet[áa]rio:\s*(.+)', texto, flags=re.IGNORECASE)
    dados["Proprietario"] = match.group(1).strip() if match else None

    match = re.search(r'(?:Nome\s*de\s*paciente|de paciente):\s*(.+)', texto, flags=re.IGNORECASE)
    dados["Paciente"] = match.group(1).strip() if match else None

    match = re.search(r'ID da anostra:\s*(\d+)', texto, flags=re.IGNORECASE)
    dados["ID_amostra"] = match.group(1).strip() if match else None

    match = re.search(r'Esp[ée]cie:\s*(.+)', texto, flags=re.IGNORECASE)
    dados["Especie"] = match.group(1).strip() if match else None

    match = re.search(r'Hora:\s*([\d\-\.: ]+)', texto, flags=re.IGNORECASE)
    dados["Hora"] = match.group(1).strip() if match else None

    texto = re.sub(r'RDW[\s\r\n]+CV', 'RDW_CV', texto, flags=re.IGNORECASE)
    texto = re.sub(r'RDW[\s\r\n]+SD', 'RDW_SD', texto, flags=re.IGNORECASE)
    texto = re.sub(r'P[\s\r\n]+LCR', 'P_LCR', texto, flags=re.IGNORECASE)
    texto = re.sub(r'P[\s\r\n]+LCC', 'P_LCC', texto, flags=re.IGNORECASE)

    tokens = re.split(r'\s+|[\r\n]+', texto)
    hemograma = {}
    i = 0
    while i < len(tokens):
        token = tokens[i].strip().upper()
        if token in marcadores_hemograma:
            marcador = token
            valor = None
            j = i + 1
            while j < len(tokens):
                t = tokens[j].strip()
                t = re.sub(r'^(L\s*|None\s*)', '', t)
                match = re.search(r'\d+[.,]?\d*', t)
                if match:
                    valor = float(match.group(0).replace(',', '.'))
                    break
                j += 1
            if valor is not None:
                hemograma[marcador] = valor
        i += 1
    dados["hemograma"] = hemograma
    return dados

def salvar_no_drive(file_bytes, nome_arquivo, mime_type):
    file_metadata = {'name': nome_arquivo, 'parents': [PARENT_FOLDER_ID]}
    media = io.BytesIO(file_bytes)
    from googleapiclient.http import MediaIoBaseUpload
    media_upload = MediaIoBaseUpload(media, mimetype=mime_type)
    drive_service.files().create(body=file_metadata, media_body=media_upload, fields='id').execute()

def preencher_template(nome, numero, texto, dados, data_exame):
    results = drive_service.files().list(
        q=f"'{PARENT_FOLDER_ID}' in parents and name='modelo_padrao.docx'",
        fields="files(id, name)"
    ).execute()
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
        p.text = p.text.replace("{{DATA}}", str(data_exame))
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
    try:
        logo = Image.open("logo_Bioapex.png")
        st.image(logo, use_column_width=True)
    except:
        st.write("🔹 Bioapex - Exames Veterinários")
    st.title("Bioapex - Exames Veterinários")

    imagem = st.file_uploader("Envie a imagem do exame", type=["jpg","png","jpeg"])
    nome = st.text_input("Nome do paciente")
    tutor = st.text_input("Nome do tutor")
    data_exame = st.date_input("Data do exame")
    email_destino = st.text_input("Enviar para email")

    # ==========================
    # VALIDAÇÃO SEGURA DE IMAGEM
    # ==========================
    if imagem is not None:
        try:
            img_pil = Image.open(imagem)
            img_pil.verify()  # verifica integridade
            imagem.seek(0)
            image_bytes = imagem.read()
            if len(image_bytes) == 0:
                st.error("Arquivo inválido ou vazio.")
                st.stop()
            st.session_state["image_bytes"] = image_bytes
        except Exception as e:
            st.error(f"Arquivo inválido ou corrompido: {e}")
            st.stop()

    # ==========================
    # PROCESSA SE JÁ EXISTE IMAGEM
    # ==========================
    if "image_bytes" in st.session_state:
        try:
            texto = realizar_ocr(st.session_state["image_bytes"])
            dados = extrair_dados(texto, marcadores_hemograma)
        except Exception as e:
            st.error(f"Erro ao processar imagem: {e}")
            st.stop()

        st.subheader("📋 Conferência e edição do Hemograma")
        hemograma_editado = {}
        cols = st.columns(3)
        for i, marcador in enumerate(marcadores_hemograma):
            valor_ocr = dados["hemograma"].get(marcador)
            with cols[i % 3]:
                hemograma_editado[marcador] = st.number_input(
                    marcador,
                    value=float(valor_ocr) if valor_ocr is not None else 0.0,
                    step=0.01,
                    key=f"input_{marcador}"
                )

        if st.button("Processar e Gerar Documento"):
            dados["hemograma"] = hemograma_editado
            numero = dados.get("ID_amostra", "")
            nome_docx = preencher_template(nome, numero, texto, dados, data_exame)
            if nome_docx:
                gerar_pdf(texto, nome_docx.replace(".docx",""))
                enviar_email(email_destino, nome_docx.replace(".docx",".pdf"))
                st.success("Processamento concluído! DOCX e PDF salvos no Drive.")
