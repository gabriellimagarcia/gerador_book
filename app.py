# app.py
import re
import hashlib
import base64
from io import BytesIO
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import pandas as pd
import requests
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="Gerador de Book", page_icon="📸", layout="wide")

# ===== Tema (Claro predominante branco, acento LARANJA e detalhes PRETO) =====
def apply_theme(dark: bool):
    ORANGE = "#FF7A00"
    ORANGE_HOVER = "#E66E00"
    BLACK = "#111111"
    GRAY_BG = "#f6f6f7"

    if dark:
        css = f"""
        <style>
        :root {{
            --accent: {ORANGE};
            --accent-hover: {ORANGE_HOVER};
            --text: #f5f5f5;
            --bg: #0e1117;
            --panel: #11151c;
            --muted: #a3a3a3;
        }}
        .stApp {{ background-color: var(--bg); color: var(--text); }}
        section[data-testid="stSidebar"] > div {{
            background: var(--panel);
            border-right: 1px solid #1b212c;
        }}
        h1, h2, h3, h4, h5, h6 {{ color: var(--text); }}
        .stTextInput input, .stNumberInput input {{
            color: var(--text); background: #0f131a; border: 1px solid #232a36;
        }}
        .stTextInput input:focus, .stNumberInput input:focus {{
            outline: none; border: 1px solid var(--accent); box-shadow: 0 0 0 1px var(--accent);
        }}
        .stButton > button, .stDownloadButton > button {{
            background: var(--accent); color: white; border: none; border-radius: 10px;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover {{
            background: {ORANGE_HOVER}; color: white;
        }}
        .stProgress > div > div {{ background-color: var(--accent); }}
        a {{ color: var(--accent); }}
        </style>
        """
    else:
        css = f"""
        <style>
        :root {{
            --accent: {ORANGE};
            --accent-hover: {ORANGE_HOVER};
            --text: {BLACK};
            --bg: #ffffff;
            --panel: {GRAY_BG};
            --muted: #5f6368;
        }}
        .stApp {{ background-color: var(--bg); color: var(--text); }}
        section[data-testid="stSidebar"] > div {{
            background: var(--panel);
            border-right: 1px solid #ececec;
        }}
        h1, h2, h3, h4, h5, h6 {{ color: var(--text); }}
        .stTextInput input, .stNumberInput input {{
            color: var(--text); background: #ffffff; border: 1px solid #dcdcdc;
        }}
        .stTextInput input:focus, .stNumberInput input:focus {{
            outline: none; border: 1px solid var(--accent); box-shadow: 0 0 0 1px var(--accent) inset;
        }}
        .stButton > button, .stDownloadButton > button {{
            background: var(--accent); color: white; border: none; border-radius: 10px;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover {{
            background: {ORANGE_HOVER}; color: white;
        }}
        .stSlider [data-baseweb="slider"] > div > div > div {{ background: rgba(255,122,0,0.2); }}
        .stSlider [data-baseweb="slider"] > div > div > div > div {{ background: var(--accent); }}
        .stProgress > div > div {{ background-color: var(--accent); }}
        a {{ color: var(--accent); }}
        </style>
        """
    st.markdown(css, unsafe_allow_html=True)

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# ===== Login (lista completa restaurada) =====
ALLOWED_USERS = {
    "lucas.costa@mkthouse.com.br": "mudar12345",
    "gabriel.garcia@mkthouse.com.br": "Peter2025!",
    "daniela.scibor@mkthouse.com.br": "mudar12345",
    "regiane.paula@mkthouse.com.br": "mudar12345",
    "pamela.fructuoso@mkthouse.com.br": "mudar12345",
    "fernanda.sabino@mkthouse.com.br": "mudar12345",
    "cacia.nogueira@mkthouse.com.br": "mudar12345",
    "edson.fortaleza@mkthouse.com.br": "mudar12345",
    "lucas.depaula@mkthouse.com.br": "mudar12345",
    "janaina.morais@mkthouse.com.br": "mudar12345",
    "debora.ramos@mkthouse.com.br": "mudar12345",
}
ALLOWED_USERS = {k.strip().lower(): v for k, v in ALLOWED_USERS.items()}

def do_login():
    st.title("🔐 Login")
    with st.form("login_form", clear_on_submit=False):
        email = st.text_input("E-mail", placeholder="seu.email@mkthouse.com.br")
        pwd = st.text_input("Senha", type="password", placeholder="••••••••")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        email_norm = (email or "").strip().lower()
        if email_norm in ALLOWED_USERS and pwd == ALLOWED_USERS[email_norm]:
            st.session_state.auth = True
            st.session_state.user_email = email_norm
            st.rerun()
        else:
            st.error("Credenciais inválidas. Verifique e tente novamente.")

if "auth" not in st.session_state:
    st.session_state.auth = False

# ===== Utilitários =====
URL_RE = re.compile(r'https?://\S+')

def extrair_links(celula):
    if pd.isna(celula): return []
    texto = str(celula).replace(",", " ").replace("(", " ").replace(")", " ").replace('"', " ").replace("'", " ")
    return [u.rstrip(").,") for u in URL_RE.findall(texto)]

def redimensionar(img: Image.Image, max_w: int, max_h: int) -> Image.Image:
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB": img = img.convert("RGB")
    img.thumbnail((max_w, max_h), resample=Image.LANCZOS)
    return img

def comprimir_jpeg_binsearch(img: Image.Image, limite_kb: int) -> BytesIO:
    lo, hi, best = 35, 95, None
    buf = BytesIO(); img.save(buf, "JPEG", quality=75, optimize=True, progressive=True, subsampling=2)
    if buf.tell()/1024 <= limite_kb: buf.seek(0); return buf
    best = buf
    while lo <= hi:
        mid = (lo+hi)//2
        buf = BytesIO(); img.save(buf, "JPEG", quality=mid, optimize=True, progressive=True, subsampling=2)
        if buf.tell()/1024 <= limite_kb: best = buf; lo = mid+1
        else: hi = mid-1
    if best is None:
        best = BytesIO(); img.save(best, "JPEG", quality=35, o
