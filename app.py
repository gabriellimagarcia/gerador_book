# === PARTE 1/10 =====================================================
# Boot, imports, logging e estilos iniciais

import re
import os
import gc
import base64
import hashlib
import time
from io import BytesIO
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
import zipfile
import numpy as np
import logging, sys
import tempfile

import streamlit as st
import pandas as pd
import requests
from PIL import Image, ImageOps, ImageDraw, ImageFilter, ImageFont, ImageFile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# -------------------------------------------------------------------
# PIL — tolerância a imagens problemáticas
# -------------------------------------------------------------------
ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = 60_000_000

# -------------------------------------------------------------------
# LOGGING
# -------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    stream=sys.stdout,
)
logger = logging.getLogger("gerador_book")

# -------------------------------------------------------------------
# CONFIG GERAL
# -------------------------------------------------------------------
st.set_page_config(page_title="Gerador de Book Ultra-Resiliente", page_icon="📸", layout="wide")
st.set_option("client.showErrorDetails", True)

# --- CSS básico (UX) ---
BASE_CSS = """
<style>
.steps {display:flex; gap:8px; align-items:center; margin:.25rem 0 .5rem;}
.step {padding:6px 10px; border-radius:999px; border:1px solid #E5E7EB; color:#111; background:#fff; font-weight:700; font-size:13px;}
.step.active {background:#FF7A00; color:#fff; border-color:#FF7A00;}
.step.sep {opacity:.5}

.img-card { border:1px solid #DDD; border-radius:10px; padding:10px; transition:all .15s ease;
  background:#fff; display:flex; flex-direction:column; gap:8px; height:100%; }
.dark .img-card { background:#0f131a; border-color:#232a36; }
.img-card:hover { box-shadow:0 8px 20px rgba(0,0,0,.06); transform:translateY(-2px); }
.group-head {display:flex; justify-content:space-between; align-items:center; padding:.25rem 0;}
.badge {display:inline-block; padding:2px 8px; border-radius:999px; border:1px solid #e5e7eb; font-size:12px;}
.reset-zone, .logout-zone {margin-top:.5rem;}
.quick-actions {display: flex; gap: 10px; margin: 10px 0;}
.quick-actions button {flex: 1;}

/* Estatísticas */
.stats-container {
    background:#f8f9fa; 
    padding:12px 15px; 
    border-radius:8px; 
    margin:10px 0;
    border-left:4px solid #FF7A00;
}
.stats-title {
    font-weight:600; 
    margin-bottom:8px;
    color:#333;
}
.stats-grid {
    display:flex; 
    gap:12px; 
    flex-wrap:wrap;
}
.stats-badge {
    padding:6px 12px;
    border-radius:6px;
    font-size:13px;
    font-weight:500;
}
.stats-success {
    background:#d4edda;
    color:#155724;
    border:1px solid #c3e6cb;
}
.stats-error {
    background:#f8d7da;
    color:#721c24;
    border:1px solid #f5c6cb;
}
.stats-warning {
    background:#fff3cd;
    color:#856404;
    border:1px solid #ffeaa7;
}
.stats-info {
    background:#d1ecf1;
    color:#0c5460;
    border:1px solid #bee5eb;
}
</style>
"""
st.markdown(BASE_CSS, unsafe_allow_html=True)

# === PARTE 2/10 =====================================================
# Continuação dos estilos (CSS login + tema)

LOGIN_CSS = """
<style>
.login-wrap { max-width: 520px; margin: 0 auto; }
.login-hero {
  background: linear-gradient(90deg, #FF7A00 0%, #FF9944 100%);
  color:#fff; padding:16px 18px; border-radius:14px 14px 0 0;
  font-weight:700; line-height:1.2; margin: 8px 0 0 0;
}
.login-card {
  border:1px solid #E7E7E7; border-top:none; background:#ffffff;
  padding:18px; border-radius:0 0 14px 14px; margin:0;
}
.stForm .stButton > button {
  background:#FF7A00 !important; color:#fff !important;
  border:none !important; border-radius:10px !important;
  font-weight:800 !important;
}
.stForm .stButton > button:hover { background:#E66E00 !important; }
</style>
"""

# -------------------------------------------------------------------
# TEMA
# -------------------------------------------------------------------
def apply_theme(dark: bool):
    ORANGE = "#FF7A00"
    ORANGE_HOVER = "#E66E00"
    BLACK = "#111111"
    GRAY_BG = "#f6f6f7"
    if dark:
        palette = f"""
        <style>
        :root {{ --accent:{ORANGE}; --accent-hover:{ORANGE_HOVER}; --text:#f5f5f5; --bg:#0e1117; --panel:#11151c; }}
        .stApp {{ background-color:var(--bg); color:var(--text); }}
        section[data-testid="stSidebar"] > div {{ background:var(--panel); border-right:1px solid #1b212c; }}
        .stButton > button, .stDownloadButton > button {{ background:var(--accent); color:#fff; border-radius:10px; border:none; }}
        .stButton > button:hover, .stDownloadButton > button:hover {{ background:var(--accent-hover); }}
        .stProgress > div > div {{ background-color:var(--accent); }}
        </style>
        """
    else:
        palette = f"""
        <style>
        :root {{ --accent:{ORANGE}; --accent-hover:{ORANGE_HOVER}; --text:{BLACK}; --bg:#ffffff; --panel:{GRAY_BG}; }}
        .stApp {{ background-color:var(--bg); color:var(--text); }}
        section[data-testid="stSidebar"] > div {{ background:var(--panel); border-right:1px solid #ececec; }}
        .stButton > button, .stDownloadButton > button {{ background:var(--accent); color:#fff; border-radius:10px; border:none; }}
        .stButton > button:hover, .stDownloadButton > button:hover {{ background:var(--accent-hover); }}
        .stProgress > div > div {{ background-color:var(--accent); }}
        </style>
        """
    st.markdown(palette, unsafe_allow_html=True)

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# === PARTE 3/10 =====================================================
# Login

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
    "david.silva@mkthouse.com.br": "mudar12345",
}
ALLOWED_USERS = {k.strip().lower(): v for k, v in ALLOWED_USERS.items()}

def do_login():
    st.markdown(LOGIN_CSS, unsafe_allow_html=True)
    st.title("🔒 Acesso Restrito")
    st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="login-hero">Use seu e-mail corporativo. Em caso de dúvidas, contate o BI.</div>', unsafe_allow_html=True)
    st.markdown('<div class="login-card">', unsafe_allow_html=True)
    with st.form("login_form", clear_on_submit=False):
        email = st.text_input("E-mail", placeholder="seu.email@mkthouse.com.br")
        pwd = st.text_input("Senha", type="password", placeholder="••••••••")
        entrar = st.form_submit_button("Entrar")
    st.markdown('</div></div>', unsafe_allow_html=True)

    if entrar:
        email_norm = (email or "").strip().lower()
        if email_norm in ALLOWED_USERS and pwd == ALLOWED_USERS[email_norm]:
            st.session_state.auth = True
            st.session_state.user_email = email_norm
            logger.info(f"Usuário autenticado: {email_norm}")
            st.rerun()
        else:
            logger.warning(f"Tentativa de login falhou: {email_norm}")
            st.error("Credenciais inválidas.")

if "auth" not in st.session_state:
    st.session_state.auth = False

# === PARTE 4/10 =====================================================
# Funções utilitárias

URL_RE = re.compile(r'https?://\S+')

def extrair_links(celula):
    """Extrai todas as URLs de uma célula, mesmo com múltiplas linhas ou caracteres especiais"""
    if pd.isna(celula):
        return []
    t = str(celula).replace("<br>", " ").replace("\n", " ").replace("\r", " ").replace(",", " ").replace("(", " ").replace(")", " ").replace('"', " ").replace("'", " ")
    return [u.rstrip(").,;:") for u in URL_RE.findall(t)]

def redimensionar(img: Image.Image, max_w: int, max_h: int) -> Image.Image:
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((max_w, max_h), resample=Image.LANCZOS)
    return img

def comprimir_jpeg_binsearch(img: Image.Image, limite_kb: int) -> BytesIO:
    lo, hi, best = 35, 95, None
    buf = BytesIO()
    img.save(buf, "JPEG", quality=75, optimize=True, progressive=True, subsampling=2)
    if buf.tell()/1024 <= limite_kb:
        buf.seek(0)
        return buf
    best = buf
    while lo <= hi:
        mid = (lo+hi)//2
        buf = BytesIO()
        img.save(buf, "JPEG", quality=mid, optimize=True, progressive=True, subsampling=2)
        if buf.tell()/1024 <= limite_kb:
            best = buf
            lo = mid+1
        else:
            hi = mid-1
    if best is None:
        best = BytesIO()
        img.save(best, "JPEG", quality=35, optimize=True, progressive=True, subsampling=2)
    best.seek(0)
    return best

def px_to_inches(px): 
    return Inches(px / 96.0)

def hex_to_rgb(hex_str: str):
    s = hex_str.strip().lstrip("#")
    if len(s) == 3:
        s = "".join([c*2 for c in s])
    return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)

def pick_contrast_color(r, g, b):
    brightness = (r*299 + g*587 + b*114) / 1000
    return (0,0,0) if brightness > 128 else (255,255,255)

# --- HASHES p/ duplicatas ---
def _sha1_bytes(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()

def _img_dhash(img: Image.Image, hash_size: int = 8) -> str:
    im = ImageOps.exif_transpose(img).convert("L").resize((hash_size + 1, hash_size), Image.LANCZOS)
    pixels = np.asarray(im, dtype=np.int16)
    diff = pixels[:, 1:] > pixels[:, :-1]
    bits = 0
    for row in diff:
        for v in row:
            bits = (bits << 1) | int(v)
    return f"{bits:0{hash_size*hash_size//4}x}"

# ---- Cache em disco (/tmp) ----
TMP_ROOT = tempfile.gettempdir()
APP_TMP_DIR = os.path.join(TMP_ROOT, "gerador_book_cache")
os.makedirs(APP_TMP_DIR, exist_ok=True)

def _save_bytes_to_tmp(ext: str, data: bytes) -> str:
    fd, path = tempfile.mkstemp(prefix="gb_", suffix=f".{ext}", dir=APP_TMP_DIR)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path

def _open_img_from_path(path: str) -> Image.Image:
    im = Image.open(path)
    im.load()
    return im

# === PARTE 5/10 =====================================================
# Qualidade da imagem + Efeitos

def _laplacian_var_gray(pil_img: Image.Image) -> float:
    g = pil_img.convert("L")
    a = np.asarray(g, dtype=np.float32)
    H, W = a.shape
    if H < 3 or W < 3:
        return 0.0
    out = (a[0:-2,1:-1] + a[1:-1,0:-2] + a[1:-1,2:] + a[2:,1:-1] - 4*a[1:-1,1:-1])
    return float(out.var())

def medir_qualidade(img: Image.Image) -> dict:
    im = ImageOps.exif_transpose(img)
    w, h = im.size
    im_small = im.copy()
    im_small.thumbnail((1024, 1024), Image.LANCZOS)
    g = im_small.convert("L")
    arr = np.asarray(g, dtype=np.float32)
    return {
        "width": int(w),
        "height": int(h),
        "megapixels": float((w*h)/1_000_000.0),
        "mean_brightness": float(arr.mean()),
        "std_contrast": float(arr.std()),
        "blur_score": _laplacian_var_gray(im_small),
    }

def _hex_to_rgba_tuple(hex_color, alpha=255):
    s = hex_color.strip().lstrip("#")
    if len(s) == 3:
        s = "".join([c*2 for c in s])
    r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    return (r, g, b, alpha)

def _apply_rounded_corners(img_rgba: Image.Image, radius: int) -> Image.Image:
    if radius <= 0:
        return img_rgba
    w, h = img_rgba.size
    mask = Image.new("L", (w, h), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, w, h], radius=radius, fill=255)
    out = img_rgba.copy()
    out.putalpha(mask)
    return out

def _apply_border_color(img_rgba: Image.Image, border_px: int, border_hex: str, radius: int) -> Image.Image:
    if border_px <= 0:
        return img_rgba
    w, h = img_rgba.size
    result = Image.new("RGBA", (w + 2*border_px, h + 2*border_px), (0,0,0,0))
    draw = ImageDraw.Draw(result)
    outer = [0, 0, result.size[0], result.size[1]]
    inner = [border_px, border_px, border_px + w, border_px + h]
    draw.rounded_rectangle(outer, radius=radius+border_px, fill=_hex_to_rgba_tuple(border_hex))
    hole = Image.new("L", result.size, 255)
    hole_draw = ImageDraw.Draw(hole)
    hole_draw.rounded_rectangle(inner, radius=radius, fill=0)
    result.putalpha(hole)
    result.alpha_composite(img_rgba, dest=(border_px, border_px))
    return result

def _apply_drop_shadow(img_rgba: Image.Image, blur: int, offset: int, opacity_pct: int) -> Image.Image:
    if blur <= 0 and offset <= 0:
        return img_rgba
    alpha = img_rgba.split()[-1]
    a = max(0, min(255, int(255 * (opacity_pct/100))))
    pad = blur + offset + 2
    w, h = img_rgba.size
    canvas = Image.new("RGBA", (w + pad, h + pad), (0,0,0,0))
    shadow = Image.new("RGBA", (w, h), (0,0,0,a))
    shadow.putalpha(alpha)
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=blur))
    canvas.alpha_composite(shadow, dest=(offset, offset))
    canvas.alpha_composite(img_rgba, dest=(0,0))
    return canvas

def apply_effects_pipeline(img_rgb: Image.Image, cfg: dict) -> Image.Image:
    out = img_rgb.convert("RGBA")
    if cfg.get("fx_round"):
        out = _apply_rounded_corners(out, int(cfg.get("fx_round_radius", 20)))
    if cfg.get("fx_border"):
        out = _apply_border_color(
            out,
            int(cfg.get("fx_border_width", 6)),
            cfg.get("fx_border_color", "#DDDDDD"),
            int(cfg.get("fx_round_radius", 20)) if cfg.get("fx_round") else 0
        )
    if cfg.get("fx_shadow"):
        out = _apply_drop_shadow(
            out,
            int(cfg.get("fx_shadow_blur", 10)),
            int(cfg.get("fx_shadow_offset", 8)),
            int(cfg.get("fx_shadow_opac", 40)),
        )
    return out

# === PARTE 6/10 =====================================================
# VERSÃO ESPECIAL JOTFORM - Download com retry, backoff e headers realistas

def baixar_processar_resiliente(session, url: str, max_w: int, max_h: int, limite_kb: int, timeout: int, fx_cfg: dict = None, max_retries: int = 8):
    """
    Versão ESPECIAL para Jotform com:
    - 8 tentativas máximas
    - Backoff exponencial (1s, 2s, 4s, 8s, 16s, 32s)
    - Headers realistas
    - Cookies persistentes
    - Tratamento especial para redirects
    """
    fx_cfg = fx_cfg or {}
    
    # Headers COMPLETOS igual a um navegador real
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Referer": "https://www.jotform.com/",
        "Sec-Fetch-Dest": "image",
        "Sec-Fetch-Mode": "no-cors",
        "Sec-Fetch-Site": "cross-site",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }
    
    # Cookies para manter sessão
    session.cookies.set("jotform_antifraud", "1", domain=".jotform.com")
    session.cookies.set("locale", "pt-BR", domain=".jotform.com")
    
    for attempt in range(max_retries):
        try:
            logger.info(f"Tentativa {attempt + 1}/{max_retries} para: {url.split('/')[-1] if '/' in url else url[:50]}")
            
            # Timeout maior a cada tentativa
            current_timeout = timeout + (attempt * 10)
            
            # Tentativa HEAD primeiro para verificar
            try:
                head_resp = session.head(url, timeout=10, headers=headers, allow_redirects=True)
                if head_resp.status_code in [301, 302, 307, 308]:
                    redirect_url = head_resp.headers.get('Location', url)
                    logger.info(f"Redirect para: {redirect_url.split('/')[-1] if '/' in redirect_url else redirect_url[:50]}")
                    url = redirect_url
            except:
                pass
            
            # Download principal
            r = session.get(url, timeout=current_timeout, stream=True, headers=headers, allow_redirects=True)
            
            # Verificar status
            if r.status_code in [403, 404, 410]:
                logger.warning(f"HTTP {r.status_code} - Link pode ter expirado")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, f"HTTP {r.status_code} - Link expirado")
                time.sleep(3 ** attempt)
                continue
            
            if r.status_code != 200:
                logger.warning(f"HTTP {r.status_code} na tentativa {attempt + 1}")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, f"HTTP {r.status_code}")
                time.sleep(2 ** attempt)
                continue
            
            # Verificar tamanho
            raw_bytes = r.content
            if len(raw_bytes) == 0:
                logger.warning(f"Arquivo vazio")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, "Arquivo vazio")
                time.sleep(2 ** attempt)
                continue
            
            # Tentar abrir imagem
            try:
                img = Image.open(BytesIO(raw_bytes))
                img.load()
                logger.info(f"✅ Imagem carregada: {img.size}")
            except Exception as e:
                logger.warning(f"Erro ao abrir imagem: {e}")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, f"Imagem corrompida: {str(e)[:100]}")
                time.sleep(2 ** attempt)
                continue
            
            # Sucesso!
            sha1_hex = _sha1_bytes(raw_bytes)
            im_small_for_hash = img.copy()
            im_small_for_hash.thumbnail((256, 256), Image.LANCZOS)
            dhash_hex = _img_dhash(im_small_for_hash)
            
            quality = medir_qualidade(img)
            img = redimensionar(img, max_w, max_h)
            
            need_alpha = bool(fx_cfg and (fx_cfg.get("fx_shadow") or fx_cfg.get("fx_round") or fx_cfg.get("fx_border")))
            if need_alpha:
                img_rgba = apply_effects_pipeline(img.convert("RGB"), fx_cfg)
                buf = BytesIO()
                img_rgba.save(buf, format="PNG", optimize=True)
                if buf.tell() / 1024 <= limite_kb:
                    path = _save_bytes_to_tmp("png", buf.getvalue())
                    w, h = img_rgba.size
                    return (url, True, path, (w, h), quality, sha1_hex, dhash_hex, None)
                bg = Image.new("RGB", img_rgba.size, (255, 255, 255))
                bg.paste(img_rgba, mask=img_rgba.split()[-1])
                buf = comprimir_jpeg_binsearch(bg, limite_kb)
                path = _save_bytes_to_tmp("jpg", buf.getvalue())
                w, h = bg.size
                return (url, True, path, (w, h), quality, sha1_hex, dhash_hex, None)
            else:
                buf = comprimir_jpeg_binsearch(img.convert("RGB"), limite_kb)
                path = _save_bytes_to_tmp("jpg", buf.getvalue())
                w, h = img.size
                return (url, True, path, (w, h), quality, sha1_hex, dhash_hex, None)
                
        except requests.exceptions.Timeout:
            logger.warning(f"Timeout tentativa {attempt + 1}")
            if attempt == max_retries - 1:
                return (url, False, None, None, None, None, None, f"Timeout após {max_retries} tentativas")
            time.sleep(3 ** attempt)
            
        except requests.exceptions.ConnectionError as e:
            logger.warning(f"Erro de conexão: {e}")
            if attempt == max_retries - 1:
                return (url, False, None, None, None, None, None, f"Erro de conexão")
            time.sleep(3 ** attempt)
            
        except Exception as e:
            logger.warning(f"Erro: {e}")
            if attempt == max_retries - 1:
                return (url, False, None, None, None, None, None, f"Erro: {str(e)[:100]}")
            time.sleep(2 ** attempt)
        
        finally:
            gc.collect()
    
    return (url, False, None, None, None, None, None, "Todas as tentativas falharam")

# === PARTE 7/10 =====================================================
# PPT helpers e funções auxiliares

def get_slots(n, prs):
    IMG_TOP = Inches(1.2)
    CONTENT_W = Inches(11)
    CONTENT_H = Inches(6)
    GAP = Inches(0.2)
    start_left = (prs.slide_width - CONTENT_W) / 2
    if n == 1:
        return [(start_left, IMG_TOP, CONTENT_W, CONTENT_H)]
    cols = n
    total_gap = GAP * (cols - 1)
    cell_w = (CONTENT_W - total_gap) / cols
    return [(start_left + c * (cell_w + GAP), IMG_TOP, cell_w, CONTENT_H) for c in range(cols)]

def add_title_and_address(slide, title_text, address_text, title_rgb=(0,0,0),
                          font_name="Radikal", title_font_size_pt=18, title_font_bold=True):
    TITLE_LEFT, TITLE_TOP, TITLE_W = Inches(0.5), Inches(0.2), Inches(12)
    tx = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_W, Inches(1))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    f = run.font
    f.name = font_name or "Radikal"
    f.size = Pt(title_font_size_pt or 18)
    f.bold = bool(title_font_bold)
    f.color.rgb = RGBColor(*title_rgb)
    p.alignment = 1
    if address_text:
        p2 = tf.add_paragraph()
        run2 = p2.add_run()
        run2.text = address_text
        f2 = run2.font
        f2.name = font_name or "Radikal"
        f2.size = Pt(max(8, (title_font_size_pt or 18) / 2))
        f2.bold = False
        f2.color.rgb = RGBColor(*title_rgb)
        p2.alignment = 1

def set_slide_bg(slide, rgb_tuple):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb_tuple)

def place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in):
    img_w_in = px_to_inches(w_px)
    img_h_in = px_to_inches(h_px)
    ratio = min(float(max_w_in) / float(img_w_in), float(max_h_in) / float(img_h_in), 1.0)
    final_w = img_w_in * ratio
    final_h = img_h_in * ratio
    x = left + (max_w_in - final_w) / 2
    y = top + (max_h_in - final_h) / 2
    slide.shapes.add_picture(file_path, x, y, width=final_w, height=final_h)

def is_portrait(w_px: int, h_px: int, tol: float = 1.05) -> bool:
    if w_px <= 0 or h_px <= 0:
        return False
    return (h_px / float(w_px)) >= tol

def move_slide_to_index(prs, old_index, new_index):
    sldIdLst = prs.slides._sldIdLst
    sld = sldIdLst[old_index]
    sldIdLst.remove(sld)
    sldIdLst.insert(new_index, sld)

def add_logo_top_right(slide, prs, logo_bytes: bytes, logo_width_in: float):
    if not logo_bytes:
        return
    left = prs.slide_width - Inches(0.5) - Inches(logo_width_in)
    top = Inches(0.2)
    slide.shapes.add_picture(BytesIO(logo_bytes), left, top, width=Inches(logo_width_in))

def add_signature_bottom_right(slide, prs, signature_bytes: bytes, signature_width_in: float,
                               bottom_margin_in: float = 0.2, right_margin_in: float = 0.2):
    if not signature_bytes:
        return
    try:
        im = Image.open(BytesIO(signature_bytes))
        w_px, h_px = im.size
        ratio = (h_px / float(w_px)) if w_px else 0.4
    except Exception:
        ratio = 0.4
    sig_h_in = signature_width_in * ratio
    left = prs.slide_width - Inches(right_margin_in) - Inches(signature_width_in)
    top = prs.slide_height - Inches(bottom_margin_in) - Inches(sig_h_in)
    slide.shapes.add_picture(BytesIO(signature_bytes), left, top, width=Inches(signature_width_in))

# === PARTE 8/10 =====================================================
# ZIP de imagens + PPT

def _path_to_jpeg_bytes(file_path: str) -> bytes:
    try:
        im = Image.open(file_path)
        fmt = (im.format or "").upper()
        if fmt in ("JPEG", "JPG"):
            with open(file_path, "rb") as f:
                return f.read()
        if im.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            im = bg
        else:
            im = im.convert("RGB")
        out = BytesIO()
        im.save(out, "JPEG", quality=88, optimize=True, progressive=True, subsampling=2)
        out.seek(0)
        return out.read()
    except Exception as e:
        logger.warning(f"_path_to_jpeg_bytes fallback: {e}")
        try:
            with open(file_path, "rb") as f:
                return f.read()
        except:
            return b""

def _sanitize_folder_name(name: str) -> str:
    safe = re.sub(r'[\\/:*?"<>|]+', ' ', str(name or "").strip())
    safe = re.sub(r'\s+', ' ', safe)
    return safe[:80] if len(safe) > 80 else safe

def montar_zip_imagens(items, resultados, excluded_urls: set) -> BytesIO:
    grupos = OrderedDict()
    for loja, endereco, url in items:
        if (url in resultados) and (url not in excluded_urls):
            grupos.setdefault(str(loja), []).append((url, resultados[url]))

    mem_zip = BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for loja, lista in grupos.items():
            pasta = _sanitize_folder_name(loja) or "Sem Nome"
            contador = 1
            for _url, (_loja, _end, file_path, (w, h), *_) in lista:
                jpeg_bytes = _path_to_jpeg_bytes(file_path)
                arquivo = f"{pasta} - {contador}.jpg"
                caminho = f"{pasta}/{arquivo}"
                try:
                    zf.writestr(caminho, jpeg_bytes)
                except Exception as e:
                    logger.warning(f"Falha ao escrever {caminho} no ZIP: {e}")
                contador += 1
    mem_zip.seek(0)
    return mem_zip

def gerar_ppt(slides_data, cfg, titulo, excluded_urls):
    """Versão simplificada para gerar PPT sem modelo"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    blank = prs.slide_layouts[6]
    
    groups = OrderedDict()
    for loja, endereco, url in slides_data["items"]:
        if url in slides_data["resultados"] and url not in excluded_urls:
            groups.setdefault(str(loja), []).append((url, slides_data["resultados"][url]))
    
    sort_mode = cfg.get("sort_mode", "Ordem original do Excel")
    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys())
    else:
        loja_keys = list(groups.keys())
    
    title_rgb = cfg.get("title_font_color_rgb", pick_contrast_color(*cfg["bg_rgb"]))
    signature_width = (cfg["logo_width_in"]/2.0) if cfg.get("auto_half_signature", True) else (cfg.get("signature_width_in") or 0.6)
    
    for loja in loja_keys:
        imgs = groups[loja]
        i = 0
        while i < len(imgs):
            if cfg["max_per_slide"] == "Automático":
                _, _, _, (w0, h0), *_ = imgs[i][1]
                per_slide = 3 if is_portrait(w0, h0) else 2
                endereco = imgs[i][1][1]
            else:
                per_slide = int(cfg["max_per_slide"])
                endereco = imgs[i][1][1]
            
            batch = imgs[i:i+per_slide]
            i += per_slide
            slide = prs.slides.add_slide(blank)
            set_slide_bg(slide, cfg["bg_rgb"])
            add_title_and_address(slide, loja, endereco, title_rgb,
                                  cfg["title_font_name"], cfg["title_font_size_pt"], cfg["title_font_bold"])
            if cfg.get("logo_bytes"):
                add_logo_top_right(slide, prs, cfg["logo_bytes"], cfg["logo_width_in"])
            if cfg.get("signature_bytes"):
                add_signature_bottom_right(slide, prs, cfg["signature_bytes"], signature_width,
                                          bottom_margin_in=cfg["signature_bottom_margin_in"],
                                          right_margin_in=cfg["signature_right_margin_in"])
            slots = get_slots(len(batch), prs)
            for (_, (_, _, file_path, (w_px, h_px), *_)), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                try:
                    place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in)
                except Exception as e:
                    logger.warning(f"Falha ao inserir imagem: {e}")
    
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# === PARTE 9/10 =====================================================
# UI de miniaturas + detecção + reset

def img_to_html_with_border(image: Image.Image, width_px: int, border_px: int, border_color: str):
    im = image.copy()
    im.thumbnail((width_px, width_px))
    buf = BytesIO()
    im.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    style = f"border:{border_px}px solid {border_color};border-radius:10px;display:block;max-width:100%;width:{width_px}px;"
    return f'<img src="data:image/png;base64,{b64}" style="{style}" />'

def render_steps(current: int):
    labels = ["Upload", "Pré-visualização", "Gerar/Exportar"]
    html = ['<div class="steps">']
    for i, txt in enumerate(labels, start=1):
        cls = "step active" if i == current else "step"
        html.append(f'<div class="{cls}">{i}. {txt}</div>')
        if i < len(labels):
            html.append('<span class="step sep">›</span>')
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)

def render_summary(items, resultados, excluded, failed_details=None):
    total_urls = len(items)
    baixadas = sum(1 for _, _, url in items if url in resultados)
    lojas = len({loja for loja, _, _ in items})
    falhas = len(failed_details) if failed_details else 0
    
    stats_html = f"""
    <div class="stats-container">
        <div class="stats-title">📊 Estatísticas</div>
        <div class="stats-grid">
            <span class="stats-badge stats-info">Lojas: {lojas}</span>
            <span class="stats-badge stats-info">Total de links: {total_urls}</span>
            <span class="stats-badge stats-success">✅ Baixadas: {baixadas}</span>
            <span class="stats-badge stats-error">❌ Falhas: {falhas}</span>
            <span class="stats-badge stats-warning">🚫 Excluídas: {len(excluded)}</span>
        </div>
    </div>
    """
    st.markdown(stats_html, unsafe_allow_html=True)
    return {"total_urls": total_urls, "baixadas": baixadas, "lojas": lojas, "falhas": falhas, "excluidas": len(excluded)}

def detectar_problemas(resultados, min_mp=0.8, min_blur=45):
    low_quality = set()
    by_sha1, by_dhash = {}, {}
    for url, tup in resultados.items():
        loja, endereco, file_path, (w, h), quality, sha1_hex, dhash_hex = tup
        if (quality.get("megapixels", 0) < min_mp) or (quality.get("blur_score", 0) < min_blur):
            low_quality.add(url)
        by_sha1.setdefault(sha1_hex, []).append(url)
        by_dhash.setdefault(dhash_hex, []).append(url)
    
    duplicates = set()
    for group in by_sha1.values():
        if len(group) > 1:
            duplicates.update(group[1:])
    for group in by_dhash.values():
        if len(group) > 1:
            duplicates.update([u for u in group[1:] if u not in duplicates])
    
    return low_quality, duplicates

def reset_app(preserve_login: bool = True):
    user = st.session_state.get("user_email")
    auth = st.session_state.get("auth", False)
    st.session_state.clear()
    
    for k in ["xlsx_key", "template_key", "logo_key", "sign_key", "download_key", "images_zip_key"]:
        st.session_state[k] = 0
    
    st.session_state.exp_plan = True
    st.session_state.exp_style = False
    st.session_state.exp_brand = False
    st.session_state.exp_fx = False
    st.session_state.exp_perf = False
    st.session_state.exp_model = False
    
    st.session_state.ppt_bytes = None
    st.session_state.images_zip_bytes = None
    st.session_state.generated = False
    st.session_state.output_filename = "Modelo_01"
    st.session_state.pipeline = {}
    st.session_state.excluded_urls = set()
    st.session_state.preview_mode = False
    st.session_state.expanded_groups = {}
    st.session_state.failed_urls = []
    st.session_state.failed_details = []
    
    if preserve_login and auth:
        st.session_state.auth = True
        st.session_state.user_email = user
        st.session_state.dark_mode = False
    st.rerun()

# === PARTE 10/10 ====================================================
# APP principal

def main_app():
    # Inicializações
    for k in ["xlsx_key", "template_key", "logo_key", "sign_key", "download_key", "images_zip_key"]:
        if k not in st.session_state:
            st.session_state[k] = 0
    for k in ["exp_plan", "exp_style", "exp_brand", "exp_fx", "exp_perf", "exp_model"]:
        if k not in st.session_state:
            st.session_state[k] = True if k == "exp_plan" else False
    if "pipeline" not in st.session_state:
        st.session_state.pipeline = {}
    if "excluded_urls" not in st.session_state:
        st.session_state.excluded_urls = set()
    if "preview_mode" not in st.session_state:
        st.session_state.preview_mode = False
    if "expanded_groups" not in st.session_state:
        st.session_state.expanded_groups = {}
    if "output_filename" not in st.session_state:
        st.session_state.output_filename = "Modelo_01"
    if "generated" not in st.session_state:
        st.session_state.generated = False
    if "ppt_bytes" not in st.session_state:
        st.session_state.ppt_bytes = None
    if "images_zip_bytes" not in st.session_state:
        st.session_state.images_zip_bytes = None
    if "failed_urls" not in st.session_state:
        st.session_state.failed_urls = []
    if "failed_details" not in st.session_state:
        st.session_state.failed_details = []
    if "quick_generate" not in st.session_state:
        st.session_state.quick_generate = False
    if "ignore_failed" not in st.session_state:
        st.session_state.ignore_failed = True

    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)
        
        st.markdown("---")
        st.subheader("🚀 Configuração Ultra-Resiliente para Jotform")
        st.caption("Otimizado especificamente para links do Jotform")
        
        with st.expander("📄 Planilha & Layout", expanded=st.session_state.exp_plan):
            loja_col = st.text_input("🛒 Coluna de LOJA", value="Selecione sua loja", key="loja_col")
            img_col = st.text_input("🖼️ Coluna de FOTOS", value="Faça o upload das fotos", key="img_col")
            use_address = st.checkbox("➕ Incluir endereço", value=False, key="use_address")
            address_col = st.text_input("🏠 Coluna de ENDEREÇO", value="Endereço", key="address_col", disabled=not use_address)
            max_per_slide = st.selectbox("📐 Fotos por slide", ["Automático", 1, 2, 3], index=0, key="max_per_slide")
            sort_mode = st.selectbox("🔤 Ordenar lojas", ["Ordem original do Excel", "Nome da loja (A→Z)"], index=0, key="sort_mode")

        with st.expander("🎨 Aparência", expanded=st.session_state.exp_style):
            bg_hex = st.color_picker("🎨 Cor de fundo", value="#FFFFFF", key="bg_hex")
            title_font_name = st.text_input("Fonte", value="Radikal", key="title_font_name")
            title_font_size_pt = st.slider("Tamanho (pt)", 8, 48, 18, 1, key="title_font_size_pt")
            title_font_bold = st.checkbox("Negrito", value=True, key="title_font_bold")
            title_font_color = st.color_picker("Cor da fonte", value="#000000", key="title_font_color")

        with st.expander("🏷️ Logo & Assinatura", expanded=st.session_state.exp_brand):
            logo_file = st.file_uploader("Logo", type=["png", "jpg", "jpeg"], key=f"logo_uploader_{st.session_state.logo_key}")
            if "logo_bytes" not in st.session_state:
                st.session_state.logo_bytes = None
            if logo_file:
                st.session_state.logo_bytes = logo_file.getvalue()
            logo_width_in = st.slider("Largura do LOGO (pol)", 0.5, 3.0, 1.2, 0.1, key="logo_width_in")
            
            signature_file = st.file_uploader("Assinatura", type=["png", "jpg", "jpeg"], key=f"signature_uploader_{st.session_state.sign_key}")
            if "signature_bytes" not in st.session_state:
                st.session_state.signature_bytes = None
            if signature_file:
                st.session_state.signature_bytes = signature_file.getvalue()
            
            signature_right_margin_in = st.slider("Margem direita assinatura", 0.0, 1.0, 0.20, 0.05, key="sig_right_margin")
            signature_bottom_margin_in = st.slider("Margem inferior assinatura", 0.0, 1.0, 0.20, 0.05, key="sig_bottom_margin")

        with st.expander("🎭 Efeitos", expanded=st.session_state.exp_fx):
            fx_shadow = st.checkbox("Sombra", value=False, key="fx_shadow")
            fx_round = st.checkbox("Borda arredondada", value=False, key="fx_round")
            fx_border = st.checkbox("Borda colorida", value=False, key="fx_border")

        with st.expander("⚡ Performance (Jotform)", expanded=st.session_state.exp_perf):
            st.warning("⚠️ Configurações críticas para Jotform")
            max_retries = st.slider("🔄 Tentativas por imagem", 5, 15, 8, 1, key="max_retries")
            timeout_base = st.slider("⏱️ Timeout base (segundos)", 30, 120, 60, 10, key="timeout_base")
            max_workers = st.slider("📡 Downloads simultâneos", 1, 4, 2, 1, key="max_workers")
            
            target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10, key="target_w")
            target_h = st.number_input("Altura máx (px)", 360, 4096, 720, 10, key="target_h")
            limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 3000, 800, 10, key="limite_kb")
            
            min_mp = st.slider("Megapixels mínimos", 0.05, 5.0, 0.3, 0.05, key="min_megapixels")
            min_blur = st.slider("Nitidez mínima", 5, 300, 25, 5, key="min_blur_score")
            
            thumb_px = st.slider("Miniaturas (px)", 120, 400, 220, 10, key="thumb_px")
            thumbs_per_row = st.slider("Miniaturas por linha", 2, 8, 4, 1, key="thumbs_per_row")
            
            ignore_failed = st.checkbox("⚠️ Ignorar falhas", value=True, key="ignore_failed")

    # Topo
    top_l, top_m, top_r = st.columns([5,1,1])
    with top_l:
        current_step = 1
        if st.session_state.get("preview_mode") and not st.session_state.get("generated"):
            current_step = 2
        if st.session_state.get("generated") or st.session_state.get("images_zip_bytes"):
            current_step = 3
        st.title("📸 Gerador de Book - Especial Jotform")
        render_steps(current_step)
        st.caption("Otimizado para Jotform com 8 tentativas, backoff exponencial e cookies persistentes")
    with top_m:
        if st.button("Resetar", key="reset_btn", use_container_width=True, type="secondary"):
            reset_app(preserve_login=True)
    with top_r:
        if st.button("Sair", key="logout_btn", use_container_width=True, type="secondary"):
            reset_app(preserve_login=False)

    main_expander = st.expander("📋 Gerador de Book - Painel Principal", expanded=True)
    
    with main_expander:
        st.subheader("1. Upload da Planilha")
        up = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"], key=f"xlsx_upload_{st.session_state.xlsx_key}")
        
        if up:
            col1, col2 = st.columns(2)
            with col1:
                btn_preview = st.button("👁️ Visualização Rápida", key="btn_preview", use_container_width=True)
            with col2:
                btn_generate_direct = st.button("🚀 Gerar PPT Direto (Jotform)", key="btn_generate_direct", use_container_width=True, type="primary")
            
            if btn_preview or btn_generate_direct:
                try:
                    df = pd.read_excel(up)
                except Exception as e:
                    st.error(f"Erro ao ler Excel: {e}")
                    st.stop()
                
                loja_col = st.session_state["loja_col"]
                img_col = st.session_state["img_col"]
                use_address = st.session_state.get("use_address", False)
                address_col = st.session_state.get("address_col", "Endereço")
                
                required_cols = [loja_col, img_col]
                if use_address:
                    required_cols.append(address_col)
                
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    st.error(f"Colunas não encontradas: {missing}")
                    st.stop()
                
                items = []
                url_line_map = {}
                for ridx, row in df.iterrows():
                    loja = str(row[loja_col]).strip()
                    endereco = str(row[address_col]).strip() if use_address else ""
                    line_no = ridx + 2
                    urls = extrair_links(row.get(img_col, ""))
                    for url in urls:
                        if url.startswith("http"):
                            items.append((loja, endereco, url))
                            url_line_map[url] = line_no
                
                # Remove duplicatas
                seen = set()
                unique_items = []
                for loja, end, url in items:
                    if url not in seen:
                        seen.add(url)
                        unique_items.append((loja, end, url))
                items = unique_items
                
                total = len(items)
                if total == 0:
                    st.warning("Nenhuma URL encontrada")
                    st.stop()
                
                st.info(f"📥 Processando {total} imagens com {st.session_state['max_retries']} tentativas cada...")
                
                # Configurar sessão
                session = requests.Session()
                adapter = requests.adapters.HTTPAdapter(pool_connections=max_workers, pool_maxsize=max_workers, max_retries=3)
                session.mount("http://", adapter)
                session.mount("https://", adapter)
                
                fx_cfg = {
                    "fx_shadow": st.session_state.get("fx_shadow", False),
                    "fx_round": st.session_state.get("fx_round", False),
                    "fx_border": st.session_state.get("fx_border", False),
                }
                
                prog = st.progress(0)
                status = st.empty()
                resultados = {}
                failed_details = []
                done = 0
                
                with ThreadPoolExecutor(max_workers=st.session_state["max_workers"]) as ex:
                    futures = {
                        ex.submit(
                            baixar_processar_resiliente,
                            session, url,
                            st.session_state["target_w"], st.session_state["target_h"],
                            st.session_state["limite_kb"], st.session_state["timeout_base"],
                            fx_cfg, st.session_state["max_retries"]
                        ): (loja, endereco, url, url_line_map.get(url, "?"))
                        for loja, endereco, url in items
                    }
                    
                    for fut in as_completed(futures):
                        loja, endereco, url, line_no = futures[fut]
                        try:
                            res = fut.result()
                        except Exception as e:
                            res = (url, False, None, None, None, None, None, str(e))
                        
                        if res and len(res) >= 2 and res[1] is True:
                            url_key = res[0]
                            file_path = res[2] if len(res) > 2 else None
                            wh = res[3] if len(res) > 3 else (0, 0)
                            quality = res[4] if len(res) > 4 else {}
                            sha1 = res[5] if len(res) > 5 else ""
                            dhash = res[6] if len(res) > 6 else ""
                            
                            if file_path and wh:
                                resultados[url_key] = (loja, endereco, file_path, wh, quality, sha1, dhash)
                            else:
                                error_msg = res[7] if len(res) > 7 else "Erro"
                                failed_details.append({"url": url, "loja": loja, "linha": line_no, "erro": error_msg})
                        else:
                            error_msg = res[7] if len(res) > 7 else "Erro"
                            failed_details.append({"url": url, "loja": loja, "linha": line_no, "erro": error_msg})
                        
                        done += 1
                        prog.progress(int(done * 100 / total))
                        status.write(f"Progresso: {done}/{total} | ✅ {len(resultados)} | ❌ {len(failed_details)}")
                
                status.write(f"✅ Concluído! Sucesso: {len(resultados)} | Falhas: {len(failed_details)}")
                
                if len(failed_details) > 0 and not st.session_state.get("ignore_failed", True):
                    st.error(f"{len(failed_details)} falhas. Interrompendo.")
                    st.stop()
                elif len(failed_details) > 0:
                    st.warning(f"⚠️ {len(failed_details)} falhas ignoradas. Continuando com {len(resultados)} imagens.")
                    if failed_details:
                        st.dataframe(pd.DataFrame(failed_details), use_container_width=True)
                
                low_q, dups = detectar_problemas(resultados, st.session_state["min_megapixels"], st.session_state["min_blur_score"])
                
                st.session_state.pipeline = {
                    "items": items,
                    "resultados": resultados,
                    "settings": {
                        "max_per_slide": st.session_state["max_per_slide"],
                        "sort_mode": st.session_state["sort_mode"],
                        "bg_rgb": hex_to_rgb(st.session_state["bg_hex"]),
                        "title_font_name": st.session_state["title_font_name"],
                        "title_font_size_pt": st.session_state["title_font_size_pt"],
                        "title_font_bold": st.session_state["title_font_bold"],
                        "title_font_color_rgb": hex_to_rgb(st.session_state["title_font_color"]),
                        "logo_bytes": st.session_state.logo_bytes,
                        "logo_width_in": st.session_state["logo_width_in"],
                        "signature_bytes": st.session_state.signature_bytes,
                        "signature_bottom_margin_in": st.session_state.get("sig_bottom_margin", 0.2),
                        "signature_right_margin_in": st.session_state.get("sig_right_margin", 0.2),
                        "effects": fx_cfg,
                    }
                }
                
                if btn_preview:
                    st.session_state.preview_mode = True
                    st.session_state.quick_generate = False
                else:
                    st.session_state.preview_mode = False
                    st.session_state.quick_generate = True
                    st.session_state.generated = False
                
                st.rerun()
        
        # Download section
        if st.session_state.get("pipeline") and st.session_state.get("quick_generate") and not st.session_state.get("ppt_bytes"):
            with st.spinner("Gerando PPT..."):
                try:
                    pipeline = st.session_state.pipeline
                    ppt_bytes = gerar_ppt(pipeline, pipeline["settings"], st.session_state.output_filename, st.session_state.excluded_urls)
                    st.session_state.ppt_bytes = ppt_bytes
                    st.session_state.generated = True
                    st.success("PPT gerado com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")
        
        if st.session_state.get("pipeline"):
            if not st.session_state.get("generated") and not st.session_state.get("quick_generate"):
                render_summary(st.session_state.pipeline["items"], 
                              st.session_state.pipeline["resultados"], 
                              st.session_state.excluded_urls,
                              st.session_state.get("failed_details", []))
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("🔄 Continuar para Gerar PPT", use_container_width=True, type="primary"):
                        st.session_state.quick_generate = True
                        st.rerun()
                with col2:
                    zip_bytes = montar_zip_imagens(st.session_state.pipeline["items"],
                                                   st.session_state.pipeline["resultados"],
                                                   st.session_state.excluded_urls)
                    st.download_button("📦 Baixar ZIP das Imagens", data=zip_bytes,
                                      file_name=f"{st.session_state.output_filename}_imagens.zip",
                                      mime="application/zip", use_container_width=True)
                with col3:
                    if st.session_state.get("ppt_bytes"):
                        st.download_button("📊 Baixar PPT", data=st.session_state.ppt_bytes,
                                          file_name=f"{st.session_state.output_filename}.pptx",
                                          mime="application/vnd.openxmlformats-officedocument.presentation.presentation",
                                          use_container_width=True)

# -------------------------------------------------------------------
# ROTEAMENTO FINAL
# -------------------------------------------------------------------
if not st.session_state.auth:
    do_login()
else:
    main_app()
