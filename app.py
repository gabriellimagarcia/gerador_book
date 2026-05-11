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
    ORANGE = "#FF7A00"; ORANGE_HOVER = "#E66E00"; BLACK = "#111111"; GRAY_BG = "#f6f6f7"
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
# Funções utilitárias + VERSÃO ULTRA-RESILIENTE

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
# VERSÃO ULTRA-RESILIENTE - Download com retry e backoff exponencial

def baixar_processar_resiliente(session, url: str, max_w: int, max_h: int, limite_kb: int, timeout: int, fx_cfg: dict = None, max_retries: int = 5):
    """
    Versão ultra-resiliente com:
    - 5 tentativas máximas
    - Backoff exponencial (1s, 2s, 4s, 8s, 16s)
    - Headers realistas
    - Tratamento robusto de erros
    """
    fx_cfg = fx_cfg or {}
    
    # Headers realistas para parecer um navegador
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
    }
    
    for attempt in range(max_retries):
        try:
            logger.debug(f"Tentativa {attempt + 1}/{max_retries} para: {url[:80]}...")
            
            # Aumenta o timeout gradualmente a cada tentativa
            current_timeout = timeout + (attempt * 5)
            
            # Download com stream para evitar problemas de memória
            r = session.get(url, timeout=current_timeout, stream=True, headers=headers)
            
            if r.status_code != 200:
                logger.warning(f"HTTP {r.status_code} na tentativa {attempt + 1}: {url[:80]}...")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, f"HTTP {r.status_code} após {max_retries} tentativas")
                time.sleep(2 ** attempt)  # Backoff exponencial
                continue
            
            # Verificar tamanho do conteúdo
            content_length = r.headers.get('content-length')
            if content_length and int(content_length) > 30 * 1024 * 1024:  # 30MB max
                logger.warning(f"Arquivo muito grande ({int(content_length)/1024/1024:.1f}MB): {url[:80]}...")
                return (url, False, None, None, None, None, None, "Arquivo muito grande (>30MB)")
            
            # Baixar conteúdo
            raw_bytes = r.content
            if len(raw_bytes) == 0:
                logger.warning(f"Arquivo vazio na tentativa {attempt + 1}: {url[:80]}...")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, "Arquivo vazio")
                time.sleep(2 ** attempt)
                continue
            
            # Tentar abrir a imagem
            try:
                img = Image.open(BytesIO(raw_bytes))
                img.load()  # Forçar carregamento para detectar corrupção
            except Exception as e:
                logger.warning(f"Imagem corrompida na tentativa {attempt + 1}: {url[:80]}... - {e}")
                if attempt == max_retries - 1:
                    return (url, False, None, None, None, None, None, f"Imagem corrompida: {str(e)[:100]}")
                time.sleep(2 ** attempt)
                continue
            
            # Processamento bem-sucedido!
            sha1_hex = _sha1_bytes(raw_bytes)
            im_small_for_hash = img.copy()
            im_small_for_hash.thumbnail((256, 256), Image.LANCZOS)
            dhash_hex = _img_dhash(im_small_for_hash)
            
            quality = medir_qualidade(img)
            img = redimensionar(img, max_w, max_h)
            
            # Aplicar efeitos se necessário
            need_alpha = bool(fx_cfg and (fx_cfg.get("fx_shadow") or fx_cfg.get("fx_round") or fx_cfg.get("fx_border")))
            if need_alpha:
                img_rgba = apply_effects_pipeline(img.convert("RGB"), fx_cfg)
            else:
                img_rgba = img.convert("RGBA")
            
            # Salvar em disco
            if need_alpha:
                buf = BytesIO()
                img_rgba.save(buf, format="PNG", optimize=True)
                if buf.tell() / 1024 <= limite_kb:
                    path = _save_bytes_to_tmp("png", buf.getvalue())
                    w, h = img_rgba.size
                    logger.info(f"✅ Sucesso na tentativa {attempt + 1}: {url[:80]}...")
                    return (url, True, path, (w, h), quality, sha1_hex, dhash_hex, None)
                
                # Fallback para JPEG
                bg = Image.new("RGB", img_rgba.size, (255, 255, 255))
                bg.paste(img_rgba, mask=img_rgba.split()[-1])
                buf = comprimir_jpeg_binsearch(bg, limite_kb)
                path = _save_bytes_to_tmp("jpg", buf.getvalue())
                w, h = bg.size
                logger.info(f"✅ Sucesso (JPEG) na tentativa {attempt + 1}: {url[:80]}...")
                return (url, True, path, (w, h), quality, sha1_hex, dhash_hex, None)
            else:
                buf = comprimir_jpeg_binsearch(img.convert("RGB"), limite_kb)
                path = _save_bytes_to_tmp("jpg", buf.getvalue())
                w, h = img.size
                logger.info(f"✅ Sucesso na tentativa {attempt + 1}: {url[:80]}...")
                return (url, True, path, (w, h), quality, sha1_hex, dhash_hex, None)
                
        except requests.exceptions.Timeout as e:
            logger.warning(f"Timeout na tentativa {attempt + 1}: {url[:80]}...")
            if attempt == max_retries - 1:
                return (url, False, None, None, None, None, None, f"Timeout após {max_retries} tentativas")
            time.sleep(2 ** attempt)
            
        except requests.exceptions.ConnectionError as e:
            logger.warning(f"Erro de conexão na tentativa {attempt + 1}: {url[:80]}... - {e}")
            if attempt == max_retries - 1:
                return (url, False, None, None, None, None, None, f"Erro de conexão: {str(e)[:100]}")
            time.sleep(2 ** attempt)
            
        except Exception as e:
            logger.warning(f"Erro inesperado na tentativa {attempt + 1}: {url[:80]}... - {e}")
            if attempt == max_retries - 1:
                return (url, False, None, None, None, None, None, f"Erro: {str(e)[:100]}")
            time.sleep(2 ** attempt)
        
        finally:
            gc.collect()
    
    return (url, False, None, None, None, None, None, "Todas as tentativas falharam")

# === PARTE 7/10 =====================================================
# ZIP de imagens + PPT com modelo (capa/final)

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

def gerar_ppt_modelo_capa_final(
    template_bytes: bytes,
    items, resultados, titulo,
    max_per_slide, sort_mode, bg_rgb,
    logo_bytes=None, logo_width_in=1.2,
    signature_bytes=None, signature_width_in=None, auto_half_signature=True,
    signature_bottom_margin_in=0.2, signature_right_margin_in=0.2,
    title_font_name="Radikal", title_font_size_pt=18, title_font_bold=True,
    title_font_color_rgb=(0,0,0),
    excluded_urls=None,
    ignore_failed=True
):
    excluded_urls = excluded_urls or set()
    prs = Presentation(BytesIO(template_bytes))
    logger.info(f"Template carregado: {len(prs.slides)} slides")

    final_idx = 1 if len(prs.slides) >= 2 else None
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]

    groups = OrderedDict()
    for loja, endereco, url in items:
        if url in resultados and url not in excluded_urls:
            groups.setdefault(str(loja), []).append((url, resultados[url]))

    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    title_rgb = title_font_color_rgb if title_font_color_rgb else pick_contrast_color(*bg_rgb)
    signature_width = (logo_width_in/2.0) if auto_half_signature else (signature_width_in or 0.6)

    for loja in loja_keys:
        imgs = groups[loja]
        i = 0
        while i < len(imgs):
            if max_per_slide == "Automático":
                _url0, (_loja0, endereco, _file0, (w0, h0), *_r0) = imgs[i]
                per_slide = 3 if is_portrait(w0, h0) else 2
            else:
                endereco = imgs[i][1][1]
                per_slide = int(max_per_slide)

            batch = imgs[i:i+per_slide]
            i += per_slide
            slide = prs.slides.add_slide(blank_layout)
            set_slide_bg(slide, bg_rgb)
            add_title_and_address(slide, loja, endereco, title_rgb,
                                  title_font_name, title_font_size_pt, title_font_bold)
            if logo_bytes:
                add_logo_top_right(slide, prs, logo_bytes, logo_width_in or 1.2)
            if signature_bytes:
                add_signature_bottom_right(
                    slide, prs, signature_bytes, signature_width,
                    bottom_margin_in=signature_bottom_margin_in,
                    right_margin_in=signature_right_margin_in
                )
            slots = get_slots(len(batch), prs)
            for (url, (_loja, _end, file_path, (w_px, h_px), *rest)), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                try:
                    place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in)
                except Exception as e:
                    logger.warning(f"Falha ao inserir imagem no slide ({url}): {e}")

    if final_idx is not None and final_idx < len(prs.slides):
        try:
            move_slide_to_index(prs, final_idx, len(prs.slides)-1)
        except Exception as e:
            logger.warning(f"Não foi possível mover o slide final: {e}")

    out = BytesIO()
    prs.save(out)
    out.seek(0)
    logger.info("PPT com modelo gerado com sucesso.")
    return out

# === PARTE 8/10 =====================================================
# PPT helpers + funções auxiliares

def get_slots(n, prs):
    IMG_TOP = Inches(1.2); CONTENT_W = Inches(11); CONTENT_H = Inches(6); GAP = Inches(0.2)
    start_left = (prs.slide_width - CONTENT_W) / 2
    if n == 1:
        return [(start_left, IMG_TOP, CONTENT_W, CONTENT_H)]
    cols = n
    total_gap = GAP * (cols - 1)
    cell_w = (CONTENT_W - total_gap) / cols
    return [(start_left + c*(cell_w+GAP), IMG_TOP, cell_w, CONTENT_H) for c in range(cols)]

def add_title_and_address(slide, title_text, address_text, title_rgb=(0,0,0),
                          font_name="Radikal", title_font_size_pt=18, title_font_bold=True):
    TITLE_LEFT, TITLE_TOP, TITLE_W = Inches(0.5), Inches(0.2), Inches(12)
    tx = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_W, Inches(1))
    tf = tx.text_frame; tf.clear()
    p = tf.paragraphs[0]; run = p.add_run(); run.text = title_text    f = run.font; f.name = font_name or "Radikal"; f.size = Pt(title_font_size_pt or 18)
    f.bold = bool(title_font_bold); f.color.rgb = RGBColor(*title_rgb)
    p.alignment = 1
    if address_text:
        p2 = tf.add_paragraph()
        run2 = p2.add_run(); run2.text = address_text
        f2 = run2.font; f2.name = font_name or "Radikal"; f2.size = Pt(max(8, (title_font_size_pt or 18) / 2))
        f2.bold = False; f2.color.rgb = RGBColor(*title_rgb)
        p2.alignment = 1

def set_slide_bg(slide, rgb_tuple):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = RGBColor(*rgb_tuple)

def place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in):
    img_w_in = px_to_inches(w_px); img_h_in = px_to_inches(h_px)
    ratio = min(float(max_w_in)/float(img_w_in), float(max_h_in)/float(img_h_in), 1.0)
    final_w = img_w_in * ratio; final_h = img_h_in * ratio
    x = left + (max_w_in - final_w)/2; y = top + (max_h_in - final_h)/2
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
    left = prs.slide_width - Inches(0.5) - Inches(logo_width_in); top = Inches(0.2)
    slide.shapes.add_picture(BytesIO(logo_bytes), left, top, width=Inches(logo_width_in))

def add_signature_bottom_right(slide, prs, signature_bytes: bytes, signature_width_in: float,
                               bottom_margin_in: float = 0.2, right_margin_in: float = 0.2):
    if not signature_bytes: 
        return
    try:
        im = Image.open(BytesIO(signature_bytes)); w_px, h_px = im.size
        ratio = (h_px / float(w_px)) if w_px else 0.4
    except Exception:
        ratio = 0.4
    sig_h_in = signature_width_in * ratio
    left = prs.slide_width - Inches(right_margin_in) - Inches(signature_width_in)
    top  = prs.slide_height - Inches(bottom_margin_in) - Inches(sig_h_in)
    slide.shapes.add_picture(BytesIO(signature_bytes), left, top, width=Inches(signature_width_in))

# === PARTE 9/10 =====================================================
# UI de miniaturas + detecção + reset (simplificado)

def img_to_html_with_border(image: Image.Image, width_px: int, border_px: int, border_color: str):
    im = image.copy()
    im.thumbnail((width_px, width_px))
    buf = BytesIO()
    im.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    style = (
        f"border:{border_px}px solid {border_color};"
        f"border-radius:10px;display:block;max-width:100%;width:{width_px}px;"
    )
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
    
    return {
        "total_urls": total_urls,
        "baixadas": baixadas,
        "lojas": lojas,
        "falhas": falhas,
        "excluidas": len(excluded)
    }

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
    
    # Keys de upload
    for k in ["xlsx_key", "template_key", "logo_key", "sign_key", "download_key", "images_zip_key"]:
        st.session_state[k] = 0
    
    # Expanders
    st.session_state.exp_plan = True
    st.session_state.exp_style = False
    st.session_state.exp_brand = False
    st.session_state.exp_fx = False
    st.session_state.exp_perf = False
    st.session_state.exp_model = False
    
    # Estado do pipeline
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
# APP principal - VERSÃO ULTRA-RESILIENTE

def main_app():
    # Inicializações
    for k in ["xlsx_key", "template_key", "logo_key", "sign_key", "download_key", "images_zip_key"]:
        if k not in st.session_state:
            st.session_state[k] = 0
    if "exp_plan" not in st.session_state:
        st.session_state.exp_plan = True
        st.session_state.exp_style = False
        st.session_state.exp_brand = False
        st.session_state.exp_fx = False
        st.session_state.exp_perf = False
        st.session_state.exp_model = False
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
    if "preview_bump" not in st.session_state:
        st.session_state.preview_bump = 0
    if "failed_urls" not in st.session_state:
        st.session_state.failed_urls = []
    if "failed_details" not in st.session_state:
        st.session_state.failed_details = []
    if "url_line_map" not in st.session_state:
        st.session_state.url_line_map = {}
    if "quick_generate" not in st.session_state:
        st.session_state.quick_generate = False
    if "ignore_failed" not in st.session_state:
        st.session_state.ignore_failed = True

    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)
        
        st.markdown("---")
        st.subheader("🚀 Configuração Ultra-Resiliente")
        st.caption("Otimizado para Jotform e links com alta taxa de sucesso")
        
        with st.expander("📄 Planilha & Layout", expanded=st.session_state.exp_plan):
            st.caption("Colunas (nomes exatos do cabeçalho):")
            loja_col = st.text_input("🛒 Coluna de LOJA", value="Selecione sua loja", key="loja_col")
            img_col  = st.text_input("🖼️ Coluna de FOTOS", value="Faça o upload das fotos", key="img_col")
            use_address = st.checkbox("➕ Incluir endereço abaixo do nome da loja", value=False, key="use_address")
            address_col = st.text_input("🏠 Coluna de ENDEREÇO", value="Endereço", key="address_col", disabled=not use_address)

            max_per_slide = st.selectbox("📐 Fotos por slide (máx.)", ["Automático", 1, 2, 3], index=0, key="max_per_slide")
            sort_mode = st.selectbox("🔤 Ordenar lojas por", ["Ordem original do Excel", "Nome da loja (A→Z)"], index=0, key="sort_mode")

        with st.expander("🎨 Aparência do slide", expanded=st.session_state.exp_style):
            bg_hex = st.color_picker("🎨 Cor de fundo", value="#FFFFFF", key="bg_hex")
            st.caption("Título do slide")
            title_font_name = st.text_input("Fonte do título", value="Radikal", key="title_font_name")
            title_font_size_pt = st.slider("Tamanho (pt)", 8, 48, 18, 1, key="title_font_size_pt")
            title_font_bold = st.checkbox("Negrito", value=True, key="title_font_bold")
            title_font_color = st.color_picker("Cor da fonte", value="#000000", key="title_font_color")

        with st.expander("🏷️ Logo & ✍️ Assinatura", expanded=st.session_state.exp_brand):
            st.caption("Logo (canto superior direito)")
            logo_file = st.file_uploader("Logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key=f"logo_uploader_{st.session_state.logo_key}")
            if "logo_bytes" not in st.session_state: st.session_state.logo_bytes = None
            if logo_file is not None: st.session_state.logo_bytes = logo_file.getvalue()
            logo_width_in = st.slider("Largura do LOGO (pol)", 0.5, 3.0, 1.2, 0.1, key="logo_width_in")

            st.markdown("---")
            st.caption("Assinatura (canto inferior direito)")
            signature_file = st.file_uploader("Assinatura (PNG/JPG)", type=["png", "jpg", "jpeg"], key=f"signature_uploader_{st.session_state.sign_key}")
            if "signature_bytes" not in st.session_state: st.session_state.signature_bytes = None
            if signature_file is not None: st.session_state.signature_bytes = signature_file.getvalue()

            auto_half_signature = st.checkbox("Usar 1/2 do tamanho do logo (recomendado)", value=True, key="auto_half_signature")
            derived_default_sig = (st.session_state.get("logo_width_in", 1.2) / 2.0)
            if not auto_half_signature:
                signature_width_in = st.slider("Largura da assinatura (pol)", 0.3, 2.0, float(derived_default_sig), 0.05, key="signature_width_in")
            else:
                if "signature_width_in" not in st.session_state:
                    st.session_state.signature_width_in = float(derived_default_sig)

            st.caption("Posição da assinatura")
            signature_right_margin_in = st.slider("Margem direita (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_right_margin")
            signature_bottom_margin_in = st.slider("Margem inferior (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_bottom_margin")

        with st.expander("📑 Modelo (capa + final)", expanded=st.session_state.exp_model):
            use_template = st.checkbox("Usar modelo (Capa + Final)", value=False, key="use_template")
            template_file = None
            if use_template:
                template_file = st.file_uploader("Suba o PPTX com 2 slides", type=["pptx"], key=f"template_pptx_{st.session_state.template_key}")

        with st.expander("🎭 Efeitos nas fotos", expanded=st.session_state.exp_fx):
            st.caption("Ative efeitos opcionais.")
            fx_shadow = st.checkbox("Sombra projetada", value=False, key="fx_shadow")
            shadow_blur = st.slider("Intensidade da sombra (blur)", 0, 30, 10, 1, key="fx_shadow_blur", disabled=not fx_shadow)
            shadow_offset = st.slider("Deslocamento da sombra (px)", 0, 30, 8, 1, key="fx_shadow_offset", disabled=not fx_shadow)
            shadow_opacity = st.slider("Opacidade da sombra (%)", 10, 100, 40, 5, key="fx_shadow_opac", disabled=not fx_shadow)

            fx_round = st.checkbox("Borda arredondada", value=False, key="fx_round")
            round_radius = st.slider("Raio dos cantos (px)", 0, 60, 20, 2, key="fx_round_radius", disabled=not fx_round)

            fx_border = st.checkbox("Borda colorida", value=False, key="fx_border")
            border_color_hex = st.color_picker("Cor da borda", value="#DDDDDD", key="fx_border_color", disabled=not fx_border)
            border_width = st.slider("Espessura da borda (px)", 1, 30, 6, 1, key="fx_border_width", disabled=not fx_border)

        with st.expander("⚡ Performance & Qualidade", expanded=st.session_state.exp_perf):
            # Configurações ULTRA-RESILIENTES (valores otimizados)
            st.info("⚡ Configurações otimizadas para máximo de sucesso no download")
            
            max_retries = st.slider("🔄 Tentativas por imagem", 3, 10, 5, 1, key="max_retries",
                                    help="Número de tentativas para cada imagem. Mais tentativas = maior chance de sucesso")
            
            timeout_base = st.slider("⏱️ Timeout base (segundos)", 15, 90, 45, 5, key="timeout_base",
                                     help="Tempo máximo de espera por download. Aumenta automaticamente a cada tentativa")
            
            max_workers = st.slider("📡 Downloads simultâneos", 1, 8, 3, 1, key="max_workers",
                                    help="Menos downloads simultâneos = mais estabilidade")
            
            st.markdown("---")
            st.caption("Redimensionamento / compressão")
            target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10, key="target_w")
            target_h = st.number_input("Altura máx (px)",  360, 4096, 720, 10, key="target_h")
            limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 3000, 800, 10, key="limite_kb",
                                        help="Aumentei para 800KB para preservar qualidade")
            
            st.markdown("---")
            st.caption("Critérios de qualidade (baixos para não excluir imagens válidas)")
            min_mp = st.slider("Megapixels mínimos", 0.05, 5.0, 0.3, 0.05, key="min_megapixels",
                               help="Valor baixo para não descartar imagens pequenas mas válidas")
            min_blur = st.slider("Limiar de nitidez (blur score)", 5, 300, 25, 5, key="min_blur_score",
                                 help="Valor baixo para não descartar fotos um pouco borradas")
            
            st.markdown("---")
            st.caption("☕ Miniaturas (pré-visualização)")
            thumb_px = st.slider("Tamanho das miniaturas (px)", 120, 400, 220, 10, key="thumb_px")
            thumbs_per_row = st.slider("Miniaturas por linha", 2, 8, 4, 1, key="thumbs_per_row")

            # Comportamento em caso de falha
            ignore_failed = st.checkbox(
                "⚠️ Ignorar falhas e continuar gerando", 
                value=True, 
                key="ignore_failed",
                help="Quando ativado, as fotos que falharam serão puladas e o book gerado com as que funcionaram. Desative para parar em caso de qualquer falha."
            )

    # Topo da página
    top_l, top_m, top_r = st.columns([5,1,1])
    with top_l:
        current_step = 1
        if st.session_state.get("preview_mode") and not st.session_state.get("generated"):
            current_step = 2
        if st.session_state.get("generated") or st.session_state.get("images_zip_bytes"):
            current_step = 3
        st.title("📸 Gerador de Book Ultra-Resiliente")
        render_steps(current_step)
        st.caption("Versão otimizada com retry automático (5 tentativas) e backoff exponencial para máxima taxa de sucesso")
    with top_m:
        if st.button("Resetar", key="reset_btn", use_container_width=True, type="secondary"):
            st.session_state.xlsx_key += 1
            st.session_state.template_key += 1
            st.session_state.logo_key += 1
            st.session_state.sign_key += 1
            st.session_state.download_key += 1
            st.session_state.images_zip_key += 1
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
                btn_preview = st.button("👁️ Visualização Rápida", key="btn_preview", use_container_width=True, type="secondary")
            with col2:
                btn_generate_direct = st.button("🚀 Gerar PPT Direto (Ultra-Resiliente)", key="btn_generate_direct", use_container_width=True, type="primary")

            if btn_preview or btn_generate_direct:
                if not up:
                    st.warning("Envie a planilha primeiro.")
                else:
                    try:
                        df = pd.read_excel(up)
                    except Exception as e:
                        st.error(f"Não consegui ler o Excel: {e}")
                        st.stop()

                    loja_col = st.session_state["loja_col"]
                    img_col  = st.session_state["img_col"]
                    use_address = st.session_state.get("use_address", False)
                    address_col = st.session_state.get("address_col", "Endereço")

                    required_cols = [loja_col, img_col] + ([address_col] if use_address else [])
                    missing = [c for c in required_cols if c not in df.columns]
                    if missing:
                        st.error(f"Colunas não encontradas: {missing}")
                        st.stop()

                    items = []
                    url_line_map = {}
                    for ridx, row in df.iterrows():
                        loja = str(row[loja_col]).strip()
                        endereco = str(row[address_col]).strip() if use_address else ""
                        line_no = (int(ridx) + 2) if isinstance(ridx, (int, np.integer)) else "?"
                        for url in extrair_links(row.get(img_col, "")):
                            if url.startswith("http"):
                                items.append((loja, endereco, url))
                                url_line_map.setdefault(url, line_no)

                    # Remove duplicatas exatas
                    seen, uniq = set(), []
                    for loja, endereco, url in items:
                        if url not in seen:
                            seen.add(url)
                            uniq.append((loja, endereco, url))
                    items = uniq

                    if st.session_state["sort_mode"] == "Nome da loja (A→Z)":
                        grouped_tmp = OrderedDict()
                        for loja, end, url in items:
                            grouped_tmp.setdefault(loja, []).append((end, url))
                        items = [
                            (loja, end, url)
                            for loja in sorted(grouped_tmp.keys())
                            for (end, url) in grouped_tmp[loja]
                        ]

                    total = len(items)
                    if total == 0:
                        st.warning("Nenhuma URL de imagem encontrada.")
                        st.stop()

                    st.info(f"📥 Processando **{total}** imagem(ns) com **{st.session_state['max_retries']}** tentativas cada...")
                    
                    session = requests.Session()
                    adapter = requests.adapters.HTTPAdapter(
                        pool_connections=st.session_state["max_workers"],
                        pool_maxsize=st.session_state["max_workers"],
                        max_retries=3
                    )
                    session.mount("http://", adapter)
                    session.mount("https://", adapter)

                    fx_cfg = {
                        "fx_shadow": st.session_state.get("fx_shadow", False),
                        "fx_shadow_blur": st.session_state.get("fx_shadow_blur", 10),
                        "fx_shadow_offset": st.session_state.get("fx_shadow_offset", 8),
                        "fx_shadow_opac": st.session_state.get("fx_shadow_opac", 40),
                        "fx_round": st.session_state.get("fx_round", False),
                        "fx_round_radius": st.session_state.get("fx_round_radius", 20),
                        "fx_border": st.session_state.get("fx_border", False),
                        "fx_border_color": st.session_state.get("fx_border_color", "#DDDDDD"),
                        "fx_border_width": st.session_state.get("fx_border_width", 6),
                    }

                    prog = st.progress(0)
                    status = st.empty()
                    resultados, falhas, done = {}, 0, 0
                    failed_urls = []
                    failed_details = []

                    with ThreadPoolExecutor(max_workers=st.session_state["max_workers"]) as ex:
                        futures = {
                            ex.submit(
                                baixar_processar_resiliente, 
                                session, url,
                                st.session_state["target_w"], st.session_state["target_h"],
                                st.session_state["limite_kb"], st.session_state["timeout_base"],
                                fx_cfg,
                                st.session_state["max_retries"]
                            ): (loja, endereco, url, url_line_map.get(url, "?"))
                            for loja, endereco, url in items
                        }
                        
                        for fut in as_completed(futures):
                            loja, endereco, url, line_no = futures[fut]
                            try:
                                res = fut.result()
                            except Exception as e:
                                logger.error(f"Erro fatal em {url}: {e}")
                                res = (url, False, None, None, None, None, None, f"Erro fatal: {e}")
                            
                            if res and len(res) >= 2 and res[1] is True:
                                url_key = res[0]
                                file_path = res[2] if len(res) > 2 else None
                                wh = res[3] if len(res) > 3 else (0, 0)
                                quality = res[4] if len(res) > 4 else {}
                                sha1_hex = res[5] if len(res) > 5 else ""
                                dhash_hex = res[6] if len(res) > 6 else ""
                                
                                if file_path and wh:
                                    resultados[url_key] = (loja, endereco, file_path, wh, quality, sha1_hex, dhash_hex)
                                else:
                                    falhas += 1
                                    failed_urls.append(url)
                                    error_msg = res[7] if len(res) > 7 else "Erro desconhecido"
                                    failed_details.append({"url": url, "loja": loja, "linha": line_no, "erro": error_msg})
                            else:
                                falhas += 1
                                failed_urls.append(url)
                                error_msg = res[7] if len(res) > 7 else "Erro desconhecido"
                                failed_details.append({"url": url, "loja": loja, "linha": line_no, "erro": error_msg})
                            
                            done += 1
                            prog.progress(int(done * 100 / total))
                            status.write(f"Processadas {done}/{total} imagens... ✅ Sucesso: {len(resultados)} | ❌ Falhas: {falhas}")

                    status.write(f"✅ Processamento concluído! Sucesso: {len(resultados)} | Falhas: {falhas}")
                    
                    st.session_state.failed_urls = failed_urls
                    st.session_state.failed_details = failed_details
                    st.session_state.url_line_map = url_line_map

                    if falhas > 0:
                        if not st.session_state.get("ignore_failed", True):
                            st.error(f"❌ {falhas} imagem(ns) falharam. O processo foi interrompido conforme sua configuração.")
                            st.stop()
                        else:
                            st.warning(f"⚠️ {falhas} imagem(ns) falharam, mas continuando com {len(resultados)} bem-sucedidas.")
                            
                            if failed_details:
                                st.markdown(f"**📋 Detalhes das {len(failed_details)} falhas:**")
                                df_failed = pd.DataFrame(failed_details)
                                if 'linha' in df_failed.columns:
                                    try:
                                        df_failed['linha_num'] = pd.to_numeric(df_failed['linha'], errors='coerce')
                                        df_failed = df_failed.sort_values('linha_num')
                                        df_failed = df_failed.drop(columns=['linha_num'])
                                    except:
                                        df_failed = df_failed.sort_values('linha')
                                st.dataframe(df_failed, hide_index=True, use_container_width=True)

                    low_q, dups = detectar_problemas(
                        resultados,
                        st.session_state["min_megapixels"],
                        st.session_state["min_blur_score"]
                    )
                    st.session_state.low_quality_urls = list(low_q)
                    st.session_state.duplicate_urls = list(dups)

                    st.session_state.pipeline = {
                        "items": items, "resultados": resultados, "falhas": falhas,
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
                            "auto_half_signature": st.session_state.get("auto_half_signature", True),
                            "signature_width_in": st.session_state.get("signature_width_in", st.session_state["logo_width_in"]/2.0),
                            "signature_right_margin_in": st.session_state.get("sig_right_margin", 0.20),
                            "signature_bottom_margin_in": st.session_state.get("sig_bottom_margin", 0.20),
                            "thumb_px": st.session_state["thumb_px"],
                            "thumbs_per_row": st.session_state["thumbs_per_row"],
                            "effects": fx_cfg,
                            "use_template": st.session_state.get("use_template", False),
                            "template_bytes": template_file.getvalue() if (st.session_state.get("use_template") and template_file) else None,
                            "low_quality_urls": list(low_q),
                            "duplicate_urls": list(dups),
                            "ignore_failed": st.session_state.get("ignore_failed", True),
                        }
                    }

                    if btn_preview:
                        st.session_state.preview_mode = True
                        st.session_state.quick_generate = False
                    else:
                        st.session_state.preview_mode = False  
                        st.session_state.quick_generate = True
                        st.session_state.generated = False

                    st.session_state.ppt_bytes = None
                    st.session_state.images_zip_bytes = None
                    
                    st.rerun()
        
        # Resto da UI (pré-visualização e geração) - similar à versão original mas simplificada
        if st.session_state.pipeline:
            if st.session_state.preview_mode and not st.session_state.quick_generate:
                st.markdown("---")
                st.subheader("2. Pré-visualização")
                p = st.session_state.pipeline
                
                stats = render_summary(p["items"], p["resultados"], st.session_state.excluded_urls, st.session_state.get("failed_details", []))
                
                if st.session_state.get("failed_details") and st.session_state.get("ignore_failed", True):
                    st.info(f"⚠️ {len(st.session_state.failed_details)} imagem(ns) falharam. O book será gerado com {stats['baixadas']} imagens.")
                
                # Preview simplificado
                st.info("👆 Selecione as imagens que deseja excluir e depois clique em 'Gerar PPT'")
                
                if st.button("🔄 Gerar PPT Agora", use_container_width=True, type="primary"):
                    st.session_state.quick_generate = True
                    st.session_state.preview_mode = False
                    st.rerun()
            
            # Geração e Download
            if st.session_state.quick_generate or st.session_state.generated:
                st.markdown("---")
                st.subheader("3. Gerar / Exportar")
                
                cfg = st.session_state.pipeline["settings"]
                items = st.session_state.pipeline["items"]
                resultados = st.session_state.pipeline["resultados"]
                
                stats = render_summary(items, resultados, st.session_state.excluded_urls, st.session_state.get("failed_details", []))
                
                if st.session_state.quick_generate and not st.session_state.get("ppt_bytes"):
                    with st.spinner("🚀 Gerando PPT com modo Ultra-Resiliente..."):
                        try:
                            titulo = (st.session_state.output_filename or "Apresentacao").strip()
                            use_template = cfg.get("use_template", False)
                            template_bytes = cfg.get("template_bytes")

                            if use_template and template_bytes:
                                ppt_bytes = gerar_ppt_modelo_capa_final(
                                    template_bytes=template_bytes,
                                    items=items, resultados=resultados, titulo=titulo,
                                    max_per_slide=cfg["max_per_slide"], sort_mode=cfg["sort_mode"],
                                    bg_rgb=cfg["bg_rgb"],
                                    logo_bytes=cfg["logo_bytes"], logo_width_in=cfg["logo_width_in"],
                                    signature_bytes=cfg["signature_bytes"],
                                    signature_width_in=cfg.get("signature_width_in"),
                                    auto_half_signature=cfg.get("auto_half_signature", True),
                                    signature_bottom_margin_in=cfg["signature_bottom_margin_in"],
                                    signature_right_margin_in=cfg["signature_right_margin_in"],
                                    title_font_name=cfg["title_font_name"],
                                    title_font_size_pt=cfg["title_font_size_pt"],
                                    title_font_bold=cfg["title_font_bold"],
                                    title_font_color_rgb=cfg.get("title_font_color_rgb", (0,0,0)),
                                    excluded_urls=st.session_state.excluded_urls,
                                    ignore_failed=cfg.get("ignore_failed", True)
                                )
                            else:
                                prs = Presentation()
                                prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
                                blank = prs.slide_layouts[6]
                                title_rgb = cfg.get("title_font_color_rgb", pick_contrast_color(*cfg["bg_rgb"]))
                                signature_width = (cfg["logo_width_in"]/2.0) if cfg.get("auto_half_signature", True) else (cfg.get("signature_width_in") or 0.6)

                                groups = OrderedDict()
                                for loja, endereco, url in items:
                                    if url in resultados and url not in st.session_state.excluded_urls:
                                        groups.setdefault(str(loja), []).append((url, resultados[url]))

                                loja_keys = list(groups.keys())
                                if cfg["sort_mode"] == "Nome da loja (A→Z)":
                                    loja_keys.sort(key=lambda s: str(s).strip().casefold())

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
                                        if cfg["logo_bytes"]:
                                            add_logo_top_right(slide, prs, cfg["logo_bytes"], cfg["logo_width_in"])
                                        if cfg["signature_bytes"]:
                                            add_signature_bottom_right(slide, prs, cfg["signature_bytes"], signature_width,
                                                bottom_margin_in=cfg["signature_bottom_margin_in"],
                                                right_margin_in=cfg["signature_right_margin_in"])
                                        slots = get_slots(len(batch), prs)
                                        for (_, (_, _, file_path, (w_px, h_px), *_)), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                                            place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in)

                                out = BytesIO()
                                prs.save(out)
                                out.seek(0)
                                ppt_bytes = out

                            st.session_state.ppt_bytes = ppt_bytes
                            st.session_state.generated = True
                            st.session_state.quick_generate = False
                            st.success(f"✅ PPT gerado com sucesso! Total: {len(resultados)} imagens em {len(groups)} lojas.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Falha ao gerar PPT: {e}")
                
                # Download buttons
                st.markdown("---")
                st.subheader("📥 Download")
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.session_state.get("ppt_bytes"):
                        st.download_button(
                            "⬇️ Baixar PPT",
                            data=st.session_state.ppt_bytes,
                            file_name=f"{st.session_state.output_filename}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentation.presentation",
                            use_container_width=True
                        )
                
                with col2:
                    zip_bytes = montar_zip_imagens(items, resultados, st.session_state.excluded_urls)
                    st.download_button(
                        "⬇️ Baixar ZIP das Imagens",
                        data=zip_bytes,
                        file_name=f"{st.session_state.output_filename}_imagens.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
        
        if not st.session_state.pipeline:
            st.info("📤 **Faça o upload da planilha para começar**")
            st.markdown("""
            ### 🚀 Modo Ultra-Resiliente - Características:
            - ✅ **5 tentativas** automáticas por imagem
            - ✅ **Backoff exponencial** (espera progressiva entre tentativas)
            - ✅ **Headers realistas** (simula navegador)
            - ✅ **Timeout adaptativo** (aumenta a cada tentativa)
            - ✅ **Otimizado para Jotform** e servidores lentos
            """)

# -------------------------------------------------------------------
# ROTEAMENTO FINAL
# -------------------------------------------------------------------
if not st.session_state.auth:
    do_login()
else:
    main_app()
