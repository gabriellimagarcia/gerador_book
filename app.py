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

# ===============================
# UX: estilo geral e barra fixa
# ===============================
BASE_CSS = """
<style>
/* barra de etapas */
.steps {display:flex; gap:8px; align-items:center; margin: 1rem 0 0.5rem;}
.step {padding:8px 12px; border-radius:999px; border:1px solid #E5E7EB; color:#111; background:#fff; font-weight:600;}
.step.active {background:#FF7A00; color:#fff; border-color:#FF7A00;}
.step.sep {opacity:0.4}

/* barra fixa de ações no rodapé */
.action-bar {
  position: fixed; left: 0; right: 0; bottom: 0;
  background: rgba(255,255,255,0.92);
  border-top: 1px solid #eee; backdrop-filter: blur(6px);
  padding: 10px 16px; z-index: 999;
}
.dark .action-bar { background: rgba(14,17,23,0.9); border-top-color:#222; }

.action-inner { max-width: 1200px; margin: 0 auto; display:flex; gap:12px; align-items:center; justify-content:flex-end; }

/* cards da pré-visualização */
.img-card {
  border:1px solid #DDD; border-radius:10px; padding:10px; transition: all .15s ease;
  background:#fff; display:flex; flex-direction:column; gap:8px; height:100%;
}
.dark .img-card { background:#0f131a; border-color:#232a36; }
.img-card:hover { box-shadow: 0 8px 20px rgba(0,0,0,0.06); transform: translateY(-2px); }

/* badge com números */
.badge {
  display:inline-block; padding:3px 8px; border-radius:999px; background:#F3F4F6; color:#111; font-size:12px; font-weight:600;
}
.dark .badge { background:#1f2635; color:#eaeaea; }

/* cabeçalho de grupo (loja) */
.group-head {
  display:flex; align-items:center; justify-content:space-between;
  padding:8px 12px; border-radius:8px; background:#F8FAFC; border:1px solid #EEF2F7; margin: 10px 0 6px;
}
.dark .group-head { background:#10151f; border-color:#1b2130; }

/* botão SAIR (dentro da área .logout-zone) */
.logout-zone .stButton > button {
  background: #E53935 !important;
  color: #fff !important;
  border: none !important;
  border-radius: 10px !important;
}
.logout-zone .stButton > button:hover {
  background: #C62828 !important;
  color: #fff !important;
}

/* marcar corpo como dark/light para CSS global */
body:has(.stApp[data-theme="dark"]) .action-bar { background: rgba(14,17,23,0.9); border-top-color:#222; }
</style>
"""
st.markdown(BASE_CSS, unsafe_allow_html=True)

# ===== Tema (claro predominante, acento laranja) =====
def apply_theme(dark: bool):
    ORANGE = "#FF7A00"; ORANGE_HOVER = "#E66E00"; BLACK = "#111111"; GRAY_BG = "#f6f6f7"
    if dark:
        palette = f"""
        <style>
        :root {{
            --accent: {ORANGE}; --accent-hover: {ORANGE_HOVER};
            --text: #f5f5f5; --bg: #0e1117; --panel: #11151c;
        }}
        .stApp {{ background-color: var(--bg); color: var(--text); }}
        section[data-testid="stSidebar"] > div {{ background: var(--panel); border-right: 1px solid #1b212c; }}
        h1,h2,h3,h4,h5,h6 {{ color: var(--text); }}
        .stButton > button, .stDownloadButton > button {{ background: var(--accent); color: #fff; border-radius:10px; border:none; }}
        .stButton > button:hover, .stDownloadButton > button:hover {{ background: var(--accent-hover); color:#fff; }}
        .stProgress > div > div {{ background-color: var(--accent); }}
        </style>
        """
    else:
        palette = f"""
        <style>
        :root {{
            --accent: {ORANGE}; --accent-hover: {ORANGE_HOVER};
            --text: {BLACK}; --bg: #ffffff; --panel: {GRAY_BG};
        }}
        .stApp {{ background-color: var(--bg); color: var(--text); }}
        section[data-testid="stSidebar"] > div {{ background: var(--panel); border-right: 1px solid #ececec; }}
        h1,h2,h3,h4,h5,h6 {{ color: var(--text); }}
        .stButton > button, .stDownloadButton > button {{ background: var(--accent); color: #fff; border-radius:10px; border:none; }}
        .stButton > button:hover, .stDownloadButton > button:hover {{ background: var(--accent-hover); color:#fff; }}
        .stProgress > div > div {{ background-color: var(--accent); }}
        </style>
        """
    st.markdown(palette, unsafe_allow_html=True)
    st.write(f"""<script>
      const root = window.parent.document.querySelector('.stApp');
      if (root) root.setAttribute('data-theme', '{'dark' if dark else 'light'}');
    </script>""", unsafe_allow_html=True)

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# ===== Login =====
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
    st.title("🔐 Acesso")
    st.write("Entre para usar o gerador de book.")
    with st.form("login_form", clear_on_submit=False):
        email = st.text_input("E-mail", placeholder="seu.email@mkthouse.com.br")
        pwd = st.text_input("Senha", type="password", placeholder="••••••••")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        email_norm = (email or "").strip().lower()
        if email_norm in ALLOWED_USERS and pwd == ALLOWED_USERS[email_norm]:
            st.session_state.auth = True
            st.session_state.user_email = email_norm
            st.toast("Login ok ✅", icon="✅")
            st.rerun()
        else:
            st.error("Credenciais inválidas.")

if "auth" not in st.session_state:
    st.session_state.auth = False

# ===== Utils =====
URL_RE = re.compile(r'https?://\S+')
def extrair_links(celula):
    if pd.isna(celula): return []
    t = str(celula).replace(",", " ").replace("(", " ").replace(")", " ").replace('"', " ").replace("'", " ")
    return [u.rstrip(").,") for u in URL_RE.findall(t)]

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
        best = BytesIO(); img.save(best, "JPEG", quality=35, optimize=True, progressive=True, subsampling=2)
    best.seek(0); return best

def baixar_processar(session, url: str, max_w: int, max_h: int, limite_kb: int, timeout: int):
    try:
        r = session.get(url, timeout=timeout, stream=True)
        if r.status_code != 200: return (url, False, None, None)
        img = Image.open(BytesIO(r.content))
        img = redimensionar(img, max_w, max_h)
        buf = comprimir_jpeg_binsearch(img, limite_kb)
        w, h = Image.open(buf).size
        buf.seek(0)
        return (url, True, buf, (w, h))
    except Exception:
        return (url, False, None, None)

def px_to_inches(px): return Inches(px / 96.0)
def hex_to_rgb(hex_str: str):
    s = hex_str.strip().lstrip("#")
    if len(s) == 3: s = "".join([c*2 for c in s])
    return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)

def pick_contrast_color(r, g, b):
    brightness = (r*299 + g*587 + b*114) / 1000
    return (0,0,0) if brightness > 128 else (255,255,255)

# Layout slots p/ 1–3 por slide
def get_slots(n, prs):
    IMG_TOP = Inches(1.2); CONTENT_W = Inches(11); CONTENT_H = Inches(6); GAP = Inches(0.2)
    start_left = (prs.slide_width - CONTENT_W) / 2
    if n == 1:
        return [(start_left, IMG_TOP, CONTENT_W, CONTENT_H)]
    cols = n
    total_gap = GAP * (cols - 1)
    cell_w = (CONTENT_W - total_gap) / cols
    return [(start_left + c*(cell_w+GAP), IMG_TOP, cell_w, CONTENT_H) for c in range(cols)]

def add_title(slide, text, title_rgb=(0,0,0), font_name="Radikal", font_size_pt=18, font_bold=True):
    TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H = Inches(0.5), Inches(0.2), Inches(12), Inches(1)
    tx = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H)
    tf = tx.text_frame; tf.clear()
    p = tf.paragraphs[0]; run = p.add_run(); run.text = text
    font = run.font
    font.name = font_name or "Radikal"
    font.size = Pt(font_size_pt or 18)
    font.bold = bool(font_bold)
    font.color.rgb = RGBColor(*title_rgb)
    p.alignment = 1

def set_slide_bg(slide, rgb_tuple):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = RGBColor(*rgb_tuple)

def place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in):
    img_w_in = px_to_inches(w_px); img_h_in = px_to_inches(h_px)
    ratio = min(float(max_w_in)/float(img_w_in), float(max_h_in)/float(img_h_in), 1.0)
    final_w = img_w_in * ratio; final_h = img_h_in * ratio
    x = left + (max_w_in - final_w)/2; y = top + (max_h_in - final_h)/2
    buf.seek(0)
    slide.shapes.add_picture(buf, x, y, width=final_w, height=final_h)

def add_logo_top_right(slide, prs, logo_bytes: bytes, logo_width_in: float):
    if not logo_bytes: return
    left = prs.slide_width - Inches(0.5) - Inches(logo_width_in); top = Inches(0.2)
    slide.shapes.add_picture(BytesIO(logo_bytes), left, top, width=Inches(logo_width_in))

def add_signature_bottom_right(slide, prs, signature_bytes: bytes, signature_width_in: float,
                               bottom_margin_in: float = 0.2, right_margin_in: float = 0.2):
    if not signature_bytes: return
    try:
        im = Image.open(BytesIO(signature_bytes)); w_px, h_px = im.size
        ratio = (h_px / float(w_px)) if w_px else 0.4
    except Exception:
        ratio = 0.4
    sig_h_in = signature_width_in * ratio
    left = prs.slide_width - Inches(right_margin_in) - Inches(signature_width_in)
    top  = prs.slide_height - Inches(bottom_margin_in) - Inches(sig_h_in)
    slide.shapes.add_picture(BytesIO(signature_bytes), left, top, width=Inches(signature_width_in))

def gerar_ppt(items, resultados, titulo, max_per_slide, sort_mode, bg_rgb,
              logo_bytes=None, logo_width_in=1.2,
              signature_bytes=None, signature_width_in=None, auto_half_signature=True,
              signature_bottom_margin_in=0.2, signature_right_margin_in=0.2,
              title_font_name="Radikal", title_font_size_pt=18, title_font_bold=True,
              excluded_urls=None):
    excluded_urls = excluded_urls or set()
    prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    blank = prs.slide_layouts[6]

    groups = OrderedDict()
    for loja, url in items:
        if url in resultados and url not in excluded_urls:
            groups.setdefault(str(loja), []).append((url, resultados[url]))

    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip() == "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    title_rgb = pick_contrast_color(*bg_rgb)
    signature_width = (logo_width_in/2.0) if auto_half_signature else (signature_width_in or 0.6)

    for loja in loja_keys:
        imgs = groups[loja]
        for i in range(0, len(imgs), max_per_slide):
            batch = imgs[i:i+max_per_slide]
            slide = prs.slides.add_slide(blank)
            set_slide_bg(slide, bg_rgb)
            add_title(slide, loja, title_rgb, title_font_name, title_font_size_pt, title_font_bold)

            if logo_bytes: add_logo_top_right(slide, prs, logo_bytes, logo_width_in or 1.2)
            if signature_bytes:
                add_signature_bottom_right(slide, prs, signature_bytes, signature_width,
                                           bottom_margin_in=signature_bottom_margin_in,
                                           right_margin_in=signature_right_margin_in)

            slots = get_slots(len(batch), prs)
            for (url, (_loja, buf, (w_px, h_px))), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in)

    out = BytesIO(); prs.save(out); out.seek(0); return out

# --- preview helpers ---
def img_to_html_with_border(image: Image.Image, width_px: int, border_px: int, border_color: str):
    im = image.copy(); im.thumbnail((width_px, width_px))
    buf = BytesIO(); im.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    style = f"border:{border_px}px solid {border_color};border-radius:10px;display:block;max-width:100%;width:{width_px}px;"
    return f'<img src="data:image/png;base64,{b64}" style="{style}" />'

def render_steps(current: int):
    labels = ["Upload", "Pré-visualização", "Gerar PPT"]
    html = ['<div class="steps">']
    for i, txt in enumerate(labels, start=1):
        cls = "step active" if i == current else "step"
        html.append(f'<div class="{cls}">{i}. {txt}</div>')
        if i < len(labels): html.append('<span class="step sep">›</span>')
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)

def render_summary(items, resultados, excluded):
    total_urls = len(items)
    baixadas = sum(1 for _, url in items if url in resultados)
    lojas = len({loja for loja, _ in items})
    st.markdown(
        f"**Resumo:** "
        f"<span class='badge'>Lojas: {lojas}</span> "
        f"<span class='badge'>Links: {total_urls}</span> "
        f"<span class='badge'>Baixadas: {baixadas}</span> "
        f"<span class='badge'>Excluídas: {len(excluded)}</span>",
        unsafe_allow_html=True
    )

# ========= PRÉ-VISUALIZAÇÃO COM EXPANDERS =========
def render_preview(items, resultados, max_per_slide, sort_mode, thumb_px: int, thumbs_per_row: int):
    """
    Pré-visualização: miniaturas em grade, com expandir/recolher por loja.
    Mantém seleção em st.session_state.excluded_urls e estado dos grupos em st.session_state.expanded_groups.
    """
    if "excluded_urls" not in st.session_state:
        st.session_state.excluded_urls = set()
    if "expanded_groups" not in st.session_state:
        st.session_state.expanded_groups = {}

    excluded = st.session_state.excluded_urls
    expanded_groups = st.session_state.expanded_groups

    # Agrupa por loja
    groups = OrderedDict()
    for loja, url in items:
        if url in resultados:
            groups.setdefault(str(loja), []).append((url, resultados[url]))

    # Ordenação de lojas
    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip() == "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    # Toolbar topo
    c1, c2, c3, c4 = st.columns([1,1,1,1])
    with c1:
        if st.button("🧹 Limpar todas as exclusões", type="secondary", use_container_width=True):
            excluded.clear()
            st.toast("Exclusões limpas", icon="🧽")
            st.rerun()
    with c2:
        if st.button("🔁 Inverter seleção", type="secondary", use_container_width=True):
            all_urls = {url for _, v in groups.items() for (url, _) in v}
            st.session_state.excluded_urls = all_urls - excluded
            st.toast("Seleção invertida", icon="🔁")
            st.rerun()
    with c3:
        if st.button("➕ Expandir todas", type="secondary", use_container_width=True):
            for loja in loja_keys:
                expanded_groups[loja] = True
            st.rerun()
    with c4:
        if st.button("➖ Recolher todas", type="secondary", use_container_width=True):
            for loja in loja_keys:
                expanded_groups[loja] = False
            st.rerun()

    st.caption(f"Excluídas até agora: **{len(excluded)}**")

    # Render por loja (expander)
    for loja in loja_keys:
        imgs = groups[loja]
        expanded_default = expanded_groups.get(loja, True)  # default expandido
        with st.expander(f"📄 {loja} — {len(imgs)} foto(s)", expanded=expanded_default):
            # Controles rápidos por loja
            lc1, lc2, lc3 = st.columns([1,1,1])
            with lc1:
                if st.button(f"Selecionar todas de {loja}", key=f"sel_all_{hash(loja)}", use_container_width=True):
                    for url, _ in imgs:
                        excluded.add(url)
                    st.rerun()
            with lc2:
                if st.button(f"Limpar seleção de {loja}", key=f"clr_sel_{hash(loja)}", use_container_width=True):
                    for url, _ in imgs:
                        excluded.discard(url)
                    st.rerun()
            with lc3:
                new_state = st.toggle("Manter este grupo expandido", value=expanded_default, key=f"exp_keep_{hash(loja)}",
                                      help="Salva a preferência para esta loja.")
                expanded_groups[loja] = bool(new_state)

            # Grid de miniaturas (thumbs_per_row colunas)
            cols = st.columns(thumbs_per_row)
            col_idx = 0
            for (url, (_loja, buf, (w_px, h_px))) in imgs:
                with cols[col_idx]:
                    try:
                        buf.seek(0)
                        im = Image.open(buf)
                    except Exception:
                        st.warning("Não foi possível pré-visualizar esta imagem.")
                        col_idx = (col_idx + 1) % thumbs_per_row
                        continue

                    is_excluded = url in excluded
                    border_px = 3 if is_excluded else 1
                    border_color = "#E53935" if is_excluded else "#DDDDDD"

                    st.markdown('<div class="img-card">', unsafe_allow_html=True)
                    st.markdown(img_to_html_with_border(im, thumb_px, border_px, border_color), unsafe_allow_html=True)

                    key = "ex_" + hashlib.md5(url.encode("utf-8")).hexdigest()
                    checked = st.checkbox("Excluir esta foto", key=key, value=is_excluded)
                    if checked:
                        excluded.add(url)
                    else:
                        excluded.discard(url)
                    st.markdown('</div>', unsafe_allow_html=True)

                col_idx = (col_idx + 1) % thumbs_per_row

        st.divider()

# ===============================
# App principal (UX melhorado)
# ===============================
def main_app():
    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)

        with st.expander("📄 Planilha & Layout", expanded=True):
            st.caption("Colunas da planilha (nomes exatos do cabeçalho):")
            loja_col = st.text_input("🛒 Coluna de LOJA", value="Selecione sua loja", key="loja_col")
            img_col  = st.text_input("🖼️ Coluna de FOTOS", value="Faça o upload das fotos", key="img_col")
            st.caption("Layout dos slides")
            max_per_slide = st.selectbox("📐 Fotos por slide (máx.)", [1, 2, 3], index=0, key="max_per_slide")
            st.caption("Ordenação")
            sort_mode = st.selectbox("🔤 Ordenar lojas por",
                ["Ordem original do Excel", "Nome da loja (A→Z)"], index=0, key="sort_mode")

        with st.expander("🎨 Aparência do slide", expanded=True):
            bg_hex = st.color_picker("🎨 Cor de fundo", value="#FFFFFF", key="bg_hex")
            st.caption("Título do slide")
            title_font_name = st.text_input("Fonte do título", value="Radikal", key="title_font_name")
            title_font_size_pt = st.slider("Tamanho (pt)", 8, 48, 18, 1, key="title_font_size_pt")
            title_font_bold = st.checkbox("Negrito", value=True, key="title_font_bold")

        with st.expander("🏷️ Logo & ✍️ Assinatura", expanded=True):
            st.caption("Logo (canto superior direito)")
            logo_file = st.file_uploader("Logo (PNG/JPG)", type=["png","jpg","jpeg"], key="logo_uploader")
            if "logo_bytes" not in st.session_state: st.session_state.logo_bytes = None
            if logo_file is not None: st.session_state.logo_bytes = logo_file.getvalue()
            logo_width_in = st.slider("Largura do LOGO (pol)", 0.5, 3.0, 1.2, 0.1, key="logo_width_in")

            st.markdown("---")
            st.caption("Assinatura (canto inferior direito)")
            signature_file = st.file_uploader("Assinatura (PNG/JPG)", type=["png","jpg","jpeg"], key="signature_uploader")
            if "signature_bytes" not in st.session_state: st.session_state.signature_bytes = None
            if signature_file is not None: st.session_state.signature_bytes = signature_file.getvalue()

            auto_half_signature = st.checkbox("Usar 1/2 do tamanho do logo (recomendado)", value=True, key="auto_half_signature")
            derived_default_sig = (st.session_state.get("logo_width_in", 1.2) / 2.0)
            if not auto_half_signature:
                signature_width_in = st.slider("Largura da assinatura (pol)", 0.3, 2.0, float(derived_default_sig), 0.05, key="signature_width_in")
            else:
                if "signature_width_in" not in st.session_state:
                    st.session_state.signature_width_in = float(derived_default_sig)

            st.caption("Posição da assinatura (quanto menor, mais encostada no canto)")
            signature_right_margin_in = st.slider("Margem direita (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_right_margin")
            signature_bottom_margin_in = st.slider("Margem inferior (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_bottom_margin")

        with st.expander("⚡ Performance & Qualidade", expanded=False):
            thumb_px = st.slider("Tamanho das miniaturas (px)", 120, 400, 220, 10, key="thumb_px")
            thumbs_per_row = st.slider("Miniaturas por linha (pré-visualização)", 2, 8, 4, 1, key="thumbs_per_row")
            st.caption("Redimensionamento / compressão")
            target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10, key="target_w")
            target_h = st.number_input("Altura máx (px)",  360, 4096, 720, 10, key="target_h")
            limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 2000, 450, 10, key="limite_kb")
            st.caption("Rede e paralelismo")
            max_workers = st.slider("Trabalhos em paralelo", 2, 32, 12, key="max_workers")
            req_timeout = st.slider("Timeout por download (s)", 5, 60, 15, key="req_timeout")

    # ===== Topo: título e botão SAIR no canto direito (vermelho)
    top_l, top_r = st.columns([6,1])
    with top_l:
        st.title("📸 Gerador de Book")
        st.caption("Monte um PPT com fotos por loja, com compressão e layout automático.")
    with top_r:
        st.markdown('<div class="logout-zone">', unsafe_allow_html=True)
        if st.button("Sair", key="logout_btn", use_container_width=True, type="secondary"):
            st.session_state.clear(); st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ===== Estados
    if "pipeline" not in st.session_state: st.session_state.pipeline = {}
    if "excluded_urls" not in st.session_state: st.session_state.excluded_urls = set()
    if "preview_mode" not in st.session_state: st.session_state.preview_mode = False
    if "expanded_groups" not in st.session_state: st.session_state.expanded_groups = {}

    # ===== Etapas e Upload
    render_steps(1 if not st.session_state.preview_mode else 2)
    up = st.file_uploader("1) Selecione ou arraste a planilha (.xlsx)", type=["xlsx"], key="xlsx_upload")

    # ===== Barra fixa de ações
    st.markdown("""
    <div class="action-bar"><div class="action-inner">
      <div id="btns"></div>
    </div></div>
    """, unsafe_allow_html=True)
    bcol1, bcol2, bcol3 = st.columns([6,1,1])
    with bcol2:
        btn_preview = st.button("🔎 Pré-visualizar", key="btn_preview", use_container_width=True)
    with bcol3:
        btn_generate = st.button("⬇️ Gerar PPT", key="btn_generate", use_container_width=True)

    # ===== Processar planilha para PRÉVIA
    if btn_preview:
        if not up:
            st.warning("Envie a planilha primeiro.")
        else:
            try:
                df = pd.read_excel(up)
            except Exception as e:
                st.error(f"Não consegui ler o Excel: {e}"); st.stop()

            loja_col = st.session_state["loja_col"]; img_col  = st.session_state["img_col"]
            missing = [c for c in [img_col, loja_col] if c not in df.columns]
            if missing: st.error(f"Colunas não encontradas: {missing}"); st.stop()

            items = []
            for _, row in df.iterrows():
                loja = str(row[loja_col]).strip()
                for url in extrair_links(row.get(img_col, "")):
                    if url.startswith("http"): items.append((loja, url))
            seen, uniq = set(), []
            for loja, url in items:
                if url not in seen:
                    seen.add(url); uniq.append((loja, url))
            items = uniq

            if st.session_state["sort_mode"] == "Nome da loja (A→Z)":
                grouped_tmp = OrderedDict()
                for loja, url in items:
                    grouped_tmp.setdefault(loja, []).append(url)
                items = [(loja, u) for loja in sorted(grouped_tmp.keys(), key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold())) for u in grouped_tmp[loja]]

            total = len(items)
            if total == 0: st.warning("Nenhuma URL de imagem encontrada."); st.stop()

            st.info(f"Baixando e processando **{total}** imagem(ns)...")
            session = requests.Session()
            adapter = requests.adapters.HTTPAdapter(pool_connections=st.session_state["max_workers"],
                                                    pool_maxsize=st.session_state["max_workers"], max_retries=2)
            session.mount("http://", adapter); session.mount("https://", adapter)
            session.headers.update({"User-Agent": "Mozilla/5.0 (GeradorBook Streamlit)"})

            prog = st.progress(0); status = st.empty()
            resultados, falhas, done = {}, 0, 0
            with ThreadPoolExecutor(max_workers=st.session_state["max_workers"]) as ex:
                futures = {ex.submit(
                    baixar_processar, session, url,
                    st.session_state["target_w"], st.session_state["target_h"],
                    st.session_state["limite_kb"], st.session_state["req_timeout"]
                ): (loja, url) for loja, url in items}
                for fut in as_completed(futures):
                    loja, url = futures[fut]
                    ok_url, ok, buf, wh = fut.result()
                    if ok: resultados[url] = (loja, buf, wh)
                    else: falhas += 1
                    done += 1; prog.progress(int(done * 100 / total))
                    status.write(f"Processadas {done}/{total} imagens...")

            status.write(f"Concluído. Falhas: {falhas}")
            st.toast("Pré-visualização pronta ✅", icon="✅")

            st.session_state.pipeline = {
                "items": items, "resultados": resultados, "falhas": falhas,
                "settings": {
                    "max_per_slide": st.session_state["max_per_slide"],
                    "sort_mode": st.session_state["sort_mode"],
                    "bg_rgb": hex_to_rgb(st.session_state["bg_hex"]),
                    "title_font_name": st.session_state["title_font_name"],
                    "title_font_size_pt": st.session_state["title_font_size_pt"],
                    "title_font_bold": st.session_state["title_font_bold"],
                    "logo_bytes": st.session_state.logo_bytes,
                    "logo_width_in": st.session_state["logo_width_in"],
                    "signature_bytes": st.session_state.signature_bytes,
                    "auto_half_signature": st.session_state.get("auto_half_signature", True),
                    "signature_width_in": st.session_state.get("signature_width_in", (st.session_state["logo_width_in"]/2.0)),
                    "signature_right_margin_in": st.session_state.get("sig_right_margin", 0.20),
                    "signature_bottom_margin_in": st.session_state.get("sig_bottom_margin", 0.20),
                    "thumb_px": st.session_state["thumb_px"],
                    "thumbs_per_row": st.session_state["thumbs_per_row"],
                }
            }
            st.session_state.preview_mode = True

    # ===== PRÉVIA
    if st.session_state.preview_mode and st.session_state.pipeline:
        render_steps(2)
        p = st.session_state.pipeline
        # Resumo
        total_urls = len(p["items"])
        baixadas = sum(1 for _, url in p["items"] if url in p["resultados"])
        lojas = len({loja for loja, _ in p["items"]})
        st.markdown(
            f"**Resumo:** "
            f"<span class='badge'>Lojas: {lojas}</span> "
            f"<span class='badge'>Links: {total_urls}</span> "
            f"<span class='badge'>Baixadas: {baixadas}</span> "
            f"<span class='badge'>Excluídas: {len(st.session_state.excluded_urls)}</span>",
            unsafe_allow_html=True
        )
        # Render com expanders
        render_preview(
            p["items"], p["resultados"],
            p["settings"]["max_per_slide"],
            p["settings"]["sort_mode"],
            p["settings"]["thumb_px"],
            p["settings"]["thumbs_per_row"]
        )
        st.info("Marque **Excluir esta foto** nas imagens que não devem ir para o PPT. Depois clique em **Gerar PPT** na barra inferior.")

    # ===== GERAR
    if btn_generate:
        if not st.session_state.pipeline:
            st.warning("Clique em **Pré-visualizar** antes de gerar.")
        else:
            render_steps(3)
            p = st.session_state.pipeline
            items = p["items"]; resultados = p["resultados"]; cfg = p["settings"]

            titulo = "Apresentacao_Relatorio_Compacta"
            ppt_bytes = gerar_ppt(
                items, resultados, titulo,
                cfg["max_per_slide"], cfg["sort_mode"], cfg["bg_rgb"],
                cfg["logo_bytes"], cfg["logo_width_in"],
                cfg["signature_bytes"], cfg["signature_width_in"],
                cfg["auto_half_signature"],
                cfg["signature_bottom_margin_in"], cfg["signature_right_margin_in"],
                cfg["title_font_name"], cfg["title_font_size_pt"], cfg["title_font_bold"],
                excluded_urls=st.session_state.excluded_urls
            )
            st.success(f"PPT gerado! (excluídas {len(st.session_state.excluded_urls)} foto(s))")
            st.download_button(
                "⬇️ Baixar PPT",
                data=ppt_bytes,
                file_name=f"{titulo}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

# ===== Roteamento =====
if not st.session_state.auth:
    do_login()
else:
    main_app()
