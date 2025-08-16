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

# ===== Tema (claro predominante, acento laranja e detalhes preto) =====
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
            background: var(--accent-hover); color: white;
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
            background: var(--accent-hover); color: white;
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

# ===== Login (lista completa) =====
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

# --- helpers de aparência ---
def hex_to_rgb(hex_str: str):
    s = hex_str.strip().lstrip("#")
    if len(s) == 3:
        s = "".join([c*2 for c in s])
    r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
    return r, g, b

def pick_contrast_color(r, g, b):
    brightness = (r*299 + g*587 + b*114) / 1000
    return (0,0,0) if brightness > 128 else (255,255,255)

# --- layout (1, 2 ou 3 por slide) ---
def get_slots(n, prs):
    IMG_TOP = Inches(1.2)
    CONTENT_W = Inches(11)
    CONTENT_H = Inches(6)
    GAP = Inches(0.2)
    start_left = (prs.slide_width - CONTENT_W) / 2
    if n == 1:
        return [(start_left, IMG_TOP, CONTENT_W, CONTENT_H)]
    else:
        cols = n
        total_gap = GAP * (cols - 1)
        cell_w = (CONTENT_W - total_gap) / cols
        slots = []
        for c in range(cols):
            left = start_left + c * (cell_w + GAP)
            slots.append((left, IMG_TOP, cell_w, CONTENT_H))
        return slots

def add_title(slide, text, title_rgb=(0,0,0), font_name="Radikal", font_size_pt=15, font_bold=True):
    TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H = Inches(0.5), Inches(0.2), Inches(12), Inches(1)
    tx = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H)
    tf = tx.text_frame; tf.clear()
    p = tf.paragraphs[0]; run = p.add_run(); run.text = text
    font = run.font
    font.name = font_name or "Radikal"   # se não vier nada, tenta Radikal
    font.size = Pt(font_size_pt or 15)
    font.bold = bool(font_bold)
    font.color.rgb = RGBColor(*title_rgb)
    p.alignment = 1

def set_slide_bg(slide, rgb_tuple):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb_tuple)

def place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in):
    img_w_in = px_to_inches(w_px)
    img_h_in = px_to_inches(h_px)
    ratio = min(float(max_w_in)/float(img_w_in), float(max_h_in)/float(img_h_in), 1.0)
    final_w = img_w_in * ratio
    final_h = img_h_in * ratio
    x = left + (max_w_in - final_w)/2
    y = top  + (max_h_in - final_h)/2
    buf.seek(0)
    slide.shapes.add_picture(buf, x, y, width=final_w, height=final_h)

def add_logo_top_right(slide, prs, logo_bytes: bytes, logo_width_in: float):
    if not logo_bytes:
        return
    left = prs.slide_width - Inches(0.5) - Inches(logo_width_in)
    top = Inches(0.2)
    slide.shapes.add_picture(BytesIO(logo_bytes), left, top, width=Inches(logo_width_in))

def add_signature_bottom_right(
    slide, prs, signature_bytes: bytes, signature_width_in: float,
    bottom_margin_in: float = 0.5, right_margin_in: float = 0.5
):
    """Assinatura no canto inferior direito com margens configuráveis."""
    if not signature_bytes:
        return
    try:
        im = Image.open(BytesIO(signature_bytes))
        w_px, h_px = im.size
        ratio = h_px / float(w_px) if w_px else 0.4
    except Exception:
        ratio = 0.4
    sig_h_in = signature_width_in * ratio

    left = prs.slide_width - Inches(right_margin_in) - Inches(signature_width_in)
    top  = prs.slide_height - Inches(bottom_margin_in) - Inches(sig_h_in)
    slide.shapes.add_picture(BytesIO(signature_bytes), left, top, width=Inches(signature_width_in))

# ===== Geração do PPT (respeita exclusões) =====
def gerar_ppt(
    items, resultados, titulo, max_per_slide, sort_mode, bg_rgb,
    logo_bytes=None, logo_width_in=1.2,
    signature_bytes=None, signature_width_in=None,
    auto_half_signature=True,
    signature_bottom_margin_in=0.2, signature_right_margin_in=0.2,
    title_font_name="Radikal", title_font_size_pt=15, title_font_bold=True,
    excluded_urls=None
):
    excluded_urls = excluded_urls or set()

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    blank = prs.slide_layouts[6]

    # agrupa por loja (respeitando exclusões)
    groups = OrderedDict()
    for loja, url in items:
        if url in resultados and url not in excluded_urls:
            groups.setdefault(str(loja), []).append((url, resultados[url]))

    # ordenação
    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip() == "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    title_rgb = pick_contrast_color(*bg_rgb)

    # tamanho default da assinatura: metade do logo (se auto_half)
    if auto_half_signature:
        base_logo_w = logo_width_in if logo_width_in else 1.2
        signature_width = (base_logo_w / 2.0)
    else:
        signature_width = signature_width_in or 0.6

    for loja in loja_keys:
        imgs = groups[loja]
        for i in range(0, len(imgs), max_per_slide):
            batch = imgs[i:i+max_per_slide]
            slide = prs.slides.add_slide(blank)
            set_slide_bg(slide, bg_rgb)
            add_title(slide, loja, title_rgb, title_font_name, title_font_size_pt, title_font_bold)

            if logo_bytes:
                add_logo_top_right(slide, prs, logo_bytes, logo_width_in or 1.2)

            if signature_bytes:
                add_signature_bottom_right(
                    slide, prs, signature_bytes, signature_width,
                    bottom_margin_in=signature_bottom_margin_in,
                    right_margin_in=signature_right_margin_in
                )

            slots = get_slots(len(batch), prs)
            for (url, (_loja, buf, (w_px, h_px))), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in)

    out = BytesIO(); prs.save(out); out.seek(0); return out

# ===== util: imagem -> HTML com borda colorida =====
def img_to_html_with_border(image: Image.Image, width_px: int, border_px: int, border_color: str):
    im = image.copy()
    im.thumbnail((width_px, width_px))
    buf = BytesIO()
    im.save(buf, format="PNG")
    data = base64.b64encode(buf.getvalue()).decode("utf-8")
    style = f"border:{border_px}px solid {border_color};border-radius:8px;display:block;max-width:100%;width:{width_px}px;"
    return f'<img src="data:image/png;base64,{data}" style="{style}" />'

# ===== PRÉ-VISUALIZAÇÃO persistente =====
def render_preview(items, resultados, max_per_slide, sort_mode, thumb_px: int):
    if "excluded_urls" not in st.session_state:
        st.session_state.excluded_urls = set()
    excluded = st.session_state.excluded_urls

    # agrupa por loja
    groups = OrderedDict()
    for loja, url in items:
        if url in resultados:
            groups.setdefault(str(loja), []).append((url, resultados[url]))

    # ordenação
    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip() == "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    # toolbar topo
    top_c1, top_c2, top_c3 = st.columns([1,1,1])
    with top_c1:
        if st.button("🔄 Fechar pré-visualização", type="secondary", use_container_width=True):
            st.session_state.preview_mode = False
            st.rerun()
    with top_c2:
        if st.button("🧹 Limpar TODAS as exclusões", type="secondary", use_container_width=True):
            excluded.clear()
            st.rerun()
    with top_c3:
        st.caption(f"Excluídas: **{len(excluded)}** foto(s)")

    # render por loja
    for loja in loja_keys:
        imgs = groups[loja]
        with st.expander(f"📄 {loja} — {len(imgs)} foto(s)", expanded=False):
            # ações por loja
            lc1, lc2, _ = st.columns([1,1,2])
            with lc1:
                if st.button(f"Selecionar todas de {loja}", key=f"sel_all_{hash(loja)}"):
                    for url, _ in imgs:
                        excluded.add(url)
                    st.rerun()
            with lc2:
                if st.button(f"Limpar seleção de {loja}", key=f"clr_sel_{hash(loja)}"):
                    for url, _ in imgs:
                        excluded.discard(url)
                    st.rerun()

            for i in range(0, len(imgs), max_per_slide):
                batch = imgs[i:i+max_per_slide]
                cols = st.columns(len(batch))
                for col, (url, (_loja, buf, (w_px, h_px))) in zip(cols, batch):
                    # miniatura com borda (vermelha se excluída)
                    try:
                        buf.seek(0)
                        im = Image.open(buf)
                    except Exception:
                        col.warning("Não foi possível pré-visualizar esta imagem.")
                        continue

                    is_excluded = url in excluded
                    border_px = 3 if is_excluded else 1
                    border_color = "#E53935" if is_excluded else "#DDDDDD"
                    html_img = img_to_html_with_border(im, thumb_px, border_px, border_color)
                    col.markdown(html_img, unsafe_allow_html=True)

                    key = "ex_" + hashlib.md5(url.encode("utf-8")).hexdigest()
                    default = is_excluded
                    checked = col.checkbox("Excluir esta foto", key=key, value=default)
                    if checked:
                        excluded.add(url)
                    else:
                        excluded.discard(url)

                st.divider()

# ===== App principal =====
def main_app():
    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)

        st.markdown("---")
        st.caption("Colunas da planilha (nomes do cabeçalho):")
        loja_col = st.text_input("Coluna de LOJA", value="Selecione sua loja", key="loja_col")
        img_col  = st.text_input("Coluna de FOTOS", value="Faça o upload das fotos", key="img_col")

        st.markdown("---")
        st.caption("Layout")
        max_per_slide = st.selectbox("Fotos por slide (máx.)", [1, 2, 3], index=0, key="max_per_slide")

        st.caption("Ordenação")
        sort_mode = st.selectbox(
            "Ordenar lojas por",
            ["Ordem original do Excel", "Nome da loja (A→Z)"],
            index=0, key="sort_mode"
        )

        st.markdown("---")
        st.caption("Aparência do slide")
        bg_hex = st.color_picker("Cor de fundo do slide", value="#FFFFFF", key="bg_hex")

        # Fonte do título
        st.caption("Fonte do título (precisa estar instalada no PC de quem abrir o PPT)")
        title_font_name = st.text_input("Nome da fonte", value="Radikal", key="title_font_name")
        title_font_size_pt = st.slider("Tamanho do título (pt)", 8, 48, 15, 1, key="title_font_size_pt")
        title_font_bold = st.checkbox("Negrito no título", value=True, key="title_font_bold")

        # Logo (topo direito)
        logo_file = st.file_uploader("Logo (PNG/JPG) — canto superior direito", type=["png","jpg","jpeg"], key="logo_uploader")
        if "logo_bytes" not in st.session_state:
            st.session_state.logo_bytes = None
        if logo_file is not None:
            st.session_state.logo_bytes = logo_file.getvalue()
        logo_width_in = st.slider("Largura do LOGO (em polegadas)", 0.5, 3.0, 1.2, 0.1, key="logo_width_in")

        # Assinatura (canto inferior direito)
        st.markdown("—")
        st.caption("Assinatura (opcional)")
        signature_file = st.file_uploader("Assinatura (PNG/JPG) — canto inferior direito", type=["png","jpg","jpeg"], key="signature_uploader")
        if "signature_bytes" not in st.session_state:
            st.session_state.signature_bytes = None
        if signature_file is not None:
            st.session_state.signature_bytes = signature_file.getvalue()

        auto_half_signature = st.checkbox("Usar 1/2 do tamanho do logo (recomendado)", value=True, key="auto_half_signature")
        derived_default_sig = (st.session_state.get("logo_width_in", 1.2) / 2.0)
        if not auto_half_signature:
            signature_width_in = st.slider(
                "Largura da ASSINATURA (em polegadas)",
                0.3, 2.0, float(derived_default_sig), 0.05,
                key="signature_width_in"
            )
        else:
            if "signature_width_in" not in st.session_state:
                st.session_state.signature_width_in = float(derived_default_sig)

        signature_right_margin_in = st.slider("Margem direita da ASSINATURA (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_right_margin")
        signature_bottom_margin_in = st.slider("Margem inferior da ASSINATURA (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_bottom_margin")

        st.markdown("---")
        st.caption("Pré-visualização")
        thumb_px = st.slider("Tamanho das miniaturas (px)", 120, 400, 220, 10, key="thumb_px")

        st.markdown("---")
        st.caption("Tamanho e compressão")
        target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10, key="target_w")
        target_h = st.number_input("Altura máx (px)",  360, 4096, 720, 10, key="target_h")
        limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 2000, 450, 10, key="limite_kb")

        st.markdown("---")
        st.caption("Rede e paralelismo")
        max_workers = st.slider("Trabalhos em paralelo", 2, 32, 12, key="max_workers")
        req_timeout = st.slider("Timeout por download (s)", 5, 60, 15, key="req_timeout")

    st.title("📸 Gerador de Book (PPT)")
    st.write("Arraste sua planilha Excel aqui (com os links das fotos).")

    up = st.file_uploader("Selecione ou arraste a planilha (.xlsx)", type=["xlsx"], key="xlsx_upload")

    # ===== Botões de fluxo =====
    st.markdown("### Etapas")
    btn_col1, btn_col2 = st.columns([1, 1])
    with btn_col1:
        btn_preview = st.button("Pré-visualizar", key="btn_preview")
    with btn_col2:
        btn_generate = st.button("Gerar & Baixar PPT", key="btn_generate")

    # ===== Estados globais =====
    if "pipeline" not in st.session_state:
        st.session_state.pipeline = {}
    if "excluded_urls" not in st.session_state:
        st.session_state.excluded_urls = set()
    if "preview_mode" not in st.session_state:
        st.session_state.preview_mode = False

    # ===== Processar planilha para PRÉVIA (preenche pipeline) =====
    if btn_preview:
        if not up:
            st.warning("Envie a planilha primeiro.")
        else:
            try:
                df = pd.read_excel(up)
            except Exception as e:
                st.error(f"Não consegui ler o Excel: {e}")
                return

            loja_col = st.session_state["loja_col"]
            img_col  = st.session_state["img_col"]

            missing = [c for c in [img_col, loja_col] if c not in df.columns]
            if missing:
                st.error(f"Colunas não encontradas: {missing}")
                return

            # lista (loja, url) sem duplicados (preserva ordem)
            items = []
            for _, row in df.iterrows():
                loja = str(row[loja_col]).strip()
                for url in extrair_links(row.get(img_col, "")):
                    if url.startswith("http"):
                        items.append((loja, url))
            seen, uniq = set(), []
            for loja, url in items:
                if url not in seen:
                    seen.add(url); uniq.append((loja, url))
            items = uniq

            # ordenação opcional por nome (mantém ordem interna de cada loja)
            if st.session_state["sort_mode"] == "Nome da loja (A→Z)":
                grouped_tmp = OrderedDict()
                for loja, url in items:
                    grouped_tmp.setdefault(loja, []).append(url)
                items = []
                for loja in sorted(grouped_tmp.keys(), key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold())):
                    for url in grouped_tmp[loja]:
                        items.append((loja, url))

            total = len(items)
            if total == 0:
                st.warning("Nenhuma URL de imagem encontrada.")
                return

            st.info(f"Serão processadas **{total}** imagens.")
            session = requests.Session()
            adapter = requests.adapters.HTTPAdapter(pool_connections=st.session_state["max_workers"],
                                                    pool_maxsize=st.session_state["max_workers"],
                                                    max_retries=2)
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

            # grava pipeline (usa SEMPRE os bytes persistidos e as preferências)
            st.session_state.pipeline = {
                "items": items,
                "resultados": resultados,
                "falhas": falhas,
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
                }
            }
            st.session_state.preview_mode = True  # mantém a prévia ativa após qualquer rerun

    # ===== Render da PRÉVIA (persistente enquanto preview_mode=True) =====
    if st.session_state.preview_mode and st.session_state.pipeline:
        p = st.session_state.pipeline
        render_preview(
            p["items"], p["resultados"],
            p["settings"]["max_per_slide"],
            p["settings"]["sort_mode"],
            p["settings"]["thumb_px"]
        )
        st.info("Marque **Excluir esta foto** nas imagens que não devem ir para o PPT. Depois clique em **Gerar & Baixar PPT**.")

    # ===== Gerar PPT =====
    if btn_generate:
        if not st.session_state.pipeline:
            st.warning("Clique em **Pré-visualizar** para processar a planilha antes de gerar.")
        else:
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
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

# ===== Roteamento =====
if not st.session_state.auth:
    do_login()
else:
    c1, c2 = st.columns([1,1])
    with c1: st.caption(f"Logado como: **{st.session_state.user_email}**")
    with c2:
        if st.button("Sair", type="secondary"):
            st.session_state.clear(); st.rerun()
    main_app()
