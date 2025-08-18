# app.py
import re
import base64
import hashlib
from io import BytesIO
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
import zipfile

import streamlit as st
import pandas as pd
import requests
from PIL import Image, ImageOps, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# -------------------------------------------------------------------
# CONFIG GERAL
# -------------------------------------------------------------------
st.set_page_config(page_title="Gerador de Book", page_icon="📸", layout="wide")

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

.badge { display:inline-block; padding:3px 8px; border-radius:999px; background:#F3F4F6; color:#111; font-size:12px; font-weight:700; }
.dark .badge { background:#1f2635; color:#eaeaea; }

.logout-zone .stButton > button {
  background:#E53935 !important; color:#fff !important;
  border:none !important; border-radius:10px !important;
}
.logout-zone .stButton > button:hover { background:#C62828 !important; }

.reset-zone .stButton > button {
  background:#FF9800 !important; color:#fff !important;
  border:none !important; border-radius:10px !important;
}
.reset-zone .stButton > button:hover { background:#F57C00 !important; }

.stDownloadButton > button {
  padding:18px 26px !important;
  font-size:18px !important;
  font-weight:800 !important;
  border-radius:12px !important;
}
</style>
"""
st.markdown(BASE_CSS, unsafe_allow_html=True)

# --- CSS extra (login mais limpo) ---
LOGIN_CSS = """
<style>
.login-hero {
  background: linear-gradient(90deg, #FF7A00 0%, #FF9944 100%);
  color:#fff; padding:16px 18px; border-radius:14px; 
  font-weight:700; margin: 8px 0 18px 0; line-height:1.2;
}
.login-card {
  border:1px solid #E7E7E7; background:#ffffff;
  padding:18px; border-radius:14px;
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

# -------------------------------------------------------------------
# DOWNLOAD & PROCESS
# -------------------------------------------------------------------
def baixar_processar(session, url: str, max_w: int, max_h: int, limite_kb: int, timeout: int, fx_cfg: dict = None):
    try:
        r = session.get(url, timeout=timeout, stream=True)
        if r.status_code != 200:
            return (url, False, None, None)

        img = Image.open(BytesIO(r.content))
        img = redimensionar(img, max_w, max_h)

        need_alpha = False
        if fx_cfg and (fx_cfg.get("fx_shadow") or fx_cfg.get("fx_round") or fx_cfg.get("fx_border")):
            img_rgba = apply_effects_pipeline(img.convert("RGB"), fx_cfg)
            need_alpha = True
        else:
            img_rgba = img.convert("RGBA")

        if need_alpha:
            buf = BytesIO(); img_rgba.save(buf, format="PNG", optimize=True)
            if buf.tell() / 1024 <= limite_kb:
                buf.seek(0); w, h = img_rgba.size
                return (url, True, buf, (w, h))
            pal = img_rgba.convert("P", palette=Image.ADAPTIVE, colors=256)
            buf = BytesIO(); pal.save(buf, format="PNG", optimize=True)
            if buf.tell() / 1024 <= limite_kb:
                buf.seek(0); w, h = img_rgba.size
                return (url, True, buf, (w, h))
            bg = Image.new("RGB", img_rgba.size, (255, 255, 255))
            bg.paste(img_rgba, mask=img_rgba.split()[-1])
            buf = comprimir_jpeg_binsearch(bg, limite_kb)
            w, h = bg.size
            return (url, True, buf, (w, h))
        else:
            buf = comprimir_jpeg_binsearch(img.convert("RGB"), limite_kb)
            w, h = img.size
            return (url, True, buf, (w, h))
    except Exception:
        return (url, False, None, None)

# -------------------------------------------------------------------
# PPT HELPERS
# -------------------------------------------------------------------
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
    p = tf.paragraphs[0]; run = p.add_run(); run.text = title_text
    f = run.font; f.name = font_name or "Radikal"; f.size = Pt(title_font_size_pt or 18)
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

def move_slide_to_index(prs, old_index, new_index):
    sldIdLst = prs.slides._sldIdLst
    sld = sldIdLst[old_index]
    sldIdLst.remove(sld)
    sldIdLst.insert(new_index, sld)

# -------------------------------------------------------------------
# GERADOR COM MODELO (capa + final)
# -------------------------------------------------------------------
def gerar_ppt_modelo_capa_final(
    template_bytes: bytes,
    items, resultados, titulo,
    max_per_slide, sort_mode, bg_rgb,
    logo_bytes=None, logo_width_in=1.2,
    signature_bytes=None, signature_width_in=None, auto_half_signature=True,
    signature_bottom_margin_in=0.2, signature_right_margin_in=0.2,
    title_font_name="Radikal", title_font_size_pt=18, title_font_bold=True,
    excluded_urls=None
):
    excluded_urls = excluded_urls or set()
    prs = Presentation(BytesIO(template_bytes))
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

    title_rgb = pick_contrast_color(*bg_rgb)
    signature_width = (logo_width_in/2.0) if auto_half_signature else (signature_width_in or 0.6)

    for loja in loja_keys:
        imgs = groups[loja]
        for i in range(0, len(imgs), max_per_slide):
            batch = imgs[i:i+max_per_slide]
            slide = prs.slides.add_slide(blank_layout)
            set_slide_bg(slide, bg_rgb)

            first_url, (_loja, endereco, _buf0, _wh0) = batch[0]
            add_title_and_address(slide, loja, endereco, title_rgb,
                                  title_font_name, title_font_size_pt, title_font_bold)

            if logo_bytes:
                add_logo_top_right(slide, prs, logo_bytes, logo_width_in or 1.2)
            if signature_bytes:
                add_signature_bottom_right(slide, prs, signature_bytes, signature_width,
                                           bottom_margin_in=signature_bottom_margin_in,
                                           right_margin_in=signature_right_margin_in)

            slots = get_slots(len(batch), prs)
            for (url, (_loja, _end, buf, (w_px, h_px))), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in)

    if final_idx is not None and final_idx < len(prs.slides):
        move_slide_to_index(prs, final_idx, len(prs.slides)-1)

    out = BytesIO(); prs.save(out); out.seek(0); return out

# -------------------------------------------------------------------
# HELPERS ZIP DE IMAGENS (pastas por loja)
# -------------------------------------------------------------------
def _buf_to_jpeg_bytes(img_buf: BytesIO) -> bytes:
    try:
        img_buf.seek(0)
        im = Image.open(img_buf)
        fmt = (im.format or "").upper()
        if fmt in ("JPEG", "JPG"):
            img_buf.seek(0); return img_buf.read()
        if im.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1]); im = bg
        else:
            im = im.convert("RGB")
        out = BytesIO()
        im.save(out, "JPEG", quality=88, optimize=True, progressive=True, subsampling=2)
        out.seek(0); return out.read()
    except Exception:
        img_buf.seek(0); return img_buf.read()

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
            for _url, (_loja, _end, buf, (w, h)) in lista:
                jpeg_bytes = _buf_to_jpeg_bytes(buf)
                arquivo = f"{pasta} - {contador}.jpg"
                caminho = f"{pasta}/{arquivo}"
                zf.writestr(caminho, jpeg_bytes)
                contador += 1
    mem_zip.seek(0)
    return mem_zip

# -------------------------------------------------------------------
# PREVIEW (UI) — com callbacks em 1 clique
# -------------------------------------------------------------------
def img_to_html_with_border(image: Image.Image, width_px: int, border_px: int, border_color: str):
    im = image.copy(); im.thumbnail((width_px, width_px))
    buf = BytesIO(); im.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    style = f"border:{border_px}px solid {border_color};border-radius:10px;display:block;max-width:100%;width:{width_px}px;"
    return f'<img src="data:image/png;base64,{b64}" style="{style}" />'

def render_steps(current: int):
    labels = ["Upload", "Pré-visualização", "Gerar/Exportar"]
    html = ['<div class="steps">']
    for i, txt in enumerate(labels, start=1):
        cls = "step active" if i == current else "step"
        html.append(f'<div class="{cls}">{i}. {txt}</div>')
        if i < len(labels): html.append('<span class="step sep">›</span>')
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)

def render_summary(items, resultados, excluded):
    total_urls = len(items)
    baixadas = sum(1 for _, _, url in items if url in resultados)
    lojas = len({loja for loja, _, _ in items})
    st.markdown(
        f"**Resumo:** "
        f"<span class='badge'>Lojas: {lojas}</span> "
        f"<span class='badge'>Links: {total_urls}</span> "
        f"<span class='badge'>Baixadas: {baixadas}</span> "
        f"<span class='badge'>Excluídas: {len(excluded)}</span>",
        unsafe_allow_html=True
    )

def _cb_select_all(loja, imgs):
    excluded = st.session_state.excluded_urls
    for url, _ in imgs:
        excluded.add(url)
    st.rerun()

def _cb_clear_group(loja, imgs):
    excluded = st.session_state.excluded_urls
    for url, _ in imgs:
        excluded.discard(url)
    st.rerun()

def render_preview(items, resultados, sort_mode, thumb_px: int, thumbs_per_row: int):
    if "excluded_urls" not in st.session_state:
        st.session_state.excluded_urls = set()
    if "expanded_groups" not in st.session_state:
        st.session_state.expanded_groups = {}

    excluded = st.session_state.excluded_urls
    expanded_groups = st.session_state.expanded_groups

    groups = OrderedDict()
    for loja, endereco, url in items:
        if url in resultados:
            groups.setdefault(str(loja), []).append((url, resultados[url]))

    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(),
                           key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    c1, c2, c3, c4 = st.columns([1,1,1,1])
    with c1:
        if st.button("🧹 Limpar todas as exclusões", type="secondary", use_container_width=True):
            excluded.clear(); st.rerun()
    with c2:
        if st.button("🔁 Inverter seleção", type="secondary", use_container_width=True):
            all_urls = {url for _, v in groups.items() for (url, _) in v}
            st.session_state.excluded_urls = all_urls - excluded
            st.rerun()
    with c3:
        if st.button("➕ Expandir todas", type="secondary", use_container_width=True):
            for loja in loja_keys: expanded_groups[loja] = True
            st.rerun()
    with c4:
        if st.button("➖ Recolher todas", type="secondary", use_container_width=True):
            for loja in loja_keys: expanded_groups[loja] = False
            st.rerun()

    st.caption(f"Excluídas até agora: **{len(excluded)}**")

    for loja in loja_keys:
        imgs = groups[loja]
        key_toggle = f"tg_{hash(loja)}"
        current_state = expanded_groups.get(loja, True)

        st.markdown(
            f"""<div class="group-head">
                <div><strong>📄 {loja}</strong> — {len(imgs)} foto(s)</div>
                <div class="badge">{'Aberto' if current_state else 'Fechado'}</div>
            </div>""",
            unsafe_allow_html=True
        )
        new_state = st.toggle("Manter este grupo visível", value=current_state, key=key_toggle)
        expanded_groups[loja] = bool(new_state)

        gc1, gc2 = st.columns([1,1])
        with gc1:
            st.button(
                f"Selecionar todas de {loja}",
                key=f"sel_all_{hash(loja)}",
                use_container_width=True,
                on_click=_cb_select_all,
                args=(loja, imgs)
            )
        with gc2:
            st.button(
                f"Limpar seleção de {loja}",
                key=f"clr_sel_{hash(loja)}",
                use_container_width=True,
                on_click=_cb_clear_group,
                args=(loja, imgs)
            )

        if expanded_groups[loja]:
            cols = st.columns(thumbs_per_row)
            col_idx = 0
            for (url, (_loja, _end, buf, (w_px, h_px))) in imgs:
                with cols[col_idx]:
                    try:
                        buf.seek(0); im = Image.open(buf)
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
                    if checked: excluded.add(url)
                    else: excluded.discard(url)
                    st.markdown('</div>', unsafe_allow_html=True)
                col_idx = (col_idx + 1) % thumbs_per_row
        st.divider()

# -------------------------------------------------------------------
# RESET (limpa PPT e ZIP)
# -------------------------------------------------------------------
def reset_app(preserve_login: bool = True):
    user = st.session_state.get("user_email")
    auth = st.session_state.get("auth", False)

    st.session_state.clear()

    # reiniciar chaves dos uploaders/botões
    st.session_state.xlsx_key = 0
    st.session_state.template_key = 0
    st.session_state.logo_key = 0
    st.session_state.sign_key = 0
    st.session_state.download_key = 0
    st.session_state.images_zip_key = 0

    # estados de expanders padrão
    st.session_state.exp_plan = True
    st.session_state.exp_style = False
    st.session_state.exp_brand = False
    st.session_state.exp_fx = False
    st.session_state.exp_perf = False
    st.session_state.exp_model = False

    # artefatos
    st.session_state.ppt_bytes = None
    st.session_state.images_zip_bytes = None
    st.session_state.generated = False
    st.session_state.output_filename = "Modelo_01"

    if preserve_login and auth:
        st.session_state.auth = True
        st.session_state.user_email = user
        st.session_state.dark_mode = False
    st.rerun()


# -------------------------------------------------------------------
# APP
# -------------------------------------------------------------------
def main_app():
    # Inicializações
    for k in ["xlsx_key","template_key","logo_key","sign_key","download_key","images_zip_key"]:
        if k not in st.session_state: st.session_state[k] = 0
    if "exp_plan" not in st.session_state:
        st.session_state.exp_plan = True
        st.session_state.exp_style = False
        st.session_state.exp_brand = False
        st.session_state.exp_fx = False
        st.session_state.exp_perf = False
        st.session_state.exp_model = False
    if "pipeline" not in st.session_state: st.session_state.pipeline = {}
    if "excluded_urls" not in st.session_state: st.session_state.excluded_urls = set()
    if "preview_mode" not in st.session_state: st.session_state.preview_mode = False
    if "expanded_groups" not in st.session_state: st.session_state.expanded_groups = {}
    if "output_filename" not in st.session_state: st.session_state.output_filename = "Modelo_01"
    if "generated" not in st.session_state: st.session_state.generated = False
    if "ppt_bytes" not in st.session_state: st.session_state.ppt_bytes = None
    if "images_zip_bytes" not in st.session_state: st.session_state.images_zip_bytes = None

    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)

        with st.expander("📄 Planilha & Layout", expanded=st.session_state.exp_plan):
            st.caption("Colunas (nomes exatos do cabeçalho):")
            loja_col = st.text_input("🛒 Coluna de LOJA", value="Selecione sua loja", key="loja_col")
            img_col  = st.text_input("🖼️ Coluna de FOTOS", value="Faça o upload das fotos", key="img_col")
            use_address = st.checkbox("➕ Incluir endereço abaixo do nome da loja", value=False, key="use_address")
            address_col = st.text_input("🏠 Coluna de ENDEREÇO", value="Endereço",
                                        key="address_col", disabled=not use_address)
            max_per_slide = st.selectbox("📐 Fotos por slide (máx.)", [1,2,3], index=0, key="max_per_slide")
            st.caption("Ordenação")
            sort_mode = st.selectbox("🔤 Ordenar lojas por",
                ["Ordem original do Excel", "Nome da loja (A→Z)"], index=0, key="sort_mode")

        with st.expander("🎨 Aparência do slide", expanded=st.session_state.exp_style):
            bg_hex = st.color_picker("🎨 Cor de fundo", value="#FFFFFF", key="bg_hex")
            st.caption("Título do slide")
            title_font_name = st.text_input("Fonte do título", value="Radikal", key="title_font_name")
            title_font_size_pt = st.slider("Tamanho (pt)", 8, 48, 18, 1, key="title_font_size_pt")
            title_font_bold = st.checkbox("Negrito", value=True, key="title_font_bold")

        with st.expander("🏷️ Logo & ✍️ Assinatura", expanded=st.session_state.exp_brand):
            st.caption("Logo (canto superior direito)")
            logo_file = st.file_uploader("Logo (PNG/JPG)", type=["png","jpg","jpeg"],
                                         key=f"logo_uploader_{st.session_state.logo_key}")
            if "logo_bytes" not in st.session_state: st.session_state.logo_bytes = None
            if logo_file is not None: st.session_state.logo_bytes = logo_file.getvalue()
            logo_width_in = st.slider("Largura do LOGO (pol)", 0.5, 3.0, 1.2, 0.1, key="logo_width_in")

            st.markdown("---")
            st.caption("Assinatura (canto inferior direito)")
            signature_file = st.file_uploader("Assinatura (PNG/JPG)", type=["png","jpg","jpeg"],
                                              key=f"signature_uploader_{st.session_state.sign_key}")
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
            thumb_px = st.slider("Tamanho das miniaturas (px)", 120, 400, 220, 10, key="thumb_px")
            thumbs_per_row = st.slider("Miniaturas por linha", 2, 8, 4, 1, key="thumbs_per_row")
            st.caption("Redimensionamento / compressão")
            target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10, key="target_w")
            target_h = st.number_input("Altura máx (px)",  360, 4096, 720, 10, key="target_h")
            limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 2000, 450, 10, key="limite_kb")
            st.caption("Rede e paralelismo")
            max_workers = st.slider("Trabalhos em paralelo", 2, 32, 12, key="max_workers")
            req_timeout = st.slider("Timeout por download (s)", 5, 60, 15, key="req_timeout")

    # Topo (título / reset / sair)
    top_l, top_m, top_r = st.columns([5,1,1])
    with top_l:
        current_step = 1
        if st.session_state.get("preview_mode") and not st.session_state.get("generated"):
            current_step = 2
        if st.session_state.get("generated") or st.session_state.get("images_zip_bytes"):
            current_step = 3
        st.title("📸 Gerador de Book")
        render_steps(current_step)
        st.caption("Monte um PPT com fotos por loja (capa e final opcionais via modelo).")
    with top_m:
        st.markdown('<div class="reset-zone">', unsafe_allow_html=True)
        if st.button("Resetar", key="reset_btn", use_container_width=True, type="secondary"):
            st.session_state.xlsx_key += 1
            st.session_state.template_key += 1
            st.session_state.logo_key += 1
            st.session_state.sign_key += 1
            st.session_state.download_key += 1
            st.session_state.images_zip_key += 1
            reset_app(preserve_login=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with top_r:
        st.markdown('<div class="logout-zone">', unsafe_allow_html=True)
        if st.button("Sair", key="logout_btn", use_container_width=True, type="secondary"):
            reset_app(preserve_login=False)
        st.markdown('</div>', unsafe_allow_html=True)

    # 1) Upload
    with st.expander("1. Upload", expanded=not st.session_state.preview_mode):
        up = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"],
                              key=f"xlsx_upload_{st.session_state.xlsx_key}")
        btn_preview = st.button("🔎 Pré-visualizar", key="btn_preview", use_container_width=True)

        if btn_preview:
            if not up:
                st.warning("Envie a planilha primeiro.")
            else:
                try:
                    df = pd.read_excel(up)
                except Exception as e:
                    st.error(f"Não consegui ler o Excel: {e}"); st.stop()

                loja_col = st.session_state["loja_col"]
                img_col  = st.session_state["img_col"]
                use_address = st.session_state.get("use_address", False)
                address_col = st.session_state.get("address_col", "Endereço")

                required_cols = [loja_col, img_col] + ([address_col] if use_address else [])
                missing = [c for c in required_cols if c not in df.columns]
                if missing: st.error(f"Colunas não encontradas: {missing}"); st.stop()

                items = []
                for _, row in df.iterrows():
                    loja = str(row[loja_col]).strip()
                    endereco = str(row[address_col]).strip() if use_address else ""
                    for url in extrair_links(row.get(img_col, "")):
                        if url.startswith("http"):
                            items.append((loja, endereco, url))

                seen, uniq = set(), []
                for loja, endereco, url in items:
                    if url not in seen:
                        seen.add(url); uniq.append((loja, endereco, url))
                items = uniq

                if st.session_state["sort_mode"] == "Nome da loja (A→Z)":
                    grouped_tmp = OrderedDict()
                    for loja, end, url in items:
                        grouped_tmp.setdefault(loja, []).append((end, url))
                    items = [(loja, end, url)
                             for loja in sorted(grouped_tmp.keys(), key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
                             for (end, url) in grouped_tmp[loja]]

                total = len(items)
                if total == 0: st.warning("Nenhuma URL de imagem encontrada."); st.stop()

                st.info(f"Baixando e processando **{total}** imagem(ns)...")
                session = requests.Session()
                adapter = requests.adapters.HTTPAdapter(pool_connections=st.session_state["max_workers"],
                                                        pool_maxsize=st.session_state["max_workers"], max_retries=2)
                session.mount("http://", adapter); session.mount("https://", adapter)
                session.headers.update({"User-Agent": "Mozilla/5.0 (GeradorBook Streamlit)"})

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

                prog = st.progress(0); status = st.empty()
                resultados, falhas, done = {}, 0, 0
                with ThreadPoolExecutor(max_workers=st.session_state["max_workers"]) as ex:
                    futures = {ex.submit(
                        baixar_processar, session, url,
                        st.session_state["target_w"], st.session_state["target_h"],
                        st.session_state["limite_kb"], st.session_state["req_timeout"],
                        fx_cfg
                    ): (loja, endereco, url) for loja, endereco, url in items}
                    for fut in as_completed(futures):
                        loja, endereco, url = futures[fut]
                        _url, ok, buf, wh = fut.result()
                        if ok: resultados[url] = (loja, endereco, buf, wh)
                        else: falhas += 1
                        done += 1; prog.progress(int(done * 100 / total))
                        status.write(f"Processadas {done}/{total} imagens...")

                status.write(f"Concluído. Falhas: {falhas}")

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
                        "signature_width_in": st.session_state.get("signature_width_in",
                                                                   (st.session_state["logo_width_in"]/2.0)),
                        "signature_right_margin_in": st.session_state.get("sig_right_margin", 0.20),
                        "signature_bottom_margin_in": st.session_state.get("sig_bottom_margin", 0.20),
                        "thumb_px": st.session_state["thumb_px"],
                        "thumbs_per_row": st.session_state["thumbs_per_row"],
                        "effects": fx_cfg,
                        "use_template": st.session_state.get("use_template", False),
                        "template_bytes": None,  # atribuído abaixo se houver upload
                    }
                }
                # template opcional
                # (precisa ser lido dentro do sidebar; se quiser habilitar, mova o uploader pra cima)
                st.session_state.preview_mode = True
                st.session_state.generated = False
                st.session_state.ppt_bytes = None
                st.session_state.images_zip_bytes = None

                # recolhe expanders
                st.session_state.exp_plan = False
                st.session_state.exp_style = False
                st.session_state.exp_brand = False
                st.session_state.exp_fx = False
                st.session_state.exp_perf = False
                st.session_state.exp_model = False
                st.rerun()

    # 2) Pré-visualização
    with st.expander("2. Pré-visualização", expanded=st.session_state.preview_mode):
        if st.session_state.preview_mode and st.session_state.pipeline:
            p = st.session_state.pipeline
            render_summary(p["items"], p["resultados"], st.session_state.excluded_urls)
            render_preview(
                p["items"], p["resultados"],
                p["settings"]["sort_mode"],
                p["settings"]["thumb_px"],
                p["settings"]["thumbs_per_row"]
            )
            st.info("Marque **Excluir esta foto** nas imagens que não devem ir para o PPT/ZIP, depois use a etapa 3.")

    # 3) Gerar / Exportar
    with st.expander("3. Gerar / Exportar", expanded=st.session_state.preview_mode):
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            st.session_state.output_filename = st.text_input(
                "Nome base do arquivo (sem extensão)",
                value=st.session_state.output_filename,
                key="output_filename_input"
            )
        with col2:
            if st.session_state.ppt_bytes:
                st.download_button(
                    "⬇️ Baixar PPT",
                    data=st.session_state.ppt_bytes,
                    file_name=f"{(st.session_state.output_filename or 'Apresentacao')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    key=f"download_{st.session_state.download_key}"
                )
            else:
                btn_generate = st.button("🧩 Gerar PPT", key="btn_generate", use_container_width=True)

        with col3:
            if st.session_state.images_zip_bytes:
                st.download_button(
                    "⬇️ Baixar Imagens (ZIP)",
                    data=st.session_state.images_zip_bytes,
                    file_name=f"{(st.session_state.output_filename or 'Imagens')}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key=f"images_zip_{st.session_state.images_zip_key}"
                )
            else:
                btn_zip = st.button("🖼️ Baixar Imagens", key="btn_zip", use_container_width=True)

        # Geração do PPT
        if (not st.session_state.ppt_bytes) and ('btn_generate' in locals()) and btn_generate:
            if not st.session_state.pipeline:
                st.warning("Faça a pré-visualização antes de gerar.")
            else:
                p = st.session_state.pipeline
                items = p["items"]; resultados = p["resultados"]; cfg = p["settings"]
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
                        excluded_urls=st.session_state.excluded_urls
                    )
                else:
                    prs = Presentation()
                    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
                    blank = prs.slide_layouts[6]
                    title_rgb = pick_contrast_color(*cfg["bg_rgb"])
                    signature_width = (cfg["logo_width_in"]/2.0) if cfg.get("auto_half_signature", True) \
                                      else (cfg.get("signature_width_in") or 0.6)

                    groups = OrderedDict()
                    for loja, endereco, url in items:
                        if url in resultados and url not in st.session_state.excluded_urls:
                            groups.setdefault(str(loja), []).append((url, resultados[url]))

                    if cfg["sort_mode"] == "Nome da loja (A→Z)":
                        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
                    else:
                        loja_keys = list(groups.keys())

                    for loja in loja_keys:
                        imgs = groups[loja]
                        for i in range(0, len(imgs), cfg["max_per_slide"]):
                            batch = imgs[i:i+cfg["max_per_slide"]]
                            slide = prs.slides.add_slide(blank)
                            set_slide_bg(slide, cfg["bg_rgb"])
                            first_url, (_loja, endereco, _buf0, _wh0) = batch[0]
                            add_title_and_address(slide, loja, endereco, title_rgb,
                                                  cfg["title_font_name"], cfg["title_font_size_pt"], cfg["title_font_bold"])
                            if cfg["logo_bytes"]:
                                add_logo_top_right(slide, prs, cfg["logo_bytes"], cfg["logo_width_in"])
                            if cfg["signature_bytes"]:
                                add_signature_bottom_right(slide, prs, cfg["signature_bytes"], signature_width,
                                                           bottom_margin_in=cfg["signature_bottom_margin_in"],
                                                           right_margin_in=cfg["signature_right_margin_in"])
                            slots = get_slots(len(batch), prs)
                            for (url, (_loja, _end, buf, (w_px, h_px))), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                                place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in)

                    out = BytesIO(); prs.save(out); out.seek(0); ppt_bytes = out

                st.session_state.ppt_bytes = ppt_bytes
                st.session_state.generated = True
                st.rerun()

        # Geração do ZIP de imagens
        if (not st.session_state.images_zip_bytes) and ('btn_zip' in locals()) and btn_zip:
            if not st.session_state.pipeline:
                st.warning("Faça a pré-visualização antes de baixar as imagens.")
            else:
                p = st.session_state.pipeline
                zip_bytes = montar_zip_imagens(
                    items=p["items"],
                    resultados=p["resultados"],
                    excluded_urls=st.session_state.excluded_urls
                )
                st.session_state.images_zip_bytes = zip_bytes
                st.rerun()

# -------------------------------------------------------------------
# ROTEAMENTO
# -------------------------------------------------------------------
if not st.session_state.auth:
    do_login()
else:
    main_app()

  
