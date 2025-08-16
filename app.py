# app.py
import re
import hashlib
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

# ===== Login simples =====
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

def add_title(slide, text, title_rgb=(0,0,0)):
    TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H = Inches(0.5), Inches(0.2), Inches(12), Inches(1)
    tx = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H)
    tf = tx.text_frame; tf.clear()
    p = tf.paragraphs[0]; run = p.add_run(); run.text = text
    font = run.font; font.name='Arial'; font.size=Pt(15); font.bold=True
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

# ===== Geração do PPT (respeita exclusões) =====
def gerar_ppt(items, resultados, titulo, max_per_slide, sort_mode, bg_rgb,
              logo_bytes=None, logo_width_in=1.2, excluded_urls=None):
    excluded_urls = excluded_urls or set()

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    blank = prs.slide_layouts[6]

    # agrupa por loja (respeitando exclusões)
    groups = OrderedDict()
    for loja, url in items:
        if url in resultados and url not in excluded_urls:
            groups.setdefault(str(loja), []).append((url, resultados[url]))  # (url, (loja, buf, (w,h)))

    # ordenação
    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip() == "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    title_rgb = pick_contrast_color(*bg_rgb)

    for loja in loja_keys:
        imgs = groups[loja]
        for i in range(0, len(imgs), max_per_slide):
            batch = imgs[i:i+max_per_slide]
            slide = prs.slides.add_slide(blank)
            set_slide_bg(slide, bg_rgb)
            add_title(slide, loja, title_rgb)
            if logo_bytes:
                add_logo_top_right(slide, prs, logo_bytes, logo_width_in)

            slots = get_slots(len(batch), prs)
            for (url, (_loja, buf, (w_px, h_px))), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                place_picture(slide, buf, w_px, h_px, left, top, max_w_in, max_h_in)

    out = BytesIO(); prs.save(out); out.seek(0); return out

# ===== PRÉ-VISUALIZAÇÃO com modo seleção =====
def render_preview(items, resultados, max_per_slide, sort_mode, select_mode: bool):
    # estados
    excluded = st.session_state.excluded_urls
    pending = st.session_state.pending_remove  # buffer de seleção

    # agrupa por loja (mantendo url)
    groups = OrderedDict()
    for loja, url in items:
        if url in resultados:
            groups.setdefault(str(loja), []).append((url, resultados[url]))  # (url, (loja, buf, (w,h)))

    # ordena lojas
    if sort_mode == "Nome da loja (A→Z)":
        loja_keys = sorted(groups.keys(), key=lambda s: (s is None or str(s).strip() == "", (s or "").strip().casefold()))
    else:
        loja_keys = list(groups.keys())

    # sumário
    st.caption(f"Excluídas permanentemente: **{len(excluded)}** · Selecionadas agora: **{len(pending)}**")

    # render por "páginas" (simulando slides)
    for loja in loja_keys:
        imgs = groups[loja]
        with st.expander(f"📄 {loja} — {len(imgs)} foto(s)", expanded=False):
            # atalhos por loja (só no modo seleção)
            if select_mode:
                c1, c2 = st.columns([1,1])
                with c1:
                    if st.button(f"Selecionar todas de {loja}", key=f"sel_all_{hash(loja)}"):
                        for url, _ in imgs:
                            if url not in excluded:
                                pending.add(url)
                        st.rerun()
                with c2:
                    if st.button(f"Limpar seleção de {loja}", key=f"clr_sel_{hash(loja)}"):
                        for url, _ in imgs:
                            if url in pending:
                                pending.remove(url)
                        st.rerun()

            # grid de miniaturas
            for i in range(0, len(imgs), max_per_slide):
                batch = imgs[i:i+max_per_slide]
                cols = st.columns(len(batch))
                for col, (url, (_loja, buf, (w_px, h_px))) in zip(cols, batch):
                    try:
                        buf.seek(0)
                        im = Image.open(buf).copy()
                        im.thumbnail((512, 512))
                        col.image(im, use_column_width=True)
                    except Exception:
                        col.warning("Não foi possível pré-visualizar esta imagem.")

                    # status/controle
                    if url in excluded:
                        col.caption("🚫 Já excluída")
                    elif select_mode:
                        # checkbox só aparece em modo seleção e NÃO em itens já excluídos
                        key = "sel_" + hashlib.md5(url.encode("utf-8")).hexdigest()
                        default = url in pending
                        checked = col.checkbox("Selecionar p/ remover", key=key, value=default)
                        if checked:
                            pending.add(url)
                        else:
                            if url in pending:
                                pending.remove(url)

# ===== App principal =====
def main_app():
    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)

        st.markdown("---")
        st.caption("Colunas da planilha (nomes do cabeçalho):")
        loja_col = st.text_input("Coluna de LOJA", value="Selecione sua loja")
        img_col  = st.text_input("Coluna de FOTOS", value="Faça o upload das fotos")

        st.markdown("---")
        st.caption("Layout")
        max_per_slide = st.selectbox("Fotos por slide (máx.)", [1, 2, 3], index=0)

        st.caption("Ordenação")
        sort_mode = st.selectbox(
            "Ordenar lojas por",
            ["Ordem original do Excel", "Nome da loja (A→Z)"],
            index=0
        )

        st.markdown("---")
        st.caption("Aparência do slide")
        bg_hex = st.color_picker("Cor de fundo do slide", value="#FFFFFF")
        logo_file = st.file_uploader("Logo (PNG/JPG) — canto superior direito", type=["png","jpg","jpeg"])
        logo_width_in = st.slider("Largura do logo (em polegadas)", 0.5, 3.0, 1.2, 0.1)

        st.markdown("---")
        st.caption("Tamanho e compressão")
        target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10)
        target_h = st.number_input("Altura máx (px)",  360, 4096, 720, 10)
        limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 2000, 450, 10)

        st.markdown("---")
        st.caption("Rede e paralelismo")
        max_workers = st.slider("Trabalhos em paralelo", 2, 32, 12)
        req_timeout = st.slider("Timeout por download (s)", 5, 60, 15)

    st.title("📸 Gerador de Book (PPT)")
    st.write("Arraste sua planilha Excel aqui (com os links das fotos).")

    up = st.file_uploader("Selecione ou arraste a planilha (.xlsx)", type=["xlsx"])

    # ===== Botões de fluxo =====
    st.markdown("### Etapas")
    btn_col1, btn_col2 = st.columns([1, 1])
    with btn_col1:
        btn_preview = st.button("Pré-visualizar", key="btn_preview")
    with btn_col2:
        btn_generate = st.button("Gerar & Baixar PPT", key="btn_generate")

    if not up:
        st.info("Envie a planilha para pré-visualizar ou gerar.")

    # ===== Estados globais =====
    if "pipeline" not in st.session_state:
        st.session_state.pipeline = {}
    if "excluded_urls" not in st.session_state:
        st.session_state.excluded_urls = set()
    if "pending_remove" not in st.session_state:
        st.session_state.pending_remove = set()
    if "select_mode" not in st.session_state:
        st.session_state.select_mode = False

    # ===== Toolbar superior: remover fotos / aplicar / cancelar / reverter =====
    tool = st.container()
    with tool:
        disabled_toolbar = not bool(st.session_state.pipeline)
        colA, colB, colC, colD = st.columns([1,1,1,1])
        if not st.session_state.select_mode:
            with colA:
                if st.button("🗑️ Remover fotos", disabled=disabled_toolbar, key="enter_select"):
                    st.session_state.select_mode = True
                    st.session_state.pending_remove = set()
                    st.rerun()
            with colD:
                if st.button("↩️ Reverter exclusões", disabled=disabled_toolbar or len(st.session_state.excluded_urls)==0, key="revert_all"):
                    st.session_state.excluded_urls.clear()
                    st.experimental_rerun()
        else:
            with colA:
                if st.button("✅ Aplicar exclusões", key="apply_sel"):
                    st.session_state.excluded_urls |= st.session_state.pending_remove
                    st.session_state.pending_remove = set()
                    st.session_state.select_mode = False
                    st.success("Exclusões aplicadas.")
                    st.experimental_rerun()
            with colB:
                if st.button("❌ Cancelar seleção", key="cancel_sel"):
                    st.session_state.pending_remove = set()
                    st.session_state.select_mode = False
                    st.experimental_rerun()
            with colD:
                st.caption(f"Selecionadas agora: {len(st.session_state.pending_remove)}")

    # ===== Processamento (preview/gerar) =====
    if (btn_preview or btn_generate) and not up:
        st.warning("Envie a planilha primeiro."); st.stop()

    if up and (btn_preview or btn_generate):
        try:
            df = pd.read_excel(up)
        except Exception as e:
            st.error(f"Não consegui ler o Excel: {e}"); st.stop()

        loja_col = st.session_state.get('loja_col', None) or st.session_state._old_state.get('loja_col') if hasattr(st.session_state, '_old_state') else None
        # (mantemos os valores do sidebar diretamente)
        loja_col = st.session_state.get('loja_col', None) or "Selecione sua loja"
        img_col  = st.session_state.get('img_col', None) or "Faça o upload das fotos"

        # checa colunas
        missing = [c for c in [img_col, loja_col] if c not in df.columns]
        if missing:
            st.error(f"Colunas não encontradas: {missing}"); st.stop()

        # lista (loja, url) sem duplicados
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

        total = len(items)
        if total == 0:
            st.warning("Nenhuma URL de imagem encontrada."); st.stop()

        st.info(f"Serão processadas **{total}** imagens.")
        session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(pool_connections=max_workers, pool_maxsize=max_workers, max_retries=2)
        session.mount("http://", adapter); session.mount("https://", adapter)
        session.headers.update({"User-Agent": "Mozilla/5.0 (GeradorBook Streamlit)"})

        prog = st.progress(0); status = st.empty()
        resultados, falhas, done = {}, 0, 0

        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {ex.submit(baixar_processar, session, url, target_w, target_h, limite_kb, req_timeout): (loja, url) for loja, url in items}
            for fut in as_completed(futures):
                loja, url = futures[fut]
                ok_url, ok, buf, wh = fut.result()
                if ok: resultados[url] = (loja, buf, wh)
                else: falhas += 1
                done += 1; prog.progress(int(done * 100 / total))
                status.write(f"Processadas {done}/{total} imagens...")

        status.write(f"Concluído. Falhas: {falhas}")

        # guarda para reuso
        st.session_state.pipeline = {
            "items": items,
            "resultados": resultados,
            "falhas": falhas,
            "settings": {
                "max_per_slide": st.session_state.get('max_per_slide', None) or max_per_slide,
                "sort_mode": st.session_state.get('sort_mode', None) or sort_mode,
                "bg_rgb": hex_to_rgb(st.session_state.get('bg_hex', None) or bg_hex),
                "logo_bytes": (logo_file.read() if logo_file else None),
                "logo_width_in": st.session_state.get('logo_width_in', None) or logo_width_in,
            }
        }

        if btn_preview:
            st.subheader("👀 Pré-visualização")
            render_preview(items, resultados, max_per_slide, sort_mode, st.session_state.select_mode)
            st.info("Use **🗑️ Remover fotos** para entrar no modo seleção, marque as fotos e clique **✅ Aplicar exclusões**.")

    # ===== Geração do PPT =====
    if btn_generate:
        if not st.session_state.pipeline:
            st.warning("Faça a pré-visualização primeiro, ou clique novamente após o processamento.")
        else:
            p = st.session_state.pipeline
            items = p["items"]; resultados = p["resultados"]; cfg = p["settings"]

            titulo = "Apresentacao_Relatorio_Compacta"
            ppt_bytes = gerar_ppt(
                items, resultados, titulo,
                cfg["max_per_slide"], cfg["sort_mode"], cfg["bg_rgb"],
                cfg["logo_bytes"], cfg["logo_width_in"],
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



