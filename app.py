# === PARTE 10/10 ====================================================
# APP (main_app) + Roteamento final - MODIFICADO E CORRIGIDO

def main_app():
    # Inicializações seguras
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
    if "pipeline" not in st.session_state: st.session_state.pipeline = {}
    if "excluded_urls" not in st.session_state: st.session_state.excluded_urls = set()
    if "preview_mode" not in st.session_state: st.session_state.preview_mode = False
    if "expanded_groups" not in st.session_state: st.session_state.expanded_groups = {}
    if "output_filename" not in st.session_state: st.session_state.output_filename = "Modelo_01"
    if "generated" not in st.session_state: st.session_state.generated = False
    if "ppt_bytes" not in st.session_state: st.session_state.ppt_bytes = None
    if "images_zip_bytes" not in st.session_state: st.session_state.images_zip_bytes = None
    if "preview_bump" not in st.session_state: st.session_state.preview_bump = 0
    if "failed_urls" not in st.session_state: st.session_state.failed_urls = []
    if "failed_details" not in st.session_state: st.session_state.failed_details = []
    if "url_line_map" not in st.session_state: st.session_state.url_line_map = {}
    if "quick_generate" not in st.session_state: st.session_state.quick_generate = False
    if "ignore_failed" not in st.session_state: st.session_state.ignore_failed = True

    with st.sidebar:
        st.header("⚙️ Preferências")
        st.session_state.dark_mode = st.toggle("Usar tema escuro", value=st.session_state.dark_mode)
        apply_theme(st.session_state.dark_mode)

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

            st.caption("Posição da assinatura (quanto menor, mais encostada no canto)")
            signature_right_margin_in = st.slider("Margem direita (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_right_margin")
            signature_bottom_margin_in = st.slider("Margem inferior (pol)", 0.0, 1.0, 0.20, 0.05, key="sig_bottom_margin")

        with st.expander("📑 Modelo (capa + final)", expanded=st.session_state.exp_model):
            use_template = st.checkbox("Usar modelo (Capa + Final)", value=False, key="use_template")
            template_file = None
            if use_template:
                template_file = st.file_uploader("Suba o PPTX com 2 slides (Capa e Final)", type=["pptx"], key=f"template_pptx_{st.session_state.template_key}")

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
            # NOVO CAMPO: Comportamento em caso de falha
            ignore_failed = st.checkbox(
                "⚠️ Ignorar falhas e continuar gerando", 
                value=True, 
                key="ignore_failed",
                help="Quando ativado, as fotos que falharam no download serão puladas e o book será gerado apenas com as que funcionaram."
            )
            
            thumb_px = st.slider("Tamanho das miniaturas (px)", 120, 400, 220, 10, key="thumb_px")
            thumbs_per_row = st.slider("Miniaturas por linha", 2, 8, 4, 1, key="thumbs_per_row")
            st.caption("Redimensionamento / compressão")
            target_w = st.number_input("Largura máx (px)", 480, 4096, 1280, 10, key="target_w")
            target_h = st.number_input("Altura máx (px)",  360, 4096, 720, 10, key="target_h")
            limite_kb = st.number_input("Tamanho máx por foto (KB)", 50, 2000, 450, 10, key="limite_kb")
            st.caption("Rede e paralelismo")
            max_workers = st.slider("Trabalhos em paralelo", 2, 32, 6, key="max_workers")
            req_timeout = st.slider("Timeout por download (s)", 5, 60, 15, key="req_timeout")
            st.caption("Critérios de qualidade (após o download)")
            min_mp = st.slider("Megapixels mínimos", 0.1, 5.0, 0.8, 0.1, key="min_megapixels")
            min_blur = st.slider("Limiar de nitidez (blur score)", 5, 300, 45, 5, key="min_blur_score")

    # Topo
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

    # 1) Upload - MODIFICADO: Botões separados
    upload_expander = st.expander("1. Upload", expanded=not st.session_state.preview_mode)
    
    with upload_expander:
        up = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"], key=f"xlsx_upload_{st.session_state.xlsx_key}")
        
        if up:
            # BOTÕES SEPARADOS
            col1, col2 = st.columns(2)
            with col1:
                btn_preview = st.button("👁️ Visualização Rápida", key="btn_preview", use_container_width=True, type="secondary")
            with col2:
                btn_generate_direct = st.button("🚀 Gerar PPT Direto", key="btn_generate_direct", use_container_width=True, type="primary")

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

                    # remove URLs exatas repetidas
                    seen, uniq = set(), []
                    for loja, endereco, url in items:
                        if url not in seen:
                            seen.add(url)
                            uniq.append((loja, endereco, url))
                    items = uniq

                    # se ordenação por loja
                    if st.session_state["sort_mode"] == "Nome da loja (A→Z)":
                        grouped_tmp = OrderedDict()
                        for loja, end, url in items:
                            grouped_tmp.setdefault(loja, []).append((end, url))
                        items = [
                            (loja, end, url)
                            for loja in sorted(grouped_tmp.keys(),
                                key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
                            for (end, url) in grouped_tmp[loja]
                        ]

                    # Cap opcional de itens para ambientes com pouca RAM
                    MAX_ITENS = 1000
                    if len(items) > MAX_ITENS:
                        st.warning(f"Muitas imagens ({len(items)}). Vou processar apenas as primeiras {MAX_ITENS} nesta rodada.")
                        items = items[:MAX_ITENS]

                    total = len(items)
                    if total == 0:
                        st.warning("Nenhuma URL de imagem encontrada.")
                        st.stop()

                    # Aviso sobre o modo de falha
                    if st.session_state.get("ignore_failed", True):
                        st.info("ℹ️ **Modo: Ignorar falhas ativado** - O sistema continuará mesmo se algumas imagens falharem.")
                    else:
                        st.warning("⚠️ **Modo: Parar em caso de falha** - O sistema parará se alguma imagem falhar.")

                    st.info(f"Baixando e processando **{total}** imagem(ns)...")
                    session = requests.Session()
                    adapter = requests.adapters.HTTPAdapter(
                        pool_connections=st.session_state["max_workers"],
                        pool_maxsize=st.session_state["max_workers"],
                        max_retries=2
                    )
                    session.mount("http://", adapter)
                    session.mount("https://", adapter)
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

                    prog = st.progress(0)
                    status = st.empty()
                    resultados, falhas, done = {}, 0, 0
                    failed_urls = []
                    failed_details = []

                    with ThreadPoolExecutor(max_workers=st.session_state["max_workers"]) as ex:
                        futures = {
                            ex.submit(
                                baixar_processar, session, url,
                                st.session_state["target_w"], st.session_state["target_h"],
                                st.session_state["limite_kb"], st.session_state["req_timeout"],
                                fx_cfg
                            ): (loja, endereco, url, url_line_map.get(url, "?"))
                            for loja, endereco, url in items
                        }
                        for fut in as_completed(futures):
                            loja, endereco, url, line_no = futures[fut]
                            try:
                                res = fut.result()
                            except Exception as e:
                                logger.error(f"Erro em download {url}: {e}")
                                res = (url, False, None, None, None, None, None, f"Exception: {e}")

                            if res and isinstance(res, (list, tuple)) and len(res) >= 2 and res[1] is True:
                                # Sucesso
                                url_key = res[0]
                                file_path = res[2] if len(res) > 2 else None
                                wh = res[3] if len(res) > 3 else (0, 0)
                                quality = res[4] if len(res) > 4 else {}
                                sha1_hex = res[5] if len(res) > 5 else ""
                                dhash_hex = res[6] if len(res) > 6 else ""

                                if (not file_path) or (not wh) or (not isinstance(wh, (list, tuple))):
                                    falhas += 1
                                    failed_urls.append(url)
                                    error_msg = res[7] if len(res) > 7 else "Erro desconhecido"
                                    failed_details.append({
                                        "url": url,
                                        "loja": loja,
                                        "linha": line_no,
                                        "erro": error_msg
                                    })
                                else:
                                    resultados[url_key] = (loja, endereco, file_path, wh, quality, sha1_hex, dhash_hex)
                            else:
                                # Falha
                                falhas += 1
                                failed_urls.append(url)
                                error_msg = res[7] if len(res) > 7 else "Erro desconhecido" if len(res) > 7 else "Erro desconhecido"
                                failed_details.append({
                                    "url": url,
                                    "loja": loja,
                                    "linha": line_no,
                                    "erro": error_msg
                                })

                            done += 1
                            prog.progress(int(done * 100 / total))
                            status.write(f"Processadas {done}/{total} imagens... (Falhas: {falhas})")

                    status.write(f"Concluído. Falhas: {falhas}")

                    # Salvar informações de falha
                    st.session_state.failed_urls = failed_urls
                    st.session_state.failed_details = failed_details
                    st.session_state.url_line_map = url_line_map

                    # Verificar se há falhas e se devemos continuar
                    if falhas > 0:
                        if not st.session_state.get("ignore_failed", True):
                            st.error(f"{falhas} imagem(ns) falharam. Desative 'Ignorar falhas' na sidebar se quiser parar.")
                            st.stop()
                        else:
                            st.warning(f"{falhas} imagem(ns) foram puladas por erro, mas continuando com as {len(resultados)} restantes.")
                            
                            # CORREÇÃO: Criar um container para as falhas sem usar expander dentro de expander
                            if failed_details:
                                # Usar uma div expansível personalizada ou simplesmente mostrar os dados
                                st.markdown(f"**📋 {len(failed_details)} Imagem(ns) com erro:**")
                                
                                # Criar DataFrame para melhor visualização
                                df_failed = pd.DataFrame(failed_details)
                                
                                # Ordenar por linha
                                if 'linha' in df_failed.columns:
                                    try:
                                        df_failed['linha_num'] = pd.to_numeric(df_failed['linha'], errors='coerce')
                                        df_failed = df_failed.sort_values('linha_num')
                                        df_failed = df_failed.drop(columns=['linha_num'])
                                    except:
                                        df_failed = df_failed.sort_values('linha')
                                
                                # Exibir tabela com scroll
                                container = st.container()
                                with container:
                                    st.dataframe(
                                        df_failed,
                                        column_config={
                                            "linha": "Linha",
                                            "loja": "Loja",
                                            "url": st.column_config.LinkColumn("URL"),
                                            "erro": "Erro"
                                        },
                                        hide_index=True,
                                        use_container_width=True,
                                        height=min(300, 35 * min(10, len(failed_details)))
                                    )
                                    
                                    # Botão para copiar lista
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        if st.button("📋 Copiar lista de falhas", key="copy_failed_list"):
                                            text_to_copy = "\n".join([f"Linha {d['linha']}: {d['url']} - {d['erro']}" 
                                                                     for d in failed_details])
                                            st.code(text_to_copy, language="text")
                                            st.success("Lista copiada para a área de transferência!")
                                    with col2:
                                        if st.button("📁 Exportar falhas (CSV)", key="export_failed_csv"):
                                            csv = df_failed.to_csv(index=False).encode('utf-8')
                                            st.download_button(
                                                label="Baixar CSV",
                                                data=csv,
                                                file_name="falhas_download.csv",
                                                mime="text/csv",
                                                key="download_failed_csv"
                                            )

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
                            "signature_width_in": st.session_state.get("signature_width_in",
                                                   (st.session_state["logo_width_in"]/2.0)),
                            "signature_right_margin_in": st.session_state.get("sig_right_margin", 0.20),
                            "signature_bottom_margin_in": st.session_state.get("sig_bottom_margin", 0.20),
                            "thumb_px": st.session_state["thumb_px"],
                            "thumbs_per_row": st.session_state["thumbs_per_row"],
                            "effects": fx_cfg,
                            "use_template": st.session_state.get("use_template", False),
                            "template_bytes": template_file.getvalue()
                                if (st.session_state.get("use_template") and template_file) else None,
                            "low_quality_urls": list(low_q),
                            "duplicate_urls": list(dups),
                            "ignore_failed": st.session_state.get("ignore_failed", True),
                        }
                    }

                    # DIFERENÇA CRÍTICA: Define o modo baseado no botão clicado
                    if btn_preview:
                        st.session_state.preview_mode = True
                        st.session_state.quick_generate = False
                    else:  # btn_generate_direct
                        st.session_state.preview_mode = False  
                        st.session_state.quick_generate = True
                        st.session_state.generated = False

                    st.session_state.ppt_bytes = None
                    st.session_state.images_zip_bytes = None

                    st.session_state.exp_plan = False
                    st.session_state.exp_style = False
                    st.session_state.exp_brand = False
                    st.session_state.exp_fx = False
                    st.session_state.exp_perf = False
                    st.session_state.exp_model = False
                    st.rerun()

    # 2) Pré-visualização - MODIFICADO: Condicional mais inteligente
    if st.session_state.preview_mode and st.session_state.pipeline and not st.session_state.quick_generate:
        preview_expander = st.expander("2. Pré-visualização", expanded=True)
        
        with preview_expander:
            p = st.session_state.pipeline

            # Mostrar estatísticas
            stats = render_summary(p["items"], p["resultados"], st.session_state.excluded_urls, st.session_state.get("failed_details", []))
            
            # Aviso se houver falhas
            if st.session_state.get("failed_details") and st.session_state.get("ignore_failed", True):
                st.info(f"""
                ⚠️ **Atenção:** {len(st.session_state.failed_details)} imagem(ns) falharam no download.
                O book será gerado apenas com as {stats['baixadas']} imagens que foram baixadas com sucesso.
                """)
                
                # Opção para ver detalhes das falhas - CORREÇÃO: Usar checkbox para controlar visibilidade
                show_failed_details = st.checkbox("Mostrar detalhes das falhas", key="show_failed_details")
                
                if show_failed_details and st.session_state.failed_details:
                    st.markdown("---")
                    st.markdown(f"**📋 Detalhes das {len(st.session_state.failed_details)} falhas:**")
                    
                    df_failed = pd.DataFrame(st.session_state.failed_details)
                    if 'linha' in df_failed.columns:
                        try:
                            df_failed['linha_num'] = pd.to_numeric(df_failed['linha'], errors='coerce')
                            df_failed = df_failed.sort_values('linha_num')
                            df_failed = df_failed.drop(columns=['linha_num'])
                        except:
                            df_failed = df_failed.sort_values('linha')
                    
                    # Exibir tabela com altura limitada
                    st.dataframe(
                        df_failed,
                        column_config={
                            "linha": "Linha",
                            "loja": "Loja",
                            "url": st.column_config.LinkColumn("URL"),
                            "erro": "Erro"
                        },
                        hide_index=True,
                        use_container_width=True,
                        height=min(400, 35 * min(15, len(st.session_state.failed_details)))
                    )
                    
                    # Botões de ação
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("📋 Copiar lista", key="copy_failed_preview"):
                            text_to_copy = "\n".join([f"Linha {d['linha']}: {d['url']} - {d['erro']}" 
                                                     for d in st.session_state.failed_details])
                            st.code(text_to_copy, language="text")
                            st.success("Lista copiada para a área de transferência!")
                    with col2:
                        if st.button("📁 Exportar CSV", key="export_failed_preview"):
                            csv = df_failed.to_csv(index=False).encode('utf-8')
                            st.download_button(
                                label="Baixar CSV",
                                data=csv,
                                file_name="falhas_detalhadas.csv",
                                mime="text/csv",
                                key="download_failed_preview"
                            )

            render_preview(
                p["items"], p["resultados"],
                p["settings"]["sort_mode"],
                p["settings"]["thumb_px"],
                p["settings"]["thumbs_per_row"]
            )
            st.info("Marque **Excluir esta foto** nas imagens que não devem ir para o PPT/ZIP. Depois, avance para a etapa 3.")

    # 3) Gerar / Exportar - MODIFICADO: Lógica otimizada
    generate_expander = st.expander("3. Gerar / Exportar", expanded=st.session_state.get("preview_mode", False) or st.session_state.get("quick_generate", False))
    
    with generate_expander:
        if not st.session_state.pipeline:
            st.info("Faça o upload da planilha e processe as imagens primeiro.")
        else:
            cfg = st.session_state.pipeline["settings"]
            items = st.session_state.pipeline["items"]
            resultados = st.session_state.pipeline["resultados"]
            
            # Mostrar estatísticas
            stats = render_summary(items, resultados, st.session_state.excluded_urls, st.session_state.get("failed_details", []))

            # GERAÇÃO DIRETA SE SOLICITADA
            if st.session_state.quick_generate and not st.session_state.get("ppt_bytes"):
                with st.spinner("🚀 Gerando PPT diretamente..."):
                    try:
                        # Aviso se houver falhas
                        if st.session_state.get("failed_details") and st.session_state.get("ignore_failed", True):
                            st.info(f"ℹ️ Gerando com {stats['falhas']} falha(s) ignorada(s). Total de imagens no book: {stats['baixadas']}")
                        
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
                            signature_width = (cfg["logo_width_in"]/2.0) if cfg.get("auto_half_signature", True) \
                                              else (cfg.get("signature_width_in") or 0.6)

                            groups = OrderedDict()
                            for loja, endereco, url in items:
                                if url in resultados and url not in st.session_state.excluded_urls:
                                    groups.setdefault(str(loja), []).append((url, resultados[url]))

                            if cfg["sort_mode"] == "Nome da loja (A→Z)":
                                loja_keys = sorted(groups.keys(),
                                    key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
                            else:
                                loja_keys = list(groups.keys())

                            for loja in loja_keys:
                                imgs = groups[loja]; i = 0
                                while i < len(imgs):
                                    if cfg["max_per_slide"] == "Automático":
                                        _url0, (_loja0, endereco, _file0, (w0, h0), *_rest0) = imgs[i]
                                        per_slide = 3 if is_portrait(w0, h0) else 2
                                    else:
                                        per_slide = int(cfg["max_per_slide"])
                                        endereco = imgs[i][1][1]

                                    batch = imgs[i:i+per_slide]; i += per_slide
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
                                    for (url, (_loja, _end, file_path, (w_px, h_px), *rest)), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                                        place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in)

                            out = BytesIO(); prs.save(out); out.seek(0); ppt_bytes = out

                        st.session_state.ppt_bytes = ppt_bytes
                        st.session_state.generated = True
                        st.session_state.quick_generate = False
                        st.success(f"✅ PPT gerado com sucesso! Total de slides gerados: {len(groups)} lojas processadas.")
                        st.rerun()
                    except Exception as e:
                        logger.exception("Falha ao gerar PPT")
                        st.error(f"Falha ao gerar PPT: {e}")
                        st.session_state.quick_generate = False

            # INTERFACE DE DOWNLOAD (comum para ambos os fluxos)
            if st.session_state.pipeline:
                # Apenas mostra prévia se não for geração direta
                if not st.session_state.quick_generate:
                    r1, r2 = st.columns([1, 5])
                    with r1:
                        if st.button("🔄 Recriar Prévia", use_container_width=True):
                            st.session_state.preview_bump = st.session_state.get("preview_bump", 0) + 1
                            st.rerun()
                    with r2:
                        st.caption("Gera novamente as miniaturas dos slides com base nas exclusões/ordem atuais.")

                    batches = build_batches_for_preview(items, resultados, cfg, st.session_state.excluded_urls)
                    st.subheader("🔎 Prévia dos primeiros slides")
                    if len(batches) == 0:
                        st.warning("Nenhum slide seria gerado com as configurações atuais.")
                    else:
                        preview_count = min(3, len(batches))
                        cols = st.columns(preview_count)
                        for idx in range(preview_count):
                            loja, end, batch = batches[idx]
                            canvas = compose_slide_preview(batch, loja, end, cfg)  # 1280x720 RGB
                            canvas = canvas.convert("RGBA")
                            W, H = canvas.width, canvas.height
                            title_rgb = cfg.get("title_font_color_rgb", pick_contrast_color(*cfg["bg_rgb"]))

                            draw = ImageDraw.Draw(canvas)
                            top_bar_h = int(110)
                            draw.rectangle([0, 0, W, top_bar_h], fill=cfg["bg_rgb"])
                            try:
                                font_title = ImageFont.truetype("arial.ttf", 28)
                                font_addr  = ImageFont.truetype("arial.ttf", 16)
                            except Exception:
                                font_title = ImageFont.load_default()
                                font_addr  = ImageFont.load_default()

                            slide_w_in, slide_h_in = 13.33, 7.5
                            x_left = int(W * (0.30 / slide_w_in))
                            y_top  = int(H * (0.20 / slide_h_in))
                            draw.text((x_left, y_top), str(loja), fill=title_rgb, font=font_title, anchor="la")
                            if end:
                                try:
                                    _, _, tw, th = draw.textbbox((0, 0), str(loja), font=font_title)
                                    y_addr = y_top + th + 6
                                except Exception:
                                    y_addr = y_top + 32
                                draw.text((x_left, y_addr), str(end), fill=title_rgb, font=font_addr, anchor="la")

                            if cfg.get("logo_bytes"):
                                try:
                                    _bio = BytesIO(cfg["logo_bytes"])
                                    logo_im = Image.open(_bio).convert("RGBA")
                                    target_w_px = max(1, int(W * (cfg["logo_width_in"] / slide_w_in)))
                                    ratio = target_w_px / float(logo_im.width)
                                    logo_im = logo_im.resize((target_w_px, int(logo_im.height * ratio)), Image.LANCZOS)
                                    x = W - int(W * (0.50 / slide_w_in)) - logo_im.width
                                    y = int(H * (0.20 / slide_h_in))
                                    canvas.alpha_composite(logo_im, (max(0, x), max(0, y)))
                                except Exception:
                                    pass

                            if cfg.get("signature_bytes"):
                                try:
                                    _bio = BytesIO(cfg["signature_bytes"])
                                    sig_im = Image.open(_bio).convert("RGBA")
                                    if cfg.get("auto_half_signature", True):
                                        sig_w_in = (cfg.get("logo_width_in", 1.2) / 2.0)
                                    else:
                                        sig_w_in = cfg.get("signature_width_in") or 0.6
                                    target_w_px = max(1, int(W * (sig_w_in / slide_w_in)))
                                    ratio = target_w_px / float(sig_im.width)
                                    sig_im = sig_im.resize((target_w_px, int(sig_im.height * ratio)), Image.LANCZOS)
                                    right_margin_in  = float(cfg.get("signature_right_margin_in", 0.20))
                                    bottom_margin_in = float(cfg.get("signature_bottom_margin_in", 0.20))
                                    x = W - int(W * (right_margin_in / slide_w_in)) - sig_im.width
                                    y = H - int(H * (bottom_margin_in / slide_h_in)) - sig_im.height
                                    canvas.alpha_composite(sig_im, (max(0, x), max(0, y)))
                                except Exception:
                                    pass

                            with cols[idx]:
                                st.image(canvas.convert("RGB"), caption=f"Slide {idx+1} — {loja}", use_column_width=True)

                # CONTROLES DE DOWNLOAD
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.session_state.output_filename = st.text_input(
                        "Nome base do arquivo (sem extensão)",
                        value=st.session_state.get("output_filename", "Modelo_01"),
                        key="output_filename_input"
                    )

                with col2:
                    if st.session_state.get("ppt_bytes"):
                        st.download_button(
                            "⬇️ Baixar PPT",
                            data=st.session_state.ppt_bytes,
                            file_name=f"{(st.session_state.output_filename or 'Apresentacao')}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentation.presentation",
                            use_container_width=True,
                            key=f"download_{st.session_state.get('download_key', 0)}"
                        )
                    elif not st.session_state.quick_generate:  # Só mostra botão de gerar se não for fluxo direto
                        btn_generate = st.button("🧩 Gerar PPT", key="btn_generate", use_container_width=True)

                with col3:
                    if st.session_state.get("images_zip_bytes"):
                        st.download_button(
                            "⬇️ Baixar Imagens (ZIP)",
                            data=st.session_state.images_zip_bytes,
                            file_name=f"{(st.session_state.output_filename or 'Imagens')}.zip",
                            mime="application/zip",
                            use_container_width=True,
                            key=f"images_zip_{st.session_state.get('images_zip_key', 0)}"
                        )
                    elif not st.session_state.quick_generate:  # Só mostra botão de gerar ZIP se não for fluxo direto
                        btn_zip = st.button("🖼️ Baixar Imagens", key="btn_zip", use_container_width=True)

                # Geração do PPT (apenas para fluxo de visualização)
                if (not st.session_state.get("ppt_bytes")) and ('btn_generate' in locals()) and btn_generate and not st.session_state.quick_generate:
                    try:
                        # Aviso se houver falhas
                        if st.session_state.get("failed_details") and st.session_state.get("ignore_failed", True):
                            st.info(f"ℹ️ Gerando com {stats['falhas']} falha(s) ignorada(s). Total de imagens no book: {stats['baixadas']}")
                        
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
                            signature_width = (cfg["logo_width_in"]/2.0) if cfg.get("auto_half_signature", True) \
                                              else (cfg.get("signature_width_in") or 0.6)

                            groups = OrderedDict()
                            for loja, endereco, url in items:
                                if url in resultados and url not in st.session_state.excluded_urls:
                                    groups.setdefault(str(loja), []).append((url, resultados[url]))

                            if cfg["sort_mode"] == "Nome da loja (A→Z)":
                                loja_keys = sorted(groups.keys(),
                                    key=lambda s: (s is None or str(s).strip()== "", (s or "").strip().casefold()))
                            else:
                                loja_keys = list(groups.keys())

                            for loja in loja_keys:
                                imgs = groups[loja]; i = 0
                                while i < len(imgs):
                                    if cfg["max_per_slide"] == "Automático":
                                        _url0, (_loja0, endereco, _file0, (w0, h0), *_rest0) = imgs[i]
                                        per_slide = 3 if is_portrait(w0, h0) else 2
                                    else:
                                        per_slide = int(cfg["max_per_slide"])
                                        endereco = imgs[i][1][1]

                                    batch = imgs[i:i+per_slide]; i += per_slide
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
                                    for (url, (_loja, _end, file_path, (w_px, h_px), *rest)), (left, top, max_w_in, max_h_in) in zip(batch, slots):
                                        place_picture(slide, file_path, w_px, h_px, left, top, max_w_in, max_h_in)

                            out = BytesIO(); prs.save(out); out.seek(0); ppt_bytes = out

                        st.session_state.ppt_bytes = ppt_bytes
                        st.session_state.generated = True
                        st.success(f"✅ PPT gerado com sucesso! Total de slides gerados: {len(groups)} lojas processadas.")
                        st.rerun()
                    except Exception as e:
                        logger.exception("Falha ao gerar PPT")
                        st.error(f"Falha ao gerar PPT: {e}")

                # Geração do ZIP (apenas para fluxo de visualização)
                if (not st.session_state.get("images_zip_bytes")) and ('btn_zip' in locals()) and btn_zip and not st.session_state.quick_generate:
                    try:
                        zip_bytes = montar_zip_imagens(
                            items=items,
                            resultados=resultados,
                            excluded_urls=st.session_state.excluded_urls
                        )
                        st.session_state.images_zip_bytes = zip_bytes
                        st.success(f"✅ ZIP gerado com sucesso! Total de imagens incluídas: {sum(1 for url in resultados if url not in st.session_state.excluded_urls)}")
                        st.rerun()
                    except Exception as e:
                        logger.exception("Falha ao montar ZIP")
                        st.error(f"Falha ao montar ZIP: {e}")

# -------------------------------------------------------------------
# ROTEAMENTO FINAL
# -------------------------------------------------------------------
if not st.session_state.auth:
    do_login()
else:
    main_app()
