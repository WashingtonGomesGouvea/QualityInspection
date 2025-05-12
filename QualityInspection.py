# QualityInspection_V2.py
import streamlit as st
import json
import pandas as pd
from datetime import datetime, timedelta
import os
from io import BytesIO
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File as SPFile # Renomear para evitar conflito
from urllib.parse import urlparse, parse_qs

# --- Configuração da Página Streamlit ---
st.set_page_config(layout="wide", page_title="Formulário de Inspeção Dinâmico V2")

# --- Constantes e Configurações ---
ROTEIROS_LOCAL_PATH = "/home/ubuntu/roteiros_final_v4.json"
UPLOAD_DIR_BASE = "uploads_inspecoes_temp" 

# --- Funções de Utilitários e Lógica de Negócio ---

@st.cache_resource(ttl=600) 
def get_sharepoint_context():
    """Carrega as credenciais e retorna um contexto do SharePoint autenticado."""
    try:
        email = st.secrets.sharepoint.email
        password = st.secrets.sharepoint.password
        site_url = st.secrets.sharepoint.site_url
        if not email or not password or not site_url:
            st.error("Credenciais do SharePoint não configuradas corretamente nos secrets.")
            return None
        return ClientContext(site_url).with_credentials(UserCredential(email, password))
    except AttributeError:
        st.error("Erro ao aceder aos secrets do SharePoint. Verifique o seu ficheiro secrets.toml.")
        return None
    except Exception as e:
        st.error(f"Falha ao conectar ao SharePoint: {e}")
        return None

@st.cache_data(ttl=300) 
def load_roteiros_config(local_fallback_path=ROTEIROS_LOCAL_PATH):
    """Tenta carregar a configuração dos roteiros do SharePoint, com fallback local."""
    ctx = get_sharepoint_context()
    roteiros_sp_url = None
    try:
        roteiros_sp_url = st.secrets.sharepoint.roteiros_file_url
    except AttributeError:
        st.warning("URL do ficheiro de roteiros no SharePoint (roteiros_file_url) não encontrado nos secrets.")

    if ctx and roteiros_sp_url:
        try:
            parsed_url = urlparse(roteiros_sp_url)
            query_params = parse_qs(parsed_url.query)
            unique_id = query_params.get("UniqueId", [None])[0]

            if not unique_id:
                raise ValueError("Não foi possível extrair o UniqueId do roteiros_file_url.")

            file_content_bytes = BytesIO()
            sp_file = ctx.web.get_file_by_id(unique_id)
            sp_file.download(file_content_bytes).execute_query()
            file_content_bytes.seek(0)
            json_data = json.loads(file_content_bytes.read().decode("utf-8"))
            st.info("Roteiros carregados do SharePoint com sucesso!")
            return json_data
        except Exception as e_sp:
            st.warning(f"Falha ao carregar roteiros do SharePoint ({roteiros_sp_url}): {e_sp}. A tentar fallback local.")
    else:
        st.info("Contexto do SharePoint ou URL dos roteiros não disponível. A tentar fallback local.")

    try:
        with open(local_fallback_path, "r", encoding="utf-8") as f:
            config = json.load(f)
            st.info(f"Roteiros carregados do ficheiro local: {local_fallback_path}")
            return config
    except FileNotFoundError:
        st.error(f"Ficheiro de roteiros local não encontrado em: {local_fallback_path}")
        return None
    except json.JSONDecodeError:
        st.error(f"Erro ao descodificar o JSON do ficheiro de roteiros local: {local_fallback_path}")
        return None
    except Exception as e_local:
        st.error(f"Erro inesperado ao carregar roteiros do ficheiro local ({local_fallback_path}): {e_local}")
        return None

# --- Inicialização da Aplicação ---
if "dados_formulario_atual" not in st.session_state:
    st.session_state.dados_formulario_atual = {}
if "form_key_counter" not in st.session_state:
    st.session_state.form_key_counter = 0
if "inspecoes_realizadas_sessao" not in st.session_state:
    st.session_state.inspecoes_realizadas_sessao = []
if "current_form_uploads" not in st.session_state:
    st.session_state.current_form_uploads = {}
if "solucao_data_preparo_dinamico" not in st.session_state:
    st.session_state.solucao_data_preparo_dinamico = datetime.now().date()
if "solucao_tipo_dinamico" not in st.session_state:
    st.session_state.solucao_tipo_dinamico = None

if not os.path.exists(UPLOAD_DIR_BASE):
    try:
        os.makedirs(UPLOAD_DIR_BASE)
    except OSError as e:
        st.warning(f"Não foi possível criar o diretório de uploads temporários: {UPLOAD_DIR_BASE}. Erro: {e}")

roteiros_config = load_roteiros_config()

# --- Funções da Barra Lateral e Seleção de Roteiro ---
def render_sidebar_form(roteiros_cfg):
    form_key = f"form_inspecao_sidebar_{st.session_state.form_key_counter}"
    with st.sidebar.form(key=form_key, clear_on_submit=False):
        st.subheader("Dados do Inspetor e Local")
        for campo_info in roteiros_cfg.get("informacoes_iniciais", []) :
            label = campo_info["label"]
            key_sidebar = f"inicial_{campo_info["key"]}" 
            obrigatorio = campo_info.get("obrigatorio", False)
            
            if campo_info["tipo"] == "texto":
                subtipo = campo_info.get("subtipo", "text")
                input_type = "password" if subtipo == "password" else "default"
                st.session_state.dados_formulario_atual[campo_info["key"]] = st.text_input(
                    label + ("*" if obrigatorio else ""), 
                    key=key_sidebar, 
                    type=input_type,
                    value=st.session_state.dados_formulario_atual.get(campo_info["key"], "")
                )
            elif campo_info["tipo"] == "data":
                default_date_val = datetime.now().date()
                current_val_str = st.session_state.dados_formulario_atual.get(campo_info["key"])
                if current_val_str:
                    try:
                        default_date_val = datetime.strptime(str(current_val_str), "%Y-%m-%d").date()
                    except (ValueError, TypeError):
                        pass 
                st.session_state.dados_formulario_atual[campo_info["key"]] = str(st.date_input(
                    label + ("*" if obrigatorio else ""), 
                    key=key_sidebar,
                    value=default_date_val
                ))
        
        st.subheader("Seleção do Setor e Processo de Inspeção")
        setores_nomes = [setor["nome"] for setor in roteiros_cfg.get("setores_inspecao", [])]
        default_setor_idx = 0
        if st.session_state.dados_formulario_atual.get("setor_selecionado_nome") in setores_nomes:
            default_setor_idx = setores_nomes.index(st.session_state.dados_formulario_atual["setor_selecionado_nome"])
        
        setor_selecionado_nome = st.selectbox(
            "Setor Inspecionado*", setores_nomes, key="sidebar_setor_selecionado_nome", index=default_setor_idx
        )
        st.session_state.dados_formulario_atual["setor_selecionado_nome"] = setor_selecionado_nome
        setor_obj = next((s for s in roteiros_cfg.get("setores_inspecao", []) if s["nome"] == setor_selecionado_nome), None)
        st.session_state.dados_formulario_atual["setor_selecionado_key"] = setor_obj["key"] if setor_obj else None

        processos_nomes = [proc["nome"] for proc in setor_obj.get("processos", [])] if setor_obj else []
        default_proc_idx = 0
        if st.session_state.dados_formulario_atual.get("processo_selecionado_nome") in processos_nomes:
            default_proc_idx = processos_nomes.index(st.session_state.dados_formulario_atual["processo_selecionado_nome"])
        
        processo_selecionado_nome = st.selectbox(
            "Processo a ser Inspecionado*", processos_nomes, key="sidebar_processo_selecionado_nome", 
            index=default_proc_idx, disabled=not bool(setor_obj)
        )
        st.session_state.dados_formulario_atual["processo_selecionado_nome"] = processo_selecionado_nome
        processo_obj = None
        if setor_obj and processo_selecionado_nome:
             processo_obj = next((p for p in setor_obj.get("processos", []) if p["nome"] == processo_selecionado_nome), None)
        st.session_state.dados_formulario_atual["processo_selecionado_key"] = processo_obj["key"] if processo_obj else None

        submitted_sidebar = st.form_submit_button("Carregar Roteiro de Inspeção")
        if submitted_sidebar:
            campos_obrigatorios_sidebar = [ci["key"] for ci in roteiros_cfg.get("informacoes_iniciais", []) if ci.get("obrigatorio")]
            if not setor_selecionado_nome: campos_obrigatorios_sidebar.append("setor_selecionado_nome") 
            if not processo_selecionado_nome: campos_obrigatorios_sidebar.append("processo_selecionado_nome")
            
            faltando = []
            for k_obr_key in campos_obrigatorios_sidebar:
                valor_atual = st.session_state.dados_formulario_atual.get(k_obr_key)
                label_campo_obr = k_obr_key
                if k_obr_key.startswith("inicial_"):
                    original_key_obr = k_obr_key.split("inicial_",1)[1]
                    campo_info_obj_obr = next((ci for ci in roteiros_cfg.get("informacoes_iniciais", []) if ci["key"] == original_key_obr), None)
                    if campo_info_obj_obr: label_campo_obr = campo_info_obj_obr["label"]
                elif k_obr_key == "setor_selecionado_nome": label_campo_obr = "Setor Inspecionado"
                elif k_obr_key == "processo_selecionado_nome": label_campo_obr = "Processo a ser Inspecionado"
                
                if isinstance(valor_atual, str) and not valor_atual.strip():
                    faltando.append(label_campo_obr)
                elif valor_atual is None:
                     faltando.append(label_campo_obr)
            
            if faltando:
                st.sidebar.error(f"Por favor, preencha os seguintes campos obrigatórios: {", ".join(faltando)}")
                st.session_state.dados_formulario_atual["sidebar_completa"] = False
            else:
                st.session_state.dados_formulario_atual["sidebar_completa"] = True
                st.session_state.dados_formulario_atual["respostas_roteiro"] = {} 
                st.session_state.current_form_uploads = {}
                st.session_state.solucao_data_preparo_dinamico = datetime.now().date()
                st.session_state.solucao_tipo_dinamico = None
                st.rerun()
    return setor_obj, processo_obj

# --- Funções de Renderização do Formulário Dinâmico, Excel e SharePoint ---
def calculate_due_date(start_date_str, rule_key, roteiros_cfg):
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
        rules = roteiros_cfg.get("regras_validade_solucoes", {})
        rule_details = rules.get(rule_key)
        if rule_details:
            if rule_details["unidade"] == "dias": return start_date + timedelta(days=rule_details["valor"])
            elif rule_details["unidade"] == "meses": return start_date + timedelta(days=rule_details["valor"] * 30) 
        return None
    except (ValueError, TypeError): return None

def normalizar_dados_para_excel(inspecao_data_list):
    dados_planos = []
    for inspecao in inspecao_data_list:
        base_info = {
            "Data/Hora da Submissão": inspecao.get("data_hora_submissao", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            "Nome do Inspetor": inspecao.get("nome_inspetor"),
            "Email do Inspetor": inspecao.get("email_inspetor"),
            "Empresa Inspecionada": inspecao.get("empresa_inspecionada"),
            "Data da Inspeção (Preenchida)": inspecao.get("data_inspecao"),
            "Setor Inspecionado": inspecao.get("setor_selecionado_nome"),
            "Processo Inspecionado": inspecao.get("processo_selecionado_nome")
        }
        respostas_roteiro = inspecao.get("respostas_roteiro", {})
        for chave_campo, valor_campo in respostas_roteiro.items():
            base_info[chave_campo] = ", ".join(map(str, valor_campo)) if isinstance(valor_campo, list) else valor_campo
        
        evidencias_info = inspecao.get("evidencias_info", {})
        for label_ev, nomes_ficheiros in evidencias_info.items():
            base_info[f"Evidência: {label_ev}"] = ", ".join(nomes_ficheiros)
        dados_planos.append(base_info)
    return dados_planos

def gerar_excel_bytes(inspecoes_data):
    if not inspecoes_data: return None
    dados_planos = normalizar_dados_para_excel(inspecoes_data if isinstance(inspecoes_data, list) else [inspecoes_data])
    df = pd.DataFrame(dados_planos)
    colunas_principais = [
        "Data/Hora da Submissão", "Nome do Inspetor", "Email do Inspetor", 
        "Empresa Inspecionada", "Data da Inspeção (Preenchida)", 
        "Setor Inspecionado", "Processo Inspecionado"
    ]
    colunas_respostas_evidencias = sorted([col for col in df.columns if col not in colunas_principais])
    df = df.reindex(columns=colunas_principais + colunas_respostas_evidencias)
    excel_bytes = BytesIO()
    df.to_excel(excel_bytes, index=False, engine="openpyxl")
    excel_bytes.seek(0)
    return excel_bytes

def salvar_inspecao_sharepoint(dados_inspecao, ctx_sp_save):
    if not ctx_sp_save: 
        st.error("Contexto do SharePoint não disponível para salvar.")
        return False
    try:
        target_folder_relative_url = "Documents/Inspeção Qualidade"
        filename = st.secrets.sharepoint.get("historico_inspecoes_filename", "registos_inspecoes_V2.xlsx")
        target_file_url_relative = f"{target_folder_relative_url}/{filename}"

        df_nova_inspecao = pd.DataFrame(normalizar_dados_para_excel([dados_inspecao]))
        df_final = df_nova_inspecao
        try:
            file_content_bytes = BytesIO()
            sp_file_obj = ctx_sp_save.web.get_file_by_server_relative_path(target_file_url_relative)
            sp_file_obj.download(file_content_bytes).execute_query()
            file_content_bytes.seek(0)
            if file_content_bytes.getbuffer().nbytes > 0:
                df_existente = pd.read_excel(file_content_bytes, engine="openpyxl")
                df_final = pd.concat([df_existente, df_nova_inspecao], ignore_index=True)
        except Exception:
            st.info(f"Ficheiro 	'{filename}' não encontrado ou vazio no SharePoint. Um novo será criado.")
        
        colunas_principais_sp = [
            "Data/Hora da Submissão", "Nome do Inspetor", "Email do Inspetor", 
            "Empresa Inspecionada", "Data da Inspeção (Preenchida)", 
            "Setor Inspecionado", "Processo Inspecionado"
        ]
        colunas_respostas_evidencias_sp = sorted([col for col in df_final.columns if col not in colunas_principais_sp])
        df_final = df_final.reindex(columns=colunas_principais_sp + colunas_respostas_evidencias_sp)

        output_excel_bytes = BytesIO()
        df_final.to_excel(output_excel_bytes, index=False, engine="openpyxl")
        output_excel_bytes.seek(0)

        target_folder = ctx_sp_save.web.get_folder_by_server_relative_path(target_folder_relative_url)
        target_folder.upload_file(filename, output_excel_bytes).execute_query()
        st.success(f"Inspeção salva com sucesso no SharePoint: {filename}")
        return True
    except Exception as e:
        st.error(f"Falha ao salvar inspeção no SharePoint: {e}")
        return False

def render_dynamic_form(processo_obj, roteiros_cfg_form, ctx_sp_form):
    if not processo_obj or "campos" not in processo_obj: return

    # --- Lógica para campos de solução fora do formulário (para reatividade) ---
    is_solucoes_process = processo_obj["key"].startswith("solucoes")
    campos_solucao_fora_form = []
    campos_solucao_dentro_form = []

    if is_solucoes_process:
        for campo_roteiro in processo_obj["campos"]:
            if campo_roteiro.get("key") in ["solucao_data_preparo", "solucao_tipo"]:
                campos_solucao_fora_form.append(campo_roteiro)
            else:
                campos_solucao_dentro_form.append(campo_roteiro)
    else:
        campos_solucao_dentro_form = processo_obj["campos"]

    # Renderizar campos de solução que precisam de reatividade FORA do formulário
    if is_solucoes_process and campos_solucao_fora_form:
        st.markdown("### Detalhes da Solução (Cálculo de Validade)")
        for campo_solucao_ext in campos_solucao_fora_form:
            label_ext = campo_solucao_ext["label"]
            key_ext = f"widget_ext_{processo_obj['key']}_{campo_solucao_ext['key']}_{st.session_state.form_key_counter}"
            obrigatorio_ext = campo_solucao_ext.get("obrigatorio", False)
            st.markdown(f"**{label_ext}**" + ("*" if obrigatorio_ext else ""))

            if campo_solucao_ext.get("key") == "solucao_data_preparo":
                st.session_state.solucao_data_preparo_dinamico = st.date_input(
                    "", 
                    value=st.session_state.solucao_data_preparo_dinamico, 
                    key=key_ext, 
                    label_visibility="collapsed"
                )
            elif campo_solucao_ext.get("key") == "solucao_tipo":
                opcoes_tipo_sol = campo_solucao_ext.get("opcoes", [])
                idx_tipo_sol = 0
                if st.session_state.solucao_tipo_dinamico and st.session_state.solucao_tipo_dinamico in opcoes_tipo_sol:
                    idx_tipo_sol = opcoes_tipo_sol.index(st.session_state.solucao_tipo_dinamico)
                elif opcoes_tipo_sol: # Define um padrão se nenhum estiver no estado e houver opções
                    st.session_state.solucao_tipo_dinamico = opcoes_tipo_sol[0]
                
                st.session_state.solucao_tipo_dinamico = st.selectbox(
                    "", 
                    opcoes_tipo_sol, 
                    index=idx_tipo_sol, 
                    key=key_ext, 
                    label_visibility="collapsed"
                )
        st.markdown("---") # Separador visual

    # --- Início do Formulário Principal ---
    with st.form(key="roteiro_form_main"): 
        respostas_roteiro = st.session_state.dados_formulario_atual.setdefault("respostas_roteiro", {})
        
        # Adicionar os valores dos campos de solução (de fora do form) às respostas para salvar
        if is_solucoes_process:
            campo_data_preparo_cfg = next((c for c in campos_solucao_fora_form if c.get("key") == "solucao_data_preparo"), None)
            campo_tipo_sol_cfg = next((c for c in campos_solucao_fora_form if c.get("key") == "solucao_tipo"), None)
            if campo_data_preparo_cfg:
                respostas_roteiro[campo_data_preparo_cfg["label"]] = str(st.session_state.solucao_data_preparo_dinamico)
            if campo_tipo_sol_cfg:
                respostas_roteiro[campo_tipo_sol_cfg["label"]] = st.session_state.solucao_tipo_dinamico

        for campo_roteiro in campos_solucao_dentro_form:
            label_roteiro = campo_roteiro["label"]
            widget_key = f"widget_{processo_obj['key']}_{campo_roteiro['key']}_{st.session_state.form_key_counter}"
            tipo_roteiro = campo_roteiro["tipo"]
            obrigatorio_roteiro = campo_roteiro.get("obrigatorio", False)
            opcoes_roteiro = campo_roteiro.get("opcoes", [])
            default_value = respostas_roteiro.get(label_roteiro) 
            campo_key = campo_roteiro.get("key")

            condicional_info = campo_roteiro.get("condicional")
            if condicional_info:
                # A lógica condicional precisa verificar os campos que estão DENTRO do form
                # ou os valores já capturados no session_state para os campos de fora.
                valor_cond_dependencia = None
                campo_cond_obj_cfg = next((c for c in processo_obj["campos"] if c["key"] == condicional_info["campo"]), None)
                if campo_cond_obj_cfg:
                    if campo_cond_obj_cfg.get("key") in ["solucao_data_preparo", "solucao_tipo"]:
                        if campo_cond_obj_cfg.get("key") == "solucao_data_preparo":
                             valor_cond_dependencia = str(st.session_state.solucao_data_preparo_dinamico)
                        elif campo_cond_obj_cfg.get("key") == "solucao_tipo":
                            valor_cond_dependencia = st.session_state.solucao_tipo_dinamico
                    else:
                        valor_cond_dependencia = respostas_roteiro.get(campo_cond_obj_cfg["label"])
                
                if not campo_cond_obj_cfg:
                    st.warning(f"Campo condicional '{condicional_info['campo']}' não encontrado na configuração do processo.")
                    continue
                if valor_cond_dependencia != condicional_info["valor"]:
                    continue
            
            st.markdown(f"**{label_roteiro}**" + ("*" if obrigatorio_roteiro else ""))

            if tipo_roteiro == "texto":
                respostas_roteiro[label_roteiro] = st.text_input("", value=default_value or "", key=widget_key, label_visibility="collapsed")
            elif tipo_roteiro == "area_texto":
                respostas_roteiro[label_roteiro] = st.text_area("", value=default_value or "", key=widget_key, label_visibility="collapsed")
            elif tipo_roteiro == "numero":
                respostas_roteiro[label_roteiro] = st.number_input("", value=float(default_value or 0.0), key=widget_key, label_visibility="collapsed", step=campo_roteiro.get("step", 1.0))
            elif tipo_roteiro == "data":
                date_val = datetime.now().date()
                if default_value: 
                    try: date_val = datetime.strptime(str(default_value), "%Y-%m-%d").date()
                    except: pass
                
                if is_solucoes_process and campo_key == "solucao_data_validade":
                    validade_calculada_str = "(Aguardando Data de Preparo e Tipo)"
                    if st.session_state.solucao_data_preparo_dinamico and st.session_state.solucao_tipo_dinamico:
                        data_preparo_str_calc = str(st.session_state.solucao_data_preparo_dinamico)
                        tipo_sol_calc = st.session_state.solucao_tipo_dinamico
                        validade_obj = calculate_due_date(data_preparo_str_calc, tipo_sol_calc, roteiros_cfg_form)
                        if validade_obj:
                            validade_calculada_str = validade_obj.strftime("%Y-%m-%d")
                        else:
                            validade_calculada_str = "N/A (Erro no cálculo)"
                    
                    st.text_input("", value=validade_calculada_str, key=widget_key, disabled=True, label_visibility="collapsed")
                    respostas_roteiro[label_roteiro] = validade_calculada_str
                else:
                    user_date_input = st.date_input("", value=date_val, key=widget_key, label_visibility="collapsed")
                    respostas_roteiro[label_roteiro] = str(user_date_input)

            elif tipo_roteiro == "selecao":
                idx = opcoes_roteiro.index(default_value) if default_value and default_value in opcoes_roteiro else 0
                user_selection = st.selectbox("", opcoes_roteiro, index=idx, key=widget_key, label_visibility="collapsed")
                respostas_roteiro[label_roteiro] = user_selection

            elif tipo_roteiro == "multi_selecao":
                current_selection = [opt for opt in default_value if opt in opcoes_roteiro] if isinstance(default_value, list) else []
                respostas_roteiro[label_roteiro] = st.multiselect("", opcoes_roteiro, default=current_selection, key=widget_key, label_visibility="collapsed")
            elif tipo_roteiro == "checkbox":
                selecionados = default_value if isinstance(default_value, list) else []
                novos_selecionados = []
                for opcao in opcoes_roteiro:
                    if st.checkbox(opcao, value=(opcao in selecionados), key=f"{widget_key}_{opcao.replace(' ', '_').lower()}"):
                        novos_selecionados.append(opcao)
                respostas_roteiro[label_roteiro] = novos_selecionados
            elif tipo_roteiro == "radio":
                idx_radio = opcoes_roteiro.index(default_value) if default_value and default_value in opcoes_roteiro else 0
                respostas_roteiro[label_roteiro] = st.radio("", opcoes_roteiro, index=idx_radio, key=widget_key, label_visibility="collapsed", horizontal=True)
            elif tipo_roteiro == "upload_multiplos":
                uploaded_files = st.file_uploader("", accept_multiple_files=True, key=widget_key, label_visibility="collapsed", type=["png", "jpg", "jpeg", "pdf", "txt", "csv", "xlsx", "docx"])
                if uploaded_files:
                    st.session_state.current_form_uploads[label_roteiro] = uploaded_files
                elif label_roteiro in st.session_state.current_form_uploads and not uploaded_files:
                    # Não remover se não houver novos uploads, para manter o estado até a submissão
                    pass 
                if label_roteiro in st.session_state.current_form_uploads:
                    st.caption(f"Ficheiros: {", ".join([f.name for f in st.session_state.current_form_uploads[label_roteiro]])}")
        
        submitted_roteiro = st.form_submit_button("Finalizar e Submeter Inspeção")
        if submitted_roteiro:
            faltando_roteiro = []
            # Verificar obrigatoriedade para campos DENTRO do form
            for cr_def in campos_solucao_dentro_form:
                if not cr_def.get("obrigatorio"): continue
                label_obr = cr_def["label"]
                valor_resp = respostas_roteiro.get(label_obr)
                cond_info_obr = cr_def.get("condicional")
                if cond_info_obr:
                    # Lógica condicional para obrigatoriedade
                    valor_cond_dep_obr = None
                    campo_cond_obj_cfg_obr = next((c for c in processo_obj["campos"] if c["key"] == cond_info_obr["campo"]), None)
                    if campo_cond_obj_cfg_obr:
                        if campo_cond_obj_cfg_obr.get("key") in ["solucao_data_preparo", "solucao_tipo"]:
                            if campo_cond_obj_cfg_obr.get("key") == "solucao_data_preparo": valor_cond_dep_obr = str(st.session_state.solucao_data_preparo_dinamico)
                            elif campo_cond_obj_cfg_obr.get("key") == "solucao_tipo": valor_cond_dep_obr = st.session_state.solucao_tipo_dinamico
                        else:
                            valor_cond_dep_obr = respostas_roteiro.get(campo_cond_obj_cfg_obr["label"])
                    if not campo_cond_obj_cfg_obr or valor_cond_dep_obr != cond_info_obr["valor"]:
                        continue
                
                is_empty_str = isinstance(valor_resp, str) and not valor_resp.strip()
                is_empty_list = isinstance(valor_resp, list) and not valor_resp
                is_upload_missing = cr_def["tipo"] == "upload_multiplos" and (label_obr not in st.session_state.current_form_uploads or not st.session_state.current_form_uploads[label_obr])
                
                if valor_resp is None or is_empty_str or is_empty_list or is_upload_missing:
                    faltando_roteiro.append(label_obr)
            
            # Verificar obrigatoriedade para campos FORA do form (se aplicável ao processo de soluções)
            if is_solucoes_process:
                for cr_ext_def in campos_solucao_fora_form:
                    if not cr_ext_def.get("obrigatorio"): continue
                    label_obr_ext = cr_ext_def["label"]
                    valor_resp_ext = None
                    if cr_ext_def.get("key") == "solucao_data_preparo": valor_resp_ext = st.session_state.solucao_data_preparo_dinamico
                    elif cr_ext_def.get("key") == "solucao_tipo": valor_resp_ext = st.session_state.solucao_tipo_dinamico
                    
                    if valor_resp_ext is None or (isinstance(valor_resp_ext, str) and not valor_resp_ext.strip()):
                        faltando_roteiro.append(label_obr_ext)

            if faltando_roteiro:
                st.error(f"Preencha os campos obrigatórios: {", ".join(faltando_roteiro)}")
            else:
                st.success("Inspeção preenchida! A processar e salvar...")
                dados_completos_inspecao = {**st.session_state.dados_formulario_atual, "data_hora_submissao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                dados_completos_inspecao["evidencias_info"] = {lbl: [f.name for f in files] for lbl, files in st.session_state.current_form_uploads.items()}
                
                ctx_to_save = get_sharepoint_context()
                if ctx_to_save:
                    salvo_sp = salvar_inspecao_sharepoint(dados_completos_inspecao, ctx_to_save)
                    if not salvo_sp:
                        st.warning("A inspeção foi registada localmente, mas não pôde ser salva no SharePoint.")
                else:
                    st.warning("Contexto do SharePoint não disponível. A inspeção será apenas registada localmente.")

                st.session_state.inspecoes_realizadas_sessao.append(dados_completos_inspecao)
                st.info(f"Inspeção adicionada à sessão. Total na sessão: {len(st.session_state.inspecoes_realizadas_sessao)}.")
                
                st.session_state.form_key_counter += 1
                st.session_state.dados_formulario_atual = {}
                st.session_state.current_form_uploads = {}
                st.session_state.solucao_data_preparo_dinamico = datetime.now().date()
                st.session_state.solucao_tipo_dinamico = None
                st.rerun()

# --- Interface Principal da Aplicação ---
st.title("Formulário de Inspeção Dinâmico V2")

if not roteiros_config:
    st.error("Aplicação indisponível: Falha ao carregar configuração dos roteiros.")
else:
    setor_obj, processo_obj = render_sidebar_form(roteiros_config)

    if st.session_state.dados_formulario_atual.get("sidebar_completa") and processo_obj:
        st.header(f"Roteiro: {st.session_state.dados_formulario_atual.get("processo_selecionado_nome", "")} ({st.session_state.dados_formulario_atual.get("setor_selecionado_nome", "")})")
        render_dynamic_form(processo_obj, roteiros_config, None) 
    elif st.session_state.dados_formulario_atual.get("sidebar_completa") and not processo_obj:
        st.warning("Processo selecionado não encontrado. Verifique a configuração ou selecione um processo válido.")

    if st.session_state.inspecoes_realizadas_sessao:
        st.sidebar.subheader("Exportar Inspeções da Sessão")
        excel_bytes_sessao = gerar_excel_bytes(st.session_state.inspecoes_realizadas_sessao)
        if excel_bytes_sessao:
            st.sidebar.download_button(
                label="Download Excel da Sessão",
                data=excel_bytes_sessao,
                file_name=f"inspecoes_sessao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.sidebar.caption(f"{len(st.session_state.inspecoes_realizadas_sessao)} inspeção(ões) registada(s) nesta sessão.")

