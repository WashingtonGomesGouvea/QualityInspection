import streamlit as st
import json
import pandas as pd
from datetime import datetime, timedelta, date
import uuid
from PIL import Image
import hashlib
import io
import base64
from typing import Dict, List, Optional
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import os

# Configuração da página
st.set_page_config(
    page_title="Sistema de Inspeção Laboratorial - Synvia",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configuração do SharePoint
SHAREPOINT_BASE_PATH = "/personal/washington_gouvea_synvia_com/Documents/Inspe%C3%A7%C3%A3o%20Qualidade/Evidencias_inspecoes"
SHAREPOINT_DADOS_PATH = f"{SHAREPOINT_BASE_PATH}/dados"
SHAREPOINT_IMAGENS_PATH = f"{SHAREPOINT_DADOS_PATH}/imagens"
SHAREPOINT_INSPECOES_PATH = f"{SHAREPOINT_DADOS_PATH}/inspecoes"
SHAREPOINT_RELATORIOS_PATH = f"{SHAREPOINT_DADOS_PATH}/relatorios"

def get_sharepoint_context(max_retries=3):
    site_url = st.secrets["sharepoint"]["site_url"]
    username = st.secrets["sharepoint"]["email"]
    password = st.secrets["sharepoint"]["password"]
    
    for attempt in range(max_retries):
        try:
            ctx_auth = AuthenticationContext(site_url)
            if ctx_auth.acquire_token_for_user(username, password):
                ctx = ClientContext(site_url, ctx_auth)
                ctx.execute_query()  # Testa a conexão
                return ctx
            else:
                st.error("Falha na autenticação: Credenciais inválidas.")
                return None
        except Exception as e:
            st.warning(f"Tentativa {attempt + 1} falhou: {str(e)}")
            if attempt == max_retries - 1:
                st.error(f"Erro ao conectar ao SharePoint após {max_retries} tentativas: {e}")
                return None
    return None

# Classe GerenciadorInspetores
class GerenciadorInspetores:
    def __init__(self, sharepoint_path=SHAREPOINT_DADOS_PATH):
        self.sharepoint_path = sharepoint_path
        self.arquivo_inspetores = f"{sharepoint_path}/inspetores.json"
        self.inspetores = {}
        self.inspetores_iniciais = {
            "Aline Cristina Felício": "aline.felicio@synvia.com",
            "Amanda Hayashi Yamanouchi Brandão": "amanda.brandao@synvia.com",
            "Anderson da Silva Alves": "anderson.alves@synvia.com",
            "Bruna Pereira Nascimento Soares": "bruna.pereira@synvia.com",
            "Caio Henrique dos Santos Alves": "caio.alves@synvia.com",
            "Cristiane Sayuri Aoki Heredia": "cristiane.heredia@synvia.com",
            "Danilo Maria de Jesus": "danilo.jesus@synvia.com",
            "Dayane Salustriano de Araújo Torette": "dayane.araujo@synvia.com",
            "Eduarda Borges Barbosa": "eduarda.barbosa@synvia.com",
            "Gabriele Pavanello da Conceição Almeida": "gabriele.pavanello@synvia.com",
            "Jéssica M. M. Beatto Correia": "jessica.beatto@synvia.com",
            "Juan de Souza Mira": "juan.mira@synvia.com",
            "Júlia Gergollete Chaparro": "julia.chaparro@synvia.com",
            "Júlio César de Oliveira Santana": "julio.oliveira@synvia.com",
            "Lauren Soares Sautirio": "lauren.sautirio@synvia.com",
            "Letícia Cândida Ribeiro": "leticia.ribeiro@synvia.com",
            "Leticia De Souza Pereira": "leticia.souza@synvia.com",
            "Luana Sayuri Aoki Amâncio": "luana.amancio@synvia.com",
            "Mariele Innocenti": "mariele.innocenti@synvia.com",
            "Naira Ferro Cintra": "naira.ferro@synvia.com",
            "Paulo Rogerio Delmonde": "paulo.delmonde@synvia.com"
        }
        # Cache para evitar múltiplas leituras do SharePoint
        if 'inspetores_cache' not in st.session_state:
            self.carregar_inspetores()
            st.session_state.inspetores_cache = self.inspetores
        else:
            self.inspetores = st.session_state.inspetores_cache

    def carregar_inspetores(self) -> None:
        ctx = get_sharepoint_context()
        if not ctx:
            self.inspetores = self.inspetores_iniciais
            return
        
        try:
            file_content = download_file_content(ctx, self.arquivo_inspetores)
            self.inspetores = json.loads(file_content.decode('utf-8'))
        except Exception:
            self.inspetores = self.inspetores_iniciais
            # Não salva automaticamente para evitar escritas desnecessárias

    def salvar_inspetores(self) -> None:
        ctx = get_sharepoint_context()
        if not ctx:
            return
        
        try:
            ctx.web.folders.add(self.sharepoint_path).execute_query()
            file_content = json.dumps(self.inspetores, ensure_ascii=False, indent=4).encode('utf-8')
            target_folder = ctx.web.get_folder_by_server_relative_url(self.sharepoint_path)
            target_folder.upload_file("inspetores.json", file_content).execute_query()
            # Atualiza o cache após salvar
            st.session_state.inspetores_cache = self.inspetores
        except Exception as e:
            st.error(f"Erro ao salvar inspetores no SharePoint: {e}")

    def adicionar_inspetor(self, nome: str, email: str) -> None:
        self.inspetores[nome] = email
        self.salvar_inspetores()

    def obter_email_por_nome(self, nome: str) -> Optional[str]:
        return self.inspetores.get(nome)

    def obter_lista_inspetores(self) -> List[str]:
        return list(self.inspetores.keys())

# Singleton para GerenciadorInspetores
_instancia = None

def obter_instancia(sharepoint_path=SHAREPOINT_DADOS_PATH):
    global _instancia
    if _instancia is None:
        _instancia = GerenciadorInspetores(sharepoint_path)
    return _instancia

# Funções de Validade
def calcular_validade_solucao(data_preparo, tipo_solucao):
    if not data_preparo:
        return None
    if isinstance(data_preparo, str):
        try:
            data_preparo = datetime.strptime(data_preparo, "%Y-%m-%d").date()
        except ValueError:
            return None
    mapeamento_validade = {
        "Água Milli-Q": 1,
        "Água Milli-Q + Ácido/Base": 7,
        "Solução Alcalina / Ácido Diluído": 7,
        "Solução Tampão / Solução Salina": 7,
        "Solvente Orgânico + Ácido/Base": 7,
        "Solvente Orgânico + Solução Tampão": 7,
        "Solvente Orgânico + Água Milli-Q": 30,
        "Solvente Orgânico + Solvente Orgânico": 30,
        "Solvente Orgânico": "Prazo do fabricante",
        "Soluções Ácidas": 30,
        "Soluções Básicas": 90,
        "Soluções Tampão não utilizadas em análises cromatográficas": 15,
        "Soluções Aquosas (incluindo tampões)": 7,
        "Soluções Aquosas/Solventes Orgânicos (fase móvel, diluentes)": 30
    }
    if tipo_solucao == "Solvente Orgânico":
        return "Prazo do fabricante"
    dias_validade = mapeamento_validade.get(tipo_solucao, 7)
    if isinstance(dias_validade, int):
        return data_preparo + timedelta(days=dias_validade)
    return dias_validade

def formatar_data_validade(data_validade):
    if isinstance(data_validade, date):
        return data_validade.strftime("%d/%m/%Y")
    return str(data_validade)

def obter_dias_restantes(data_validade):
    if not isinstance(data_validade, date):
        return None
    hoje = date.today()
    return (data_validade - hoje).days

# Funções de Imagem
def salvar_imagem(imagem, prefixo="evidencia", sharepoint_path=SHAREPOINT_IMAGENS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        st.error("Não foi possível conectar ao SharePoint para salvar a imagem.")
        return None
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"{prefixo}_{timestamp}_{uuid.uuid4().hex[:8]}.jpg"
    caminho_arquivo = f"{sharepoint_path}/{nome_arquivo}"
    
    try:
        # Cria a pasta se não existir
        ctx.web.folders.add(sharepoint_path).execute_query()
        
        # Prepara o buffer da imagem
        if isinstance(imagem, bytes):
            buffer = io.BytesIO(imagem)
        elif isinstance(imagem, Image.Image):
            buffer = io.BytesIO()
            imagem.save(buffer, format="JPEG")
        else:
            img = Image.open(io.BytesIO(imagem))
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG")
        buffer.seek(0)
        
        # Faz o upload
        target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_path)
        target_folder.upload_file(nome_arquivo, buffer.getvalue()).execute_query()
        return caminho_arquivo
    except Exception as e:
        st.error(f"Erro ao salvar imagem no SharePoint: {e}")
        return None

def componente_imagem(chave, label="Adicionar evidência visual", sharepoint_path=SHAREPOINT_IMAGENS_PATH):
    col1, col2 = st.columns(2)
    caminho_key = f"imagem_path_{chave}"
    hash_key = f"imagem_hash_{chave}"
    with col1:
        arquivo_upload = st.file_uploader(
            f"{label} (Upload)", type=["jpg", "jpeg", "png"], key=f"upload_{chave}"
        )
    with col2:
        usar_camera = st.checkbox("Usar câmera", key=f"camera_check_{chave}")
        if usar_camera:
            imagem_camera = st.camera_input("Capturar imagem", key=f"camera_{chave}")
            if imagem_camera:
                conteudo = imagem_camera.getvalue()
                hash_imagem = hashlib.md5(conteudo).hexdigest()
                if st.session_state.get(hash_key) != hash_imagem:
                    st.session_state[hash_key] = hash_imagem
                    st.session_state[caminho_key] = salvar_imagem(
                        conteudo, f"evidencia_{chave}", sharepoint_path
                    )
    if arquivo_upload:
        conteudo = arquivo_upload.getvalue()
        hash_imagem = hashlib.md5(conteudo).hexdigest()
        if st.session_state.get(hash_key) != hash_imagem:
            st.session_state[hash_key] = hash_imagem
            st.session_state[caminho_key] = salvar_imagem(
                conteudo, f"evidencia_{chave}", sharepoint_path
            )
    return st.session_state.get(caminho_key)

def imagem_para_base64(caminho_imagem):
    ctx = get_sharepoint_context()
    if not ctx or not caminho_imagem:
        st.error("Conexão com SharePoint falhou ou caminho da imagem inválido.")
        return None
    
    try:
        ctx.web.get_file_by_server_relative_url(caminho_imagem).properties.execute_query()
        file_content = download_file_content(ctx, caminho_imagem) # Lê o conteúdo da imagem (retorna bytes)
        return base64.b64encode(file_content).decode('utf-8')
    except Exception as e:
        st.error(f"Erro ao carregar imagem do SharePoint: {e}")
        return None

# Funções de Exportação
def exportar_para_csv(dados, nome_arquivo, sharepoint_base=SHAREPOINT_DADOS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        return None
    
    try:
        dados_planos = {}
        def achatar_dict(d, prefixo=""):
            for k, v in d.items():
                if isinstance(v, dict):
                    achatar_dict(v, f"{prefixo}{k}_")
                elif isinstance(v, list):
                    dados_planos[f"{prefixo}{k}"] = ", ".join(map(str, v))
                else:
                    dados_planos[f"{prefixo}{k}"] = v
        achatar_dict(dados)
        df = pd.DataFrame([dados_planos])
        buffer = io.StringIO()
        df.to_csv(buffer, index=False, encoding='utf-8-sig')
        buffer.seek(0)
        file_path = f"{sharepoint_base}/relatorios/{nome_arquivo}"
        ctx.web.folders.add(f"{sharepoint_base}/relatorios").execute_query()
        target_folder = ctx.web.get_folder_by_server_relative_url(f"{sharepoint_base}/relatorios")
        target_folder.upload_file(nome_arquivo, buffer.getvalue().encode('utf-8')).execute_query()
        return file_path
    except Exception as e:
        st.error(f"Erro ao exportar para CSV no SharePoint: {e}")
        return None

def exportar_para_excel(dados, nome_arquivo, sharepoint_base=SHAREPOINT_DADOS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        return None
    
    try:
        dados_planos = {}
        def achatar_dict(d, prefixo=""):
            for k, v in d.items():
                if isinstance(v, dict):
                    achatar_dict(v, f"{prefixo}{k}_")
                elif isinstance(v, list):
                    dados_planos[f"{prefixo}{k}"] = ", ".join(map(str, v))
                else:
                    dados_planos[f"{prefixo}{k}"] = v
        achatar_dict(dados)
        df = pd.DataFrame([dados_planos])
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        file_path = f"{sharepoint_base}/relatorios/{nome_arquivo}"
        ctx.web.folders.add(f"{sharepoint_base}/relatorios").execute_query()
        target_folder = ctx.web.get_folder_by_server_relative_url(f"{sharepoint_base}/relatorios")
        target_folder.upload_file(nome_arquivo, buffer.getvalue()).execute_query()
        return file_path
    except Exception as e:
        st.error(f"Erro ao exportar para Excel no SharePoint: {e}")
        return None

# Funções de Processamento de Dados
def processar_dados_para_exportacao(dados):
    dados_processados = {}
    info_basicas = dados.get('informacoes_basicas', {})
    dados_processados['ID_Inspecao'] = dados.get('id_inspecao', '')
    dados_processados['Data_Inspecao'] = info_basicas.get('data_inspecao', '')
    dados_processados['Inspetor'] = info_basicas.get('nome_inspetor', '')
    dados_processados['Email_Inspetor'] = info_basicas.get('email_inspetor', '')
    dados_processados['Empresa'] = info_basicas.get('empresa', '')
    dados_processados['Setor'] = info_basicas.get('setor', '')
    dados_processados['Laboratorio'] = info_basicas.get('laboratorio', '')
    dados_processados['Processo'] = dados.get('processo_selecionado', '')
    dados_processados['Timestamp'] = dados.get('timestamp', '')
    
    dados_form = dados.get('dados_formulario', {})
    processo = dados.get('processo_selecionado', '')
    
    if processo == "Soluções":
        identificacao = dados_form.get('identificacao_controle', {})
        dados_processados['Codigo_Solucao'] = identificacao.get('codigo_solucao', '')
        dados_processados['Codigo_Padrao'] = identificacao.get('codigo_padrao', '')
        dados_processados['Etiqueta_Integra'] = identificacao.get('etiqueta_integra', '')
        dados_processados['Cadeia_Custodia'] = identificacao.get('cadeia_custodia', '')
        dados_processados['Substancia_Controlada'] = identificacao.get('substancia_controlada', '')
        dados_processados['Data_Recebimento_Padrao'] = identificacao.get('data_recebimento', '')
        dados_processados['Data_Preparo_Solucao'] = identificacao.get('data_preparo', '')
        dados_processados['Tipo_Solucao'] = identificacao.get('tipo_solucao', '')
        data_validade = identificacao.get('data_validade', '')
        if data_validade and isinstance(data_validade, str) and data_validade != "Prazo do fabricante":
            try:
                data_validade_formatada = datetime.fromisoformat(data_validade).strftime('%d/%m/%Y')
            except:
                data_validade_formatada = data_validade
        else:
            data_validade_formatada = data_validade
        dados_processados['Data_Validade_Solucao'] = data_validade_formatada
        anotacoes = dados_form.get('anotacoes_registro', {})
        dados_processados['Numero_Livro'] = anotacoes.get('numero_livro', '')
        dados_processados['Lacre'] = anotacoes.get('lacre', '')
        dados_processados['FOR'] = anotacoes.get('for', '')
        dados_processados['Classificacao_Correta'] = dados_form.get('classificacao_risco', '')
        dados_processados['Armazenamento_Adequado'] = dados_form.get('armazenamento_adequado', '')
        avaliacao = dados_form.get('avaliacao_conformidade', {})
        for erro, avaliacao_erro in avaliacao.items():
            dados_processados[f'Avaliacao_{erro.replace(" ", "_").replace("/", "_")}'] = avaliacao_erro
    
    elif processo == "Rastreabilidade de amostra":
        setor = info_basicas.get('setor', '')
        if setor == "Synvia Labs":
            identificacao = dados_form.get('identificacao_amostra', {})
            dados_processados['Etiqueta_Integra'] = identificacao.get('etiqueta_integra', '')
            dados_processados['Codigo_Amostra'] = identificacao.get('codigo_amostra', '')
            dados_processados['Data_Recebimento'] = identificacao.get('data_recebimento', '')
            dados_processados['Ativo'] = identificacao.get('ativo', '')
            dados_processados['Codigo_MBA'] = identificacao.get('codigo_mba', '')
            dados_processados['Armazenado_Corretamente'] = identificacao.get('armazenado_corretamente', '')
            racks = dados_form.get('identificacao_racks', {})
            dados_processados['Estudo'] = racks.get('estudo', '')
            dados_processados['Ensaio'] = racks.get('ensaio', '')
            dados_processados['Validade'] = racks.get('validade', '')
            dados_processados['Armazenamento_Adequado'] = racks.get('armazenamento_adequado', '')
        else:
            acompanhamento = dados_form.get('acompanhamento_amostra', {})
            dados_processados['Codigo_Amostra_Acompanhada'] = acompanhamento.get('codigo_amostra_acompanhada', '')
            dados_processados['Codigo_Lote_Acompanhado'] = acompanhamento.get('codigo_lote_acompanhado', '')
            dados_processados['Tipo_Amostra'] = ', '.join(acompanhamento.get('tipo_amostra', []))
            lcms = dados_form.get('lcms', {})
            dados_processados['TAG_LCMS'] = lcms.get('tag_lcms', '')
            dados_processados['Numero_Livro_LCMS'] = lcms.get('numero_livro_lcms', '')
            dados_processados['Data_Injecao'] = lcms.get('data_injecao', '')
            dados_processados['Horario_Injecao'] = lcms.get('horario_injecao', '')
            dados_processados['Criterios_Curva'] = lcms.get('criterios_curva', '')
            controles = dados_form.get('controles_rejeicoes', {})
            for controle, resultado in controles.items():
                dados_processados[f'Controle_{controle}'] = resultado
            extracao = dados_form.get('extracao', {})
            dados_processados['Numero_Livro_Extracao'] = extracao.get('numero_livro_extracao', '')
            dados_processados['Data_Inicio_Extracao'] = extracao.get('data_inicio_extracao', '')
            dados_processados['Horario_Entrada_Extracao'] = extracao.get('horario_entrada_extracao', '')
            dados_processados['Horario_Saida_Extracao'] = extracao.get('horario_saida_extracao', '')
            centrifuga = dados_form.get('centrifuga', {})
            dados_processados['TAG_Centrifuga'] = centrifuga.get('tag_centrifuga', '')
            dados_processados['Numero_Livro_Centrifuga'] = centrifuga.get('numero_livro_centrifuga', '')
            dados_processados['Horario_Entrada_Centrifuga'] = centrifuga.get('horario_entrada_centrifuga', '')
            dados_processados['Horario_Saida_Centrifuga'] = centrifuga.get('horario_saida_centrifuga', '')
            ultrassom = dados_form.get('ultrassom', {})
            dados_processados['Numero_Livro_Ultrassom'] = ultrassom.get('numero_livro_ultrassom', '')
            dados_processados['Data_Anotacao_Ultrassom'] = ultrassom.get('data_anotacao_ultrassom', '')
            dados_processados['Horario_Entrada_Ultrassom'] = ultrassom.get('horario_entrada_ultrassom', '')
            dados_processados['Horario_Saida_Ultrassom'] = ultrassom.get('horario_saida_ultrassom', '')
            transporte = dados_form.get('transporte', {})
            dados_processados['Numero_Pacote'] = transporte.get('numero_pacote', '')
            dados_processados['Data_Recebimento_Pacote'] = transporte.get('data_recebimento_pacote', '')
            dados_processados['Horario_Recebimento_Pacote'] = transporte.get('horario_recebimento_pacote', '')
            dados_processados['Transportadora'] = transporte.get('transportadora', '')
    
    elif processo == "Equipamentos":
        identificacao = dados_form.get('identificacao', {})
        dados_processados['TAG'] = identificacao.get('tag', '')
        dados_processados['Logbook'] = identificacao.get('logbook', '')
        dados_processados['Calibracao_Valida'] = identificacao.get('calibracao_valida', '')
        dados_processados['Numero_Certificado'] = identificacao.get('num_certificado', '')
        dados_processados['Proxima_Calibracao'] = identificacao.get('proxima_calibracao', '')
        dados_processados['Anotacao_Logbook'] = identificacao.get('anotacao_logbook', '')
        dados_processados['Anotacao_Outros'] = identificacao.get('anotacao_outros', '')
        dados_processados['Equipamento'] = dados_form.get('equipamento_selecionado', '')
        campos_especificos = dados_form.get('campos_especificos', {})
        for categoria, detalhes in campos_especificos.items():
            if isinstance(detalhes, dict):
                for item, valor in detalhes.items():
                    dados_processados[f'{categoria}_{item}'] = valor
    
    elif processo == "Monitoramento ambiental":
        info_logbook = dados_form.get('info_logbook', {})
        dados_processados['Numero_Logbook'] = info_logbook.get('numero_logbook', '')
        dados_processados['TAG_Equipamento'] = info_logbook.get('tag_equipamento', '')
        dados_processados['Data_Abertura'] = info_logbook.get('data_abertura', '')
        dados_processados['Localizacao'] = info_logbook.get('localizacao', '')
        dados_processados['Ocorrencias'] = ', '.join(dados_form.get('ocorrencias', []))
        dados_processados['Integridade_Dados'] = ', '.join(dados_form.get('integridade_dados', []))
        dados_processados['Condicoes_Logbook'] = ', '.join(dados_form.get('condicoes_logbook', []))
        equipamentos = dados_form.get('equipamentos_associados', {})
        dados_processados['TAG_Termo'] = equipamentos.get('tag_termo', '')
        dados_processados['Num_Logbook_Monit'] = equipamentos.get('num_logbook_monit', '')
        dados_processados['Num_Certificado'] = equipamentos.get('num_certificado', '')
        dados_processados['Data_Calibracao'] = equipamentos.get('data_calibracao', '')
        dados_processados['Registros_3Meses'] = ', '.join(dados_form.get('registros_3meses', []))
    
    else:
        if 'info_logbook' in dados_form:
            info_logbook = dados_form.get('info_logbook', {})
            dados_processados['Numero_Logbook'] = info_logbook.get('numero_logbook', '')
            dados_processados['TAG_Equipamento'] = info_logbook.get('tag_equipamento', '')
            dados_processados['Data_Abertura'] = info_logbook.get('data_abertura', '')
            dados_processados['Localizacao'] = info_logbook.get('localizacao', '')
        if 'integridade_dados' in dados_form:
            dados_processados['Integridade_Dados'] = ', '.join(dados_form.get('integridade_dados', []))
        if 'avaliacao_detalhada' in dados_form:
            avaliacao = dados_form.get('avaliacao_detalhada', {})
            for erro, avaliacao_erro in avaliacao.items():
                dados_processados[f'Avaliacao_{erro.replace(" ", "_").replace("/", "_")}'] = avaliacao_erro
        if 'condicoes_logbook' in dados_form:
            dados_processados['Condicoes_Logbook'] = ', '.join(dados_form.get('condicoes_logbook', []))
    
    dados_processados['Evidencia_Visual'] = dados_form.get('evidencia_visual', '')
    dados_processados['Observacoes'] = dados_form.get('observacoes', '')
    return dados_processados

def exportar_lista_completa_inspecoes(sharepoint_base=SHAREPOINT_DADOS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        return None
    
    arquivo_inspecoes = f"{sharepoint_base}/inspecoes/inspecoes.json"
    
    try:
        file_content = download_file_content(ctx, arquivo_inspecoes)  # Lê o conteúdo do arquivo (retorna bytes)
        inspecoes = json.loads(file_content.decode('utf-8')) if file_content else []
        
        dados_relatorio = []
        for insp in inspecoes:
            dados_processados = processar_dados_para_exportacao(insp)
            dados_relatorio.append(dados_processados)
        
        if dados_relatorio:
            df = pd.DataFrame(dados_relatorio)
            nome_arquivo = f"relatorio_completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            caminho_arquivo = f"{sharepoint_base}/relatorios/{nome_arquivo}"
            
            buffer = io.StringIO()
            df.to_csv(buffer, index=False, encoding='utf-8-sig')
            buffer.seek(0)
            ctx.web.folders.add(f"{sharepoint_base}/relatorios").execute_query()
            target_folder = ctx.web.get_folder_by_server_relative_url(f"{sharepoint_base}/relatorios")
            target_folder.upload_file(nome_arquivo, buffer.getvalue().encode('utf-8')).execute_query()
            
            nome_arquivo_excel = nome_arquivo.replace('.csv', '.xlsx')
            caminho_excel = f"{sharepoint_base}/relatorios/{nome_arquivo_excel}"
            buffer_excel = io.BytesIO()
            df.to_excel(buffer_excel, index=False, engine='openpyxl')
            buffer_excel.seek(0)
            target_folder.upload_file(nome_arquivo_excel, buffer_excel.getvalue()).execute_query()
            
            return caminho_arquivo
        return None
    except Exception as e:
        st.error(f"Erro ao exportar lista completa de inspeções: {e}")
        return None

# Funções de Inspeção
def salvar_inspecao(dados, sharepoint_base=SHAREPOINT_DADOS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        st.error("Não foi possível conectar ao SharePoint para salvar a inspeção.")
        return None
    
    id_inspecao = f"insp_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
    dados['id_inspecao'] = id_inspecao
    dados['timestamp'] = datetime.now().isoformat()
    arquivo_inspecoes = f"{sharepoint_base}/inspecoes/inspecoes.json"
    
    try:
        # Cria a pasta se não existir
        ctx.web.folders.add(f"{sharepoint_base}/inspecoes").execute_query()
        
        # Tenta ler o arquivo existente
        try:
            file_content = download_file_content(ctx, arquivo_inspecoes)
            inspecoes = json.loads(file_content.decode('utf-8')) if file_content else []
        except Exception:
            inspecoes = []  # Inicializa com lista vazia se o arquivo não existe
        
        inspecoes.append(dados)
        file_content = json.dumps(inspecoes, ensure_ascii=False, indent=4).encode('utf-8')
        target_folder = ctx.web.get_folder_by_server_relative_url(f"{sharepoint_base}/inspecoes")
        target_folder.upload_file("inspecoes.json", file_content).execute_query()
        
        # Exporta para CSV e Excel
        nome_arquivo_csv = f"{id_inspecao}.csv"
        nome_arquivo_excel = f"{id_inspecao}.xlsx"
        exportar_para_csv(dados, nome_arquivo_csv, sharepoint_base)
        exportar_para_excel(dados, nome_arquivo_excel, sharepoint_base)
        
        # Atualiza o cache
        st.session_state.inspecoes_cache = inspecoes
        
        return id_inspecao
    except Exception as e:
        st.error(f"Erro ao salvar inspeção no SharePoint: {e}")
        return None

def gerar_relatorio(id_inspecao, sharepoint_base=SHAREPOINT_DADOS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        return None
    
    arquivo_inspecoes = f"{sharepoint_base}/inspecoes/inspecoes.json"
    
    try:
        file_content = download_file_content(ctx, arquivo_inspecoes)  # Lê o conteúdo do arquivo (retorna bytes)
        inspecoes = json.loads(file_content.decode('utf-8')) if file_content else []
        
        inspecao = next((i for i in inspecoes if i.get('id_inspecao') == id_inspecao), None)
        if not inspecao:
            st.error(f"Inspeção com ID {id_inspecao} não encontrada.")
            return None
        
        nome_arquivo = f"relatorio_{id_inspecao}.csv"
        caminho_arquivo = exportar_para_csv(inspecao, nome_arquivo, sharepoint_base)
        nome_arquivo_excel = f"relatorio_{id_inspecao}.xlsx"
        exportar_para_excel(inspecao, nome_arquivo_excel, sharepoint_base)
        
        return caminho_arquivo
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        return None

def listar_inspecoes(sharepoint_base=SHAREPOINT_DADOS_PATH):
    ctx = get_sharepoint_context()
    if not ctx:
        st.error("Não foi possível conectar ao SharePoint para listar inspeções.")
        return []
    
    arquivo_inspecoes = f"{sharepoint_base}/inspecoes/inspecoes.json"
    
    try:
        ctx.web.get_file_by_server_relative_url(arquivo_inspecoes).properties.execute_query()
        file_content = download_file_content(ctx, arquivo_inspecoes)  # Lê o conteúdo do arquivo (retorna bytes)
        
        inspecoes = json.loads(file_content.decode('utf-8')) if file_content else []
        
        st.session_state.inspecoes_cache = inspecoes
        
        return [
            {
                'id_inspecao': insp.get('id_inspecao', ''),
                'data_inspecao': insp.get('informacoes_basicas', {}).get('data_inspecao', ''),
                'nome_inspetor': insp.get('informacoes_basicas', {}).get('nome_inspetor', ''),
                'empresa': insp.get('informacoes_basicas', {}).get('empresa', ''),
                'setor': insp.get('informacoes_basicas', {}).get('setor', ''),
                'processo': insp.get('processo_selecionado', '')
            }
            for insp in inspecoes
        ]
    except Exception as e:
        st.error(f"Erro ao listar inspeções: {e}")
        return []

# Componentes de Interface
def tabela_avaliacao_erros(chave, erros=None):
    if erros is None:
        erros = [
            "Falta de assinatura/rubrica",
            "Falta de preenchimento de dados em dia/horário",
            "Rasura",
            "Cancelamento incorreto",
            "Correção incorreta",
            "Não está na sequência cronológica esperada",
            "Dados fora da especificação, sem tomada de ações",
            "TAG incorreta",
            "Dados ilegíveis"
        ]
    opcoes = ["0 erros", "1 a 5 erros", "6 a 10 erros", "Mais de 10 erros"]
    st.write("### Avaliação Detalhada")
    resultados = {}
    cols = st.columns([3] + [1] * len(opcoes))
    with cols[0]:
        st.write("**Erro**")
    for i, opcao in enumerate(opcoes):
        with cols[i+1]:
            st.write(f"**{opcao}**")
    for erro in erros:
        cols = st.columns([3] + [1] * len(opcoes))
        with cols[0]:
            st.write(erro)
        chave_erro = f"{chave}_{erro.replace(' ', '_').replace('/', '_')}"
        selecao = st.radio(
            f"Selecione para {erro}",
            opcoes,
            key=chave_erro,
            label_visibility="collapsed",
            horizontal=True
        )
        resultados[erro] = selecao
    return resultados

def componente_condicoes_logbook(chave):
    st.write("### Condições do Logbook")
    opcoes = [
        "Páginas rasgadas pelo lacre",
        "Lombada rasgada",
        "Página molhada com dados borrados",
        "Páginas rasgadas nas anotações",
        "Logbook íntegro",
        "Lacre íntegro (legível e inteiro)"
    ]
    selecao = st.multiselect(
        "Selecione as condições observadas:",
        opcoes,
        key=f"condicoes_logbook_{chave}"
    )
    return selecao

def componente_integridade_dados(chave):
    st.write("### Integridade de Dados")
    opcoes = [
        "Falta de assinatura/rubrica",
        "Falta de preenchimento de dados em dia/horário",
        "Rasura",
        "Cancelamento incorreto",
        "Não está na sequência cronológica esperada",
        "O logbook não apresenta problemas de preenchimento"
    ]
    selecao = st.multiselect(
        "Selecione os problemas observados:",
        opcoes,
        key=f"integridade_dados_{chave}"
    )
    if st.checkbox("Outro problema não listado", key=f"outro_check_{chave}"):
        outro = st.text_area(
            "Descreva o problema:",
            key=f"outro_texto_{chave}"
        )
        if outro:
            selecao.append(f"Outro: {outro}")
    return selecao

def componente_info_logbook(chave, localizacoes=None):
    if localizacoes is None:
        localizacoes = [
            "Sala de extração",
            "Sala do massas",
            "Sala de preparo de amostras",
            "Sala de pesagem",
            "Sala de acondicionamento/freezer",
            "Sala de preparo",
            "Sala de oncológicos",
            "Sala de CG",
            "Sala de HPLC",
            "Sala de paramentação (Micro)",
            "Sala de preparo e esterilização (Micro)",
            "Sala de análise A (Micro)",
            "Sala de análise B (Micro)",
            "Sala de incubação (Micro)",
            "Sala de leitura (Micro)",
            "Sala de descontaminação (Micro)"
        ]
    st.write("### Informações do Logbook")
    numero_logbook = st.text_input(
        "Qual o número do logbook que será inspecionado?*",
        key=f"numero_logbook_{chave}"
    )
    tag_equipamento = st.text_input(
        "TAG do equipamento, se aplicável (preencher com 'NA' se não houver)",
        key=f"tag_equipamento_{chave}"
    )
    data_abertura = st.date_input(
        "Data de abertura do logbook",
        key=f"data_abertura_{chave}"
    )
    todas_localizacoes = localizacoes + ["Outro"]
    localizacao = st.selectbox(
        "Qual a localização do logbook que será inspecionado?",
        todas_localizacoes,
        key=f"localizacao_{chave}"
    )
    if localizacao == "Outro":
        localizacao_outro = st.text_input(
            "Especifique a localização:",
            key=f"localizacao_outro_{chave}"
        )
        localizacao = localizacao_outro
    return {
        "numero_logbook": numero_logbook,
        "tag_equipamento": tag_equipamento,
        "data_abertura": data_abertura.isoformat() if data_abertura else None,
        "localizacao": localizacao
    }

# Formulários de Processo
def processo_monitoramento_ambiental():
    st.header("🧪 Monitoramento Ambiental")
    info_logbook = componente_info_logbook("monit_amb")
    st.write("### Descrição das Ocorrências")
    ocorrencias = st.multiselect(
        "Selecione as ocorrências observadas:",
        [
            "Falta de assinatura/rubrica",
            "Falta de preenchimento de dados em dia/horário",
            "Rasura",
            "Cancelamento incorreto",
            "Não está na sequência cronológica esperada",
            "O logbook não apresenta problemas de preenchimento"
        ],
        key="ocorrencias_monit_amb"
    )
    if st.checkbox("Outra ocorrência não listada", key="outro_check_ocorrencias"):
        outro = st.text_area(
            "Descreva a ocorrência:",
            key="outro_texto_ocorrencias"
        )
        if outro:
            ocorrencias.append(f"Outro: {outro}")
    integridade_dados = componente_integridade_dados("monit_amb")
    condicoes_logbook = componente_condicoes_logbook("monit_amb")
    st.write("### Evidências Visuais")
    caminho_evidencia = componente_imagem("monit_amb", "Adicionar foto das evidências, se aplicável")
    st.write("### Equipamentos Associados")
    tag_termo = st.text_input(
        "TAG do termo-higrômetro:",
        key="tag_termo_monit_amb"
    )
    num_logbook_monit = st.text_input(
        "Nº do logbook de monitoramento ambiental:",
        key="num_logbook_monit_amb"
    )
    num_certificado = st.text_input(
        "Nº do certificado:",
        key="num_certificado_monit_amb"
    )
    data_calibracao = st.date_input(
        "Data da última calibração:",
        key="data_calibracao_monit_amb"
    )
    st.write("### Avaliação dos Registros dos Últimos 3 Meses")
    registros_3meses = st.multiselect(
        "Selecione as situações observadas:",
        [
            "Registros feitos corretamente",
            "Temperatura fora de especificação, sem justificativa",
            "Temperatura fora de especificação, com justificativa",
            "Ausência de registro sem justificativa"
        ],
        key="registros_3meses_monit_amb"
    )
    st.write("### Observações Gerais")
    observacoes = st.text_area(
        "Observações pertinentes:",
        key="observacoes_monit_amb"
    )
    return {
        "processo": "Monitoramento Ambiental",
        "info_logbook": info_logbook,
        "ocorrencias": ocorrencias,
        "integridade_dados": integridade_dados,
        "condicoes_logbook": condicoes_logbook,
        "evidencia_visual": caminho_evidencia,
        "equipamentos_associados": {
            "tag_termo": tag_termo,
            "num_logbook_monit": num_logbook_monit,
            "num_certificado": num_certificado,
            "data_calibracao": data_calibracao.isoformat() if data_calibracao else None
        },
        "registros_3meses": registros_3meses,
        "observações": observações
    }

def processo_equipamentos():
    st.header("⚙️ Equipamentos")
    st.write("### Identificação do Equipamento")
    tag = st.text_input("TAG*", key="tag_equipamento")
    logbook = st.text_input(
        "Logbook do equipamento (incluir número do livro, FOR e lacre, se aplicável)",
        key="logbook_equipamento"
    )
    calibracao_valida = st.radio(
        "Calibração ou manutenção dentro da validade?",
        ["Sim", "Não", "Vencimento próximo"],
        key="calibracao_valida"
    )
    num_certificado = st.text_input("Número do certificado", key="num_certificado_equip")
    proxima_calibracao = st.date_input("Próxima calibração", key="proxima_calibracao")
    anotacao_logbook = st.radio(
        "Última verificação/calibração do equipamento foi anotada no logbook, estando de acordo com a etiqueta?",
        ["Sim", "Não", "Outros"],
        key="anotacao_logbook"
    )
    anotacao_outros = st.text_input("Especifique:", key="anotacao_outros") if anotacao_logbook == "Outros" else None
    st.write("### Selecionar Equipamento Inspecionado")
    equipamentos = [
        "Agitador de Microplacas", "Agitador de tubos", "Centrífuga", "Centrífuga de microplacas",
        "Centrífuga refrigerada", "Chuveiro lava olhos", "Ducha oftálmica", "Contador de células",
        "Cronômetro digital", "Desumidificador", "Espectrofotômetro para microplacas",
        "Estufa de secagem e esterilização", "Freezer", "Ultrafreezer", "Geladeira",
        "Geladeira duplex", "Homogeneizador de Microtubos", "Lavadora de Microplacas",
        "Leitora multicanal para microplacas", "Micropipetas", "Micropipeta eletrônica",
        "Micropipeta multicanal", "Microscópio binocular Acromático", "Osmose reversa",
        "Pipetador automático", "Relógio Digital", "Termobloco incubador", "Termociclador",
        "Termo-higrômetro", "Termômetro digital, Max. Min.", "Termômetro digital infravermelho",
        "Balança analítica", "Capela de exaustão", "Concentrador de amostras",
        "Homogeneizador", "Cromatógrafo Líquido Acoplado a Espectrômetro de Massas",
        "Espectrômetro de massa com plasma indutivamente acoplado"
    ]
    equipamento_selecionado = st.selectbox("Selecione o equipamento inspecionado:", equipamentos, key="equipamento_selecionado")
    campos_especificos = {}
    if equipamento_selecionado == "Balança analítica":
        st.write("### Balanças (Verificação diária)")
        campos_especificos["balanca"] = tabela_avaliacao_erros("balanca")
    elif equipamento_selecionado in ["Micropipetas", "Micropipeta eletrônica", "Micropipeta multicanal"]:
        st.write("### Micropipetas (PLT Unit)")
        erros_plt = ["Dias sem registro", "Resultado fora da especificação", "Campos sem preenchimento"]
        campos_especificos["micropipeta_plt"] = tabela_avaliacao_erros("micropipeta_plt", erros_plt)
        st.write("### Micropipetas (Verificação Gravimétrica)")
        erros_gravimetrica = ["Resultado fora da especificação", "Campos sem preenchimento", "Verificação fora da data especificada"]
        opcoes_gravimetrica = ["0 erros", "1 erro", "2 erros", "3 erros", "4 ou mais erros"]
        st.write("#### Avaliação Detalhada")
        resultados_gravimetrica = {}
        cols = st.columns([3] + [1] * len(opcoes_gravimetrica))
        with cols[0]:
            st.write("**Erro**")
        for i, opcao in enumerate(opcoes_gravimetrica):
            with cols[i+1]:
                st.write(f"**{opcao}**")
        for erro in erros_gravimetrica:
            cols = st.columns([3] + [1] * len(opcoes_gravimetrica))
            with cols[0]:
                st.write(erro)
            chave_erro = f"gravimetrica_{erro.replace(' ', '_').replace('/', '_')}"
            selecao = st.radio(
                f"Selecione para {erro}",
                opcoes_gravimetrica,
                key=chave_erro,
                label_visibility="collapsed",
                horizontal=True
            )
            resultados_gravimetrica[erro] = selecao
        campos_especificos["micropipeta_gravimetrica"] = resultados_gravimetrica
    st.write("### Observações Gerais")
    observacoes = st.text_area("Observações pertinentes:", key="observacoes_equip")
    st.write("### Evidências Visuais")
    caminho_evidencia = componente_imagem("equipamentos", "Inserir evidências, se aplicável")
    return {
        "processo": "Equipamentos",
        "identificacao": {
            "tag": tag,
            "logbook": logbook,
            "calibracao_valida": calibracao_valida,
            "num_certificado": num_certificado,
            "proxima_calibracao": proxima_calibracao.isoformat() if proxima_calibracao else None,
            "anotacao_logbook": anotacao_logbook,
            "anotacao_outros": anotacao_outros
        },
        "equipamento_selecionado": equipamento_selecionado,
        "campos_especificos": campos_especificos,
        "observacoes": observacoes,
        "evidencia_visual": caminho_evidencia
    }

def processo_solucoes():
    st.header("🧪 Soluções")
    st.write("### Identificação e Controle")
    codigo_solucao = st.text_input("Código da Solução*", key="codigo_solucao")
    codigo_padrao = st.text_input("Código do padrão utilizado*", key="codigo_padrao")
    etiqueta_integra = st.radio("Etiqueta de recebimento de reagente e identificação de solução estão íntegras?", ["Sim", "Não"], key="etiqueta_integra")
    cadeia_custodia = st.radio("Cadeia de custódia (FOR-401) preenchida?", ["Sim", "Não"], key="cadeia_custodia")
    substancia_controlada = st.radio("Substância controlada pela Portaria nº 344/98?", ["Sim", "Não"], key="substancia_controlada")
    data_recebimento = st.date_input("Data de recebimento do padrão", key="data_recebimento_solucao")
    data_preparo = st.date_input("Data de preparo da solução", key="data_preparo_solucao")
    st.write("### Tipo de Solução")
    tipos_solucao = [
        "Água Milli-Q", "Água Milli-Q + Ácido/Base", "Solução Alcalina / Ácido Diluído",
        "Solução Tampão / Solução Salina", "Solvente Orgânico + Ácido/Base",
        "Solvente Orgânico + Solução Tampão", "Solvente Orgânico + Água Milli-Q",
        "Solvente Orgânico + Solvente Orgânico", "Solvente Orgânico", "Soluções Ácidas",
        "Soluções Básicas", "Soluções Tampão não utilizadas em análises cromatográficas",
        "Soluções Aquosas (incluindo tampões)", "Soluções Aquosas/Solventes Orgânicos (fase móvel, diluentes)"
    ]
    tipo_solucao = st.selectbox("Selecione o tipo de solução:", tipos_solucao, key="tipo_solucao")
    if data_preparo:
        data_validade_calculada = calcular_validade_solucao(data_preparo, tipo_solucao)
        if isinstance(data_validade_calculada, str):
            st.info(f"Validade da solução: {data_validade_calculada}")
            data_validade = st.date_input("Data de validade da solução (conforme fabricante)", key="data_validade_solucao")
        else:
            st.info(f"Validade calculada: {data_validade_calculada.strftime('%d/%m/%Y')}")
            data_validade = data_validade_calculada
    else:
        data_validade = st.date_input("Data de validade da solução", key="data_validade_solucao")
    st.write("### Anotações e Registro")
    numero_livro = st.text_input("Número do livro", key="numero_livro_solucao")
    lacre = st.text_input("Lacre", key="lacre_solucao")
    for_selecionado = st.radio("FOR", ["FOR-297: Soluções gerais", "FOR-298: Soluções padrão"], key="for_solucao")
    st.write("### Classificação de Risco")
    classificacao_correta = st.radio("Classificação está correta?", ["Sim", "Não"], key="classificacao_correta")
    if classificacao_correta == "Sim":
        st.write("""
        | Cor | Significado |
        |-----|-------------|
        | Laranja | Reagente sem classificação de risco |
        | Amarelo | Reagente altamente reativo, explosivo, reativo com água, fortemente oxidante e pirofórico |
        | Azul | Potencial severo à saúde (inalação, ingestão, absorção) |
        | Vermelho | Inflamável ou combustível |
        | Branco | Corrosivo |
        """)
    st.write("### Armazenamento")
    armazenamento_adequado = st.radio("O armazenamento está adequado?", ["Sim", "Não"], key="armazenamento_adequado_solucao")
    st.write("### Avaliação da Conformidade")
    avaliacao_conformidade = tabela_avaliacao_erros("solucoes")
    st.write("### Evidências Visuais")
    caminho_evidencia = componente_imagem("solucoes", "Inserir evidências, se aplicável")
    st.write("### Observações Gerais")
    observacoes = st.text_area("Observações pertinentes:", key="observacoes_solucoes")
    return {
        "processo": "Soluções",
        "identificacao_controle": {
            "codigo_solucao": codigo_solucao,
            "codigo_padrao": codigo_padrao,
            "etiqueta_integra": etiqueta_integra,
            "cadeia_custodia": cadeia_custodia,
            "substancia_controlada": substancia_controlada,
            "data_recebimento": data_recebimento.isoformat() if data_recebimento else None,
            "data_preparo": data_preparo.isoformat() if data_preparo else None,
            "tipo_solucao": tipo_solucao,
            "data_validade": data_validade.isoformat() if isinstance(data_validade, date) else data_validade
        },
        "anotacoes_registro": {
            "numero_livro": numero_livro,
            "lacre": lacre,
            "for": for_selecionado
        },
        "classificacao_risco": classificacao_correta,
        "armazenamento_adequado": armazenamento_adequado,
        "avaliacao_conformidade": avaliacao_conformidade,
        "evidencia_visual": caminho_evidencia,
        "observacoes": observacoes
    }

def processo_rastreabilidade_amostra_labs():
    st.header("🧬 Rastreabilidade de Amostra (Synvia Labs)")
    st.write("### Identificação da Amostra")
    etiqueta_integra = st.radio("Etiqueta de identificação da amostra está íntegra?", ["Sim", "Não"], key="etiqueta_integra_labs")
    codigo_amostra = st.text_input("Código da Amostra*", key="codigo_amostra_labs")
    data_recebimento = st.date_input("Data de recebimento da amostra", key="data_recebimento_amostra_labs")
    ativo = st.text_input("Ativo (ECD / ECCD / ELD)", key="ativo_labs")
    codigo_mba = st.text_input("Código MBA (ECD / ECCD / ELD)", key="codigo_mba_labs")
    armazenado_corretamente = st.radio("Armazenado corretamente?", ["Sim", "Não"], key="armazenado_corretamente_labs")
    st.write("### Identificação das Rack's")
    estudo = st.text_input("Estudo", key="estudo_labs")
    ensaio = st.text_input("Ensaio", key="ensaio_labs")
    validade = st.date_input("Validade", key="validade_rack_labs")
    armazenamento_adequado = st.radio("Armazenamento está adequado?", ["Sim", "Não"], key="armazenamento_adequado_labs")
    st.write("### Evidências Visuais")
    caminho_evidencia = componente_imagem("rastreabilidade_labs", "Fotos das evidências, se aplicável")
    st.write("### Observações Gerais")
    observacoes = st.text_area("Observações pertinentes:", key="observacoes_rastreabilidade_labs")
    return {
        "processo": "Rastreabilidade de Amostra (Labs)",
        "identificacao_amostra": {
            "etiqueta_integra": etiqueta_integra,
            "codigo_amostra": codigo_amostra,
            "data_recebimento": data_recebimento.isoformat() if data_recebimento else None,
            "ativo": ativo,
            "codigo_mba": codigo_mba,
            "armazenado_corretamente": armazenado_corretamente
        },
        "identificacao_racks": {
            "estudo": estudo,
            "ensaio": ensaio,
            "validade": validade.isoformat() if validade else None,
            "armazenamento_adequado": armazenamento_adequado
        },
        "evidencia_visual": caminho_evidencia,
        "observacoes": observacoes
    }

def processo_rastreabilidade_amostra_tox():
    st.header("🧬 Rastreabilidade de Amostra (Synvia Tox)")
    st.write("### Acompanhamento da amostra")
    codigo_amostra_acompanhada = st.text_input("Código da amostra acompanhada*", key="codigo_amostra_acompanhada")
    codigo_lote_acompanhado = st.text_input("Código do lote acompanhado", key="codigo_lote_acompanhado")
    tipo_amostra = st.multiselect(
        "A amostra é de triagem ou confirmatório? (máximo 2 opções)",
        ["Triagem", "Confirmatório Geral", "Confirmatório THC"],
        key="tipo_amostra",
        max_selections=2
    )
    st.write("### LC-MS/MS")
    tag_lcms = st.text_input("TAG LC-MS/MS", key="tag_lcms")
    numero_livro_lcms = st.text_input("Número do livro LC-MS/MS", key="numero_livro_lcms")
    data_injecao = st.date_input("Data da injeção no LC-MS/MS", key="data_injecao")
    horario_injecao = st.time_input("Horário da injeção anotado no logbook", key="horario_injecao")
    criterios_curva = st.radio(
        "Os critérios da curva de calibração feitos no dia foram atendidos?",
        [
            "Curva injetada, critérios atendidos",
            "Curva injetada, critérios não atendidos",
            "Curva não injetada (sem registro)"
        ],
        key="criterios_curva"
    )
    st.write("### Controles e Rejeições")
    st.write("Os controles para o lote foram aprovados?")
    controles = ["CQA", "CQM", "CQB"]
    opcoes_rejeicao = ["Todos aprovados", "1 Rejeição", "2 Rejeições", "3 Rejeições", "Acima de 4 rejeições"]
    resultados_controles = {}
    cols = st.columns([2] + [1] * len(opcoes_rejeicao))
    with cols[0]:
        st.write("**Item**")
    for i, opcao in enumerate(opcoes_rejeicao):
        with cols[i+1]:
            st.write(f"**{opcao}**")
    for controle in controles:
        cols = st.columns([2] + [1] * len(opcoes_rejeicao))
        with cols[0]:
            st.write(controle)
        chave_controle = f"controle_{controle}"
        selecao = st.radio(
            f"Selecione para {controle}",
            opcoes_rejeicao,
            key=chave_controle,
            label_visibility="collapsed",
            horizontal=True
        )
        resultados_controles[controle] = selecao
    st.write("### Extração")
    numero_livro_extracao = st.text_input("Número do Livro Ata de Extração", key="numero_livro_extracao")
    data_inicio_extracao = st.date_input("Data de início da extração", key="data_inicio_extracao")
    horario_entrada_extracao = st.time_input("Horário de entrada na extração", key="horario_entrada_extracao")
    horario_saida_extracao = st.time_input("Horário de saída da extração", key="horario_saida_extracao")
    st.write("### Centrífuga 2")
    tag_centrifuga = st.text_input("TAG Centrífuga 2", key="tag_centrifuga")
    numero_livro_centrifuga = st.text_input("Número do Livro Ata da Centrífuga 2", key="numero_livro_centrifuga")
    horario_entrada_centrifuga = st.time_input("Horário de entrada na centrífuga", key="horario_entrada_centrifuga")
    horario_saida_centrifuga = st.time_input("Horário de saída da centrífuga", key="horario_saida_centrifuga")
    st.write("### Ultrassom")
    numero_livro_ultrassom = st.text_input("Número do Livro Ata de Ultrassom", key="numero_livro_ultrassom")
    data_anotacao_ultrassom = st.date_input("Data da Anotação no Ultrassom", key="data_anotacao_ultrassom")
    horario_entrada_ultrassom = st.time_input("Horário de entrada no ultrassom", key="horario_entrada_ultrassom")
    horario_saida_ultrassom = st.time_input("Horário de saída do ultrassom", key="horario_saida_ultrassom")
    st.write("### Transporte")
    numero_pacote = st.text_input("Número do pacote", key="numero_pacote")
    data_recebimento_pacote = st.date_input("Data de recebimento do pacote", key="data_recebimento_pacote")
    horario_recebimento_pacote = st.time_input("Horário de recebimento do pacote", key="horario_recebimento_pacote")
    transportadora = st.text_input("Transportadora", key="transportadora")
    st.write("### Evidências Visuais")
    caminho_evidencia = componente_imagem("rastreabilidade_tox", "Fotos das evidências, se aplicável")
    st.write("### Observações Geral")
    observacoes = st.text_area("Observações pertinentes:", key="observacoes_rastreabilidade_tox")
    return {
        "processo": "Rastreabilidade de Amostra (Tox)",
        "acompanhamento_amostra": {
            "codigo_amostra_acompanhada": codigo_amostra_acompanhada,
            "codigo_lote_acompanhado": codigo_lote_acompanhado,
            "tipo_amostra": tipo_amostra
        },
        "lcms": {
            "tag_lcms": tag_lcms,
            "numero_livro_lcms": numero_livro_lcms,
            "data_injecao": data_injecao.isoformat() if data_injecao else None,
            "horario_injecao": horario_injecao.isoformat() if horario_injecao else None,
            "criterios_curva": criterios_curva
        },
        "controles_rejeicoes": resultados_controles,
        "extracao": {
            "numero_livro_extracao": numero_livro_extracao,
            "data_inicio_extracao": data_inicio_extracao.isoformat() if data_inicio_extracao else None,
            "horario_entrada_extracao": horario_entrada_extracao.isoformat() if horario_entrada_extracao else None,
            "horario_saida_extracao": horario_saida_extracao.isoformat() if horario_saida_extracao else None
        },
        "centrifuga": {
            "tag_centrifuga": tag_centrifuga,
            "numero_livro_centrifuga": numero_livro_centrifuga,
            "horario_entrada_centrifuga": horario_entrada_centrifuga.isoformat() if horario_entrada_centrifuga else None,
            "horario_saida_centrifuga": horario_saida_centrifuga.isoformat() if horario_saida_centrifuga else None
        },
        "ultrassom": {
            "numero_livro_ultrassom": numero_livro_ultrassom,
            "data_anotacao_ultrassom": data_anotacao_ultrassom.isoformat() if data_anotacao_ultrassom else None,
            "horario_entrada_ultrassom": horario_entrada_ultrassom.isoformat() if horario_entrada_ultrassom else None,
            "horario_saida_ultrassom": horario_saida_ultrassom.isoformat() if horario_saida_ultrassom else None
        },
        "transporte": {
            "numero_pacote": numero_pacote,
            "data_recebimento_pacote": data_recebimento_pacote.isoformat() if data_recebimento_pacote else None,
            "horario_recebimento_pacote": horario_recebimento_pacote.isoformat() if horario_recebimento_pacote else None,
            "transportadora": transportadora
        },
        "evidencia_visual": caminho_evidencia,
        "observacoes": observacoes
    }

def processo_generico(nome_processo, setor):
    st.header(f"{nome_processo} ({setor})")
    info_logbook = componente_info_logbook(f"{nome_processo.lower().replace(' ', '_')}_{setor.lower().replace(' ', '_')}")
    integridade_dados = componente_integridade_dados(f"{nome_processo.lower().replace(' ', '_')}_{setor.lower().replace(' ', '_')}")
    avaliacao_detalhada = tabela_avaliacao_erros(f"{nome_processo.lower().replace(' ', '_')}_{setor.lower().replace(' ', '_')}")
    condicoes_logbook = componente_condicoes_logbook(f"{nome_processo.lower().replace(' ', '_')}_{setor.lower().replace(' ', '_')}")
    st.write("### Evidências Visuais")
    caminho_evidencia = componente_imagem(
        f"{nome_processo.lower().replace(' ', '_')}_{setor.lower().replace(' ', '_')}",
        "Adicionar foto das evidências, se aplicável"
    )
    st.write("### Observações Gerais")
    observacoes = st.text_area(
        "Observações pertinentes:",
        key=f"observacoes_{nome_processo.lower().replace(' ', '_')}_{setor.lower().replace(' ', '_')}"
    )
    return {
        "processo": f"{nome_processo} ({setor})",
        "info_logbook": info_logbook,
        "integridade_dados": integridade_dados,
        "avaliacao_detalhada": avaliacao_detalhada,
        "condicoes_logbook": condicoes_logbook,
        "evidencia_visual": caminho_evidencia,
        "observacoes": observacoes
    }

# Função Principal
def main():
    st.title("Sistema de Inspeção Laboratorial - Synvia")
    st.write("Sistema para registro e acompanhamento de inspeções nos setores Synvia Labs e Synvia Tox.")

    if 'etapa_atual' not in st.session_state:
        st.session_state.etapa_atual = 'informacoes_basicas'
    if 'dados_inspecao' not in st.session_state:
        st.session_state.dados_inspecao = {}

    gerenciador_inspetores = obter_instancia()

    with st.sidebar:
        st.header("Navegação")
        if st.button("Nova Inspeção"):
            st.session_state.etapa_atual = 'informacoes_basicas'
            st.session_state.dados_inspecao = {}
            st.rerun()
        if 'etapas_concluidas' in st.session_state:
            st.write("### Etapas Concluídas")
            for etapa in st.session_state.etapas_concluidas:
                st.write(f"✅ {etapa}")
        st.write("### Histórico de Inspeções")
        inspecoes = listar_inspecoes()
        if inspecoes:
            inspecao_selecionada = st.selectbox(
                "Selecione uma inspeção anterior:",
                options=[f"{insp['data_inspecao']} - {insp['empresa']} - {insp['processo']}" for insp in inspecoes],
                format_func=lambda x: x,
                key="historico_inspecoes"
            )
            if st.button("Carregar Inspeção", key="btn_carregar_inspecao"):
                idx = [f"{insp['data_inspecao']} - {insp['empresa']} - {insp['processo']}" for insp in inspecoes].index(inspecao_selecionada)
                id_inspecao = inspecoes[idx]['id_inspecao']
                ctx = get_sharepoint_context()
                if ctx:
                    try:
                        arquivo_inspecoes = f"{SHAREPOINT_DADOS_PATH}/inspecoes/inspecoes.json"
                        file_content = download_file_content(ctx, arquivo_inspecoes)
                        todas_inspecoes = json.loads(file_content.decode('utf-8')) if file_content else []
                        inspecao = next((i for i in todas_inspecoes if i.get('id_inspecao') == id_inspecao), None)
                        if inspecao:
                            st.session_state.dados_inspecao = inspecao
                            st.session_state.etapa_atual = 'conclusao'
                            st.rerun()
                        else:
                            st.error(f"Inspeção com ID {id_inspecao} não encontrada.")
                    except Exception as e:
                        st.error(f"Erro ao carregar inspeção: {e}")
        st.write("### Exportação de Dados")
        if st.button("Exportar Lista Completa", key="btn_exportar_sidebar"):
            caminho_relatorio = exportar_lista_completa_inspecoes()
            ctx = get_sharepoint_context()
            if caminho_relatorio and ctx:
                try:
                    file = ctx.web.get_file_by_server_relative_url(caminho_relatorio)
                    csv_data = file.read().execute_query()
                    st.download_button(
                        label="Baixar Relatório Completo CSV",
                        data=csv_data,
                        file_name=os.path.basename(caminho_relatorio),
                        mime="text/csv",
                        key="download_csv_sidebar"
                    )
                    caminho_excel = caminho_relatorio.replace('.csv', '.xlsx')
                    file_excel = ctx.web.get_file_by_server_relative_url(caminho_excel)
                    excel_data = file_excel.read().execute_query()
                    st.download_button(
                        label="Baixar Relatório Completo Excel",
                        data=excel_data,
                        file_name=os.path.basename(caminho_excel),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel_sidebar"
                    )
                except Exception as e:
                    st.error(f"Não foi possível gerar o relatório completo: {e}")

    # ----------- ETAPAS PRINCIPAIS DO FORMULÁRIO -----------
    if st.session_state.etapa_atual == 'informacoes_basicas':
        st.header("🔹 Informações da Inspeção")
        lista_inspetores = gerenciador_inspetores.obter_lista_inspetores()
        nome_inspetor = st.selectbox("Nome do Inspetor*", options=[""] + lista_inspetores, key="nome_inspetor")
        novo_inspetor = st.checkbox("Adicionar novo inspetor")
        if novo_inspetor:
            novo_nome = st.text_input("Nome do novo inspetor")
            novo_email = st.text_input("Email do novo inspetor")
            if st.button("Adicionar", key="btn_adicionar_inspetor") and novo_nome and novo_email:
                gerenciador_inspetores.adicionar_inspetor(novo_nome, novo_email)
                st.success(f"Inspetor {novo_nome} adicionado com sucesso!")
                lista_inspetores = gerenciador_inspetores.obter_lista_inspetores()
                nome_inspetor = novo_nome
                st.rerun()
        email_inspetor = gerenciador_inspetores.obter_email_por_nome(nome_inspetor) if nome_inspetor else ""
        st.text_input("Email do Inspetor*", value=email_inspetor, disabled=True)
        empresa = st.text_input("Empresa a ser Inspecionada*")
        data_inspecao = st.date_input("Data da Inspeção*")
        setor = st.radio("Setor / Setor-Inspeção*", ["Synvia Labs", "Synvia Tox"])
        laboratorio = None
        if setor == "Synvia Labs":
            laboratorio = st.selectbox(
                "Escolha o Laboratório",
                [
                    "Equivalência Farmacêutica (Campinas)",
                    "Equivalência Farmacêutica (Paulínia)",
                    "Bioequivalência",
                    "Laboratório de Microbiologia",
                    "Laboratório Clínico"
                ]
            )
        if st.button("Avançar para Seleção de Processo", key="btn_avancar_processo"):
            if not nome_inspetor or not empresa or not data_inspecao:
                st.error("Por favor, preencha todos os campos obrigatórios.")
            else:
                st.session_state.dados_inspecao['informacoes_basicas'] = {
                    "nome_inspetor": nome_inspetor,
                    "email_inspetor": email_inspetor,
                    "empresa": empresa,
                    "data_inspecao": data_inspecao.isoformat(),
                    "setor": setor,
                    "laboratorio": laboratorio
                }
                st.session_state.etapa_atual = 'selecao_processo'
                if 'etapas_concluidas' not in st.session_state:
                    st.session_state.etapas_concluidas = []
                st.session_state.etapas_concluidas.append("Informações da Inspeção")
                st.rerun()

    elif st.session_state.etapa_atual == 'selecao_processo':
        st.header("🔹 Escolher o Processo a ser Inspecionado")
        setor = st.session_state.dados_inspecao['informacoes_basicas']['setor']
        processos_disponiveis = {
            "Synvia Labs": [
                "Soluções",
                "Rastreabilidade de amostra",
                "Equipamentos",
                "Monitoramento ambiental",
                "Controle de temperatura de equipamentos",
                "Controle de temperatura ambiente"
            ],
            "Synvia Tox": [
                "Rastreabilidade de amostra",
                "Controle de temperatura de equipamentos",
                "Controle de temperatura ambiente"
            ]
        }
        processo_selecionado = st.selectbox(
            "Selecione o processo a ser inspecionado:",
            processos_disponiveis[setor],
            key="processo_selecionado"
        )
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Voltar", key="btn_voltar_processo"):
                st.session_state.etapa_atual = 'informacoes_basicas'
                st.session_state.etapas_concluidas.remove("Informações da Inspeção")
                st.rerun()
        with col2:
            if st.button("Avançar", key="btn_avancar_formulario"):
                st.session_state.dados_inspecao['processo_selecionado'] = processo_selecionado
                st.session_state.etapa_atual = 'formulario_processo'
                if 'etapas_concluidas' not in st.session_state:
                    st.session_state.etapas_concluidas = []
                st.session_state.etapas_concluidas.append("Seleção de Processo")
                st.rerun()

    elif st.session_state.etapa_atual == 'formulario_processo':
        processo = st.session_state.dados_inspecao['processo_selecionado']
        setor = st.session_state.dados_inspecao['informacoes_basicas']['setor']

        if processo == "Soluções":
            dados_formulario = processo_solucoes()
        elif processo == "Rastreabilidade de amostra" and setor == "Synvia Labs":
            dados_formulario = processo_rastreabilidade_amostra_labs()
        elif processo == "Rastreabilidade de amostra" and setor == "Synvia Tox":
            dados_formulario = processo_rastreabilidade_amostra_tox()
        elif processo == "Equipamentos":
            dados_formulario = processo_equipamentos()
        elif processo == "Monitoramento ambiental":
            dados_formulario = processo_monitoramento_ambiental()
        else:
            dados_formulario = processo_generico(processo, setor)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Voltar", key="btn_voltar_formulario"):
                st.session_state.etapa_atual = 'selecao_processo'
                try:
                    st.session_state.etapas_concluidas.remove("Seleção de Processo")
                except ValueError:
                    pass  # Evita erro se "Seleção de Processo" não estiver na lista
                st.rerun()
        with col2:
            if st.button("Salvar e Finalizar", key="btn_finalizar_formulario"):
                st.session_state.dados_inspecao['dados_formulario'] = dados_formulario
                id_inspecao = salvar_inspecao(st.session_state.dados_inspecao)
                if id_inspecao:
                    caminho_relatorio = gerar_relatorio(id_inspecao)
                    if caminho_relatorio:
                        st.session_state.dados_inspecao['caminho_relatorio'] = caminho_relatorio
                    st.session_state.etapa_atual = 'conclusao'
                    if 'etapas_concluidas' not in st.session_state:
                        st.session_state.etapas_concluidas = []
                    st.session_state.etapas_concluidas.append("Formulário do Processo")
                    st.rerun()
                else:
                    st.error("Erro ao salvar a inspeção. Tente novamente.")

    elif st.session_state.etapa_atual == 'conclusao':
        st.header("✅ Inspeção Finalizada")
        st.write("### Resumo da Inspeção")
        info_basicas = st.session_state.dados_inspecao['informacoes_basicas']
        processo = st.session_state.dados_inspecao['processo_selecionado']
        st.write(f"**Inspetor:** {info_basicas['nome_inspetor']}")
        st.write(f"**Empresa:** {info_basicas['empresa']}")
        try:
            st.write(f"**Data:** {datetime.fromisoformat(info_basicas['data_inspecao']).strftime('%d/%m/%Y')}")
        except ValueError:
            st.write(f"**Data:** {info_basicas['data_inspecao']} (formato inválido)")
        st.write(f"**Setor:** {info_basicas['setor']}")
        if info_basicas.get('laboratorio'):
            st.write(f"**Laboratório:** {info_basicas['laboratorio']}")
        st.write(f"**Processo:** {processo}")

        ctx = get_sharepoint_context()
        if 'caminho_relatorio' in st.session_state.dados_inspecao and ctx:
            caminho_relatorio = st.session_state.dados_inspecao['caminho_relatorio']
            try:
                file = ctx.web.get_file_by_server_relative_url(caminho_relatorio)
                csv_data = file.read().execute_query()
                st.download_button(
                    label="Baixar Relatório CSV",
                    data=csv_data,
                    file_name=os.path.basename(caminho_relatorio),
                    mime="text/csv",
                    key="download_csv_conclusao"
                )
                caminho_excel = caminho_relatorio.replace('.csv', '.xlsx')
                file_excel = ctx.web.get_file_by_server_relative_url(caminho_excel)
                excel_data = file_excel.read().execute_query()
                st.download_button(
                    label="Baixar Relatório Excel",
                    data=excel_data,
                    file_name=os.path.basename(caminho_excel),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_conclusao"
                )
            except Exception as e:
                st.warning(f"Erro ao baixar relatórios: {e}. Gere o relatório novamente se necessário.")

        st.write("### Evidência Visual")
        evidencia = st.session_state.dados_inspecao.get('dados_formulario', {}).get('evidencia_visual')
        if evidencia and ctx:
            try:
                file = ctx.web.get_file_by_server_relative_url(evidencia)
                image_data = file.read().execute_query()
                imagem = Image.open(io.BytesIO(image_data))
                st.image(imagem, caption="Evidência Visual da Inspeção",  use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao carregar imagem: {e}")
        else:
            st.info("Nenhuma evidência visual foi adicionada.")

        st.write("### Ações Adicionais")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Nova Inspeção", key="btn_nova_inspecao"):
                st.session_state.etapa_atual = 'informacoes_basicas'
                st.session_state.dados_inspecao = {}
                st.session_state.etapas_concluidas = []
                st.rerun()
        with col2:
            if st.button("Voltar ao Formulário", key="btn_voltar_conclusao"):
                st.session_state.etapa_atual = 'formulario_processo'
                try:
                    st.session_state.etapas_concluidas.remove("Formulário do Processo")
                except ValueError:
                    pass  # Evita erro se "Formulário do Processo" não estiver na lista
                st.rerun()

if __name__ == "__main__":
    main()
