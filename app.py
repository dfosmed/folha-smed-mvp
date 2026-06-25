import os
import io
import re
import logging
import streamlit as st
import pandas as pd
import importlib
from dotenv import load_dotenv

# Setup basic logging (OWASP A09: Security Logging and Monitoring Failures)
logging.basicConfig(
    filename='app.log',
    level=logging.ERROR,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def secure_filename_custom(filename):
    """Sanitize filename to prevent Path Traversal attacks."""
    return re.sub(r'[^a-zA-Z0-9_.-]', '_', os.path.basename(filename))

load_dotenv()
import src.extrator_pdf
import src.classificador_rubricas
import src.gerador_excel
import src.gerador_contabilizacao
import src.gerar_consignados

importlib.reload(src.extrator_pdf)
importlib.reload(src.classificador_rubricas)
importlib.reload(src.gerador_excel)
importlib.reload(src.gerador_contabilizacao)
importlib.reload(src.gerar_consignados)

from src.extrator_pdf import extrair_dados_pdf
from src.classificador_rubricas import classificar_rubricas
from src.gerador_excel import gerar_excel_extraido
from src.gerador_contabilizacao import parse_and_fill_contabilizacao
from src.gerar_consignados import gerar_consignados_excel

# Garantir que as pastas existam
os.makedirs("temp", exist_ok=True)
os.makedirs("outputs", exist_ok=True)

st.set_page_config(page_title="Extrator de Folha SMED", page_icon="💹", layout="wide", initial_sidebar_state="collapsed")

# Custom CSS for UI/UX improvements (Dark Premium Theme)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Dark Premium Background */
    .stApp {
        background-color: #0B0F19;
        background-image: radial-gradient(circle at 50% 0%, #17243b 0%, #0B0F19 70%);
        color: #E2E8F0;
    }
    
    /* Headers */
    h1, h2, h3 {
        color: #F8FAFC !important;
        font-weight: 800 !important;
        letter-spacing: -0.02em;
    }
    
    /* Paragraphs */
    .stMarkdown p {
        color: #94A3B8;
        font-size: 1.05rem;
        font-weight: 300;
    }
    
    /* Stylish Buttons */
    .stButton>button {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: #ffffff !important;
        border-radius: 12px;
        border: 1px solid rgba(16, 185, 129, 0.3);
        padding: 0.6rem 1.2rem;
        font-weight: 600;
        letter-spacing: 0.02em;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 20px -2px rgba(16, 185, 129, 0.4);
    }
    .stButton>button p {
        color: #ffffff !important;
        font-weight: 600;
    }
    .stButton>button:hover {
        transform: translateY(-3px) scale(1.02);
        box-shadow: 0 10px 25px -3px rgba(16, 185, 129, 0.6);
        border: 1px solid rgba(16, 185, 129, 0.5);
    }
    .stButton>button:active {
        transform: translateY(1px);
    }
    
    /* Metrics Styling */
    div[data-testid="stMetricValue"] {
        color: #34D399;
        font-size: 2.5rem;
        font-weight: 800;
        text-shadow: 0 2px 10px rgba(52, 211, 153, 0.2);
    }
    div[data-testid="stMetricLabel"] {
        color: #94A3B8 !important;
        font-weight: 600;
        font-size: 1rem;
    }
    
    /* Alerts & Notifications */
    .stAlert {
        background-color: rgba(30, 41, 59, 0.7) !important;
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
        color: #E2E8F0;
    }
    
    /* File Uploader Glassmorphism */
    div[data-testid="stFileUploader"] {
        background-color: rgba(15, 23, 42, 0.6);
        backdrop-filter: blur(12px);
        border-radius: 16px;
        padding: 25px;
        border: 2px dashed rgba(16, 185, 129, 0.4);
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(16, 185, 129, 0.8);
        background-color: rgba(15, 23, 42, 0.8);
    }
    
    /* Dataframes/Tables */
    [data-testid="stDataFrame"] {
        background-color: rgba(15, 23, 42, 0.8);
        border-radius: 12px;
        padding: 10px;
        border: 1px solid rgba(255, 255, 255, 0.05);
    }
    
    /* Input Fields */
    div[data-baseweb="input"] {
        background-color: rgba(15, 23, 42, 0.8) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 10px !important;
        color: white !important;
    }
    div[data-baseweb="input"]:focus-within {
        border-color: #10B981 !important;
        box-shadow: 0 0 0 1px #10B981 !important;
    }
    .stTextInput input {
        color: white !important;
        background-color: transparent !important;
    }
</style>
""", unsafe_allow_html=True)

# Authentication State
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "login_attempts" not in st.session_state:
    st.session_state["login_attempts"] = 0

# Login Flow
if not st.session_state["authenticated"]:
    st.markdown("<h1 style='text-align: center; margin-top: 10vh;'>🔐 Acesso Restrito</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Por favor, insira a senha para acessar o Extrator de Folha SMED.</p>", unsafe_allow_html=True)
    
    # Brute Force Protection (OWASP A07)
    if st.session_state["login_attempts"] >= 5:
        st.error("Muitas tentativas falhas. Acesso bloqueado temporariamente por segurança.")
        logger.warning("Bloqueio temporário de segurança ativado devido a múltiplas falhas de login.")
        st.stop()
        
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            senha = st.text_input("Senha", type="password", placeholder="Digite sua senha aqui")
            submit = st.form_submit_button("Acessar Sistema", use_container_width=True)
            
            if submit:
                # Require env var to be properly set
                env_pass = os.getenv("APP_PASSWORD")
                if env_pass and senha == env_pass:
                    st.session_state["authenticated"] = True
                    st.session_state["login_attempts"] = 0
                    st.rerun()
                else:
                    st.session_state["login_attempts"] += 1
                    logger.warning(f"Tentativa de login falha. Total: {st.session_state['login_attempts']}")
                    st.error("Senha incorreta. Acesso negado.")
    st.stop()

# Main App Header
st.title("💹 Extrator de Folha SMED")
st.markdown("<p style='margin-bottom: 2rem;'>Faça upload do PDF consolidado da folha para gerar a planilha extraída e classificada de forma automática e eficiente.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📥 Arraste ou selecione o PDF da folha", type=["pdf"])

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("### ⚙️ Configuração de Dotações")
st.markdown("Marque abaixo quais dotações deseja aplicar para cada grupo. É possível selecionar mais de uma por linha.")

if 'df_config_dotacoes' not in st.session_state:
    data = []
    seen = set()
    for k, v in src.gerador_contabilizacao.CONFIG_BLOCOS.items():
        dc = v.get("Dotacao_CCusto")
        dr = v.get("Dotacao_Regime")
        if dc and dr:
            comb = (dc, dr)
            if comb not in seen:
                seen.add(comb)
                data.append({
                    "Centro de Custo": dc,
                    "Regime": dr,
                    "DOTAÇÃO 25%": False,
                    "DOTAÇÃO FUNDEB": False,
                    "DOTAÇÃO EXTRA 25%": False
                })
    st.session_state['df_config_dotacoes'] = pd.DataFrame(data)

edited_df = st.data_editor(
    st.session_state['df_config_dotacoes'],
    hide_index=True,
    disabled=["Centro de Custo", "Regime"],
    use_container_width=True
)
st.session_state['df_config_dotacoes'] = edited_df

if 'processamento_concluido' not in st.session_state:
    st.session_state['processamento_concluido'] = False

if st.button("🚀 Iniciar Processamento", use_container_width=True):
    if uploaded_file is None:
        st.warning("Por favor, faça o upload de um arquivo PDF antes de processar.")
    else:
        # Prevenção DoS: Verifica tamanho do arquivo (Max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            st.error("O arquivo excedeu o limite de segurança de 50MB. Operação cancelada.")
            st.stop()
            
        # Sanitização: Path Traversal Protection
        safe_filename = secure_filename_custom(uploaded_file.name)
        
        # Salva arquivo temporário
        temp_pdf_path = os.path.join("temp", safe_filename)
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        try:
            with st.spinner("Extraindo dados do PDF... Isso pode demorar alguns minutos dependendo do tamanho do arquivo."):
                df_original = extrair_dados_pdf(temp_pdf_path)
            
            if df_original is None or df_original.empty:
                st.error("Nenhum dado pôde ser extraído. O PDF pode não conter texto extraível ou não estar no formato esperado.")
            else:
                with st.spinner("Classificando rubricas..."):
                    df_classificado = classificar_rubricas(df_original)
                
                qtd_linhas_extraidas = len(df_original)
                
                if "Página" in df_original.columns:
                    qtd_paginas = df_original["Página"].nunique()
                elif "Pagina" in df_original.columns:
                    qtd_paginas = df_original["Pagina"].nunique()
                else:
                    qtd_paginas = 0
                
                df_pendencias = df_classificado[df_classificado["Status"] == "Pendente"]
                qtd_pendencias = len(df_pendencias)
                qtd_classificadas = qtd_linhas_extraidas - qtd_pendencias
                
                # Salva métricas no state
                st.session_state['metrics'] = {
                    'paginas': qtd_paginas,
                    'linhas': qtd_linhas_extraidas,
                    'classificadas': qtd_classificadas,
                    'pendentes': qtd_pendencias
                }
                st.session_state['df_classificado'] = df_classificado
                st.session_state['df_pendencias'] = df_pendencias
                
                # Gera o arquivo Excel
                with st.spinner("Gerando arquivo Excel..."):
                    excel_bytes = gerar_excel_extraido(df_classificado, df_original)
                    excel_filename = safe_filename.replace('.pdf', '.PDF').replace('.PDF', '_classificado.xlsx')
                    excel_path = os.path.join("outputs", excel_filename)
                    try:
                        with open(excel_path, "wb") as f:
                            f.write(excel_bytes)
                    except Exception as ex:
                        pass
                    
                    st.session_state['excel_bytes'] = excel_bytes
                    st.session_state['excel_filename'] = excel_filename
                
                # Preparar Contabilização
                contabilizacao_sucesso = False
                with st.spinner("Gerando arquivo de Contabilização..."):
                    try:
                        import base64
                        df_resumo = pd.read_excel(io.BytesIO(excel_bytes), sheet_name='Resumo Geral')
                        
                        modelo_b64 = os.getenv("MODELO_XLSX_B64")
                        if modelo_b64:
                            modelo_source = base64.b64decode(modelo_b64)
                        else:
                            modelo_source = 'modelo_contabilização.xlsx'
                            
                        # Build config dict from dataframe
                        config_dotacoes_dict = {}
                        for _, row in st.session_state['df_config_dotacoes'].iterrows():
                            selecionadas = []
                            if row.get("DOTAÇÃO 25%"): selecionadas.append("DOTAÇÃO 25%")
                            if row.get("DOTAÇÃO FUNDEB"): selecionadas.append("DOTAÇÃO FUNDEB")
                            if row.get("DOTAÇÃO EXTRA 25%"): selecionadas.append("DOTAÇÃO EXTRA 25%")
                            cc = str(row.get("Centro de Custo", "")).strip().upper()
                            rg = str(row.get("Regime", "")).strip().upper()
                            config_dotacoes_dict[f"{cc}_{rg}"] = selecionadas
                            
                        contabilizacao_bytes = parse_and_fill_contabilizacao(df_resumo, modelo_source, config_dotacoes_dict)
                        contabilizacao_filename = safe_filename.replace('.pdf', '.PDF').replace('.PDF', '_contabilizacao.xlsx')
                        contabilizacao_path = os.path.join("outputs", contabilizacao_filename)
                        
                        try:
                            with open(contabilizacao_path, "wb") as f:
                                f.write(contabilizacao_bytes)
                        except Exception as ex:
                            pass
                        
                        st.session_state['contabilizacao_bytes'] = contabilizacao_bytes
                        st.session_state['contabilizacao_filename'] = contabilizacao_filename
                        contabilizacao_sucesso = True
                    except Exception as e:
                        logger.error(f"Erro ao gerar contabilizacao: {str(e)}", exc_info=True)
                        st.error("Ocorreu um erro interno ao gerar o arquivo de contabilização.")
                        st.session_state['contabilizacao_bytes'] = None
                        
                # Preparar Consignados
                consignados_sucesso = False
                with st.spinner("Gerando arquivo de Consignados..."):
                    try:
                        consignados_bytes = gerar_consignados_excel(df_resumo)
                        consignados_filename = safe_filename.replace('.pdf', '.PDF').replace('.PDF', '_consignados.xlsx')
                        consignados_path = os.path.join("outputs", consignados_filename)
                        
                        try:
                            with open(consignados_path, "wb") as f:
                                f.write(consignados_bytes)
                        except Exception as ex:
                            pass
                        
                        st.session_state['consignados_bytes'] = consignados_bytes
                        st.session_state['consignados_filename'] = consignados_filename
                        consignados_sucesso = True
                    except Exception as e:
                        logger.error(f"Erro ao gerar consignados: {str(e)}", exc_info=True)
                        st.error("Ocorreu um erro interno ao gerar o arquivo de consignados.")
                        st.session_state['consignados_bytes'] = None
                
                st.session_state['processamento_concluido'] = True
                
        except Exception as e:
            logger.error(f"Erro geral no processamento: {str(e)}", exc_info=True)
            st.error("Ocorreu um erro inesperado no processamento do arquivo. Consulte o log do sistema ou contate o suporte.")
            
        finally:
            try:
                if os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
            except:
                pass

# Exibição dos resultados e botões (Fora do bloco if st.button para persistir após download)
if st.session_state.get('processamento_concluido'):
    st.success("✅ Extração e Classificação concluídas com sucesso!")
    
    st.markdown("### 📊 Visão Geral")
    m = st.session_state['metrics']
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Páginas", m['paginas'])
    col2.metric("Total Linhas", m['linhas'])
    col3.metric("Classificadas", m['classificadas'])
    col4.metric("Pendentes", m['pendentes'])
    
    st.markdown("<hr style='margin: 2rem 0;'>", unsafe_allow_html=True)
    st.markdown("### 🏷️ Classificação de Rubricas")
    
    qtd_pend = m['pendentes']
    if qtd_pend == 0:
        st.success("Todas as rubricas foram classificadas.")
    else:
        st.warning(f"Existem {qtd_pend} rubricas pendentes de classificação.")
    
    with st.expander("Prévia das Rubricas Classificadas", expanded=True):
        st.dataframe(st.session_state['df_classificado'], width="stretch")
        
    with st.expander("Pendências", expanded=(qtd_pend > 0)):
        if qtd_pend > 0:
            st.dataframe(st.session_state['df_pendencias'], width="stretch")
        else:
            st.info("Nenhuma pendência encontrada.")

    st.markdown("### 📥 Baixar Arquivos Gerados")
    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    
    with col_btn1:
        st.download_button(
            label="⬇️ Baixar Excel Completo",
            data=st.session_state['excel_bytes'],
            file_name=st.session_state['excel_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col_btn2:
        if st.session_state.get('contabilizacao_bytes'):
            st.download_button(
                label="⬇️ Baixar Contabilização",
                data=st.session_state['contabilizacao_bytes'],
                file_name=st.session_state['contabilizacao_filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("A contabilização não pôde ser gerada.")
            
    with col_btn3:
        if st.session_state.get('consignados_bytes'):
            st.download_button(
                label="⬇️ Baixar Consignados",
                data=st.session_state['consignados_bytes'],
                file_name=st.session_state['consignados_filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("Consignados não gerado.")
