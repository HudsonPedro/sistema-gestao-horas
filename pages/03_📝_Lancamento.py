import streamlit as st
import pandas as pd
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Lançamento de Atividades", layout="wide")

# 2. CSS PARA PADRONIZAÇÃO (LOGO NO TOPO E CAIXA DE USUÁRIO)
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] { padding-top: 0rem !important; margin-top: -30px !important; }
        [data-testid="stSidebarHeader"] { padding: 0px !important; text-align: center !important; }
        .user-block { background-color: #f0f2f6; padding: 10px; border-radius: 8px; margin-top: -15px; }
        h1 { color: #004a87; }
    </style>
""", unsafe_allow_html=True)

# 3. FUNÇÃO PARA CARREGAR AS LEGENDAS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://google.com"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

# 4. SIDEBAR PADRÃO
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    u_email = st.user.get("email") or "hudson.valente@crti.com.br"
    st.markdown(f"<div class='user-block'><span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br><span style='font-size: 11px; color: #555;'>{u_email}</span></div>", unsafe_allow_html=True)
    st.title("Menu Principal")
    if st.button("🏠 Home", use_container_width=True): st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True): st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📄 Relatórios", use_container_width=True): st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True): st.switch_page("pages/03_📝_Lancamento.py")

# 5. CARREGAMENTO E LÓGICA DE FILTROS
try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
    # Coluna E (índice 4) para Situação
    lista_situacao = sorted(df_leg.iloc[:, 4].dropna().unique().tolist())
    # Coluna LOCAL (Supondo que seja a coluna G ou índice 6 - ajuste se necessário)
    lista_local = ["Remoto", "Presencial"] # Padrão caso não queira puxar da planilha
except:
    lista_clientes = ["Erro ao carregar"]
    lista_situacao = ["Concluído", "Em Aberto", "Pendente"]
    lista_local = ["Remoto", "Presencial"]

st.title("📝 Lançamento de Atividades")

# 6. INPUTS DINÂMICOS (Fora do formulário para o filtro funcionar)
col_top1, col_top2, col_top3 = st.columns(3)
with col_top1:
    data_atendimento = st.date_input("DATA", datetime.now())
    cliente_selecionado = st.selectbox("CLIENTE", options=lista_clientes)

# Filtra Solicitante baseado no Cliente
try:
    filtro_solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    lista_solicitantes = sorted(filtro_solicitantes)
except:
    lista_solicitantes = []

# 7. FORMULÁRIO DE LANÇAMENTO
with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        ra = st.text_input("RA (Número)")
        situacao_ra = st.selectbox("SITUACAO_RA", options=lista_situacao)
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")
        # LOCAL movido para baixo do Consultor
        local = st.selectbox("LOCAL", options=lista_local)

    with col2:
        hr_inicio = st.time_input("HR_INICIO")
        hr_fim = st.time_input("HR_FIM")
        solicitante = st.selectbox("SOLICITANTE", options=lista_solicitantes)
        forma = st.text_input("FORMA")

    with col3:
        hr_inicio_d = st.time_input("HR_INICIO_D (Desloc)")
        hr_fim_d = st.time_input("HR_FIM_D (Desloc)")
        km_d = st.number_input("KM_D", min_value=0.0, step=0.1)
        forma_d = st.text_input("FORMA_D")

    st.markdown("---")
    observacoes = st.text_area("OBSERVAÇÕES")
    participante = st.text_input("PARTICIPANTE")
    descricao_d = st.text_area("DESCRICAO_D")

    btn_enviar = st.form_submit_button("🚀 Salvar na Planilha")

if btn_enviar:
    # Para o envio funcionar, você precisará da lógica de gravação (gspread ou st.connection)
    st.success(f"Dados de {cliente_selecionado} processados com sucesso!")
