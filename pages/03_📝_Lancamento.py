import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Lançamento de Atividades", layout="wide")

# 2. CSS PARA PADRONIZAÇÃO
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] { padding-top: 0rem !important; margin-top: -30px !important; }
        [data-testid="stSidebarHeader"] { padding: 0px !important; text-align: center !important; }
        .user-block { background-color: #f0f2f6; padding: 10px; border-radius: 8px; margin-top: -15px; }
        h1 { color: #004a87; }
    </style>
""", unsafe_allow_html=True)

# 3. FUNÇÕES DE DADOS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTtNKWayx3w7y8FPuV_hsaYWcZsB6ftUBKpJALkFOnlYxLEbNfu3LH0y76qxQsGhg/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

from google.oauth2 import service_account

def conectar_google_sheets():
    # LINKS OBRIGATÓRIOS para o Google liberar a gravação
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

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

# 5. CARREGAMENTO DE LISTAS
try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
    lista_situacao = sorted(df_leg.iloc[:, 4].dropna().unique().tolist())
except:
    lista_clientes = ["Erro ao carregar"]
    lista_situacao = ["Concluído", "Pendente"]

st.title("📝 Lançamento de Atividades")

# 6. FILTRO DINÂMICO
col_top1, col_top2 = st.columns([1, 2])
with col_top1:
    data_atendimento = st.date_input("DATA", datetime.now())
    cliente_selecionado = st.selectbox("CLIENTE", options=lista_clientes)

try:
    lista_solicitantes = sorted(df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist())
except:
    lista_solicitantes = []

# 7. FORMULÁRIO
with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        ra = st.text_input("RA (Número)")
        situacao_ra = st.selectbox("SITUACAO_RA", options=lista_situacao)
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")
        local = st.selectbox("LOCAL", ["Remoto", "Presencial"])

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

    btn_enviar = st.form_submit_button("🚀 Gravar na Planilha")

# 8. LÓGICA DE GRAVAÇÃO
# 8. LÓGICA DE GRAVAÇÃO (Substitua apenas este bloco final)
if btn_enviar:
    try:
        with st.spinner("Gravando dados na planilha..."):
            client = conectar_google_sheets()
            
            # ID da planilha (confirmado)
            planilha_id = "1vTtNKWayx3w7y8FPuV_hsaYWcZsB6ftUBKpJALkFOnlYxLEbNfu3LH0y76qxQsGhg"
            sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ESouVyrO4Qlqm0oyQiFBbXVOEtbCLnCf/edit?gid=545789165#gid=545789165")
            
            # Ajuste dinâmico: tenta "Maio 2026", se não conseguir, tenta o nome exato que você definiu
            try:
                # Esta linha gera "Maio 2026" automaticamente baseada na data selecionada
                meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
                nome_aba_dinamico = f"{meses[data_atendimento.month - 1]} {data_atendimento.year}"
                aba = sheet.worksheet(nome_aba_dinamico)
            except:
                # Caso a lógica acima falhe, tenta o nome manual
                aba = sheet.worksheet("Maio 2026")
            
            # Organiza a linha para a planilha
                    # Organiza a linha na ordem exata das colunas da sua planilha:
        # A=DATA, B=HR_INICIO, C=HR_FIM, D=CLIENTE, E=RA... e assim por diante
            nova_linha = [
                data_atendimento.strftime('%d/%m/%Y'), # Coluna A (Data)
                hr_inicio.strftime('%H:%M'),           # Coluna B
                hr_fim.strftime('%H:%M'),             # Coluna C
                cliente_selecionado,                   # Coluna D
                ra,                                    # Coluna E
                situacao_ra,                           # Coluna F
                observacoes,                           # Coluna G
                consultor,                             # Coluna H
                solicitante,                           # Coluna I
                participante,                          # Coluna J
                forma,                                 # Coluna K
                local,                                 # Coluna L
                hr_inicio_d.strftime('%H:%M'),         # Coluna M
                hr_fim_d.strftime('%H:%M'),           # Coluna N
                km_d,                                  # Coluna O
                forma_d,                               # Coluna P
                descricao_d                            # Coluna Q
            ]
        
        # Grava na planilha          
            aba.append_row(nova_linha)
            st.success(f"✅ Lançamento para {cliente_selecionado} gravado com sucesso!")
            st.balloons()
            
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ Erro: A aba '{nome_aba_dinamico}' não foi encontrada na planilha.")
    except Exception as e:
        st.error(f"Erro ao gravar: {e}")
