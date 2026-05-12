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
    lista_local = sorted(df_leg["Clientes"].dropna().unique().tolist())
except:
    lista_clientes = ["Erro ao carregar"]
    lista_situacao = ["Concluído", "Pendente"]
    lista_local = ["Erro ao Carregar"]

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
        local = st.selectbox("LOCAL", options=lista_situacao)
        lista_local = sorted(df_leg[df_leg["Clientes"] == cliente_selecionado]["Local"].dropna().unique().tolist())

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
# 8. LÓGICA DE GRAVAÇÃO
if btn_enviar:
    try:
        with st.spinner("Conectando ao Google Sheets..."):
            client = conectar_google_sheets()
            
            # ATENÇÃO: Verifique se este ID é o da planilha NOVA (sem .xlsx no nome)
            planilha_id = "1ESouVyrO4Qlqm0oyQiFBbXVOEtbCLnCf"
            
            # Tenta abrir a planilha
            try:
                sheet = client.open_by_key(planilha_id)
            except Exception as e:
                st.error("❌ Planilha não encontrada! Verifique o ID ou se você compartilhou a NOVA planilha com o e-mail da conta de serviço.")
                st.stop()

            # Define o nome da aba exatamente como está no Google
            nome_aba = "Maio 2026" 
            
            try:
                aba = sheet.worksheet(nome_aba)
            except:
                st.error(f"❌ A aba '{nome_aba}' não foi encontrada. Verifique se o nome está idêntico na planilha.")
                st.stop()
            
            # Organiza a linha conforme sua estrutura
            nova_linha = [
                data_atendimento.strftime('%d/%m/%Y'), 
                hr_inicio.strftime('%H:%M'),           
                hr_fim.strftime('%H:%M'),             
                cliente_selecionado,                   
                ra,                                    
                situacao_ra,                           
                observacoes,                           
                consultor,                             
                solicitante,                           
                participante,                          
                forma,                                 
                local,                                 
                hr_inicio_d.strftime('%H:%M'),         
                hr_fim_d.strftime('%H:%M'),           
                km_d,                                  
                forma_d,                               
                descricao_d                            
            ]
        
            # Grava na planilha
            aba.append_row(nova_linha, value_input_option='USER_ENTERED')
            st.success(f"✅ Sucesso! Lançamento para {cliente_selecionado} gravado.")
            st.balloons()
            
    except Exception as e:
        st.error(f"Erro inesperado: {e}")

        st.error(f"Erro ao gravar: {e}")
