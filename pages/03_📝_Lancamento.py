import streamlit as st
import pandas as pd
from datetime import datetime

# 1. Configuração e Estilo (Mesmo padrão das outras páginas)
st.set_page_config(page_title="Input de Dados", layout="wide")

st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        .user-block { background-color: #f0f2f6; padding: 10px; border-radius: 8px; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

# 2. Sidebar (Copie sua Sidebar padrão aqui)
with st.sidebar:
    st.image("hptechNova.jpg", use_container_width=True)
    # ... (seu código de login/usuário e botões de navegação)
    if st.button("🏠 Home", use_container_width=True): st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True): st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📄 Relatórios", use_container_width=True): st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True): st.switch_page("pages/03_📝_Lancamento.py")

# 3. Formulário de Input
st.title("📝 Lançamento de Atividades")

with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        data_atendimento = st.date_input("DATA", datetime.now())
        cliente = st.selectbox("CLIENTE", ["Cliente A", "Cliente B", "Cliente C"]) # Substitua pela sua lista
        ra = st.text_input("RA (Número)")
        situacao_ra = st.selectbox("SITUACAO_RA", ["Concluído", "Em Aberto", "Pendente"])
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")

    with col2:
        hr_inicio = st.time_input("HR_INICIO")
        hr_fim = st.time_input("HR_FIM")
        solicitante = st.selectbox("SOLICITANTE", ["Nome 1", "Nome 2"])
        forma = st.text_input("FORMA")
        local = st.selectbox("LOCAL", ["Remoto", "Presencial"])

    with col3:
        hr_inicio_d = st.time_input("HR_INICIO_D (Desloc)")
        hr_fim_d = st.time_input("HR_FIM_D (Desloc)")
        km_d = st.number_input("KM_D", min_value=0.0, step=0.1)
        forma_d = st.text_input("FORMA_D")

    st.markdown("---")
    observacoes = st.text_area("OBSERVAÇÕES")
    participantes = st.text_input("PARTICIPANTE")
    descricao_d = st.text_area("DESCRICAO_D")

    btn_enviar = st.form_submit_button("Salvar na Planilha")

if btn_enviar:
    # Aqui entrará a lógica para salvar no Google Sheets
    st.success("Dados capturados com sucesso! (Conecte sua API do Google para gravar)")

