import streamlit as st
import pandas as pd
from datetime import datetime

# 1. Função para carregar os dados da aba "Legendas"
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://google.com"
    try:
        # Lê especificamente a aba "Legendas"
        df_legendas = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
        
        # Puxa as colunas solicitadas, removendo vazios e duplicados
        clientes = sorted(df_legendas["Clientes"].dropna().unique().tolist())
        solicitantes = sorted(df_legendas["Solicitante1"].dropna().unique().tolist())
        return clientes, solicitantes
    except Exception as e:
        # Caso a aba mude de nome ou dê erro, retorna listas padrão para não travar o app
        return ["Erro ao carregar Clientes"], ["Erro ao carregar Solicitantes"]

# Carregando as listas antes de montar o formulário
lista_clientes, lista_solicitantes = carregar_legendas()

# --- Interface ---
st.title("📝 Lançamento de Atividades")

with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        data_atendimento = st.date_input("DATA", datetime.now())
        # Busca da aba Legendas
        cliente = st.selectbox("CLIENTE", options=lista_clientes)
        ra = st.text_input("RA (Número)")
        situacao_ra = st.selectbox("SITUACAO_RA", ["Concluído", "Em Aberto", "Pendente"])
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")

    with col2:
        hr_inicio = st.time_input("HR_INICIO")
        hr_fim = st.time_input("HR_FIM")
        # Busca da aba Legendas
        solicitante = st.selectbox("SOLICITANTE", options=lista_solicitantes)
        forma = st.text_input("FORMA")
        local = st.selectbox("LOCAL", ["Remoto", "Presencial"])

    with col3:
        hr_inicio_d = st.time_input("HR_INICIO_D (Desloc)")
        hr_fim_d = st.time_input("HR_FIM_D (Desloc)")
        km_d = st.number_input("KM_D", min_value=0.0, step=0.1)
        forma_d = st.text_input("FORMA_D")

    # RECOLOCANDO OS CAMPOS QUE HAVIAM SUMIDO:
    st.markdown("---")
    observacoes = st.text_area("OBSERVAÇÕES", help="Detalhes do atendimento")
    participante = st.text_input("PARTICIPANTE")
    descricao_d = st.text_area("DESCRICAO_D", help="Descrição do deslocamento")

    st.markdown("<br>", unsafe_allow_html=True)
    btn_enviar = st.form_submit_button("Salvar na Planilha", type="primary")

if btn_enviar:
    st.info("Funcionalidade de salvamento em desenvolvimento (Requer API de Escrita do Google).")
