import streamlit as st
import pandas as pd
from datetime import datetime

# 1. Carregamento dos dados da aba "Legendas"
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTtNKWayx3w7y8FPuV_hsaYWcZsB6ftUBKpJALkFOnlYxLEbNfu3LH0y76qxQsGhg/pub?output=xlsx"
    # Lendo especificamente a aba Legendas
    df_legendas = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df_legendas

try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
    
    # BUSCANDO A COLUNA E (SITUAÇÃO) DA ABA LEGENDAS
    # Supondo que o nome da coluna E seja 'Situacao' ou o índice 4
    lista_situacao = sorted(df_leg.iloc[:, 4].dropna().unique().tolist())
    
    # BUSCANDO A COLUNA DO CAMPO LOCAL (Suponha que seja a coluna F ou índice 5)
    lista_local = sorted(df_leg.iloc[:, 5].dropna().unique().tolist()) if df_leg.shape[1] > 5 else ["Remoto", "Presencial"]
except:
    lista_situacao = ["Concluído", "Em Aberto", "Pendente"]
    lista_local = ["Remoto", "Presencial"]

# --- FORMULÁRIO ---
with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        ra = st.text_input("RA (Número)")
        # SITUACAO_RA AGORA PUXA DA COLUNA E
        situacao_ra = st.selectbox("SITUACAO_RA", options=lista_situacao)
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")
        # LOCAL MOVIDO PARA CÁ (Abaixo de Consultor)
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
    # RECOLOCANDO OS CAMPOS QUE SUMIRAM:
    observacoes = st.text_area("OBSERVAÇÕES")
    participante = st.text_input("PARTICIPANTE")
    descricao_d = st.text_area("DESCRICAO_D")

    btn_enviar = st.form_submit_button("Salvar na Planilha")

if btn_enviar:
    # Lógica de salvar...
    st.success(f"Lançamento para {cliente_selecionado} registrado!")
