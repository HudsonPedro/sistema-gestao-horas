import streamlit as st
import pandas as pd
from datetime import datetime

# 1. Função para carregar dados (Igual ao Dashboard)
@st.cache_data(ttl=600)
def carregar_dados():
    url = "https://google.com"
    dict_abas = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    return dict_abas

# Carrega os dados para popular os selects
try:
    dict_abas = carregar_dados()
    # Usamos a aba do mês atual ou a primeira disponível para pegar as listas
    aba_referencia = list(dict_abas.keys())[0] 
    df_ref = dict_abas[aba_referencia]
    
    # Extrai listas únicas removendo valores vazios
    lista_clientes = sorted(df_ref["CLIENTE"].dropna().unique().tolist())
    lista_solicitantes = sorted(df_ref["SOLICITANTE"].dropna().unique().tolist())
    lista_situacao = sorted(df_ref["SITUACAO_RA"].dropna().unique().tolist())
except:
    # Fallback caso a planilha falhe
    lista_clientes = ["Erro ao carregar"]
    lista_solicitantes = ["Erro ao carregar"]
    lista_situacao = ["Concluído", "Pendente", "Em Aberto"]

# --- Interface ---
st.title("📝 Lançamento de Atividades")

with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.date_input("DATA", datetime.now())
        # AQUI OS CAMPOS PUXANDO DA PLANILHA:
        st.selectbox("CLIENTE", options=lista_clientes)
        st.text_input("RA (Número)")
        st.selectbox("SITUACAO_RA", options=lista_situacao)
        st.text_input("CONSULTOR", value="Hudson Valente")

    with col2:
        st.time_input("HR_INICIO")
        st.time_input("HR_FIM")
        st.selectbox("SOLICITANTE", options=lista_solicitantes)
        st.text_input("FORMA")
        st.selectbox("LOCAL", ["Remoto", "Presencial"])

    # ... (restante das colunas col3 e campos de texto)
    
    st.form_submit_button("Salvar na Planilha")
