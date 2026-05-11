import streamlit as st
import pandas as pd
from datetime import datetime

# 1. Carregamento dos dados da aba "Legendas"
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://google.com"
    # Lendo especificamente a aba Legendas
    df_legendas = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df_legendas

try:
    df_leg = carregar_legendas()
    # Pega a lista de clientes únicos para o primeiro select
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except Exception as e:
    st.error(f"Erro ao carregar aba Legendas: {e}")
    lista_clientes = ["Erro ao carregar"]

st.title("📝 Lançamento de Atividades")

# 2. LÓGICA DE FILTRO DINÂMICO (Fora do formulário para atualizar em tempo real)
# O Cliente precisa ser selecionado fora do st.form para o Solicitante atualizar na hora
col_c1, col_c2, col_c3 = st.columns(3)

with col_c1:
    data_atendimento = st.date_input("DATA", datetime.now())
    cliente_selecionado = st.selectbox("CLIENTE", options=lista_clientes)

# Filtra os solicitantes baseados no cliente escolhido na aba Legendas
# Importante: Sua aba Legendas deve ter as colunas "Clientes" e "Solicitante1" na mesma linha
try:
    filtro_solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    lista_solicitantes = sorted(filtro_solicitantes)
except:
    lista_solicitantes = []

# 3. FORMULÁRIO PARA OS DEMAIS CAMPOS
with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Cliente e Data já foram pegos acima, mas precisamos deles aqui para o envio
        ra = st.text_input("RA (Número)")
        situacao_ra = st.selectbox("SITUACAO_RA", ["Concluído", "Em Aberto", "Pendente"])
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")

    with col2:
        hr_inicio = st.time_input("HR_INICIO")
        hr_fim = st.time_input("HR_FIM")
        # Este campo agora é dinâmico!
        solicitante = st.selectbox("SOLICITANTE", options=lista_solicitantes)
        forma = st.text_input("FORMA")
        local = st.selectbox("LOCAL", ["Remoto", "Presencial"])

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
