# Por Hudson Valente - HPTECH
import streamlit as st
import pandas as pd
import requests
import io

st.set_page_config(page_title="Histórico de Envios - HPTECH", page_icon="📜", layout="wide")

st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        .user-block { background-color: #f0f2f6; padding: 8px; border-radius: 8px; margin-top: -10px; }
        h1 { color: #b0231d; }
    </style>
""", unsafe_allow_html=True)

# SIDEBAR PADRÃO DO SEU ECOSSISTEMA
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    st.markdown("---")
    u_email = st.user.get("email") or "hudson.valente@crti.com.br"
    st.markdown(f'<div class="user-block">👤 <b>Usuário Logado</b><br><span style="font-size: 11px; color: #555;">{u_email}</span></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.title("Menu Principal")
    if st.button("🏠 Home", use_container_width=True): st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True): st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📄 Relatórios", use_container_width=True): st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True): st.switch_page("pages/03_📝_Lancamento.py")
    if st.button("💰 Medição Mensal", use_container_width=True): st.switch_page("pages/04_💰_Medicao_Mensal.py")
    if st.button("📜 Histórico de Envios", use_container_width=True): st.switch_page("pages/05_📜_Historico.py")

st.title("📜 Histórico e Auditoria de Envios")
st.markdown("---")

@st.cache_data(ttl=60)
def carregar_logs_auditoria():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    response = requests.get(url)
    df_dict = pd.read_excel(io.BytesIO(response.content), sheet_name=None)
    if "HISTORICO_ENVIOS" in df_dict:
        return df_dict["HISTORICO_ENVIOS"]
    return pd.DataFrame(columns=["DATA_ENVIO", "TIPO_RELATORIO", "IDENTIFICADOR", "DESTINATARIO", "USUARIO_LOGADO"])

with st.spinner("Carregando logs de auditoria..."):
    df_logs = carregar_logs_auditoria()

if df_logs.empty:
    st.info("Nenhum relatório foi enviado através do sistema até o momento.")
else:
    # Ordena para mostrar os envios mais recentes no topo da tabela
    df_logs = df_logs.iloc[::-1]
    
    # Filtros rápidos na tela
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        filtro_tipo = st.multiselect("Filtrar por Tipo de Documento:", options=list(df_logs["TIPO_RELATORIO"].unique()), default=list(df_logs["TIPO_RELATORIO"].unique()))
    with col_f2:
        busca_destinatario = st.text_input("Buscar por E-mail do Destinatário:")
        
    # Aplica os filtros digitados
    df_filtrado = df_logs[df_logs["TIPO_RELATORIO"].isin(filtro_tipo)]
    if busca_destinatario:
        df_filtrado = df_filtrado[df_filtrado["DESTINATARIO"].str.contains(busca_destinatario, case=False, na=False)]
        
    st.markdown("### Lançamentos Documentados no Google Sheets")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
