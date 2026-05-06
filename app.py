import streamlit as st
import locale

# 1. Configuração da Página (Deve ser a primeira linha de código Streamlit)
st.set_page_config(
    page_title="CRTI - Sistema de Gestão",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Configurar o idioma para Português (para datas, etc)
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, 'pt_BR')

# 3. Estilização CSS para aproximar do visual da imagem (Azul CRTI)
st.markdown("""
    <style>
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
    .main {
        background-color: #ffffff;
    }
    h1 {
        color: #004a87;
    }
    </style>
    """, unsafe_allow_html=True)

# 4. Menu Lateral (Sidebar)
with st.sidebar:
    st.image("crti.jpg", use_container_width=True)
    st.title("Menu Principal")
    st.info("Selecione uma das opções acima para navegar no sistema.")
    
    st.divider()
    st.caption("v1.0.0 - Unificado")
    st.caption("© 2024 CRTI Sistemas")

# 5. Conteúdo da Tela Principal (Home)
st.title("Bem-vindo ao Sistema Unificado CRTI")
st.markdown(f"### Olá! Escolha uma das seções no menu lateral para começar.")

# Criando cards visuais para a Home
col1, col2 = st.columns(2)

with col1:
    st.subheader("📈 Dashboard")
    st.write("Visualize indicadores, métricas e o desempenho em tempo real.")
    if st.button("Abrir Dashboards"):
        st.switch_page("pages/01_📊_Dashboard.py")

with col2:
    st.subheader("📄 Relatórios")
    st.write("Gere documentos detalhados e exporte dados para análise.")
    if st.button("Abrir Relatórios"):
        st.switch_page("pages/02_📄_Relatorios.py")

st.divider()
st.write("Qualquer dúvida, entre em contato com o suporte interno.")

