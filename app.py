import streamlit as st
import locale

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="HPTECH Sistema de Gestão",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. CSS PARA ESCONDER O MENU PADRÃO E CENTRALIZAR LOGO
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {
            padding-top: 0rem !important;
            margin-top: -30px !important;
        }
        [data-testid="stSidebarHeader"] {
            padding: 0px !important;
            text-align: center !important;
        }
        [data-testid="stSidebar"] img {
            display: block;
            margin-left: auto;
            margin-right: auto;
            width: 85%;
        }
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -15px;
            text-align: left;
        }
        h1 { color: #004a87; }
    </style>
""", unsafe_allow_html=True)

# 3. SIDEBAR COM NOVO BOTÃO
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    
    # Identificação do Usuário
    u_email = st.user.get("email") or "hudson.valente@crti.com.br"
    st.markdown(f"""
        <div class="user-block">
            <span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br>
            <span style='font-size: 11px; color: #555;'>{u_email}</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.title("Menu Principal")
    
    # Navegação Atualizada
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True):
        st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📄 Relatórios", use_container_width=True):
        st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True):
        st.switch_page("pages/03_📝_Lancamento.py")
    
    st.divider()
    st.caption("v1.5 - Unificado")

# 4. CONTEÚDO DA HOME COM NOVO CARD
st.title("Bem-vindo ao Sistema de Gestão CRTI")
st.markdown("### Selecione uma das seções abaixo para começar.")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📈 Dashboard")
    st.write("Visualize indicadores e métricas.")
    if st.button("Ir para Dashboards", key="btn_dash"):
        st.switch_page("pages/01_📊_Dashboard.py")

with col2:
    st.subheader("📄 Relatórios")
    st.write("Gere PDFs e envie por e-mail.")
    if st.button("Ir para Relatórios", key="btn_rel"):
        st.switch_page("pages/02_📄_Relatorios.py")

with col3:
    st.subheader("📝 Lançamento")
    st.write("Insira novos dados na planilha.")
    if st.button("Novo Lançamento", key="btn_input"):
        st.switch_page("pages/03_📝_Lancamento.py")

st.divider()
st.info("Sistema integrado com Google Sheets.")
