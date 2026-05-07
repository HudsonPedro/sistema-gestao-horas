import streamlit as st
import locale

# 1. CONFIGURAÇÃO DA PÁGINA (Sempre a primeira linha!)
st.set_page_config(
    page_title="CRTI - Sistema de Gestão",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* O BLOCO QUE VOCÊ PERGUNTOU ENTRA AQUI */
        .user-block {
            background-color: #f0f2f6;
            padding: 10px;
            border-radius: 8px;
            margin-top: 10px;
            margin-bottom: 10px;
            border: 1px solid #e0e0e0;
        }
    </style>
""", unsafe_allow_html=True)


# 2. CSS PARA ESCONDER O MENU PADRÃO (Movi para cá para garantir o topo)
st.markdown("""
    <style>
        /* Esconde o menu nativo do Streamlit */
        [data-testid="stSidebarNav"] {display: none;}
        
        /* Ajusta cores do ERP */
        [data-testid="stSidebar"] { background-color: #f8f9fa; }
        h1 { color: #004a87; }
        
        /* Remove o espaçamento extra que o menu escondido deixa */
        div[data-testid="stSidebarUserContent"] { padding-top: 0rem; }
    </style>
""", unsafe_allow_html=True)

# 3. IDIOMA E LOGIN
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, 'pt_BR')

# 4. SIDEBAR PERSONALIZADA (Logo agora será a primeira coisa visual)
with st.sidebar:
    # 1. Logo
    st.image("crti.jpg", use_container_width=True)
    
    # 2. Lógica de Usuário Simplificada (Evita que o app caia se o e-mail falhar)
    try:
        email_google = st.user.get("email")
        u_display = email_google if email_google else "Usuário Local / Teste"
    except:
        u_display = "Usuário Local / Teste"

    st.markdown(f"""
        <div class="user-block">
            <span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br>
            <span style='font-size: 12px; color: #555;'>{u_display}</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.title("Menu Principal")
    
    # Botões de Navegação
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("app.py")
        
    if st.button("📊 Dashboard", use_container_width=True):
        st.switch_page("pages/01_📊_Dashboard.py")
        
    if st.button("📄 Relatórios", use_container_width=True):
        st.switch_page("pages/02_📄_Relatorios.py")
    
    st.divider()
    st.caption("v1.0.0 06052026")
    st.caption("©2026 Direitos reservados.")
    st.caption("HPTech Informática ME.")

# 5. CONTEÚDO DA HOME (Igual ao seu PDF)
st.title("Bem-vindo ao Sistema de Gestão de Horas")
st.markdown("### Olá! Escolha uma das seções no menu lateral para começar.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📊 Dashboard")
    st.write("Visualize indicadores e métricas em tempo real.")
    if st.button("Ir para Dashboards", key="btn_dash"):
        st.switch_page("pages/01_📊_Dashboard.py")

with col2:
    st.subheader("📄 Relatórios")
    st.write("Gere documentos detalhados e exporte dados.")
    if st.button("Ir para Relatórios", key="btn_rel"):
        st.switch_page("pages/02_📄_Relatorios.py")

st.divider()
st.info("Qualquer dúvida, entre em contato com o suporte interno.")
