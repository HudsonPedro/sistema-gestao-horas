import streamlit as st
import locale
import base64

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="HPTECH Sistema de Gestão",
    page_icon="hptechICO.png",
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

# 2. CSS PARA OCULTAR O MENU E FORÇAR A LOGO NO TOPO
st.markdown("""
    <style>
        /* Esconde o menu de páginas padrão do Streamlit */
        [data-testid="stSidebarNav"] {display: none;}
        
        /* Zera o espaçamento do topo para a logo subir */
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* Cor do título para o padrão azul CRTI */
        h1 { color: #b0231d; } /*#004a87 = AZUL CRTI*/
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
        /* Esconde o menu nativo */
        [data-testid="stSidebarNav"] {display: none;}
        
        /* FORÇA A LOGO PARA O TOPO ABSOLUTO */
        [data-testid="stSidebarContent"] {
            padding-top: 0rem !important;
        }

        /* Ajuste da logo para não encostar nas laterais */
        [data-testid="stSidebarHeader"] {
            padding-top: 0rem !important;
        }

        /* Estilo da caixinha de usuário */
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -10px; /* Puxa a caixinha um pouco para cima */
        }
    </style>
""", unsafe_allow_html=True)


# 3. SIDEBAR COM NOVO BOTÃO
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    st.markdown("---")
    # Identificação do Usuário
    u_email = st.user.get("email") or "hudson.valente@crti.com.br"
    st.markdown(f"""
        <div class="user-block">
            <span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br>
            <span style='font-size: 11px; color: #555;'>{u_email}</span>
        </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
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
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")
    
# 4. CONTEÚDO DA HOME COM NOVO CARD
#st.title("Bem-vindo ao Sistema de Gestão")
#st.markdown("### Selecione uma das seções abaixo para começar.")
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Tenta carregar a imagem que está no seu repositório GitHub
try:
    img_base64 = get_image_base64("hptechICO.png")
    
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <h1 style="margin: 0; font-size: 2.5rem;">Bem-vindo ao Sistema de Gestão</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
        st.markdown("## Selecione uma das seções abaixo para começar.")
    )
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("Bem-vindo ao Sistema de Gestão HPTECH")

st.markdown("---")

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
