import streamlit as st
import streamlit_authenticator as stauth
import locale
import base64

# =============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (DEVE SER A PRIMEIRA LINHA EXECUTADA)
# =============================================================================
st.set_page_config(
    page_title="HPTECH Sistema de Gestão",
    page_icon="hptech.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# 2. SISTEMA DE AUTENTICAÇÃO
# =============================================================================
# Cadastro de usuários com senhas criptografadas
credentials = {
    "usernames": {
        "admin": {
            "name": "Administrador",
            # Senha real: 
            "password": "$2b$12$e61m2rV9YfJzE6W1O8mEbe5D3JkHnK7f8z7G4H3X2B1C0D9E8F7G.", 
            "email": "hudsonpedro@gmail.com"
        },
        "usuario1": {
            "name": "Hudson Valente",
            # Senha real: 
            "password": "$2b$12$v9YfJzE6W1O8mEbe5D3JkHe61m2rV9YfJzE6W1O8mEbe5D3Jk.",
            "email": "hudson.pedro@hotmail.com"
        }
    }
}

authenticator = stauth.Authenticate(
    credentials,
    cookie_name="cookie_sistema_seguro",
    key="chave_secreta_aleatoria_123",
    cookie_expiry_days=30
)

# Renderiza o formulário de login centralizado na tela
# COMO DEVE FICAR SEGUINDO A NOVA DOCUMENTAÇÃO:
# 1. SUBSTITUA ESTA LINHA (Linha 46):
# name, authentication_status, username = authenticator.login(location='main', clear_on_submit=False)

# POR ESTA VERSÃO CORRETA DA VERSÃO 0.4.2:
# =============================================================================
# 2. SISTEMA DE AUTENTICAÇÃO
# =============================================================================

# ... (Mantenha o dicionário 'credentials' e a inicialização do 'authenticator' como estão)

# Renderiza o formulário de login centralizado na tela
authentication_status = authenticator.login(location='main')

# CORREÇÃO DA VALIDAÇÃO DO STATUS:
if st.session_state.get("authentication_status") is False:
    st.error("Usuário ou senha incorretos.")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("Por favor, digite seu usuário e senha para acessar.")
    st.stop()

# Se chegou aqui, o login foi bem-sucedido!
username = st.session_state["username"]
name = st.session_state["name"]


# Se o código passar daqui, significa que o usuário está LOGADO com sucesso!


# =============================================================================
# 3. SEU SISTEMA EM PRODUÇÃO ORIGINAL (INALTERADO)
# =============================================================================

# Estilos CSS originais
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
    
    # Identificação do Usuário conectado dinamicamente pelo login
    st.markdown(f"""
    <div class="user-block">
    <span style='font-size: 14px;'> <b>Usuário Logado</b></span><br>
    <span style='font-size: 11px; color: #555;'>{name} ({username})</span>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.title("Menu Principal")
    
    # Navegação Atualizada
    if st.button(" Home", use_container_width=True):
        st.switch_page("app.py")
    if st.button(" Dashboard", use_container_width=True):
        st.switch_page("pages/01_ _Dashboard.py")
    if st.button(" Lançamento de Horas", use_container_width=True):
        st.switch_page("pages/03_ _Lancamento.py")
    if st.button(" Relatórios RA", use_container_width=True):
        st.switch_page("pages/02_ _Relatorios.py") 
    if st.button(" Medição Mensal", use_container_width=True):
        st.switch_page("pages/04_ _Medicao_Mensal.py")
    if st.button(" Termo Homologação", use_container_width=True): 
        st.switch_page("pages/05_ _Termos.py")
    if st.button(" Termo Encerramento", use_container_width=True): 
        st.switch_page("pages/06_ _Termo_Encerramento.py")
    if st.button(" Termo Presencial", use_container_width=True): 
        st.switch_page("pages/07_ _Termo_Treinamento_Presencial.py")
    if st.button(" Reembolso de KM", use_container_width=True): 
        st.switch_page("pages/08_ _Reembolso_KM.py")
        
    st.divider()
    # Adiciona botão de Logout nativo no rodapé do menu lateral
    authenticator.logout('Sair do Sistema', 'sidebar')

# 4. CONTEÚDO DA HOME COM NOVO CARD
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Tenta carregar a imagem que está no repositório GitHub
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
    )
    st.markdown("Selecione uma das seções abaixo para começar.")
except:
    st.title("Bem-vindo ao Sistema de Gestão HPTECH")
    st.markdown("---")

col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)

with col1:
    st.subheader(" Dashboard")
    st.write("Visualize indicadores e métricas.")
    if st.button("Ir para Dashboards", key="btn_dash"):
        st.switch_page("pages/01_ _Dashboard.py")

with col2:
    st.subheader(" Lançamento")
    st.write("Insira novos dados na Base.")
    if st.button("Lançamento de Horas", key="btn_input"):
        st.switch_page("pages/03_ _Lancamento.py")

with col3:
    st.subheader(" Relatórios RA")
    st.write("Gere RA e envie o lote por e-mail.")
    if st.button("Gere o RA", key="btn_rel"):
        st.switch_page("pages/02_ _Relatorios.py")

with col4:
    st.subheader(" Medição")
    st.write("Fechamento mensal consolidado.")
    if st.button("Nova Medição", key="btn_med"):
        st.switch_page("pages/04_ _Medicao_Mensal.py")

with col5:
    st.subheader(" Homologação")
    st.write("Gere os Termos de Homologação.")
    if st.button("Novo Termo", key="btn_termo_homolog"):
        st.switch_page("pages/05_ _Termos.py")

with col6:
    st.subheader(" Encerramento")
    st.write("Gere os Termos de Encerramento.")
    if st.button("Novo Encerramento", key="btn_termo_encerra"):
        st.switch_page("pages/06_ _Termo_Encerramento.py")

with col7:
    st.subheader(" Presencial")
    st.write("Gere o Termo Presencial.")
    if st.button("Novo Termo Presencial", key="btn_termo_presencial"):
        st.switch_page("pages/07_ _Termo_Treinamento_Presencial.py")

with col8:
    st.subheader(" Reembolso de KM")
    st.write("Gere o Reembolso KM.")
    if st.button("Novo Reembolso KM", key="btn_reembolso_km"):
        st.switch_page("pages/08_ _Reembolso_KM.py")

st.divider()
st.info("Sistema integrado HPtech Informática ME.")
st.caption("v1.1 - 14072026")
st.caption("Todos os direitos reservados")
st.caption("Copyright ©2026 HPtech Informática ME")
