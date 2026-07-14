import streamlit as st
import locale
import base64

# =============================================================================
# 1. BLOCO DE LOGIN NATIVO (ADICIONADO NO TOPO DO ARQUIVO)
# =============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.set_page_config(page_title="HPTECH Sistema de Gestão", page_icon="hptech.png", layout="centered")
    st.title("🔑 HPTECH - Controle de Acesso")
    
    with st.form("formulario_login"):
        usuario_input = st.text_input("Username")
        senha_input = st.text_input("Password", type="password")
        botao_entrar = st.form_submit_button("Login", use_container_width=True)
        
        if botao_entrar:
            if usuario_input == "admin" and senha_input == "Admin@2026":
                st.session_state["autenticado"] = True
                st.session_state["u_email"] = "hudsonpedro@gmail.com"
                st.rerun()
            elif usuario_input == "usuario1" and senha_input == "Mudar@123":
                st.session_state["autenticado"] = True
                st.session_state["u_email"] = "hudson.pedro@hotmail.com"
                st.rerun()
            else:
                st.error("❌ Usuário ou senha incorretos.")
    st.stop()  # Trava o carregamento se não estiver autenticado


# =============================================================================
# 2. SEU SISTEMA EM PRODUÇÃO ORIGINAL (ABSOLUTAMENTE INTACTO A PARTIR DAQUI)
# =============================================================================

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
 page_title="HPTECH Sistema de Gestão",
 page_icon="hptech.png",
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
 u_email = st.session_state.get("u_email") or "hudson.valente@crti.com.br"
 st.markdown(f"""
 <div class="user-block">
 <span style='font-size: 14px;'> 👤 <b>Usuário Logado</b></span><br>
 <span style='font-size: 11px; color: #555;'>{u_email}</span>
 </div>
 """, unsafe_allow_html=True)
 st.markdown("---")
 st.title("Menu Principal")
 
 # Navegação Atualizada
 if st.button("🏠 Home", use_container_width=True):
  st.switch_page("app.py")
 if st.button("📊 Dashboard", use_container_width=True):
  st.switch_page("pages/01_ _Dashboard.py")
 if st.button("📝 Lançamento de Horas", use_container_width=True):
  st.switch_page("pages/03_ _Lancamento.py")
 if st.button("📄 Relatórios RA", use_container_width=True):
  st.switch_page("pages/02_ _Relatorios.py") 
 if st.button("💰 Medição Mensal", use_container_width=True):
  st.switch_page("pages/04_ _Medicao_Mensal.py")
 if st.button("📜 Termo Homologação", use_container_width=True): 
  st.switch_page("pages/05_ _Termos.py")
 if st.button("📄 Termo Encerramento", use_container_width=True): 
  st.switch_page("pages/06_ _Termo_Encerramento.py")
 if st.button("🚗 Termo Presencial", use_container_width=True): 
  st.switch_page("pages/07_ _Termo_Treinamento_Presencial.py")
 if st.button("💰 Reembolso de KM", use_container_width=True): 
  st.switch_page("pages/08_ _Reembolso_KM.py")
 
 st.divider()
 
 if st.button("🚪 Sair", use_container_width=True):
  st.session_state["autenticado"] = False
  st.rerun()
  
 st.caption("v1.0 - 11052026")
 st.caption("Todos os direitos reservados")
 st.caption("Copyright ©2026 HPtech Informática ME")
 
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
 <h1 style="margin: 0; font-size: 2.5rem;">Ben-vindo ao Sistema de Gestão</h1>
 <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
 </div>
 """, 
 unsafe_allow_html=True
 )
 st.markdown("Selecione uma das seções abaixo para começar.")
except:
 # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
st.title("Bem-vindo ao Sistema de Gestão HPTECH")
st.markdown("---")

col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)

with col1:
    st.subheader("📈 Dashboard")
    st.write("Visualize indicadores e métricas.")
    if st.button("Ir para Dashboards", key="btn_dash"):
        st.switch_page("pages/01_📊_Dashboard.py")

with col2:
    st.subheader("📝 Lançamento")
    st.write("Insira novos dados na Base.")
    if st.button("Lançamento de Horas", key="btn_input"):
        st.switch_page("pages/03_📝_Lancamento.py")

with col3:
    st.subheader("📄 Relatórios RA")
    st.write("Gere RA e envie o lote por e-mail.")
    if st.button("Gere o RA", key="btn_rel"):
        st.switch_page("pages/02_📄_Relatorios.py")
        
with col4:
    st.subheader("💰 Medição")
    st.write("Fechamento mensal consolidado.")
    if st.button("Nova Medição", key="btn_med"):
        st.switch_page("pages/04_💰_Medicao_Mensal.py")

with col5:
    st.subheader("📋 Homologação")
    st.write("Gere os Termos de Homologação.")
    if st.button("Novo Termo", key="btn_termo_homolog"):
        st.switch_page("pages/05_📋_Termos.py")

with col6:
    st.subheader("📑 Encerramento")
    st.write("Gere os Termos de Encerramento.")
    if st.button("Novo Encerramento", key="btn_termo_encerra"):
        st.switch_page("pages/06_📑_Termo_Encerramento.py")

with col7:
    st.subheader("🚗 Presencial")
    st.write("Gere o Termo Presencial.")
    if st.button("Novo Termo Presencial", key="btn_termo_presencial"):
        st.switch_page("pages/07_🚗_Termo_Treinamento_Presencial.py")

with col8:
    st.subheader("💰 Reembolso de KM")
    st.write("Gere o Reembolso KM.")
    if st.button("Novo Reembolso KM", key="btn_reembolso_km"):
        st.switch_page("pages/08_💰_Reembolso_KM.py")

st.divider()
st.info("Sistema integrado HPtech Informática ME.")
st.info("v1.0 - 11052026 | Todos os direitos reservados.")
st.caption("Copyright ©2026 HPtech Informática ME")
