import streamlit as st
import locale
import base64
import hashlib
import psycopg2
import os

# =============================================================================
# BANCO DE DADOS DE USUÁRIOS SEGURO (SQLITE)
# =============================================================================
# =============================================================================
# BANCO DE DADOS DE USUÁRIOS SEGURO (SQLITE) - LIMPO SEM CREDENCIAIS EXPOSTAS
# =============================================================================
def conectar_banco():

    conn = psycopg2.connect(
        st.secrets["DATABASE_URL"]
    )

    cursor = conn.cursor()

    return conn, cursor

def criptografar_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

# Inicializa apenas a estrutura da tabela em produção (Sem injetar dados via código)
conn, cursor = conectar_banco()
conn.close()


# =============================================================================
# 1. BLOCO DE LOGIN
# =============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.set_page_config(page_title="HPTECH Sistema de Gestão", page_icon="hptech.png", layout="wide")
    col_esq, col_centro, col_dir = st.columns([1.5, 1, 1.5])
    
    with col_centro:
        st.write("") 
        try:
            with open("hptechICO.png", "rb") as img_file:
                img_base64 = base64.b64encode(img_file.read()).decode()
            st.markdown(f'<div style="text-align: center; margin-bottom: 20px;"><img src="data:image/png;base64,{img_base64}" style="height: 200px;"></div>', unsafe_allow_html=True)
        except:
            pass 

        st.markdown("<h5 style='text-align: center; margin-bottom: 10px;'>🔑 HPTECH - Controle de Acesso</h5>", unsafe_allow_html=True)
        
        with st.form("formulario_login"):
            usuario_input = st.text_input("Usuário").strip().lower()
            senha_input = st.text_input("Senha", type="password")
            botao_entrar = st.form_submit_button("Login", use_container_width=True)
            
            if botao_entrar:
                conn, cursor = conectar_banco()
                cursor.execute("SELECT nome, senha_hash, email, status FROM usuarios WHERE username=?", (usuario_input,))
                user_data = cursor.fetchone()
                conn.close()
                
                if user_data:
                    nome, senha_hash_db, email, status = user_data
                    if status == "Bloqueado":
                        st.error("❌ Este usuário está bloqueado. Contate o administrador.")
                    elif criptografar_senha(senha_input) == str(senha_hash_db):
                        st.session_state["autenticado"] = True
                        st.session_state["u_email"] = email
                        st.session_state["u_name"] = nome
                        st.session_state["u_user"] = usuario_input
                        st.rerun()
                    else:
                        st.error("❌ Usuário ou senha incorretos.")
                else:
                    st.error("❌ Usuário ou senha incorretos.")
                    
        #st.markdown("<p style='text-align: center; color: #777; margin-top: 15px;'>Sistema Integrado HPtech Informática ME.</p>", unsafe_allow_html=True)
        _, col_centro, _ = st.columns([1, 40, 1])
        with col_centro:
            st.info("Sistema Integrado HPtech Informática\n v1.1|14072026|Copyright ©2026.", icon="ℹ️")

    st.stop()

u_email = st.session_state["u_email"]
u_name = st.session_state["u_name"]
u_user = st.session_state["u_user"]

# =============================================================================
# 2. SEU SISTEMA EM PRODUÇÃO ORIGINAL (RESTAURADO E HIGIENIZADO)
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
 [data-testid="stSidebarNav"] {display: none;}
 [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
 h1 { color: #b0231d; }
 </style>
""", unsafe_allow_html=True)

st.markdown("""
 <style>
 [data-testid="stSidebarNav"] {display: none;}
 [data-testid="stSidebarContent"] {
 padding-top: 0rem !important;
 }
 [data-testid="stSidebarHeader"] {
 padding-top: 0rem !important;
 }
 .user-block {
 background-color: #f0f2f6;
 padding: 8px;
 border-radius: 8px;
 margin-top: -10px;
 }
 </style>
""", unsafe_allow_html=True)

# 3. SIDEBAR COM NOVO BOTÃO
with st.sidebar:
 st.image("hptechNova.png", use_container_width=True)
 st.markdown("---")
 
 st.markdown(f"""
 <div class="user-block">
 <span style='font-size: 14px;'> 👤 <b>Usuário Logado</b></span><br>
 <span style='font-size: 11px; color: #555;'>{u_email}</span>
 </div>
 """, unsafe_allow_html=True)
 st.markdown("---")
 
 # --- NOVO: EXPANDER PARA ALTERAR A PRÓPRIA SENHA ---
 with st.expander("🔐 Alterar Minha Senha"):
  with st.form("form_trocar_senha", clear_on_submit=True):
   senha_atual = st.text_input("Senha Atual", type="password")
   nova_senha = st.text_input("Nova Senha", type="password")
   confirmar_nova = st.text_input("Confirme a Nova Senha", type="password")
   
   if st.form_submit_button("Atualizar Senha", use_container_width=True):
    if nova_senha != confirmar_nova:
     st.error("❌ As novas senhas não coincidem.")
    elif len(nova_senha) < 6:
     st.error("❌ A senha deve ter no mínimo 6 caracteres.")
    else:
     conn, cursor = conectar_banco()
     # Verifica se a senha atual está correta no banco
     cursor.execute("SELECT senha_hash FROM usuarios WHERE username=?", (u_user,))
     senha_db = cursor.fetchone()[0]
     
     if criptografar_senha(senha_atual) == senha_db:
      # Atualiza pela nova senha criptografada em SHA-256
      cursor.execute("UPDATE usuarios SET senha_hash=? WHERE username=?", (criptografar_senha(nova_senha), u_user))
      conn.commit()
      st.success("✅ Senha alterada com sucesso!")
     else:
      st.error("❌ Senha atual incorreta.")
     conn.close()
    
 st.title("Menu Principal")
 
 # Navegação Atualizada (Mapeamento limpo de caracteres ocultos)
 if st.button("🏠 Home", use_container_width=True):
     st.switch_page("app.py")
 if st.button("📊 Dashboard", use_container_width=True):
     st.switch_page("pages/01_📊_Dashboard.py")
 if st.button("📝 Lançamento de Horas", use_container_width=True):
     st.switch_page("pages/03_📝_Lancamento.py")
 if st.button("📄 Relatórios RA", use_container_width=True):
     st.switch_page("pages/02_📄_Relatorios.py")    
 if st.button("💰 Medição Mensal", use_container_width=True):
     st.switch_page("pages/04_💰_Medicao_Mensal.py")
 if st.button("📋 Termo Homologação", use_container_width=True): 
     st.switch_page("pages/05_📋_Termos.py")
 if st.button("📑 Termo Encerramento", use_container_width=True): 
     st.switch_page("pages/06_📑_Termo_Encerramento.py")
 if st.button("🚗 Termo Presencial", use_container_width=True): 
     st.switch_page("pages/07_🚗_Termo_Treinamento_Presencial.py")
 if st.button("💰 Reembolso de KM", use_container_width=True): 
     st.switch_page("pages/08_💰_Reembolso_KM.py")
  # --- NOVO: ABA DE ADMINISTRAÇÃO EXCLUSIVA PARA O ADMIN ---
 if u_user == "admin":
  st.markdown("---")
  if st.button("⚙️ Gerenciar Usuários", use_container_width=True):
   st.session_state["pagina_admin"] = True
  else:
   if "pagina_admin" not in st.session_state:
    st.session_state["pagina_admin"] = False
       
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
# =============================================================================
# INTERFACE VISUAL DE GERENCIAMENTO DE USUÁRIOS
# =============================================================================
if u_user == "admin" and st.session_state.get("pagina_admin"):
    st.markdown("---")
    st.header("⚙️ Painel de Controle de Usuários")
    
    aba1, aba2, aba3 = st.tabs(["🆕 Cadastrar", "✏️ Alterar / Bloquear", "❌ Excluir"])
    
    with aba1:
        with st.form("cadastrar_user"):
            new_user = st.text_input("Usuário (Login)").strip().lower()
            new_name = st.text_input("Nome Completo")
            new_email = st.text_input("E-mail")
            new_pass = st.text_input("Senha", type="password")
            if st.form_submit_button("Salvar Novo Usuário"):
                if new_user and new_pass:
                    conn, cursor = conectar_banco()
                    try:
                        cursor.execute("INSERT INTO usuarios VALUES (?,?,?,?,?)", (new_user, new_name, criptografar_senha(new_pass), new_email, "Ativo"))
                        conn.commit()
                        st.success("Usuário cadastrado com sucesso!")
                    except:
                        st.error("Este Usuário já existe.")
                    conn.close()
                else:
                    st.warning("Preencha o Usuário e a Senha.")

    with aba2:
        conn, cursor = conectar_banco()
        cursor.execute("SELECT username, nome, email, status FROM usuarios")
        lista_users = cursor.fetchall()
        conn.close()
        
        user_sel = st.selectbox("Selecione o usuário para modificar", [u[0] for u in lista_users if u[0] != 'admin'])
        
        if user_sel:
            curr_data = [u for u in lista_users if u[0] == user_sel][0]
            with st.form("alterar_user"):
                alt_name = st.text_input("Nome", value=curr_data[1])
                alt_email = st.text_input("E-mail", value=curr_data[2])
                alt_status = st.selectbox("Status", ["Ativo", "Bloqueado"], index=0 if curr_data[3] == "Ativo" else 1)
                alt_pass = st.text_input("Nova Senha (deixe em branco para não alterar)", type="password")
                
                if st.form_submit_button("Atualizar Dados"):
                    conn, cursor = conectar_banco()
                    if alt_pass:
                        cursor.execute("UPDATE usuarios SET nome=?, email=?, status=?, senha_hash=? WHERE username=?", (alt_name, alt_email, alt_status, criptografar_senha(alt_pass), user_sel))
                    else:
                        cursor.execute("UPDATE usuarios SET nome=?, email=?, status=? WHERE username=?", (alt_name, alt_email, alt_status, user_sel))
                    conn.commit()
                    conn.close()
                    st.success("Usuário atualizado!")

    with aba3:
        user_del = st.selectbox("Selecione o usuário para DELETAR permanentemente", [u[0] for u in lista_users if u[0] != 'admin'])
        if st.button("⚠️ CONFIRMAR EXCLUSÃO DEFINITIVA", type="primary"):
            if user_del:
                conn, cursor = conectar_banco()
                cursor.execute("DELETE FROM usuarios WHERE username=?", (user_del,))
                conn.commit()
                conn.close()
                st.success(f"Usuário {user_del} removido do sistema.")
                st.rerun()

    if st.button("Voltar para a Home"):
        st.session_state["pagina_admin"] = False
        st.rerun()

st.divider()
st.info("Sistema integrado HPtech Informática ME.")
st.info("v1.0 - 14072026 | Todos os direitos reservados.")
st.caption("Copyright ©2026 HPtech Informática ME")
