import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import subprocess
import base64
import time
import importlib.util

# Descobre o caminho absoluto da pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --- CARREGAMENTO DINÂMICO DO MOTOR DE E-MAIL (Evita o SyntaxError com emojis) ---
caminho_pg2 = os.path.join(BASE_DIR, "pages", "02_📄_Relatorios.py")
try:
    spec = importlib.util.spec_from_file_location("relatorios_modulo", caminho_pg2)
    relatorios_modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(relatorios_modulo)
    enviar_relatorio_email = relatorios_modulo.enviar_relatorio_email
except Exception as e:
    st.error(f"Erro ao carregar o motor de e-mails da página 02: {e}")

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Termos HPTECH", 
    page_icon="hptech.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. CSS PARA OCULTAR O MENU E FORÇAR A LOGO NO TOPO
st.markdown("""
    <style>
        /* Esconde o menu de páginas padrão do Streamlit */
        [data-testid="stSidebarNav"] {display: none;}
        
        /* Zera o espaçamento do topo para a logo subir */
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* Ajuste da logo para não encostar nas laterais */
        [data-testid="stSidebarHeader"] {padding-top: 0rem !important;}
        
        /* Cor do título para o padrão azul CRTI */
        h1 { color: #b0231d; }
        
        /* Estilo da caixinha de usuário */
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -10px; /* Puxa a caixinha um pouco para cima */
        }
    </style>
""", unsafe_allow_html=True)

# 3. FUNÇÕES DE DADOS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

# 4. SIDEBAR COM MENU INTEGRADO E ENVIAR POR E-MAIL UNIFICADO
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
    
    # Navegação com rotas e emojis exatos do seu sistema
    if st.button("🏠 Home", use_container_width=True): st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True): st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📄 Relatórios", use_container_width=True): st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True): st.switch_page("pages/03_📝_Lancamento.py")
    if st.button("💰 Medição Mensal", use_container_width=True): st.switch_page("pages/04_💰_Medicao_Mensal.py")
    if st.button("📄 Termo Homologação", use_container_width=True): st.switch_page("pages/05_📄_Termos.py")
    if st.button("📄 Termo Encerramento", use_container_width=True): st.switch_page("pages/06_📑_Termo_Encerramento.py")
    
    # --- INTERFACE DE CONFIGURAÇÃO E DISPARO DE E-MAIL ISOLADA COM KEYS ÚNICAS ---
    st.markdown("---")
    st.header("📬 Disparo de Termos")
    email_destinatario = st.text_input("Enviar para (Destinatário):", value="financeiro@crti.com.br", key="termo_email_dest")
    senha_app = st.text_input("Senha App Gmail:", value="fzau tvih zlsn xadi", type="password", key="termo_email_pass")
    btn_enviar_emails = st.button("🚀 **ENVIAR TERMOS POR E-MAIL**", type="primary", use_container_width=True, key="termo_email_btn")
    
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# 5. CARREGAMENTO DE LISTAS
try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except Exception as e:
    st.error(f"Erro ao carregar os clientes da planilha: {e}")
    lista_clientes = []

# Título Customizado com Logo em Base64
def get_image_base64(path):
    with open(path, "rb") as img_file: return base64.b64encode(img_file.read()).decode()

try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(f'<div style="display: flex; align-items: center;"><h1 style="margin: 0; font-size: 2.5rem;">Emissão de Termos de Encerramento</h1><img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;"></div>', unsafe_allow_html=True)
except:
    st.title("📄 Emissão de Termos de Encerramento")

st.markdown("---")

# SELETOR DO DOCUMENTO
tipo_documento = st.selectbox(
    "Selecione o Tipo de Documento que deseja emitir:",
    ["Termo de Homologação e Encerramento Geral", "Documento de Não Homologação (Apenas Pendências)"]
)

st.markdown("---")

# Lista Completa de Módulos
TODOS_MODULOS = [
    "Compras", "Suprimentos e Estoque", "Frota - Equipamentos", 
    "Contratos e Medições de Terceiros", "Custos e Resultados", 
    "Apropriações e Apontamentos", "Produção", "Financeiro", 
    "Contábil", "Patrimonial", "Fiscal", "CRTI Buscador", 
    "CRTI Emissor NFe/NFCe", "CRTI Emissor CTe", "CRTI Emissor MDFe", 
    "CRTI Emissor NFSe", "CRTI Emissor Fatura de Locação", 
    "Gestão de Vendas (Produção)", "Gestão de Vendas (Agronegócio)", 
    "Engenharia, Contratos e Medições de Obras", "Locação de Equipamentos", 
    "Qualidade/Avaliação/Documentação", "Cadastros Globais", "Configuração do Sistema"
]

# --- ENTRADAS DO FORMULÁRIO ---
cliente_selecionado = st.selectbox("Nome do Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    if solicitantes:
        gerente_cliente_sugerido = str(solicitantes[0]).strip()
            
gerente_cliente = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
gerente_crti = "SUELLEN GOMES"

# --- INTERFACE DINÂMICA COMPATÍVEL ---
if tipo_documento == "Termo de Homologação e Encerramento Geral":
    col_datas_1, col_datas_2 = st.columns(2)
    with col_datas_1:
        data_inicio = st.date_input("Data de início da Implantação:", datetime.now())
    with col_datas_2:
        data_fim = st.date_input("Data da homologação da implantação:", datetime.now())

    st.markdown("---")
    st.subheader("Configuração dos Módulos")
    modulos_homologados = st.multiselect("Selecione os Módulos HOMOLOGADOS:", options=TODOS_MODULOS)

    dados_homologados_tabela = []
    if modulos_homologados:
        dt_virada_unica = st.date_input("Data de Início em Produção (Válida para todos os homologados):", datetime.now())
        data_virada_formatada = dt_virada_unica.strftime("%d/%m/%Y")
        for mod in modulos_homologados:
            dados_homologados_tabela.append({"nome": mod, "data": data_virada_formatada})

    opcoes_restantes = [m for m in TODOS_MODULOS if m not in modulos_homologados]
    modulos_nao_homologados = st.multiselect("Selecione os Módulos NÃO HOMOLOGADOS:", options=opcoes_restantes)

else:
    data_fim = st.date_input("Data do Documento Auxiliar:", datetime.now())
    st.markdown("---")
    modulos_nao_homologados = st.multiselect("Selecione os Módulos NÃO HOMOLOGADOS (Pendentes):", options=TODOS_MODULOS)

# Tradução da data por extenso
meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_fim.day} de {meses_br[data_fim.month - 1]} de {data_fim.year}"

# Pasta física temporária estável local para guardar cópias dos anexos
PASTA_TERMOS = "termos_emitidos"
os.makedirs(PASTA_TERMOS, exist_ok=True)

# --- 6. PROCESSAMENTO E EMISSÃO DO RELATÓRIO (SUA LÓGICA DO TERMO GERAL INTACTA) ---
if st.button("Gerar Documento Selecionado", type="primary"):
    if not cliente_selecionado:
        st.warning("Por favor, selecione um cliente para prosseguir.")
    elif tipo_documento == "Documento de Não Homologação (Apenas Pendências)" and not modulos_nao_homologados:
        st.warning("Por favor, selecione ao menos um módulo não homologado.")
    else:
        with st.spinner("⏳ Gerando termos customizados (Word e PDF)..."):
            try:
                if tipo_documento == "Termo de Homologação e Encerramento Geral":
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "encerramento.docx")
                else:
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "naohomologado.docx")
                
                doc = DocxTemplate(caminho_modelo)
                
                if modulos_nao_homologados:
                    texto_nao_homologados_str = "\n".join([f"• {mod}" for mod in modulos_nao_homologados])
                else:
                    texto_nao_homologados_str = "Nenhum módulo pendente nesta fase."

                if tipo_documento == "Termo de Homologação e Encerramento Geral":
                    # SUA LÓGICA VERTICAL DO TERMO GERAL ORIGINAL QUE DEU CERTO
                    nomes_homologados_str = "\n".join([str(item['nome']) for item in dados_homologados_tabela])
                    datas_homologados_str = "\n".join([str(item['data']) for item in dados_homologados_tabela])
                    
                    contexto = {
                        "cliente": cliente_selecionado,
                        "gerente_crti": gerente_crti,
                        "gerente_cliente": gerente_cliente,
                        "data_inicio": data_inicio.strftime("%d/%m/%Y"),
                        "data_fim": data_fim.strftime("%d/%m/%Y"),
                        "data_extenso": data_extenso_str,
                        
                        "nomes_homologados": nomes_homologados_str,
                        "datas_homologados": datas_homologados_str,
                        "texto_nao_homologados": texto_nao_homologados_str,
                        
                        " nomes_homologados ": nomes_homologados_str,
                        " datas_homologados ": datas_homologados_str,
                        " texto_nao_homologados ": texto_nao_homologados_str
                    }
                else:
                    contexto = {
                        "cliente": cliente_selecionado,
                        "gerente_cliente": gerente_cliente,
                        "data_extenso": data_extenso_str,
                        "texto_nao_homologados": texto_nao_homologados_str,
                        " texto_nao_homologados ": texto_nao_homologados_str
                    }
                
                doc.render(contexto)
                
                buffer_docx = io.BytesIO()
                doc.save(buffer_docx)
                buffer_docx.seek(0)
                
                prefixo = "Termo_Geral" if tipo_documento == "Termo de Homologação e Encerramento Geral" else "Doc_Não_Homologação"
                nome_download_bonito = f"{prefixo} - {cliente_selecionado}".replace("/", "-")
                
                caminho_pdf_fisico = os.path.join(PASTA_TERMOS, f"{nome_download_bonito}.pdf")
                caminho_docx_fisico = os.path.join(PASTA_TERMOS, f"{nome_download_bonito}.docx")
                
                doc.save(caminho_docx_fisico)
                
                cmd = f"libreoffice --headless --convert-to pdf --outdir {PASTA_TERMOS} {caminho_docx_fisico}"
                subprocess.run(cmd, shell=True, check=True)
                
                with open(caminho_pdf_fisico, "rb") as f:
                    buffer_pdf = io.BytesIO(f.read())
                
                st.success("✨ Documento gerado e salvo na base de envios com sucesso!")
                
                col_down1, col_down2 = st.columns(2)
                with col_down1:
                    st.download_button(label="📥 Baixar Termo em PDF (.pdf)", data=buffer_pdf, file_name=f"{nome_download_bonito}.pdf", mime="application/pdf", use_container_width=True)
                with col_down2:
                    st.download_button(label="📥 Baixar Termo em Word (.docx)", data=buffer_docx, file_name=f"{nome_download_bonito}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao processar o arquivo físico: {e}")

# =========================================================================
# 7. POP-UP DE CONFIRMAÇÃO DO DISPARO DE TERMOS
# =========================================================================
@st.dialog("📧 Confirmação de Disparo de Termos")
def confirmar_envio_termos_popup(arquivos_validos):
    st.write("Você tem certeza que deseja disparar o termo gerado por e-mail?")
    st.write(f"• **Destinatário:** `{email_destinatario}`")
    st.write(f"• **Arquivos em anexo:** PDF e Word do termo ativo")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("Sim, Disparar Termo", use_container_width=True):
            with st.spinner("Enviando e-mail..."):
                sucesso, msg = enviar_relatorio_email(
                    arquivos_validos, "://gmail.com", 587, "hudson.valente@crti.com.br", senha_app, email_destinatario
                )
                if sucesso:
                    st.success("🎉 Termo enviado com sucesso para a análise!")
                    st.balloons()
                    time.sleep(4)
                else:
                    st.error(msg)
                    time.sleep(4)
                st.rerun()
                
    with col_p2:
        if st.button("Não, Cancelar", use_container_width=True):
            st.rerun()

# --- GATILHO DA SIDEBAR QUE CHAMA O POP-UP ---
if btn_enviar_emails:
    import glob
    arquivos_pasta = glob.glob(os.path.join(PASTA_TERMOS, "*.*"))
    arquivos_validos = [f for f in arquivos_pasta if f.endswith(".pdf") or f.endswith(".xlsx") or f.endswith(".docx")]
    
    if not arquivos_validos:
        st.sidebar.warning("⚠️ Gere o documento na tela primeiro antes de disparar.")
    else:
        confirmar_envio_termos_popup(arquivos_validos)
