#Por Hudson Valente - HPTECH
#Criado em: 16/05/2026
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

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --- CARREGAMENTO DINÂMICO DO MOTOR DA PÁGINA 02 (Evita SyntaxError com emojis) ---
caminho_pg2 = os.path.join(BASE_DIR, "pages", "02_📄_Relatorios.py")
enviar_relatorio_email_sistema = None

try:
    if os.path.exists(caminho_pg2):
        spec = importlib.util.spec_from_file_location("relatorios_modulo_termos", caminho_pg2)
        relatorios_modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(relatorios_modulo)
        # Importa estritamente a função homologada que já funciona no seu app
        enviar_relatorio_email_sistema = relatorios_modulo.enviar_relatorio_email
except Exception as e:
    pass

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Termos HPTECH", page_icon="hptechICO.png", layout="wide")

# 2. CSS PARA OCULTAR O MENU E FORÇAR A LOGO NO TOPO
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        [data-testid="stSidebarHeader"] {padding-top: 0rem !important;}
        h1 { color: #b0231d; }
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -10px;
        }
    </style>
""", unsafe_allow_html=True)

# 3. SIDEBAR COM MENU INTEGRADO UNIFICADO
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    st.markdown("---")
    u_email = st.user.get("email") or "hudson.valente@crti.com.br"
    st.markdown(f"""
        <div class="user-block">
            <span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br>
            <span style='font-size: 11px; color: #555;'>{u_email}</span>
        </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.title("Menu Principal")
    
    if st.button("🏠 Home", use_container_width=True): st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True): st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📄 Relatórios", use_container_width=True): st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True): st.switch_page("pages/03_📝_Lancamento.py")
    if st.button("💰 Medição Mensal", use_container_width=True): st.switch_page("pages/04_💰_Medicao_Mensal.py")
    if st.button("📋 Termo Homologação", use_container_width=True): st.switch_page("pages/05_📋_Termos.py")
    if st.button("📑 Termo Encerramento", use_container_width=True): st.switch_page("pages/06_📑_Termo_Encerramento.py")
    
    # --- CONFIGURAÇÃO E DISPARO DE E-MAIL NA SIDEBAR ---
    st.markdown("---")
    st.header("📬 Disparo de Termos")
    email_destinatario = st.text_input("Enviar para (Destinatário):", value="financeiro@crti.com.br", key="t_email_dest")
    senha_app = st.text_input("Senha App Gmail:", value="fzau tvih zlsn xadi", type="password", key="t_email_pass")
    btn_enviar_emails = st.button("🚀 **ENVIAR TERMOS POR E-MAIL**", type="primary", use_container_width=True, key="t_email_btn")
    
    st.divider()
    st.caption("v1.0 - 11052026")

# 4. CARREGAMENTO DE LISTAS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except:
    df_leg = pd.DataFrame()
    lista_clientes = []

def get_image_base64(path):
    with open(path, "rb") as img_file: return base64.b64encode(img_file.read()).decode()

try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(f'<div style="display: flex; align-items: center;"><h1 style="margin: 0; font-size: 2.5rem;">Emissão de Termos de Encerramento</h1><img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;"></div>', unsafe_allow_html=True)
except:
    st.title("📄 Emissão de Termos de Encerramento")

st.markdown("---")

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

cliente_selecionado = st.selectbox("Nome do Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    if solicitantes: 
        gerente_cliente_sugerido = str(solicitantes[0]).strip()
        
gerente_cliente = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
gerente_crti = "SUELLEN GOMES"

tipo_documento = st.selectbox(
    "Selecione o Tipo de Documento que deseja emitir:",
    ["Termo de Homologação e Encerramento Geral", "Documento de Não Homologação (Apenas Pendências)"]
)

st.markdown("---")

# --- INTERFACE DINÂMICA ---
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

meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_fim.day} de {meses_br[data_fim.month - 1]} de {data_fim.year}"

# Pasta física temporária local para segurar as cópias dos arquivos antes de anexar
PASTA_TERMOS = os.path.join(BASE_DIR, "termos_emitidos")
os.makedirs(PASTA_TERMOS, exist_ok=True)

# --- GERAÇÃO DOS RELATÓRIOS ---
if st.button("Gerar Documento Selecionado", type="primary"):
    if not cliente_selecionado:
        st.warning("Por favor, selecione um cliente para prosseguir.")
    elif tipo_documento == "Documento de Não Homologação (Apenas Pendências)" and not modulos_nao_homologados:
        st.warning("Por favor, selecione ao menos um módulo não homologado.")
    else:
        with st.spinner("⏳ Gerando termos customizados (Word e PDF)..."):
            try:
                if tipo_documento == "Termo de Homologação e Encerramento Geral":
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "Lincoln_Pedro_Termos_encerramento.docx")
                    if not os.path.exists(caminho_modelo):
                        caminho_modelo = os.path.join(BASE_DIR, "modelos", "encerramento.docx")
                else:
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "naohomologado.docx")
                
                doc = DocxTemplate(caminho_modelo)
                
                if modulos_nao_homologados:
                    texto_nao_homologados_str = "\n".join([f"• {mod}" for mod in modulos_nao_homologados])
                else:
                    texto_nao_homologados_str = "Nenhum módulo pendente nesta fase."

                if tipo_documento == "Termo de Homologação e Encerramento Geral":
                    nomes_homologados_str = "\n".join([str(item['nome']) for item in dados_homologados_tabela])
                    datas_homologados_str = "\n".join([str(item['data']) for item in dados_homologados_tabela])
                    
                    contexto = {
                        "cliente": cliente_selecionado,
                        "gerente_crti": "SUELLEN GOMES",
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
                
                # Garante um nome limpo temporário para o Linux executar sem falhas de escape
                caminho_docx_fisico = os.path.join(PASTA_TERMOS, "temp_processamento.docx")
                doc.save(caminho_docx_fisico)
                
                cmd = f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TERMOS}" "{caminho_docx_fisico}"'
                subprocess.run(cmd, shell=True, check=True)
                
                caminho_pdf_gerado_lib = os.path.join(PASTA_TERMOS, "temp_processamento.pdf")
                
                prefixo = "Termo_Geral" if tipo_documento == "Termo de Homologação e Encerramento Geral" else "Doc_Não_Homologação"
                nome_download_bonito = f"{prefixo} - {cliente_selecionado}".replace("/", "-")
                
                caminho_pdf_final = os.path.join(PASTA_TERMOS, f"{nome_download_bonito}.pdf")
                caminho_docx_final = os.path.join(PASTA_TERMOS, f"{nome_download_bonito}.docx")
                
                if os.path.exists(caminho_pdf_final): os.remove(caminho_pdf_final)
                if os.path.exists(caminho_docx_final): os.remove(caminho_docx_final)
                os.rename(caminho_pdf_gerado_lib, caminho_pdf_final)
                os.rename(caminho_docx_fisico, caminho_docx_final)
                
                with open(caminho_pdf_final, "rb") as f:
                    buffer_pdf = io.BytesIO(f.read())
                    
                with open(caminho_docx_final, "rb") as f:
                    buffer_docx = io.BytesIO(f.read())
                
                st.success("✨ Documento gerado e pronto para envio!")
                
                col_down1, col_down2 = st.columns(2)
                with col_down1:
                    st.download_button(label="📥 Baixar Termo em PDF (.pdf)", data=buffer_pdf, file_name=f"{nome_download_bonito}.pdf", mime="application/pdf", use_container_width=True)
                with col_down2:
                    st.download_button(label="📥 Baixar Termo em Word (.docx)", data=buffer_docx, file_name=f"{nome_download_bonito}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao processar o arquivo físico: {e}")

# =========================================================================
# 5. POP-UP DE CONFIRMAÇÃO DO DISPARO DE TERMOS
# =========================================================================
@st.dialog("📧 Confirmação de Disparo de Termos")
def confirmar_envio_termos_popup_p4(email, arquivos_lote):
    st.write("Você tem certeza que deseja disparar o termo gerado por e-mail?")
    st.write(f"• **Destinatário:** `{email}`")
    st.write(f"• **Arquivos em anexo:** PDF e Word do termo ativo")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("Sim, Disparar Termo", use_container_width=True):
            if enviar_relatorio_email_sistema is not None:
                with st.spinner("Compilando anexos do termo e enviando pelo canal estável..."):
                    # Dispara chamando de forma transparente o mesmo barramento seguro da página 02
                    sucesso, msg = enviar_relatorio_email_sistema(
                        arquivos_lote, "://gmail.com", 587, "hudson.valente@crti.com.br", senha_app, email
                    )
                    if sucesso:
                        st.success("🎉 Termo enviado com sucesso para a análise!")
                        st.balloons()
                        time.sleep(4)
                    else:
                        st.error(msg)
                        time.sleep(4)
                    st.rerun()
            else:
                st.error("O motor de e-mails da página 02 não pôde ser instanciado.")
                
    with col_p2:
        if st.button("Não, Cancelar", use_container_width=True):
            st.rerun()

# --- GATILHO DA SIDEBAR QUE CHAMA O POP-UP ---
if btn_enviar_emails:
    import glob
    arquivos_pasta = glob.glob(os.path.join(PASTA_TERMOS, "*.pdf")) + glob.glob(os.path.join(PASTA_TERMOS, "*.docx"))
    
    if not arquivos_pasta:
        st.sidebar.warning("⚠️ Gere o documento na tela primeiro antes de disparar.")
    else:
        confirmar_envio_termos_popup_p4(email_destinatario, arquivos_pasta)
