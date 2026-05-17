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
import smtplib
import zipfile
import glob
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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

# 3. MOTOR DE DISPARO SMTP DA SUA PÁGINA 04 (BLINDADO CONTRA ERRO NONAME)
def enviar_email_termos_logica_p04(email_destino, nome_documento, arquivos_anexos):
    email_remetente = st.secrets["smtp"]["usuario"]
    senha_remetente = st.secrets["smtp"]["senha"]
    smtp_server = st.secrets["smtp"]["servidor"]
    smtp_porta = int(st.secrets["smtp"]["porta"])
    
    # Cria um assunto limpo removendo caracteres longos demais
    assunto_limpo = str(nome_documento).replace("_", " ").strip()
    
    msg = MIMEMultipart()
    msg["From"] = email_remetente
    msg["To"] = email_destino
    msg["Subject"] = f"{assunto_limpo} - HUDSON VALENTE"
    
    corpo = f"""<html><body><p>Prezada Sra. Amanda, espero que se encontre bem.</p><br>
    <p>Segue em anexo o arquivo de homologação <b>{assunto_limpo}</b> para análise e assinatura institucional.</p><br>
    <p>Atenciosamente,<br><br>Hudson Valente<br>HPtech Informática ME</p></body></html>"""
    msg.attach(MIMEText(corpo, "html"))
    
    for caminho_arquivo in arquivos_anexos:
        try:
            with open(caminho_arquivo, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encode_base64(part)
                
                # CORREÇÃO DEFINITIVA: Força um nome curto e limpo diretamente na extensão da aba do Gmail
                extensao = ".pdf" if caminho_arquivo.lower().endswith(".pdf") else ".docx"
                nome_anexo_institucional = f"Documento_Homologacao{extensao}"
                
                # Monta o cabeçalho padrão de forma estrita exigida pelo protocolo RFC do Gmail
                part.add_header("Content-Disposition", "attachment", filename=nome_anexo_institucional)
                msg.attach(part)
        except Exception as e:
            return False, f"Falha ao acoplar anexo: {str(e)}"
            
    try:
        server = smtplib.SMTP(smtp_server, smtp_porta)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, email_destino, msg.as_string())
        server.quit()
        return True, "E-mail enviado com sucesso!"
    except Exception as e:
        return False, f"Falha no envio SMTP (Segurança de Rede): {str(e)}"

# 4. SIDEBAR COM OS EMOJIS EXATOS DA SUA ÁRVORE DE ARQUIVOS
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
    
    # INTERFACE DE DISPARO DA SIDEBAR
    st.markdown("---")
    st.header("📬 Disparo de Termos")
    email_destinatario = st.text_input("Enviar para (Destinatário):", value="financeiro@crti.com.br", key="enc_email_dest_key")
    btn_enviar_emails = st.button("🚀 **ENVIAR TERMOS POR E-MAIL**", type="primary", use_container_width=True, key="enc_email_btn_key")
    
    # BOTAO AUXILIAR DE LIMPEZA NA SIDEBAR
    st.markdown("---")
    st.header("🗑️ Gerenciamento")
    if st.button("🗑️ Limpar Todos os Termos Emitidos", use_container_width=True):
        arquivos_limpeza = glob.glob(os.path.join(BASE_DIR, "termos_emitidos", "*.*"))
        for arq in arquivos_limpeza:
            try: os.remove(arq)
            except: pass
        st.success("Pasta de histórico esvaziada!")
        time.sleep(1.5)
        st.rerun()

    st.divider()
    st.caption("v1.0 - 11052026")

# 5. CARREGAMENTO DE LISTAS
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
        gerente_cliente_sugerido = str(solicitantes).strip()
        
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

PASTA_TERMOS = os.path.join(BASE_DIR, "termos_emitidos")
os.makedirs(PASTA_TERMOS, exist_ok=True)

# --- 6. GERAÇÃO DOS RELATÓRIOS ---
if st.button("Gerar Documento Selecionado", type="primary"):
    if not cliente_selecionado:
        st.warning("Por favor, selecione um cliente para prosseguir.")
    elif tipo_documento == "Documento de Não Homologação (Apenas Pendências)" and not modulos_nao_homologados:
        st.warning("Por favor, selecione ao menos um módulo não homologado.")
    else:
        with st.spinner("⏳ Gerando termos customizados (Word e PDF)..."):
            try:
                for arquivo_antigo in glob.glob(os.path.join(PASTA_TERMOS, "*.*")):
                    try: os.remove(arquivo_antigo)
                    except: pass

                if tipo_documento == "Termo de Homologação e Encerramento Geral":
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "Lincoln_Pedro_Termos_Campanha_encerramento.docx")
                    if not os.path.exists(caminho_modelo):
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
                
                prefixo = "Termo_Geral" if tipo_documento == "Termo de Homologação e Encerramento Geral" else "Doc_Não_Homologação"
                nome_limpo_arquivo = f"{prefixo}_{cliente_selecionado}".replace(" ", "_").replace("/", "-")
                
                caminho_docx_ativo = os.path.join(PASTA_TERMOS, f"{nome_limpo_arquivo}.docx")
                doc.save(caminho_docx_ativo)
                
                cmd = f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TERMOS}" "{caminho_docx_ativo}"'
                subprocess.run(cmd, shell=True, check=True)
                
                st.success("✨ Termo gerado com sucesso e armazenado na tela!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar o arquivo físico: {e}")

# --- PAINEL VISUAL DE DOWNLOAD E ZIP COMPACTADO ---
arquivos_gerados_base = glob.glob(os.path.join(PASTA_TERMOS, "*.pdf")) + glob.glob(os.path.join(PASTA_TERMOS, "*.docx"))

if arquivos_gerados_base:
    st.markdown("---")
    st.subheader("📥 Download dos Arquivos Emitidos")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for arq_caminho in arquivos_gerados_base:
            zip_file.write(arq_caminho, os.path.basename(arq_caminho))
    zip_buffer.seek(0)
    
    st.download_button(
        label="🎁 **BAIXAR TODOS OS TERMOS CONFIGURADOS (ZIP)**",
        data=zip_buffer,
        file_name=f"Termos_HPTECH_{cliente_selecionado}.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    with st.expander("📄 Ver arquivos individuais...", expanded=True):
        col_grade = st.columns(2)
        for idx, arq_caminho in enumerate(sorted(arquivos_gerados_base)):
            nome_real = os.path.basename(arq_caminho)
            with open(arq_caminho, "rb") as f_leitura:
                conteudo_bytes = f_leitura.read()
            
            col_alvo = col_grade[idx % 2]
            extensao_label = "PDF" if nome_real.endswith(".pdf") else "Word"
            mime_tipo = "application/pdf" if nome_real.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            
            col_alvo.download_button(
                label=f"📥 Baixar {extensao_label}: {nome_real}",
                data=conteudo_bytes,
                file_name=nome_real,
                mime=mime_tipo,
                use_container_width=True,
                key=f"btn_dl_{idx}"
            )

# =========================================================================
# 7. POP-UP DE CONFIRMAÇÃO DO DISPARO DE TERMOS
# =========================================================================
@st.dialog(" Confirmação de Disparo de Termos")
def confirmar_envio_termos_popup_final(email, arquivos_lote):
    st.write("Você tem certeza que deseja disparar o termo gerado por e-mail?")
    st.write(f"• **Destinatário:** `{email}`")
    st.write(f"• **Arquivos em anexo:** PDF e Word do lote ativo")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("Sim, Disparar Termo", use_container_width=True):
            with st.spinner("Compilando lote de anexos e enviando..."):
                if arquivos_lote:
                    # Resolve o bug do noname fixando o assunto de forma estruturada e limpa
                    primeiro_arq = arquivos_lote
                    nome_doc_assunto = os.path.basename(primeiro_arq).replace("_", " ").replace(".pdf", "").replace(".docx", "")
                else:
                    nome_doc_assunto = "Termo de Homologação"
                
                ok, r_msg = enviar_email_termos_logica_p04(email, nome_doc_assunto, arquivos_lote)
                
                if ok:
                    st.success(f" {r_msg}")
                    st.balloons()
                    time.sleep(4)
                else:
                    st.error(r_msg)
                    time.sleep(4)
                st.rerun()
                
    with col_p2:
        if st.button("Não, Cancelar", use_container_width=True):
            st.rerun()

# --- GATILHO DA SIDEBAR QUE CHAMA O POP-UP ---
if btn_enviar_emails:
    if not arquivos_gerados_base:
        st.sidebar.warning(" Gere o documento na tela primeiro antes de disparar.")
    else:
        confirmar_envio_termos_popup_final(email_destinatario, arquivos_gerados_base)
