#Por Hudson Valente - HPTECH
#Criado em: 16/05/2026
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import subprocess
import zipfile
import glob
import smtplib
import time
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Pasta mestre isolada para os termos de homologação emitidos nesta tela
PASTA_TERMOS_H = os.path.join(BASE_DIR, "termos_homologacao_emitidos")
os.makedirs(PASTA_TERMOS_H, exist_ok=True)

# 1. CONFIGURAÇÃO DA PÁGINA
#st.set_page_config(page_title="Gerador de Termos HPTECH", page_icon="hptech.png", layout="wide")

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Gerador de Termos HPTECH",
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

# 3. MOTOR DE DISPARO SMTP DA SUA PÁGINA 04
def enviar_email_homologacao_logica_p04(string_destinatarios, cliente_nome, arquivos_anexos, gerente_cliente_nome):
    email_remetente = st.secrets["smtp"]["usuario"]
    senha_remetente = st.secrets["smtp"]["senha"]
    smtp_server = st.secrets["smtp"]["servidor"]
    smtp_porta = int(st.secrets["smtp"]["porta"])
    
    lista_destinatarios = [email.strip() for email in string_destinatarios.split(",") if email.strip()]
    if not lista_destinatarios:
        return False, "Nenhum e-mail válido foi inserido."
        
    msg = MIMEMultipart()
    msg["From"] = email_remetente
    msg["To"] = ", ".join(lista_destinatarios)
    msg["Subject"] = f"TERMO DE HOMOLOGAÇÃO DE TODOS OS MÓDULOS CRTI ERP – {cliente_nome}"
    
    corpo_html = f"""
    <html>
    <body>
        <p>Prezada Sra. Amanda, espero que se encontre bem.</p>
        <p>Segue em anexo o pacote contendo os <b>Termos de Homologação</b> referentes aos módulos do sistema CRTI ERP implantado no cliente <b>{cliente_nome}</b>.</p>
        <p>Os documentos já foram validados e está pronto para análise e assinatura institucional do Sr(a). {gerente_cliente_nome}.</p>
        <br>
        <p>Me coloco à inteira disposição para possíveis esclarecimentos.</p>
        <p>Com Gratidão!!,<br><br>Hudson Valente</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(corpo_html, "html"))
    
    for caminho_arquivo in arquivos_anexos:
        nome_original = os.path.basename(caminho_arquivo)
        # Blindagem do Gmail contra o bug noname: mantém os nomes limpos de rede
        nome_arquivo_limpo = nome_original.replace(" ", "_")
            
        try:
            with open(caminho_arquivo, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encode_base64(part)
                part.add_header("Content-Disposition", "attachment", filename=str(nome_arquivo_limpo))
                msg.attach(part)
        except Exception as e:
            return False, f"Falha ao acoplar anexo: {str(e)}"
            
    try:
        server = smtplib.SMTP(smtp_server, smtp_porta)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, lista_destinatarios, msg.as_string())
        server.quit()
        return True, "E-mail com o lote enviado com sucesso para todos os destinatários!"
    except Exception as e:
        return False, f"Falha no envio SMTP (Segurança de Rede): {str(e)}"

# 4. SIDEBAR COM SEU MENU CUSTOMIZADO E DISPARO INTEGRADO
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
    st.caption("Separe os e-mails usando vírgula (,)")
    email_destinatario = st.text_input("Enviar para (Destinatários):", value="financeiro@crti.com.br", key="hom_email_dest_key")
    btn_enviar_emails = st.button("🚀 **ENVIAR PACOTE EM LOTE**", type="primary", use_container_width=True, key="hom_email_btn_key")
    
    # BOTAO AUXILIAR DE LIMPEZA NA SIDEBAR
    st.markdown("---")
    st.header("🗑️ Gerenciamento")
    if st.button("🗑️ Limpar Termos de Homologação", use_container_width=True):
        arquivos_limpeza = glob.glob(os.path.join(PASTA_TERMOS_H, "*.*"))
        for arq in arquivos_limpeza:
            try: os.remove(arq)
            except: pass
        st.success("Pasta de histórico esvaziada!")
        time.sleep(1.5)
        st.rerun()

    st.divider()
    st.caption("v1.1 - 17052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# 5. CARREGAMENTO DE LISTAS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except Exception as e:
    st.error(f"Erro ao carregar os clientes da planilha: {e}")
    lista_clientes = []

# Cabeçalho Principal Visual
def get_image_base64(path):
    with open(path, "rb") as img_file: return base64.b64encode(img_file.read()).decode()

try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(f'<div style="display: flex; align-items: center;"><h1 style="margin: 0; font-size: 2.5rem;">Gerador de Termos de Homologação</h1><img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 140px;"></div>', unsafe_allow_html=True)
    st.markdown("Selecione o cliente e os módulos para gerar os Termos de Homologação em lote.")
except:
    st.title("📋 Gerador de Termos de Homologação")

st.markdown("---")

# Lista de Módulos Unificada
TODOS_MODULOS = [
    "Compras", "Suprimentos e Estoque", "Frota - Equipamentos", 
    "Contratos e Medições de Terceiros", "Custos e Resultados", 
    "Financeiro", "Contábil", "Patrimonial", "Fiscal", "CRTI Buscador", 
    "CRTI Emissor NFe/NFCe", "CRTI Emissor CTe", "CRTI Emissor MDFe", 
    "CRTI Emissor NFSe", "CRTI Emissor Fatura de Locação", 
    "Gestão de Vendas (Produção)", "Gestão de Vendas (Agronegócio)", 
    "Engenharia, Contratos e Medições de Obras", "Locação de Equipamentos", 
    "Qualidade/Avaliação/Documentação"
]

modulos_selecionados = st.multiselect("Selecione os Módulos de Treinamento que deseja Homologar:", TODOS_MODULOS)

MAPA_MODELOS = {
    "Compras": "compras.docx", "Suprimentos e Estoque": "suprimentos.docx", "Frota - Equipamentos": "frotas.docx",
    "Contratos e Medições de Terceiros": "terceiros.docx", "Custos e Resultados": "custos.docx", "Financeiro": "financeiro.docx",
    "Contábil": "contabil.docx", "Patrimonial": "patrimonial.docx", "Fiscal": "fiscal.docx", "CRTI Buscador": "buscador.docx",
    "CRTI Emissor NFe/NFCe": "nfe.docx", "CRTI Emissor CTe": "cte.docx", "CRTI Emissor MDFe": "mdfe.docx",
    "CRTI Emissor NFSe": "nfse.docx", "CRTI Emissor Fatura de Locação": "fatura.docx", "Gestão de Vendas (Produção)": "vendas.docx",
    "Gestão de Vendas (Agronegócio)": "agronegocio.docx", "Engenharia, Contratos e Medições de Obras": "engenharia.docx",
    "Locação de Equipamentos": "locacao.docx", "Qualidade/Avaliação/Documentação": "qualidade.docx"
}

if lista_clientes:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes)
else:
    cliente_selecionado = st.text_input("Digite o Nome do Cliente manualmente:")

# Busca e higienização automática do gerente solicitante do Sheets
gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    if solicitantes: 
        gerente_cliente_sugerido = str(df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().iloc[0]).strip()
        
gerente_cliente_name = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)

data_selecionada = st.date_input("Data da Homologação:", datetime.now())
data_formatada = data_selecionada.strftime("%d/%m/%Y")

meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_selecionada.day} de {meses_br[data_selecionada.month - 1]} de {data_selecionada.year}"

# --- 6. GERAÇÃO EM LOTE DOS TERMOS DE HOMOLOGAÇÃO ---
if st.button("Gerar Todos os Termos Selecionados", type="primary", use_container_width=True):
    if not cliente_selecionado or cliente_selecionado == "Erro ao carregar":
        st.warning("Por favor, selecione um cliente válido.")
    elif not modulos_selecionados:
        st.warning("Por favor, selecione ao menos um módulo de treinamento.")
    else:
        with st.spinner("⏳ Compilando lote de termos selecionados (Word e PDF)..."):
            try:
                for arquivo_antigo in glob.glob(os.path.join(PASTER_TERMOS_H, "*.*") if 'PASTER_TERMOS_H' in locals() else os.path.join(PASTA_TERMOS_H, "*.*")):
                    try: os.remove(arquivo_antigo)
                    except: pass

                for modulo in modulos_selecionados:
                    modelo_nome_arquivo = MAPA_MODELOS.get(modulo, "compras.docx")
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", modelo_nome_arquivo)
                    
                    if not os.path.exists(caminho_modelo):
                        caminho_modelo = os.path.join(BASE_DIR, "modelos", f"{modulo.lower()}.docx")
                    if not os.path.exists(caminho_modelo):
                        caminho_modelo = os.path.join(BASE_DIR, "modelos", f"{modulo}.docx")

                    if os.path.exists(caminho_modelo):
                        doc = DocxTemplate(caminho_modelo)
                        contexto = {
                            "cliente": cliente_selecionado,
                            "data": data_formatada,
                            "data_homologacao": data_formatada,
                            "gerente_cliente": gerente_cliente_name,
                            "gerente_crti": "SUELLEN GOMES",
                            "data_extenso": data_extenso_str
                        }
                        doc.render(contexto)
                        
                        nome_mod_limpo = modulo.replace(" ", "_").replace("/", "-")
                        cliente_limpo = cliente_selecionado.replace(" ", "_").replace("/", "-")
                        nome_base = f"Termo_de_Homologacao_{nome_mod_limpo}_{cliente_limpo}"
                        
                        arquivo_docx_temporario = os.path.join(PASTA_TERMOS_H, f"{nome_base}.docx")
                        doc.save(arquivo_docx_temporario)
                        
                        cmd = f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TERMOS_H}" "{arquivo_docx_temporario}"'
                        subprocess.run(cmd, shell=True, check=True)

                st.success("✨ Todos os Termos de Homologação selecionados foram compilados com sucesso!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar e salvar o lote de Termos: {e}")

# --- 7. PAINEL VISUAL DE DOWNLOAD E ZIP COMPACTADO ---
arquivos_gerados_h = glob.glob(os.path.join(PASTA_TERMOS_H, "*.pdf")) + glob.glob(os.path.join(PASTA_TERMOS_H, "*.docx"))

if arquivos_gerados_h:
    st.markdown("---")
    st.subheader("📥 Download do Pacote de Homologação Emitido")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for arq_caminho in arquivos_gerados_h:
            zip_file.write(arq_caminho, os.path.basename(arq_caminho))
    zip_buffer.seek(0)
    
    st.download_button(
        label="🎁 **BAIXAR PACOTE DE HOMOLOGAÇÃO COMPLETO (ZIP)**",
        data=zip_buffer,
        file_name=f"Pacote_Homologacao_{cliente_selecionado}.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    with st.expander("📄 Ver e baixar arquivos individuais...", expanded=True):
        col_grade = st.columns(2)
        for idx, arq_caminho in enumerate(sorted(arquivos_gerados_h)):
            nome_real = os.path.basename(arq_caminho).replace("_", " ")
            with open(arq_caminho, "rb") as f_leitura: conteudo_bytes = f_leitura.read()
            
            col_alvo = col_grade[idx % 2]
            extensao_label = "PDF" if arq_caminho.endswith(".pdf") else "Word"
            mime_tipo = "application/pdf" if arq_caminho.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            
            col_alvo.download_button(
                label=f"📥 Baixar {extensao_label}: {nome_real}",
                data=conteudo_bytes,
                file_name=nome_real,
                mime=mime_tipo,
                use_container_width=True,
                key=f"btn_dl_hom_lote_{idx}"
            )

# =========================================================================
# 8. POP-UP DE CONFIRMAÇÃO DO DISPARO POR E-MAIL
# =========================================================================
@st.dialog(" Confirmação de Disparo de Termos")
def confirmar_envio_homologacao_popup(email, arquivos_lote):
    st.write("Você tem certeza que deseja disparar o lote de termos de homologação gerados por e-mail?")
    st.write(f"• **Destinatários:** `{email}`")
    st.write(f"• **Arquivos em anexo:** Todos os PDFs e Words dos módulos ativos selecionados")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("Sim, Disparar Homologações", use_container_width=True):
            with st.spinner("Conectando ao barramento SMTP seguro e enviando lote..."):
                # CORREÇÃO DEFINITIVA: Sincronizado removendo a variável individual inexistente
                ok, r_msg = enviar_email_homologacao_logica_p04(
                    email, cliente_selecionado, arquivos_lote, gerente_cliente_name
                )
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
    if not arquivos_gerados_h:
        st.sidebar.warning("⚠️ Gere os Termos na tela primeiro antes de disparar.")
    else:
        confirmar_envio_homologacao_popup(email_destinatario, arquivos_gerados_h)
