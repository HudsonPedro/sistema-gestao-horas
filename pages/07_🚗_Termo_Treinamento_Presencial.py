#Por Hudson Valente - HPTECH
#Criado em: 08/06/2026
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

# Pasta mestre isolada para os termos de treinamento presencial emitidos
PASTA_TREINAMENTO_P = os.path.join(BASE_DIR, "termos_treinamento_presencial")
os.makedirs(PASTA_TREINAMENTO_P, exist_ok=True)

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Treinamento Presencial HPTECH", page_icon="hptech.png", layout="wide")

# 2. CSS PARA OCULTAR O MENU PADRÃO E APLICAR SEU DESIGN
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
    h1 { color: #b0231d; }
    .user-block {
    background-color: #f0f2f6;
    padding: 8px;
    border-radius: 8px;
    margin-top: -10px;
    }
    </style>
""", unsafe_allow_html=True)

# 3. MOTOR DE DISPARO SMTP DA SUA PÁGINA 04 (ADAPTADO PARA SUPORTE MULTI-DESTINATÁRIO)
def enviar_email_treinamento_presencial(string_destinatarios, cliente_nome, arquivos_anexos, gerente_cliente_nome):
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
    msg["Subject"] = f"Termo de Confirmação de Treinamento Presencial – {cliente_nome}"
    
    corpo_html = f"""
    <html>
    <body>
        <p>Prezados(as), espero que se encontre bem.</p>
        <p>Segue em anexo o <b>Termo de Confirmação de Treinamento Presencial</b> referente às visitas realizadas no cliente <b>{cliente_nome}</b>.</p>
        <p>Os documentos detalham as agendas, módulos validados e os responsáveis participantes orientados.</p>
        <p>Favor colher a assinatura institucional do Sr(a). {gerente_cliente_nome}.</p>
        <br>
        <p>Me coloco à inteira disposição para possíveis esclarecimentos.</p>
        <p>Atenciosamente,<br><b>Hudson Valente</b><br>HPtech Informática ME</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(corpo_html, "html"))
    
    for caminho_arquivo in arquivos_anexos:
        nome_original = os.path.basename(caminho_arquivo)
        extensao = ".pdf" if nome_original.lower().endswith(".pdf") else ".docx"
        
        # Blindagem do Gmail contra o bug noname: higieniza cabeçalhos removendo acentos no tráfego MIME
        if "Presencial" in nome_original:
            nome_arquivo_limpo = f"Termo_Confirmacao_Treinamento_Presencial{extensao}"
        else:
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
        return True, "Pacote de treinamento presencial enviado com sucesso!"
    except Exception as e:
        return False, f"Falha no envio SMTP (Segurança de Rede): {str(e)}"

# 4. SIDEBAR COM MENU INTEGRADO ATUALIZADO (INCLUINDO A NOVA PÁGINA 07)
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
    email_destinatario = st.text_input("Enviar para (Destinatários):", value="financeiro@crti.com.br", key="pres_email_dest_key")
    btn_enviar_emails = st.button("🚀 **ENVIAR TERMO POR E-MAIL**", type="primary", use_container_width=True, key="pres_email_btn_key")
    
    # BOTAO AUXILIAR DE LIMPEZA NA SIDEBAR
    st.markdown("---")
    st.header("🗑️ Gerenciamento")
    if st.button("🗑️ Limpar Termos Presenciais", use_container_width=True):
        arquivos_limpeza = glob.glob(os.path.join(PASTA_TREINAMENTO_P, "*.*"))
        for arq in arquivos_limpeza:
            try: os.remove(arq)
            except: pass
        st.success("Pasta de histórico esvaziada!")
        time.sleep(1.5)
        st.rerun()

    st.divider()
    st.caption("v1.0 - 08062026")

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

# Título Principal da Tela
st.title("🚗 Confirmação de Treinamento Presencial")
st.write("Insira os parâmetros de visita para compilar os termos institucionais.")
st.markdown("---")

# --- FORMULÁRIO DE ENTRADAS BASEADO NAS TAGS DO DOCUMENTO ---
cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")

# Coleta dinâmica do solicitante (gerente do cliente)
gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado["Solicitante1"].dropna().unique().tolist()
    if solicitantes: 
        gerente_cliente_sugerido = str(solicitantes).strip()

col_dados_1, col_dados_2 = st.columns(2)
with col_dados_1:
    solicitante_nome = st.text_input("Nome do Responsável / Administrador (Cliente):", value=gerente_cliente_sugerido)
    participantes_txt = st.text_area("Responsáveis pelo recebimento dos treinamentos (Participantes):", placeholder="Ex:\nJoão Silva - Compras\nMaria Souza - Financeiro")
with col_dados_2:
    consultor_nome = st.text_input("Consultor Implantador (CRTI):", value="HUDSON VALENTE")
    modulos_treinados = st.multiselect("Módulos treinados presencialmente:", options=TODOS_MODULOS)

st.markdown("##### Agenda e Horários da Visita Presencial")
col_t1, col_t2, col_t3, col_t4 = st.columns(4)
with col_t1:
    data_inicio_v = st.date_input("Período da visita (De):", datetime.now())
with col_t2:
    data_fim_v = st.date_input("Período da visita (Até):", datetime.now())
with col_t3:
    hora_inicio_txt = st.text_input("Horário de Início:", value="08:00")
with col_t4:
    hora_fim_txt = st.text_input("Horário de Término:", value="17:00")

# Campos das Colunas AB e AC informadas no layout mestre
st.markdown("##### Detalhamento Técnico")
desc_presencial = st.text_area("Descrição das atividades realizadas (Coluna AB):", placeholder="Descreva os processos validados, alinhamento de custos, etc.")
obs_presencial = st.text_area("Observações adicionais (Coluna AC):", placeholder="Insira observações relevantes ou pendências mapeadas nesta visita.")

# Formatações das strings de data por extenso e período
periodo_completo_str = f"{data_inicio_v.strftime('%d/%m/%Y')} até {data_fim_v.strftime('%d/%m/%Y')}"
data_emissao_str = data_fim_v.strftime("%d/%m/%Y")
# --- 6. PROCESSAMENTO E VALIDAÇÃO DA GRAVAÇÃO ---
if st.button("Gerar Termo de Treinamento Presencial", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Por favor, identifique a Empresa Cliente.")
    elif not modulos_treinados:
        st.warning("Selecione ao menos um módulo treinado.")
    else:
        with st.spinner("⏳ Processando modelo 'presencial.docx' e convertendo..."):
            try:
                # Esvazia a pasta antes da nova compilação para não acumular lixo
                for arquivo_antigo in glob.glob(os.path.join(PASTA_TREINAMENTO_P, ".")):
                    try:
                        os.remove(arquivo_antigo)
                    except:
                        pass

                # Procura pelo arquivo de modelo padrão na pasta modelos
                caminho_modelo = os.path.join(BASE_DIR, "modelos", "presencial.docx")
                if not os.path.exists(caminho_modelo):
                    st.error("⚠️ Erro: O arquivo mestre 'presencial.docx' não foi localizado na sua pasta 'modelos' do GitHub. Faça o upload do arquivo para prosseguir.")
                else:
                    doc = DocxTemplate(caminho_modelo)
                    modulos_formatados_txt = "\n".join((f"• {m}" for m in modulos_treinados))

                    # Alinha rigorosamente todas as tags mapeadas na Page 1 do seu PDF
                    contexto = {
                        "cliente": cliente_selecionado,
                        "participantes": participantes_txt,
                        "consultor": consultor_nome,
                        "data": data_emissao_str,
                        "hora_inicio": hora_inicio_txt,
                        "hora_fim": hora_fim_txt,
                        "modulos": modulos_formatados_txt,
                        "desc_pres": desc_presencial,
                        "obs_pres": obs_presencial,
                        "solicitante": solicitante_nome
                    }

                    doc.render(contexto)

                    nome_base_arquivo = f"Termo_Treinamento_Presencial_{cliente_selecionado}".replace(" ", "_").replace("/", "-")
                    caminho_docx_final = os.path.join(PASTA_TREINAMENTO_P, f"{nome_base_arquivo}.docx")
                    doc.save(caminho_docx_final)

                    # Converte via LibreOffice headless para PDF
                    cmd = f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TREINAMENTO_P}" "{caminho_docx_final}"'
                    subprocess.run(cmd, shell=True, check=True)

                    st.success("✨ Lote de treinamento presencial gerado e disponível na tela!")
                    time.sleep(1)
                    st.rerun()

            except Exception as e:
                st.error(f"Erro ao compilar documento físico: {e}")
# --- 7. PAINEL VISUAL DE DOWNLOADS E ZIP COMPACTADO ---
arquivos_gerados_p = glob.glob(os.path.join(PASTA_TREINAMENTO_P, ".pdf")) + \
                     glob.glob(os.path.join(PASTA_TREINAMENTO_P, ".docx"))

if arquivos_gerados_p:
    st.markdown("---")
    st.subheader("📥 Download dos Documentos de Visita Presencial")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for arq_caminho in arquivos_gerados_p:
            zip_file.write(arq_caminho, os.path.basename(arq_caminho))

    zip_buffer.seek(0)
    st.download_button(
        label="🎁 BAIXAR TODOS OS DOCUMENTOS DE VISITA (ZIP)",
        data=zip_buffer,
        file_name=f"Lote_Treinamento_Presencial_{cliente_selecionado}.zip",
        mime="application/zip",
        use_container_width=True
    )

    with st.expander("📄 Ver e baixar arquivos individuais...", expanded=True):
        col_grade = st.columns(2)
        for idx, arq_caminho in enumerate(sorted(arquivos_gerados_p)):
            nome_real = os.path.basename(arq_caminho).replace("_", " ")
            with open(arq_caminho, "rb") as f_leitura:
                conteudo_bytes = f_leitura.read()

            col_alvo = col_grade[idx % 2]
            extensao_label = "PDF" if arq_caminho.endswith(".pdf") else "Word"
            mime_tipo = "application/pdf" if arq_caminho.endswith(".pdf") else \
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

            col_alvo.download_button(
                label=f"📥 Baixar {extensao_label}: {nome_real}",
                data=conteudo_bytes,
                file_name=nome_real,
                mime=mime_tipo,
                use_container_width=True,
                key=f"btn_dl_pres_{idx}"
            )

# =========================================================================
# 8. POP-UP DE CONFIRMAÇÃO DO DISPARO POR E-MAIL
# =========================================================================
@st.dialog(" Confirmação de Disparo de Termos")
def confirmar_envio_presencial_popup(email, arquivos_lote):
    st.write("Você tem certeza que deseja disparar o termo de treinamento presencial gerado?")
    st.write(f"• Destinatários: {email}")
    st.write("• Arquivos em anexo: Lote contendo o Word e PDF ativos mapeados")
    st.markdown("---")

    col_p1, col_p2 = st.columns(2)

    with col_p1:
        if st.button("Sim, Disparar Documentos", use_container_width=True):
            with st.spinner("Compilando anexos limpos e transmitindo por canal SMTP..."):
                ok, r_msg = enviar_email_treinamento_presencial(
                    email,
                    cliente_selecionado,
                    arquivos_lote,
                    solicitante_nome
                )
                if ok:
                    st.success(f"{r_msg}")
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
    if not archivos_gerados_p:
        st.sidebar.warning("⚠️ Gere o documento na tela primeiro antes de disparar.")
    else:
        confirmar_envio_presencial_popup(email_destinatario, archivos_gerados_p)
