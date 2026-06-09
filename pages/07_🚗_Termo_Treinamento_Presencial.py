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
    .user-block { background-color: #f0f2f6; padding: 8px; border-radius: 8px; margin-top: -10px; }
    </style>
""", unsafe_allow_html=True)

# 3. MOTOR DE DISPARO SMTP DA SUA PÁGINA 04
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
    msg["Subject"] = f"Termo de Confirmacao de Treinamento Presencial – {cliente_nome}"
    
    corpo_html = f"""
    <html>
    <body>
        <p>Prezados(as), espero que se encontre bem.</p>
        <p>Segue em anexo o <b>Termo de Confirmação de Treinamento Presencial</b> referente às visitas e consultorias realizadas no cliente <b>{cliente_nome}</b>.</p>
        <p>O documento detalha o histórico completo de dias, horários, escopos validados e observações técnicas por blocos de atendimento.</p>
        <p>Favor colher a assinatura institucional do Sr(a). {gerente_cliente_nome}.</p>
        <br>
        <p>Me coloco à inteira disposição para possíveis esclarecimentos.</p>
        <p>Atenciosamente,<br><b>Hudson Valente</b><br>HPtech Informática ME</p>
    </body>
    </html>
    """
    msg["To"] = ", ".join(lista_destinatarios)
    msg.attach(MIMEText(corpo_html, "html"))
    
    for caminho_arquivo in arquivos_anexos:
        nome_original = os.path.basename(caminho_arquivo)
        extensao = ".pdf" if nome_original.lower().endswith(".pdf") else ".docx"
        nome_arquivo_limpo = f"Termo_Confirmacao_Treinamento_Presencial_{cliente_nome}".replace(" ", "_") + extensao
            
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
# 4. SIDEBAR COM MENU INTEGRADO ATUALIZADO
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
    if st.button("🚗 Treinamento Presencial", use_container_width=True): st.switch_page("pages/07_🚗_Termo_Treinamento_Presencial.py")
    
    st.markdown("---")
    st.header("📬 Disparo de Termos")
    st.caption("Separe os e-mails usando vírgula (,)")
    email_destinatario = st.text_input("Enviar para (Destinatários):", value="financeiro@crti.com.br", key="pres_email_dest_key")
    btn_enviar_emails = st.button("🚀 **ENVIAR TERMO POR E-MAIL**", type="primary", key="pres_email_btn_key")
    
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

# 5. CARREGAMENTO DOS DADOS DA PLANILHA GOOGLE
@st.cache_data(ttl=300)
def carregar_dados_planilha():
    url_legendas = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df_leg = pd.read_excel(url_legendas, sheet_name="Legendas", engine='openpyxl')
    df_dados = pd.read_excel(url_legendas, sheet_name="Lancamentos", engine='openpyxl')
    return df_leg, df_dados

try:
    df_leg, df_dados = carregar_dados_planilha()
    lista_clientes = sorted(df_dados["CLIENTE"].dropna().unique().tolist())
except:
    df_leg = pd.DataFrame()
    df_dados = pd.DataFrame()
    lista_clientes = []

st.title("🚗 Confirmação de Treinamento Presencial (Por Blocos)")
st.write("O sistema extrai os dados e descrições diretamente da planilha de lançamentos.")
st.markdown("---")

cliente_selecionado = st.selectbox("Selecione o Cliente para Filtro Dinâmico:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    if solicitantes: 
        gerente_cliente_sugerido = str(solicitantes[0]).strip()

solicitante_nome = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
consultor_nome = st.text_input("Consultor Implantador (CRTI):", value="HUDSON VALENTE")
data_emissao = st.date_input("Data de Emissão do Termo:", datetime.now())

meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_emissao.day} de {meses_br[data_emissao.month - 1]} de {data_emissao.year}"
# --- 6. PROCESSAMENTO E FILTRAGEM REGRAS DE NEGÓCIO ---
if st.button("Gerar Relatório de Atendimentos Presenciais", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Selecione um cliente válido.")
    elif df_dados.empty:
        st.error("Erro: A base de dados da planilha não pôde ser lida.")
    else:
        with st.spinner("⏳ Filtrando lançamentos ativos em elaboração..."):
            try:
                df_dados["SITUAÇÃO"] = df_dados["SITUAÇÃO"].astype(str).str.strip()
                atendimentos_cliente = df_dados[
                    (df_dados["CLIENTE"] == cliente_selecionado) & 
                    (df_dados["SITUAÇÃO"] == "Em Elaboração") & 
                    (df_dados["RA"].notna())
                ].copy()
                
                if atendimentos_cliente.empty:
                    st.warning(f"Nenhum lançamento com RA válido e Situação 'Em Elaboração' foi localizado para o cliente '{cliente_selecionado}'.")
                else:
                    for antigo in glob.glob(os.path.join(PASTA_TREINAMENTO_P, "*.*")):
                        try: os.remove(antigo)
                        except: pass

                    atendimentos_cliente["DATA"] = pd.to_datetime(atendimentos_cliente["DATA"], errors='coerce')
                    atendimentos_cliente = atendimentos_cliente.sort_values(by="DATA")
                    
                    data_inicio_ra_str = atendimentos_cliente["DATA"].min().strftime("%d/%m/%Y")
                    data_fim_ra_str = atendimentos_cliente["DATA"].max().strftime("%d/%m/%Y")
                    periodo_visita_total = f"{data_inicio_ra_str} até {data_fim_ra_str}"

                    lista_atendimentos_word = []
                    lista_observacoes_gerais = []
                    
                    for idx, linha in sorted(atendimentos_cliente.iterrows(), key=lambda x: x[1]["DATA"]):
                        dt_str = linha["DATA"].strftime("%d/%m/%Y") if pd.notnull(linha["DATA"]) else ""
                        desc_pres_val = str(linha.get("DESCRIÇÃO ATENDIMENTO", "")).strip()
                        obs_pres_val = str(linha.get("OBSERVAÇÃO", "")).strip()
                        modulo_val = str(linha.get("MÓDULO / ATIVIDADE", "")).strip()
                        
                        lista_atendimentos_word.append({
                            "modulos": modulo_val,
                            "data_dia": dt_str,
                            "hora_inicio": str(linha.get("ENTRADA", "08:00")),
                            "hora_fim": str(linha.get("SAÍDA", "17:00")),
                            "desc_pres": desc_pres_val
                        })
                        
                        if obs_pres_val and obs_pres_val.lower() != "nan":
                            lista_observacoes_gerais.append(f"• Data {dt_str}: {obs_pres_val}")

                    resumao_geral_ac = "\n".join(lista_observacoes_gerais) if lista_observacoes_gerais else "Nenhuma observação técnica registrada."

                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "presencial.docx")
                    
                    if not os.path.exists(caminho_modelo):
                        st.error("⚠️ O modelo 'presencial.docx' não foi localizado.")
                    else:
                        doc = DocxTemplate(caminho_modelo)
                        contexto = {
                            "cliente": cliente_selecionado,
                            "consultor": consultor_nome,
                            "periodo_visita": periodo_visita_total,
                            "solicitante": solicitante_nome,
                            "data_extenso": data_extenso_str,
                            "atendimentos": lista_atendimentos_word,
                            "obs_geral_resumo": resumao_geral_ac
                        }
                        doc.render(contexto)
                        
                        nome_final = f"Termo_Treinamento_Presencial_{cliente_selecionado}".replace(" ", "_").replace("/", "-")
                        caminho_docx = os.path.join(PASTA_TREINAMENTO_P, f"{nome_final}.docx")
                        doc.save(caminho_docx)
                        
                        subprocess.run(f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TREINAMENTO_P}" "{caminho_docx}"', shell=True, check=True)
                        st.success("✨ Relatório gerado com sucesso!")
                        time.sleep(1)
                        st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar lote: {e}")

# --- 7. PAINEL VISUAL DE DOWNLOADS E ZIP ---
arquivos_gerados_p = glob.glob(os.path.join(PASTA_TREINAMENTO_P, "*.pdf")) + glob.glob(os.path.join(PASTA_TREINAMENTO_P, "*.docx"))

if arquivos_gerados_p:
    st.markdown("---")
    st.subheader("📥 Download do Histórico de Visitas Gerado")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for arq_caminho in arquivos_gerados_p:
            zip_file.write(arq_caminho, os.path.basename(arq_caminho))
    zip_buffer.seek(0)
    
    st.download_button(label="🎁 **BAIXAR PACOTE COMPLETO (ZIP)**", data=zip_buffer, file_name=f"Lote_Presencial_{cliente_selecionado}.zip", mime="application/zip", use_container_width=True)
    
    with st.expander("📄 Ver e baixar arquivos individuais...", expanded=True):
        col_grade = st.columns(2)
        for idx, arq_caminho in enumerate(sorted(arquivos_gerados_p)):
            nome_real = os.path.basename(arq_caminho).replace("_", " ")
            with open(arq_caminho, "rb") as f_leitura: conteudo_bytes = f_leitura.read()
            col_alvo = col_grade[idx % 2]
            extensao_label = "PDF" if arq_caminho.endswith(".pdf") else "Word"
            mime_tipo = "application/pdf" if arq_caminho.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            col_alvo.download_button(label=f"📥 Baixar {extensao_label}: {nome_real}", data=conteudo_bytes, file_name=nome_real, mime=mime_tipo, use_container_width=True, key=f"btn_dl_pres_v3_{idx}")

# =========================================================================
# Pop-up de Confirmação
# =========================================================================
@st.dialog(" Confirmação de Disparo de Termos")
def confirmar_envio_presencial_popup(email, arquivos_lote):
    st.write("Você tem certeza que deseja disparar o termo de treinamento presencial gerado por e-mail?")
    st.write(f"• **Destinatários:** `{email}`")
    st.write(f"• **Arquivos em anexo:** Lote contendo o Word e PDF com os blocos de visitas")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("Sim, Disparar Histórico", use_container_width=True):
            with st.spinner("Compilando anexos e enviando..."):
                ok, r_msg = enviar_email_treinamento_presencial(email, cliente_selecionado, arquivos_lote, solicitante_nome)
                if ok:
                    st.success(f" {r_msg}")
                    st.balloons()
                    time.sleep(4)
                else:
                    st.error(r_msg)
                    time.sleep(4)
                st.rerun()
    with col_p2:
        if st.button("Não, Cancelar", use_container_width=True): st.rerun()

if btn_enviar_emails:
    if not arquivos_gerados_p: st.sidebar.warning("⚠️ Mapeie os dados na tela primeiro antes de disparar.")
    else: confirmar_envio_presencial_popup(email_destinatario, arquivos_gerados_p)
