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
import requests
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
    msg["Subject"] = f"Termo de Confirmação de Treinamento Presencial – {cliente_nome}"
    
    corpo_html = f"""
    <html>
    <body>
        <p>Prezados(as), espero que se encontre bem.</p>
        <p>Segue em anexo o <b>Termo de Confirmação de Treinamento Presencial</b>, referente às visitas e consultorias realizadas no cliente <b>{cliente_nome}</b>.</p>
        <p>O documento detalha o histórico completo de visitas por atendimentos.</p>
        <p>Favor colher a assinatura institucional do Sr(a). {gerente_cliente_nome}.</p>
        <br>
        <p>Me coloco à inteira disposição para possíveis esclarecimentos.</p>
        <p>Atenciosamente,<br><b>Hudson Valente</b><br></p>
    </body>
    </html>
    """
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
       
    st.markdown("---")
    st.header("📬 Disparo de Termos")
    st.caption("Separe os e-mails usando vírgula (,)")
    email_destinatario = st.text_input("Enviar para (Destinatários):",  value="financeiro@crti.com.br,tayna@crti.com.br,kamille.voitach@crti.com.br,camille@crti.com.br", key="pres_email_dest_key")
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

    st.divider()
    st.caption("v1.0 - 08062026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# URL oficial mestre da planilha publicada
URL_PLANILHA_MUDANCA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"

# 5. CACHE DE ALTA VELOCIDADE EXCLUSIVO PARA CARREGAR A TELA INICIAL INSTANTANEAMENTE
@st.cache_data(ttl=300)
def carregar_dados():
    response = requests.get(URL_PLANILHA_MUDANCA, timeout=30, stream=True)
    xl = pd.ExcelFile(io.BytesIO(response.content))
    df_leg = pd.read_excel(xl, sheet_name="Legendas", engine='openpyxl')
    abas_reais = xl.sheet_names
    abas_meses = [a for a in abas_reais if a not in ["Legendas", "Config", "Dashboard", "Parâmetros", "Parametros"]]
    return df_leg, abas_meses
st.sidebar.header("⚙️ Configurações GERAIS")
if st.sidebar.button("🔄 Atualizar Base de Dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()    
try:
    df_leg, lista_abas_meses = carregar_dados()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except:
    df_leg = pd.DataFrame()
    lista_abas_meses = []
    lista_clientes = []

# Função para converter imagem local para Base64 (para funcionar dentro do HTML)
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Tenta carregar a imagem que está no repositório GitHub
try:
    img_base64 = get_image_base64("hptechICO.png")
    
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <h1 style="margin: 0; font-size: 2.5rem;">Confirmação de Treinamento Presencial</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.write("O sistema extrai os dados e descrições diretamente da base de lançamentos do mês selecionado.")
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("🚗 Confirmação de Treinamento Presencial")
#st.title("🚗 Confirmação de Treinamento Presencial")
#st.write("O sistema extrai os dados e descrições diretamente da base de lançamentos do mês selecionado.")
st.markdown("---")

cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")
aba_mes_selecionada = st.selectbox("Selecione o Mês do Atendimento:", lista_abas_meses) if lista_abas_meses else st.text_input("Mês:")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna()
    if not solicitantes_df.empty:
        # Extrai a string limpa do nome do gerente para sumir com aspas e colchetes
        gerente_cliente_sugerido = str(solicitantes_df.iloc[0]).strip()

solicitante_nome = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
consultor_nome = st.text_input("Consultor Implantador (CRTI):", value="HUDSON VALENTE")

data_emissao = st.date_input("Data de Emissão do Termo:", datetime.now())
meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_emissao.day} de {meses_br[data_emissao.month - 1]} de {data_emissao.year}"
# --- 6. PROCESSAMENTO E FILTRAGEM REGRAS DE NEGÓCIO ---
# --- 6. PROCESSAMENTO E FILTRAGEM DINÂMICA DA ABA SELECIONADA ---
# --- 6. PROCESSAMENTO E FILTRAGEM DINÂMICA DA ABA SELECIONADA ---

if st.button("Gerar Relatório de Atendimentos Presenciais", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Selecione um cliente válido.")
    else:
        with st.spinner(f"⏳ Processando atendimentos do mês: '{aba_mes_selecionada}'..."):
            try:
                res_dados = requests.get(URL_PLANILHA_MUDANCA, timeout=45, stream=True)
                xl_dados = pd.ExcelFile(io.BytesIO(res_dados.content))
                df_dados = pd.read_excel(xl_dados, sheet_name=aba_mes_selecionada, engine='openpyxl')
                
                df_dados = df_dados.reset_index(drop=True)
                df_dados.columns = df_dados.columns.str.strip()
                
                col_cliente = "CLIENTE"
                col_situacao = "SITUACAO_RA"
                col_ra = "RA"
                col_data = "DATA"
                
                if col_situacao in df_dados.columns:
                    df_dados[col_situacao] = df_dados[col_situacao].astype(str).str.strip()
                
                atendimentos_cliente = df_dados[
                    (df_dados[col_cliente] == cliente_selecionado) & 
                    (df_dados[col_situacao] == "Em Elaboração") & 
                    (df_dados[col_ra].notna())
                ].copy()
                
                if atendimentos_cliente.empty:
                    st.warning(f"⚠️ Nenhum lançamento com SITUACAO_RA = 'Em Elaboração' foi localizado na aba '{aba_mes_selecionada}'.")
                else:
                    for antigo in glob.glob(os.path.join(PASTA_TREINAMENTO_P, "*.*")):
                        try: os.remove(antigo)
                        except: pass

                    atendimentos_cliente = atendimentos_cliente.reset_index(drop=True)
                    atendimentos_cliente[col_data] = pd.to_datetime(atendimentos_cliente[col_data], errors='coerce')
                    atendimentos_cliente = atendimentos_cliente.sort_values(by=col_data)
                    
                    data_inicio_ra_str = atendimentos_cliente[col_data].min().strftime("%d/%m/%Y")
                    data_fim_ra_str = atendimentos_cliente[col_data].max().strftime("%d/%m/%Y")
                    periodo_visita_total = f"{data_inicio_ra_str} até {data_fim_ra_str}"

                    lista_atendimentos_word = []
                    lista_observacoes_gerais = []
                    
                    # Nome exato e estrito das colunas da sua planilha de dados
                    col_part = "PARTICIPANTE"
                    col_desc = "DESC_PRES"
                    col_obs = "OBS_PRES"
                    col_modulo = "MÓDULO / ATIVIDADE" if "MÓDULO / ATIVIDADE" in df_dados.columns else "MODULO / ATIVIDADE"
                    
                    # CORREÇÃO DEFINITIVA: Vincula as colunas de horário reais da planilha
                    col_hr_inicio = "HR_INICIO"
                    col_hr_fim = "HR_FIM"
                    
                    dict_part = atendimentos_cliente[col_part].to_dict() if col_part in atendimentos_cliente.columns else {}
                    dict_desc = atendimentos_cliente[col_desc].to_dict() if col_desc in atendimentos_cliente.columns else {}
                    dict_obs = atendimentos_cliente[col_obs].to_dict() if col_obs in atendimentos_cliente.columns else {}
                    dict_modulo = atendimentos_cliente[col_modulo].to_dict() if col_modulo in atendimentos_cliente.columns else {}
                    dict_data_dia = atendimentos_cliente[col_data].to_dict()
                    
                    # Extrai os dicionários usando os termos corretos passados por você
                    dict_hr_ini = atendimentos_cliente[col_hr_inicio].to_dict() if col_hr_inicio in atendimentos_cliente.columns else {}
                    dict_hr_fim = atendimentos_cliente[col_hr_fim].to_dict() if col_hr_fim in atendimentos_cliente.columns else {}

                    for idx in atendimentos_cliente.index:
                        dt_objeto = dict_data_dia.get(idx)
                        dt_str = dt_objeto.strftime("%d/%m/%Y") if pd.notnull(dt_objeto) else ""
                        
                        part_val = str(dict_part.get(idx, "")).strip()
                        desc_pres_val = str(dict_desc.get(idx, "")).strip()
                        obs_pres_val = str(dict_obs.get(idx, "")).strip()
                        modulo_val = str(dict_modulo.get(idx, "")).strip()
                        
                        # Captura os horários reais individuais da respectiva linha (manhã ou tarde)
                        hora_ini_raw = str(dict_hr_ini.get(idx, "08:00")).strip()
                        hora_fim_raw = str(dict_hr_fim.get(idx, "12:00")).strip()
                        
                        hora_ini_val = hora_ini_raw[:5] if ":" in hora_ini_raw else hora_ini_raw
                        hora_fim_val = hora_fim_raw[:5] if ":" in hora_fim_raw else hora_fim_raw
                        
                        participante_final = part_val
                        if not participante_final or participante_final.lower() in ["nan", "", "none"]:
                            participante_final = str(solicitante_nome).strip()
                        
                        lista_atendimentos_word.append({
                            "modulos": modulo_val if modulo_val.lower() != "nan" else "",
                            "modulo": modulo_val if modulo_val.lower() != "nan" else "",
                            "participantes": participante_final,
                            "participante": participante_final,
                            "data_dia": dt_str,
                            "data": dt_str,
                            "hora_inicio": hora_ini_val,
                            "hora_fim": hora_fim_val,
                            "desc_pres": desc_pres_val if desc_pres_val.lower() != "nan" else ""
                        })
                        
                        if obs_pres_val and obs_pres_val.lower() != "nan" and obs_pres_val.strip() != "":
                            lista_observacoes_gerais.append(obs_pres_val.strip())

                    resumao_geral_ac = "\n".join(lista_observacoes_gerais) if lista_observacoes_gerais else "Nenhuma observação técnica registrada."
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "presencial.docx")
                    
                    if not os.path.exists(caminho_modelo):
                        st.error("⚠️ O modelo 'presencial.docx' não foi localizado na pasta 'modelos'.")
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
                st.error(f"Erro ao processar lote na aba selecionada: {e}")

# --- PAINEL VISUAL DE DOWNLOADS E ZIP ---
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

# Pop-up de Confirmação
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
    if not arquivos_gerados_p: 
        st.sidebar.warning("⚠️ Mapeie os dados na tela primeiro antes de disparar.")
    else: 
        confirmar_envio_presencial_popup(email_destinatario, arquivos_gerados_p)
