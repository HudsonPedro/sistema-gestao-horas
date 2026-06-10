#Por Hudson Valente - HPTECH
#Criado em: 10/06/2026
import io
import os 
import smtplib
import glob
import time
import requests
import base64
import pandas as pd
import streamlit as st
import xlsxwriter
from fpdf import FPDF
from datetime import datetime
from PIL import Image as PILImage
from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Pasta mestre isolada para os relatórios de reembolso emitidos
PASTA_REEMBOLSO_KM = os.path.join(BASE_DIR, "termos_reembolso_km")
os.makedirs(PASTA_REEMBOLSO_KM, exist_ok=True)

# 1. CONFIGURAÇÃO DA PÁGINA INSTITUCIONAL
st.set_page_config(page_title="Reembolso de KM - HPTECH", page_icon="hptech.png", layout="wide")

# 2. CSS PARA DESIGN AVANÇADO DA INTERFACE
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
    h1 { color: #b0231d; }
    .user-block { background-color: #f0f2f6; padding: 8px; border-radius: 8px; margin-top: -10px; }
    </style>
""", unsafe_allow_html=True)

def formatar_br(valor):
    try: 
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: 
        return "0,00"

# --- CLASSE PDF LIMPA SEM QUADROS (MODO PAISAGEM / LANDSCAPE) ---
class PDFReembolsoKM(FPDF):
    def header(self):
        pass
    def footer(self):
        pass
# 3. SIDEBAR COM MENU INTEGRADO E BOTÃO LIMPAR HISTÓRICO CORRIGIDO
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
    if st.button("💰 Reembolso de KM", use_container_width=True): st.switch_page("pages/08_💰_Reembolso_KM.py")
    
    st.markdown("---")
    st.header("🗑️ Gerenciamento")
    if st.button("🗑️ Limpar Histórico de KM", use_container_width=True):
        arquivos_limpeza = glob.glob(os.path.join(PASTA_REEMBOLSO_KM, "*.*"))
        for arq in arquivos_limpeza:
            try: os.remove(arq)
            except: pass
        if "km_pdf_p04" in st.session_state: del st.session_state["km_pdf_p04"]
        if "km_xlsx_p04" in st.session_state: del st.session_state["km_xlsx_p04"]
        st.success("Histórico de reembolsos esvaziado!")
        time.sleep(1.5)
        st.rerun()
    st.divider()
    st.caption("v1.0 - 10062026")

# URL mestre da planilha publicada
URL_PLANILHA_MUDANCA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
ENDERECO_CRTI_FIXO = "Rua Padre Anchieta, 2050 - Bigorrilho"

@st.cache_data(ttl=300)
def carregar_estrutura_abas_km():
    response = requests.get(URL_PLANILHA_MUDANCA, timeout=30, stream=True)
    xl = pd.ExcelFile(io.BytesIO(response.content))
    df_leg = pd.read_excel(xl, sheet_name="Legendas", engine='openpyxl')
    abas_reais = xl.sheet_names
    abas_meses = [a for a in abas_reais if a not in ["Legendas", "Config", "Dashboard", "Parâmetros", "Parametros"]]
    return df_leg, abas_meses

try:
    df_leg, lista_abas_meses = carregar_estrutura_abas_km()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except:
    df_leg = pd.DataFrame()
    lista_abas_meses = []
    lista_clientes = []

col_f1, col_f2 = st.columns(2)
with col_f1:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Cliente:")
    aba_mes_selecionada = st.selectbox("Selecione o Mês de Referência (Aba):", lista_abas_meses) if lista_abas_meses else st.text_input("Mês:")
with col_f2:
    vlr_abast = st.number_input("Valor Unitário Último Abastecimento (R$):", min_value=0.0, value=7.49, step=0.01)
    email_target = st.text_input("Destinatário do Relatório:", "suellen@crti.com.br")

st.markdown("### 📄 Comprovantes")
comprovante_file = st.file_uploader("Subir Comprovante de Abastecimento (Imagem PNG/JPG):", type=["png", "jpg", "jpeg"])
st.markdown("---")

endereco_cliente_sugerido = "Rua Pascoal Carignano, 675 - Ferraria, Campo Largo"
if not df_leg.empty and cliente_selecionado:
    try:
        solicitantes_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Endereco"].dropna()
        if not solicitantes_df.empty:
            texto_cru = str(solicitantes_df.to_list()).strip()
            endereco_cliente_sugerido = texto_cru.replace("['", "").replace("']", "").replace("[", "").replace("]", "")
    except: pass

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    try:
        sol_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna()
        if not sol_df.empty: 
            gerente_cliente_sugerido = str(sol_df.to_list()).strip().replace("['", "").replace("']", "").replace("[", "").replace("]", "")
    except: pass

solicitante_nome = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
data_emissao = st.date_input("Data de Emissão do Termo:", datetime.now())
meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_emissao.day} de {meses_br[data_emissao.month - 1]} de {data_emissao.year}"
# --- FUNÇÃO DE DISPARO SMTP PADRÃO DA PRODUÇÃO ---
def enviar_email_reembolso_km(email_destino, cliente, pdf_data, xlsx_data, n_pdf, n_xlsx):
    e_remetente = st.secrets["smtp"]["usuario"]
    senha_remetente = st.secrets["smtp"]["senha"]
    smtp_server = st.secrets["smtp"]["servidor"]
    smtp_porta = int(st.secrets["smtp"]["porta"])
    
    msg = MIMEMultipart()
    msg["From"] = e_remetente
    msg["To"] = email_destino
    msg["Subject"] = f"Relatorio Reembolso KM - {cliente}"
    
    corpo = f"<html><body><p>Prezada Suellen,</p><p>Segue em anexo o relatório consolidado de reembolso de KM rodado e a planilha Excel referente ao cliente <b>{cliente}</b>.</p><br><p>Atenciosamente,<br>Hudson Valente</p></body></html>"
    msg.attach(MIMEText(corpo, "html"))
    
    for b_data, nome_arquivo in [(pdf_data, n_pdf), (xlsx_data, n_xlsx)]:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(b_data)
        encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename=\"{nome_arquivo}\"")
        msg.attach(part)
        
    try:
        server = smtplib.SMTP(smtp_server, smtp_porta)
        server.starttls()
        server.login(e_remetente, senha_remetente)
        server.sendmail(e_remetente, email_destino, msg.as_string())
        server.quit()
        return True, "Relatório de KM enviado com sucesso!"
    except Exception as e:
        return False, f"Falha no envio: {str(e)}"

# --- 6. GERAÇÃO DO LAYOUT IDENTICO AO MODELO SOLICITADO ---
# --- 6. GERAÇÃO E PROCESSAMENTO EM CÓDIGO ---
if st.button("Gerar Relatório de Reembolso de KM", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Selecione um cliente válido.")
    else:
        with st.spinner("⏳ Compilando dados de percurso dinâmico..."):
            try:
                res_dados = requests.get(URL_PLANILHA_MUDANCA, timeout=45, stream=True)
                xl_dados = pd.ExcelFile(io.BytesIO(res_dados.content))
                df_dados = pd.read_excel(xl_dados, sheet_name=aba_mes_selecionada, engine='openpyxl')
                
                df_dados = df_dados.reset_index(drop=True)
                df_dados.columns = df_dados.columns.str.strip()
                
                col_cliente = "CLIENTE"
                col_situacao = "SITUACAO_RA"
                col_ra = "RA"
                col_km_d = "KM_D"
                col_data = "DATA"
                col_local_planilha = "LOCAL"
                
                if col_situacao in df_dados.columns:
                    df_dados[col_situacao] = df_dados[col_situacao].astype(str).str.strip()
                
                # Filtra os lançamentos válidos com KM_D maior que zero e situação em Elaboração
                atendimentos_km = df_dados[
                    (df_dados[col_cliente] == cliente_selecionado) & 
                    (df_dados[col_situacao] == "Em Elaboração") & 
                    (df_dados[col_ra].notna()) & 
                    (df_dados[col_km_d].notna()) & (df_dados[col_km_d] > 0)
                ].copy()
                
                if atendimentos_km.empty:
                    st.warning(f"⚠️ Nenhum lançamento ativo em elaboração com KM_D preenchido foi localizado para o cliente '{cliente_selecionado}'.")
                else:
                    atendimentos_km[col_data] = pd.to_datetime(atendimentos_km[col_data], errors='coerce')
                    atendimentos_km = atendimentos_km.sort_values(by=col_data).reset_index(drop=True)
                    
                    lista_linhas = []
                    t_km = 0.0
                    t_vlr = 0.0
                    
                    dict_km = atendimentos_km[col_km_d].to_dict()
                    dict_dt = atendimentos_km[col_data].to_dict()
                    dict_loc = atendimentos_km[col_local_planilha].to_dict() if col_local_planilha in atendimentos_km.columns else {}
                    
                    for idx in atendimentos_km.index:
                        dt_obj = dict_dt.get(idx)
                        dt_str = dt_obj.strftime("%d/%m/%Y") if pd.notnull(dt_obj) else ""
                        km_f = float(dict_km.get(idx, 0))
                        loc_celula = str(dict_loc.get(idx, "")).strip().lower()
                        
                        # REQUISITO EXATO: Decide e inverte a rota com base na coluna LOCAL da planilha mestre
                        if "cliente" in loc_celula:
                            percurso = "Ida"
                            orig = ENDERECO_CRTI_FIXO
                            dest = endereco_cliente_sugerido
                        else:
                            percurso = "Volta"
                            orig = endereco_cliente_sugerido
                            dest = ENDERECO_CRTI_FIXO
                            
                        vlr_f = km_f * (vlr_abast * 0.25)
                        t_km += km_f
                        t_vlr += vlr_f
                        
                        lista_linhas.append([dt_str, cliente_selecionado, orig, dest, percurso, f"{km_f:.0f}", f"R$ {formatar_br(vlr_abast)}", f"R$ {formatar_br(vlr_f)}"])

                    # 1. ARQUIVO EXCEL ESPELHADO (.XLSX)
                    df_excel = pd.DataFrame(lista_linhas)
                    df_excel.columns = ["DATA", "CLIENTE", "LOCAL ORIGEM", "LOCAL DESTINO", "Percurso", "DISTÂNCIA (KM)", "Vlr Unit. Ultimo Abast.", "TOTAL"]
                    df_excel.loc[len(df_excel)] = ["Total KM", f"{t_km:.0f}", "Valor Total", f"R$ {formatar_br(t_vlr)}", "", "", "", ""]
                    
                    buffer_xlsx = io.BytesIO()
                    with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                        df_excel.to_excel(writer, sheet_name="Reembolso", index=False)
                    st.session_state["km_xlsx_p04"] = buffer_xlsx.getvalue()

                    # 2. DOCUMENTO PDF IDÊNTICO (MODO PAISAGEM / LANDSCAPE)
                    pdf = PDFReembolsoKM(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    
                    ARQUIVO_LOGO = "crti.jpg"
                    if os.path.exists(ARQUIVO_LOGO):
                        pdf.image(ARQUIVO_LOGO, x=15, y=10, w=35)
                    
                    pdf.set_font("Arial", "B", 12)
                    pdf.text(120, 16, "RELATÓRIO PARA REEMBOLSO DE KM RODADO")
                    
                    pdf.set_font("Arial", "B", 9)
                    pdf.text(15, 32, "IMPLANTADOR CRTI: HUDSON PEDRO SALES VALENTE")
                    
                    pdf.set_y(36); pdf.set_x(15)
                    pdf.set_fill_color(226, 239, 218); pdf.set_font("Arial", "B", 7.5)
                    
                    headers = ["DATA", "CLIENTE", "LOCAL ORIGEM", "LOCAL DESTINO", "Percurso", "DISTÂNCIA (KM)", "Vlr Unit. Ultimo Abast.", "TOTAL"]
                    h_widths = [16, 28, 63, 63, 14, 21, 23, 19]
                    
                    for idx_h, txt in enumerate(headers): 
                        pdf.cell(h_widths[idx_h], 6, txt, border=1, align="C", fill=True)
                    
                    pdf.set_font("Arial", "", 6.5)
                    for row in lista_linhas:
                        pdf.ln(6); pdf.set_x(15)
                        for col_i, val in enumerate(row):
                            align_cell = "L" if col_i in [2, 3] else "C"
                            pdf.cell(h_widths[col_i], 6, str(val), border=1, align=align_cell)
                        
                    # Rodapé de totalização consolidado deslocado para as posições corretas
                    pdf.ln(6); pdf.set_x(15); pdf.set_font("Arial", "B", 7.5)
                    pdf.cell(16 + 28 + 63 + 63, 6, "", border=0)
                    pdf.cell(14, 6, "Total KM", border=1, align="R", fill=True)
                    pdf.cell(21, 6, f"{t_km:.0f}", border=1, align="C")
                    pdf.cell(23, 6, "Valor Total", border=1, align="R", fill=True)
                    pdf.cell(19, 6, f"R$ {formatar_br(t_vlr)}", border=1, align="C")
                    
                    # Página de Anexo exclusiva no PDF caso tenha comprovante enviado
                    caminho_imagem_disco = os.path.join(PASTA_REEMBOLSO_KM, "comprovante_km_final.jpg")
                    if comprovante_file:
                        img_pil_conv = PILImage.open(io.BytesIO(comprovante_file.read()))
                        img_pil_conv.convert("RGB").save(caminho_imagem_disco, "JPEG")
                        
                        pdf.add_page()
                        pdf.set_font("Arial", "B", 11)
                        pdf.text(15, 15, "COMPROVANTE DE ABASTECIMENTO ANEXADO")
                        pdf.image(caminho_imagem_disco, x=15, y=22, w=110)
                    
                    st.session_state["km_pdf_p04"] = pdf.output(dest="S").encode("latin1")
                    
                    if comprovante_file and os.path.exists(caminho_imagem_disco):
                        try: os.remove(caminho_imagem_disco)
                        except: pass
                        
                    st.success("✨ Relatórios gerados idênticos com sucesso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro na compilação: {e}")

# --- PAINEL VISUAL DE DOWNLOADS DA SESSÃO ---
if "km_pdf_p04" in st.session_state:
    st.markdown("---")
    n_pdf = f"Reembolso_KM_{cliente_selecionado}_{aba_mes_selecionada}.pdf"
    n_xlsx = f"Reembolso_KM_{cliente_selecionado}_{aba_mes_selecionada}.xlsx"
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(label="📥 **BAIXAR RELATÓRIO PDF**", data=st.session_state["km_pdf_p04"], file_name=n_pdf, mime="application/pdf", use_container_width=True)
    with col_dl2:
        st.download_button(label="📥 **BAIXAR PLANILHA EXCEL**", data=st.session_state["km_xlsx_p04"], file_name=n_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.markdown("---")

@st.dialog("Confirmação de Envio por E-mail")
def confirmar_envio_km_popup():
    st.write(f"Deseja disparar o relatório de KM do cliente **{cliente_selecionado}** para `{email_target}`?")
    st.markdown("---")
    cp1, cp2 = st.columns(2)
    with cp1:
        if st.button("Sim, Enviar", use_container_width=True):
            with st.spinner("Disparando e-mail corporativo..."):
                if "km_pdf_p04" in st.session_state:
                    ok, msg = enviar_email_reembolso_km(email_target, cliente_selecionado, st.session_state["km_pdf_p04"], st.session_state["km_xlsx_p04"], f"Reembolso_KM_{cliente_selecionado}.pdf", f"Reembolso_KM_{cliente_selecionado}.xlsx")
                    if ok:
                        st.success(msg)
                        st.balloons()
                        time.sleep(4)
                    else: st.error(msg)
                else:
                    st.error("Gere o relatório na tela primeiro antes de disparar.")
            st.rerun()
        with cp2:
            if st.button("Não, Cancelar", use_container_width=True): 
                st.rerun()

if st.button("🚀 **ENVIAR REEMBOLSO POR E-MAIL**", type="primary", use_container_width=True):
    if not email_target: 
        st.error("Insira um endereço de e-mail válido.")
    else: 
        confirmar_envio_km_popup()
