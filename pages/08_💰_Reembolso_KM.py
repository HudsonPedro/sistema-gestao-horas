#Por Hudson Valente - HPTECH
#Criado em: 09/06/2026
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

# 2. CSS PARA OCULTAR O MENU PADRÃO E DEIXAR A TELA BONITA
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
    h1 { color: #b0231d; }
    .user-block { background-color: #f0f2f6; padding: 8px; border-radius: 8px; margin-top: -10px; }
    </style>
""", unsafe_allow_html=True)

# Função mestre de formatação para Real Brasileiro legítimo (R$ 9,99)
def formatar_br(valor):
    try:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"

# --- CLASSE PDF LIMPA (PADRÃO REQUISITO SOLICITADO) ---
class PDFReembolsoKM(FPDF):
    def header(self):
        pass
    def footer(self):
        pass
# 3. SIDEBAR COM MENU INTEGRADO E BOTÃO LIMPAR HISTÓRICO OPERACIONAL
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
    st.caption("v1.0 - 09062026")

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

st.title("💰 Relatório para Reembolso de KM Rodado")
st.write("Layout estruturado via código interno, gerando documentos XLSX e PDF idênticos de forma simultânea.")
st.markdown("---")

col_f1, col_f2 = st.columns(2)
with col_f1:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Cliente:")
    aba_mes_selecionada = st.selectbox("Selecione o Mês de Referência (Aba):", lista_abas_meses) if lista_abas_meses else st.text_input("Mês:")
with col_f2:
    vlr_abast = st.number_input("Valor Unitário Último Abastecimento (R$):", min_value=0.0, value=7.49, step=0.01)
    email_target = st.text_input("Destinatário do Relatório:", "hudson.valente@crti.com.br")

# INTERFACE VISUAL RESTAURADA: Área bonita para upload do cupom de abastecimento
st.markdown("### 📄 Comprovantes")
comprovante_file = st.file_uploader("Subir Comprovante de Abastecimento (Imagem PNG/JPG):", type=["png", "jpg", "jpeg"])
st.markdown("---")

endereco_cliente_sugerido = "Rua Pascoal Carignano, 675 - Ferraria, Campo Largo"
if not df_leg.empty and cliente_selecionado:
    try:
        solicitantes_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Endereco"].dropna()
        if not solicitantes_df.empty:
            endereco_cliente_sugerido = str(solicitantes_df.to_list()[0]).strip()
    except: pass

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    try:
        sol_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna()
        if not sol_df.empty: 
            gerente_cliente_sugerido = str(sol_df.to_list()[0]).strip()
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
    
    corpo = f"<html><body><p>Prezada Suellen,</p><p>Segue em anexo o relatório consolidado de reembolso de KM rodado referente ao cliente <b>{cliente}</b>.</p><br><p>Atenciosamente,<br>Hudson Valente</p></body></html>"
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
        server.login(e_remetente, senate_remetente if 'senate_remetente' in locals() else senha_remetente)
        server.sendmail(e_remetente, email_destino, msg.as_string())
        server.quit()
        return True, "Relatório de KM enviado com sucesso!"
    except Exception as e:
        return False, f"Falha no envio: {str(e)}"

# --- GERAÇÃO E PROCESSAMENTO DINÂMICO ---
if st.button("Gerar Relatório de Reembolso de KM", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Selecione um cliente válido.")
    else:
        with st.spinner("⏳ Compilando dados de percurso..."):
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
                
                if col_situacao in df_dados.columns:
                    df_dados[col_situacao] = df_dados[col_situacao].astype(str).str.strip()
                
                atendimentos_km = df_dados[
                    (df_dados[col_cliente] == cliente_selecionado) & 
                    (df_dados[col_situacao] == "Em Elaboração") & 
                    (df_dados[col_ra].notna()) & 
                    (df_dados[col_km_d].notna()) & (df_dados[col_km_d] > 0)
                ].copy()
                
                if atendimentos_km.empty:
                    st.warning(f"⚠️ Nenhum lançamento ativo localizado para o cliente '{cliente_selecionado}' na aba '{aba_mes_selecionada}'.")
                else:
                    atendimentos_km = atendimentos_km.reset_index(drop=True)
                    atendimentos_km[col_data] = pd.to_datetime(atendimentos_km[col_data], errors='coerce')
                    atendimentos_km = atendimentos_km.sort_values(by=col_data).reset_index(drop=True)
                    
                    lista_linhas = []
                    t_km = 0.0
                    t_vlr = 0.0
                    
                    dict_km = atendimentos_km[col_km_d].to_dict()
                    dict_dt = atendimentos_km[col_data].to_dict()
                    
                    # O laço itera de forma contínua acumulando os valores corretamente
                    for r_idx, idx in enumerate(atendimentos_km.index):
                        dt_obj = dict_dt.get(idx)
                        dt_str = dt_obj.strftime("%d/%m/%Y") if pd.notnull(dt_obj) else ""
                        km_f = float(dict_km.get(idx, 0))
                        
                        # Regra mestre de alternância baseada na posição física da linha filtrada
                        if r_idx % 2 == 0:
                            percurso, orig, dest = "Ida", ENDERECO_CRTI_FIXO, endereco_cliente_sugerido
                        else:
                            percurso, orig, dest = "Volta", endereco_cliente_sugerido, ENDERECO_CRTI_FIXO
                            
                        vlr_f = km_f * (vlr_abast * 0.25)
                        t_km += km_f
                        t_vlr += vlr_f
                        
                        lista_linhas.append([dt_str, cliente_selecionado, orig, dest, percurso, f"{km_f:.0f}", f"R$ {formatar_br(vlr_abast)}", f"R$ {formatar_br(vlr_f)}"])

                    # 1. CONSTRUTOR EXCEL ESPELHADO (.XLSX)
                    df_excel = pd.DataFrame(lista_linhas)
                    df_excel.columns = ["DATA", "CLIENTE", "LOCAL ORIGEM", "LOCAL DESTINO", "Percurso", "DISTÂNCIA (KM)", "Vlr Unit. Ultimo Abast.", "TOTAL"]
                    df_excel.loc[len(df_excel)] = ["Total KM", f"{t_km:.0f}", "Valor Total", f"R$ {formatar_br(t_vlr)}", "", "", "", ""]
                    
                    buffer_xlsx = io.BytesIO()
                    with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                        df_excel.to_excel(writer, sheet_name="Reembolso", index=False)
                    st.session_state["km_xlsx_p04"] = buffer_xlsx.getvalue()

                    # 2. CONSTRUTOR PDF IDÊNTICO COM PROPORÇÃO DE MARGENS RECALCULADA
                    pdf = PDFReembolsoKM()
                    pdf.add_page()
                    
                    ARQUIVO_LOGO = "crti.jpg"
                    if os.path.exists(ARQUIVO_LOGO):
                        pdf.image(ARQUIVO_LOGO, x=15, y=10, w=32)
                    
                    pdf.set_font("Arial", "B", 11)
                    pdf.set_text_color(0, 0, 0)
                    pdf.text(100, 16, "RELATÓRIO PARA REEMBOLSO DE KM RODADO")
                    
                    pdf.set_font("Arial", "B", 8)
                    pdf.text(15, 30, "IMPLANTADOR CRTI: HUDSON PEDRO SALES VALENTE")
                    
                    pdf.set_y(34); pdf.set_x(15)
                    pdf.set_fill_color(245, 245, 245); pdf.set_font("Arial", "B", 6.5)
                    
                    headers = ["DATA", "CLIENTE", "LOCAL ORIGEM", "LOCAL DESTINO", "Percurso", "DISTÂNCIA (KM)", "Vlr Unit. Ultimo Abast.", "TOTAL"]
                    # DIMENSÕES CORRIGIDAS: Larguras perfeitas em mm para fechar as margens da folha sem estourar para o lado
                    h_widths = [15, 28, 41, 41, 13, 14, 15, 13]
                    
                    for idx_h, txt in enumerate(headers): 
                        pdf.cell(h_widths[idx_h], 6, txt, border=1, align="C", fill=True)
                    
                    pdf.set_font("Arial", "", 5.5)
                    for row in lista_linhas:
                        pdf.ln(6); pdf.set_x(15)
                        for col_i, val in enumerate(row):
                            align_cell = "L" if col_i in [2, 3] else "C"
                            pdf.cell(h_widths[col_i], 6, str(val), border=1, align=align_cell)
                        
                    # Rodapé de totalização consolidado e cravado com os 104 KM acumulados
                    pdf.ln(6); pdf.set_x(15); pdf.set_font("Arial", "B", 6.5)
                    pdf.cell(15 + 28 + 41, 6, "", border=0) # Pula Data, Cliente e Origem
                    pdf.cell(41, 6, "Total KM", border=1, align="R", fill=True)
                    pdf.cell(13, 6, f"{t_km:.0f}", border=1, align="C")
                    pdf.cell(14, 6, "", border=0) # Pula Vlr Unit
                    pdf.cell(15, 6, "Valor Total", border=1, align="R", fill=True)
                    pdf.cell(13, 6, f"R$ {formatar_br(t_vlr)}", border=1, align="C")
                    
                    st.session_state["km_pdf_p04"] = pdf.output(dest="S").encode("latin1")
                    st.success("✨ Relatório gerado com sucesso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro na compilação: {e}")

# --- PAINEL VISUAL DE DOWNLOADS ---
if "km_pdf_p04" in st.session_state:
    st.markdown("---")
    n_pdf = f"Reembolso_KM_{cliente_selecionado}_{aba_mes_selecionada}.pdf"
    n_xlsx = f"Reembolso_KM_{cliente_selecionado}_{aba_mes_selecionada}.xlsx"
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(label="📥 **BAIXAR RELATÓRIO PDF**", data=st.session_state["km_pdf_p04"], file_name=n_pdf, mime="application/pdf", use_container_width=True)
    with col_dl2:
        st.download_button(label="📥 **BAIXAR PLANILHA EXCEL**", data=st.session_state["km_xlsx_p04"], file_name=n_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# GATILHO FIXO E ALINHADO PARA O DISPARO SMTP DO FECHAMENTO MENSAL
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
                else: 
                    st.error(msg)
                st.rerun()
        with cp2:
            if st.button("Não, Cancelar", use_container_width=True): 
                st.rerun()

if st.button("🚀 **ENVIAR REEMBOLSO POR E-MAIL**", type="primary", use_container_width=True):
    if not email_target: 
        st.error("Insira um endereço de e-mail válido.")
    else: 
        confirmar_envio_km_popup()
