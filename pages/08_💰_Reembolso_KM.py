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

# Pasta mestre para guardar o histórico gerado
PASTA_REEMBOLSO_KM = os.path.join(BASE_DIR, "termos_reembolso_km")
os.makedirs(PASTA_REEMBOLSO_KM, exist_ok=True)

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Reembolso de KM - HPTECH", page_icon="hptech.png", layout="wide")

# 2. CSS PARA OCULTAR O MENU PADRÃO E APLICAR SEU DESIGN
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
    h1 { color: #b0231d; }
    .user-block { background-color: #f0f2f6; padding: 8px; border-radius: 8px; margin-top: -10px; }
    </style>
""", unsafe_allow_html=True)

def formatar_br(valor):
    try: return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return "0,00"

# --- MOTOR CLASSE PDF REVISADA (PADRÃO DA PÁGINA 04) ---
class PDFReembolsoKM(FPDF):
    def moldura_topo(self, x, y, w, h, cliente, total_km, total_vlr):
        self.set_draw_color(180, 180, 180) 
        self.set_line_width(0.4)
        self.rect(x, y, w, h)
        self.set_font("Arial", "", 9)
        self.set_text_color(0, 0, 0)
        
        self.text(x + 4, y + 6, f"Parceiro: CR Tecnologia da Informação Ltda")
        self.text(x + 4, y + 12, f"Implantador CRTI: HUDSON PEDRO SALES VALENTE")
        self.text(x + 4, y + 18, f"Empresa Cliente: {cliente}")
        self.text(x + 4, y + 24, f"Data de Emissão: {datetime.now().strftime('%d/%m/%Y')}")
        
        self.line(x + 110, y, x + 110, y + h)
        self.set_font("Arial", "B", 10)
        self.text(x + 114, y + 8, "Resumo do Reembolso:")
        self.set_font("Arial", "", 9)
        self.text(x + 114, y + 16, f"Total Geral de KM: {total_km:.0f} KM")
        self.set_font("Arial", "B", 10)
        self.text(x + 114, y + 24, f"Valor do Reembolso: R$ {formatar_br(total_vlr)}")
# 3. SIDEBAR COM MENU INTEGRADO ATUALIZADO (PADRÃO INTERNACIONAL)
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
        for arq in glob.glob(os.path.join(PASTA_REEMBOLSO_KM, "*.*")):
            try: os.remove(arq)
            except: pass
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
st.write("Layout gerado do zero de forma interna no código, seguindo o padrão da Medição Mensal.")
st.markdown("---")

col_f1, col_f2 = st.columns(2)
with col_f1:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Cliente:")
    aba_mes_selecionada = st.selectbox("Selecione o Mês de Referência (Aba):", lista_abas_meses) if lista_abas_meses else st.text_input("Mês:")
with col_f2:
    vlr_abast = st.number_input("Valor Unitário Último Abastecimento (R$):", min_value=0.0, value=7.49, step=0.01)
    email_target = st.text_input("Destinatário do Relatório:", "suellen@crti.com.br")

endereco_cliente_sugerido = "Rua Pascoal Carignano, 675, Ferraria - Campo Largo"
if not df_leg.empty and cliente_selecionado:
    try:
        end_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Endereco"].dropna()
        if not end_df.empty: endereco_cliente_sugerido = str(end_df.values).strip()
    except: pass
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
    
    corpo = f"<html><body><p>Prezada Suellen,</p><p>Segue em anexo o relatório consolidado de reembolso de KM rodado e planilha referente aos atendimentos do cliente <b>{cliente}</b>.</p><br><p>Atenciosamente,<br>Hudson Valente</p></body></html>"
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

# --- 6. GERAÇÃO E PROCESSAMENTO EM CÓDIGO ---
if st.button("Gerar Relatório de Reembolso de KM", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Selecione um cliente válido.")
    else:
        with st.spinner("⏳ Compilando dados por dicionário nativo..."):
            try:
                res_dados = requests.get(URL_PLANILHA_MUDANCA, timeout=45, stream=True)
                xl_dados = pd.ExcelFile(io.BytesIO(res_dados.content))
                df_dados = pd.read_excel(xl_dados, sheet_name=aba_mes_selecionada, engine='openpyxl')
                
                df_dados = df_dados.reset_index(drop=True)
                df_dados.columns = df_dados.columns.str.strip()
                
                # CORREÇÃO CRUCIAL: Aponta para o nome correto da coluna mestre com acentuação
                col_cliente = "CLIENTE"
                col_situacao = "SITUAÇÃO"
                col_ra = "RA"
                col_km_d = "KM_D"
                col_data = "DATA"
                
                if col_situacao in df_dados.columns:
                    df_dados[col_situacao] = df_dados[col_situacao].astype(str).str.strip()
                
                # Filtra os lançamentos ativos com quilometragem válida e situação Em Elaboração
                atendimentos_km = df_dados[
                    (df_dados[col_cliente] == cliente_selecionado) & 
                    (df_dados[col_situacao] == "Em Elaboração") & 
                    (df_dados[col_km_d].notna()) & (df_dados[col_km_d] > 0)
                ].copy()
                
                if atendimentos_km.empty:
                    st.warning(f"⚠️ Nenhum lançamento ativo localizado para '{cliente_selecionado}' nesta aba.")
                else:
                    atendimentos_km[col_data] = pd.to_datetime(atendimentos_km[col_data], errors='coerce')
                    atendimentos_km = atendimentos_km.sort_values(by=col_data).reset_index(drop=True)
                    
                    lista_linhas = []
                    t_km = 0.0
                    t_vlr = 0.0
                    
                    dict_km = atendimentos_km[col_km_d].to_dict()
                    dict_dt = atendimentos_km[col_data].to_dict()
                    
                    for idx in atendimentos_km.index:
                        dt_obj = dict_dt.get(idx)
                        dt_str = dt_obj.strftime("%d/%m/%Y") if pd.notnull(dt_obj) else ""
                        km_f = float(dict_km.get(idx, 0))
                        
                        if idx % 2 == 0:
                            percurso, orig, dest = "Ida", ENDERECO_CRTI_FIXO, endereco_cliente_sugerido
                        else:
                            percurso, orig, dest = "Volta", endereco_cliente_sugerido, ENDERECO_CRTI_FIXO
                            
                        vlr_f = km_f * (vlr_abast * 0.25)
                        t_km += km_f
                        t_vlr += vlr_f
                        
                        lista_linhas.append([dt_str, cliente_selecionado, orig, dest, percurso, f"{km_f:.0f}", f"R$ {vlr_abast:,.2f}", f"R$ {vlr_f:,.2f}"])

                    # 1. MOTOR DO EXCEL (XlsxWriter - ALINHADO COM A PÁGINA 04)
                    out_xlsx = io.BytesIO()
                    wb = xlsxwriter.Workbook(out_xlsx, {"in_memory": True})
                    ws = wb.add_worksheet("Reembolso")
                    ws.hide_gridlines(2)
                    
                    f_tit = wb.add_format({"bold": True, "size": 13, "font_name": "Arial"})
                    f_head = wb.add_format({"bold": True, "bg_color": "#F5F5F5", "border": 1, "align": "center", "font_name": "Arial", "size": 9})
                    f_cel = wb.add_format({"border": 1, "align": "center", "font_name": "Arial", "size": 9})
                    f_tot = wb.add_format({"bold": True, "bg_color": "#F5F5F5", "border": 1, "font_name": "Arial", "size": 9})
                    
                    ws.set_column("A:A", 12); ws.set_column("B:B", 25); ws.set_column("C:C", 35); ws.set_column("D:D", 35)
                    ws.set_column("E:E", 10); ws.set_column("F:F", 10); ws.set_column("G:G", 14); ws.set_column("H:H", 14)
                    
                    ws.write("A2", "Relatório de Reembolso de KM Rodado", f_tit)
                    headers = ["data_dia", "cliente", "local_origem", "local_destino", "percurso", "km", "vlr_unit", "total"]
                    for c_idx, txt in enumerate(headers): ws.write(3, c_idx, txt, f_head)
                    
                    for r_idx, row in enumerate(lista_linhas):
                        for c_idx, val in enumerate(row): ws.write(4 + r_idx, c_idx, val, f_cel)
                        
                    l_f = 4 + len(lista_linhas)
                    ws.write(l_f, 0, "Total KM", f_tot); ws.write(l_f, 1, f"{t_km:.0f}", f_cel)
                    ws.write(l_f, 2, "Valor Total", f_tot); ws.write(l_f, 3, f"R$ {formatar_br(t_vlr)}", f_cel)
                    wb.close()
                    st.session_state["km_xlsx_p04"] = out_xlsx.getvalue()

                    # 2. MOTOR DO PDF (FPDF - ALINHADO COM A PÁGINA 04)
                    pdf = PDFReembolsoKM()
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 14)
                    pdf.text(15, 18, "Relatório Reembolso de KM Rodado")
                    pdf.moldura_topo(15, 24, 180, 28, cliente_selecionado, t_km, t_vlr)
                    
                    pdf.set_y(58); pdf.set_x(15)
                    pdf.set_fill_color(245, 245, 245); pdf.set_font("Arial", "B", 7)
                    
                    h_widths = [15, 25, 43, 43, 14, 12, 14, 14]
                    for idx_h, txt in enumerate(headers): pdf.cell(h_widths[idx_h], 6, txt, border=1, align="C", fill=True)
                    
                    pdf.set_font("Arial", "", 6.5)
                    for row in lista_linhas:
                        pdf.set_x(15)
                        for col_i, val in enumerate(row):
                            pdf.cell(h_widths[col_i], 6, str(val), border=1, align="C")
                        pdf.ln(6)
                        
                    pdf.set_x(15); pdf.set_font("Arial", "B", 7)
                    pdf.cell(15, 6, "TOTAL KM", border=1, align="C", fill=True)
                    pdf.cell(25, 6, f"{t_km:.0f}", border=1, align="C")
                    pdf.cell(43, 6, "VALOR TOTAL", border=1, align="C", fill=True)
                    pdf.cell(43, 6, f"R$ {formatar_br(t_vlr)}", border=1, align="C")
                    
                    st.session_state["km_pdf_p04"] = pdf.output(dest="S").encode("latin1")
                    st.success("✨ Documentos estruturados com sucesso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro na compilação: {e}")
# --- DOWNLOADS E POP-UP MESTRE DE DISPARO ---
if "km_pdf_p04" in st.session_state:
    st.markdown("---")
    n_pdf = f"Reembolso_KM_{cliente_selecionado}_{aba_mes_selecionada}.pdf"
    n_xlsx = f"Reembolso_KM_{cliente_selecionado}_{aba_mes_selecionada}.xlsx"
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(label="📥 **BAIXAR RELATÓRIO PDF**", data=st.session_state["km_pdf_p04"], file_name=n_pdf, mime="application/pdf", use_container_width=True)
    with col_dl2:
        st.download_button(label="📥 **BAIXAR PLANILHA EXCEL**", data=st.session_state["km_xlsx_p04"], file_name=n_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# CORREÇÃO CRUCIAL: Isola e deixa a caixa de disparo de e-mail fixa e visível na interface principal
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
