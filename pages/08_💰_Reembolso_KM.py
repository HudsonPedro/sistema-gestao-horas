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

# --- CLASSE PDF LIMPA SEM QUADROS (ORIENTAÇÃO PAISAGEM JÁ VALIDADA) ---
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
ENDERECO_CRTI_PADRAO = "Rua Padre Anchieta, 2050, Bigorrilho - Curitiba/PR"

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
st.markdown("---")

col_f1, col_f2 = st.columns(2)
with col_f1:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Cliente:")
    aba_mes_selecionada = st.selectbox("Selecione o Mês de Referência:", lista_abas_meses) if lista_abas_meses else st.text_input("Mês:")
with col_f2:
    vlr_abast = st.number_input("Valor Unitário Último Abastecimento (R$):", min_value=0.0, value=6.49, step=0.01)
    email_target = st.text_input("Destinatário do Relatório:", "financeiro@crti.com.br")

# CAPTURA DINÂMICA DE ENDEREÇOS DA PLANILHA LEGENDAS
endereco_cliente_map = ""
endereco_crti_erp_map = ENDERECO_CRTI_PADRAO

if not df_leg.empty and cliente_selecionado:
    try:
        linha_c = df_leg[df_leg["Clientes"] == cliente_selecionado]
        if not linha_c.empty and "Endereco" in df_leg.columns:
            endereco_cliente_map = str(linha_c["Endereco"].values[0]).strip()
        
        linha_crti = df_leg[df_leg["Clientes"].astype(str).str.contains("CRTI", case=False, na=False)]
        if not linha_crti.empty and "Endereco" in df_leg.columns:
            endereco_crti_erp_map = str(linha_crti["Endereco"].values[0]).strip()
    except:
        pass

if not endereco_cliente_map or endereco_cliente_map.lower() == "nan":
    endereco_cliente_map = "Endereço não localizado na aba Legendas"

# EXIBE AS CAIXAS DE TEXTO COM OS ENDEREÇOS DINÂMICOS NA INTERFACE
st.markdown("### 🗺️ Configuração de Rota e Percurso")
col_end1, col_end2 = st.columns(2)
with col_end1:
    end_ida_input = st.text_input("Endereço (Empresa Cliente):", value=endereco_cliente_map)
with col_end2:
    end_crti_input = st.text_input("Endereço (CRTI ERP):", value=endereco_crti_erp_map)

st.markdown("### 📄 Comprovantes")
comprovante_file = st.file_uploader("Subir Comprovante de Abastecimento (Imagem PNG/JPG):", type=["png", "jpg", "jpeg"])
st.markdown("---")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    try:
        if "Solicitante1" in df_leg.columns:
            sol_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna()
            if not sol_df.empty: gerente_cliente_sugerido = str(sol_df.values[0]).strip()
    except: pass

solicitante_nome = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
data_emissao = st.text_input("Data de Emissão:", value=datetime.now().strftime("%d/%m/%Y"))

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
    
    corpo = f"<html><body><p>Prezada Sra. Amanda, espero que se encontre bem,</p><p>Segue em anexo o relatório de reembolso de KM rodado e o comprovante de abastecimentos, referente ao atendimento presencial no cliente <b>{cliente} no dia: {data_emissao}.</b>.</p><br><p>Com Gratidão,<br>Hudson Valente</p></body></html>"
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
# --- REQUISITO ANEXO: ARMAZENAMENTO E PRÉ-VISUALIZAÇÃO EM TEMPO REAL ---
lista_linhas_preview = []
t_km_acumulado = 0.0
t_vlr_acumulado = 0.0

if cliente_selecionado and aba_mes_selecionada:
    try:
        res_dados = requests.get(URL_PLANILHA_MUDANCA, timeout=30)
        xl_dados = pd.ExcelFile(io.BytesIO(res_dados.content))
        df_dados = pd.read_excel(xl_dados, sheet_name=aba_mes_selecionada, engine='openpyxl')
        df_dados.columns = df_dados.columns.str.strip()
        
        col_cliente = "CLIENTE"
        col_situacao = "SITUACAO_RA"
        col_km_d = "KM_D"
        col_data = "DATA"
        col_idavolta_planilha = "IDA_VOLTA"
        
        if col_situacao in df_dados.columns:
            df_dados[col_situacao] = df_dados[col_situacao].astype(str).str.strip()
            
        atendimentos_filtrados = df_dados[
            (df_dados[col_cliente] == cliente_selecionado) & 
            (df_dados[col_situacao] == "Em Elaboração") & 
            (df_dados[col_km_d].notna()) & (df_dados[col_km_d] > 0)
        ].copy()
        
        if not atendimentos_filtrados.empty:
            atendimentos_filtrados[col_data] = pd.to_datetime(atendimentos_filtrados[col_data], errors='coerce')
            atendimentos_filtrados = atendimentos_filtrados.sort_values(by=col_data)
            
            for idx in atendimentos_filtrados.index:
                dt_obj = atendimentos_filtrados.loc[idx, col_data]
                dt_str = dt_obj.strftime("%d/%m/%Y") if pd.notnull(dt_obj) else ""
                km_f = float(atendimentos_filtrados.loc[idx, col_km_d])
                celula_ad = str(atendimentos_filtrados.loc[idx, col_idavolta_planilha]).strip().upper() if col_idavolta_planilha in atendimentos_filtrados.columns else ""
                
                if "IDA" in celula_ad:
                    percurso_linha = "Ida"
                    origem_linha = end_crti_input
                    destino_linha = end_ida_input
                else:
                    percurso_linha = "Volta"
                    origem_linha = end_ida_input
                    destino_linha = end_crti_input
                    
                vlr_f = km_f * (vlr_abast * 0.25)
                t_km_acumulado += km_f
                t_vlr_acumulado += vlr_f
                
                lista_linhas_preview.append([
                    dt_str, cliente_selecionado, origem_linha, destino_linha, 
                    percurso_linha, f"{km_f:.0f}", f"R$ {formatar_br(vlr_abast)}", f"R$ {formatar_br(vlr_f)}"
                ])
                
            st.markdown("### 📋 Lançamentos de Ida e Volta")
            headers_tabela = ["DATA", "CLIENTE", "LOCAL ORIGEM", "LOCAL DESTINO", "PERCURSO", "DISTÂNCIA(KM)", "VLR. UNI. ABAST.", "TOTAL"]
            df_preview_tela = pd.DataFrame(lista_linhas_preview, columns=headers_tabela)
            st.dataframe(df_preview_tela, use_container_width=True)
            
            st.info(f"📈 **Resumo Acumulado na Tela:** {t_km_acumulado:.0f} KM Rodados | Valor Total Reembolso: **R$ {formatar_br(t_vlr_acumulado)}**")
    except Exception as e:
        st.error(f"Erro no processamento mestre: {e}")

# --- MOTOR DE COMPILAÇÃO E COMPOSIÇÃO DOS ARQUIVOS ---
if st.button("Gerar Relatório de Reembolso de KM", type="primary", use_container_width=True):
    if not lista_linhas_preview:
        st.error("Gere os lançamentos na tela antes de exportar.")
    else:
        with st.spinner("⏳ Estruturando arquivos gêmeos PDF e Excel..."):
            try:
                caminho_imagem_disco = ""
                if comprovante_file:
                    caminho_imagem_disco = os.path.join(PASTA_REEMBOLSO_KM, "comprovante_km_final.jpg")
                    img_pil_conv = PILImage.open(io.BytesIO(comprovante_file.read()))
                    img_pil_conv.convert("RGB").save(caminho_imagem_disco, "JPEG")

                headers = ["DATA", "CLIENTE", "LOCAL ORIGEM", "LOCAL DESTINO", "PERCURSO", "DISTÂNCIA(KM)", "VLR. UNI. ABAST.", "TOTAL"]
                ARQUIVO_LOGO = "crti.jpg"

                # 1. PLANILHA EXCEL ESPELHADA (.XLSX) - AJUSTADA EM FOLHA ÚNICA PAISAGEM COM LOGO
                out_xlsx = io.BytesIO()
                wb = xlsxwriter.Workbook(out_xlsx, {"in_memory": True})
                ws = wb.add_worksheet("Reembolso")
                ws.hide_gridlines(2)
                
                # FORÇA PÁGINA ÚNICA HORIZONTAL NO EXCEL SEM QUEBRAS DE LINHA PONTILHADAS
                ws.set_landscape()
                ws.set_paper(9) # Papel A4
                ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
                ws.fit_to_pages(1, 0) # Trava em 1 página de largura máxima
                
                f_tit = wb.add_format({"bold": True, "size": 11, "font_name": "Arial"})
                f_sub = wb.add_format({"bold": True, "size": 9, "font_name": "Arial"})
                f_head = wb.add_format({"bold": True, "bg_color": "#E2EFDA", "border": 1, "align": "center", "font_name": "Arial", "size": 8.5})
                f_cel = wb.add_format({"border": 1, "align": "center", "font_name": "Arial", "size": 8})
                f_cel_l = wb.add_format({"border": 1, "align": "left", "font_name": "Arial", "size": 8})
                f_tot = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "align": "right", "font_name": "Arial", "size": 8.5})
                
                # Dimensões exatas das colunas do Excel espelhadas com o PDF
                ws.set_column("A:A", 14)
                ws.set_column("B:B", max(35, len(cliente_selecionado) + 4))
                ws.set_column("C:C", 52)
                ws.set_column("D:D", 52)
                ws.set_column("E:E", 14)
                ws.set_column("F:F", 18)
                ws.set_column("G:G", 22)
                ws.set_column("H:H", 18)
                
                # Insere a logo física na célula A2 para matar o texto provisório
                if os.path.exists(ARQUIVO_LOGO):
                    ws.insert_image("A2", ARQUIVO_LOGO, {"x_scale": 1.8, "y_scale": 1.8}) #0.48
                
                ws.write("D2", "RELATÓRIO PARA REEMBOLSO DE KM RODADO", f_tit)
                ws.write("A4", f"IMPLANTADOR CRTI: HUDSON VALENTE", f_sub)
                
                for c_idx, txt in enumerate(headers): 
                    ws.write(5, c_idx, txt, f_head)
                
                for r_idx, row in enumerate(lista_linhas_preview):
                    for c_idx, val in enumerate(row):
                        # Clientes e Endereços (índices 1, 2 e 3) alinhados à esquerda, o resto centralizado
                        fmt_c = f_cel_l if c_idx in [1, 2, 3] else f_cel
                        ws.write(6 + r_idx, c_idx, val, fmt_c)
                    
                l_f = 6 + len(lista_linhas_preview)
                ws.write(l_f, 4, "Total KM", f_tot)
                ws.write(l_f, 5, f"{t_km_acumulado:.0f}", f_cel)
                ws.write(l_f, 6, "Valor Total", f_tot)
                ws.write(l_f, 7, f"R$ {formatar_br(t_vlr_acumulado)}", f_cel)

                ws.write("A4", "COMPROVANTE DE ABASTECIMENTO ANEXADO", f_sub)
                
                if caminho_imagem_disco and os.path.exists(caminho_imagem_disco):
                    ws.insert_image(l_f + 3, 2, caminho_imagem_disco, {"x_scale": 1.0, "y_scale": 1.0}) #{"x_scale": 0.42, "y_scale": 0.42})
                    
                wb.close()
                st.session_state["km_xlsx_p04"] = out_xlsx.getvalue()

                # 2. RELATÓRIO PDF IDÊNTICO EM MODO PAISAGEM (MANTIDO IMPECÁVEL)
                pdf = PDFReembolsoKM(orientation='L', unit='mm', format='A4')
                pdf.add_page()
                
                if os.path.exists(ARQUIVO_LOGO):
                    pdf.image(ARQUIVO_LOGO, x=15, y=10, w=35)
                
                pdf.set_font("Arial", "B", 12)
                pdf.text(120, 16, "RELATÓRIO PARA REEMBOLSO DE KM RODADO")
                
                pdf.set_font("Arial", "B", 9)
                pdf.text(15, 32, "IMPLANTADOR CRTI: HUDSON VALENTE")
                
                pdf.set_y(36); pdf.set_x(15)
                pdf.set_fill_color(226, 239, 218); pdf.set_font("Arial", "B", 7.5)
                
                h_widths = [16, 32, 58, 58, 18, 22, 31, 23]
                for idx_h, txt in enumerate(headers): 
                    pdf.cell(h_widths[idx_h], 6, txt, border=1, align="C", fill=True)
                
                pdf.set_font("Arial", "", 6.5)
                for row in lista_linhas_preview:
                    pdf.ln(6); pdf.set_x(15)
                    for col_i, val in enumerate(row):
                        align_cell = "L" if col_i in [1, 2, 3] else "C"
                        pdf.cell(h_widths[col_i], 6, str(val), border=1, align=align_cell)
                    
                pdf.ln(6); pdf.set_x(15); pdf.set_font("Arial", "B", 7.5)
                pdf.cell(16 + 32 + 58 + 58, 6, "", border=0)
                pdf.cell(18, 6, "Total KM", border=1, align="R", fill=True)#16
                pdf.cell(22, 6, f"{t_km_acumulado:.0f}", border=1, align="C")
                pdf.cell(31, 6, "Valor Total", border=1, align="R", fill=True) #34
                pdf.cell(23, 6, f"R$ {formatar_br(t_vlr_acumulado)}", border=1, align="C")
                
                if caminho_imagem_disco and os.path.exists(caminho_imagem_disco):
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 11)
                    pdf.text(15, 15, "COMPROVANTE DE ABASTECIMENTO ANEXADO")
                    pdf.image(caminho_imagem_disco, x=20, y=27, w=115) #x=15, y=22, w=110)
                
                st.session_state["km_pdf_p04"] = pdf.output(dest="S").encode("latin1")
                
                if caminho_imagem_disco and os.path.exists(caminho_imagem_disco):
                    try: os.remove(caminho_imagem_disco)
                    except: pass
                        
                st.success("✨ Relatórios gêmeos gerados com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro na compilação: {e}")
    # --- PAINEL VISUAL DE DOWNLOADS ---
if "km_pdf_p04" in st.session_state:
    st.markdown("---")
    n_pdf = f"Relatorio_Reembolso_KM_{cliente_selecionado}.pdf"
    n_xlsx = f"Relatorio_Reembolso_KM_{cliente_selecionado}.xlsx"
    
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
