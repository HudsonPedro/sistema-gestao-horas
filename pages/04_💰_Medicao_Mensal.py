# Por Hudson Valente - HPTECH
# Gerador de Medição Mensal Automatizado - PDF + XLSX
import io
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from fpdf import FPDF
import streamlit as st
import xlsxwriter
import pandas as pd
import requests

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Medição Mensal - HPTECH", 
    page_icon="hptechICO.png", 
    layout="wide"
)

# 2. DESIGN CSS PADRÃO DO SEU APP
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -10px;
        }
        h1 { color: #b0231d; }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE CONVERSÃO E FORMATAÇÃO BRASILEIRA ---
def horas_para_decimal(tempo_str):
    """Converte '138:25:00' para float seguro"""
    tempo_str = str(tempo_str).strip()
    if ":" not in tempo_str:
        return 0.0
    partes = tempo_str.split(":")
    horas = int(partes[0])
    minutos = int(partes[1]) if len(partes) > 1 else 0
    return horas + (minutos / 60.0)

def formatar_br(valor):
    """Formata número para o padrão brasileiro: 11.073,33"""
    try:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"

# --- CARREGAR BASE DE DADOS DO GOOGLE SHEETS ---
@st.cache_data(ttl=600)
def carregar_planilha_todas_abas():
    url = https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx
    response = requests.get(url)
    return pd.read_excel(io.BytesIO(response.content), sheet_name=None, engine='openpyxl')

# 3. SIDEBAR COM MENU DE NAVEGAÇÃO E ATUALIZAÇÃO DA BASE
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
    
    st.divider()
    st.markdown("### Configurações GERAIS")
    if st.button("🔄 Atualizar Planilha", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.title("💰 Medição Mensal de Prestação de Serviços")
st.markdown("---")

# Conectando à planilha
with st.spinner("Analisando dados das planilhas..."):
    try:
        dict_abas = carregar_planilha_todas_abas()
        abas_disponiveis = list(dict_abas.keys())
    except Exception as e:
        st.error(f"Erro ao baixar planilha base: {e}")
        st.stop()

# --- CONTROLES DE FILTRO DIRETOS NA TELA ---
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    aba_selecionada = st.selectbox("**Selecione o Mês de Faturamento:**", abas_disponiveis)
with col_f2:
    numero_medicao = st.number_input("**Número da Medição:**", min_value=1, value=37)
with col_f3:
    valor_hora = st.number_input("**Preço da Hora (R$):**", min_value=0.0, value=80.00, step=5.0)

# Inicializa o dicionário com os dados fixos do cliente antes do processamento
dados_faturamento = {
    "parceiro": "CR Tecnologia da Informação Ltda",
    "endereco": "Rua Padre Anchieta, 2050 - Bairro Bigorrilho",
    "cidade_uf": "Curitiba - PR",
    "cep": "80730-000",
    "cnpj": "04.616.592/0001-21",
    "numero_medicao": numero_medicao,
    "data_inicio": "01/04/2026",
    "data_fim": "30/04/2026",
    "mes_ano": str(aba_selecionada).lower()[:6],
    "descricao_servico": "Prestação de serviços de consultoria Implantação",
    "qtd_horas": "0:00:00",
    "preco_unitario": valor_hora,
    "preco_total": 0.0,
}

# Processando dados da aba selecionada
df_mes = dict_abas[aba_selecionada].copy()

# Normaliza as colunas de texto para evitar erros com células nulas ou vazias
df_mes["SITUACAO_RA"] = df_mes["SITUACAO_RA"].fillna("").astype(str).str.strip().str.lower()
df_mes["TOTAL_HR"] = df_mes["TOTAL_HR"].fillna("").astype(str).str.strip()

# Filtra estritamente os registros que estão em elaboração neste mês
df_filtrado = df_mes[df_mes["SITUACAO_RA"] == "em elaboração"]

# Realiza o somatório convertendo as strings de horas da planilha para segundos totais
total_segundos = 0
for val in df_filtrado["TOTAL_HR"]:
    val_str = str(val).strip()
    if ":" in val_str:
        try:
            partes = val_str.split(":")
            h = int(partes[0])
            m = int(partes[1]) if len(partes) > 1 else 0
            s = int(partes[2]) if len(partes) > 2 else 0
            total_segundos += (h * 3600) + (m * 60) + s
        except (ValueError, IndexError):
            continue

# Reconverte o total de segundos para o formato estruturado HH:MM:SS
horas_inteiras = int(total_segundos // 3600)
minutos_restantes = int((total_segundos % 3600) // 60)
segundos_restantes = int(total_segundos % 60)

total_horas_faturar = f"{horas_inteiras}:{minutos_restantes:02d}:{segundos_restantes:02d}"

# Efetua a conversão para decimal e calcula o preço final do faturamento
horas_dec = horas_para_decimal(total_horas_faturar)
preco_total_calculado = horas_dec * valor_hora

# Tenta capturar o intervalo real de datas da coluna DATA da planilha
if "DATA" in df_mes.columns:
    df_mes["DATA_DT"] = pd.to_datetime(df_mes["DATA"], errors="coerce")
    df_validas = df_mes.dropna(subset=["DATA_DT"])
    if not df_validas.empty:
        dados_faturamento["data_inicio"] = df_validas["DATA_DT"].min().strftime("%d/%m/%Y")
        dados_faturamento["data_fim"] = df_validas["DATA_DT"].max().strftime("%d/%m/%Y")

# Injeta os valores reais calculados no dicionário de contexto corporativo
dados_faturamento["qtd_horas"] = total_horas_faturar
dados_faturamento["preco_total"] = preco_total_calculado

# --- PRÉVIA DOS RESULTADOS CALCULADOS ---
st.markdown("### 📊 Resumo do Faturamento Calculado da Planilha")
col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Total de Horas Encontradas", total_horas_faturar)
col_m2.metric("Preço Unitário da Hora", f"R$ {formatar_br(valor_hora)}")
col_m3.metric("Preço Total Calculado", f"R$ {formatar_br(preco_total_calculado)}")

# --- CLASSE DO PDF DA MEDIÇÃO (MOLDURAS VERMELHAS) ---
class PDFMedicaoNovo(FPDF):
    def moldura_topo(self, x, y, w, h, dados):
        self.set_draw_color(255, 0, 0)
        self.set_line_width(0.5)
        self.rect(x, y, w, h)
        self.set_font("Arial", "", 9)
        self.set_text_color(0, 0, 0)
        linhas = [
            f"Parceiro:       {dados['parceiro']}",
            f"Endereço:     {dados['endereco']}",
            f"Cidade / UF:  {dados['cidade_uf']}",
            f"CEP:             {dados['cep']}",
            f"CNPJ:           {dados['cnpj']}",
        ]
        curr_y = y + 4
        for linha in linhas:
            self.text(x + 4, curr_y, linha)
            curr_y += 5
        self.line(x + 110, y, x + 110, y + h)
        self.text(x + 114, y + 8, "Medição Número:")
        self.set_font("Arial", "B", 10)
        self.text(x + 160, y + 8, str(dados["numero_medicao"]))
        self.set_font("Arial", "", 9)
        self.text(x + 114, y + 18, f"Período:  {dados['data_inicio']}    até    {dados['data_fim']}")

def gerar_pdf_medicao_nova(dados):
    pdf = PDFMedicaoNovo(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.text(15, 20, "Medição Mensal de Prestação de Serviços")
    pdf.moldura_topo(15, 28, 180, 32, dados)
    pdf.set_font("Arial", "B", 10)
    pdf.text(15, 70, "* Serviços Executados")
    pdf.set_y(74); pdf.set_x(15)
    pdf.set_draw_color(180, 180, 180); pdf.set_fill_color(245, 245, 245); pdf.set_font("Arial", "B", 8)
    headers = [("Mês/Ano", 18), ("Item", 10), ("Descrição", 72), ("Unidade", 15), ("Qtd", 20), ("Preço Unitário", 22), ("Preço Total", 23)]
    for txt, w in headers: pdf.cell(w, 7, txt, border=1, align="C", fill=True)
    pdf.set_y(81); pdf.set_x(15); pdf.set_font("Arial", "", 8)
    pdf.cell(18, 10, dados["mes_ano"], border=1, align="C")
    pdf.cell(10, 10, "1", border=1, align="C")
    pdf.cell(72, 10, dados["descricao_servico"], border=1, align="L")
    pdf.cell(15, 10, "HR", border=1, align="C")
    pdf.cell(20, 10, dados["qtd_horas"], border=1, align="C")
    pdf.cell(22, 10, formatar_br(dados["preco_unitario"]), border=1, align="C")
    pdf.set_draw_color(255, 0, 0)
    pdf.cell(23, 10, formatar_br(dados["preco_total"]), border=1, align="C")
    pdf.set_y(91); pdf.set_x(15); pdf.set_draw_color(180, 180, 180)
    pdf.cell(28, 7, "", border=0); pdf.set_font("Arial", "B", 8)
    pdf.cell(72, 7, "TOTAL", border=1, align="L", fill=True); pdf.cell(57, 7, "", border=0)
    pdf.set_draw_color(255, 0, 0)
    pdf.cell(23, 7, formatar_br(dados["preco_total"]), border=1, align="C")
    pdf.set_text_color(100, 100, 100); pdf.set_font("Arial", "I", 7)
    pdf.text(115, 104, "* Duplicatas a serem emitidas")
    pdf.text(115, 107, f"HP SERVIÇOS ADM, valor total de R$ {formatar_br(dados['preco_total'])}")
    pdf.set_draw_color(255, 0, 0); pdf.rect(15, 115, 180, 35)
    pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", "B", 9); pdf.text(18, 120, "* De acordo com a Medição Mensal")
    pdf.set_draw_color(150, 150, 150); pdf.line(20, 140, 85, 140); pdf.line(125, 140, 190, 140)
    pdf.set_font("Arial", "", 8); pdf.text(20, 144, "HP SERVIÇOS ADM"); pdf.text(125, 144, "CRTI")
    return pdf.output(dest="S").encode("latin1")

# --- GERADOR PLANILHA EXCEL ---
def gerar_xlsx_medicao_nova(dados):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet("Medição")
    fmt_titulo = workbook.add_format({"bold": True, "size": 14})
    fmt_borda_vermelha = workbook.add_format({"border": 1, "border_color": "red"})
    fmt_header = workbook.add_format({"bold": True, "bg_color": "#F5F5F5", "border": 1, "border_color": "#B0B0B0", "align": "center"})
    fmt_celula = workbook.add_format({"border": 1, "border_color": "#B0B0B0", "align": "center"})
    fmt_celula_esq = workbook.add_format({"border": 1, "border_color": "#B0B0B0", "align": "left"})
    fmt_total_vermelho = workbook.add_format({"bold": True, "border": 1, "border_color": "red", "align": "center"})
    worksheet.set_column("A:G", 15); worksheet.set_column("C:C", 40)
    worksheet.write("A2", "Medição Mensal de Prestação de Serviços", fmt_titulo)
    worksheet.merge_range("A4:D8", "", fmt_borda_vermelha)
    worksheet.write("A4", f"  Parceiro: {dados['parceiro']}")
    worksheet.write("A5", f"  Endereço: {dados['endereco']}")
    worksheet.write("A6", f"  Cidade/UF: {dados['cidade_uf']}")
    worksheet.write("A7", f"  CEP: {dados['cep']}")
    worksheet.write("A8", f"  CNPJ: {dados['cnpj']}")
    worksheet.merge_range("E4:G8", "", fmt_borda_vermelha)
    worksheet.write("E4", "  Medição Número:")
    worksheet.write("F4", dados["numero_medicao"], workbook.add_format({"bold": True}))
    worksheet.write("E6", f"  Período: {dados['data_inicio']} até {dados['data_fim']}")
    worksheet.write("A10", "* Serviços Executados", workbook.add_format({"bold": True}))
    headers = ["Mês/Ano", "Item", "Descrição", "Unidade", "Qtd", "Preço Unitário", "Preço Total"]
    for col, h in enumerate(headers): worksheet.write(10, col, h, fmt_header)
    worksheet.write(11, 0, dados["mes_ano"], fmt_celula)
    worksheet.write(11, 1, "1", fmt_celula)
    worksheet.write(11, 2, dados["descricao_servico"], fmt_celula_esq)
    worksheet.write(11, 3, "HR", fmt_celula)
    worksheet.write(11, 4, dados["qtd_horas"], fmt_celula)
    worksheet.write(11, 5, formatar_br(dados["preco_unitario"]), fmt_celula)
    worksheet.write(11, 6, formatar_br(dados["preco_total"]), fmt_total_vermelho)
    worksheet.write(12, 2, "TOTAL", fmt_header)
    worksheet.write(12, 6, formatar_br(dados["preco_total"]), fmt_total_vermelho)
    workbook.close()
    return output.getvalue()

def enviar_email_medicao_nova(email_destino, dados, pdf_bytes, xlsx_bytes):
    email_remetente = st.secrets["smtp"]["usuario"]
    senha_remetente = st.secrets["smtp"]["senha"]
    smtp_server = st.secrets["smtp"]["servidor"]
    smtp_porta = int(st.secrets["smtp"]["porta"])
    msg = MIMEMultipart(); msg["From"] = email_remetente; msg["To"] = email_destino
    msg["Subject"] = f"Medição Mensal de Prestação de Serviços N° {dados['numero_medicao']} - CRTI"
    corpo = f"<html><body><p>Olá,</p><p>Seguem anexados os relatórios da medição de serviços calculada de {dados['data_inicio']} até {dados['data_fim']}.</p><p>Total faturado: R$ {formatar_br(dados['preco_total'])}</p></body></html>"
    msg.attach(MIMEText(corpo, "html"))
    
    for b_data, ext in [(pdf_bytes, "pdf"), (xlsx_bytes, "xlsx")]:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(b_data); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename=Medicao_{dados['numero_medicao']}.{ext}")
        msg.attach(part)
    try:
        server = smtplib.SMTP(smtp_server, smtp_porta); server.starttls()
        server.login(email_remetente, senha_remetente); server.sendmail(email_remetente, email_destino, msg.as_string()); server.quit()
        return True, "E-mail enviado com sucesso!"
    except Exception as e: return False, f"Falha no envio SMTP: {str(e)}"

# --- EXIBIÇÃO DOS BOTÕES VISUAIS ---
st.markdown("---")
col_b1, col_b2 = st.columns(2)
with col_b1:
    st.download_button(label="📥 Baixar PDF da Medição (Formatado BR)", data=gerar_pdf_medicao_nova(dados_faturamento), file_name=f"Medicao_{numero_medicao}.pdf", mime="application/pdf", use_container_width=True)
with col_b2:
    st.download_button(label="📊 Baixar Excel da Medição (Formatado BR)", data=gerar_xlsx_medicao_nova(dados_faturamento), file_name=f"Medicao_{numero_medicao}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.markdown("---")
email_target = st.text_input("Destinatário da Medição:", "financeiro@crti.com.br")
if st.button("🚀 Enviar Medição por E-mail", use_container_width=True):
    with st.spinner("Compilando anexos formatados e enviando..."):
        p_b = gerar_pdf_medicao_nova(dados_faturamento)
        x_b = gerar_xlsx_medicao_nova(dados_faturamento)
        ok, r_msg = enviar_email_medicao_nova(email_target, dados_faturamento, p_b, x_b)
    if ok: st.success(r_msg); st.balloons()
    else: st.error(r_msg)
