# Por Hudson Valente - HPTECH 14/05/2026 09:12
# Gerador de Medição Mensal Automatizado - PDF + XLSX
import io
import os  
import smtplib
from email.encoders import encode_base64 # CORREÇÃO DA SINTAXE DE IMPORTAÇÃO
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from fpdf import FPDF
import streamlit as st
import xlsxwriter
import pandas as pd
import requests
from datetime import datetime, timedelta
import base64

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="HPTECH Sistema de Gestão",
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

# --- FUNÇÕES DE CONVERSÃO E FORMATAÇÃO BRASILEIRA ---
def horas_para_decimal(tempo_str):
    tempo_str = str(tempo_str).strip()
    if ":" not in tempo_str:
        return 0.0
    partes = tempo_str.split(":")
    horas = int(partes[0])
    minutos = int(partes[1]) if len(partes) > 1 else 0
    return horas + (minutos / 60.0)

def formatar_br(valor):
    try:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"

# --- CARREGAR BASE DE DADOS DO GOOGLE SHEETS ---
#def carregar_planilha_todas_abas():
@st.cache_data(ttl=600)
def carregar_base_de_dados():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    response = requests.get(url)
    return pd.read_excel(io.BytesIO(response.content), sheet_name=None, engine='openpyxl')

# 3. SIDEBAR COM MENU DE NAVEGAÇÃO
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
    
      # Navegação Atualizada
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True):
        st.switch_page("pages/01_📊_Dashboard.py")
    if st.button("📝 Lançamento de Horas", use_container_width=True):
        st.switch_page("pages/03_📝_Lancamento.py")
    if st.button("📄 Relatórios RA", use_container_width=True):
        st.switch_page("pages/02_📄_Relatorios.py")    
    if st.button("📋 Termo Homologação", use_container_width=True): 
       st.switch_page("pages/05_📋_Termos.py")
    if st.button("📑 Termo Encerramento", use_container_width=True): 
        st.switch_page("pages/06_📑_Termo_Encerramento.py")
    if st.button("🚗 Termo Presencial", use_container_width=True): 
        st.switch_page("pages/07_🚗_Termo_Treinamento_Presencial.py")
    if st.button("💰 Reembolso de KM", use_container_width=True): 
        st.switch_page("pages/08_💰_Reembolso_KM.py")
             
    
    st.divider()
    st.sidebar.header("⚙️ Configurações GERAIS")
    if st.sidebar.button("🔄 Atualizar Base de Dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
  
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")


#st.title("💰 Medição Mensal de Prestação de Serviços")
#st.markdown("---")
# 4. CONTEÚDO DA HOME COM NOVO CARD
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Tenta carregar a imagem que está no repositório GitHub
try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <h1 style="margin: 0; font-size: 2.5rem;">Medição Mensal de Prestação de Serviço</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.markdown("Fechamento mensal consolidado.")
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("Medição Mensal de Prestação de Serviço")
st.markdown("---")

# Conectando à planilha
with st.spinner("Analisando dados..."):
    try:
        dict_abas = carregar_base_de_dados()
        abas_disponiveis = list(dict_abas.keys())
    except Exception as e:
        st.error(f"Erro ao baixar base de dados: {e}")
        st.stop()

# --- CONTROLES DE FILTRO DIRETOS NA TELA ---
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    aba_selecionada = st.selectbox("**Selecione o Mês de Faturamento:**", abas_disponiveis)
with col_f2:
    numero_medicao = st.number_input("**Número da Medição:**", min_value=1, value=40)
with col_f3:
    valor_hora = st.number_input("**Preço da Hora (R$):**", min_value=0.0, value=80.00, step=5.0)

# LÓGICA DE TRATAMENTO DINÂMICO DE DATAS DO CALENDÁRIO
try:
    nome_aba_limpo = str(aba_selecionada).strip()
    partes_abas = nome_aba_limpo.split(" ")
    
    meses_map = {
        "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
        "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
    }
    
    mes_nome_doc = partes_abas[0].lower()
    ano_doc = int(partes_abas[1]) if len(partes_abas) > 1 else datetime.now().year
    mes_num_doc = meses_map.get(mes_nome_doc, datetime.now().month)
    
    data_inicio_real = f"01/{mes_num_doc:02d}/{ano_doc}"
    
    if mes_num_doc == 12:
        ultimo_dia_mes = 31
    else:
        ultimo_dia_mes = (datetime(ano_doc, mes_num_doc + 1, 1) - timedelta(days=1)).day
        
    data_fim_real = f"{ultimo_dia_mes:02d}/{mes_num_doc:02d}/{ano_doc}"
    mes_ano_tabela = f"{mes_nome_doc}/{ano_doc}"
except:
    data_inicio_real = "01/05/2026"
    data_fim_real = "31/05/2026"
    mes_ano_tabela = "maio/2026"

# Inicializa o dicionário com as datas corrigidas
dados_faturamento = {
    "parceiro": "CR Tecnologia da Informação Ltda",
    "endereco": "Rua Padre Anchieta, 2050 - Bairro Bigorrilho",
    "cidade_uf": "Curitiba - PR",
    "cep": "80730-000",
    "cnpj": "04.616.592/0001-21",
    "numero_medicao": numero_medicao,
    "data_inicio": data_inicio_real,
    "data_fim": data_fim_real,
    "mes_ano": mes_ano_tabela,
    "descricao_servico": "Prestação de serviços de consultoria Implantação",
    "qtd_horas": "0:00:00",
    "preco_unitario": valor_hora,
    "preco_total": 0.0,
}

df_mes = dict_abas[aba_selecionada].copy()
df_mes["TOTAL_HR"] = df_mes["TOTAL_HR"].fillna("").astype(str).str.strip()

coluna_usuario = None
for col in df_mes.columns:
    if str(col).upper() in ["EMAIL", "CONSULTOR", "NOME", "USUARIO"]:
        coluna_usuario = col
        break

if coluna_usuario:
    df_mes[coluna_usuario] = df_mes[coluna_usuario].fillna("").astype(str).str.strip()
    lista_consultores = sorted(list(df_mes[coluna_usuario].unique()))
    index_padrao = 0
    for i, c in enumerate(lista_consultores):
        if "hudson" in c.lower() or "valente" in c.lower():
            index_padrao = i
            break
            
    consultor_sel = st.selectbox("**Filtrar por Consultor / Usuário:**", lista_consultores, index=index_padrao)
    df_filtrado = df_mes[df_mes[coluna_usuario] == consultor_sel]
else:
    df_filtrado = df_mes.copy()

total_segundos = 0
for val in df_filtrado["TOTAL_HR"]:
    if not val or val == "" or val in ["00:00:00", "00:00", "0.0", "0"]:
        continue
    if hasattr(val, "hour") and hasattr(val, "minute"):
        total_segundos += (val.hour * 3600) + (val.minute * 60) + getattr(val, "second", 0)
    elif isinstance(val, pd.Timedelta):
        total_segundos += int(val.total_seconds())
    else:
        val_str = str(val).strip()
        if ":" in val_str:
            try:
                if " " in val_str:
                    val_str = val_str.split(" ")[-1]
                partes = val_str.split(":")
                h = int(partes[0])
                m = int(partes[1]) if len(partes) > 1 else 0
                s = int(partes[2]) if len(partes) > 2 else 0
                total_segundos += (h * 3600) + (m * 60) + s
            except (ValueError, IndexError):
                continue

horas_inteiras = int(total_segundos // 3600)
minutos_restantes = int((total_segundos % 3600) // 60)
segundos_restantes = int(total_segundos % 60)
total_horas_faturar = f"{horas_inteiras}:{minutos_restantes:02d}:{segundos_restantes:02d}"

horas_dec = horas_para_decimal(total_horas_faturar)
preco_total_calculado = horas_dec * valor_hora

dados_faturamento["qtd_horas"] = total_horas_faturar
dados_faturamento["preco_total"] = preco_total_calculado

# --- PRÉVIA DOS RESULTADOS ---
st.markdown("### Resumo do Faturamento Calculado da Base de Dados")
col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Total de Horas Encontradas", total_horas_faturar)
col_m2.metric("Preço Unitário da Hora", f"R$ {formatar_br(valor_hora)}")
col_m3.metric("Preço Total Calculado", f"R$ {formatar_br(preco_total_calculado)}")

# --- CLASSE DO PDF REVISADA (SEM RECUO, SEM VERMELHO, SEM CAIXA EMBAIXO) ---
class PDFMedicaoNovo(FPDF):
    def moldura_topo(self, x, y, w, h, dados):
        self.set_draw_color(180, 180, 180)  
        self.set_line_width(0.4)
        self.rect(x, y, w, h)
        self.set_font("Arial", "", 9)
        self.set_text_color(0, 0, 0)
        
        linhas = [
            f"Parceiro:     {dados['parceiro']}",
            f"Endereço:   {dados['endereco']}",
            f"Cidade / UF: {dados['cidade_uf']}",
            f"CEP:           {dados['cep']}",
            f"CNPJ:         {dados['cnpj']}",
        ]
        curr_y = y + 5
        for linha in linhas:
            self.text(x + 4, curr_y, linha)
            curr_y += 5.2
            
        self.line(x + 110, y, x + 110, y + h)
        self.text(x + 114, y + 8, "Medição Número:")
        self.set_font("Arial", "B", 10)
        self.text(x + 160, y + 8, str(dados["numero_medicao"]))
        self.set_font("Arial", "", 9)
        self.text(x + 114, y + 18, f"Período: {dados['data_inicio']} até {dados['data_fim']}")

def gerar_pdf_medicao_nova(dados):
    pdf = PDFMedicaoNovo(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    ARQUIVO_LOGO = "crti.jpg"
    if os.path.exists(ARQUIVO_LOGO):
        pdf.image(ARQUIVO_LOGO, x=150, y=10, w=45)
        
    pdf.set_font("Arial", "B", 15)
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
    pdf.cell(23, 10, formatar_br(dados["preco_total"]), border=1, align="C")
    
    pdf.set_y(91); pdf.set_x(15)
    pdf.cell(28, 7, "", border=0); pdf.set_font("Arial", "B", 8)
    pdf.cell(72, 7, "TOTAL", border=1, align="L", fill=True); pdf.cell(57, 7, "", border=0)
    pdf.cell(23, 7, formatar_br(dados["preco_total"]), border=1, align="C")
    
    pdf.set_text_color(100, 100, 100); pdf.set_font("Arial", "I", 7)
    pdf.text(115, 104, "* Duplicatas a serem emitidas")
    pdf.text(115, 107, f"HPtech Informática ME, valor total de R$ {formatar_br(dados['preco_total'])}")
    pdf.text(115, 110, "Banco: Santander Ag. 0809 CC: 01055895-8")
    pdf.text(115, 113, "Pix: hudsonpedro@gmail.com")
    
    # ASSINATURAS LIMPAS
    pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", "B", 9)
    pdf.text(15, 122, "* De acordo com a Medição Mensal")
    pdf.set_draw_color(180, 180, 180)
    pdf.line(15, 142, 85, 142); pdf.line(125, 142, 195, 142)
    pdf.set_font("Arial", "", 8); pdf.text(15, 146, "HPtech Informática ME"); pdf.text(125, 146, "CR Tecnologia da Informação Ltda")
    return pdf.output(dest="S").encode("latin1")

# --- GERADOR PLANILHA EXCEL CORRIGIDO (FIM DOS CORTES VERTICAIS) ---
def gerar_xlsx_medicao_nova(dados):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet("Medição")
    
    # Oculta as linhas de grade padrão do Excel
    worksheet.hide_gridlines(2)
    
    # DEFINIÇÃO DE ESTILOS E FONTES
    fmt_titulo = workbook.add_format({"bold": True, "size": 15, "font_name": "Arial"})
    fmt_negrito = workbook.add_format({"bold": True, "font_name": "Arial", "size": 10})
    fmt_regular = workbook.add_format({"font_name": "Arial", "size": 9})
    
    # ESTILOS DAS BORDAS EXTERNAS DOS QUADROS (ALINHADOS)
    cor_borda = "#B4B4B4"
    fmt_canto_esq_sup = workbook.add_format({"left": 1, "top": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    fmt_canto_esq_inf = workbook.add_format({"left": 1, "bottom": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    fmt_linha_esq      = workbook.add_format({"left": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    
    fmt_canto_dir_sup = workbook.add_format({"right": 1, "top": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    fmt_canto_dir_inf = workbook.add_format({"right": 1, "bottom": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    fmt_linha_dir      = workbook.add_format({"right": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    
    fmt_tampa_sup     = workbook.add_format({"top": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    fmt_tampa_inf     = workbook.add_format({"bottom": 1, "border_color": cor_borda, "font_name": "Arial", "size": 9})
    fmt_miolo_limpo   = workbook.add_format({"font_name": "Arial", "size": 9})
    
    # ESTILOS DA TABELA DE ITENS
    fmt_header = workbook.add_format({
        "bold": True, "bg_color": "#F5F5F5", "border": 1, "border_color": cor_borda, 
        "align": "center", "valign": "vcenter", "font_name": "Arial", "size": 9
    })
    fmt_celula = workbook.add_format({
        "border": 1, "border_color": cor_borda, "align": "center", "valign": "vcenter", "font_name": "Arial", "size": 9
    })
    fmt_celula_esq = workbook.add_format({
        "border": 1, "border_color": cor_borda, "align": "left", "valign": "vcenter", "font_name": "Arial", "size": 9
    })
    fmt_total_label = workbook.add_format({
        "bold": True, "bg_color": "#F5F5F5", "border": 1, "border_color": cor_borda, 
        "align": "left", "valign": "vcenter", "font_name": "Arial", "size": 9
    })
    fmt_total_valor = workbook.add_format({
        "bold": True, "border": 1, "border_color": cor_borda, "align": "center", "valign": "vcenter", "font_name": "Arial", "size": 9
    })

    # DEFINIÇÃO DE LARGURAS DE COLUNAS EQUILIBRADAS
    worksheet.set_column("A:A", 14)  # Mês/Ano
    worksheet.set_column("B:B", 8)   # Item
    worksheet.set_column("C:C", 48)  # Descrição
    worksheet.set_column("D:D", 10)  # Unidade
    worksheet.set_column("E:E", 14)  # Qtd
    worksheet.set_column("F:F", 16)  # Preço Unitário
    worksheet.set_column("G:G", 18)  # Preço Total
    
    # Título Principal
    worksheet.write("A2", "Medição Mensal de Prestação de Serviços", fmt_titulo)
    
    # Inserção da Logo da CRTI (Fixada no limite direito absoluto)
    ARQUIVO_LOGO = "crti.jpg"
    if os.path.exists(ARQUIVO_LOGO):
        worksheet.insert_image("G1", ARQUIVO_LOGO, {
            "x_scale": 1.85, 
            "y_scale": 1.85, 
            "x_offset": -15, 
            "y_offset": 5
        })
    
    # =========================================================================
    # QUADRO 1: PARCEIRO (COLUNAS A ATÉ D)
    # =========================================================================
    worksheet.write("A4", f"  Parceiro: {dados['parceiro']}", fmt_canto_esq_sup)
    worksheet.write("B4", "", fmt_tampa_sup); worksheet.write("C4", "", fmt_tampa_sup); worksheet.write("D4", "", fmt_canto_dir_sup)
    
    worksheet.write("A5", f"  Endereço: {dados['endereco']}", fmt_linha_esq)
    worksheet.write("B5", "", fmt_miolo_limpo); worksheet.write("C5", "", fmt_miolo_limpo); worksheet.write("D5", "", fmt_linha_dir)
    
    worksheet.write("A6", f"  Cidade / UF: {dados['cidade_uf']}", fmt_linha_esq)
    worksheet.write("B6", "", fmt_miolo_limpo); worksheet.write("C6", "", fmt_miolo_limpo); worksheet.write("D6", "", fmt_linha_dir)
    
    worksheet.write("A7", f"  CEP:       {dados['cep']}", fmt_linha_esq)
    worksheet.write("B7", "", fmt_miolo_limpo); worksheet.write("C7", "", fmt_miolo_limpo); worksheet.write("D7", "", fmt_linha_dir)
    
    worksheet.write("A8", f"  CNPJ:     {dados['cnpj']}", fmt_canto_esq_inf)
    worksheet.write("B8", "", fmt_tampa_inf); worksheet.write("C8", "", fmt_tampa_inf); worksheet.write("D8", "", fmt_canto_dir_inf)

    # =========================================================================
    # QUADRO 2: MEDIÇÃO (COLUNAS E ATÉ G) - CORRIGIDO SEM QUEBRA VERTICAL
    # =========================================================================
    # Linha 4 (Mescla F e G para fechar a ponta direita de forma uniforme)
    worksheet.write("E4", "  Medição Número:", fmt_canto_esq_sup)
    worksheet.write("F4", dados["numero_medicao"], fmt_negrito)
    worksheet.write("G4", "", fmt_canto_dir_sup)
    
    # Linha 5
    worksheet.write("E5", "", fmt_linha_esq); worksheet.write("F5", "", fmt_miolo_limpo); worksheet.write("G5", "", fmt_linha_dir)
    
    # Linha 6 (Texto do período longo avança pelas colunas F e G)
    worksheet.write("E6", f"  Período: {dados['data_inicio']} até {dados['data_fim']}", fmt_linha_esq)
    worksheet.write("F6", "", fmt_miolo_limpo); worksheet.write("G6", "", fmt_linha_dir)
    
    # Linha 7
    worksheet.write("E7", "", fmt_linha_esq); worksheet.write("F7", "", fmt_miolo_limpo); worksheet.write("G7", "", fmt_linha_dir)
    
    # Linha 8
    worksheet.write("E8", "", fmt_canto_esq_inf); worksheet.write("F8", "", fmt_tampa_inf); worksheet.write("G8", "", fmt_canto_dir_inf)
    
    # =========================================================================
    # TABELA DE ITENS EXECUTADOS
    # =========================================================================
    worksheet.write("A10", "* Serviços Executados", fmt_negrito)
    
    headers = ["Mês/Ano", "Item", "Descrição", "Unidade", "Qtd", "Preço Unitário", "Preço Total"]
    for col_idx, text in enumerate(headers):
        worksheet.write(10, col_idx, text, fmt_header)
        
    worksheet.write(11, 0, dados["mes_ano"], fmt_celula)
    worksheet.write(11, 1, "1", fmt_celula)
    worksheet.write(11, 2, dados["descricao_servico"], fmt_celula_esq)
    worksheet.write(11, 3, "HR", fmt_celula)
    worksheet.write(11, 4, dados["qtd_horas"], fmt_celula)
    worksheet.write(11, 5, formatar_br(dados["preco_unitario"]), fmt_celula)
    worksheet.write(11, 6, formatar_br(dados["preco_total"]), fmt_celula)
    
    # Linha do TOTAL perfeitamente alinhada à coluna G
    worksheet.write(12, 0, "", fmt_celula)
    worksheet.write(12, 1, "", fmt_celula)
    worksheet.write(12, 2, "TOTAL", fmt_total_label)
    worksheet.write(12, 3, "", fmt_celula)
    worksheet.write(12, 4, "", fmt_celula)
    worksheet.write(12, 5, "", fmt_celula)
    worksheet.write(12, 6, formatar_br(dados["preco_total"]), fmt_total_valor)
    
    # Notas adicionais de rodapé
    worksheet.write("E14", "* Duplicatas a serem emitidas", workbook.add_format({"italic": True, "size": 7, "font_name": "Arial", "font_color": "#646464"}))
    worksheet.write("E15", f"HP SERVIÇOS ADM, valor total de R$ {formatar_br(dados['preco_total'])}", workbook.add_format({"italic": True, "size": 7, "font_name": "Arial", "font_color": "#646464"}))
    
    # Seção de Assinaturas (Alinhada às margens da folha)
    worksheet.write("A17", "* De acordo com a Medição Mensal", fmt_negrito)
    
    fmt_linha_assinatura = workbook.add_format({"top": 1, "top_color": cor_borda, "align": "left", "font_name": "Arial", "size": 8})
    worksheet.write("A20", "HPtech Informática ME", fmt_linha_assinatura)
    worksheet.write("F20", "CR Tecnologia da Informação Ltda", fmt_linha_assinatura)
    
    workbook.close()
    return output.getvalue()

# --- DISPARO SMTP CORRIGIDO (FIM DO BUG NONAME) ---
def enviar_email_medicao_nova(email_destino, dados, pdf_bytes, xlsx_bytes, nome_pdf, nome_xlsx):
    email_remetente = st.secrets["smtp"]["usuario"]
    senha_remetente = st.secrets["smtp"]["senha"]
    smtp_server = st.secrets["smtp"]["servidor"]
    smtp_porta = int(st.secrets["smtp"]["porta"])
    
    msg = MIMEMultipart()
    msg["From"] = email_remetente
    msg["To"] = email_destino
    
    # O assunto do e-mail assume dinamicamente o nome limpo do arquivo
    msg["Subject"] = str(nome_pdf).replace(".pdf", "")
    
    corpo = f"""<html><body><p>Prezada Sra. Camille Borges, espero que se encontre bem,</p><br>
    <p>Conforme solicitado, segue em anexo a medição para aprovação, autorizando a emissão da NFS-e referente aos serviços de implantação no período de ({dados['data_inicio']} até {dados['data_fim']}).</p><br>
    <p>De acordo como já informado ao setor financeiro as cuidados da Sra. Amanda, informo que a minha conta do Itaú está em processo de encerramento, agora o meu pix é hudsonpedro@gmail.com, favor realizar os depósitos na minha conta Santander Ag. 0809 cc 01055895-8.</p><br><br>
    <p>Com gratidão!<br><br>Hudson Valente</p></body></html>"""
    msg.attach(MIMEText(corpo, "html"))
    
    # Lista estruturada mapeando os bytes com seus respectivos nomes dinâmicos passados pela UI
    lote_anexos = [
        (pdf_bytes, nome_pdf),
        (xlsx_bytes, nome_xlsx)
    ]
    
    for b_data, nome_arquivo in lote_anexos:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(b_data)
        encode_base64(part)
        
        # O uso de aspas duplas escapadas \" garante que o Gmail leia nomes com espaços e símbolos sem quebrar
        part.add_header("Content-Disposition", f"attachment; filename=\"{nome_arquivo}\"")
        msg.attach(part)
        
    try:
        server = smtplib.SMTP(smtp_server, smtp_porta)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, email_destino, msg.as_string())
        server.quit()
        return True, "E-mail enviado com sucesso!"
    except Exception as e: 
        return False, f"Falha no envio SMTP: {str(e)}"

# --- BOTÕES VISUAIS E EXPORTAÇÃO DIRETAL ---
st.markdown("---")
col_b1, col_b2 = st.columns(2)

# Nomenclatura oficial padronizada sem risco de quebra de codificação
# Trocamos o caractere '°' por 'N' ou espaço para total compatibilidade com o servidor Linux
nome_pdf = f"Medicao N {dados_faturamento['numero_medicao']} - {dados_faturamento['mes_ano']} (Implantacao) - HUDSON.pdf"
nome_xlsx = f"Medicao N {dados_faturamento['numero_medicao']} - {dados_faturamento['mes_ano']} (Implantacao) - HUDSON.xlsx"

with col_b1:
    st.download_button(
        label=" Baixar PDF da Medição", 
        data=gerar_pdf_medicao_nova(dados_faturamento), 
        file_name=nome_pdf, 
        mime="application/pdf", 
        use_container_width=True
    )
with col_b2:
    st.download_button(
        label=" Baixar Excel da Medição", 
        data=gerar_xlsx_medicao_nova(dados_faturamento), 
        file_name=nome_xlsx, 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        use_container_width=True
    )

# --- BOTÃO PRINCIPAL QUE ABRE O POPUP ---
st.markdown("---")
email_target = st.text_input("Destinatário da Medição:", "suellen@crti.com.br")

# 1. Definição da função do Pop-up (st.dialog)
import time # Garante a importação do controle de tempo no script

# 1. Definição da função do Pop-up com persistência de mensagem
@st.dialog("⚠️ Confirmação de Envio")
def confirmar_envio_popup(email, dados):
    st.write(f"Você tem certeza que deseja enviar a medição para o e-mail **{email}**?")
    st.write(f"**Total de Horas:** {dados['qtd_horas']} | **Valor Total:** R$ {formatar_br(dados['preco_total'])}")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("✅ Sim, Enviar", use_container_width=True):
            with st.spinner("Compilando anexos formatados e enviando..."):
                p_b = gerar_pdf_medicao_nova(dados)
                x_b = gerar_xlsx_medicao_nova(dados)
                ok, r_msg = enviar_email_medicao_nova(email, dados, p_b, x_b, nome_pdf, nome_xlsx)
                
            if ok:
                st.success(f"🎉 {r_msg}") # Exibe o alerta verde na tela
                st.balloons()            # Sobe os balões de celebração
                time.sleep(10)            # Segura a mensagem na tela por 10 segundos antes de fechar
            else:
                st.error(r_msg)
                time.sleep(4)            # Dá mais tempo para ler caso ocorra algum erro
            
            # Atualiza a página e fecha o pop-up automaticamente
            st.rerun()
            
    with col_p2:
        if st.button("❌ Não, Cancelar", use_container_width=True):
            st.rerun()

# 2. Gatilho para abrir o pop-up na tela (Mantém igual)
if st.button("🚀 Enviar Medição por E-mail", use_container_width=True):
    if not email_target:
        st.error("Por favor, preencha o e-mail do destinatário.")
    else:
        confirmar_envio_popup(email_target, dados_faturamento)
