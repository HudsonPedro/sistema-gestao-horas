#Por Hudson Valente - HPTECH
#em: 27/04/2026 - 19:55h
import streamlit as st
import base64 #==novo imagem ao lado no título ===#

# 1. CONFIGURAÇÃO DA PÁGINA (Sempre a primeira linha de código!)
st.set_page_config(
    page_title="Gerador de Relatórios HPTECH", 
    page_icon="hptechICO.png", 
    layout="wide"
)

st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* O BLOCO PADRÃO*/
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
        
        /* Cor do título para o padrão vermelho*/
        h1 { color: #b0231d; } 
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


# 3. SIDEBAR COM NOVO BOTÃO
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    st.markdown("---")
    # Identificação do Usuário
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
    if st.button("📄 Relatórios", use_container_width=True):
        st.switch_page("pages/02_📄_Relatorios.py")
    if st.button("📝 Lançamento", use_container_width=True):
        st.switch_page("pages/03_📝_Lancamento.py")
    if st.button("💰 Medição Mensal", use_container_width=True):
        st.switch_page("pages/04_💰_Medicao_Mensal.py")
  
#   versionamento 
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")
    
# 4. IMPORTS PESADOS
import os
import glob
from datetime import timedelta, time, datetime
import pandas as pd
import locale
from fpdf import FPDF
import xlsxwriter 
from PIL import Image
import shutil 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

if "relatorios_gerados" not in st.session_state:
    st.session_state.relatorios_gerados = False

def data_por_extenso_pt(dt):
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    dia = dt.day
    mes = meses[dt.month]
    ano = dt.year
    return f"{dia:02d} de {mes} de {ano}"

# =========================================================================
# AQUI ENTRA O PASSO 2: COLE A FUNÇÃO DE LOG EXATAMENTE NESTE ESPAÇO
# =========================================================================
def registrar_historico_sheets(tipo, identificador, destinatario):
    """Grava o log do envio com sucesso na aba HISTORICO_ENVIOS do Google Sheets"""
    import gspread
    from google.oauth2.service_account import Credentials
    from datetime import datetime
    
    try:
        escopos = ["googleapis.com"]
        credenciais = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=escopos)
        cliente_sheets = gspread.authorize(credenciais)
        
        id_planilha = "1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE"
        planilha = cliente_sheets.open_by_key(id_planilha)
        aba_log = planilha.worksheet("HISTORICO_ENVIOS")
        
        data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        usuario_atual = st.user.get("email") or "hudson.valente@crti.com.br"
        
        nova_linha = [data_atual, tipo, str(identificador), destinatario, usuario_atual]
        aba_log.append_row(nova_linha)
    except Exception as e:
        st.sidebar.error(f"Erro ao registrar log de auditoria: {e}")
# =========================================================================

PASTA_SAIDA = "relatorios"
LOGO = "crti.jpg"
os.makedirs(PASTA_SAIDA, exist_ok=True)

class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.ra_numero = None

    def header(self):
        if os.path.exists(LOGO):
            self.image(LOGO, x=160, y=10, w=40)
        self.set_font("Arial", "B", 12)
        ra_mostrar = str(self.ra_numero) if self.ra_numero != 0 else "S/N"
        self.cell(0, 15, f"RELATÓRIO DE ATENDIMENTO Nº {ra_mostrar}", ln=True, align="L")

def enviar_relatorio_email(arquivos_anexos, servidor_smtp, porta, email_remetente, senha, destinatario):
    if not arquivos_anexos: 
        return False, "Nenhum arquivo para anexar."
    
    if isinstance(arquivos_anexos, list):
        primeiro_arquivo = arquivos_anexos[0] 
    else:
        primeiro_arquivo = arquivos_anexos
    
    nome_base = os.path.basename(primeiro_arquivo).replace(".pdf", "").replace(".xlsx", "").strip()
    
    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = destinatario
    msg['Subject'] = f"{nome_base} - HUDSON VALENTE"

    corpo = f"""Prezada Sra. Amanda, espero que se encontre bem.

Segue em anexo o {nome_base} (em formatos PDF e Excel) para análise e assinatura.

Atenciosamente,

Hudson Valente"""
    
    msg.attach(MIMEText(corpo, 'plain'))
    
    arquivos_para_enviar = arquivos_anexos if isinstance(arquivos_anexos, list) else [arquivos_anexos]
    
    for caminho_arquivo in arquivos_para_enviar:
        nome_arquivo_original = os.path.basename(caminho_arquivo)
        try:
            with open(caminho_arquivo, "rb") as attachment:
                if nome_arquivo_original.lower().endswith(".pdf"):
                    part = MIMEBase('application', 'pdf')
                elif nome_arquivo_original.lower().endswith(".xlsx"):
                    part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=nome_arquivo_original)
            msg.attach(part)
        except Exception as e:
            return False, f"Erro ao anexar {nome_arquivo_original}: {str(e)}"
    
    try:
        server = smtplib.SMTP(servidor_smtp, porta)
        server.starttls()
        server.login(email_remetente, senha)
        server.sendmail(email_remetente, destinatario, msg.as_string())
        server.quit()
        return True, f"✅ Enviado com sucesso: {nome_base}"
    except Exception as e:
        return False, f"❌ Erro de conexão no envio: {str(e)}"

# ====== STREAMLIT UI ======
st.set_page_config(page_title="Gerador de Relatórios HPTECH", page_icon="📄", layout="wide")

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
            <h1 style="margin: 0; font-size: 2.5rem;">Gerador Automático de Relatórios</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("🔥 Gerador Automático de Relatórios HPTECH")

st.markdown("---")

# --- LÊ A PLANILHA TODA (TODAS AS ABAS) ---
@st.cache_data(ttl=600) 
def carregar_planilha_todas_abas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    dict_abas = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    return dict_abas

st.sidebar.header("⚙️ Configurações GERAIS")

if st.sidebar.button("🔄 Atualizar Base de Dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

with st.spinner("⏳ Analisando Dados..."):
    try:
        dict_abas = carregar_planilha_todas_abas()
        abas_disponiveis = list(dict_abas.keys())
    except Exception as e:
        st.error(f"❌ Erro ao baixar planilha: Verifique se possui a biblioteca openpyxl instalada. Erro: {e}")
        st.stop()
        
st.sidebar.markdown("### 📂 Seleção da Base")

#====== NOVO CARIMBAR RELATORIO JA GERADOS E LIMPAR ===========
def limpar_estado():
    st.session_state.relatorios_gerados = False

aba_selecionada = st.sidebar.selectbox("**Selecione o Mês:**", abas_disponiveis, on_change=limpar_estado)

# --- DADOS DA ABA SELECIONADA ---
df_completo = dict_abas[aba_selecionada].copy()
df_completo["DATA"] = pd.to_datetime(df_completo["DATA"], errors="coerce", dayfirst=True)

# Descobre a menor e maior data presentes SOMENTE na aba selecionada
min_data_aba = df_completo["DATA"].min()
max_data_aba = df_completo["DATA"].max()
# Se a aba estiver vazia, define um padrão de segurança
if pd.isnull(min_data_aba): min_data_aba = datetime(2026, 1, 1)
if pd.isnull(max_data_aba): max_data_aba = datetime(2026, 1, 31)

st.sidebar.markdown("### 📅 Filtro de Datas")
data_inicio_selecionada = st.sidebar.date_input("**Data Início**", value=pd.to_datetime(min_data_aba).date())
data_fim_selecionada = st.sidebar.date_input("**Data Fim**", value=pd.to_datetime(max_data_aba).date())

btn_gerar = st.sidebar.button("🚀 **GERAR RELATÓRIOS**", type="primary", use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.header("📧 Disparo de E-mails")
email_destinatario = st.sidebar.text_input("Enviar para (Destinatário):", value="financeiro@crti.com.br")
senha_app = st.sidebar.text_input("Senha App Gmail:", value="fzau tvih zlsn xadi", type="password")
btn_enviar_emails = st.sidebar.button("📨 **ENVIAR POR E-MAIL**", type="primary", use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.header("🗑️ Gerenciamento de Arquivos")
if st.sidebar.button("🗑️ **Limpar Relatórios Antigos**", type="secondary", use_container_width=True):
    arquivos_para_remover = glob.glob(os.path.join(PASTA_SAIDA, "*.pdf")) + glob.glob(os.path.join(PASTA_SAIDA, "*.xlsx"))
    if not arquivos_para_remover:
        st.sidebar.warning("ℹ️ Nenhum relatório para remover.")
    else:
        for arquivo in arquivos_para_remover: os.remove(arquivo)
        st.sidebar.success("✅ Relatórios removidos!")
        st.rerun()  

col1, col2 = st.columns([2.5, 1])
with col1:
    # ==== TABELA DETALHADA ====
    st.subheader(f"📊 Prévia dos Dados do Mês: {aba_selecionada}")
    st.dataframe(df_completo.head(125), use_container_width=True)

with col2:
    st.subheader("📈 Status Atual")
    st.metric(f"Linhas em {aba_selecionada}", len(df_completo))
    st.info("1º Selecione o Mês.\n\n2º Selecione as datas.\n\n3º Clique em GERAR RELATÓRIOS.")

# ==========================================
# GERAÇÃO DOS RELATÓRIOS 
# ==========================================
if btn_gerar:
    st.session_state.relatorios_gerados = True 
    st.markdown("---")
    resumo_status = st.empty() 
    
    with st.spinner("🎯 Gerando arquivos PDF e Excel idênticos..."):
        df = df_completo.copy()
        
        data_ini_pd = pd.to_datetime(data_inicio_selecionada)
        data_fim_pd = pd.to_datetime(data_fim_selecionada)
        df = df[(df["DATA"] >= data_ini_pd) & (df["DATA"] <= data_fim_pd)]

        if df.empty:
            st.error("⚠️ Nenhum registro encontrado para estas datas nesta aba.")
            st.stop()

        cols_obr = ["CLIENTE", "OBSERVAÇÕES", "CONSULTOR", "SOLICITANTE", "PARTICIPANTE", "FORMA", "RA", "LOCAL", "SITUACAO_RA", "HR_INICIO_D", "HR_FIM_D", "TOTAL_HR_D", "KM_D", "FORMA_D", 'DESCRICAO']
        for col in cols_obr:
            if col not in df.columns: df[col] = ""
            else: df[col] = df[col].fillna("").astype(str)

        df = df[df["RA"].str.strip() != ""]
        df = df[df["CLIENTE"].str.strip() != ""]
        df = df[df["SITUACAO_RA"].str.strip() == "Em Elaboração"]

        grupos = df.groupby(["CLIENTE", "RA"], as_index=False)
        arquivos_saida = []

        for (cliente, ra), grupo in grupos:
            solicitante = str(grupo["SOLICITANTE"].iloc[0]).strip()
            consultor = str(grupo["CONSULTOR"].iloc[0]).strip()
            participante_padrao = str(grupo["PARTICIPANTE"].iloc[0]).strip()
            local = str(grupo["LOCAL"].iloc[0]).strip()

            data_inicio_rel = grupo["DATA"].min().strftime("%d/%m/%Y")
            data_fim_rel = grupo["DATA"].max().strftime("%d/%m/%Y")
            dt_obj = grupo["DATA"].max().to_pydatetime()
            data_rodape = data_por_extenso_pt(dt_obj)

            # ------ CÁLCULO DE HORAS ------
            total_seg, total_seg_d = 0, 0
            
            for val in grupo["TOTAL_HR"]:
                val_str = str(val).strip() 
                if ":" in val_str:
                    try: 
                        p = val_str.split(":")
                        total_seg += int(p[0]) * 3600 + (int(p[1]) * 60 if len(p) > 1 else 0)
                    except: pass
            total_hr_str = f"{int(total_seg // 3600):02d}:{int((total_seg % 3600) // 60):02d}"

            for val in grupo["TOTAL_HR_D"]:
                val_str = str(val).strip()
                if ":" in val_str:
                    try: 
                        p = val_str.split(":")
                        total_seg_d += int(p[0]) * 3600 + (int(p[1]) * 60 if len(p) > 1 else 0)
                    except: pass
            total_hr_str_d = f"{int(total_seg_d // 3600):02d}:{int((total_seg_d % 3600) // 60):02d}"
            # -----------------------------------------------

            total_km = 0
            tem_desl = False
            for val in grupo["KM_D"]:
                v = str(val).strip().replace(',', '.')
                if v not in ["", "nan", "None", "0", "0.0"]:
                    try: 
                        kv=float(v); total_km += kv; tem_desl = True if kv > 0 else False
                    except: pass

            texto_ra = str(ra).strip()
            ra_str = str(int(float(texto_ra))) if texto_ra != "" and texto_ra.replace('.', '', 1).isdigit() else "S/N"
            
            nome_base = f"RA Nº {ra_str} {cliente[:25].upper().replace(' ', ' ')}"
            file_pdf = os.path.join(PASTA_SAIDA, nome_base + ".pdf")
            file_xlsx = os.path.join(PASTA_SAIDA, nome_base + ".xlsx")

            # --- 1. GERAÇÃO PDF ---
            pdf = PDF()
            pdf.ra_numero = ra_str
            pdf.add_page()
            pdf.set_font("Arial", "BU", 10); pdf.cell(105, 5, "DADOS GERAIS:", ln=False); pdf.cell(0, 5, "RESUMO DO RELATÓRIO:", ln=True)
            pdf.set_font('Arial', 'B', 10); pdf.cell(14, 5, "Cliente: ", ln=False); pdf.set_font("Arial", "", 8); pdf.cell(0, 5, cliente, ln=False)
            pdf.set_x(115); pdf.set_font('Arial', 'B', 10); pdf.cell(39, 5, "Período: ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(0, 5, f"{data_inicio_rel} até {data_fim_rel}", ln=True)
            pdf.set_x(10); pdf.set_font('Arial', 'B', 10); pdf.cell(20, 5, "Solicitante: ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(0, 5, solicitante, ln=False)
            pdf.set_x(114.5); pdf.set_font('Arial', 'B', 10); pdf.cell(39.5, 5, "Total de Horas: ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(0, 5, total_hr_str, ln=True)
            pdf.set_x(10); pdf.set_font('Arial', 'B', 10); pdf.cell(38, 5, "Tipo de Atendimento: ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(66.5, 5, "Implantação do Sistema CRTI ERP", ln=False)
            pdf.set_x(114.5); pdf.set_font('Arial', 'B', 10); pdf.cell(39.5, 5, "Total Deslocamento: ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(0, 5, total_hr_str_d, ln=True)
            pdf.set_x(10); pdf.set_font('Arial', 'B', 10); pdf.cell(45, 5, "Unidade de Atendimento: ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(60, 5, local, ln=False)
            pdf.set_x(115); pdf.set_font('Arial', 'B', 10); pdf.cell(39, 5, "Distância (KM): ", ln=False); pdf.set_font("Arial", "", 10); pdf.cell(0, 5, f"{total_km} km", ln=True)
            pdf.ln(5); pdf.set_font("Arial", "B", 10); pdf.set_fill_color(0, 112, 192); pdf.set_text_color(255, 255, 255); pdf.cell(190, 10, "DESCRIÇÃO DAS ATIVIDADES", border=1, ln=True, fill=True, align="C"); pdf.set_text_color(0, 0, 0)

            for _, linha in grupo.iterrows():
                if pdf.get_y() > 245: pdf.add_page()
                y_i = pdf.get_y(); x_i = 10
                da = pd.to_datetime(linha["DATA"]).strftime("%d/%m/%Y") if pd.notnull(linha["DATA"]) else ""
                hi = str(linha["HR_INICIO"])[0:5] if pd.notnull(linha["HR_INICIO"]) else "00:00"
                hf = str(linha["HR_FIM"])[0:5] if pd.notnull(linha["HR_FIM"]) else "00:00"
                tt = str(linha["TOTAL_HR"])[0:5] if pd.notnull(linha["TOTAL_HR"]) else "00:00"
                ob = str(linha["OBSERVAÇÕES"]).strip()

                pdf.set_xy(x_i + 2, y_i + 2)
                pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Data: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(30, 5, da, ln=False)
                pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Hora Início: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(15, 5, hi, ln=False)
                pdf.set_font('Arial', 'B', 10); pdf.cell(20, 5, "Hora Final: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(38, 5, hf, ln=False)
                pdf.set_font('Arial', 'B', 10); pdf.cell(30, 5, "Total: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(0, 5, tt, ln=True)
                pdf.set_x(x_i + 2); pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Consultor: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(67, 5, consultor, ln=False)
                pdf.set_font('Arial', 'B', 10); pdf.cell(42, 5, "Forma de Atendimento: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(0, 5, str(linha.get("FORMA", "Remoto")).strip(), ln=True)
                pdf.set_x(x_i + 2); pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Atividade: ", ln=False); pdf.set_font('Arial', '', 10); pdf.multi_cell(0, 5, ob if ob else "-")
                pdf.set_x(x_i + 2); pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Participante: ", ln=False); pdf.set_font('Arial', '', 10); pdf.multi_cell(0, 5, str(linha["PARTICIPANTE"]).strip() or participante_padrao)
                
                pdf.set_y(pdf.get_y() + 2); pdf.rect(x_i, y_i, 190, pdf.get_y() - y_i)

            if tem_desl:
                if pdf.get_y() > 220: pdf.add_page()
                pdf.ln(2)
                pdf.set_font("Arial", "B", 10); pdf.set_fill_color(0, 112, 192); pdf.set_text_color(255, 255, 255)
                pdf.cell(190, 10, "DESLOCAMENTOS", border=1, ln=True, fill=True, align="C")
                pdf.set_text_color(0, 0, 0)
                for _, linha in grupo.iterrows():
                    km_s = str(linha.get("KM_D", "")).strip().replace(',', '.')
                    if km_s in ["", "nan", "None", "0", "0.0"]: continue 
                    if pdf.get_y() > 245: pdf.add_page()
                    y_i = pdf.get_y(); x_i = 10
                    dd = pd.to_datetime(linha["DATA"]).strftime("%d/%m/%Y") if pd.notnull(linha["DATA"]) else ""
                    hi_d = str(linha["HR_INICIO_D"])[0:5] if pd.notnull(linha["HR_INICIO_D"]) else "00:00"
                    hf_d = str(linha["HR_FIM_D"])[0:5] if pd.notnull(linha["HR_FIM_D"]) else "00:00"
                    tt_d = str(linha["TOTAL_HR_D"])[0:5] if pd.notnull(linha["TOTAL_HR_D"]) else "00:00"
                    ds_d = str(linha.get("DESCRICAO_D", "")).strip()
                    fm_d = str(linha.get("FORMA_D", "Carro Próprio")).strip()

                    pdf.set_xy(x_i + 2, y_i + 2)
                    pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Data: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(30, 5, dd, ln=False)
                    pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Hora Início: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(15, 5, hi_d, ln=False)
                    pdf.set_font('Arial', 'B', 10); pdf.cell(20, 5, "Hora Final: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(38, 5, hf_d, ln=False)
                    pdf.set_font('Arial', 'B', 10); pdf.cell(30, 5, "Total: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(0, 5, tt_d, ln=True)
                    pdf.set_x(x_i + 2); pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Distância: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(30, 5, f"{km_s} km", ln=False)
                    pdf.set_font('Arial', 'B', 10); pdf.cell(45, 5, "Forma de Deslocamento: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(30, 5, fm_d, ln=False)
                    pdf.set_font('Arial', 'B', 10); pdf.cell(20, 5, "Consultor: ", ln=False); pdf.set_font('Arial', '', 10); pdf.cell(0, 5, consultor, ln=True)
                    pdf.set_x(x_i + 2); pdf.set_font('Arial', 'B', 10); pdf.cell(22, 5, "Descrição: ", ln=False); pdf.set_font('Arial', '', 10); pdf.multi_cell(0, 5, ds_d if ds_d else "-")
                    pdf.set_y(pdf.get_y() + 2); pdf.rect(x_i, y_i, 190, pdf.get_y() - y_i)

            if pdf.get_y() > 220: pdf.add_page()
            pdf.ln(5); pdf.cell(0, 8, f"Curitiba, {data_rodape}.", ln=True); pdf.ln(4)
            pdf.set_text_color(255, 0, 0); pdf.multi_cell(0, 5, "As horas referentes aos atendimentos e despesas de viagens serão faturadas conforme acerto prévio. Declaro que os serviços descritos neste relatório foram prestados e confirmado como aceitos.", align="C"); pdf.set_text_color(0, 0, 0)
            pdf.ln(10)
            pdf.cell(90, 8, "__________________________________", align="C"); pdf.cell(10); pdf.cell(90, 8, "__________________________________", align="C", ln=True)
            pdf.cell(90, 2, consultor, align="C"); pdf.cell(7); pdf.cell(90, 2, solicitante, align="C", ln=True)
            pdf.set_font("Arial", "", 8); pdf.cell(90, 7, "CRTI", align="C"); pdf.cell(7); pdf.cell(90, 7, cliente, align="C", ln=True)
            pdf.cell(90, 2, f"RELATÓRIO DE ATENDIMENTO Nº {ra_str}", align="C"); pdf.cell(7); pdf.cell(90, 2, f"RELATÓRIO DE ATENDIMENTO Nº {ra_str}", ln=True, align="C")
            pdf.output(file_pdf)
            arquivos_saida.append(file_pdf)

            # ========================================================
            # 2. GERAÇÃO EXCEL
            # ========================================================
            wb = xlsxwriter.Workbook(file_xlsx)
            ws = wb.add_worksheet("Relatório")
            
            ws.set_paper(9) 
            ws.fit_to_pages(1, 0) 
            ws.set_margins(left=0.2, right=0.2, top=0.4, bottom=0.4) 
            
            ws.set_column('A:A', 13); ws.set_column('B:B', 15); ws.set_column('C:C', 13); ws.set_column('D:D', 11)  
            ws.set_column('E:E', 13); ws.set_column('F:F', 13); ws.set_column('G:G', 8); ws.set_column('H:H', 11)  
            
            f_title = wb.add_format({'bold': True, 'font_size': 13})
            f_hdr_u = wb.add_format({'bold': True, 'underline': True, 'font_size': 10})
            f_bold = wb.add_format({'bold': True, 'font_size': 10, 'valign': 'vcenter'})
            f_norm = wb.add_format({'font_size': 10, 'valign': 'vcenter'}) 
            f_blue = wb.add_format({'bold': True, 'bg_color': '#0070c0', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            f_red = wb.add_format({'font_color': 'red', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
            f_center = wb.add_format({'align': 'center', 'valign': 'top', 'font_size': 10})
            f_sign = wb.add_format({'top': 1, 'align': 'center', 'font_size': 10, 'valign': 'top'})
            
            f_TL = wb.add_format({'top': 1, 'left': 1, 'bold': True, 'valign': 'vcenter'})
            f_T  = wb.add_format({'top': 1, 'valign': 'vcenter'})
            f_T_b= wb.add_format({'top': 1, 'bold': True, 'valign': 'vcenter'})
            f_TR = wb.add_format({'top': 1, 'right': 1, 'valign': 'vcenter'})
            f_ML = wb.add_format({'left': 1, 'bold': True, 'valign': 'top'})
            f_M  = wb.add_format({'valign': 'top'})
            f_M_b= wb.add_format({'bold': True, 'valign': 'top'})
            f_merge_mid = wb.add_format({'valign': 'top', 'text_wrap': True, 'right': 1})
            f_merge_bot = wb.add_format({'bottom': 1, 'valign': 'top', 'text_wrap': True, 'right': 1})
            f_BL = wb.add_format({'bottom': 1, 'left': 1, 'bold': True, 'valign': 'top'})
            
            if os.path.exists(LOGO):
                ws.insert_image('F1', LOGO, {'x_offset': 10, 'y_offset': 2, 'x_scale': 2.00, 'y_scale': 2.00})
            
            ws.write('A2', f"RELATÓRIO DE ATENDIMENTO Nº {ra_str}", f_title)
            ws.write('A4', "DADOS GERAIS:", f_hdr_u)
            ws.write('F4', "RESUMO DO RELATÓRIO:", f_hdr_u) 
            
            ws.write('A5', "Cliente:", f_bold); ws.merge_range('B5:D5', cliente[:45], f_norm)
            ws.write('F5', "Período:", f_bold); ws.merge_range('G5:H5', f"{data_inicio_rel} até {data_fim_rel}", f_norm)
            ws.write('A6', "Solicitante:", f_bold); ws.merge_range('B6:D6', solicitante, f_norm)
            ws.write('F6', "Total de Horas:", f_bold); ws.merge_range('G6:H6', total_hr_str, f_norm)
            ws.write('A7', "Tipo de Atend.:", f_bold); ws.merge_range('B7:D7', "Implantação CRTI", f_norm)
            ws.write('F7', "Deslocamento:", f_bold); ws.merge_range('G7:H7', total_hr_str_d, f_norm)
            ws.write('A8', "Unidade:", f_bold); ws.merge_range('B8:D8', local, f_norm)
            ws.write('F8', "Distância (KM):", f_bold); ws.merge_range('G8:H8', f"{total_km:.2f} km".replace('.', ','), f_norm)

            row = 10
            ws.merge_range(f'A{row}:H{row}', "DESCRIÇÃO DAS ATIVIDADES", f_blue); ws.set_row(row - 1, 18); row += 1

            for _, linha in grupo.iterrows():
                d = pd.to_datetime(linha["DATA"]).strftime("%d/%m/%Y") if pd.notnull(linha["DATA"]) else ""
                ob = str(linha["OBSERVAÇÕES"]).strip()
                fo = str(linha.get("FORMA", "Remoto")).strip()
                pa = str(linha["PARTICIPANTE"]).strip() or participante_padrao
                hi = str(linha["HR_INICIO"])[0:5] if pd.notnull(linha["HR_INICIO"]) else "00:00"
                hf = str(linha["HR_FIM"])[0:5] if pd.notnull(linha["HR_FIM"]) else "00:00"
                tt = str(linha["TOTAL_HR"])[0:5] if pd.notnull(linha["TOTAL_HR"]) else "00:00"

                ws.write(f'A{row}', "Data:", f_TL); ws.write(f'B{row}', d, f_T); ws.write(f'C{row}', "Hora Início:", f_T_b)
                ws.write(f'D{row}', hi, f_T); ws.write(f'E{row}', "Hora Final:", f_T_b); ws.write(f'F{row}', hf, f_T)
                ws.write(f'G{row}', "Total:", f_T_b); ws.write(f'H{row}', tt, f_TR); row += 1
                
                ws.write(f'A{row}', "Consultor:", f_ML); ws.merge_range(f'B{row}:D{row}', consultor, f_M)
                ws.merge_range(f'E{row}:F{row}', "Forma de Atendimento:", f_M_b); ws.merge_range(f'G{row}:H{row}', fo, f_merge_mid); row += 1
                
                linhas_obs = max(1, len(ob) // 90 + 1)
                ws.write(f'A{row}', "Atividade:", f_ML); ws.merge_range(f'B{row}:H{row}', ob if ob else "-", f_merge_mid)
                ws.set_row(row - 1, 15 * linhas_obs); row += 1
                
                linhas_pa = max(1, len(pa) // 90 + 1)
                ws.write(f'A{row}', "Participante:", f_BL); ws.merge_range(f'B{row}:H{row}', pa if pa else "-", f_merge_bot)
                ws.set_row(row - 1, 15 * linhas_pa); row += 1

            if tem_desl:
                ws.merge_range(f'A{row}:H{row}', "DESLOCAMENTOS", f_blue); ws.set_row(row - 1, 18); row += 1
                for _, linha in grupo.iterrows():
                    km_s = str(linha.get("KM_D", "")).strip().replace(',', '.')
                    if km_s in ["", "nan", "None", "0", "0.0"]: continue 
                    
                    dd = pd.to_datetime(linha["DATA"]).strftime("%d/%m/%Y") if pd.notnull(linha["DATA"]) else ""
                    hi_d = str(linha["HR_INICIO_D"])[0:5] if pd.notnull(linha["HR_INICIO_D"]) else "00:00"
                    hf_d = str(linha["HR_FIM_D"])[0:5] if pd.notnull(linha["HR_FIM_D"]) else "00:00"
                    tt_d = str(linha["TOTAL_HR_D"])[0:5] if pd.notnull(linha["TOTAL_HR_D"]) else "00:00"
                    ds_d = str(linha.get("DESCRICAO_D", "")).strip()
                    fm_d = str(linha.get("FORMA_D", "Carro Próprio")).strip()

                    ws.write(f'A{row}', "Data:", f_TL); ws.write(f'B{row}', dd, f_T); ws.write(f'C{row}', "Hora Início:", f_T_b); ws.write(f'D{row}', hi_d, f_T)
                    ws.write(f'E{row}', "Hora Final:", f_T_b); ws.write(f'F{row}', hf_d, f_T); ws.write(f'G{row}', "Total:", f_T_b); ws.write(f'H{row}', tt_d, f_TR); row += 1
                    
                    ws.write(f'A{row}', "Distância:", f_ML); ws.write(f'B{row}', f"{km_s} km", f_M)
                    ws.merge_range(f'C{row}:D{row}', "Forma Desloc.:", f_M_b); ws.merge_range(f'E{row}:F{row}', fm_d, f_M)
                    ws.write(f'G{row}', "Consultor:", f_M_b); ws.write(f'H{row}', consultor, f_merge_mid); row += 1
                    
                    linhas_ds = max(1, len(ds_d) // 90 + 1)
                    ws.write(f'A{row}', "Descrição:", f_BL); ws.merge_range(f'B{row}:H{row}', ds_d if ds_d else "-", f_merge_bot)
                    ws.set_row(row - 1, 15 * linhas_ds); row += 1

            row += 2
            
            ws.merge_range(f'A{row}:D{row}', f"Curitiba, {data_rodape}.", f_norm); row += 2
            
            ws.merge_range(f'A{row}:H{row}', "As horas referentes aos atendimentos e despesas de viagens serão faturadas conforme acerto prévio. Declaro que os serviços descritos neste relatório foram prestados e confirmado como aceitos.", f_red)
            ws.set_row(row - 1, 30); row += 5
            
            ws.merge_range(f'A{row}:C{row}', consultor, f_sign); ws.merge_range(f'F{row}:H{row}', solicitante, f_sign); row+=1
            ws.merge_range(f'A{row}:C{row}', "CRTI", f_center); ws.merge_range(f'F{row}:H{row}', cliente[:30], f_center); row+=1
            ws.merge_range(f'A{row}:C{row}', f"RELATÓRIO DE ATENDIMENTO Nº {ra_str}", f_center); ws.merge_range(f'F{row}:H{row}', f"RELATÓRIO DE ATENDIMENTO Nº {ra_str}", f_center)

            ws.hide_gridlines(2)
            wb.close()
            arquivos_saida.append(file_xlsx)

        # ========================================================
        # --- ARQUIVO ZIP ---
        # ========================================================
        resumo_status.success(f"✅ Feito! Foram gerados **{len(arquivos_saida)}** arquivos da aba '{aba_selecionada}'.")
        
        # Cria um nome legal para o arquivo ZIP
        nome_zip = f"Relatorios_{aba_selecionada.replace(' ', '_')}"
        caminho_zip = f"{nome_zip}.zip"
        
        # O Python compacta a pasta 'relatorios' todinha dentro desse ZIP
        shutil.make_archive(nome_zip, 'zip', PASTA_SAIDA)
        
        st.markdown("### 📥 Download dos Arquivos")
        
        # Botão GIgante para baixar o ZIP
        with open(caminho_zip, "rb") as f_zip:
             st.download_button(
                 label="📦 BAIXAR TODOS OS RELATÓRIOS (ZIP)",
                 data=f_zip,
                 file_name=caminho_zip,
                 mime="application/zip",
                 type="primary",
                 use_container_width=True
             )
             
        # Coloquei os arquivos soltos dentro de um botão que abre e fecha pra não poluir a tela
        with st.expander("Ver arquivos individualmente..."):
            cols_dw = st.columns(3)
            for i, path in enumerate(arquivos_saida):
                nome_arq = os.path.basename(path)
                with open(path, "rb") as file:
                    ext = "📄 PDF" if ".pdf" in nome_arq.lower() else "📊 Excel"
                    cols_dw[i % 3].download_button(label=f"Baixar {ext}: {nome_arq[:15]}...", data=file, file_name=nome_arq, key=path)
# Verifica se os relatórios já foram gerados nesta sessão NOVOOOOOOOOOOOOOOOO
if st.session_state.relatorios_gerados:
    # Busca os arquivos na pasta para garantir que eles ainda existem
    arquivos_pasta = glob.glob(os.path.join(PASTA_SAIDA, "*.*"))
    arquivos_exibicao = [f for f in arquivos_pasta if f.endswith(".pdf") or f.endswith(".xlsx")]
    
    if arquivos_exibicao:
        st.subheader("📥 Baixar Relatórios Gerados")
        cols_dw = st.columns(3)
        for i, path in enumerate(arquivos_exibicao):
            nome_arq = os.path.basename(path)
            with open(path, "rb") as file:
                ext = " PDF" if ".pdf" in nome_arq.lower() else " Excel"
                cols_dw[i % 3].download_button(
                    label=f"Baixar {ext}: {nome_arq[:15]}...", 
                    data=file, 
                    file_name=nome_arq, 
                    key=f"btn_{path}" # Chave única para não dar conflito
                )
# =========================================================================
# 1. FUNÇÃO DO POP-UP COM O SEU MOTOR DE ENVIO ORIGINAL (CORRIGIDO)
# =========================================================================
@st.dialog("⚠️ Confirmação de Disparo em Lote")
def confirmar_envio_atendimentos_popup(arquivos_validos):
    # Agrupa os relatórios por RA usando exatamente a sua lógica original
    agrupados_por_ra = {}
    for arq in arquivos_validos:
        base = arq.replace(".pdf", "").replace(".xlsx", "")
        if base not in agrupados_por_ra: 
            agrupados_por_ra[base] = []
        agrupados_por_ra[base].append(arq)
    
    total_envios = len(agrupados_por_ra)
    
    st.write("Você tem certeza que deseja disparar os relatórios de atendimento em lote?")
    st.write(f"• **Destinatário Cadastrado:** `{email_destinatario}`")
    st.write(f"• **Total de e-mails a serem gerados:** {total_envios} mensagens")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("✅ Sim, Disparar Todos", use_container_width=True):
            import time
            
            st.subheader("🚀 Iniciando disparo...")
            sucessos = 0
            barra = st.progress(0)
            
            # Executa o seu laço original de envio com o host correto: smtp.gmail.com
            for i, (base, lista_anexos) in enumerate(agrupados_por_ra.items()):
                sucesso, msg = enviar_relatorio_email(
                    lista_anexos, "smtp.gmail.com", 587, "hudson.valente@crti.com.br", senha_app, email_destinatario
                )
                if sucesso:
                    sucessos += 1
                    st.success(msg)
                else: 
                    st.error(msg)
                barra.progress((i + 1) / total_envios)
                
            st.info(f"🎉 **{sucessos}** de **{total_envios}** e-mails enviados.")
            
            # Segura o feedback visual na tela por 3 segundos antes de fechar o pop-up
            if sucessos == total_envios:
                st.balloons()
                time.sleep(10)
            else:
                time.sleep(5)
                
            st.rerun()
            
    with col_p2:
        if st.button("❌ Não, Cancelar", use_container_width=True):
            st.rerun()

# =========================================================================
# 2. LÓGICA DO GATILHO DA SIDEBAR QUE CHAMA O POP-UP
# =========================================================================
if btn_enviar_emails:
    st.markdown("---")
    arquivos_pasta = glob.glob(os.path.join(PASTA_SAIDA, "*.*"))
    arquivos_validos = [f for f in arquivos_pasta if f.endswith(".pdf") or f.endswith(".xlsx")]
    
    if not arquivos_validos:
        st.warning("⚠️ Gere os relatórios primeiro.")
        st.stop()
        
    # Abre o pop-up de validação passando a lista de arquivos prontos
    confirmar_envio_atendimentos_popup(arquivos_validos)



