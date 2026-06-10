#Por Hudson Valente - HPTECH
#Criado em: 09/06/2026
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from datetime import datetime
import io
import os
import subprocess
import glob
import requests

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Pasta mestre isolada para os relatórios de reembolso emitidos
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
# 3. SIDEBAR COM MENU INTEGRADO ATUALIZADO
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
    st.divider()
    st.caption("v1.0 - 09062026")

# URL mestre da planilha publicada
URL_PLANILHA_MUDANCA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
ENDERECO_CRTI_FIXO = "Rua Padre Anchieta, 2050"

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
st.write("O sistema extrai os dados e distâncias diretamente da planilha de lançamentos do mês selecionado.")
st.markdown("---")

cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")
aba_mes_selecionada = st.selectbox("Selecione o Mês do Atendimento (Aba da Planilha):", lista_abas_meses) if lista_abas_meses else st.text_input("Aba do Mês:")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna()
    if not solicitantes_df.empty:
        gerente_cliente_sugerido = str(solicitantes_df.values[0]).strip()

# Busca o endereço do cliente dinamicamente na aba Legendas, coluna 'Endereco'
endereco_cliente_sugerido = "Endereço do Cliente, Campo Largo"
if not df_leg.empty and cliente_selecionado:
    try:
        end_df = df_leg[df_leg["Clientes"] == cliente_selecionado]["Endereco"].dropna()
        if not end_df.empty:
            endereco_cliente_sugerido = str(end_df.values[0]).strip()
    except:
        pass

solicitante_nome = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)
vlr_abast = st.number_input("Valor Unitário Último Abastecimento (R$):", min_value=0.0, value=5.50, step=0.01)
comprovante_file = st.file_uploader("Subir Comprovante de Abastecimento (Imagem PNG/JPG) [Requisito 2]:", type=["png", "jpg", "jpeg"])

data_emissao = st.date_input("Data de Emissão do Termo:", datetime.now())
meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_emissao.day} de {meses_br[data_emissao.month - 1]} de {data_emissao.year}"
# --- 6. PROCESSAMENTO E FILTRAGEM DINÂMICA DA ABA SELECIONADA ---
if st.button("Gerar Relatório de Reembolso de KM", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Selecione um cliente válido.")
    else:
        with st.spinner(f"⏳ Processando atendimentos da aba '{aba_mes_selecionada}'..."):
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
                col_km_d = "KM_D"
                
                if col_situacao in df_dados.columns:
                    df_dados[col_situacao] = df_dados[col_situacao].astype(str).str.strip()
                
                # Filtra os lançamentos ativos com quilometragem válida e situação "Em Elaboração"
                atendimentos_cliente = df_dados[
                    (df_dados[col_cliente] == cliente_selecionado) & 
                    (df_dados[col_situacao] == "Em Elaboração") & 
                    (df_dados[col_ra].notna()) & 
                    (df_dados[col_km_d].notna()) & 
                    (df_dados[col_km_d] > 0)
                ].copy()
                
                if atendimentos_cliente.empty:
                    st.warning(f"⚠️ Nenhum lançamento de KM ativo em elaboração foi localizado na aba '{aba_mes_selecionada}'.")
                else:
                    for antigo in glob.glob(os.path.join(PASTA_REEMBOLSO_KM, "*.*")):
                        try: os.remove(antigo)
                        except: pass

                    atendimentos_cliente = atendimentos_cliente.reset_index(drop=True)
                    atendimentos_cliente[col_data] = pd.to_datetime(atendimentos_cliente[col_data], errors='coerce')
                    atendimentos_cliente = atendimentos_cliente.sort_values(by=col_data).reset_index(drop=True)
                    
                    lista_atendimentos_word = []
                    total_km = 0.0
                    total_valor = 0.0
                    
                    # Dicionários de carregamento rápido baseados na Lógica da Página 02
                    dict_km = atendimentos_cliente[col_km_d].to_dict()
                    dict_data_dia = atendimentos_cliente[col_data].to_dict()

                    for idx in atendimentos_cliente.index:
                        dt_objeto = dict_data_dia.get(idx)
                        dt_str = dt_objeto.strftime("%d/%m/%Y") if pd.notnull(dt_objeto) else ""
                        km_linha = float(dict_km.get(idx, 0))
                        
                        # Requisito 3: Regra de alternância de percurso Ida (Par) e Volta (Ímpar)
                        if idx % 2 == 0:
                            percurso = "Ida"
                            local_origem = ENDERECO_CRTI_FIXO
                            local_destino = endereco_cliente_sugerido
                        else:
                            percurso = "Volta"
                            local_origem = endereco_cliente_sugerido
                            local_destino = ENDERECO_CRTI_FIXO
                        
                        # Requisito: TOTAL = DISTÂNCIA * (Vlr Unit * 0.25)
                        valor_linha = km_linha * (vlr_abast * 0.25)
                        
                        total_km += km_linha
                        total_valor += valor_linha
                        
                        lista_atendimentos_word.append({
                            "data_dia": dt_str,
                            "cliente": cliente_selecionado,
                            "local_origem": local_origem,
                            "local_destino": local_destino,
                            "percurso": percurso,
                            "km": f"{km_linha:.0f}",
                            "vlr_unit": f"R$ {vlr_abast:,.2f}",
                            "total": f"R$ {valor_linha:,.2f}"
                        })

                    # 1. EXCEL (.XLSX) COM TOTALIZADORES NO FIM
                    df_excel = pd.DataFrame(lista_atendimentos_word)
                    df_excel.loc[len(df_excel)] = ["Total KM", f"{total_km:.0f}", "Valor Total", f"R$ {total_valor:,.2f}", "", "", "", ""]
                    
                    buffer_xlsx = io.BytesIO()
                    with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                        df_excel.to_excel(writer, sheet_name="Reembolso_KM", index=False)
                    st.session_state["km_xlsx_ready"] = buffer_xlsx.getvalue()

                    # 2. PROCESSO WORD (.DOCX) -> PDF (VIA LIBREOFFICE)
                    caminho_modelo = os.path.join(BASE_DIR, "modelos", "reembolso.docx")
                    
                    if not os.path.exists(caminho_modelo):
                        st.error("⚠️ O modelo físico 'reembolso.docx' não foi localizado na pasta 'modelos'.")
                    else:
                        doc = DocxTemplate(caminho_modelo)
                        
                        # Requisito 2: Se houver arquivo enviado, prepara o anexo de imagem dentro do Word
                        contexto = {
                            "cliente": cliente_selecionado,
                            "total_km": f"{total_km:.0f}",
                            "valor_total": f"R$ {total_valor:,.2f}",
                            "data_extenso": data_extenso_str,
                            "kms": lista_atendimentos_word
                        }
                        
                        if comprovante_file:
                            caminho_temp_img = os.path.join(PASTA_REEMBOLSO_KM, "temp_comp.jpg")
                            with open(caminho_temp_img, "wb") as f:
                                f.write(comprovante_file.getbuffer())
                            contexto["comprovante"] = InlineImage(doc, caminho_temp_img, width=Inches(4.5))
                        else:
                            contexto["comprovante"] = "Nenhum comprovante anexado."
                            
                        doc.render(contexto)
                        
                        nome_final = f"Relatorio_Reembolso_KM_{cliente_selecionado}".replace(" ", "_").replace("/", "-")
                        caminho_docx = os.path.join(PASTA_REEMBOLSO_KM, f"{nome_final}.docx")
                        doc.save(caminho_docx)
                        
                        # Converte para PDF nativo pelo LibreOffice do servidor Linux
                        subprocess.run(f'libreoffice --headless --convert-to pdf --outdir "{PASTA_REEMBOLSO_KM}" "{caminho_docx}"', shell=True, check=True)
                        
                        caminho_pdf_gerado = os.path.join(PASTA_REEMBOLSO_KM, f"{nome_final}.pdf")
                        if os.path.exists(caminho_pdf_gerado):
                            with open(caminho_pdf_gerado, "rb") as f_pdf:
                                st.session_state["km_pdf_ready"] = f_pdf.read()
                        
                        if comprovante_file and os.path.exists(caminho_temp_img):
                            try: os.remove(caminho_temp_img)
                            except: pass
                            
                        st.success("✨ Relatório gerado com sucesso!")
                        time.sleep(1)
                        st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar lote na aba selecionada: {e}")

# --- PAINEL VISUAL DE DOWNLOADS ---
if "km_pdf_ready" in st.session_state:
    st.markdown("---")
    st.subheader("📥 Download do Relatório de KM Rodado")
    
    col_dl1, col_dl2 = st.columns(2)
    col_dl1.download_button(
        label="📥 **BAIXAR RELATÓRIO EM PDF**",
        data=st.session_state["km_pdf_ready"],
        file_name=f"Relatorio_Reembolso_KM_{cliente_selecionado}.pdf",
        mime="application/pdf",
        use_container_width=True
    )
    col_dl2.download_button(
        label="📥 **BAIXAR PLANILHA EM XLSX**",
        data=st.session_state["km_xlsx_ready"],
        file_name=f"Relatorio_Reembolso_KM_{cliente_selecionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
