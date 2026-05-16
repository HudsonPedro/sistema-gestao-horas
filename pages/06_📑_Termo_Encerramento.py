import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import subprocess
import base64

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

st.set_page_config(page_title="Termo de Encerramento HPTECH", page_icon="hptech.png", layout="wide")

# CSS para unificar o design e ocultar menu nativo
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        .user-block { background-color: #f0f2f6; padding: 8px; border-radius: 8px; margin-top: -10px; }
        h1 { color: #b0231d; }
    </style>
""", unsafe_allow_html=True)

# Menu Lateral Padrão do Sistema
with st.sidebar:
    st.image("hptechNova.png", use_container_width=True)
    st.markdown("---")
    u_email = st.user.get("email") or "hudson.valente@crti.com.br"
    st.markdown(f"<div class='user-block'><span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br><span style='font-size: 11px; color: #555;'>{u_email}</span></div>", unsafe_allow_html=True)
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
    if st.button("📋 Termo Homologação", use_container_width=True): 
       st.switch_page("pages/05_📋_Termos.py")
    if st.button("📑 Termo Encerramento", use_container_width=True): 
        st.switch_page("pages/06_📑_Termo_Encerramento.py")
      
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# Título Customizado com Logo
def get_image_base64(path):
    with open(path, "rb") as img_file: return base64.b64encode(img_file.read()).decode()
try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(f'<div style="display: flex; align-items: center;"><h1 style="margin: 0; font-size: 2.5rem;">Termo de Encerramento</h1><img src="data:image/png;base64,{img_base64}" style="height: 180px;"></div>', unsafe_allow_html=True)
except:
    st.title("📜 Termo de Homologação e Encerramento")

st.markdown("---")

# Carregamento da base via link público Excel
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    return pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')

try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except:
    df_leg = pd.DataFrame()
    lista_clientes = []

# Lista padronizada de módulos da CRTI
TODOS_MODULOS = [
    "Compras", 
    "Suprimentos e Estoque", 
    "Frota - Equipamentos", 
    "Contratos e Medições de Terceiros", 
    "Custos e Resultados", 
    "Apropriações e Apontamentos", 
    "Produção", 
    "Financeiro", 
    "Contábil", 
    "Patrimonial",  
    "Fiscal", 
    "CRTI Buscador",  
    "CRTI Emissor NFe/NFCe", 
    "CRTI Emissor CTe",  
    "CRTI Emissor MDFe",  
    "CRTI Emissor NFSe", 
    "CRTI Emissor Fatura de Locação",  
    "Gestão de Vendas (Produção)",  
    "Gestão de Vendas (Agronegócio)",  
    "Engenharia, Contratos e Medições de Obras", 
    "Locação de Equipamentos", 
    "Qualidade/Avaliação/Documentação", 
    "Cadastros Globais", 
    "Configuração do Sistema"
]

# 1. ENTRADAS DO FORMULÁRIO (DADOS DO PROJETO)
col1, col2 = st.columns(2)
with col1:
    cliente_selecionado = st.selectbox("Nome do Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")
    gerente_crti = st.text_input("Gerente de implantação na CRTI:", value="SUELLEN GOMES", disabled=True)
    
    # Busca dinâmica do gerente do cliente com base na coluna 'Solicitante1' (Coluna B)
    gerente_cliente_sugerido = ""
    if not df_leg.empty and cliente_selecionado:
        solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
        if solicitantes:
            gerente_cliente_sugerido = solicitantes[0]
            
    gerente_cliente = st.text_input("Gerente de Implantação na EMPRESA CLIENTE:", value=gerente_cliente_sugerido)

with col2:
    data_inicio = st.date_input("Data de início da Implantação:", datetime.now())
    data_fim = st.date_input("Data da homologação da implantação:", datetime.now())

st.markdown("---")

# 2. SELEÇÃO DINÂMICA DE MÓDULOS (HOMOLOGADOS E NÃO HOMOLOGADOS)
# 2. SELEÇÃO DINÂMICA DE MÓDULOS (HOMOLOGADOS E NÃO HOMOLOGADOS)
st.subheader("Configuração dos Módulos")
modulos_homologados = st.multiselect("Selecione os Módulos HOMOLOGADOS:", options=TODOS_MODULOS)

# Cria apenas UM campo de data único para todos os módulos homologados
dados_homologados_tabela = []
if modulos_homologados:
    # Um único calendário na tela
    dt_virada_unica = st.date_input("Data de Início em Produção (Válida para todos os módulos homologados):", datetime.now())
    data_virada_formatada = dt_virada_unica.strftime("%d/%m/%Y")
    
    # Preenche a lista que vai para a tabela do Word com a mesma data para todos
    for mod in modulos_homologados:
        dados_homologados_tabela.append({
            "nome": mod, 
            "data": data_virada_formatada
        })

# Módulos não homologados continuam iguais
opcoes_restantes = [m for m in TODOS_MODULOS if m not in modulos_homologados]
modulos_nao_homologados = st.multiselect("Selecione os Módulos NÃO HOMOLOGADOS:", options=opcoes_restantes)

# 3. CONVERSÃO DE DATA POR EXTENSO (Mês em português baseado na data_fim)
meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_fim.day} de {meses_br[data_fim.month - 1]} de {data_fim.year}"

# 4. GERAÇÃO DOS ARQUIVOS (WORD E PDF)
# 4. GERAÇÃO DOS ARQUIVOS (WORD E PDF)1
if st.button("Gerar Termo de Encerramento", type="primary"):
    if not cliente_selecionado:
        st.warning("Selecione um cliente para prosseguir.")
    elif not modulos_homologados and not modulos_nao_homologados:
        st.warning("Selecione ao menos um módulo para compor o documento.")
    else:
        with st.spinner("⏳ Gerando termos customizados (Word e PDF)..."):
            try:
                caminho_modelo = os.path.join(BASE_DIR, "modelos", "encerramento.docx")
                doc = DocxTemplate(caminho_modelo)
                
                contexto = {
                    "cliente": cliente_selecionado,
                    "gerente_crti": gerente_crti,
                    "gerente_cliente": gerente_cliente,
                    "data_inicio": data_inicio.strftime("%d/%m/%Y"),
                    "data_fim": data_fim.strftime("%d/%m/%Y"),
                    "homologados": dados_homologados_tabela,
                    "nao_homologados": modulos_nao_homologados,
                    "data_extenso": data_extenso_str
                }
                
                doc.render(contexto)
                
                # Gerar o arquivo Word em memória para o download nativo
                buffer_docx = io.BytesIO()
                doc.save(buffer_docx)
                buffer_docx.seek(0)
                
                # NOME INTERNO (BLINDADO): Sem espaços para o Linux não se perder
                arquivo_docx_temporario = "temp_termo_geral.docx"
                arquivo_pdf_gerado = "temp_termo_geral.pdf"
                
                # Salva o arquivo de trabalho temporário no disco do servidor
                doc.save(arquivo_docx_temporario)
                
                # Executa o LibreOffice usando o nome interno padronizado
                cmd = f"libreoffice --headless --convert-to pdf {arquivo_docx_temporario}"
                subprocess.run(cmd, shell=True, check=True)
                
                # Lê o PDF gerado de volta para a memória com segurança
                with open(arquivo_pdf_gerado, "rb") as f:
                    buffer_pdf = io.BytesIO(f.read())
                
                # Exclusão segura dos arquivos intermediários usando o nome fixado
                if os.path.exists(arquivo_docx_temporario): os.remove(arquivo_docx_temporario)
                if os.path.exists(arquivo_pdf_gerado): os.remove(arquivo_pdf_gerado)
                
                st.success("✨ Documentos gerados com sucesso!")
                
                # NOME DE DOWNLOAD (BONITO): Aqui sim pode conter espaços e o nome do cliente formatado
                nome_download_bonito = f"Termo de Homologação e Encerramento - {cliente_selecionado}".replace("/", "-")
                
                col_down1, col_down2 = st.columns(2)
                with col_down1:
                    st.download_button(
                        label="📥 Baixar Termo em PDF (.pdf)", 
                        data=buffer_pdf, 
                        file_name=f"{nome_download_bonito}.pdf", 
                        mime="application/pdf", 
                        use_container_width=True
                    )
                with col_down2:
                    st.download_button(
                        label="📥 Baixar Termo em Word (.docx)", 
                        data=buffer_docx, 
                        file_name=f"{nome_download_bonito}.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"Erro ao processar o documento físico: {e}")
