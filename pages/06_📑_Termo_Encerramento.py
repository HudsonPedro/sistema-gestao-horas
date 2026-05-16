import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
from datetime import datetime
import io
import os
import subprocess
import base64

# Descobre o caminho absoluto da pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Termos HPTECH", 
    page_icon="hptech.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. CSS PARA OCULTAR O MENU E FORÇAR A LOGO NO TOPO
st.markdown("""
    <style>
        /* Esconde o menu de páginas padrão do Streamlit */
        [data-testid="stSidebarNav"] {display: none;}
        
        /* Zera o espaçamento do topo para a logo subir */
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* Ajuste da logo para não encostar nas laterais */
        [data-testid="stSidebarHeader"] {padding-top: 0rem !important;}
        
        /* Cor do título para o padrão azul CRTI */
        h1 { color: #b0231d; }
        
        /* Estilo da caixinha de usuário */
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -10px; /* Puxa a caixinha um pouco para cima */
        }
    </style>
""", unsafe_allow_html=True)

# 3. FUNÇÕES DE DADOS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

# 4. SIDEBAR COM MENU INTEGRADO (Emojis exatos do seu repositório)
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
    
    # Navegação com rotas limpas do seu sistema
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
    if st.button("📄 Termo Homologação", use_container_width=True):
        st.switch_page("pages/05_📄_Termos.py")
    if st.button("📄 Termo Encerramento", use_container_width=True):
        st.switch_page("pages/06_📄_Termo_Encerramento.py")
    
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# 5. CARREGAMENTO DE LISTAS
try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except Exception as e:
    st.error(f"Erro ao carregar os clientes da planilha: {e}")
    lista_clientes = []

# Título Customizado com Logo em Base64
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <h1 style="margin: 0; font-size: 2.5rem;">Termo de Homologação e Encerramento</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.markdown("Selecione o cliente e os módulos para gerar o Termo de Homologação e Encerramento.")
except:
    st.title("📄 Termo de Homologação e Encerramento")

st.markdown("---")

# Lista Completa de Módulos
TODOS_MODULOS = [
    "Compras", "Suprimentos e Estoque", "Frota - Equipamentos", 
    "Contratos e Medições de Terceiros", "Custos e Resultados", 
    "Apropriações e Apontamentos", "Produção", "Financeiro", 
    "Contábil", "Patrimonial", "Fiscal", "CRTI Buscador", 
    "CRTI Emissor NFe/NFCe", "CRTI Emissor CTe", "CRTI Emissor MDFe", 
    "CRTI Emissor NFSe", "CRTI Emissor Fatura de Locação", 
    "Gestão de Vendas (Produção)", "Gestão de Vendas (Agronegócio)", 
    "Engenharia, Contratos e Medições de Obras", "Locação de Equipamentos", 
    "Qualidade/Avaliação/Documentação", "Cadastros Globais", "Configuração do Sistema"
]

# --- ENTRADAS DO FORMULÁRIO ---
col1, col2 = st.columns(2)
with col1:
    if lista_clientes:
        cliente_selecionado = st.selectbox("Nome do Cliente:", lista_clientes)
    else:
        cliente_selecionado = st.text_input("Nome do Cliente manualmente:")
        
    gerente_crti = st.text_input("Gerente de implantação na CRTI:", value="SUELLEN GOMES", disabled=True)
    
    # Busca dinâmica do solicitante na planilha
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
st.subheader("Configuração dos Módulos")

# Seleção Dinâmica de Módulos Homologados
modulos_homologados = st.multiselect("Selecione os Módulos HOMOLOGADOS:", options=TODOS_MODULOS)

# Campo único de data de virada de produção
dados_homologados_tabela = []
if modulos_homologados:
    dt_virada_unica = st.date_input("Data de Início em Produção (Válida para todos os homologados):", datetime.now())
    data_virada_formatada = dt_virada_unica.strftime("%d/%m/%Y")
    for mod in modulos_homologados:
        dados_homologados_tabela.append({"nome": mod, "data": data_virada_formatada})

# Seleção de Módulos Não Homologados (Oculta os já selecionados)
opcoes_restantes = [m for m in TODOS_MODULOS if m not in modulos_homologados]
modulos_nao_homologados = st.multiselect("Selecione os Módulos NÃO HOMOLOGADOS:", options=opcoes_restantes)

# Tradução da data por extenso (Curitiba, DD de MM de AAAA)
meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_fim.day} de {meses_br[data_fim.month - 1]} de {data_fim.year}"

# --- 6. PROCESSAMENTO E EMISSÃO DO ARQUIVO ---
if st.button("Gerar Termo de Encerramento Geral", type="primary"):
    if not cliente_selecionado:
        st.warning("Por favor, selecione um cliente para prosseguir.")
    elif not modulos_homologados and not modulos_nao_homologados:
        st.warning("Selecione ao menos um módulo (homologado ou não) para compor o documento.")
    else:
        with st.spinner("⏳ Gerando termos customizados (Word e PDF)..."):
            try:
                # Carrega o arquivo oficial 'encerramento.docx'
                caminho_modelo = os.path.join(BASE_DIR, "modelos", "encerramento.docx")
                doc = DocxTemplate(caminho_modelo)
                
                # --- PROCESSAMENTO RICHTEXT PARA AS COLUNAS HORIZONTAIS ---
                rt_nomes = RichText()
                for i, item in enumerate(dados_homologados_tabela):
                    rt_nomes.add(str(item['nome']))
                    if i < len(dados_homologados_tabela) - 1:
                        rt_nomes.add('\n')
                
                rt_datas = RichText()
                for i, item in enumerate(dados_homologados_tabela):
                    rt_datas.add(str(item['data']))
                    if i < len(dados_homologados_tabela) - 1:
                        rt_datas.add('\n')

                rt_nao_homologados = RichText()
                if modulos_nao_homologados:
                    for i, mod in enumerate(modulos_nao_homologados):
                        rt_nao_homologados.add(f"•\t{mod}")
                        if i < len(modulos_nao_homologados) - 1:
                            rt_nao_homologados.add('\n')
                else:
                    rt_nao_homologados.add("Nenhum módulo pendente nesta fase.")
                
                # --- MAPA DE CONTEXTO DUPLO PARA AS DUAS VERSÕES DE TAGS DO WORD ---
                contexto = {
                    "cliente": cliente_selecionado,
                    "gerente_crti": gerente_crti,
                    "gerente_cliente": gerente_cliente,
                    "data_inicio": data_inicio.strftime("%d/%m/%Y"),
                    "data_fim": data_fim.strftime("%d/%m/%Y"),
                    "data_extenso": data_extenso_str,
                    
                    # Suporte para tags sem espaços no Word
                    "nomes_homologados": rt_nomes,
                    "datas_homologados": rt_datas,
                    "texto_nao_homologados": rt_nao_homologados,
                    
                    # Suporte para tags com espaços no Word (Exatamente como nas fotos)
                    " nomes_homologados ": rt_nomes,
                    " datas_homologados ": rt_datas,
                    " texto_nao_homologados ": rt_nao_homologados
                }
                
                doc.render(contexto)
                
                # Conversão em memória para o Streamlit Cloud Linux
                buffer_docx = io.BytesIO()
                doc.save(buffer_docx)
                buffer_docx.seek(0)
                
                arquivo_docx_temporario = "temp_termo_geral.docx"
                arquivo_pdf_gerado = "temp_termo_geral.pdf"
                
                doc.save(arquivo_docx_temporario)
                
                # Comando headless do LibreOffice para criar o PDF limpo
                cmd = f"libreoffice --headless --convert-to pdf {arquivo_docx_temporario}"
                subprocess.run(cmd, shell=True, check=True)
                
                with open(arquivo_pdf_gerado, "rb") as f:
                    buffer_pdf = io.BytesIO(f.read())
                
                # Limpeza de resíduos de segurança do servidor Linux
                if os.path.exists(arquivo_docx_temporario): os.remove(arquivo_docx_temporario)
                if os.path.exists(arquivo_pdf_gerado): os.remove(arquivo_pdf_gerado)
                
                st.success("✨ Termo Geral gerado com sucesso!")
                
                nome_download_bonito = f"Termo de Homologação e Encerramento - {cliente_selecionado}".replace("/", "-")
                
                # Botões lado a lado idênticos à versão estável
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
                st.error(f"Erro ao processar o arquivo físico: {e}")
