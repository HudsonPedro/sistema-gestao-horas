import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import subprocess
import base64

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Gerador de Termos HPTECH", page_icon="hptech.png", layout="wide")

# 2. CSS PARA OCULTAR O MENU PADRÃO E APLICAR SEU DESIGN
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

# 3. SIDEBAR COM SEU MENU CUSTOMIZADO
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
    
    # Navegação com rotas limpas para evitar erros de compilação
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("app.py")
    if st.button("📊 Dashboard", use_container_width=True):
        st.switch_page("01_📊_Dashboard")
    if st.button("📄 Relatórios", use_container_width=True):
        st.switch_page("02_📄_Relatorios")
    if st.button("📝 Lançamento", use_container_width=True):
        st.switch_page("03_📝_Lancamento")
    if st.button("💰 Medição Mensal", use_container_width=True):
        st.switch_page("04_💰_Medicao_Mensal")
    if st.button("📋 Termo Homologação", use_container_width=True): 
       st.switch_page("05_📋_Termos.py")
    if st.button("📑 Termo Encerramento", use_container_width=True): 
        st.switch_page("06_📑_Termo_Encerramento.py")
      
    
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# 4. CONTEÚDO PRINCIPAL (COM LOGO)
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

try:
    img_base64 = get_image_base64("hptechICO.png")
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <h1 style="margin: 0; font-size: 2.5rem;">Gerador de Termos de Homologação</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.markdown("Selecione o cliente e o módulo para gerar o Termo de Homologação.")
except:
    st.title("📄 Gerador de Termos de Homologação")

st.markdown("---")

@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except Exception as e:
    st.error(f"Erro ao carregar os clientes da planilha: {e}")
    lista_clientes = []

modulo = st.selectbox(
    "Selecione o Módulo de Treinamento:",
    ["Compras", 
     "Suprimentos e Estoque", 
     "Frota - Equipamentos", 
     "Contratos e Medições de Terceiros", 
     "Custos e Resultados", 
     "Financeiro", 
     "CRTI Emissor Nfe/NFCe", 
     "CRTI Emissor CTe", 
     "CRTI Emissor MDFe", 
     "CRTI Emissor NFSe", 
     "Gestão de Vendas (Produção)", 
     "Engenharia, Contratos e Medições de Obras", 
     "Locação de Equipamentos"
    ]
)

# MAPA CORRIGIDO (Removido o parêntese incorreto na linha do suprimentos)
MAPA_MODELOS = {
    "Compras": os.path.join(BASE_DIR, "modelos", "compras.docx"),
    "Suprimentos e Estoque": os.path.join(BASE_DIR, "modelos", "suprimentos.docx"),
    "Frota - Equipamentos": os.path.join(BASE_DIR, "modelos", "frotas.docx"),
    "Contratos e Medições de Terceiros": os.path.join(BASE_DIR, "modelos", "terceiros.docx"),
    "Custos e Resultados": os.path.join(BASE_DIR, "modelos", "custos.docx"),
    "Financeiro": os.path.join(BASE_DIR, "modelos", "financeiro.docx"),
    "CRTI Emissor Nfe/NFCe": os.path.join(BASE_DIR, "modelos", "nfe.docx"),
    "CRTI Emissor CTe": os.path.join(BASE_DIR, "modelos", "cte.docx"),
    "CRTI Emissor MDFe": os.path.join(BASE_DIR, "modelos", "mdfe.docx"),
    "CRTI Emissor NFSe": os.path.join(BASE_DIR, "modelos", "nfse.docx"),
    "Gestão de Vendas (Produção)": os.path.join(BASE_DIR, "modelos", "vendas.docx"),
    "Engenharia, Contratos e Medições de Obras": os.path.join(BASE_DIR, "modelos", "engenharia.docx"),
    "Locação de Equipamentos": os.path.join(BASE_DIR, "modelos", "locacao.docx")
}

if lista_clientes:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes)
else:
    cliente_selecionado = st.text_input("Digite o Nome do Cliente manualmente:")

data_selecionada = st.date_input("Data da Homologação:", datetime.now())
data_formatada = data_selecionada.strftime("%d/%m/%Y")

if st.button("Gerar Documento", type="primary"):
    if not cliente_selecionado or cliente_selecionado == "Erro ao carregar":
        st.warning("Por favor, selecione um cliente válido.")
    else:
        with st.spinner("⏳ Gerando arquivos (Word e PDF)..."):
            try:
                caminho_modelo = MAPA_MODELOS.get(modulo)
                doc = DocxTemplate(caminho_modelo)
                
                contexto = {
                    "cliente": cliente_selecionado,
                    "data": data_formatada
                }
                
                doc.render(contexto)
                
                buffer_docx = io.BytesIO()
                doc.save(buffer_docx)
                buffer_docx.seek(0)
                
                nome_base = f"Termo de Homologacao {modulo.replace(' ', ' ')}-{cliente_selecionado.replace(' ', ' ')}"
                arquivo_docx_temporario = f"{nome_base}.docx"
                arquivo_pdf_gerado = f"{nome_base}.pdf"
                
                doc.save(arquivo_docx_temporario)
                
                cmd = f"libreoffice --headless --convert-to pdf {arquivo_docx_temporario}"
                subprocess.run(cmd, shell=True, check=True)
                
                with open(arquivo_pdf_gerado, "rb") as f:
                    buffer_pdf = io.BytesIO(f.read())
                
                if os.path.exists(arquivo_docx_temporario):
                    os.remove(arquivo_docx_temporario)
                if os.path.exists(arquivo_pdf_gerado):
                    os.remove(arquivo_pdf_gerado)
                
                st.success("✨ Documentos gerados com sucesso!")
                
                # BOTÕES DE DOWNLOAD CORRIGIDOS E FINALIZADOS LADO A LADO
                col_down1, col_down2 = st.columns(2)
                
                with col_down1:
                    st.download_button(
                        label="📥 Baixar em PDF (.pdf)",
                        data=buffer_pdf,
                        file_name=f"{nome_base}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                with col_down2:
                    st.download_button(
                        label="📥 Baixar em Word (.docx)",
                        data=buffer_docx,
                        file_name=f"{nome_base}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                
            except subprocess.CalledProcessError:
                st.error("Erro ao converter para PDF. Verifique se o arquivo packages.txt com o texto 'libreoffice' foi criado na raiz do GitHub.")
            except FileNotFoundError:
                st.error(f"Erro: O arquivo de modelo não foi localizado em: {caminho_modelo}")
            except Exception as e:
                st.error(f"Ocorreu um erro ao processar o documento: {e}")
