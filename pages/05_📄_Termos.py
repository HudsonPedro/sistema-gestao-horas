import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io

# 1. CONFIGURAÇÃO DA PÁGINA (Deve ser a primeira linha de comandos)
st.set_page_config(page_title="Gerador de Termos CRTI", page_icon="hptech.png", layout="wide")

# 2. CSS PARA OCULTAR O MENU PADRÃO E APLICAR SEU DESIGN
st.markdown("""
    <style>
        /* Esconde o menu nativo do Streamlit */
        [data-testid="stSidebarNav"] {display: none;}
        
        /* Zera o espaçamento do topo */
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* Estilo da caixinha de usuário */
        .user-block {
            background-color: #f0f2f6;
            padding: 8px;
            border-radius: 8px;
            margin-top: -10px;
        }
        
        /* Cor do título principal */
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
    if st.button("📋 Gerador de Termos", use_container_width=True):
        st.switch_page("05_📄_Termos")
    
    st.divider()
    st.caption("v1.0 - 11052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# 4. CONTEÚDO PRINCIPAL
st.title("📄 Gerador de Termos de Homologação")
st.write("Selecione o cliente e o módulo para gerar o documento customizado.")

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
    ["Gestão de Compras", "Gestão de Suprimentos"]
)

# CAMINHOS AJUSTADOS EXATAMENTE COMO ESTÃO NO SEU GITHUB (Com "DO_MODULO_DE_GESTÃO" removendo acentos)
MAPA_MODELOS = {
    "Gestão de Compras": "modelos/TERMO_DE_HOMOLOGACAO_DO_MODULO_DE_COMPRAS.docx",
    "Gestão de Suprimentos": "modelos/TERMO_DE_HOMOLOGACAO_DO_MODULO_DE_SUPRIMENTOS.docx"
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
        try:
            caminho_modelo = MAPA_MODELOS.get(modulo)
            doc = DocxTemplate(caminho_modelo)
            
            contexto = {
                "cliente": cliente_selecionado,
                "data": data_formatada
            }
            
            doc.render(contexto)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✨ Documento gerado com sucesso!")
            
            st.download_button(
                label=f"📥 Baixar Termo de {modulo} (.docx)",
                data=buffer,
                file_name=f"Termo_{modulo.replace(' ', '_')}_{cliente_selecionado}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except FileNotFoundError:
            st.error(f"Erro: O arquivo de modelo não foi localizado em: {caminho_modelo}")
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o documento: {e}")
