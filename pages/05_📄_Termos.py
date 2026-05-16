import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io

st.set_page_config(page_title="Gerador de Termos CRTI", page_icon="📄", layout="centered")

st.title("📄 Gerador de Termos de Homologação")
st.write("Selecione o cliente e o módulo para gerar o documento customizado.")

# --- 1. FUNÇÕES DE DADOS (Idêntica à sua lógica de produção) ---
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

# --- 2. CARREGAMENTO DE LISTAS ---
try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except Exception as e:
    st.error(f"Erro ao carregar os clientes da planilha: {e}")
    lista_clientes = []

# --- 3. INTERFACE DO USUÁRIO ---

modulo = st.selectbox(
    "Selecione o Módulo de Treinamento:",
    ["Gestão de Compras", "Gestão de Suprimentos"]
)

# Caminhos exatos dos arquivos mapeados na pasta 'modelos' do seu repositório
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

# --- 4. PROCESSAMENTO DO DOCUMENTO ---

if st.button("Gerar Documento", type="primary"):
    if not cliente_selecionado or cliente_selecionado == "Erro ao carregar":
        st.warning("Por favor, selecione um cliente válido.")
    else:
        try:
            # Seleciona o arquivo correto com base no módulo escolhido
            caminho_modelo = MAPA_MODELOS.get(modulo)
            doc = DocxTemplate(caminho_modelo)
            
            # Substitui as tags {{ cliente }} e {{ data }} que estão no seu Word
            contexto = {
                "cliente": cliente_selecionado,
                "data": data_formatada
            }
            
            doc.render(contexto)
            
            # Cria o arquivo em memória temporária para o Streamlit Cloud disponibilizar
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✨ Documento gerado com sucesso!")
            
            # Botão nativo para baixar o arquivo pronto em Word (.docx)
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
