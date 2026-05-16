import streamlit as st
from streamlit_gsheets import GSheetsConnection
from docxtpl import DocxTemplate
from datetime import datetime
import io

st.set_page_config(page_title="Gerador de Termos CRTI", page_icon="📄", layout="centered")

st.title("📄 Gerador de Termos de Homologação")
st.write("Selecione o cliente e o módulo para gerar o documento físico customizado.")

# --- 1. CONEXÃO COM GOOGLE SHEETS ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl="5m")
    
    # ATENÇÃO: Substitua 'Cliente' pelo nome exato da coluna da sua planilha google
    lista_clientes = df['Cliente'].dropna().unique().tolist()
except Exception as e:
    st.error(f"Erro ao conectar com o Google Sheets: {e}")
    lista_clientes = []

# --- 2. INTERFACE DO USUÁRIO ---

modulo = st.selectbox(
    "Selecione o Módulo de Treinamento:",
    ["Gestão de Compras", "Gestão de Suprimentos"]
)

# Caminhos exatos dos arquivos que estão na sua pasta 'modelos'
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

# --- 3. PROCESSAMENTO DO DOCUMENTO ---

if st.button("Gerar Documento", type="primary"):
    if not cliente_selecionado:
        st.warning("Por favor, selecione ou digite o nome de um cliente.")
    else:
        try:
            # Busca o modelo correto baseado na escolha do selectbox
            caminho_modelo = MAPA_MODELOS.get(modulo)
            doc = DocxTemplate(caminho_modelo)
            
            # Substitui as tags {{ cliente }} e {{ data }} mapeadas no Word
            contexto = {
                "cliente": cliente_selecionado,
                "data": data_formatada
            }
            
            doc.render(contexto)
            
            # Salva o arquivo na memória temporária do servidor
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✨ Documento gerado com sucesso!")
            
            # Botão de download nativo do Streamlit
            st.download_button(
                label=f"📥 Baixar Termo de {modulo} (.docx)",
                data=buffer,
                file_name=f"Termo_{modulo.replace(' ', '_')}_{cliente_selecionado}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except FileNotFoundError:
            st.error(f"Erro: O arquivo de modelo não foi localizado em: {caminho_modelo}")
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")

