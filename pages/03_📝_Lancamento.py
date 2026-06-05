#Por Hptec Informatica 
#("v1.0 - 11052026") #16:43 sem alterações

import datetime
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import base64

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Lançamento de Atividades", 
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


# 3. FUNÇÕES DE DADOS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

from google.oauth2 import service_account

def conectar_google_sheets():
    # LINKS OBRIGATÓRIOS para o Google liberar a gravação
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

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
    if st.button("📝 Lançamento de Horas", use_container_width=True):
        st.switch_page("pages/03_📝_Lancamento.py")
    if st.button("📄 Relatórios RA", use_container_width=True):
        st.switch_page("pages/02_📄_Relatorios.py")    
    if st.button("💰 Medição Mensal", use_container_width=True):
        st.switch_page("pages/04_💰_Medicao_Mensal.py")
    if st.button("📋 Termo Homologação", use_container_width=True): 
       st.switch_page("pages/05_📋_Termos.py")
    if st.button("📑 Termo Encerramento", use_container_width=True): 
        st.switch_page("pages/06_📑_Termo_Encerramento.py")
             
    st.divider()
    st.caption("v1.0001-05062026") ###==> INCLUSÃO PENDENCIAS <== v1.0-11052026
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")

# 5. CARREGAMENTO DE LISTAS
try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
    lista_situacao = sorted(df_leg.iloc[:, 4].dropna().unique().tolist())
    
except:
    lista_clientes = ["Erro ao carregar"]
    lista_situacao = ["Concluído", "Pendente"]

#st.title("📝 Lançamento de Atividades")
# Função para converter imagem local para Base64 (para funcionar dentro do HTML)
def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Tenta carregar a imagem que está no seu repositório GitHub
try:
    img_base64 = get_image_base64("hptechICO.png")
    
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <h1 style="margin: 0; font-size: 2.5rem;">Lançamento de Atividades</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("📝 Lançamento de Atividades")

st.markdown("---")

# 6. FILTRO DINÂMICO
col_top1, col_top2 = st.columns([1, 2])
with col_top1:
    data_atendimento = st.date_input("DATA", datetime.now())
    cliente_selecionado = st.selectbox("CLIENTE", options=lista_clientes)

try:
    lista_solicitantes = sorted(df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist())
except:
    lista_solicitantes = []
try:
    lista_local = sorted(df_leg[df_leg["Clientes"] == cliente_selecionado]["Local"].dropna().unique().tolist())
except:
    lista_local = []

# 7. FORMULÁRIO
with st.form("form_lancamento", clear_on_submit=True):
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        ra = st.text_input("RA (Número)")
        situacao_ra = st.selectbox("SITUACAO", options=lista_situacao[2])
        consultor = st.text_input("CONSULTOR", value="Hudson Valente")
        local = st.selectbox("LOCAL", options=lista_local)
       
    with col2:
        hr_inicio = st.time_input("HORA INICIO")
        hr_fim = st.time_input("HORA FIM")
        solicitante = st.selectbox("SOLICITANTE", options=lista_solicitantes)
        forma = st.selectbox("FORMA", ["Remoto", "Presencial"])
        
    with col3:
        hr_inicio_d = st.time_input("HORA INICIO(Desloc)")
        hr_fim_d = st.time_input("HORA FIM(Desloc)")
        km_d = st.number_input("KM(Desloc)", min_value=0.0, step=0.1)
        forma_d = st.selectbox("FORMA(Desloc)", [" ", "Carro Próprio", "Ônibus", "Taxi/Uber", "Avião"])

    with col4:
        descricao_p = st.text_input("DESCRICAO(Pendência)")
        responsavel_p = st.selectbox("RESPONSÁVEL(Pendência)", [" ","Cliente", "CRTI"])
        status_p = st.selectbox("STATUS(Pendência)", [" ", "Pendente", "Realizado"])
           
    st.markdown("---")
    observacoes = st.text_area("OBSERVAÇÕES")
    participante = st.text_input("PARTICIPANTES")
    descricao_d = st.text_area("DESCRICAO(Desloc)")
    

    btn_enviar = st.form_submit_button("🚀 Gravar na Base de Dados")

# 8. LÓGICA DE GRAVAÇÃO (AJUSTADA)
if btn_enviar:
    with st.spinner("⏳ Localizando data na base de dados..."):
        try:
            client = conectar_google_sheets()
            planilha_id = "1m__s5DERX8Lca7r9hi5oZR3HRtmA6yZjVeqqGXyisp4"
            sheet = client.open_by_key(planilha_id)
            
            # Nome da aba dinâmico baseado na data escolhida (Ex: "Maio 2026")
            meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                     "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            nome_aba = f"{meses[data_atendimento.month - 1]} {data_atendimento.year}"
            
            aba = sheet.worksheet(nome_aba)
            
            # --- LÓGICA DE BUSCA DE LINHA (EVITA SOBREPOSIÇÃO) ---
            data_procurada = data_atendimento.strftime('%d/%m/%Y')
            coluna_datas = aba.col_values(1)   # Coluna A (Datas)
            coluna_clientes = aba.col_values(9) # Coluna I (Clientes) - Ajuste se necessário

            try:
                # 1. Acha a primeira ocorrência da data
                linha_destino = coluna_datas.index(data_procurada) + 1
                
                # 2. Enquanto a linha atual tiver a mesma data E o cliente não estiver vazio, desce para próxima linha
                # Isso permite lançar várias atividades no mesmo dia
                while (linha_destino <= len(coluna_datas) and 
                       coluna_datas[linha_destino-1] == data_procurada and 
                       linha_destino <= len(coluna_clientes) and 
                       coluna_clientes[linha_destino-1].strip() != ""):
                    linha_destino += 1
                
                # 3. Validação: se mudou a data ou acabou a planilha, algo está errado
                if linha_destino > len(coluna_datas) or coluna_datas[linha_destino-1] != data_procurada:
                     st.warning("⚠️ Não há mais linhas em branco disponíveis para este dia.")
                     st.stop()

            except ValueError:
                st.error(f"❌ A data {data_procurada} não foi encontrada na coluna A.")
                st.stop()


            # Preparamos os valores para as colunas específicas (seguindo sua imagem)
            # Nota: Ajuste os índices das colunas conforme sua planilha real
            # update_cells ou update_range é mais preciso que append_row aqui
            
            # Exemplo de mapeamento baseado na sua imagem:
            # B: HR_INICIO, C: HR_FIM, D: TOTAL_HR (fórmula), G: CLIENTE...
            
            # O range deve ser a LETRA da coluna seguida do número da LINHA encontrada
            valores_atualizacao = [
                {'range': f'D{linha_destino}', 'values': [[hr_inicio.strftime('%H:%M')]]},
                {'range': f'E{linha_destino}', 'values': [[hr_fim.strftime('%H:%M')]]},
                {'range': f'I{linha_destino}', 'values': [[cliente_selecionado]]},
                {'range': f'J{linha_destino}', 'values': [[ra]]},
                {'range': f'L{linha_destino}', 'values': [[situacao_ra]]},
                {'range': f'M{linha_destino}', 'values': [[observacoes]]},
                {'range': f'N{linha_destino}', 'values': [[consultor]]},
                {'range': f'O{linha_destino}', 'values': [[solicitante]]},
                {'range': f'P{linha_destino}', 'values': [[participante]]},
                {'range': f'Q{linha_destino}', 'values': [[forma]]},
                {'range': f'R{linha_destino}', 'values': [[local]]},
                {'range': f'S{linha_destino}', 'values': [[hr_inicio_d.strftime('%H:%M')]]},
                {'range': f'T{linha_destino}', 'values': [[hr_fim_d.strftime('%H:%M')]]},
                {'range': f'V{linha_destino}', 'values': [[km_d]]},
                {'range': f'W{linha_destino}', 'values': [[forma_d]]},
                {'range': f'X{linha_destino}', 'values': [[descricao_d]]},
                {'range': f'Y{linha_destino}', 'values': [[descricao_p]]},
                {'range': f'Z{linha_destino}', 'values': [[responsavel_p]]},
                {'range': f'AA{linha_destino}', 'values': [[status_p]]}
            ]

            aba.batch_update(valores_atualizacao, value_input_option='USER_ENTERED')
            
            st.success(f"✅ Lançamento realizado com sucesso na linha {linha_destino}!")
            st.balloons()

        except Exception as e:
            st.error(f"❌ Erro crítico: {e}")

