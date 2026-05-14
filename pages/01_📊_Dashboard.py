#Por Hptec Informatica 
#("v1.0 - 27042026") #16:43 sem alterações
import streamlit as st
import base64

# 1. CONFIGURAÇÃO DA PÁGINA (Sempre a primeira linha!)
#st.set_page_config(page_title="Dashboard HPTECH", page_icon="📊", layout="wide")

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Dashboard HPTECH",
    page_icon="hptechICO.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="stSidebarContent"] {padding-top: 0rem !important;}
        
        /* BLOCO USER*/
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
        
        /* Cor do título para o padrão vermelho */
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
        
    # Botão Navegação Atualizada
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
               
# --- INÍCIO DA LÓGICA DO DASHBOARD ---
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta

#st.title("📊 Dashboard de Horas Trabalhadas")
#st.markdown("---")

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
            <h1 style="margin: 0; font-size: 2.5rem;">Dashboard de Horas Trabalhadas</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("📊 Dashboard de Horas Trabalhadas")

st.markdown("---")

# ==== CONFIGURAÇÃO DE VALOR ====
VALOR_HORA = 80.00  # <--- ALTERE AQUI O VALOR DA SUA HORA

# Carrega os mesmos dados do Google Sheets

@st.cache_data(ttl=600)
def carregar_dados():
    # URL de Publicação na Web, forçando saída XLSX
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    
    # sheet_name=None carrega todas as abas em um dicionário
    dict_abas = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    return dict_abas

# --- LOGICA DE EXECUÇÃO ---
try:
    dict_abas = carregar_dados()
    abas_disponiveis = list(dict_abas.keys())
except Exception as e:
    st.error(f"❌ Erro ao conectar: {e}")
    st.info("💡 Dica: Verifique se em 'Arquivo > Compartilhar > Publicar na Web', a opção 'Todo o documento' está selecionada como 'Microsoft Excel'.")
    st.stop()
   
if st.sidebar.button("🔄 Atualizar Base de Dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()
    
with st.spinner("⏳ Analisando Dados..."):
    try:
        dict_abas = carregar_dados()
        abas_disponiveis = list(dict_abas.keys())
    except Exception as e:
        st.error(f"❌ Erro ao baixar planilha: Verifique se possui a biblioteca openpyxl instalada. Erro: {e}")
        st.stop()
# Selector de mês
col1, col2 = st.columns(2)
with col1:
    mes_selecionado = st.selectbox("📅 Selecione o Mês:", list(dict_abas.keys()))

df = dict_abas[mes_selecionado].copy()

# Tratamento inicial
df = df.dropna(subset=["CLIENTE", "CONSULTOR", "TOTAL_HR"])
df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)

# Converte TOTAL_HR para minutos
def converter_para_minutos(valor):
    try:
        val_str = str(valor).strip()
        if ":" in val_str:
            p = val_str.split(":")
            return int(p[0]) * 60 + int(p[1])
    except:
        pass
    return 0

df["minutos"] = df["TOTAL_HR"].apply(converter_para_minutos)
df["horas_decimal"] = df["minutos"] / 60
df["valor_total"] = df["horas_decimal"] * VALOR_HORA

# ==== FILTROS ====
st.sidebar.markdown("### 🔍 Filtros")

# Garantir que não existam valores nulos e converter para string antes de ordenar
clientes_unicos = sorted(df["CLIENTE"].dropna().astype(str).unique())
cliente_filtro = st.sidebar.multiselect("Clientes:", clientes_unicos, default=clientes_unicos)

consultores_unicos = sorted(df["CONSULTOR"].dropna().astype(str).unique())
consultor_filtro = st.sidebar.multiselect("Consultores:", consultores_unicos, default=consultores_unicos)

df_filtrado = df[(df["CLIENTE"].astype(str).isin(cliente_filtro)) & 
                  (df["CONSULTOR"].astype(str).isin(consultor_filtro))]

# ==== MÉTRICAS PRINCIPAIS ====
st.markdown("### 📈 Resumo Geral")
col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)

total_horas = df_filtrado["horas_decimal"].sum()
total_financeiro = df_filtrado["valor_total"].sum()
total_dias = df_filtrado["DATA"].nunique()
media_horas_dia = total_horas / total_dias if total_dias > 0 else 0
clientes_atendidos = df_filtrado["CLIENTE"].nunique()

def formatar_horas_relogio(horas_decimais):
    total_minutos = int(round(horas_decimais * 60))
    horas = total_minutos // 60
    minutos = total_minutos % 60
    return f"{horas:02d}:{minutos:02d}"

with col_m1:
    # De 36.2h para 36:12
    horas_formatadas = formatar_horas_relogio(total_horas)
    st.metric("⏱️ Total de Horas", horas_formatadas)
# Formata como 2,900.00 -> troca vírgula por "X" -> troca ponto por vírgula -> troca "X" por ponto
valor_formatado = f"{total_financeiro:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
with col_m2:
    st.metric("💰 Total Financeiro", f"R$ {valor_formatado}")
with col_m3:
    st.metric("📅 Dias Trabalhados", int(total_dias))
with col_m4:
    # De 7.2h para 07:12
    media_formatada = formatar_horas_relogio(media_horas_dia)
    st.metric("📊 Média/Dia", media_formatada)
with col_m5:
    st.metric("🏢 Clientes", int(clientes_atendidos))

# ==== GRÁFICOS ====
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs(["📊 Por Cliente", "👨‍💼 Por Consultor", "📈 Timeline", "🎯 Distribuição"])

with tab1:
    col_t1, col_t2 = st.columns([1, 2])
    
# Agrupamento para tabela e gráfico
    resumo_cliente = df_filtrado.groupby("CLIENTE").agg({
        "horas_decimal": "sum",
        "valor_total": "sum"
    }).sort_values("horas_decimal", ascending=True)
    
    with col_t1:
        st.markdown("**Resumo Financeiro**")
        st.dataframe(
            resumo_cliente.sort_values("horas_decimal", ascending=False)
            .style.format({"horas_decimal": "{:.1f}h", "valor_total": "R$ {:,.2f}"}),
            use_container_width=True
        )

    with col_t2:
        fig = px.bar(
            resumo_cliente,
            x="horas_decimal",
            y=resumo_cliente.index,
            orientation="h",
            title="Horas por Cliente",
            labels={"horas_decimal": "Horas", "CLIENTE": "Cliente"},
            color="valor_total",
            color_continuous_scale="Reds",
            text_auto='.1f'
        )
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    horas_por_consultor = df_filtrado.groupby("CONSULTOR")["horas_decimal"].sum().sort_values(ascending=True)
    fig = px.bar(
        x=horas_por_consultor.values,
        y=horas_por_consultor.index,
        orientation="h",
        title="Horas Trabalhadas por Consultor",
        labels={"x": "Horas", "y": "Consultor"},
        color=horas_por_consultor.values,
        color_continuous_scale="Greens",
        text_auto='.1f'
    )
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    horas_por_data = df_filtrado.groupby(df_filtrado["DATA"].dt.date)["horas_decimal"].sum()
    fig = px.line(
        x=horas_por_data.index,
        y=horas_por_data.values,
        title="Evolução de Horas por Data",
        labels={"x": "Data", "y": "Horas"},
        markers=True
    )
    st.plotly_chart(fig, use_container_width=True)

with tab4:
    df_sun = df_filtrado.groupby(["CLIENTE", "CONSULTOR"])["horas_decimal"].sum().reset_index()
    df_sun = df_sun[df_sun["horas_decimal"] > 0]
    
    if not df_sun.empty:
        fig = px.sunburst(
            df_sun,
            path=["CLIENTE", "CONSULTOR"],
            values="horas_decimal",
            title="Distribuição Hierárquica: Cliente → Consultor",
            color="horas_decimal",
            color_continuous_scale="Viridis"
        )
        st.plotly_chart(fig, use_container_width=True)

# ==== TABELA DETALHADA ====
st.markdown("---")
st.markdown("### 📋 Detalhamento por Dia")
tabela_detalhe = df_filtrado[["DATA", "CLIENTE", "CONSULTOR", "OBSERVAÇÕES", "TOTAL_HR"]].copy()
st.dataframe(tabela_detalhe.sort_values("DATA", ascending=False), use_container_width=True)

# ==== EXPORTAR RELATÓRIO ====
st.markdown("---")
if st.button("📥 Preparar Arquivo para Exportar"):
    arquivo_excel = "dashboard_horas.xlsx"
    with pd.ExcelWriter(arquivo_excel) as writer:
        df_filtrado.groupby("CLIENTE").agg({"horas_decimal": "sum", "valor_total": "sum"}).to_excel(writer, sheet_name="Por Cliente")
        df_filtrado.to_excel(writer, sheet_name="Detalhes", index=False)
    
    with open(arquivo_excel, "rb") as file:
        st.download_button(label="Clique aqui para Baixar o Excel", data=file, file_name=arquivo_excel)
