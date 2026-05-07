import streamlit as st

# 1. CONFIGURAÇÃO DA PÁGINA (Sempre a primeira linha!)
st.set_page_config(page_title="Dashboard CRTI", page_icon="📊", layout="wide")

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
        h1 { color: #004a87; }
    </style>
""", unsafe_allow_html=True)

# 3. SIDEBAR PERSONALIZADA
with st.sidebar:
    # 1. Logo
    st.image("crti.jpg", use_container_width=True)
    
    # Versão simplificada sem busca de e-mail para destravar o app
    st.markdown("""
        <div class="user-block">
            <span style='font-size: 14px;'>👤 <b>Usuário Logado</b></span><br>
            <span style='font-size: 12px; color: #555;'>hudson.valente@crti.com.br</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.title("Menu Principal")
    
    # 3. Navegação Manual
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("app.py")
        
    if st.button("📊 Dashboard", use_container_width=True):
        st.switch_page("pages/01_📊_Dashboard.py")
        
    if st.button("📄 Relatórios", use_container_width=True):
        st.switch_page("pages/02_📄_Relatorios.py")
    
    st.divider()

# --- INÍCIO DA LÓGICA DO DASHBOARD ---
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta

st.title("📊 Dashboard de Horas Trabalhadas")
st.markdown("---")

# [Mantenha o restante do seu código original de carregamento e gráficos aqui]

# ==== CONFIGURAÇÃO DE VALOR ====
VALOR_HORA = 80.00  # <--- ALTERE AQUI O VALOR DA SUA HORA

# Carrega os mesmos dados do Google Sheets
@st.cache_data(ttl=600)
def carregar_dados():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTtNKWayx3w7y8FPuV_hsaYWcZsB6ftUBKpJALkFOnlYxLEbNfu3LH0y76qxQsGhg/pub?output=xlsx"
    dict_abas = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    return dict_abas

#dict_abas = carregar_dados()
st.sidebar.header("⚙️ Configurações GERAIS")

if st.sidebar.button("🔄 Atualizar Planilha Google", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

with st.spinner("⏳ Analisando arquivo do Google Sheets e buscando abas..."):
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
clientes_unicos = sorted(df["CLIENTE"].unique())
cliente_filtro = st.sidebar.multiselect("Clientes:", clientes_unicos, default=clientes_unicos)

consultores_unicos = sorted(df["CONSULTOR"].unique())
consultor_filtro = st.sidebar.multiselect("Consultores:", consultores_unicos, default=consultores_unicos)

df_filtrado = df[(df["CLIENTE"].isin(cliente_filtro)) & (df["CONSULTOR"].isin(consultor_filtro))]

# ==== MÉTRICAS PRINCIPAIS ====
st.markdown("### 📈 Resumo Geral")
col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)

total_horas = df_filtrado["horas_decimal"].sum()
total_financeiro = df_filtrado["valor_total"].sum()
total_dias = df_filtrado["DATA"].nunique()
media_horas_dia = total_horas / total_dias if total_dias > 0 else 0
clientes_atendidos = df_filtrado["CLIENTE"].nunique()

with col_m1:
    st.metric("⏱️ Total de Horas", f"{total_horas:.1f}h")
with col_m2:
    st.metric("💰 Total Financeiro", f"R$ {total_financeiro:,.2f}")
with col_m3:
    st.metric("📅 Dias Trabalhados", int(total_dias))
with col_m4:
    st.metric("📊 Média/Dia", f"{media_horas_dia:.1f}h")
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
            color_continuous_scale="Blues",
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
