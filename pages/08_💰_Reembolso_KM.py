#Por Hudson Valente - HPTECH
#Criado em: 09/06/2026
import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import requests
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image as PILImage

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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
    
    st.divider()
    st.caption("v1.0 - 09062026")

# URL oficial mestre da planilha publicada
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
st.write("Módulo automatizado para consolidação de distâncias, cálculo de reembolso e anexo de abastecimento.")
st.markdown("---")

col_f1, col_f2 = st.columns(2)
with col_f1:
    cliente_selecionado = st.selectbox("Selecione o Cliente:", lista_clientes)
    aba_mes_selecionada = st.selectbox("Selecione o Mês de Referência (Aba):", lista_abas_meses)
with col_f2:
    vlr_abast = st.number_input("Valor Unitário Último Abastecimento (R$):", min_value=0.0, value=5.50, step=0.01)
    comprovante_file = st.file_uploader("Subir Comprovante de Abastecimento (Imagem PNG/JPG):", type=["png", "jpg", "jpeg"])

# Descobre o endereço do cliente automaticamente da aba Legendas (Coluna F)
endereco_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    try:
        col_end = [c for c in df_leg.columns if "ENDERECO" in c.upper() or "ENDEREÇO" in c.upper()][0]
        end_df = df_leg[df_leg["Clientes"] == cliente_selecionado][col_end].dropna()
        if not end_df.empty:
            endereco_cliente_sugerido = str(end_df.iloc[0]).strip()
    except:
        endereco_cliente_sugerido = "Endereço não localizado nas Legendas"
if st.button("Gerar Relatório de Reembolso de KM", type="primary", use_container_width=True):
    if not cliente_selecionado or not aba_mes_selecionada:
        st.warning("Preencha todos os seletores para processar os dados.")
    else:
        with st.spinner("⏳ Baixando registros da aba do mês e consolidando quilometragem..."):
            try:
                res_dados = requests.get(URL_PLANILHA_MUDANCA, timeout=45, stream=True)
                xl_dados = pd.ExcelFile(io.BytesIO(res_dados.content))
                df_dados = pd.read_excel(xl_dados, sheet_name=aba_mes_selecionada, engine='openpyxl')
                
                df_dados = df_dados.reset_index(drop=True)
                df_dados.columns = df_dados.columns.str.upper().str.strip()
                
                col_cliente = "CLIENTE"
                col_km_d = "KM_D"
                col_data = "DATA"
                
                if col_km_d not in df_dados.columns:
                    col_km_d = next((c for c in df_dados.columns if "KM" in c or "DISTANCIA" in c), None)

                if not col_km_d:
                    st.error(f"⚠️ Erro: Coluna 'KM_D' (Coluna V) não foi localizada na aba '{aba_mes_selecionada}'.")
                else:
                    # Filtra os dados da planilha usando o Cliente e a distância preenchida como chaves
                    dados_filtrados = df_dados[
                        (df_dados[col_cliente] == cliente_selecionado) & 
                        (df_dados[col_km_d].notna()) & 
                        (df_dados[col_km_d] > 0)
                    ].copy()
                    
                    if dados_filtrados.empty:
                        st.warning(f"Nenhum registro com KM informado foi localizado para '{cliente_selecionado}' no mês.")
                    else:
                        dados_filtrados[col_data] = pd.to_datetime(dados_filtrados[col_data], errors='coerce')
                        dados_filtrados = dados_filtrados.sort_values(by=col_data).reset_index(drop=True)
                        
                        linhas_relatorio = []
                        total_km = 0.0
                        total_valor = 0.0
                        
                        # Percorre as linhas aplicando as regras estritas de percurso Ida e Volta
                        for idx, linha in dados_filtrados.iterrows():
                            dt_str = linha[col_data].strftime("%d/%m/%Y") if pd.notnull(linha[col_data]) else ""
                            km_linha = float(linha[col_km_d])
                            
                            # Regra de alternância: Linhas pares tratadas como Ida, ímpares como Volta
                            if idx % 2 == 0:
                                percurso = "Ida"
                                orig = ENDERECO_CRTI_FIXO
                                dest = endereco_cliente_sugerido
                            else:
                                percurso = "Volta"
                                orig = endereco_cliente_sugerido
                                dest = ENDERECO_CRTI_FIXO
                                
                            # Cálculo matemático oficial do requisito: DISTÂNCIA * (Vlr Unit * 0.25)
                            valor_linha = km_linha * (vlr_abast * 0.25)
                            
                            total_km += km_linha
                            total_valor += valor_linha
                            
                            linhas_relatorio.append({
                                "DATA": dt_str,
                                "CLIENTE": cliente_selecionado,
                                "LOCAL ORIGEM": orig,
                                "LOCAL DESTINO": dest,
                                "Percurso": percurso,
                                "DISTÂNCIA (KM)": km_linha,
                                "Vlr Unit. Ultimo Abast.": f"R$ {vlr_abast:,.2f}",
                                "TOTAL": f"R$ {valor_linha:,.2f}"
                            })
                        
                        df_relatorio_final = pd.DataFrame(linhas_relatorio)
                        
                        # Adiciona a linha de totalizadores no fim do dataframe do Excel
                        df_excel = df_relatorio_final.copy()
                        df_excel.loc[len(df_excel)] = ["Total KM", f"{total_km:.0f}", "Valor Total", f"R$ {total_valor:,.2f}", "", "", "", ""]
                        
                        # ----------------------------------------------------
                        # COMPILADOR 1: CONSTRUÇÃO EXCEL (.XLSX)
                        # ----------------------------------------------------
                        buffer_xlsx = io.BytesIO()
                        with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                            df_excel.to_excel(writer, sheet_name="Reembolso", index=False)
                        buffer_xlsx.seek(0)
                        
                        # ----------------------------------------------------
                        # COMPILADOR 2: CONSTRUÇÃO PDF VIA REPORTLAB
                        # ----------------------------------------------------
                        buffer_pdf = io.BytesIO()
                        doc = SimpleDocTemplate(buffer_pdf, pagesize=letter, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
                        styles = getSampleStyleSheet()
                        
                        estilo_titulo = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#b0231d'), alignment=1)
                        estilo_texto = ParagraphStyle('TextStyle', parent=styles['Normal'], fontSize=8)
                        estilo_header = ParagraphStyle('HeadStyle', parent=styles['Normal'], fontSize=8, textColor=colors.white, fontName='Helvetica-Bold')
                        
                        elementos = []
                        elementos.append(Paragraph("RELATÓRIO PARA REEMBOLSO DE KM RODADO", estilo_titulo))
                        elementos.append(Spacer(1, 15))
                        elementos.append(Paragraph(f"IMPLANTADOR CRTI: HUDSON PEDRO SALES VALENTE", estilo_texto))
                        elementos.append(Spacer(1, 10))
                        
                        # Monta a estrutura da tabela do PDF
                        dados_tabela = [[Paragraph(c, estilo_header) for c in df_relatorio_final.columns]]
                        for _, r in df_relatorio_final.iterrows():
                            dados_tabela.append([Paragraph(str(r[c]), estilo_texto) for c in df_relatorio_final.columns])
                            
                        # Adiciona a linha de totalização no PDF
                        dados_tabela.append([
                            Paragraph("<b>Total KM</b>", estilo_texto), Paragraph(f"<b>{total_km:.0f}</b>", estilo_texto),
                            Paragraph("<b>Valor Total</b>", estilo_texto), Paragraph(f"<b>R$ {total_valor:,.2f}</b>", estilo_texto),
                            Paragraph("", estilo_texto), Paragraph("", estilo_texto), Paragraph("", estilo_texto), Paragraph("", estilo_texto)
                        ])
                        
                        tabela_pdf = Table(dados_tabela, colWidths=[50, 70, 110, 110, 45, 45, 55, 60])
                        tabela_pdf.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#b0231d')),
                            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                            ('TOPPADDING', (0,0), (-1,-1), 4),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                        ]))
                        elementos.append(tabela_pdf)
                        
                        # REQUISITO DO ANEXO: Insere a imagem do comprovante logo abaixo do relatório no PDF
                        if comprovante_file:
                            elementos.append(Spacer(1, 20))
                            elementos.append(Paragraph("<b>COMPROVANTE DE ABASTECIMENTO ANEXADO:</b>", estilo_texto))
                            elementos.append(Spacer(1, 10))
                            
                            img_data = comprovante_file.read()
                            img_pil = PILImage.open(io.BytesIO(img_data))
                            
                            # Redimensiona mantendo a proporção para caber na folha do PDF
                            img_pil.thumbnail((400, 400))
                            img_temp_path = os.path.join(BASE_DIR, "temp_comprovante.jpg")
                            img_pil.convert('RGB').save(img_temp_path, "JPEG")
                            
                            elementos.append(Image(img_temp_path, width=img_pil.width, height=img_pil.height))
                            
                        doc.build(elementos)
                        buffer_pdf.seek(0)
                        
                        # Remove o arquivo temporário de imagem se ele tiver sido criado
                        if comprovante_file and os.path.exists(img_temp_path):
                            os.remove(img_temp_path)
                            
                        st.success("✨ Relatório gerado com sucesso em ambos os formatos!")
                        st.markdown("---")
                        
                        col_dl1, col_dl2 = st.columns(2)
                                                col_dl1.download_button(
                            label="📥 **BAIXAR RELATÓRIO EM PDF**",
                            data=buffer_pdf,
                            file_name=f"Relatorio_Reembolso_KM_{cliente_selecionado}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        col_dl2.download_button(
                            label="📥 **BAIXAR PLANILHA EM XLSX**",
                            data=buffer_xlsx,
                            file_name=f"Relatorio_Reembolso_KM_{cliente_selecionado}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"Erro ao compilar relatório de reembolso de KM: {e}")

