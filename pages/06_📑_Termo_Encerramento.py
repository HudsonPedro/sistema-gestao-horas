#Por Hudson Valente - HPTECH
#Criado em: 16/05/2026
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import subprocess
import base64
import zipfile
import glob
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64

# Descobre a pasta raiz do projeto de forma segura para o Linux
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Pasta mestre dos termos emitidos
PASTA_TERMOS = os.path.join(BASE_DIR, "termos_emitidos")
os.makedirs(PASTA_TERMOS, exist_ok=True)

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="HPTECH Sistema de Gestão",
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


# 3. MOTOR DE DISPARO SMTP DA SUA PÁGINA 04 (BLINDADO CONTRA ERRO DE ACENTUAÇÃO E NONAME)
def enviar_email_termos_logica_p04(string_destinatarios, cliente_nome, data_virada_str, campos_fiscal, arquivos_anexos, gerente_cliente_nome):
    email_remetente = st.secrets["smtp"]["usuario"]
    senha_remetente = st.secrets["smtp"]["senha"]
    smtp_server = st.secrets["smtp"]["servidor"]
    smtp_porta = int(st.secrets["smtp"]["porta"])
    
    lista_destinatarios = [email.strip() for email in string_destinatarios.split(",") if email.strip()]
    if not lista_destinatarios:
        return False, "Nenhum e-mail válido foi inserido."
        
    msg = MIMEMultipart()
    msg["From"] = email_remetente
    msg["To"] = ", ".join(lista_destinatarios)
    msg["Subject"] = f"Solicitação da disponibilidade do suporte para a VIRADADA EM PRODUÇÃO – {cliente_nome}"
    
    corpo_html = f"""
    <html>
    <body>
        <p>Prezados(as), espero que se encontre bem.</p>
        <p>Solicito a disponibilidade do suporte para a VIRADA EM PRODUÇÃO – <b>{cliente_nome}</b>.</p>
        <p>O processo da MUDANÇA NA APLICAÇÃO PARA PRODUÇÃO da <b>{cliente_nome}</b> - (PARÂMETROS E MIGRATE) será no dia <b>{data_virada_str}</b>, data PREVISTA para emissão dos últimos documentos fiscais.</p>
        
        <p><b>Tipo de virada:</b> {campos_fiscal['tipo_virada']}<br>
        <b>Período da Virada:</b> {campos_fiscal['periodo_virada']}<br>
        <b>Solicitante:</b> {gerente_cliente_nome}</p>
        
        <p>Em seguida, os últimos documentos emitidos serão encaminhados para a mudança da aplicação em produção.</p>
        
        <p><b>Último NF-e:</b> {campos_fiscal['nfe']}<br>
        <b>Último MDF-e:</b> {campos_fiscal['mdfe']}<br>
        <b>Último CT-e:</b> {campos_fiscal['cte']}<br>
        <b>Último NFS-e:</b> {campos_fiscal['nfse']}<br>
        <b>Últimos Boletos:</b> {campos_fiscal['boletos']}</p>
        
        <p>A Sra. Amanda, favor.</p>
        
        <p style="color: #b0231d; font-weight: bold;">Enviar os anexos: Termo de Homologação e Encerramento de Implantação do CRTI ERP e o Termo de NÃO Homologação de Módulos do CRTI ERP para assinatura institucional com as informações da data de corte atualizada.</p>
        
        <p style="color: #b0231d; font-weight: bold;">Para virar o sistema em produção é obrigatória a assinatura de ambas as partes. O cliente Sr(a). {gerente_cliente_nome} está ao par e no aguardo dos documentos para assinaturas.</p>
        
        <br>
        <p>Caso encontre alguma divergência, favor criticar para as devidas correções.<br>
        Me coloco à inteira disposição para possíveis esclarecimentos.</p>
        <p>Com Gratidão!!<br><br>Hudson Valente</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(corpo_html, "html"))
    
    for caminho_arquivo in arquivos_anexos:
        nome_original = os.path.basename(caminho_arquivo)
        extensao = ".pdf" if nome_original.lower().endswith(".pdf") else ".docx"
        
        # BLINDAGEM DO GMAIL: Força nomes 100% livres de acentos e sem a palavra 'NÃO' no protocolo MIME
        if "NÃO" in nome_original or "Nao" in nome_original or "naohomologado" in nome_original:
            nome_arquivo_limpo = f"Termo de NAO Homologacao de Modulos do CRTI ERP{extensao}"
        else:
            nome_arquivo_limpo = f"Termo de Homologacao e Encerramento de Implantacao do CRTI ERP{extensao}"
            
        try:
            with open(caminho_arquivo, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encode_base64(part)
                # Adiciona o cabeçalho forçando string limpa e sem acentuação para o Gmail ler
                part.add_header("Content-Disposition", "attachment", filename=str(nome_arquivo_limpo))
                msg.attach(part)
        except Exception as e:
            return False, f"Falha ao acoplar anexo: {str(e)}"
            
    try:
        server = smtplib.SMTP(smtp_server, smtp_porta)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, lista_destinatarios, msg.as_string())
        server.quit()
        return True, "E-mail com o pacote de virada enviado com sucesso para todos!"
    except Exception as e:
        return False, f"Falha no envio SMTP (Segurança de Rede): {str(e)}"

# 4. SIDEBAR COM MENU INTEGRADO
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
    
    # INTERFACE DE DISPARO DA SIDEBAR
    st.markdown("---")
    st.header("📬 Disparo de Termos")
    st.caption("Separe os e-mails usando vírgula (,)")
    email_destinatario = st.text_input("Enviar para (Destinatários):", value="financeiro@crti.com.br,tayna@crti.com.br,kamille.voitach@crti.com.br,camille@crti.com.br", key="enc_email_dest_key")
    btn_enviar_emails = st.button("🚀 **ENVIAR PACOTE COMPLETO**", type="primary", use_container_width=True, key="enc_email_btn_key")
    
    # BOTAO AUXILIAR DE LIMPEZA NA SIDEBAR
    st.markdown("---")
    st.header("🗑️ Gerenciamento")
    if st.button("🗑️ Limpar Termos Emitidos", use_container_width=True):
        arquivos_limpeza = glob.glob(os.path.join(PASTA_TERMOS, "*.*"))
        for arq in arquivos_limpeza:
            try: os.remove(arq)
            except: pass
        st.success("Pasta de histórico esvaziada!")
        time.sleep(1.5)
        st.rerun()

    st.divider()
    st.caption("v1.1 - 16052026")
    st.caption("Todos os direitos reservados")
    st.caption("Copyright ©2026 HPtech Informática ME")
    
# 4. CONTEÚDO PRINCIPAL
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
            <h1 style="margin: 0; font-size: 2.5rem;">Gerador de Termos de Encerramento</h1>
            <img src="data:image/png;base64,{img_base64}" style="margin-left: 0px; height: 180px;">
        </div>
        """, 
        unsafe_allow_html=True
    )
    st.markdown("Selecione o cliente para gerar os Termos.")
except:
    # Caso a imagem mude de nome ou não seja encontrada, mantém apenas o texto
    st.title("📄 Gerador de Termos de Encerramento")

st.markdown("---")

# 5. CARREGAMENTO DE LISTAS
@st.cache_data(ttl=600)
def carregar_legendas():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQABOlTPSx3-hKS7qPIXNl8jODyzQBF-_FVMR4JX3o0WNBmsl5OVPQUi0cNfZ1TMEShcH3hmHIL-kE/pub?output=xlsx"
    df = pd.read_excel(url, sheet_name="Legendas", engine='openpyxl')
    return df

try:
    df_leg = carregar_legendas()
    lista_clientes = sorted(df_leg["Clientes"].dropna().unique().tolist())
except:
    df_leg = pd.DataFrame()
    lista_clientes = []

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

cliente_selecionado = st.selectbox("Nome do Cliente:", lista_clientes) if lista_clientes else st.text_input("Nome do Cliente:")

gerente_cliente_sugerido = ""
if not df_leg.empty and cliente_selecionado:
    solicitantes = df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().unique().tolist()
    if solicitantes: 
        gerente_cliente_sugerido = str(df_leg[df_leg["Clientes"] == cliente_selecionado]["Solicitante1"].dropna().iloc[0]).strip()
        
gerente_cliente_name = st.text_input("Gerente de Implantação na EMPRESA CLIENTE (Solicitante):", value=gerente_cliente_sugerido)
gerente_crti = "SUELLEN GOMES"

st.markdown("---")
st.subheader("Configuração Unificada dos Termos")

col_datas_1, col_datas_2 = st.columns(2)
with col_datas_1:
    data_inicio = st.date_input("Data de início da Implantação Geral:", datetime.now())
with col_datas_2:
    data_fim = st.date_input("Data da homologação / emissão do documento:", datetime.now())

st.markdown("##### Parâmetros do Lote de Mudança")
col_v1, col_v2, col_v3 = st.columns(3)
with col_v1:
    tipo_virada_sel = st.selectbox("Tipo de virada:", ["Remoto", "Presencial"])
with col_v2:
    dt_inicio_v = st.date_input("Período Virada (Início):", datetime.now())
with col_v3:
    dt_fim_v = st.date_input("Período Virada (Fim):", datetime.now())

periodo_virada_str = f"{dt_inicio_v.strftime('%d/%m/%Y')} a {dt_fim_v.strftime('%d/%m/%Y')}"

col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
with col_f1: status_nfe = st.selectbox("Último NF-e:", ["NÃO EMITE", "SIM"])
with col_f2: status_mdfe = st.selectbox("Último MDF-e:", ["NÃO EMITE", "SIM"])
with col_f3: status_cte = st.selectbox("Último CT-e:", ["NÃO EMITE", "SIM"])
with col_f4: status_nfse = st.selectbox("Último NFS-e:", ["SIM", "NÃO EMITE"])
with col_f5: status_boletos = st.selectbox("Últimos Boletos:", ["NÃO EMITE", "SIM"])

campos_fiscais_payload = {
    "tipo_virada": tipo_virada_sel, "periodo_virada": periodo_virada_str,
    "nfe": status_nfe, "mdfe": status_mdfe, "cte": status_cte, "nfse": status_nfse, "boletos": status_boletos
}

st.markdown("---")
col_mod_1, col_mod_2 = st.columns(2)
with col_mod_1:
    modulos_homologados = st.multiselect("Selecione os Módulos HOMOLOGADOS:", options=TODOS_MODULOS)
    dados_homologados_tabela = []
    if modulos_homologados:
        dt_virada_unica = st.date_input("Data de Início em Produção (Para todos os homologados):", datetime.now())
        data_virada_formatada = dt_virada_unica.strftime("%d/%m/%Y")
        for mod in modulos_homologados:
            dados_homologados_tabela.append({"nome": mod, "data": data_virada_formatada})
    else:
        data_virada_formatada = datetime.now().strftime("%d/%m/%Y")

with col_mod_2:
    opcoes_restantes = [m for m in TODOS_MODULOS if m not in modulos_homologados]
    modulos_nao_homologados = st.multiselect("Selecione os Módulos NÃO HOMOLOGADOS (Pendências):", options=opcoes_restantes)

meses_br = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_extenso_str = f"{data_fim.day} de {meses_br[data_fim.month - 1]} de {data_fim.year}"

# --- 6. GERAÇÃO EM LOTE DOS DOIS DOCUMENTOS DE UMA VEZ ---
if st.button("Gerar Todos os Termos do Cliente", type="primary", use_container_width=True):
    if not cliente_selecionado:
        st.warning("Por favor, selecione um cliente para prosseguir.")
    else:
        with st.spinner("⏳ Compilando lote completo (Word e PDF)..."):
            try:
                for arquivo_antigo in glob.glob(os.path.join(PASTA_TERMOS, "*.*")):
                    try: os.remove(arquivo_antigo)
                    except: pass

                texto_nao_homologados_str = "\n".join([f"• {mod}" for mod in modulos_nao_homologados]) if modulos_nao_homologados else "Nenhum módulo pendente nesta fase."
                nomes_homologados_str = "\n".join([str(item['nome']) for item in dados_homologados_tabela])
                datas_homologados_str = "\n".join([str(item['data']) for item in dados_homologados_tabela])

                # --- DOCUMENTO 1: TERMO GERAL ---
                if modulos_homologados:
                    caminho_m1 = os.path.join(BASE_DIR, "modelos", "Lincoln_Pedro_Termos_Campanha_encerramento.docx")
                    if not os.path.exists(caminho_m1): caminho_m1 = os.path.join(BASE_DIR, "modelos", "Lincoln_Pedro_Termos_encerramento.docx")
                    if not os.path.exists(caminho_m1): caminho_m1 = os.path.join(BASE_DIR, "modelos", "encerramento.docx")
                    
                    doc1 = DocxTemplate(caminho_m1)
                    ctx1 = {
                        "cliente": cliente_selecionado, "gerente_crti": "SUELLEN GOMES", "gerente_cliente": gerente_cliente_name,
                        "data_inicio": data_inicio.strftime("%d/%m/%Y"), "data_fim": data_fim.strftime("%d/%m/%Y"), "data_extenso": data_extenso_str,
                        "nomes_homologados": nomes_homologados_str, "datas_homologados": datas_homologados_str, "texto_nao_homologados": texto_nao_homologados_str,
                        " nomes_homologados ": nomes_homologados_str, " datas_homologados ": datas_homologados_str, " texto_nao_homologados ": texto_nao_homologados_str
                    }
                    doc1.render(ctx1)
                    n_arq1 = f"Termo_de_Homologacao_e_Encerramento_de_Implantacao_do_CRTI_ERP"
                    c_docx1 = os.path.join(PASTA_TERMOS, f"{n_arq1}.docx")
                    doc1.save(c_docx1)
                    subprocess.run(f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TERMOS}" "{c_docx1}"', shell=True, check=True)

                # --- DOCUMENTO 2: NÃO HOMOLOGAÇÃO ---
                if modulos_nao_homologados:
                    caminho_m2 = os.path.join(BASE_DIR, "modelos", "naohomologado.docx")
                    doc2 = DocxTemplate(caminho_m2)
                    ctx2 = {
                        "cliente": cliente_selecionado, "gerente_cliente": gerente_cliente_name, "data_extenso": data_extenso_str,
                        "texto_nao_homologados": texto_nao_homologados_str, " texto_nao_homologados ": texto_nao_homologados_str
                    }
                    doc2.render(ctx2)
                    n_arq2 = f"Termo_de_Nao_Homologacao_de_Modulos_do_CRTI_ERP"
                    c_docx2 = os.path.join(PASTA_TERMOS, f"{n_arq2}.docx")
                    doc2.save(c_docx2)
                    subprocess.run(f'libreoffice --headless --convert-to pdf --outdir "{PASTA_TERMOS}" "{c_docx2}"', shell=True, check=True)

                st.success("✨ Todos os documentos do lote foram gerados e armazenados!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar lote físico: {e}")

# --- PAINEL VISUAL DE DOWNLOAD E ZIP COMPACTADO ---
arquivos_gerados_base = glob.glob(os.path.join(PASTA_TERMOS, "*.pdf")) + glob.glob(os.path.join(PASTA_TERMOS, "*.docx"))

if arquivos_gerados_base:
    st.markdown("---")
    st.subheader("📥 Download do Pacote Completo Emitido")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for arq_caminho in arquivos_gerados_base:
            zip_file.write(arq_caminho, os.path.basename(arq_caminho))
    zip_buffer.seek(0)
    
    st.download_button(
        label="🎁 **BAIXAR TODOS OS TERMOS CONFIGURADOS (ZIP)**",
        data=zip_buffer,
        file_name=f"Pacote_Termos_HPTECH_{cliente_selecionado}.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    with st.expander("📄 Ver e baixar arquivos individuais do lote...", expanded=True):
        col_grade = st.columns(2)
        for idx, arq_caminho in enumerate(sorted(arquivos_gerados_base)):
            nome_real = os.path.basename(arq_caminho).replace("_", " ")
            with open(arq_caminho, "rb") as f_leitura: conteudo_bytes = f_leitura.read()
            
            col_alvo = col_grade[idx % 2]
            extensao_label = "PDF" if arq_caminho.endswith(".pdf") else "Word"
            mime_tipo = "application/pdf" if arq_caminho.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            
            col_alvo.download_button(
                label=f"📥 Baixar {extensao_label}: {nome_real}",
                data=conteudo_bytes,
                file_name=nome_real,
                mime=mime_tipo,
                use_container_width=True,
                key=f"btn_dl_{idx}"
            )

# =========================================================================
# 7. POP-UP DE CONFIRMAÇÃO DO DISPARO DO LOTE COMPLETO
# =========================================================================
@st.dialog(" Confirmação de Disparo de Termos")
def confirmar_envio_termos_popup_final(email, arquivos_lote):
    st.write("Você tem certeza que deseja disparar os termos gerados por e-mail?")
    st.write(f"• **Destinatários:** `{email}`")
    st.write(f"• **Arquivos em anexo:** PDF e Word de todos os termos gerados na tela")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("Sim, Disparar Lote", use_container_width=True):
            with st.spinner("Compilando lote completo de anexos e enviando..."):
                ok, r_msg = enviar_email_termos_logica_p04(
                    email, cliente_selecionado, data_virada_formatada, campos_fiscais_payload, arquivos_lote, gerente_cliente_name
                )
                if ok:
                    st.success(f" {r_msg}")
                    st.balloons()
                    time.sleep(4)
                else:
                    st.error(r_msg)
                    time.sleep(4)
                st.rerun()
                
    with col_p2:
        if st.button("Não, Cancelar", use_container_width=True):
            st.rerun()

# --- GATILHO DA SIDEBAR QUE CHAMA O POP-UP ---
if btn_enviar_emails:
    if not arquivos_gerados_base:
        st.sidebar.warning(" Gere os documentos na tela primeiro antes de disparar.")
    else:
        confirmar_envio_termos_popup_final(email_destinatario, arquivos_gerados_base)
