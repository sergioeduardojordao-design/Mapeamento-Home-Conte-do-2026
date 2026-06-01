import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import os
import time
import folium
from streamlit_folium import st_folium
import random

# Configurações iniciais
HEX_HOME = "#E47673"
EXCEL_FILE = "banco_criadoras.xlsx"

st.set_page_config(
    page_title="Home Conteúdo - Dashboard",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilização CSS
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; background-color: #f8f9fa; }}
    .stTabs [data-baseweb="tab-panel"] {{
        background: rgba(255, 255, 255, 0.8); border-radius: 20px; padding: 25px;
        border: 1px solid rgba(255, 255, 255, 0.3); box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.05);
        backdrop-filter: blur(10px);
    }}
    .stButton>button {{
        background-color: {HEX_HOME}; color: white; border-radius: 12px; border: none;
        padding: 12px 24px; font-weight: 600; box-shadow: 0 4px 15px rgba(228, 118, 115, 0.3);
        transition: all 0.3s ease; width: 100%;
    }}
    .stButton>button:hover {{ transform: translateY(-2px); background-color: #d16260; }}
    .stDataFrame {{ border-radius: 15px; overflow: hidden; border: 1px solid #eee; }}
    .logo-container {{ display: flex; justify-content: center; align-items: center; flex-direction: column; padding: 20px; }}
    </style>
""", unsafe_allow_html=True)

# Dicionário de Coordenadas
COORDENADAS_UF = {
    "SP": (-23.5505, -46.6333), "RJ": (-22.9068, -43.1729), "MG": (-19.9167, -43.9345), "ES": (-20.2976, -40.2958),
    "PR": (-25.4284, -49.2733), "RS": (-30.0346, -51.2177), "SC": (-27.5954, -48.5480), "DF": (-15.7801, -47.9292),
    "GO": (-16.6869, -49.2648), "MT": (-15.6010, -56.0974), "MS": (-20.4428, -54.6464), "BA": (-12.9714, -38.5014),
    "PE": (-8.0543, -34.8813), "CE": (-3.7173, -38.5434), "MA": (-2.5297, -44.3028), "AL": (-9.6658, -35.7350),
    "PB": (-7.1150, -34.8631), "PI": (-5.0920, -42.8038), "RN": (-5.7945, -35.2110), "SE": (-10.9111, -37.0717),
    "AM": (-3.1190, -60.0217), "PA": (-1.4558, -48.4902), "RO": (-8.7612, -63.9039), "AP": (0.0340, -51.0694),
    "TO": (-10.1674, -48.3277), "AC": (-9.9740, -67.8076), "RR": (2.8197, -60.6732)
}

# Funções Utilitárias
def obter_coordenadas(uf):
    lat, lng = COORDENADAS_UF.get(str(uf).strip().upper(), (-14.2350, -51.9253))
    return lat + random.uniform(-0.15, 0.15), lng + random.uniform(-0.15, 0.15)

def limpar_valor_numerico(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).startswith('='): return 0
    try:
        return int(float(str(val).strip().replace(" ", "").replace(".", "").replace(",", "")))
    except: return 0

def limpar_taxa_float(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).startswith('='): return 0.0
    try:
        return float(str(val).replace("%", "").strip().replace(",", "."))
    except: return 0.0

def descobrir_linha_cabecalho():
    if not os.path.exists(EXCEL_FILE): return 1
    df_bruto = pd.read_excel(EXCEL_FILE, header=None)
    for i, row in df_bruto.iterrows():
        if "CRIADORA" in [str(x).upper().strip() for x in row.values]: return i
    return 0

def achar_ultima_linha_com_texto(ws):
    for row in range(ws.max_row, 0, -1):
        celula = ws.cell(row=row, column=3).value
        if celula and str(celula).strip() != "" and str(celula).upper() != "CRIADORA": return row
    return ws.max_row

def carregar_e_processar_dados():
    if not os.path.exists(EXCEL_FILE): return pd.DataFrame()
    try:
        linha_idx = descobrir_linha_cabecalho()
        wb_links = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
        ws_links = wb_links.active
        wb_dados = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        ws_dados = wb_dados.active
        
        colunas_nomes = [str(ws_dados.cell(linha_idx + 1, col).value).strip() for col in range(1, ws_dados.max_column + 1)]
        dados = []
        for r_idx in range(linha_idx + 2, ws_dados.max_row + 1):
            if str(ws_dados.cell(r_idx, 3).value).strip() in ["", "None", "CRIADORA"]: continue
            linha = {}
            for c_idx, c_name in enumerate(colunas_nomes, start=1):
                if not c_name or c_name == "None": continue
                cell_dados = ws_dados.cell(r_idx, c_idx)
                cell_links = ws_links.cell(r_idx, c_idx)
                if c_name in ["YOUTUBE", "INSTAGRAM", "TIK TOK", "FACEBOOK", "KWAI"] and cell_links.hyperlink:
                    linha[c_name] = cell_links.hyperlink.target
                else:
                    linha[c_name] = cell_dados.value if cell_dados.value is not None else ""
            dados.append(linha)
        
        df = pd.DataFrame(dados)
        df["PERFIL"] = df["PORTE"].astype(str).str.strip().replace(["", "None", "nan", "0", "0.0"], "Não Definido") if "PORTE" in df.columns else "Não Definido"
        
        # Normalização
        for col in ["CRIADORA", "CIDADE", "UF", "PRINCIPAL REDE", "E-mail de contato", "Telefone", "OBS"]:
            if col in df.columns: df[col] = df[col].fillna("").astype(str).str.strip().replace(["None", "nan", "0", "0.0"], "")
            
        df["Link YouTube"] = df["YOUTUBE"].fillna("").astype(str) if "YOUTUBE" in df.columns else ""
        df["Link Instagram"] = df["INSTAGRAM"].fillna("").astype(str) if "INSTAGRAM" in df.columns else ""
        df["Link TikTok"] = df["TIK TOK"].fillna("").astype(str) if "TIK TOK" in df.columns else ""
        
        for col in df.columns:
            if col not in ["YOUTUBE", "INSTAGRAM", "TIK TOK", "CRIADORA", "CIDADE", "UF", "COMPLEMENTO", "E-mail de contato", "Telefone", "OBS", "PORTE", "PERFIL", "PRINCIPAL REDE", "Link YouTube", "Link Instagram", "Link TikTok"]:
                if "TAXA" not in col.upper() and "ENGAJAMENTO" not in col.upper():
                    df[col] = df[col].apply(limpar_valor_numerico)
                else:
                    df[col] = df[col].apply(limpar_taxa_float)
        return df
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
        return pd.DataFrame()

def salvar_novo_registro_excel(dados):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        linha_idx = descobrir_linha_cabecalho()
        colunas_nomes = [str(ws.cell(linha_idx + 1, col).value).strip() for col in range(1, ws.max_column + 1)]
        proxima_linha = achar_ultima_linha_com_texto(ws) + 1
        font_link = Font(color="0563C1", underline="single")
        
        for idx, col_name in enumerate(colunas_nomes, start=1):
            if col_name == "CRIADORA": ws.cell(proxima_linha, idx, dados['nome'])
            elif col_name == "OBS": ws.cell(proxima_linha, idx, dados['obs'])
            elif col_name == "CIDADE": ws.cell(proxima_linha, idx, dados['cidade'])
            elif col_name == "UF": ws.cell(proxima_linha, idx, dados['uf'])
            elif col_name == "PORTE": ws.cell(proxima_linha, idx, dados['porte_manual'])
            elif col_name == "PRINCIPAL REDE": ws.cell(proxima_linha, idx, dados['rede_manual'])
            elif col_name == "YOUTUBE" and dados['rede_detectada'] == "YouTube":
                c = ws.cell(proxima_linha, idx, dados['nome']); c.hyperlink = dados['link_principal']; c.font = font_link
            elif col_name == "INSTAGRAM" and dados['rede_detectada'] == "Instagram":
                c = ws.cell(proxima_linha, idx, dados['nome']); c.hyperlink = dados['link_principal']; c.font = font_link
            elif col_name == "TIK TOK" and dados['rede_detectada'] == "TikTok":
                c = ws.cell(proxima_linha, idx, dados['nome']); c.hyperlink = dados['link_principal']; c.font = font_link
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False

# Lógica da Aplicação
if 'page' not in st.session_state: st.session_state.page = 'entrada'

if st.session_state.page == 'entrada':
    st.markdown("<div class='logo-container'>", unsafe_allow_html=True)
    if os.path.exists("logotipo.png"): st.image("logotipo.png", width=280)
    st.title("🏠 Home Conteúdo - Mapeamento Estratégico")
    c_l, c_r = st.columns(2)
    if c_l.button("👥 ACESSO EQUIPE"): st.session_state.page = 'admin'; st.rerun()
    if c_r.button("🔍 ACESSO VISITANTE"): st.session_state.page = 'visitante'; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

elif st.session_state.page == 'admin':
    df_base = carregar_e_processar_dados()
    if st.button("⬅ Sair"): st.session_state.page = 'entrada'; st.rerun()
    if df_base.empty: st.error("Planilha vazia ou não encontrada.")
    else:
        tab1, tab2, tab3 = st.tabs(["📊 Dados", "➕ Novo", "📈 Relatórios"])
        with tab1:
            busca = st.text_input("Buscar:")
            df_f = df_base[df_base.apply(lambda row: row.astype(str).str.contains(busca, case=False).any(), axis=1)] if busca else df_base
            st.dataframe(df_f, use_container_width=True)
        with tab2:
            st.subheader("Cadastro")
            # Adicione aqui os campos de input do seu formulário
            if st.button("💾 Salvar"): pass
        with tab3:
            st.metric("Total de Criadoras", len(df_base))

elif st.session_state.page == 'visitante':
    df_base = carregar_e_processar_dados()
    if st.button("⬅ Voltar"): st.session_state.page = 'entrada'; st.rerun()
    st.write("Visualização restrita para visitantes.")
    st.dataframe(df_base.head(5), use_container_width=True)A
