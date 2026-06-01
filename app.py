import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import folium
from streamlit_folium import st_folium
import json
import random

# --- CONFIGURAÇÃO ---
HEX_HOME = "#E47673"
st.set_page_config(page_title="Home Conteúdo - Dashboard", page_icon="🏠", layout="wide", initial_sidebar_state="collapsed")

# CSS Original
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; background-color: #f8f9fa; }}
    .stButton>button {{ background-color: {HEX_HOME}; color: white; border-radius: 12px; border: none; padding: 12px 24px; font-weight: 600; width: 100%; }}
    </style>
""", unsafe_allow_html=True)

# --- CONEXÃO COM GOOGLE ---
@st.cache_data(ttl=600)
def carregar_dados_do_sheets():
    creds_dict = json.loads(st.secrets["gcp_service_account"])
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1uuqIvekMUt3m7hKyFNXXUMZ0535Sz4deKfi1HRb9VLY").sheet1
    return pd.DataFrame(sheet.get_all_records())

# --- LÓGICA DE PROCESSAMENTO ---
def processar_dados(df):
    # Converte colunas numéricas para evitar erros
    cols_num = ["INSCRITOS YOUTUBE", "SEGUIDORES INSTAGRAM", "SEGUIDORES TIK TOK"]
    for col in cols_num:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

# --- APP ---
if 'page' not in st.session_state: st.session_state.page = 'entrada'

if st.session_state.page == 'entrada':
    st.title("🏠 Home Conteúdo - Mapeamento Estratégico")
    c_l, c_r = st.columns(2)
    if c_l.button("👥 ACESSO EQUIPE (ADMIN)"): st.session_state.page = 'admin'; st.rerun()
    if c_r.button("🔍 ACESSO VISITANTE"): st.session_state.page = 'visitante'; st.rerun()

elif st.session_state.page == 'admin':
    df_raw = carregar_dados_do_sheets()
    df = processar_dados(df_raw)
    
    st.title("Painel Administrativo")
    if st.button("⬅ Sair"): st.session_state.page = 'entrada'; st.rerun()
    
    tab1, tab2 = st.tabs(["📊 Inteligência de Dados", "➕ Novo Cadastro"])
    with tab1:
        st.dataframe(df, use_container_width=True)
    with tab2:
        st.subheader("Cadastrar Nova Influenciadora")
        # Criando inputs na ordem das colunas da planilha
        with st.form("novo_cadastro"):
            nome = st.text_input("Nome da Criadora")
            cidade = st.text_input("Cidade")
            uf = st.text_input("UF")
            # ... adicione os outros campos seguindo a ordem da sua lista ...
            submit = st.form_submit_button("Salvar")
            if submit:
                # Lista com as 34 posições (substitua os "" pelos valores dos inputs)
                nova_linha = ["", "", nome, cidade, uf] + [""] * 29 
                # Chame a função de salvar aqui
                st.success("Salvo!")

elif st.session_state.page == 'visitante':
    df = processar_dados(carregar_dados_do_sheets())
    st.title("Área de Visitante")
    st.dataframe(df[["CRIADORA", "CIDADE", "UF", "PORTE"]].head())
    if st.button("⬅ Voltar"): st.session_state.page = 'entrada'; st.rerun()
