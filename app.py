import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import random
import gspread
from google.oauth2.service_account import Credentials
import json

# Cores oficiais da Home Conteúdo
HEX_HOME = "#E47673"

st.set_page_config(page_title="Home Conteúdo - Dashboard", page_icon="🏠", layout="wide", initial_sidebar_state="collapsed")

# --- CONEXÃO GOOGLE SHEETS ---
@st.cache_resource
def conectar_sheets():
    creds_dict = json.loads(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    return client.open_by_key("1uuqIvekMUt3m7hKyFNXXUMZ0535Sz4deKfi1HRb9VLY").sheet1

# --- FUNÇÕES DE PROCESSAMENTO ---
@st.cache_data(ttl=600)
def carregar_dados_do_sheets():
    ws = conectar_sheets()
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Aplicar limpezas
    for col in df.columns:
        if "TAXA" in col.upper() or "ENGAJAMENTO" in col.upper():
            df[col] = pd.to_numeric(df[col].astype(str).str.replace("%", "").str.replace(",", "."), errors='coerce')
        elif col not in ["CRIADORA", "CIDADE", "UF", "PRINCIPAL REDE", "OBS", "E-mail de contato", "Telefone"]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

# --- INTERFACE ---
if 'page' not in st.session_state: st.session_state.page = 'entrada'

if st.session_state.page == 'entrada':
    st.title("🏠 Home Conteúdo - Mapeamento Estratégico")
    c_l, c_r = st.columns(2)
    if c_l.button("👥 ACESSO EQUIPE"): st.session_state.page = 'admin'; st.rerun()
    if c_r.button("🔍 ACESSO VISITANTE"): st.session_state.page = 'visitante'; st.rerun()

elif st.session_state.page == 'admin':
    df_base = carregar_dados_do_sheets()
    if st.button("⬅ Voltar"): st.session_state.page = 'entrada'; st.rerun()
    
    tab1, tab2 = st.tabs(["📊 Inteligência de Dados", "🗺️ Mapa"])
    
    with tab1:
        st.subheader("Filtros")
        uf_f = st.multiselect("UF", df_base["UF"].unique())
        df_f = df_base[df_base["UF"].isin(uf_f)] if uf_f else df_base
        st.dataframe(df_f, use_container_width=True)
        
    with tab2:
        st.subheader("Mapa de Atuação")
        # Aqui você coloca o seu código original do Folium, usando df_f
        st.write("Mapa em desenvolvimento...")

elif st.session_state.page == 'visitante':
    df_base = carregar_dados_do_sheets()
    if st.button("⬅ Voltar"): st.session_state.page = 'entrada'; st.rerun()
    st.subheader("Visualização Visitante")
    st.dataframe(df_base.head(10))
