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
    return client.open("Mapeamento 2026").sheet1

# --- FUNÇÕES DE APOIO ---
def limpar_valor_numerico(val):
    try: return int(float(str(val).replace(" ", "").replace(".", "").replace(",", "")))
    except: return 0

def limpar_taxa_float(val):
    try: return float(str(val).replace("%", "").replace(",", "."))
    except: return 0.0

@st.cache_data(ttl=300)
def carregar_dados():
    ws = conectar_sheets()
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    # Limpeza básica e conversão de tipos
    for col in df.columns:
        if "TAXA" in col.upper() or "ENGAJAMENTO" in col.upper():
            df[col] = df[col].apply(limpar_taxa_float)
    return df

# --- INTERFACE ---
if 'page' not in st.session_state: st.session_state.page = 'entrada'

if st.session_state.page == 'entrada':
    st.title("🏠 Home Conteúdo - Mapeamento Estratégico")
    c_l, c_r = st.columns(2)
    if c_l.button("👥 ACESSO EQUIPE"): st.session_state.page = 'admin'; st.rerun()
    if c_r.button("🔍 ACESSO VISITANTE"): st.session_state.page = 'visitante'; st.rerun()

elif st.session_state.page == 'admin':
    df_base = carregar_dados()
    if st.button("⬅ Voltar"): st.session_state.page = 'entrada'; st.rerun()
    st.title("Painel Administrativo")
    st.dataframe(df_base, use_container_width=True)

elif st.session_state.page == 'visitante':
    df_base = carregar_dados()
    if st.button("⬅ Voltar"): st.session_state.page = 'entrada'; st.rerun()
    st.title("🔍 Área do Visitante")
    st.dataframe(df_base.head(5), use_container_width=True)
