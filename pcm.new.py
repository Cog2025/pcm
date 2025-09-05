import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import gspread
from google.oauth2.service_account import Credentials
import os

# ==============================
# CONFIGURAÇÕES
# ==============================
st.set_page_config(page_title="Inventário de Equipamentos", layout="wide")
st.title("📊 Inventário de Equipamentos")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
CREDS_FILE = "google_credentials.json"  # precisa do JSON da conta de serviço
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1r-UFzs0pDAixrT6hk8c9X6Ax8J9-AnJ8hD5twaRRECI/edit?usp=sharing"
PLANILHA_EQUIPAMENTOS = "Modelo Árvore de Ativos"

# ==============================
# CONEXÃO COM GOOGLE SHEETS
# ==============================
def connect_to_google_sheets():
    if os.path.exists(CREDS_FILE):
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def fetch_sheet_as_df(worksheet):
    data = worksheet.get_all_values()
    if not data:
        return pd.DataFrame()
    headers = [h.strip() for h in data.pop(0)]
    return pd.DataFrame(data, columns=headers)

# ==============================
# CARREGAR DADOS
# ==============================
client = connect_to_google_sheets()
workbook = client.open_by_url(SPREADSHEET_URL)
worksheet = workbook.worksheet(PLANILHA_EQUIPAMENTOS)
df = fetch_sheet_as_df(worksheet)

st.success("✅ Dados carregados da planilha do Google Sheets!")

# ==============================
# FILTROS
# ==============================
st.sidebar.header("Filtros")
filtros = {}
for col in df.columns:
    if df[col].dropna().nunique() > 0:
        selecao = st.sidebar.multiselect(
            f"Filtrar {col}", options=sorted(df[col].dropna().unique().tolist())
        )
        if selecao:
            filtros[col] = selecao

df_filtrado = df.copy()
for col, valores in filtros.items():
    df_filtrado = df_filtrado[df_filtrado[col].isin(valores)]

# ==============================
# EXIBIR TABELA
# ==============================
st.subheader("📋 Tabela de Equipamentos")
st.dataframe(df_filtrado, use_container_width=True)

# ==============================
# GRÁFICO
# ==============================
st.subheader("📈 Análises Gráficas")
colunas_categoricas = [col for col in df_filtrado.columns if df_filtrado[col].dtype == "object"]
if colunas_categoricas:
    coluna_grafico = st.selectbox("Escolha uma coluna para visualizar a distribuição:", colunas_categoricas)
    contagem = df_filtrado[coluna_grafico].value_counts().reset_index()
    contagem.columns = [coluna_grafico, "Quantidade"]
    fig = px.bar(contagem, x=coluna_grafico, y="Quantidade", title=f"Distribuição por {coluna_grafico}")
    st.plotly_chart(fig, use_container_width=True)

# ==============================
# FORMULÁRIO NOVO EQUIPAMENTO
# ==============================
st.subheader("➕ Adicionar Novo Equipamento")
with st.form("novo_equipamento"):
    nome = st.text_input("Nome")
    codigo = st.text_input("Código")
    codigo_pai = st.text_input("Código Pai/Local")
    fabricante = st.text_input("Fabricante")
    modelo = st.text_input("Modelo")
    tipo = st.text_input("Tipo")
    classificacao1 = st.text_input("Classificação 1")
    classificacao2 = st.text_input("Classificação 2")
    peso = st.text_input("Peso")
    outro2 = st.text_input("Outro 2")
    serial = st.text_input("Serial")
    centro_custos = st.text_input("Centro de Custos")
    criticidade = st.text_input("CRITICIDADE")
    media_uso = st.text_input("Média diária de uso do equipamento")

    submitted = st.form_submit_button("Adicionar")

    if submitted:
        novo_registro = [
            nome, codigo, codigo_pai, fabricante, modelo, tipo,
            classificacao1, classificacao2, peso, outro2, serial,
            centro_custos, criticidade, media_uso
        ]
        worksheet.append_row(novo_registro, value_input_option="USER_ENTERED")
        st.success("✅ Novo equipamento adicionado com sucesso!")

# ==============================
# EXPORTAR XLSX
# ==============================
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Equipamentos")
    return output.getvalue()

excel_data = to_excel(df_filtrado)
st.download_button(
    label="📥 Baixar tabela filtrada (Excel)",
    data=excel_data,
    file_name="equipamentos_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
