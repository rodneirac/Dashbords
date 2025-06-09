import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image
from io import BytesIO
import calendar

st.set_page_config(page_title="Dashboard - Descontos DEC", layout="wide")

# Exibe logomarca centralizada ao lado do título
col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    display_logo = Image.open("logo-supermix-pq.png")
    st.image(display_logo, width=180)

# Upload de arquivo
st.sidebar.title("📤 Upload de Planilha")
uploaded_file = st.sidebar.file_uploader("Envie a planilha .xlsx atualizada:", type=["xlsx"])

# Lê o arquivo enviado ou usa padrão
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.sidebar.success("Arquivo carregado com sucesso!")
else:
    st.sidebar.warning("Usando planilha padrão: EXPORT_20250604_114410.XLSX")
    df = pd.read_excel("EXPORT_20250604_114410.XLSX")

# Filtrar somente motivo DEC
df = df[df["Cód.Motivo"] == "DEC"].copy()

# Preparo de colunas
if df["Data Criação"].dtype != 'datetime64[ns]':
    df["Data Criação"] = pd.to_datetime(df["Data Criação"])
df["Ano"] = df["Data Criação"].dt.year
df["Mês"] = df["Data Criação"].dt.month

# Sidebar
st.sidebar.title("📊 Filtros de Análise")
filial = st.sidebar.multiselect("Filial (Divisão)", options=sorted(df["Divisão"].unique()), default=None)
ano = st.sidebar.multiselect("Ano", options=sorted(df["Ano"].unique()), default=None)

# Filtro dependente de mês baseado no ano
if ano:
    meses_filtrados = df[df["Ano"].isin(ano)]["Mês"].unique()
else:
    meses_filtrados = df["Mês"].unique()

meses_filtrados.sort()
mes_map = {calendar.month_name[m]: m for m in meses_filtrados if m in calendar.month_name}
mes_nomes = [calendar.month_name[m] for m in meses_filtrados if m in calendar.month_name]
mes_nome = st.sidebar.multiselect("Mês", options=mes_nomes, default=None)
mes = [mes_map[m] for m in mes_nome] if mes_nome else None

situacao = st.sidebar.multiselect("Situação (coluna W)", options=sorted(df[df.columns[22]].dropna().unique()), default=None)

# Aplicar filtros
if filial:
    df = df[df["Divisão"].isin(filial)]
if ano:
    df = df[df["Ano"].isin(ano)]
if mes:
    df = df[df["Mês"].isin(mes)]
if situacao:
    df = df[df[df.columns[22]].isin(situacao)]

# KPIs
total_desconto = df["Desconto"].sum()
total_solicitacoes = len(df)
total_filiais = df["Divisão"].nunique()

# Título
st.title("Dashboard Interativo - Descontos Motivo DEC")
st.markdown("---")

# KPIs
col1, col2, col3 = st.columns(3)
col1.metric("Desconto Total", f"R$ {total_desconto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
col2.metric("Solicitações", total_solicitacoes)
col3.metric("Filiais Envolvidas", total_filiais)

# Gráfico de barras - Desconto por filial
st.subheader("Desconto Total por Filial")
fig_filial = px.bar(df.groupby("Divisão")["Desconto"].sum().reset_index().sort_values("Desconto", ascending=False),
                    x="Divisão", y="Desconto", text_auto=True)
st.plotly_chart(fig_filial, use_container_width=True)

# Linha do tempo - Evolução mensal
df["Mês/Ano"] = pd.to_datetime(df["Data Criação"].dt.to_period("M").astype(str))
df_mensal = df.groupby("Mês/Ano")["Desconto"].sum().reset_index()
df_mensal = df_mensal.sort_values("Mês/Ano")

st.subheader("Evolução Mensal dos Descontos")
fig_mensal = go.Figure()
fig_mensal.add_trace(go.Bar(x=df_mensal["Mês/Ano"], y=df_mensal["Desconto"], name="Desconto"))
fig_mensal.add_trace(go.Scatter(x=df_mensal["Mês/Ano"], y=df_mensal["Desconto"].rolling(3).mean(),
                                mode='lines+markers', name="Média Móvel 3 meses", line=dict(color='orange')))
fig_mensal.update_layout(xaxis_title="Mês/Ano", yaxis_title="Desconto", hovermode="x unified")
st.plotly_chart(fig_mensal, use_container_width=True)

# Pizza - Nível 1
st.subheader("Distribuição por Nível 1")
fig_n1 = px.pie(df, names="Nível 1 Descrição", values="Desconto", hole=0.4)
st.plotly_chart(fig_n1, use_container_width=True)

# Tabela por motivo detalhado
st.subheader("Tabela por Nível 1 + Nível 2")
tab_motivo = df.groupby(["Nível 1 Descrição", "Nível 2 Descrição"])["Desconto"].sum().reset_index().sort_values("Desconto", ascending=False)
st.dataframe(tab_motivo, use_container_width=True)

# Ranking por usuário
st.subheader("Ranking de Usuários")
rk_user = df.groupby("Usuário")["Desconto"].sum().reset_index().sort_values("Desconto", ascending=False)
st.dataframe(rk_user, use_container_width=True)

st.markdown("---")
st.markdown("Relatório dinâmico criado via Streamlit. Para análise completa, aplique filtros na barra lateral.")
