import pandas as pd
import plotly.express as px
import streamlit as st

# Carregar os dados (substitua pelo seu caminho se for local)
df = pd.read_excel("EXPORT_20250604_114410.XLSX")

# Filtrar somente motivo DEC
df = df[df["Cód.Motivo"] == "DEC"].copy()

# Preparo de colunas
if df["Data Criação"].dtype != 'datetime64[ns]':
    df["Data Criação"] = pd.to_datetime(df["Data Criação"])
df["Mês/Ano"] = df["Data Criação"].dt.to_period("M").astype(str)

# Sidebar
st.sidebar.title("Filtros")
filial = st.sidebar.multiselect("Filial (Divisão)", options=sorted(df["Divisão"].unique()), default=None)
status = st.sidebar.multiselect("Status", options=sorted(df["Status"].unique()), default=None)

# Aplicar filtros
if filial:
    df = df[df["Divisão"].isin(filial)]
if status:
    df = df[df["Status"].isin(status)]

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
df_mensal = df.groupby("Mês/Ano")["Desconto"].sum().reset_index()
st.subheader("Evolução Mensal dos Descontos")
fig_mensal = px.line(df_mensal, x="Mês/Ano", y="Desconto", markers=True)
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
