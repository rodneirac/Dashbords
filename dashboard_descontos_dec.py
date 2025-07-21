import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image
import calendar

st.set_page_config(page_title="Dashboard - Monitoramento Instruções", layout="wide")

# Exibe logomarca centralizada ao lado do título
col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    display_logo = Image.open("logo-supermix-pq.png")
    st.image(display_logo, width=180)

# Upload de arquivo
st.sidebar.title("📤 Upload de Planilha")
uploaded_file = st.sidebar.file_uploader("Envie a planilha .xlsx atualizada:", type=["xlsx"])

# Nome padrão da planilha no GitHub ou local
FILE_DEFAULT = "DADOSATUAL.XLSX"

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.sidebar.success("Arquivo carregado com sucesso!")
else:
    st.sidebar.warning(f"Usando planilha padrão: {FILE_DEFAULT}")
    df = pd.read_excel(FILE_DEFAULT)

# Preparo de colunas (datas, anos, meses)
if df["Data Criação"].dtype != 'datetime64[ns]':
    df["Data Criação"] = pd.to_datetime(df["Data Criação"])
df["Ano"] = df["Data Criação"].dt.year
df["Mês"] = df["Data Criação"].dt.month

# Sidebar de filtros
st.sidebar.title("📊 Filtros de Análise")
filial = st.sidebar.multiselect("Filial (Divisão)", options=sorted(df["Divisão"].dropna().unique()))
ano = st.sidebar.multiselect("Ano", options=sorted(df["Ano"].dropna().unique()))
meses_disponiveis = df["Mês"].dropna().unique() if not ano else df[df["Ano"].isin(ano)]["Mês"].dropna().unique()
meses_disponiveis = [int(m) for m in meses_disponiveis if pd.notna(m)]
meses_disponiveis.sort()
mes_nomes = [calendar.month_name[m] for m in meses_disponiveis]
mes_map = {calendar.month_name[m]: m for m in meses_disponiveis}
mes_nome = st.sidebar.multiselect("Mês", options=mes_nomes)
mes = [mes_map[m] for m in mes_nome] if mes_nome else None

# Filtros aplicados
df_filtrado = df.copy()
if filial:
    df_filtrado = df_filtrado[df_filtrado["Divisão"].isin(filial)]
if ano:
    df_filtrado = df_filtrado[df_filtrado["Ano"].isin(ano)]
if mes:
    df_filtrado = df_filtrado[df_filtrado["Mês"].isin(mes)]

# Funções auxiliares
def resumo_kpi(df, motivo, kpi_dias=False, kpi_valor=False):
    dados = df[df["Cód.Motivo"] == motivo].copy()
    qtd = len(dados)
    media_dias = dados["Dias"].mean() if kpi_dias else None
    soma_valor = dados["Desconto"].sum() if kpi_valor else None
    return qtd, media_dias, soma_valor

# KPIs principais (cards)
col1, col2, col3 = st.columns(3)
# PRL
qtd_prl, media_dias_prl, _ = resumo_kpi(df_filtrado, "PRL", kpi_dias=True)
col1.metric("Solicitações PRL", qtd_prl)
col1.metric("Média Dias PRL", f"{media_dias_prl:.1f}" if media_dias_prl else "-")
# ALT
qtd_alt, media_dias_alt, _ = resumo_kpi(df_filtrado, "ALT", kpi_dias=True)
col2.metric("Solicitações ALT", qtd_alt)
col2.metric("Média Dias ALT", f"{media_dias_alt:.1f}" if media_dias_alt else "-")
# DEC
qtd_dec, _, soma_dec = resumo_kpi(df_filtrado, "DEC", kpi_valor=True)
col3.metric("Solicitações DEC", qtd_dec)
col3.metric("Desconto Total DEC", f"R$ {soma_dec:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

st.markdown("---")

# Tabelas e gráficos por filial e nível 1

st.subheader("Resumo PRL (por Filial e Nível 1 Descrição)")
df_prl = df_filtrado[df_filtrado["Cód.Motivo"] == "PRL"]
tab_prl = df_prl.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Dias', 'count'),
    Soma_Dias=('Dias', 'sum'),
    Media_Dias=('Dias', 'mean')
).reset_index()
st.dataframe(tab_prl, use_container_width=True)
if not tab_prl.empty:
    fig_prl = px.bar(tab_prl, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                     title="Solicitações PRL por Filial e Nível 1")
    st.plotly_chart(fig_prl, use_container_width=True)

st.subheader("Resumo ALT (por Filial e Nível 1 Descrição)")
df_alt = df_filtrado[df_filtrado["Cód.Motivo"] == "ALT"]
tab_alt = df_alt.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Dias', 'count'),
    Soma_Dias=('Dias', 'sum'),
    Media_Dias=('Dias', 'mean')
).reset_index()
st.dataframe(tab_alt, use_container_width=True)
if not tab_alt.empty:
    fig_alt = px.bar(tab_alt, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                     title="Solicitações ALT por Filial e Nível 1")
    st.plotly_chart(fig_alt, use_container_width=True)

st.subheader("Resumo DEC (por Filial e Nível 1 Descrição)")
df_dec = df_filtrado[df_filtrado["Cód.Motivo"] == "DEC"]
tab_dec = df_dec.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Desconto', 'count'),
    Soma_Desconto=('Desconto', 'sum')
).reset_index()
st.dataframe(tab_dec, use_container_width=True)
if not tab_dec.empty:
    fig_dec = px.bar(tab_dec, x="Divisão", y="Soma_Desconto", color="Nível 1 Descrição", barmode="group",
                     title="Desconto DEC por Filial e Nível 1")
    st.plotly_chart(fig_dec, use_container_width=True)

st.markdown("---")
st.markdown("Relatório dinâmico por instrução: PRL (prorrogação), ALT (alteração) e DEC (desconto). Refine a análise usando os filtros laterais.")
