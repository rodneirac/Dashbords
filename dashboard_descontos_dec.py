import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image
import calendar

st.set_page_config(page_title="Análise Instruções Bancárias", layout="wide")

# Logo e título
col_logo1, col_logo2, col_logo3 = st.columns([1, 6, 1])
with col_logo2:
    display_logo = Image.open("logo-supermix-pq.png")
    st.image(display_logo, width=180)
    st.markdown("<h1 style='text-align: center; margin-top: 0;'>Análise Instruções Bancárias</h1>", unsafe_allow_html=True)
st.markdown("---")

# Nome padrão da planilha (leitura fixa - no github/local)
FILE_DEFAULT = "DADOSATUAL.XLSX"
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

# Função para KPIs PRL + ALT
def resumo_prl_alt(df):
    df_prl_alt = df[df["Cód.Motivo"].isin(["PRL", "ALT"])].copy()
    qtd = len(df_prl_alt)
    media_dias = df_prl_alt["Dias"].mean() if not df_prl_alt.empty else None
    return qtd, media_dias

# Função para KPIs DEC + ALT desconto
def resumo_desconto_dec_alt(df):
    return df[df["Cód.Motivo"].isin(["DEC", "ALT"])]["Desconto"].sum()

# KPIs principais (cards)
col1, col2 = st.columns(2)
# PRL + ALT
qtd_prl_alt, media_dias_prl_alt = resumo_prl_alt(df_filtrado)
col1.metric("Solicitações PRL + ALT", qtd_prl_alt)
col1.metric("Média Dias PRL + ALT", f"{media_dias_prl_alt:.1f}" if media_dias_prl_alt else "-")
# DEC + ALT Desconto
qtd_dec = len(df_filtrado[df_filtrado["Cód.Motivo"] == "DEC"])
desconto_total_dec_alt = resumo_desconto_dec_alt(df_filtrado)
col2.metric("Solicitações DEC", qtd_dec)
col2.metric(
    "Desconto Total (DEC + ALT)",
    f"R$ {desconto_total_dec_alt:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)

st.markdown("---")

# Tabelas e gráficos por filial e nível 1

st.subheader("Resumo PRL + ALT (por Filial e Nível 1 Descrição)")
df_prl_alt = df_filtrado[df_filtrado["Cód.Motivo"].isin(["PRL", "ALT"])]
tab_prl_alt = df_prl_alt.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Dias', 'count'),
    Soma_Dias=('Dias', 'sum'),
    Media_Dias=('Dias', 'mean')
).reset_index()
st.dataframe(tab_prl_alt, use_container_width=True)
if not tab_prl_alt.empty:
    fig_prl_alt = px.bar(tab_prl_alt, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                         title="Solicitações PRL + ALT por Filial e Nível 1")
    st.plotly_chart(fig_prl_alt, use_container_width=True)

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
