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

# Linhas para cada motivo
df_prl = df_filtrado[df_filtrado["Cód.Motivo"] == "PRL"]
df_dec = df_filtrado[df_filtrado["Cód.Motivo"] == "DEC"]
df_alt = df_filtrado[df_filtrado["Cód.Motivo"] == "ALT"]
df_bxs = df_filtrado[df_filtrado["Cód.Motivo"] == "BXS"]
df_can = df_filtrado[df_filtrado["Cód.Motivo"] == "CAN"]
df_ref = df_filtrado[df_filtrado["Cód.Motivo"] == "REF"]

# KPIs (cards)
qtd_prl_card = len(df_prl) + len(df_alt)
media_dias_prl_alt = pd.concat([df_prl, df_alt])["Dias"].mean() if not pd.concat([df_prl, df_alt]).empty else None
qtd_dec_card = len(df_dec) + len(df_alt)
desconto_total_dec_alt = pd.concat([df_dec, df_alt])["Desconto"].sum()
qtd_bxs = len(df_bxs)
desconto_total_bxs = df_bxs["Desconto"].sum()
df_cancel = pd.concat([df_can, df_ref])
qtd_cancel = len(df_cancel)
montante_cancel = df_cancel["Montante"].sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Qtde Prorrogações", qtd_prl_card)
col1.metric("Média dias Prorrogações", f"{media_dias_prl_alt:.1f}" if media_dias_prl_alt else "-")
col2.metric("Qtde Solicitações Desc/Abatimentos", qtd_dec_card)
col2.metric(
    "Desconto Total (DEC + ALT)",
    f"R$ {desconto_total_dec_alt:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)
col3.metric("Solicitações BXS", qtd_bxs)
col3.metric(
    "Montante BXS",
    f"R$ {desconto_total_bxs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)
col4.metric("Cancelamentos (CAN + REF)", qtd_cancel)
col4.metric(
    "Montante Cancelado",
    f"R$ {montante_cancel:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)

st.markdown("---")

# Helper para formatar coluna como reais
def format_reais(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Tabelas e gráficos PRL + ALT
st.subheader("Resumo PRL + ALT (por Filial e Nível 1 Descrição)")
df_prl_alt = pd.concat([df_prl, df_alt])
tab_prl_alt = df_prl_alt.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Dias', 'count'),
    Soma_Dias=('Dias', 'sum'),
    Media_Dias=('Dias', 'mean')
).reset_index()
# Ordenar do maior para o menor soma_dias
tab_prl_alt = tab_prl_alt.sort_values("Soma_Dias", ascending=False)
# Formatar valores
tab_prl_alt["Soma_Dias"] = tab_prl_alt["Soma_Dias"].apply(lambda x: f"{x:.0f}")
tab_prl_alt["Media_Dias"] = tab_prl_alt["Media_Dias"].apply(lambda x: f"{x:.1f}")
st.dataframe(tab_prl_alt, use_container_width=True)
if not tab_prl_alt.empty:
    fig_prl_alt = px.bar(tab_prl_alt, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                         title="Solicitações PRL + ALT por Filial e Nível 1")
    st.plotly_chart(fig_prl_alt, use_container_width=True)

# Tabelas e gráficos DEC + ALT
st.subheader("Resumo DEC + ALT (por Filial e Nível 1 Descrição)")
df_dec_alt = pd.concat([df_dec, df_alt])
tab_dec_alt = df_dec_alt.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Desconto', 'count'),
    Soma_Desconto=('Desconto', 'sum')
).reset_index()
tab_dec_alt = tab_dec_alt.sort_values("Soma_Desconto", ascending=False)
tab_dec_alt["Soma_Desconto"] = tab_dec_alt["Soma_Desconto"].apply(format_reais)
st.dataframe(tab_dec_alt, use_container_width=True)
if not tab_dec_alt.empty:
    fig_dec_alt = px.bar(tab_dec_alt, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                         title="Solicitações DEC + ALT por Filial e Nível 1")
    st.plotly_chart(fig_dec_alt, use_container_width=True)

# Tabelas e gráficos BXS
st.subheader("Resumo BXS (por Filial e Nível 1 Descrição)")
tab_bxs = df_bxs.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Desconto', 'count'),
    Soma_Desconto=('Desconto', 'sum')
).reset_index()
tab_bxs = tab_bxs.sort_values("Soma_Desconto", ascending=False)
tab_bxs["Soma_Desconto"] = tab_bxs["Soma_Desconto"].apply(format_reais)
st.dataframe(tab_bxs, use_container_width=True)
if not tab_bxs.empty:
    fig_bxs = px.bar(tab_bxs, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                     title="Solicitações BXS por Filial e Nível 1")
    st.plotly_chart(fig_bxs, use_container_width=True)

# Tabelas e gráficos Cancelamentos
st.subheader("Resumo Cancelamentos (CAN + REF) (por Filial e Nível 1 Descrição)")
tab_cancel = df_cancel.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Montante', 'count'),
    Soma_Montante=('Montante', 'sum')
).reset_index()
tab_cancel = tab_cancel.sort_values("Soma_Montante", ascending=False)
tab_cancel["Soma_Montante"] = tab_cancel["Soma_Montante"].apply(format_reais)
st.dataframe(tab_cancel, use_container_width=True)
if not tab_cancel.empty:
    fig_cancel = px.bar(tab_cancel, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                        title="Solicitações Canceladas por Filial e Nível 1")
    st.plotly_chart(fig_cancel, use_container_width=True)

st.markdown("---")
st.markdown("Relatório dinâmico por instrução: PRL (prorrogação), ALT (alteração), DEC (desconto), BXS (baixa) e Cancelamentos (CAN/REF). Refine a análise usando os filtros laterais.")
