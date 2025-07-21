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

# ========== Agrupamento dos motivos ==========
def agrupa_motivo(cod):
    cod = str(cod).strip().upper()
    if cod in ["XXX", "CEN", "DEV", "YYY"]:
        return None
    elif cod in ["PRL", "ALT"]:
        return "Prorrogação"
    elif cod == "DEC":
        return "Desconto/Abat."
    elif cod in ["CAN", "REF"]:
        return "Cancelamento"
    elif cod == "BXS":
        return "Baixa de Saldo"
    else:
        return cod  # Se surgir novo código não previsto

df["Motivo Agrupado"] = df["Cód.Motivo"].apply(agrupa_motivo)

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

# Filtrar só motivos válidos para o filtro (eliminando os ocultos)
motivos_disponiveis = [m for m in sorted(df["Motivo Agrupado"].dropna().unique()) if m not in [None]]
motivo = st.sidebar.multiselect("Motivo", options=motivos_disponiveis)

# Filtros aplicados
df_filtrado = df[df["Motivo Agrupado"].notna()].copy()
if filial:
    df_filtrado = df_filtrado[df_filtrado["Divisão"].isin(filial)]
if ano:
    df_filtrado = df_filtrado[df_filtrado["Ano"].isin(ano)]
if mes:
    df_filtrado = df_filtrado[df_filtrado["Mês"].isin(mes)]
if motivo:
    df_filtrado = df_filtrado[df_filtrado["Motivo Agrupado"].isin(motivo)]

# ==== KPIs agrupando conforme regras ====
# Prorrogação = PRL + ALT
df_prorrog = df_filtrado[df_filtrado["Motivo Agrupado"] == "Prorrogação"]
# Desconto/Abat = DEC + ALT
df_desc_abat = df_filtrado[df_filtrado["Motivo Agrupado"] == "Desconto/Abat."]
# Baixa de Saldo = BXS
df_baixa = df_filtrado[df_filtrado["Motivo Agrupado"] == "Baixa de Saldo"]
# Cancelamento = CAN + REF
df_cancel = df_filtrado[df_filtrado["Motivo Agrupado"] == "Cancelamento"]

qtd_prl_card = len(df_prorrog)
media_dias_prl = df_prorrog["Dias"].mean() if not df_prorrog.empty else None
qtd_desc_card = len(df_desc_abat)
desconto_total_dec_abat = df_desc_abat["Desconto"].sum() if not df_desc_abat.empty else 0
qtd_bxs = len(df_baixa)
desconto_total_bxs = df_baixa["Desconto"].sum() if not df_baixa.empty else 0
qtd_cancel = len(df_cancel)
montante_cancel = df_cancel["Montante"].sum() if not df_cancel.empty else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("Solicitações Prorrogação", qtd_prl_card)
col1.metric("Média Dias Prorrogação", f"{media_dias_prl:.1f}" if media_dias_prl else "-")
col2.metric("Solicitações Desconto/Abat.", qtd_desc_card)
col2.metric(
    "Desconto Total (Desconto/Abat.)",
    f"R$ {desconto_total_dec_abat:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)
col3.metric("Solicitações Baixa de Saldo", qtd_bxs)
col3.metric(
    "Montante Baixa de Saldo",
    f"R$ {desconto_total_bxs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)
col4.metric("Cancelamentos", qtd_cancel)
col4.metric(
    "Montante Cancelado",
    f"R$ {montante_cancel:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)

st.markdown("---")

# GRÁFICO DE RANKING GERAL COM TOOLTIP DETALHADO POR MOTIVO AGRUPADO
st.subheader("Ranking de Filiais por Total de Solicitações (Agrupado por Tipo)")

# Pivot para tooltip detalhado
tooltip_pivot = (
    df_filtrado
    .pivot_table(index="Divisão", columns="Motivo Agrupado", values="Data Criação", aggfunc="count", fill_value=0)
    .reset_index()
)
tooltip_pivot["Qtde_Solicitações"] = tooltip_pivot.drop(columns=["Divisão"]).sum(axis=1)

# Ordenar do maior para menor
tooltip_pivot = tooltip_pivot.sort_values("Qtde_Solicitações", ascending=False)

motivos_agrup = [col for col in tooltip_pivot.columns if col not in ["Divisão", "Qtde_Solicitações"]]

fig_qtde = px.bar(
    tooltip_pivot,
    x="Divisão",
    y="Qtde_Solicitações",
    text="Qtde_Solicitações",
    title="Ranking de Filiais por Total de Solicitações",
    color_discrete_sequence=["#800040"],
    hover_data={m: True for m in motivos_agrup} | {"Qtde_Solicitações": True, "Divisão": True}
)
hovertemplate = "<b>Filial:</b> %{x}<br><b>Total de Solicitações:</b> %{y}<br>"
for m in motivos_agrup:
    hovertemplate += f"<b>{m}:</b> %{{customdata[{motivos_agrup.index(m)}]}}<br>"
fig_qtde.update_traces(
    texttemplate='%{text}',
    textposition='outside',
    hovertemplate=hovertemplate
)
fig_qtde.update_layout(
    yaxis_title="Total de Solicitações",
    xaxis_title="Filial",
    uniformtext_minsize=8,
    uniformtext_mode='hide'
)
st.plotly_chart(fig_qtde, use_container_width=True)

# ==== GRÁFICOS DE PIZZA POR NÍVEL 1 DESCRIÇÃO (COM "EFEITO 3D") ====
st.subheader("Distribuição por Nível 1 Descrição")

# Prorrogação
col_pie1, col_pie2 = st.columns(2)
with col_pie1:
    pizza_prl = (
        df_prorrog.groupby("Nível 1 Descrição")
        .size()
        .reset_index(name="Qtde")
        .sort_values("Qtde", ascending=False)
    )
    fig_pie_prl = px.pie(
        pizza_prl,
        names="Nível 1 Descrição",
        values="Qtde",
        hole=0.4,
        title="Prorrogação"
    )
    fig_pie_prl.update_traces(textinfo='percent+label', pull=[0.08]*len(pizza_prl))
    st.plotly_chart(fig_pie_prl, use_container_width=True)

# Desconto/Abatimento
with col_pie2:
    pizza_desc_abat = (
        df_desc_abat.groupby("Nível 1 Descrição")["Desconto"]
        .sum()
        .reset_index()
        .sort_values("Desconto", ascending=False)
    )
    fig_pie_desc_abat = px.pie(
        pizza_desc_abat,
        names="Nível 1 Descrição",
        values="Desconto",
        hole=0.4,
        title="Desconto/Abat."
    )
    fig_pie_desc_abat.update_traces(textinfo='percent+label', pull=[0.08]*len(pizza_desc_abat))
    st.plotly_chart(fig_pie_desc_abat, use_container_width=True)

# Baixa de Saldo
col_pie3, col_pie4 = st.columns(2)
with col_pie3:
    pizza_baixa = (
        df_baixa.groupby("Nível 1 Descrição")["Desconto"]
        .sum()
        .reset_index()
        .sort_values("Desconto", ascending=False)
    )
    fig_pie_baixa = px.pie(
        pizza_baixa,
        names="Nível 1 Descrição",
        values="Desconto",
        hole=0.4,
        title="Baixa de Saldo"
    )
    fig_pie_baixa.update_traces(textinfo='percent+label', pull=[0.08]*len(pizza_baixa))
    st.plotly_chart(fig_pie_baixa, use_container_width=True)

# Cancelamento
with col_pie4:
    pizza_cancel = (
        df_cancel.groupby("Nível 1 Descrição")["Montante"]
        .sum()
        .reset_index()
        .sort_values("Montante", ascending=False)
    )
    fig_pie_cancel = px.pie(
        pizza_cancel,
        names="Nível 1 Descrição",
        values="Montante",
        hole=0.4,
        title="Cancelamento"
    )
    fig_pie_cancel.update_traces(textinfo='percent+label', pull=[0.08]*len(pizza_cancel))
    st.plotly_chart(fig_pie_cancel, use_container_width=True)

# Helper para formatar coluna como reais
def format_reais(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Tabelas e gráficos Desconto/Abatimento
st.subheader("Resumo Desconto/Abat. (por Filial e Nível 1 Descrição)")
tab_desc_abat = df_desc_abat.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Desconto', 'count'),
    Soma_Desconto=('Desconto', 'sum')
).reset_index()
tab_desc_abat = tab_desc_abat.sort_values("Soma_Desconto", ascending=False)
tab_desc_abat["Soma_Desconto"] = tab_desc_abat["Soma_Desconto"].apply(format_reais)
st.dataframe(tab_desc_abat, use_container_width=True)
if not tab_desc_abat.empty:
    fig_desc_abat = px.bar(tab_desc_abat, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                         title="Solicitações Desconto/Abat. por Filial e Nível 1")
    st.plotly_chart(fig_desc_abat, use_container_width=True)

# Tabelas e gráficos Baixa de Saldo
st.subheader("Resumo Baixa de Saldo (por Filial e Nível 1 Descrição)")
tab_baixa = df_baixa.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Desconto', 'count'),
    Soma_Desconto=('Desconto', 'sum')
).reset_index()
tab_baixa = tab_baixa.sort_values("Soma_Desconto", ascending=False)
tab_baixa["Soma_Desconto"] = tab_baixa["Soma_Desconto"].apply(format_reais)
st.dataframe(tab_baixa, use_container_width=True)
if not tab_baixa.empty:
    fig_baixa = px.bar(tab_baixa, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                     title="Solicitações Baixa de Saldo por Filial e Nível 1")
    st.plotly_chart(fig_baixa, use_container_width=True)

# Tabelas e gráficos Cancelamento
st.subheader("Resumo Cancelamento (por Filial e Nível 1 Descrição)")
tab_cancel = df_cancel.groupby(["Divisão", "Nível 1 Descrição"]).agg(
    Qtde=('Montante', 'count'),
    Soma_Montante=('Montante', 'sum')
).reset_index()
tab_cancel = tab_cancel.sort_values("Soma_Montante", ascending=False)
tab_cancel["Soma_Montante"] = tab_cancel["Soma_Montante"].apply(format_reais)
st.dataframe(tab_cancel, use_container_width=True)
if not tab_cancel.empty:
    fig_cancel = px.bar(tab_cancel, x="Divisão", y="Qtde", color="Nível 1 Descrição", barmode="group",
                        title="Solicitações Cancelamento por Filial e Nível 1")
    st.plotly_chart(fig_cancel, use_container_width=True)

st.markdown("---")
st.markdown("Relatório dinâmico por instrução: Prorrogação, Desconto/Abat., Baixa de Saldo e Cancelamento. Refine a análise usando os filtros laterais.")
