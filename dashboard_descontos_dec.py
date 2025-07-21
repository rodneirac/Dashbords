import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image
import calendar
from streamlit_extras.metric_cards import style_metric_cards

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

# ==== CARDS MODERNOS COM STREAMLIT-EXTRAS ====
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown("### ⏳ Prorrogação")
    st.metric("Solicitações", qtd_prl_card, help="Total de solicitações de prorrogação")
    st.metric("Média Dias", f"{media_dias_prl:.1f}" if media_dias_prl else "-", help="Média de dias das prorrogações")
with col2:
    st.markdown("### 💰 Desconto/Abat.")
    st.metric("Solicitações", qtd_desc_card, help="Total de solicitações de desconto/abatimento")
    st.metric("Desconto Total", f"R$ {desconto_total_dec_abat:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), help="Soma dos descontos concedidos")
with col3:
    st.markdown("### 🏦 Baixa de Saldo")
    st.metric("Solicitações", qtd_bxs, help="Total de baixas de saldo")
    st.metric("Montante", f"R$ {desconto_total_bxs:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), help="Soma dos valores baixados")
with col4:
    st.markdown("### ❌ Cancelamentos")
    st.metric("Solicitações", qtd_cancel, help="Total de cancelamentos")
    st.metric("Montante", f"R$ {montante_cancel:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), help="Soma dos valores cancelados")

style_metric_cards(
    background_color="#F9F9F9",
    border_left_color="#800040",
    border_color="#DDD",
    border_radius_px=12,
    box_shadow=True
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

# ==== GRÁFICOS DE PIZZA POR NÍVEL 1 DESCRIÇÃO (COM TOOLTIP DETALHADO DE NÍVEL 2) ====
st.subheader("Distribuição por Nível 1 Descrição")

# ===== PRORROGAÇÃO (Qtde de solicitações) =====
df_tmp = df_prorrog.copy()
df_n2 = (
    df_tmp.groupby(["Nível 1 Descrição", "Nível 2 Descrição"]).size().reset_index(name="Qtde")
)
def tooltip_n2_prl(nivel1):
    fatia = df_n2[df_n2["Nível 1 Descrição"] == nivel1]
    if fatia.empty:
        return "-"
    return "<br>".join([f"{n2}: {qtde}" for n2, qtde in zip(fatia["Nível 2 Descrição"], fatia["Qtde"])])

pizza_prl = (
    df_prorrog.groupby("Nível 1 Descrição")
    .size()
    .reset_index(name="Qtde")
    .sort_values("Qtde", ascending=False)
)
pizza_prl["TooltipN2"] = pizza_prl["Nível 1 Descrição"].apply(tooltip_n2_prl)

col_pie1, col_pie2 = st.columns(2)
with col_pie1:
    fig_pie_prl = px.pie(
        pizza_prl,
        names="Nível 1 Descrição",
        values="Qtde",
        hole=0.4,
        title="Prorrogação",
        custom_data=["TooltipN2"]
    )
    fig_pie_prl.update_traces(
        textinfo='percent+label', pull=[0.08]*len(pizza_prl),
        hovertemplate="<b>%{label}</b><br>Qtd: %{value}<br><br><b>Qtde por Nível 2:</b><br>%{customdata[0]}<extra></extra>"
    )
    st.plotly_chart(fig_pie_prl, use_container_width=True)

# ===== DESCONTO/ABATIMENTO (soma de Desconto) =====
df_tmp = df_desc_abat.copy()
df_n2 = (
    df_tmp.groupby(["Nível 1 Descrição", "Nível 2 Descrição"])["Desconto"].sum().reset_index()
)
def tooltip_n2_desc(nivel1):
    fatia = df_n2[df_n2["Nível 1 Descrição"] == nivel1]
    if fatia.empty:
        return "-"
    return "<br>".join([f"{n2}: R$ {soma:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for n2, soma in zip(fatia["Nível 2 Descrição"], fatia["Desconto"])])

pizza_desc_abat = (
    df_desc_abat.groupby("Nível 1 Descrição")["Desconto"]
    .sum()
    .reset_index()
    .sort_values("Desconto", ascending=False)
)
pizza_desc_abat["TooltipN2"] = pizza_desc_abat["Nível 1 Descrição"].apply(tooltip_n2_desc)

with col_pie2:
    fig_pie_desc_abat = px.pie(
        pizza_desc_abat,
        names="Nível 1 Descrição",
        values="Desconto",
        hole=0.4,
        title="Desconto/Abat.",
        custom_data=["TooltipN2"]
    )
    fig_pie_desc_abat.update_traces(
        textinfo='percent+label', pull=[0.08]*len(pizza_desc_abat),
        hovertemplate="<b>%{label}</b><br>Soma: %{value:,.2f}<br><br><b>Nível 2:</b><br>%{customdata[0]}<extra></extra>"
    )
    st.plotly_chart(fig_pie_desc_abat, use_container_width=True)

# ===== BAIXA DE SALDO (soma de Desconto) =====
col_pie3, col_pie4 = st.columns(2)
df_tmp = df_baixa.copy()
df_n2 = (
    df_tmp.groupby(["Nível 1 Descrição", "Nível 2 Descrição"])["Desconto"].sum().reset_index()
)
def tooltip_n2_baixa(nivel1):
    fatia = df_n2[df_n2["Nível 1 Descrição"] == nivel1]
    if fatia.empty:
        return "-"
    return "<br>".join([f"{n2}: R$ {soma:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for n2, soma in zip(fatia["Nível 2 Descrição"], fatia["Desconto"])])

pizza_baixa = (
    df_baixa.groupby("Nível 1 Descrição")["Desconto"]
    .sum()
    .reset_index()
    .sort_values("Desconto", ascending=False)
)
pizza_baixa["TooltipN2"] = pizza_baixa["Nível 1 Descrição"].apply(tooltip_n2_baixa)

with col_pie3:
    fig_pie_baixa = px.pie(
        pizza_baixa,
        names="Nível 1 Descrição",
        values="Desconto",
        hole=0.4,
        title="Baixa de Saldo",
        custom_data=["TooltipN2"]
    )
    fig_pie_baixa.update_traces(
        textinfo='percent+label', pull=[0.08]*len(pizza_baixa),
        hovertemplate="<b>%{label}</b><br>Soma: %{value:,.2f}<br><br><b>Nível 2:</b><br>%{customdata[0]}<extra></extra>"
    )
    st.plotly_chart(fig_pie_baixa, use_container_width=True)

# ===== CANCELAMENTO (soma de Montante) =====
df_tmp = df_cancel.copy()
df_n2 = (
    df_tmp.groupby(["Nível 1 Descrição", "Nível 2 Descrição"])["Montante"].sum().reset_index()
)
def tooltip_n2_cancel(nivel1):
    fatia = df_n2[df_n2["Nível 1 Descrição"] == nivel1]
    if fatia.empty:
        return "-"
    return "<br>".join([f"{n2}: R$ {soma:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for n2, soma in zip(fatia["Nível 2 Descrição"], fatia["Montante"])])

pizza_cancel = (
    df_cancel.groupby("Nível 1 Descrição")["Montante"]
    .sum()
    .reset_index()
    .sort_values("Montante", ascending=False)
)
pizza_cancel["TooltipN2"] = pizza_cancel["Nível 1 Descrição"].apply(tooltip_n2_cancel)

with col_pie4:
    fig_pie_cancel = px.pie(
        pizza_cancel,
        names="Nível 1 Descrição",
        values="Montante",
        hole=0.4,
        title="Cancelamento",
        custom_data=["TooltipN2"]
    )
    fig_pie_cancel.update_traces(
        textinfo='percent+label', pull=[0.08]*len(pizza_cancel),
        hovertemplate="<b>%{label}</b><br>Soma: %{value:,.2f}<br><br><b>Nível 2:</b><br>%{customdata[0]}<extra></extra>"
    )
    st.plotly_chart(fig_pie_cancel, use_container_width=True)

# Helper para formatar coluna como reais
def format_reais(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

st.markdown("---")
st.markdown("Relatório dinâmico por instrução: Prorrogação, Desconto/Abat., Baixa de Saldo e Cancelamento. Refine a análise usando os filtros laterais.")

st.markdown("---")
st.subheader("Top 15 Filiais - Indicadores por Volume e Média")

col_top1, col_top2 = st.columns(2)

# 1. Top 15 - Média de Dias em Prorrogação
with col_top1:
    st.markdown("##### Média de Dias de Prorrogação por Filial (Top 15)")
    df_media_prl = (
        df_prorrog.groupby("Divisão")["Dias"]
        .mean()
        .reset_index()
        .sort_values("Dias", ascending=False)
        .head(15)
    )
    fig_media_prl = px.bar(
        df_media_prl, x="Dias", y="Divisão", orientation="h",
        labels={"Dias": "Média Dias", "Divisão": "Filial"},
        text="Dias"
    )
    fig_media_prl.update_layout(
        yaxis={'categoryorder':'total ascending'},
        height=500,
        margin=dict(l=50, r=30, t=30, b=30),
        showlegend=False
    )
    fig_media_prl.update_traces(
        marker_color="#3a75c4",
        texttemplate='%{text:.1f}', textposition="outside"
    )
    st.plotly_chart(fig_media_prl, use_container_width=True)

# 2. Top 15 - Desconto/Abatimento
with col_top2:
    st.markdown("##### Volume de Descontos/Abat. por Filial (Top 15)")
    df_top_desc = (
        df_desc_abat.groupby("Divisão")["Desconto"]
        .sum()
        .reset_index()
        .sort_values("Desconto", ascending=False)
        .head(15)
    )
    fig_top_desc = px.bar(
        df_top_desc, x="Desconto", y="Divisão", orientation="h",
        labels={"Desconto": "Volume R$", "Divisão": "Filial"},
        text="Desconto"
    )
    fig_top_desc.update_layout(
        yaxis={'categoryorder':'total ascending'},
        height=500,
        margin=dict(l=50, r=30, t=30, b=30),
        showlegend=False
    )
    fig_top_desc.update_traces(
        marker_color="#0b8d69",
        texttemplate='R$ %{text:,.2f}', textposition="outside"
    )
    st.plotly_chart(fig_top_desc, use_container_width=True)

col_top3, col_top4 = st.columns(2)

# 3. Top 15 - Baixa de Saldo
with col_top3:
    st.markdown("##### Volume de Baixa de Saldo por Filial (Top 15)")
    df_top_baixa = (
        df_baixa.groupby("Divisão")["Desconto"]
        .sum()
        .reset_index()
        .sort_values("Desconto", ascending=False)
        .head(15)
    )
    fig_top_baixa = px.bar(
        df_top_baixa, x="Desconto", y="Divisão", orientation="h",
        labels={"Desconto": "Volume R$", "Divisão": "Filial"},
        text="Desconto"
    )
    fig_top_baixa.update_layout(
        yaxis={'categoryorder':'total ascending'},
        height=500,
        margin=dict(l=50, r=30, t=30, b=30),
        showlegend=False
    )
    fig_top_baixa.update_traces(
        marker_color="#da853b",
        texttemplate='R$ %{text:,.2f}', textposition="outside"
    )
    st.plotly_chart(fig_top_baixa, use_container_width=True)

# 4. Top 15 - Cancelamentos
with col_top4:
    st.markdown("##### Volume de Cancelamentos por Filial (Top 15)")
    df_top_cancel = (
        df_cancel.groupby("Divisão")["Montante"]
        .sum()
        .reset_index()
        .sort_values("Montante", ascending=False)
        .head(15)
    )
    fig_top_cancel = px.bar(
        df_top_cancel, x="Montante", y="Divisão", orientation="h",
        labels={"Montante": "Volume R$", "Divisão": "Filial"},
        text="Montante"
    )
    fig_top_cancel.update_layout(
        yaxis={'categoryorder':'total ascending'},
        height=500,
        margin=dict(l=50, r=30, t=30, b=30),
        showlegend=False
    )
    fig_top_cancel.update_traces(
        marker_color="#a43e3e",
        texttemplate='R$ %{text:,.2f}', textposition="outside"
    )
    st.plotly_chart(fig_top_cancel, use_container_width=True)
