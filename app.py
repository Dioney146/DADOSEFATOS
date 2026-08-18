"""
DADOS E FATOS — Indicadores de Roteirização por Estado
Delly's Food Service

Lê os relatórios de rotas do RoadNet (xlsx/xls/csv) e monta uma apresentação
com os indicadores diários: veículos, ocupação, drop size, paradas x entregas
e peso x capacidade.

Fontes de dados aceitas:
  1) Arquivos na pasta ./dados do repositório (ex.: dados/AM.xlsx, dados/MG.xlsx)
  2) Upload pela barra lateral (o estado é detectado pelo nome do arquivo)
"""

from __future__ import annotations

import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Dados e Fatos | Roteirização",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

PASTA_DADOS = Path(__file__).parent / "dados"

# Estados atendidos. A chave é o código usado no nome do arquivo.
ESTADOS = {
    "AM": "Amazonas",
    "BA": "Bahia",
    "DF": "Distrito Federal",
    "ES": "Espírito Santo",
    "MG": "Minas Gerais",
    "MT": "Mato Grosso",
    "SP": "São Paulo",
    "SPW": "São Paulo (W Food)",
    "SP3P": "São Paulo (3P)",
}

# Paleta
COR_TEXTO = "#14161A"
COR_SUAVE = "#8A9299"
COR_AZUL = "#5D87B0"
COR_AZUL_CLARO = "#8FB3D4"
COR_ESCURA = "#2C3440"
COR_TRILHO = "#E6E8EA"
COR_BORDA = "#DCDFE3"

COLUNAS_ESPERADAS = [
    "ID",
    "Descrição",
    "Número de paradas",
    "Número de Ordens",
    "Entrega Total Peso",
    "Entrega Total Valor",
    "Capacidade Peso",
    "Equipamento",
    "Distância total",
    "Tipos de equipamento",
    "Sessão de roteirização",
    "Estado",
]

# ──────────────────────────────────────────────────────────────────────────────
# ESTILO
# ──────────────────────────────────────────────────────────────────────────────

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;600;700&family=Barlow+Condensed:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

.stApp { background: #F4F4F2; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 1.6rem; padding-bottom: 2.5rem; max-width: 1500px; }

html, body, [class*="css"] { color: #14161A; font-family: 'Archivo', sans-serif; }

/* Cabeçalho */
.cab {
    display: flex; align-items: flex-end; justify-content: space-between;
    gap: 24px; border-bottom: 2px solid #14161A; padding-bottom: 10px; margin-bottom: 22px;
}
.cab-titulo {
    font-family: 'Barlow Condensed', sans-serif; font-weight: 700; letter-spacing: .02em;
    font-size: 44px; line-height: .95; text-transform: uppercase; margin: 0;
}
.cab-sub {
    font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: .18em;
    text-transform: uppercase; color: #8A9299; margin-top: 6px;
}
.cab-meta { display: flex; gap: 30px; }
.meta-item { text-align: right; }
.meta-rot {
    font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: .18em;
    text-transform: uppercase; color: #8A9299;
}
.meta-val {
    font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 26px; line-height: 1.1;
}

/* Painéis (st.container(border=True)) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > div > div > .painel-topo),
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: #FFFFFF; border-color: #DCDFE3 !important; border-radius: 3px;
}
section[data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] { background: transparent; }
.painel-topo {
    display: flex; align-items: baseline; justify-content: space-between;
    border-bottom: 1px solid #ECEEF0; padding-bottom: 8px; margin-bottom: 6px;
}
.painel-titulo {
    font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 20px;
    text-transform: uppercase; letter-spacing: .03em;
}
.painel-nota {
    font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: .14em;
    text-transform: uppercase; color: #8A9299;
}
.painel-vazio {
    font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: #8A9299; padding: 30px 0;
}

/* Faixa de números acima dos gráficos */
.faixa { display: flex; gap: 2px; margin: 2px 0 0 0; }
.faixa-cel { flex: 1; text-align: center; }
.faixa-n1 {
    font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 15px; color: #5D87B0;
}
.faixa-n2 {
    font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 15px; color: #2C3440;
}

/* Rodapé de fonte */
.fonte {
    font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: .14em;
    text-transform: uppercase; color: #8A9299; margin-top: 26px;
    border-top: 1px solid #DCDFE3; padding-top: 10px;
}

section[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid #DCDFE3; }
section[data-testid="stSidebar"] h2 {
    font-family: 'Barlow Condensed', sans-serif; text-transform: uppercase; letter-spacing: .04em;
}
div[data-testid="stMetricValue"] { font-family: 'Barlow Condensed', sans-serif; }
</style>
"""


# ──────────────────────────────────────────────────────────────────────────────
# LEITURA E TRATAMENTO
# ──────────────────────────────────────────────────────────────────────────────

def normalizar(texto: str) -> str:
    """Minúsculas, sem acento e sem espaços extras."""
    if texto is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(texto))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", txt).strip().lower()


def br_para_float(valor) -> float:
    """Converte números em formato pt-BR ('1.420,8677') ou já numéricos."""
    if valor is None:
        return float("nan")
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip()
    if not txt or txt in {"-", "--"}:
        return float("nan")
    txt = re.sub(r"[^\d,.\-]", "", txt)
    if not txt:
        return float("nan")
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return float("nan")


def ler_arquivo(nome: str, conteudo: bytes) -> pd.DataFrame:
    """Lê xlsx, xls, HTML disfarçado de xls, csv ou texto delimitado."""
    buf = io.BytesIO(conteudo)
    tentativas = []

    if nome.lower().endswith((".xlsx", ".xlsm")):
        tentativas.append(lambda: pd.read_excel(io.BytesIO(conteudo), dtype=str))
    if nome.lower().endswith(".xls"):
        tentativas.append(lambda: pd.read_excel(io.BytesIO(conteudo), dtype=str))
        tentativas.append(lambda: pd.read_html(io.BytesIO(conteudo))[0].astype(str))
    tentativas.append(lambda: pd.read_excel(io.BytesIO(conteudo), dtype=str))
    tentativas.append(lambda: pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python",
                                          dtype=str, encoding="utf-8"))
    tentativas.append(lambda: pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python",
                                          dtype=str, encoding="latin-1"))

    erro_final = None
    for tentativa in tentativas:
        try:
            df = tentativa()
            if df is not None and len(df.columns) > 1:
                return df
        except Exception as exc:  # noqa: BLE001
            erro_final = exc
    raise ValueError(f"Não foi possível ler o arquivo {nome}: {erro_final}")


def mapear_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Reconhece as colunas pelo nome, tolerante a acento/caixa/espaços."""
    mapa = {normalizar(c): c for c in df.columns}
    renomear = {}
    for alvo in COLUNAS_ESPERADAS:
        chave = normalizar(alvo)
        if chave in mapa:
            renomear[mapa[chave]] = alvo
    return df.rename(columns=renomear)


def extrair_data(sessao: str):
    """A data da operação vem no início da 'Sessão de roteirização'."""
    achado = re.search(r"(\d{2}/\d{2}/\d{4})", str(sessao))
    if achado:
        return pd.to_datetime(achado.group(1), format="%d/%m/%Y", errors="coerce")
    achado = re.search(r"(\d{2}/\d{2})", str(sessao))
    if achado:
        return pd.to_datetime(achado.group(1) + "/" + str(pd.Timestamp.today().year),
                              format="%d/%m/%Y", errors="coerce")
    return pd.NaT


def detectar_estado(nome_arquivo: str, df: pd.DataFrame) -> str:
    """Descobre o estado pelo nome do arquivo; se falhar, pelo prefixo da descrição."""
    base = normalizar(Path(nome_arquivo).stem).upper()
    base = re.sub(r"[^A-Z0-9]", " ", base)
    fichas = base.split()
    for codigo in sorted(ESTADOS, key=len, reverse=True):
        if codigo in fichas or base.startswith(codigo):
            return codigo
    if "Descrição" in df.columns:
        prefixos = (
            df["Descrição"].dropna().astype(str)
            .str.extract(r"^([A-Za-z]{2,3})\s*-", expand=False).dropna().str.upper()
        )
        if not prefixos.empty:
            mais_comum = prefixos.value_counts().index[0]
            if mais_comum in ESTADOS:
                return mais_comum
    return "N/D"


def tratar(df: pd.DataFrame, nome_arquivo: str) -> pd.DataFrame:
    """Deixa a base pronta para análise: colunas numéricas, data e estado."""
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df = mapear_colunas(df)

    faltando = [c for c in ["Sessão de roteirização", "Entrega Total Peso", "Capacidade Peso"]
                if c not in df.columns]
    if faltando:
        raise ValueError(f"{nome_arquivo}: colunas ausentes {faltando}")

    numericas = {
        "Número de paradas": "PARADAS",
        "Número de Ordens": "ENTREGAS",
        "Entrega Total Peso": "PESO",
        "Entrega Total Valor": "VALOR",
        "Capacidade Peso": "CAPACIDADE",
        "Distância total": "DISTANCIA",
    }
    for origem, destino in numericas.items():
        df[destino] = df[origem].map(br_para_float) if origem in df.columns else float("nan")

    df["DATA"] = df["Sessão de roteirização"].map(extrair_data)
    df["ROTA"] = df["ID"].astype(str) if "ID" in df.columns else ""
    df["VEICULO"] = df["Equipamento"].astype(str).str.strip() if "Equipamento" in df.columns else ""
    df["TIPO_VEICULO"] = (
        df["Tipos de equipamento"].astype(str).str.strip().replace({"": "NÃO INFORMADO", "nan": "NÃO INFORMADO"})
        if "Tipos de equipamento" in df.columns else "NÃO INFORMADO"
    )
    df["STATUS"] = (
        df["Estado"].astype(str).str.strip().replace({"": "NÃO INFORMADO", "nan": "NÃO INFORMADO"})
        if "Estado" in df.columns else "NÃO INFORMADO"
    )
    df["UF"] = detectar_estado(nome_arquivo, df)
    df["ARQUIVO"] = Path(nome_arquivo).name

    df = df.dropna(subset=["DATA"])
    df = df[df["CAPACIDADE"].fillna(0) > 0]
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def carregar(arquivos: tuple[tuple[str, bytes], ...]) -> pd.DataFrame:
    """Lê e trata todos os arquivos (cacheado pelo conteúdo)."""
    bases = []
    problemas = []
    for nome, conteudo in arquivos:
        try:
            bases.append(tratar(ler_arquivo(nome, conteudo), nome))
        except Exception as exc:  # noqa: BLE001
            problemas.append(f"{nome}: {exc}")
    if problemas:
        st.session_state["_problemas"] = problemas
    if not bases:
        return pd.DataFrame()
    return pd.concat(bases, ignore_index=True)


def arquivos_da_pasta() -> list[tuple[str, bytes]]:
    if not PASTA_DADOS.exists():
        return []
    achados = []
    for caminho in sorted(PASTA_DADOS.iterdir()):
        if caminho.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".csv"} and not caminho.name.startswith("~$"):
            achados.append((caminho.name, caminho.read_bytes()))
    return achados


# ──────────────────────────────────────────────────────────────────────────────
# INDICADORES
# ──────────────────────────────────────────────────────────────────────────────

def indicadores_por_dia(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por data com todos os indicadores da apresentação."""
    if df.empty:
        return pd.DataFrame()

    agrupado = df.groupby("DATA").agg(
        ROTAS=("ROTA", "count"),
        VEICULOS=("VEICULO", pd.Series.nunique),
        PARADAS=("PARADAS", "sum"),
        ENTREGAS=("ENTREGAS", "sum"),
        PESO=("PESO", "sum"),
        CAPACIDADE=("CAPACIDADE", "sum"),
        VALOR=("VALOR", "sum"),
        DISTANCIA=("DISTANCIA", "sum"),
    ).reset_index()

    agrupado["OCUPACAO"] = agrupado["PESO"] / agrupado["CAPACIDADE"].replace(0, pd.NA)
    agrupado["MEDIA_PARADAS"] = agrupado["PARADAS"] / agrupado["ROTAS"].replace(0, pd.NA)
    agrupado["DROP_PARADA"] = agrupado["PESO"] / agrupado["PARADAS"].replace(0, pd.NA)
    agrupado["DROP_VEICULO"] = agrupado["PESO"] / agrupado["VEICULOS"].replace(0, pd.NA)
    agrupado["DROP_ROTA"] = agrupado["PESO"] / agrupado["ROTAS"].replace(0, pd.NA)
    agrupado["ROTULO"] = agrupado["DATA"].dt.strftime("%d/%m")
    return agrupado.sort_values("DATA").reset_index(drop=True)


def num(valor, casas: int = 0) -> str:
    """Formata número no padrão pt-BR."""
    if valor is None or pd.isna(valor):
        return "—"
    texto = f"{float(valor):,.{casas}f}"
    return texto.replace(",", "@").replace(".", ",").replace("@", ".")


# ──────────────────────────────────────────────────────────────────────────────
# COMPONENTES VISUAIS
# ──────────────────────────────────────────────────────────────────────────────

LAYOUT_BASE = dict(
    margin=dict(l=8, r=8, t=6, b=26),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Archivo, sans-serif", size=11, color=COR_TEXTO),
    showlegend=False,
    hoverlabel=dict(font_size=12, font_family="Archivo, sans-serif"),
)

EIXO_X = dict(showgrid=False, zeroline=False, linecolor=COR_BORDA,
              tickfont=dict(size=10, color=COR_SUAVE))
EIXO_Y = dict(showgrid=True, gridcolor="#F0F1F3", zeroline=False, showline=False,
              tickfont=dict(size=10, color=COR_SUAVE))


def titulo_painel(titulo: str, nota: str = "") -> None:
    """Cabeçalho de um painel (usado dentro de um st.container com borda)."""
    st.markdown(
        f'''<div class="painel-topo">
            <span class="painel-titulo">{titulo}</span>
            <span class="painel-nota">{nota}</span></div>''',
        unsafe_allow_html=True,
    )


LIMITE_FAIXA = 16  # acima disso os números por dia ficam ilegíveis e são omitidos


def faixa_numeros(valores: list[str], cor: str = "azul") -> None:
    """Linha de números alinhada com as colunas do gráfico abaixo."""
    if len(valores) > LIMITE_FAIXA:
        return
    classe = "faixa-n1" if cor == "azul" else "faixa-n2"
    celulas = "".join(f'<div class="faixa-cel {classe}">{v}</div>' for v in valores)
    st.markdown(f'<div class="faixa">{celulas}</div>', unsafe_allow_html=True)


def grafico_barras(dados: pd.DataFrame, coluna: str, altura: int = 190) -> go.Figure:
    fig = go.Figure(
        go.Bar(
            x=dados["ROTULO"], y=dados[coluna],
            marker_color=COR_AZUL, marker_line_width=0, width=0.62,
            hovertemplate="%{x}<br>%{y}<extra></extra>",
        )
    )
    fig.update_layout(**LAYOUT_BASE, height=altura, bargap=0.35)
    fig.update_xaxes(**EIXO_X)
    fig.update_yaxes(**EIXO_Y, visible=False)
    return fig


def grafico_barras_horizontais(dados: pd.DataFrame, coluna: str, altura: int = 260) -> go.Figure:
    dados = dados.iloc[::-1]
    maximo = float(dados[coluna].max() or 1)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=dados["ROTULO"], x=[maximo] * len(dados), orientation="h",
        marker_color=COR_TRILHO, marker_line_width=0, hoverinfo="skip", width=0.55,
    ))
    fig.add_trace(go.Bar(
        y=dados["ROTULO"], x=dados[coluna], orientation="h",
        marker_color=COR_AZUL, marker_line_width=0, width=0.55,
        text=[num(v, 1) for v in dados[coluna]], textposition="outside",
        textfont=dict(family="Barlow Condensed, sans-serif", size=14, color=COR_TEXTO),
        hovertemplate="%{y}<br>%{x:.1f} kg<extra></extra>",
    ))
    fig.update_layout(**LAYOUT_BASE, height=altura, barmode="overlay", bargap=0.25)
    fig.update_xaxes(visible=False, range=[0, maximo * 1.18])
    fig.update_yaxes(showgrid=False, zeroline=False, linecolor="rgba(0,0,0,0)",
                     tickfont=dict(size=10, color=COR_SUAVE))
    return fig


def grafico_duas_linhas(dados: pd.DataFrame, col_a: str, col_b: str,
                        nome_a: str, nome_b: str, altura: int = 230) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dados["ROTULO"], y=dados[col_a], name=nome_a, mode="lines+markers",
        line=dict(color=COR_ESCURA, width=2), marker=dict(size=6, symbol="square"),
        yaxis="y", hovertemplate=f"{nome_a}: %{{y:.1f}}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=dados["ROTULO"], y=dados[col_b], name=nome_b, mode="lines+markers",
        line=dict(color=COR_AZUL, width=2), marker=dict(size=6, symbol="square"),
        yaxis="y2", hovertemplate=f"{nome_b}: %{{y:.0f}}<extra></extra>",
    ))
    base = {k: v for k, v in LAYOUT_BASE.items() if k != "showlegend"}
    fig.update_layout(
        **base, height=altura, showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=10, color=COR_SUAVE)),
        yaxis=dict(**EIXO_Y, visible=False),
        yaxis2=dict(overlaying="y", side="right", visible=False),
    )
    fig.update_xaxes(**EIXO_X)
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# PÁGINA
# ──────────────────────────────────────────────────────────────────────────────

def barra_lateral() -> tuple[pd.DataFrame, str]:
    st.sidebar.markdown("## Base de dados")
    enviados = st.sidebar.file_uploader(
        "Relatório de rotas (RoadNet)",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
        help="O estado é identificado pelo nome do arquivo. Ex.: AM.xlsx, MG_1.xlsx.",
    )

    arquivos = arquivos_da_pasta()
    if enviados:
        arquivos = arquivos + [(f.name, f.getvalue()) for f in enviados]

    if not arquivos:
        return pd.DataFrame(), ""

    st.session_state.pop("_problemas", None)
    df = carregar(tuple(arquivos))
    for aviso in st.session_state.get("_problemas", []):
        st.sidebar.warning(aviso)
    if df.empty:
        return df, ""

    st.sidebar.markdown("## Filtros")
    ufs = sorted(df["UF"].unique())
    uf = st.sidebar.selectbox(
        "Estado", ufs, format_func=lambda c: f"{c} — {ESTADOS.get(c, 'Não identificado')}"
    )
    df = df[df["UF"] == uf]

    status_disponiveis = sorted(df["STATUS"].unique())
    padrao = [s for s in status_disponiveis if normalizar(s).startswith("conclu")] or status_disponiveis
    status = st.sidebar.multiselect("Status da rota", status_disponiveis, default=padrao)
    if status:
        df = df[df["STATUS"].isin(status)]

    if not df.empty:
        d_min, d_max = df["DATA"].min().date(), df["DATA"].max().date()
        periodo = st.sidebar.date_input("Período", value=(d_min, d_max),
                                        min_value=d_min, max_value=d_max, format="DD/MM/YYYY")
        if isinstance(periodo, tuple) and len(periodo) == 2:
            df = df[(df["DATA"].dt.date >= periodo[0]) & (df["DATA"].dt.date <= periodo[1])]

    tipos = sorted(df["TIPO_VEICULO"].unique())
    escolhidos = st.sidebar.multiselect("Tipo de veículo", tipos, default=tipos)
    if escolhidos:
        df = df[df["TIPO_VEICULO"].isin(escolhidos)]

    st.sidebar.markdown("## Drop size")
    base_drop = st.sidebar.radio(
        "Calcular o drop por", ["Parada", "Veículo", "Rota"], index=0,
        help="Drop = peso entregue dividido pela base escolhida.",
    )

    return df, base_drop


def cabecalho(uf: str, resumo: pd.DataFrame) -> None:
    if resumo.empty:
        periodo, dias, rotas = "—", "0", "0"
    else:
        periodo = f"{resumo['ROTULO'].iloc[0]}–{resumo['ROTULO'].iloc[-1]}"
        dias = num(len(resumo))
        rotas = num(resumo["ROTAS"].sum())

    st.markdown(
        f"""
        <div class="cab">
          <div>
            <p class="cab-titulo">Indicadores por dia</p>
            <div class="cab-sub">{uf} · {ESTADOS.get(uf, 'Estado não identificado')} · Delly's Food Service</div>
          </div>
          <div class="cab-meta">
            <div class="meta-item"><div class="meta-rot">Período</div><div class="meta-val">{periodo}</div></div>
            <div class="meta-item"><div class="meta-rot">Dias</div><div class="meta-val">{dias}</div></div>
            <div class="meta-item"><div class="meta-rot">Rotas</div><div class="meta-val">{rotas}</div></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def linha_um(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str) -> None:
    c1, c2, c3 = st.columns([1.05, 1.05, 0.9], gap="medium")

    with c1, st.container(border=True):
        titulo_painel("Veículos", "rotas / dia")
        faixa_numeros([num(v) for v in resumo["ROTAS"]], cor="escuro")
        st.plotly_chart(grafico_barras(resumo, "VEICULOS"), width="stretch",
                        config={"displayModeBar": False}, key="g_veiculos")

    with c2, st.container(border=True):
        titulo_painel("Ocupação", "peso ÷ capacidade")
        faixa_numeros([f"{num(v * 100)}%" if pd.notna(v) else "—" for v in resumo["OCUPACAO"]],
                      cor="escuro")
        st.plotly_chart(grafico_barras(resumo, "OCUPACAO"), width="stretch",
                        config={"displayModeBar": False}, key="g_ocupacao")

    with c3, st.container(border=True):
        titulo_painel("Drop", rotulo_drop)
        altura_drop = max(232, 22 * len(resumo))
        st.plotly_chart(grafico_barras_horizontais(resumo, coluna_drop, altura=altura_drop),
                        width="stretch", config={"displayModeBar": False}, key="g_drop")


def linha_dois(resumo: pd.DataFrame) -> None:
    c1, c2 = st.columns(2, gap="medium")

    with c1, st.container(border=True):
        titulo_painel("Paradas × entregas", "entregas / média de paradas")
        faixa_numeros([num(v) for v in resumo["ENTREGAS"]], cor="azul")
        faixa_numeros([num(v, 1) for v in resumo["MEDIA_PARADAS"]], cor="escuro")
        st.plotly_chart(
            grafico_duas_linhas(resumo, "MEDIA_PARADAS", "ENTREGAS", "Média paradas", "Entregas"),
            width="stretch", config={"displayModeBar": False}, key="g_paradas",
        )

    with c2, st.container(border=True):
        titulo_painel("Peso × capacidade por dia", "capacidade / peso (kg)")
        faixa_numeros([num(v) for v in resumo["CAPACIDADE"]], cor="azul")
        faixa_numeros([num(v) for v in resumo["PESO"]], cor="escuro")
        st.plotly_chart(
            grafico_duas_linhas(resumo, "PESO", "CAPACIDADE", "Peso (kg)", "Capacidade (kg)"),
            width="stretch", config={"displayModeBar": False}, key="g_peso",
        )


def tabela_detalhe(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str) -> None:
    tabela = pd.DataFrame({
        "Data": resumo["ROTULO"],
        "Rotas": resumo["ROTAS"],
        "Veículos": resumo["VEICULOS"],
        "Paradas": resumo["PARADAS"].round(0),
        "Entregas": resumo["ENTREGAS"].round(0),
        "Média paradas": resumo["MEDIA_PARADAS"].round(1),
        "Peso (kg)": resumo["PESO"].round(0),
        "Capacidade (kg)": resumo["CAPACIDADE"].round(0),
        "Ocupação": (resumo["OCUPACAO"] * 100).round(1),
        rotulo_drop: resumo[coluna_drop].round(1),
    })
    total = {
        "Data": "TOTAL",
        "Rotas": resumo["ROTAS"].sum(),
        "Veículos": resumo["VEICULOS"].max(),
        "Paradas": resumo["PARADAS"].sum().round(0),
        "Entregas": resumo["ENTREGAS"].sum().round(0),
        "Média paradas": round(resumo["PARADAS"].sum() / max(resumo["ROTAS"].sum(), 1), 1),
        "Peso (kg)": resumo["PESO"].sum().round(0),
        "Capacidade (kg)": resumo["CAPACIDADE"].sum().round(0),
        "Ocupação": round(resumo["PESO"].sum() / max(resumo["CAPACIDADE"].sum(), 1) * 100, 1),
        rotulo_drop: round(resumo[coluna_drop].mean(), 1),
    }
    tabela = pd.concat([tabela, pd.DataFrame([total])], ignore_index=True)

    with st.expander("Ver tabela e baixar os dados"):
        st.dataframe(tabela, width="stretch", hide_index=True)
        st.download_button(
            "Baixar indicadores em CSV",
            tabela.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name="indicadores_por_dia.csv",
            mime="text/csv",
        )


def main() -> None:
    st.markdown(CSS, unsafe_allow_html=True)

    df, base_drop = barra_lateral()

    if df.empty:
        st.markdown('<p class="cab-titulo">Dados e fatos</p>', unsafe_allow_html=True)
        st.markdown(
            '<div class="cab-sub">Indicadores de roteirização · Delly\'s Food Service</div>',
            unsafe_allow_html=True,
        )
        st.info(
            "Envie o relatório de rotas do RoadMet/RoadNet na barra lateral, ou coloque os "
            "arquivos na pasta **dados/** do repositório (ex.: `dados/AM.xlsx`) para o site "
            "abrir já preenchido."
        )
        return

    coluna_drop = {"Parada": "DROP_PARADA", "Veículo": "DROP_VEICULO", "Rota": "DROP_ROTA"}[base_drop]
    rotulo_drop = {"Parada": "kg / parada", "Veículo": "kg / veículo", "Rota": "kg / rota"}[base_drop]

    resumo = indicadores_por_dia(df)
    if resumo.empty:
        st.warning("Nenhuma rota no filtro selecionado.")
        return

    cabecalho(df["UF"].iloc[0], resumo)
    linha_um(resumo, coluna_drop, rotulo_drop)
    st.write("")
    linha_dois(resumo)
    tabela_detalhe(resumo, coluna_drop, rotulo_drop)

    st.markdown(
        f'<div class="fonte">Fonte: relatório de rotas RoadNet · '
        f'{num(len(df))} rotas processadas · arquivos: {", ".join(sorted(df["ARQUIVO"].unique()))}</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
