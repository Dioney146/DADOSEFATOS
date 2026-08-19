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

try:  # arquivo opcional: sem ele, o site roda só no modo apresentação
    from painel_arrastavel import renderizar_painel
except ImportError:  # noqa: BLE001
    renderizar_painel = None

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
COR_TRILHO = "#E2E4E6"
COR_FUNDO_GRAFICO = "#F1F1EF"
COR_GRADE = "#C9CCCF"
COR_BORDA = "#DCDFE3"

TAMANHO_PADRAO = {"altura": 232, "largura": "Padrão", "modo_painel": False}

LARGURAS_PAGINA = {
    "Estreita": "1180px",
    "Padrão": "1560px",
    "Larga": "1900px",
    "Tela cheia": "100%",
}

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
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700&family=Barlow+Condensed:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

/* ── Ícones do Streamlit ─────────────────────────────────────────────────────
   Se a rede bloquear o Google Fonts, a fonte de ícones não carrega e o
   navegador imprime o NOME do ícone ("double_arrow_right"). Aqui o texto é
   zerado sempre e o desenho é reposto com caracteres comuns, sem depender
   de nenhuma fonte externa.                                                  */
span[data-testid="stIconMaterial"], .material-symbols-rounded, .material-icons {
    font-size: 0 !important; line-height: 0 !important;
    width: 20px !important; height: 20px !important;
    overflow: hidden !important; white-space: nowrap;
    display: inline-flex !important; align-items: center; justify-content: center;
    color: transparent !important;
}
span[data-testid="stIconMaterial"]::after,
.material-symbols-rounded::after, .material-icons::after {
    content: ""; font-family: 'Archivo', Arial, sans-serif;
    font-size: 15px; line-height: 1; color: #14161A;
}
/* Botão de abrir/fechar a barra lateral */
header[data-testid="stHeader"] span[data-testid="stIconMaterial"]::after,
button[data-testid="stExpandSidebarButton"] span[data-testid="stIconMaterial"]::after,
button[data-testid="stBaseButton-headerNoPadding"] span[data-testid="stIconMaterial"]::after,
div[data-testid="stSidebarCollapseButton"] span[data-testid="stIconMaterial"]::after {
    content: "»"; font-size: 17px; font-weight: 700;
}
/* Setas de blocos recolhíveis */
div[data-testid="stExpander"] span[data-testid="stIconMaterial"]::after {
    content: "▾"; font-size: 13px;
}

/* ── Base: o site é claro, independente do tema do navegador/Streamlit ────── */
.stApp, [data-testid="stAppViewContainer"] { background: #F4F4F2 !important; }
header[data-testid="stHeader"] {
    background: #F4F4F2 !important; height: 3rem; z-index: 999;
}
.block-container { padding-top: 3.4rem; padding-bottom: 2.5rem; max-width: 1560px; }
html, body, .stApp, .stApp p, .stApp span, .stApp label, .stApp li,
.stApp h1, .stApp h2, .stApp h3, .stApp h4 {
    color: #14161A;
    font-family: 'Archivo', Arial, sans-serif;
}

/* ── Cabeçalho ───────────────────────────────────────────────────────────── */
.cab {
    display: flex; align-items: flex-end; justify-content: space-between;
    gap: 24px; border-bottom: 2px solid #14161A; padding-bottom: 10px; margin-bottom: 18px;
}
.cab-titulo {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; letter-spacing: .02em;
    font-size: 42px; line-height: .95; text-transform: uppercase; margin: 0; color: #14161A;
}
.cab-sub {
    font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 11px; letter-spacing: .16em;
    text-transform: uppercase; color: #7C858D; margin-top: 6px;
}
.cab-meta { display: flex; gap: 28px; }
.meta-item { text-align: right; }
.meta-rot {
    font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 10px; letter-spacing: .16em;
    text-transform: uppercase; color: #7C858D;
}
.meta-val {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 26px;
    line-height: 1.1; color: #14161A;
}

/* ── Cartões de resumo ───────────────────────────────────────────────────── */
.kpis { display: flex; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
.kpi {
    flex: 1; min-width: 150px; background: #FFFFFF; border: 1px solid #DCDFE3;
    border-top: 3px solid #5D87B0; border-radius: 3px; padding: 10px 14px 12px 14px;
}
.kpi-rot {
    font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 10px; letter-spacing: .14em;
    text-transform: uppercase; color: #7C858D;
}
.kpi-val {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 30px;
    line-height: 1.1; color: #14161A;
}
.kpi-nota { font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 10px; color: #7C858D; }

/* ── Painéis (st.container com borda) ────────────────────────────────────── */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: #FFFFFF; border: 1px solid #14161A !important; border-radius: 0 !important;
    padding: 2px 4px;
}
/* Painéis lado a lado terminam na mesma altura */
div[data-testid="stHorizontalBlock"] { align-items: stretch; }
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { display: flex; }
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]
    div[data-testid="stVerticalBlockBorderWrapper"] { width: 100%; height: 100%; }
section[data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] {
    background: transparent; border-color: transparent !important;
}
.painel-topo {
    display: flex; align-items: baseline; justify-content: space-between;
    padding: 2px 2px 10px 2px; margin-bottom: 2px;
}
.painel-titulo {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 23px;
    text-transform: uppercase; letter-spacing: .01em; color: #14161A;
}
.painel-nota {
    font-family: 'Archivo', Arial, sans-serif; font-size: 11px; font-weight: 500;
    letter-spacing: .16em; text-transform: uppercase; color: #7C858D;
}

/* Paginação dos slides */
.paginacao {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 20px;
    text-align: center; color: #14161A; padding-top: 4px;
}

/* ── Faixa de números acima dos gráficos ─────────────────────────────────── */
.faixa { display: flex; gap: 2px; margin: 0 0 2px 0; }
.faixa-cel { flex: 1; text-align: center; }
.faixa-n1 {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 17px; color: #5D87B0;
}
.faixa-n2 {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 17px; color: #14161A;
}

/* ── Mini estatísticas (painel de drop) ──────────────────────────────────── */
.mini { display: flex; gap: 12px; margin: 6px 0 4px 0; }
.mini-item { flex: 1; }
.mini-rot {
    font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 9px; letter-spacing: .12em;
    text-transform: uppercase; color: #7C858D;
}
.mini-val {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 19px;
    line-height: 1.1; color: #14161A;
}
.mini-dia { font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 10px; color: #7C858D; }

/* ── Rodapé ──────────────────────────────────────────────────────────────── */
.fonte {
    font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 10px; letter-spacing: .12em;
    text-transform: uppercase; color: #7C858D; margin-top: 22px;
    border-top: 1px solid #DCDFE3; padding-top: 10px;
}

/* ── Barra lateral: clara e legível em qualquer tema ─────────────────────── */
section[data-testid="stSidebar"], section[data-testid="stSidebar"] > div {
    background: #FFFFFF !important; border-right: 1px solid #DCDFE3;
}
section[data-testid="stSidebar"] * { color: #14161A; }
section[data-testid="stSidebar"] h2 {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; text-transform: uppercase;
    letter-spacing: .04em; font-size: 17px; margin-bottom: .2rem;
}
section[data-testid="stSidebar"] label p { font-size: 12px; font-weight: 600; }

/* Campos de formulário sempre claros */
div[data-baseweb="select"] > div, div[data-baseweb="input"] > div,
div[data-testid="stFileUploaderDropzone"], div[data-baseweb="popover"] li {
    background: #FFFFFF !important; border-color: #DCDFE3 !important; color: #14161A !important;
}
div[data-baseweb="popover"] ul { background: #FFFFFF !important; }
/* Etiquetas dos multiselects */
span[data-baseweb="tag"] {
    background: #2C3440 !important; border-radius: 2px !important;
}
span[data-baseweb="tag"] span, span[data-baseweb="tag"] svg { color: #FFFFFF !important; fill: #FFFFFF !important; }

/* Impressão (Ctrl+P) — sai igual a um slide em PDF */
@media print {
    section[data-testid="stSidebar"], header[data-testid="stHeader"],
    div[data-testid="stExpander"], .stButton, div[data-testid="stSegmentedControl"] { display: none !important; }
    .block-container { padding: 0 !important; max-width: 100% !important; }
    div[data-testid="stVerticalBlockBorderWrapper"] { break-inside: avoid; }
}

/* Botão de alternância do drop */
div[data-testid="stSegmentedControl"] button { font-size: 12px; }
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
    plot_bgcolor=COR_FUNDO_GRAFICO,
    font=dict(family="Archivo, sans-serif", size=11, color=COR_TEXTO),
    showlegend=False,
    hoverlabel=dict(font_size=12, font_family="Archivo, sans-serif"),
)

EIXO_X = dict(showgrid=False, zeroline=False, linecolor=COR_TEXTO, linewidth=1,
              ticks="", tickfont=dict(size=10, color=COR_SUAVE))
EIXO_Y = dict(showgrid=True, gridcolor=COR_GRADE, griddash="dot", zeroline=False, showline=False,
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


def eixo_datas(fig: go.Figure, dados: pd.DataFrame) -> None:
    """Com muitos dias, mostra um rótulo sim, outro não, para não embolar."""
    passo = 1 if len(dados) <= 18 else 2 if len(dados) <= 40 else 3
    fig.update_xaxes(**EIXO_X, type="category", dtick=passo, tickangle=-90 if len(dados) > 12 else 0)


def grafico_barras(dados: pd.DataFrame, coluna: str, altura: int = 232) -> go.Figure:
    fig = go.Figure(
        go.Bar(
            x=dados["ROTULO"], y=dados[coluna],
            marker_color=COR_AZUL, marker_line_width=0, width=0.62,
            hovertemplate="%{x}<br>%{y}<extra></extra>",
        )
    )
    fig.update_layout(**LAYOUT_BASE, height=altura, bargap=0.35)
    eixo_datas(fig, dados)
    fig.update_yaxes(**EIXO_Y, showticklabels=False)
    return fig


def cores_pela_media(valores, media: float) -> list[str]:
    """Dias acima da média em azul cheio; abaixo, em azul claro."""
    return [COR_AZUL if (pd.notna(v) and v >= media) else COR_AZUL_CLARO for v in valores]


def grafico_drop_por_dia(dados: pd.DataFrame, coluna: str, altura: int = 232) -> go.Figure:
    """Drop em ordem cronológica, no mesmo eixo dos demais painéis do dia."""
    media = float(dados[coluna].mean())
    fig = go.Figure(
        go.Bar(
            x=dados["ROTULO"], y=dados[coluna], width=0.62,
            marker_color=cores_pela_media(dados[coluna], media), marker_line_width=0,
            hovertemplate="%{x}<br>%{y:.1f} kg<extra></extra>",
        )
    )
    fig.add_hline(
        y=media, line_width=1, line_dash="dot", line_color=COR_ESCURA,
        annotation_text=f"média {num(media, 1)}", annotation_position="top left",
        annotation_font=dict(family="IBM Plex Mono, monospace", size=9, color=COR_SUAVE),
    )
    fig.update_layout(**LAYOUT_BASE, height=altura, bargap=0.35)
    eixo_datas(fig, dados)
    fig.update_yaxes(**EIXO_Y, showticklabels=False)
    return fig


def grafico_drop_ranking(dados: pd.DataFrame, coluna: str, limite: int = 12,
                         altura: int = 232) -> go.Figure:
    """Dias ordenados do maior para o menor drop, com trilho de comparação."""
    media = float(dados[coluna].mean())
    ranking = dados.dropna(subset=[coluna]).sort_values(coluna, ascending=False).head(limite)
    ranking = ranking.iloc[::-1]
    maximo = float(ranking[coluna].max() or 1)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=ranking["ROTULO"], x=[maximo] * len(ranking), orientation="h",
        marker_color=COR_TRILHO, marker_line_width=0, hoverinfo="skip", width=0.55,
    ))
    fig.add_trace(go.Bar(
        y=ranking["ROTULO"], x=ranking[coluna], orientation="h", width=0.55,
        marker_color=cores_pela_media(ranking[coluna], media), marker_line_width=0,
        text=[num(v, 1) for v in ranking[coluna]], textposition="outside",
        textfont=dict(family="Barlow Condensed, sans-serif", size=13, color=COR_TEXTO),
        hovertemplate="%{y}<br>%{x:.1f} kg<extra></extra>",
    ))
    fig.add_vline(y0=0, y1=1, x=media, line_width=1, line_dash="dot", line_color=COR_ESCURA)
    base = {k: v for k, v in LAYOUT_BASE.items() if k != "plot_bgcolor"}
    fig.update_layout(**base, height=altura, barmode="overlay", bargap=0.25,
                      plot_bgcolor="rgba(0,0,0,0)")
    fig.update_xaxes(visible=False, range=[0, maximo * 1.18])
    fig.update_yaxes(showgrid=False, zeroline=False, linecolor="rgba(0,0,0,0)",
                     tickfont=dict(size=10, color=COR_SUAVE))
    return fig


def mini_estatisticas_drop(dados: pd.DataFrame, coluna: str) -> None:
    """Média do período e os dias de maior e menor drop."""
    validos = dados.dropna(subset=[coluna])
    if validos.empty:
        return
    melhor = validos.loc[validos[coluna].idxmax()]
    pior = validos.loc[validos[coluna].idxmin()]
    blocos = [
        ("Média", num(validos[coluna].mean(), 1), "período"),
        ("Maior", num(melhor[coluna], 1), melhor["ROTULO"]),
        ("Menor", num(pior[coluna], 1), pior["ROTULO"]),
    ]
    html = "".join(
        f'<div class="mini-item"><div class="mini-rot">{rot}</div>'
        f'<div class="mini-val">{val}</div><div class="mini-dia">{dia}</div></div>'
        for rot, val, dia in blocos
    )
    st.markdown(f'<div class="mini">{html}</div>', unsafe_allow_html=True)


def grafico_duas_linhas(dados: pd.DataFrame, col_a: str, col_b: str,
                        nome_a: str, nome_b: str, altura: int = 250) -> go.Figure:
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
        yaxis=dict(**EIXO_Y, showticklabels=False),
        yaxis2=dict(overlaying="y", side="right", visible=False),
    )
    eixo_datas(fig, dados)
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
        return pd.DataFrame(), "Parada", "10", TAMANHO_PADRAO

    st.session_state.pop("_problemas", None)
    df = carregar(tuple(arquivos))
    for aviso in st.session_state.get("_problemas", []):
        st.sidebar.warning(aviso)
    if df.empty:
        return df, "Parada", "10", TAMANHO_PADRAO

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

    st.sidebar.markdown("## Apresentação")
    dias_slide = st.sidebar.selectbox(
        "Dias por slide", ["10", "15", "20", "Todos"], index=0,
        help="Como no slide impresso: cada tela mostra um bloco de dias.",
    )

    altura = st.sidebar.slider(
        "Altura dos gráficos (px)", min_value=150, max_value=520, value=232, step=10,
        help="Vale para os cinco painéis. Aumente para projetar, reduza para caber na tela.",
    )
    modo_painel = renderizar_painel is not None and st.sidebar.toggle(
        "Modo painel (arrastar e redimensionar)", value=False,
        help="Monta os cinco gráficos numa grade livre, como no canvas do Power BI. "
             "O layout fica salvo neste navegador.",
    )
    largura = st.sidebar.select_slider(
        "Largura da página", options=["Estreita", "Padrão", "Larga", "Tela cheia"],
        value="Padrão",
    )

    return df, base_drop, dias_slide, {"altura": altura, "largura": largura,
                                       "modo_painel": modo_painel}


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


def navegar_slides(resumo: pd.DataFrame, dias_slide: str) -> pd.DataFrame:
    """Divide o período em blocos de dias, como as páginas de um slide."""
    if dias_slide == "Todos" or len(resumo) <= int(dias_slide):
        return resumo

    tamanho = int(dias_slide)
    total = -(-len(resumo) // tamanho)
    atual = min(st.session_state.get("slide", 1), total)

    esq, meio, dir_, _ = st.columns([0.09, 0.16, 0.09, 0.66])
    if esq.button("‹", key="slide_ant", disabled=atual == 1, width="stretch"):
        atual -= 1
    if dir_.button("›", key="slide_prox", disabled=atual == total, width="stretch"):
        atual += 1
    st.session_state["slide"] = atual
    meio.markdown(
        f'<div class="paginacao">{atual} / {total}</div>', unsafe_allow_html=True
    )
    return resumo.iloc[(atual - 1) * tamanho: atual * tamanho].reset_index(drop=True)


def linha_kpis(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str) -> None:
    """Resumo do período em cinco cartões."""
    rotas = resumo["ROTAS"].sum()
    peso = resumo["PESO"].sum()
    capacidade = resumo["CAPACIDADE"].sum()
    cartoes = [
        ("Rotas", num(rotas), f"{num(rotas / max(len(resumo), 1), 1)} por dia"),
        ("Veículos", num(resumo["VEICULOS"].max()), f"pico em um dia · {num(resumo['VEICULOS'].mean(), 1)} na média"),
        ("Ocupação", f"{num(peso / max(capacidade, 1) * 100, 1)}%", f"{num(peso)} de {num(capacidade)} kg"),
        ("Entregas", num(resumo["ENTREGAS"].sum()), f"{num(resumo['PARADAS'].sum())} paradas"),
        ("Drop médio", num(resumo[coluna_drop].mean(), 1), rotulo_drop),
    ]
    html = "".join(
        f'<div class="kpi"><div class="kpi-rot">{rot}</div>'
        f'<div class="kpi-val">{val}</div><div class="kpi-nota">{nota}</div></div>'
        for rot, val, nota in cartoes
    )
    st.markdown(f'<div class="kpis">{html}</div>', unsafe_allow_html=True)


# Espaço ocupado pelo seletor "Por dia / Ranking" e pelas mini estatísticas do Drop
COMPENSACAO_DROP = 92


def linha_um(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str,
             altura: int = 232) -> None:
    c1, c2, c3 = st.columns([1.05, 1.05, 0.9], gap="medium")
    altura_lateral = altura + COMPENSACAO_DROP

    with c1, st.container(border=True):
        titulo_painel("Veículos", "rotas / dia")
        faixa_numeros([num(v) for v in resumo["ROTAS"]], cor="escuro")
        st.plotly_chart(grafico_barras(resumo, "VEICULOS", altura=altura_lateral), width="stretch",
                        config={"displayModeBar": False}, key="g_veiculos")

    with c2, st.container(border=True):
        titulo_painel("Ocupação", "peso ÷ capacidade")
        faixa_numeros([f"{num(v * 100)}%" if pd.notna(v) else "—" for v in resumo["OCUPACAO"]],
                      cor="escuro")
        st.plotly_chart(grafico_barras(resumo, "OCUPACAO", altura=altura_lateral), width="stretch",
                        config={"displayModeBar": False}, key="g_ocupacao")

    with c3, st.container(border=True):
        titulo_painel("Drop", rotulo_drop)
        visao = st.segmented_control(
            "Visão do drop", ["Por dia", "Ranking"], default="Por dia",
            key="visao_drop", label_visibility="collapsed",
        ) or "Por dia"
        mini_estatisticas_drop(resumo, coluna_drop)
        if visao == "Por dia":
            faixa_numeros([num(v, 1) for v in resumo[coluna_drop]], cor="escuro")
            st.plotly_chart(grafico_drop_por_dia(resumo, coluna_drop, altura=altura),
                            width="stretch", config={"displayModeBar": False}, key="g_drop_dia")
        else:
            limite = 12
            st.plotly_chart(grafico_drop_ranking(resumo, coluna_drop, limite=limite, altura=altura),
                            width="stretch", config={"displayModeBar": False}, key="g_drop_rank")
            if len(resumo) > limite:
                st.markdown(
                    f'<div class="painel-nota">{limite} maiores de {len(resumo)} dias</div>',
                    unsafe_allow_html=True,
                )


def linha_dois(resumo: pd.DataFrame, altura: int = 250) -> None:
    c1, c2 = st.columns(2, gap="medium")

    with c1, st.container(border=True):
        titulo_painel("Paradas × entregas", "entregas / média de paradas")
        faixa_numeros([num(v) for v in resumo["ENTREGAS"]], cor="azul")
        faixa_numeros([num(v, 1) for v in resumo["MEDIA_PARADAS"]], cor="escuro")
        st.plotly_chart(
            grafico_duas_linhas(resumo, "MEDIA_PARADAS", "ENTREGAS", "Média paradas", "Entregas",
                                altura=altura),
            width="stretch", config={"displayModeBar": False}, key="g_paradas",
        )

    with c2, st.container(border=True):
        titulo_painel("Peso × capacidade por dia", "capacidade / peso (kg)")
        faixa_numeros([num(v) for v in resumo["CAPACIDADE"]], cor="azul")
        faixa_numeros([num(v) for v in resumo["PESO"]], cor="escuro")
        st.plotly_chart(
            grafico_duas_linhas(resumo, "PESO", "CAPACIDADE", "Peso (kg)", "Capacidade (kg)",
                                altura=altura),
            width="stretch", config={"displayModeBar": False}, key="g_peso",
        )


def figuras_do_painel(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str) -> list[dict]:
    """As mesmas cinco figuras do modo apresentação, prontas para a grade livre."""
    return [
        {"id": "veiculos", "titulo": "Veículos", "nota": "rotas / dia",
         "figura": grafico_barras(resumo, "VEICULOS")},
        {"id": "ocupacao", "titulo": "Ocupação", "nota": "peso ÷ capacidade",
         "figura": grafico_barras(resumo, "OCUPACAO")},
        {"id": "drop", "titulo": "Drop", "nota": rotulo_drop,
         "figura": grafico_drop_por_dia(resumo, coluna_drop)},
        {"id": "paradas", "titulo": "Paradas × entregas", "nota": "entregas / média de paradas",
         "figura": grafico_duas_linhas(resumo, "MEDIA_PARADAS", "ENTREGAS",
                                       "Média paradas", "Entregas")},
        {"id": "peso", "titulo": "Peso × capacidade", "nota": "capacidade / peso (kg)",
         "figura": grafico_duas_linhas(resumo, "PESO", "CAPACIDADE",
                                       "Peso (kg)", "Capacidade (kg)")},
    ]


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

    df, base_drop, dias_slide, tamanho = barra_lateral()
    st.markdown(
        f'<style>.block-container {{ max-width: {LARGURAS_PAGINA[tamanho["largura"]]} !important; }}</style>',
        unsafe_allow_html=True,
    )

    if df.empty:
        st.markdown(
            '<div class="abertura">'
            '<p class="cab-titulo">Dados e fatos</p>'
            '<div class="cab-sub">Indicadores de roteirização · Delly\'s Food Service</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        st.info(
            "Envie o relatório de rotas do RoadNet na barra lateral, ou coloque os arquivos "
            "na pasta **dados/** do repositório (ex.: `dados/AM.xlsx`) para o site abrir "
            "já preenchido."
        )
        return

    coluna_drop = {"Parada": "DROP_PARADA", "Veículo": "DROP_VEICULO", "Rota": "DROP_ROTA"}[base_drop]
    rotulo_drop = {"Parada": "kg / parada", "Veículo": "kg / veículo", "Rota": "kg / rota"}[base_drop]

    resumo = indicadores_por_dia(df)
    if resumo.empty:
        st.warning("Nenhuma rota no filtro selecionado.")
        return

    pagina = navegar_slides(resumo, dias_slide)
    cabecalho(df["UF"].iloc[0], pagina)
    if tamanho["modo_painel"] and renderizar_painel is None:
        st.warning(
            "O modo painel precisa do arquivo **painel_arrastavel.py** na mesma pasta do "
            "app.py, no repositório. Mostrando o modo apresentação."
        )
    if tamanho["modo_painel"] and renderizar_painel is not None:
        renderizar_painel(
            figuras_do_painel(pagina, coluna_drop, rotulo_drop),
            chave=df["UF"].iloc[0],
            altura_celula=max(50, tamanho["altura"] // 4),
            altura_total=max(560, tamanho["altura"] * 3),
        )
    else:
        linha_um(pagina, coluna_drop, rotulo_drop, altura=tamanho["altura"])
        linha_dois(pagina, altura=tamanho["altura"] + COMPENSACAO_DROP)
    with st.expander("Resumo do período inteiro"):
        linha_kpis(resumo, coluna_drop, rotulo_drop)
    tabela_detalhe(resumo, coluna_drop, rotulo_drop)

    st.markdown(
        f'<div class="fonte">Fonte: relatório de rotas RoadNet · '
        f'{num(len(df))} rotas processadas · arquivos: {", ".join(sorted(df["ARQUIVO"].unique()))}</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
