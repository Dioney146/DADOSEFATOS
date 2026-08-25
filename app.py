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

import base64
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
    page_title="Indicadores Delly's",
    page_icon=str(Path(__file__).parent / "assets" / "favicon.png")
    if (Path(__file__).parent / "assets" / "favicon.png").exists() else "📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PASTA_DADOS = Path(__file__).parent / "dados"
PASTA_ASSETS = Path(__file__).parent / "assets"
ARQUIVO_LOGO = PASTA_ASSETS / "logo.png"
ARQUIVO_ICONE = PASTA_ASSETS / "favicon.png"

# Estados atendidos. A chave é o código usado no nome do arquivo.
ESTADOS = {
    "TODOS": "Consolidado · todos os estados",
    "AM": "Amazonas",
    "BA": "Bahia",
    "DF": "Distrito Federal",
    "ES": "Espírito Santo",
    "MG": "Minas Gerais",
    "MG_NF": "Minas Gerais (NF)",
    "MT": "Mato Grosso",
    "SP": "São Paulo",
    "SP_WFS": "São Paulo (W Food)",
    "SP_3P": "São Paulo (3P)",
}

MESES_CURTOS = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
                7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}

# Paleta
COR_TEXTO = "#14161A"
COR_SUAVE = "#8B949E"
COR_AZUL = "#5D87B0"
COR_AZUL_CLARO = "#8FB3D4"
COR_ESCURA = "#2C3440"
COR_TRILHO = "#E2E4E6"
COR_FUNDO_GRAFICO = "rgba(0,0,0,0)"
COR_GRADE = "#E4E8EC"
COR_BORDA = "#E3E7EB"

TAMANHO_PADRAO = {"altura": 100, "modo_painel": False}

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
    "SEMANA",
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
.stApp, [data-testid="stAppViewContainer"] { background: #F2F5F9 !important; }
header[data-testid="stHeader"] {
    background: #F2F5F9 !important; height: 2.6rem; z-index: 999;
}
.block-container {
    padding-top: 1.7rem !important; padding-bottom: 6px !important;
    padding-left: clamp(8px, 0.8vw, 20px) !important;
    padding-right: clamp(8px, 0.8vw, 20px) !important;
    max-width: 100% !important; width: 100%; margin: 0;
}
/* O conteúdo do Streamlit não pode limitar a largura por conta própria */
[data-testid="stAppViewBlockContainer"], [data-testid="stMainBlockContainer"] {
    max-width: none !important;
}
[data-testid="stMain"] .stMainBlockContainer { width: 100%; }
html, body, .stApp, .stApp p, .stApp span, .stApp label, .stApp li,
.stApp h1, .stApp h2, .stApp h3, .stApp h4 {
    color: #14161A;
    font-family: 'Archivo', Arial, sans-serif;
}

/* ── Barra superior: identificação à esquerda, mini KPIs à direita ──────── */
.cab {
    display: flex; align-items: center; justify-content: space-between;
    gap: 20px; margin: -2px 0 10px 0;
    background: #FFFFFF; border: 1px solid #EAEEF3; border-radius: 14px;
    padding: clamp(8px, 0.7vw, 14px) clamp(12px, 1vw, 20px);
    box-shadow: 0 1px 2px rgba(20, 22, 26, .03), 0 6px 20px rgba(20, 22, 26, .04);
}
.cab-id { display: flex; align-items: center; gap: 12px; }
.cab-barra {
    width: 4px; height: clamp(30px, 2.6vw, 44px); border-radius: 4px;
    background: linear-gradient(180deg, #5D87B0 0%, #2C3440 100%); display: block;
}
.cab-titulo {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif;
    font-weight: 700; letter-spacing: .01em;
    font-size: clamp(24px, 2.0vw, 38px); line-height: 1; text-transform: uppercase;
    margin: 0; color: #14161A;
}
.cab-sub {
    font-family: 'Archivo', Arial, sans-serif; font-size: clamp(9px, .72vw, 11px);
    font-weight: 600; letter-spacing: .13em; text-transform: uppercase;
    color: #8B949E; margin-top: 3px;
}
.cab-sep { color: #C7CFD8; margin: 0 3px; }

.cab-meta { display: flex; gap: 8px; }
.kpi-chip {
    background: #FFFFFF; border: 1px solid #E7EBF0; border-radius: 12px;
    padding: 7px 16px; min-width: 92px; text-align: right;
    box-shadow: 0 1px 2px rgba(20, 22, 26, .04);
}
.kpi-chip.destaque { background: #14161A; border-color: #14161A; }
.kpi-rot {
    font-family: 'Archivo', Arial, sans-serif; font-size: 9px; font-weight: 600;
    letter-spacing: .14em; text-transform: uppercase; color: #98A2AE;
}
.kpi-val {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700;
    font-size: clamp(18px, 1.45vw, 27px); line-height: 1.15; color: #14161A;
    white-space: nowrap;
}
.kpi-chip.destaque .kpi-val { color: #FFFFFF; }
.kpi-chip.destaque .kpi-rot { color: #8FB3D4; }

/* ── Painéis (st.container com borda) ────────────────────────────────────── */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: #FFFFFF; border: 1px solid #EAEEF3 !important; border-radius: 14px !important;
    padding: clamp(8px, 0.7vw, 14px) clamp(8px, 0.7vw, 16px) clamp(4px, 0.4vw, 8px);
    box-shadow: 0 1px 2px rgba(20, 22, 26, .03), 0 6px 20px rgba(20, 22, 26, .045);
    transition: box-shadow .18s ease, transform .18s ease;
}
div[data-testid="stMain"] div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 2px 4px rgba(20, 22, 26, .05), 0 10px 28px rgba(45, 90, 140, .10);
    transform: translateY(-1px);
}
div[data-testid="stVerticalBlockBorderWrapper"] .js-plotly-plot,
div[data-testid="stVerticalBlockBorderWrapper"] .stPlotlyChart { width: 100% !important; }
/* Espaçamento enxuto entre painéis e entre as linhas */
div[data-testid="stHorizontalBlock"] {
    gap: clamp(6px, 0.7vw, 16px) !important; margin-bottom: clamp(6px, 0.7vw, 16px);
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stElementContainer"] {
    margin-bottom: 0 !important;
}
div[data-testid="stVerticalBlock"] { gap: clamp(6px, 0.7vw, 16px) !important; }

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
    display: flex; align-items: center; justify-content: space-between;
    gap: 10px; padding: 0 0 6px 0; margin: 0;
}
.painel-titulo {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700;
    font-size: clamp(16px, 1.15vw, 23px);
    text-transform: uppercase; letter-spacing: .01em; color: #14161A; line-height: 1;
}
.painel-nota {
    font-family: 'Archivo', Arial, sans-serif; font-size: clamp(8px, .62vw, 10px);
    font-weight: 600; letter-spacing: .12em; text-transform: uppercase; color: #5D87B0;
    background: #EFF4FA; border-radius: 20px; padding: 3px 10px; white-space: nowrap;
}
/* Legenda integrada ao cabeçalho do card */
.painel-series { display: flex; gap: 12px; align-items: center; }
.serie {
    display: inline-flex; align-items: center; gap: 6px;
    font-family: 'Archivo', Arial, sans-serif; font-size: clamp(8px, .62vw, 10px);
    font-weight: 600; letter-spacing: .1em; text-transform: uppercase; color: #6B7684;
    white-space: nowrap;
}
.serie i { width: 14px; height: 3px; border-radius: 2px; display: inline-block; }

/* Abas do topo coladas no cabeçalho */
div[data-testid="stTabs"] [data-baseweb="tab-list"] { gap: 18px; margin-bottom: 2px; }
div[data-testid="stTabs"] [data-baseweb="tab-panel"] { padding-top: 4px; }
div[data-testid="stTabs"] [data-baseweb="tab"] { padding: 2px 0; }

/* ── Cards de semana (topo da aba diária) ───────────────────────────────── */
div[data-testid="stHorizontalBlock"] button[kind="secondary"],
div[data-testid="stHorizontalBlock"] button[kind="primary"] {
    line-height: 1.1; border-radius: 10px; padding: 6px 4px; min-height: 0;
    font-family: 'Archivo', Arial, sans-serif; font-size: 11px; font-weight: 700;
    letter-spacing: .04em; text-transform: uppercase;
    box-shadow: 0 1px 2px rgba(20, 22, 26, .04);
    transition: transform .15s ease, box-shadow .15s ease;
}
div[data-testid="stHorizontalBlock"] button[kind="secondary"] {
    background: #FFFFFF; border: 1px solid #E7EBF0; color: #48525E;
}
div[data-testid="stHorizontalBlock"] button[kind="secondary"]:hover {
    border-color: #5D87B0; color: #14161A; transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(45, 90, 140, .12);
}
div[data-testid="stHorizontalBlock"] button[kind="primary"] {
    background: #2F5B87; border: 1px solid #2F5B87; color: #FFFFFF;
}

/* Paginação dos slides */
.paginacao {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700; font-size: 20px;
    text-align: center; color: #14161A; padding-top: 4px;
}

/* ── Faixa de números acima dos gráficos ─────────────────────────────────── */
.faixa {
    display: grid; width: 100%; max-width: 100%; margin: 2px 0 4px 0;
    padding: 0; overflow: hidden; box-sizing: border-box;
}
.faixa-cel {
    min-width: 0; max-width: 100%; text-align: center;
    overflow: hidden; white-space: nowrap; text-overflow: clip;
    line-height: 1.15; font-size: inherit;
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700;
}
.faixa-eixo {
    display: grid; width: 100%; max-width: 100%; box-sizing: border-box;
    margin: 2px 0 0 0; padding: 0; overflow: hidden;
    font-family: 'Archivo', Arial, sans-serif; font-weight: 600; color: #98A2AE;
    letter-spacing: .02em;
}
.eixo-cel { min-width: 0; text-align: center; overflow: hidden; white-space: nowrap; }

/* Hierarquia de cor: valor principal em grafite, apoio em azul, secundário em cinza */
.faixa-forte { color: #14161A; }
.faixa-azul { color: #3F6E9C; }
.faixa-suave { color: #7A8794; font-weight: 600; }

/* ── Mini estatísticas (painel de drop) ──────────────────────────────────── */
.mini {
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px;
    background: #F7F9FC; border: 1px solid #EEF2F7; border-radius: 10px;
    padding: 8px 10px; margin: 2px 0 6px 0;
}
.mini-item { text-align: center; }
.mini-item + .mini-item { border-left: 1px solid #E7ECF2; }
.mini-rot {
    font-family: 'Archivo', Arial, sans-serif; font-size: 8px; font-weight: 600;
    letter-spacing: .14em; text-transform: uppercase; color: #98A2AE;
}
.mini-val {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif; font-weight: 700;
    font-size: clamp(20px, 1.5vw, 28px); line-height: 1.05; color: #14161A;
}
.mini-item:first-child .mini-val { color: #3F6E9C; }
.mini-dia {
    font-family: 'Archivo', Arial, sans-serif; font-size: 9px; color: #98A2AE; letter-spacing: .04em;
}

/* ── Visão por semana ────────────────────────────────────────────────────── */
.legenda { display: flex; gap: 8px 20px; justify-content: flex-end;
           flex-wrap: wrap; margin: -6px 0 10px 0; }
.leg-item {
    font-size: 11px; letter-spacing: .08em; text-transform: uppercase; color: #14161A;
    display: inline-flex; align-items: center; gap: 7px;
}
.leg-cor { width: 12px; height: 12px; display: inline-block; }
.sem-faixa {
    display: grid; width: 100%; box-sizing: border-box;
    border-top: 1px solid #E3E7EB; padding-top: 8px; margin: 0 0 4px 0;
}
.sem-cel { text-align: center; min-width: 0; overflow: hidden; }
.sem-val {
    font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif;
    font-weight: 700; font-size: 27px; line-height: 1.05; color: #14161A;
    white-space: nowrap;
}
/* Valores compridos (peso, capacidade) em corpo menor, para não se colarem */
.sem-cel.compacta .sem-val { font-size: 19px; }
.sem-cel.compacta .sem-rot { font-size: 9px; letter-spacing: .04em; }
.sem-rot {
    font-family: 'IBM Plex Mono', 'Consolas', monospace; font-size: 10px;
    letter-spacing: .1em; text-transform: uppercase; color: #7C858D;
}

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

/* Telas estreitas: os painéis passam a ocupar a linha inteira, sem rolagem lateral */
@media (max-width: 900px) {
    div[data-testid="stHorizontalBlock"] { flex-wrap: wrap; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important; min-width: 100% !important;
    }
    .cab { flex-wrap: wrap; gap: 10px; }
}
html, body { overflow-x: hidden; }

/* Componente invisível que reajusta os gráficos ao mudar a largura */
iframe[title="st.iframe"], iframe[height="0"] { height: 0 !important; display: block; }
div[data-testid="stElementContainer"]:has(iframe[height="0"]) { display: none !important; }

/* Menu do canto (System/Light/Dark) escondido: o site é sempre claro */
#MainMenu, [data-testid="stMainMenu"] { display: none !important; }

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


def rotulo_semana(data: pd.Timestamp, varios_meses: bool) -> str:
    """Semana do mês, de segunda a domingo. S1 é a semana em que cai o dia 1."""
    if pd.isna(data):
        return ""
    primeiro = data.replace(day=1)
    inicio_mes = primeiro - pd.Timedelta(days=primeiro.weekday())
    inicio_semana = data - pd.Timedelta(days=data.weekday())
    numero = int((inicio_semana - inicio_mes).days // 7) + 1
    return f"{MESES_CURTOS[data.month]}/S{numero}" if varios_meses else f"S{numero}"


def detectar_estado(nome_arquivo: str, df: pd.DataFrame) -> str:
    """Descobre o estado pelo nome do arquivo; se falhar, pelo prefixo da descrição."""
    base = re.sub(r"[^A-Z0-9]", "", normalizar(Path(nome_arquivo).stem).upper())
    fichas = re.sub(r"[^A-Z0-9]", " ", normalizar(Path(nome_arquivo).stem).upper()).split()
    for codigo in sorted(ESTADOS, key=len, reverse=True):
        compacto = codigo.replace("_", "")
        if base.startswith(compacto) or codigo in fichas or compacto in fichas:
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
    placa = df["Equipamento"].astype(str).str.strip() if "Equipamento" in df.columns else ""
    df["PLACA"] = placa
    df["TIPO_VEICULO"] = (
        df["Tipos de equipamento"].astype(str).str.strip().replace({"": "NÃO INFORMADO", "nan": "NÃO INFORMADO"})
        if "Tipos de equipamento" in df.columns else "NÃO INFORMADO"
    )
    df["STATUS"] = (
        df["Estado"].astype(str).str.strip().replace({"": "NÃO INFORMADO", "nan": "NÃO INFORMADO"})
        if "Estado" in df.columns else "NÃO INFORMADO"
    )
    df["UF"] = detectar_estado(nome_arquivo, df)
    # No consolidado, a mesma placa em estados diferentes é outra frota:
    # o estado entra na chave para a contagem não juntar veículos distintos.
    df["VEICULO"] = df["UF"] + "·" + df["PLACA"].astype(str)
    df["ARQUIVO"] = Path(nome_arquivo).name

    df = df.dropna(subset=["DATA"])
    df = df[df["CAPACIDADE"].fillna(0) > 0]

    # A coluna SEMANA do arquivo é usada quando existe; senão é calculada aqui.
    varios_meses = df["DATA"].dt.month.nunique() > 1 if not df.empty else False
    if "SEMANA" in df.columns and df["SEMANA"].notna().any():
        df["SEMANA"] = df["SEMANA"].astype(str).str.strip()
    else:
        df["SEMANA"] = df["DATA"].map(lambda d: rotulo_semana(d, varios_meses))

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
    """
    Procura planilhas na pasta dados/ e também na raiz do repositório, para o
    caso de o arquivo ter sido enviado ao lado do app.py.
    """
    extensoes = {".xlsx", ".xlsm", ".xls", ".csv"}
    achados: list[tuple[str, bytes]] = []
    vistos: set[str] = set()

    for pasta in (PASTA_DADOS, Path(__file__).parent):
        if not pasta.exists():
            continue
        for caminho in sorted(pasta.iterdir()):
            if not caminho.is_file() or caminho.suffix.lower() not in extensoes:
                continue
            if caminho.name.startswith("~$") or caminho.name in vistos:
                continue
            vistos.add(caminho.name)
            achados.append((caminho.name, caminho.read_bytes()))
    return achados


# ──────────────────────────────────────────────────────────────────────────────
# INDICADORES
# ──────────────────────────────────────────────────────────────────────────────

def indicadores_por_dia(df: pd.DataFrame, por: str = "Dia") -> pd.DataFrame:
    """Uma linha por data (ou por semana) com os indicadores da apresentação."""
    if df.empty:
        return pd.DataFrame()

    if por == "Semana":
        chave = df["SEMANA"]
        ordem = df.groupby("SEMANA")["DATA"].min()
    else:
        chave = df["DATA"]

    agrupado = df.groupby(chave).agg(
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
    agrupado["OCUPACAO_PCT"] = agrupado["OCUPACAO"] * 100
    agrupado["MEDIA_PARADAS"] = agrupado["PARADAS"] / agrupado["ROTAS"].replace(0, pd.NA)
    agrupado["DROP_PARADA"] = agrupado["PESO"] / agrupado["PARADAS"].replace(0, pd.NA)
    agrupado["DROP_VEICULO"] = agrupado["PESO"] / agrupado["VEICULOS"].replace(0, pd.NA)
    agrupado["DROP_ROTA"] = agrupado["PESO"] / agrupado["ROTAS"].replace(0, pd.NA)
    if por == "Semana":
        agrupado["ROTULO"] = agrupado["SEMANA"]
        agrupado["EIXO"] = agrupado["SEMANA"]
        agrupado["ORDEM"] = agrupado["SEMANA"].map(ordem)
        agrupado["INICIO"] = agrupado["SEMANA"].map(df.groupby("SEMANA")["DATA"].min())
        agrupado["FIM"] = agrupado["SEMANA"].map(df.groupby("SEMANA")["DATA"].max())
        agrupado["DIAS"] = agrupado["SEMANA"].map(df.groupby("SEMANA")["DATA"].nunique())
        # Veículos da semana = soma dos veículos usados em cada dia
        # (uma placa que rodou 5 dias conta 5 vezes).
        por_dia = df.groupby(["SEMANA", "DATA"])["VEICULO"].nunique()
        agrupado["VEICULOS_DISTINTOS"] = agrupado["VEICULOS"]
        agrupado["VEICULOS"] = agrupado["SEMANA"].map(por_dia.groupby("SEMANA").sum())
        pronto = agrupado.sort_values("ORDEM").drop(columns="ORDEM").reset_index(drop=True)
        pronto["HOVER"] = [
            f"{r} · {i.strftime('%d/%m')}–{f.strftime('%d/%m')}"
            for r, i, f in zip(pronto["ROTULO"], pronto["INICIO"], pronto["FIM"])
        ]
        return pronto

    agrupado["ROTULO"] = agrupado["DATA"].dt.strftime("%d/%m")
    # O eixo mostra só o dia; se o recorte cruzar meses, mantém dia/mês para não repetir
    um_mes = agrupado["DATA"].dt.to_period("M").nunique() <= 1
    agrupado["EIXO"] = agrupado["DATA"].dt.strftime("%d" if um_mes else "%d/%m")
    agrupado["HOVER"] = agrupado["DATA"].dt.strftime("%d/%m/%Y")  # tooltip com a data inteira
    return agrupado.sort_values("DATA").reset_index(drop=True)


def toneladas(valor) -> str:
    """Peso em toneladas, curto: 67.900 kg vira 68t."""
    if valor is None or pd.isna(valor):
        return "—"
    return f"{num(valor / 1000)}t"


def num(valor, casas: int = 0) -> str:
    """Formata número no padrão pt-BR."""
    if valor is None or pd.isna(valor):
        return "—"
    texto = f"{float(valor):,.{casas}f}"
    return texto.replace(",", "@").replace(".", ",").replace("@", ".")


# ──────────────────────────────────────────────────────────────────────────────
# COMPONENTES VISUAIS
# ──────────────────────────────────────────────────────────────────────────────

CONFIG_GRAFICO = {"displayModeBar": False, "responsive": True}

LAYOUT_BASE = dict(
    margin=dict(l=0, r=0, t=4, b=22),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor=COR_FUNDO_GRAFICO,
    font=dict(family="Archivo, sans-serif", size=11, color=COR_TEXTO),
    showlegend=False,
    hoverlabel=dict(font_size=12, font_family="Archivo, sans-serif",
                    bgcolor="#FFFFFF", bordercolor="#E7EBF0",
                    font=dict(color=COR_TEXTO)),
    hovermode="x unified",
    transition=dict(duration=220, easing="cubic-in-out"),
)

EIXO_X = dict(showgrid=False, zeroline=False, showline=False, ticks="",
              showspikes=True, spikemode="toaxis", spikethickness=1,
              spikedash="dot", spikecolor="#C7D2DD", spikesnap="cursor",
              tickfont=dict(size=10, color=COR_SUAVE))
EIXO_Y = dict(showgrid=True, gridcolor=COR_GRADE, griddash="solid", gridwidth=1,
              zeroline=False, showline=False, nticks=4,
              tickfont=dict(size=10, color=COR_SUAVE))


def titulo_painel(titulo: str, nota: str = "", linha: int = 0,
                  series: list[tuple[str, str]] | None = None) -> None:
    """
    Cabeçalho do painel: título à esquerda, unidade à direita.

    `series` transforma a legenda do gráfico em etiquetas do próprio cabeçalho
    (cor + nome), tirando a legenda de dentro da área de plotagem.
    """
    marca = f'<i class="marca-linha-{linha}"></i>' if linha else ""
    if series:
        etiquetas = "".join(
            f'<span class="serie"><i style="background:{cor}"></i>{nome}</span>'
            for nome, cor in series
        )
        direita = f'<span class="painel-series">{etiquetas}</span>'
    else:
        direita = f'<span class="painel-nota">{nota}</span>'
    st.markdown(
        f'''<div class="painel-topo">{marca}
            <span class="painel-titulo">{titulo}</span>
            {direita}</div>''',
        unsafe_allow_html=True,
    )


LIMITE_FAIXA = 45  # até aqui os números cabem; acima disso ficam só no tooltip


def entre(valor: float, minimo: float, maximo: float) -> float:
    """Mantém o valor dentro de um intervalo."""
    return max(minimo, min(maximo, valor))


def escala(quantidade: int, com_poucos: float, com_muitos: float) -> float:
    """
    Interpola entre o tamanho usado com 7 colunas e o usado com 31; de 31 em
    diante segue comprimindo até 62 colunas (dois meses), com metade do passo.
    """
    proporcao = entre((quantidade - 7) / 24, 0, 1)
    valor = com_poucos + (com_muitos - com_poucos) * proporcao
    if quantidade > 31:
        extra = entre((quantidade - 31) / 31, 0, 1) * 0.5
        valor += (com_muitos - com_poucos) * extra
    return valor


def corpo_do_eixo(quantidade: int) -> int:
    """Rótulos do eixo: 13px com 7 dias, caindo até 8px com 31."""
    return max(7, int(round(escala(quantidade, 13, 8))))


def largura_da_barra(quantidade: int) -> tuple[float, float]:
    """
    Barra e vão proporcionais: com poucos dias a barra é mais estreita e o vão
    maior; com o mês cheio a barra engorda e o vão encolhe, sem sobrepor.
    """
    vao = max(0.08, escala(quantidade, 0.42, 0.14))
    return 1.0 - vao, vao


# Fração da célula que o número pode ocupar: o resto vira respiro entre eles.
OCUPACAO_DA_CELULA = 0.78


def corpo_dos_valores(quantidade: int) -> tuple[int, str]:
    """Corpo da fonte e espaçamento, interpolados pela quantidade de colunas."""
    corpo = int(round(escala(quantidade, 19, 9)))
    aperto = escala(quantidade, 0.0, -0.05)
    return corpo, f"{aperto:.3f}em"


def faixa_numeros(valores: list[str], cor: str = "forte", ajuste: int = 0,
                  fracao: float = 0.30) -> None:
    """
    Linha de números alinhada com as colunas do gráfico.

    O corpo da fonte tem duas travas: uma pela quantidade de colunas e outra
    pela largura real disponível na tela (fracao = quanto da largura da janela
    este card ocupa). Vence a menor, então o número nunca passa da célula.
    """
    if not valores or len(valores) > LIMITE_FAIXA:
        return

    quantidade = len(valores)
    corpo, espaco = corpo_dos_valores(quantidade)
    mais_longo = max((len(str(v)) for v in valores), default=2)
    if quantidade > 16 and mais_longo > 2:
        corpo -= mais_longo - 2
    corpo = max(6, corpo + ajuste)

    # Largura de cada célula em vw. O número só pode usar parte dela (o resto é
    # o vão que impede um encostar no outro); 0,58em é a largura de um dígito.
    por_celula = (fracao * 100) / quantidade
    teto_vw = (por_celula * OCUPACAO_DA_CELULA) / (mais_longo * 0.58)

    classe = {"azul": "faixa-azul", "suave": "faixa-suave"}.get(cor, "faixa-forte")
    st.markdown(
        f'<div class="faixa" style="grid-template-columns: repeat({quantidade}, minmax(0, 1fr)); '
        f'font-size: min({corpo}px, {teto_vw:.3f}vw); letter-spacing:{espaco}">'
        + "".join(f'<div class="faixa-cel {classe}">{v}</div>' for v in valores)
        + "</div>",
        unsafe_allow_html=True,
    )


def faixa_eixo(valores: list[str], fracao: float = 0.30) -> None:
    """
    Rótulos dos dias, na mesma grade dos números acima.

    Desenhar o eixo em HTML (e não dentro do gráfico) garante que cada dia caia
    exatamente na mesma coluna do valor correspondente, independentemente da
    largura que o Plotly resolva usar.
    """
    if not valores:
        return
    quantidade = len(valores)
    passo = passo_dos_rotulos(quantidade, fracao)
    corpo = corpo_do_eixo(quantidade)
    celulas = "".join(
        f'<div class="eixo-cel">{v if i % passo == 0 else ""}</div>'
        for i, v in enumerate(valores)
    )
    st.markdown(
        f'<div class="faixa-eixo" '
        f'style="grid-template-columns: repeat({quantidade}, minmax(0, 1fr)); '
        f'font-size:{corpo}px">{celulas}</div>',
        unsafe_allow_html=True,
    )


def passo_dos_rotulos(quantidade: int, fracao: float) -> int:
    """
    De quantos em quantos dias escrever o rótulo do eixo.

    Estima a largura de cada coluna (em pixels, numa janela de 1600px) e compara
    com o espaço que o texto ocupa. Se não couber com folga, escreve dia sim,
    dia não — as barras continuam todas lá, só o rótulo é que alterna.
    """
    largura_coluna = (fracao * 1600) / max(quantidade, 1)
    largura_texto = corpo_do_eixo(quantidade) * 2 * 0.62 + 6  # 2 dígitos + respiro
    passo = 1
    while largura_coluna * passo < largura_texto and passo < 5:
        passo += 1
    return passo


def eixo_datas(fig: go.Figure, dados: pd.DataFrame, fracao: float = 0.30) -> None:
    """
    Todas as colunas aparecem e o eixo começa e termina rente à borda, para os
    rótulos dos dias ficarem na mesma vertical dos números da faixa acima.
    """
    quantidade = len(dados)
    fig.update_layout(autosize=True, margin=dict(l=0, r=0, t=4, b=2))
    # Os rótulos dos dias são desenhados fora do gráfico (ver faixa_eixo),
    # para ficarem alinhados com os números da faixa de cima.
    fig.update_xaxes(**EIXO_X, type="category", showticklabels=False,
                     automargin=False, range=[-0.5, quantidade - 0.5])


def grafico_barras(dados: pd.DataFrame, coluna: str, altura: int = 300,
                   fracao: float = 0.30, formato: str = "%{y:,.0f}") -> go.Figure:
    """`formato` define como o valor aparece no tooltip (ex.: '%{y:.0%}')."""
    largura, vao = largura_da_barra(len(dados))
    fig = go.Figure(
        go.Bar(
            x=dados["EIXO"], y=dados[coluna], customdata=dados["HOVER"],
            marker_color=COR_AZUL, marker_line_width=0, width=largura,
            marker=dict(cornerradius=4),
            hovertemplate="<b>%{customdata}</b><br>" + formato + "<extra></extra>",
        )
    )
    fig.update_layout(**LAYOUT_BASE, height=altura, bargap=vao)
    eixo_datas(fig, dados, fracao)
    fig.update_yaxes(**EIXO_Y, showticklabels=False)
    return fig


def cores_pela_media(valores, media: float) -> list[str]:
    """Dias acima da média em azul cheio; abaixo, em azul claro."""
    return [COR_AZUL if (pd.notna(v) and v >= media) else COR_AZUL_CLARO for v in valores]


def grafico_drop_por_dia(dados: pd.DataFrame, coluna: str, altura: int = 250,
                         fracao: float = 0.30) -> go.Figure:
    """Drop em ordem cronológica, no mesmo eixo dos demais painéis do dia."""
    media = float(dados[coluna].mean())
    largura, vao = largura_da_barra(len(dados))
    fig = go.Figure(
        go.Bar(
            x=dados["EIXO"], y=dados[coluna], width=largura, customdata=dados["HOVER"],
            marker_color=cores_pela_media(dados[coluna], media), marker_line_width=0,
            marker=dict(cornerradius=4),
            hovertemplate="<b>%{customdata}</b><br>%{y:,.0f} kg<extra></extra>",
        )
    )
    fig.add_hline(
        y=media, line_width=1, line_dash="dot", line_color=COR_ESCURA,
        annotation_text=f"média {num(media)}", annotation_position="top left",
        annotation_font=dict(family="IBM Plex Mono, monospace", size=9, color=COR_SUAVE),
    )
    fig.update_layout(**LAYOUT_BASE, height=altura, bargap=vao)
    eixo_datas(fig, dados, fracao)
    fig.update_yaxes(**EIXO_Y, showticklabels=False)
    return fig


def ranking_estados(df: pd.DataFrame, coluna: str) -> pd.Series:
    """
    Valor do indicador em cada estado, no período filtrado.

    Usa a mesma conta do painel, só que aplicada a um estado por vez: ocupação
    e drop vêm dos totais do estado; veículos, da média diária de placas.
    """
    valores = {}
    for uf, parte in df.groupby("UF"):
        dia = indicadores_por_dia(parte, por="Dia")
        if dia.empty:
            continue
        if coluna == "VEICULOS":
            valores[uf] = dia["VEICULOS"].mean()
        elif coluna == "OCUPACAO_PCT":
            valores[uf] = dia["PESO"].sum() / max(dia["CAPACIDADE"].sum(), 1) * 100
        elif coluna == "DROP_PARADA":
            valores[uf] = dia["PESO"].sum() / max(dia["PARADAS"].sum(), 1)
        elif coluna == "DROP_VEICULO":
            valores[uf] = dia["PESO"].sum() / max(dia["VEICULOS"].sum(), 1)
        elif coluna == "DROP_ROTA":
            valores[uf] = dia["PESO"].sum() / max(dia["ROTAS"].sum(), 1)
        else:
            valores[uf] = dia[coluna].mean()
    return pd.Series(valores, dtype="float64")


def mini_indicadores(dados: pd.DataFrame, coluna: str, sufixo: str = "",
                     rotulos: tuple[str, str, str] = ("Média", "Máximo", "Mínimo"),
                     por_estado: pd.Series | None = None) -> None:
    """
    Resumo do indicador dentro do próprio card.

    Por padrão mostra a média do período e os dias de maior e menor valor.
    No consolidado (`por_estado` preenchido), o maior e o menor passam a ser
    os ESTADOS, com a sigla no lugar da data.
    """
    validos = dados.dropna(subset=[coluna])
    if validos.empty:
        return

    if por_estado is not None and len(por_estado.dropna()) > 1:
        classificado = por_estado.dropna()
        topo, base = classificado.idxmax(), classificado.idxmin()
        blocos = [
            (rotulos[0], num(validos[coluna].mean()) + sufixo, "período"),
            ("Maior estado", num(classificado[topo]) + sufixo, topo),
            ("Menor estado", num(classificado[base]) + sufixo, base),
        ]
        html = "".join(
            f'<div class="mini-item"><div class="mini-rot">{rot}</div>'
            f'<div class="mini-val">{val}</div><div class="mini-dia">{dia}</div></div>'
            for rot, val, dia in blocos
        )
        st.markdown(f'<div class="mini">{html}</div>', unsafe_allow_html=True)
        return

    melhor = validos.loc[validos[coluna].idxmax()]
    pior = validos.loc[validos[coluna].idxmin()]
    blocos = [
        (rotulos[0], num(validos[coluna].mean()) + sufixo, "período"),
        (rotulos[1], num(melhor[coluna]) + sufixo, melhor["ROTULO"]),
        (rotulos[2], num(pior[coluna]) + sufixo, pior["ROTULO"]),
    ]
    html = "".join(
        f'<div class="mini-item"><div class="mini-rot">{rot}</div>'
        f'<div class="mini-val">{val}</div><div class="mini-dia">{dia}</div></div>'
        for rot, val, dia in blocos
    )
    st.markdown(f'<div class="mini">{html}</div>', unsafe_allow_html=True)


def grafico_duas_linhas(dados: pd.DataFrame, col_a: str, col_b: str,
                        nome_a: str, nome_b: str, altura: int = 330,
                        fracao: float = 0.485) -> go.Figure:
    fig = go.Figure()
    marcador = max(2.5, escala(len(dados), 7, 3))
    traco = max(1.1, escala(len(dados), 2.4, 1.4))
    fig.add_trace(go.Scatter(
        x=dados["EIXO"], y=dados[col_a], name=nome_a, mode="lines+markers",
        customdata=dados["HOVER"],
        line=dict(color=COR_ESCURA, width=traco, shape="spline", smoothing=0.9, dash="dot"),
        marker=dict(size=marcador, symbol="circle"),
        yaxis="y", hovertemplate=f"{nome_a}: %{{y:,.0f}}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=dados["EIXO"], y=dados[col_b], name=nome_b, mode="lines+markers",
        customdata=dados["HOVER"],
        line=dict(color=COR_AZUL, width=traco + 0.6, shape="spline", smoothing=0.9),
        marker=dict(size=marcador, symbol="circle"),
        yaxis="y2", hovertemplate=f"{nome_b}: %{{y:,.0f}}<extra></extra>",
    ))
    base = {k: v for k, v in LAYOUT_BASE.items() if k not in {"showlegend", "hovermode"}}
    fig.update_layout(
        **base, height=altura, showlegend=False, hovermode="x unified",
        yaxis={**EIXO_Y, "showticklabels": False},
        yaxis2=dict(overlaying="y", side="right", visible=False),
    )
    eixo_datas(fig, dados, fracao)
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# PÁGINA
# ──────────────────────────────────────────────────────────────────────────────

def periodo_do_mes_atual(df: pd.DataFrame):
    """
    Período que o site abre por padrão: o mês corrente.

    Se ainda não houver rotas do mês de hoje na base (fim de mês, arquivo não
    atualizado), cai para o mês mais recente que existir nos dados.
    """
    hoje = pd.Timestamp.today().normalize()
    mes_alvo = hoje.to_period("M")
    meses = df["DATA"].dt.to_period("M")
    if not (meses == mes_alvo).any():
        mes_alvo = meses.max()

    do_mes = df.loc[meses == mes_alvo, "DATA"]
    return do_mes.min().date(), do_mes.max().date()


def barra_lateral() -> tuple[pd.DataFrame, str]:
    # Os dados vêm apenas do repositório (pasta dados/ ou raiz), sem upload na tela.
    arquivos = arquivos_da_pasta()

    if not arquivos:
        return pd.DataFrame(), "Parada", "31", TAMANHO_PADRAO

    st.session_state.pop("_problemas", None)
    df = carregar(tuple(arquivos))
    for aviso in st.session_state.get("_problemas", []):
        st.sidebar.warning(aviso)
    if df.empty:
        return df, "Parada", "31", TAMANHO_PADRAO

    with st.sidebar.expander("Arquivos carregados"):
        mapa = (df.groupby(["ARQUIVO", "UF"]).size().reset_index(name="Rotas")
                  .rename(columns={"ARQUIVO": "Arquivo", "UF": "Estado"}))
        st.dataframe(mapa, width="stretch", hide_index=True, key="arquivos_carregados")

    st.sidebar.markdown("## Filtros")
    ufs = ["TODOS"] + sorted(df["UF"].unique())
    uf = st.sidebar.selectbox(
        "Estado", ufs,
        format_func=lambda c: ("Todos os estados" if c == "TODOS"
                               else f"{c} — {ESTADOS.get(c, 'Não identificado')}"),
    )
    if uf != "TODOS":
        df = df[df["UF"] == uf]
    df.attrs["selecao"] = uf

    # Sem filtro de status nem de tipo de veículo: todas as rotas entram nos gráficos.

    if not df.empty:
        d_min, d_max = df["DATA"].min().date(), df["DATA"].max().date()
        inicio_padrao, fim_padrao = periodo_do_mes_atual(df)
        periodo = st.sidebar.date_input("Período", value=(inicio_padrao, fim_padrao),
                                        min_value=d_min, max_value=d_max, format="DD/MM/YYYY")
        if isinstance(periodo, tuple) and len(periodo) == 2:
            df = df[(df["DATA"].dt.date >= periodo[0]) & (df["DATA"].dt.date <= periodo[1])]

    st.sidebar.markdown("## Drop size")
    base_drop = st.sidebar.radio(
        "Calcular o drop por", ["Parada", "Veículo", "Rota"], index=0,
        help="Drop = peso entregue dividido pela base escolhida.",
    )

    st.sidebar.markdown("## Apresentação")
    dias_slide = st.sidebar.selectbox(
        "Colunas por slide", ["10", "15", "20", "31", "Todos"], index=3,
        help="Como no slide impresso: cada tela mostra um bloco de dias.",
    )

    altura = st.sidebar.slider(
        "Altura dos painéis (% da tela)", min_value=60, max_value=140, value=100, step=5,
        help="100% faz os cinco painéis preencherem exatamente a altura da janela.",
    )
    modo_painel = renderizar_painel is not None and st.sidebar.toggle(
        "Modo painel (arrastar e redimensionar)", value=False,
        help="Monta os cinco gráficos numa grade livre, como no canvas do Power BI. "
             "O layout fica salvo neste navegador.",
    )
    return df, base_drop, dias_slide, {"altura": altura, "modo_painel": modo_painel}


@st.cache_data(show_spinner=False)
def ajustar_graficos_ao_redimensionar() -> None:
    """
    Redesenha os gráficos quando a largura da página muda.

    Abrir ou fechar a barra lateral muda a largura do conteúdo, mas não dispara
    o evento de redimensionamento da janela — por isso o Plotly continuava com a
    largura antiga e a linha aparecia esticada até a segunda interação. Aqui um
    observador acompanha o tamanho real do conteúdo e avisa o Plotly.
    """
    import streamlit.components.v1 as componentes

    componentes.html(
        """
        <script>
        (function () {
          const pai = window.parent;
          if (!pai || !pai.document) { return; }

          let ultimo = 0;
          function avisar() {
            const largura = pai.document.body.clientWidth;
            if (Math.abs(largura - ultimo) < 2) { return; }
            ultimo = largura;
            pai.dispatchEvent(new Event("resize"));
            // o Plotly redimensiona cada gráfico já desenhado na página
            const graficos = pai.document.querySelectorAll(".js-plotly-plot");
            if (pai.Plotly) {
              graficos.forEach((g) => { try { pai.Plotly.Plots.resize(g); } catch (e) {} });
            }
          }

          try {
            new pai.ResizeObserver(() => { avisar(); setTimeout(avisar, 260); })
              .observe(pai.document.body);
          } catch (e) {
            setInterval(avisar, 500);   // navegador antigo: verifica de tempos em tempos
          }
        })();
        </script>
        """,
        height=0,
    )


def selo_da_marca() -> str:
    """Selo redondo com o 'D' da marca, fixo no canto superior direito."""
    arquivo = ARQUIVO_ICONE if ARQUIVO_ICONE.exists() else ARQUIVO_LOGO
    if not arquivo.exists():
        return ""
    dados = base64.b64encode(arquivo.read_bytes()).decode()
    return (
        '<img src="data:image/png;base64,' + dados + '" alt="Dellys" '
        'style="position:fixed; top:52px; right:16px; z-index:1000; '
        'width:40px; height:40px; border-radius:12px; object-fit:cover; '
        'background:#FFFFFF; border:1px solid #E3E7EB; padding:2px; '
        'box-shadow:0 1px 2px rgba(20,22,26,.06); pointer-events:none;">'
    )


def cabecalho(uf: str, resumo: pd.DataFrame, por: str = "Dia") -> None:
    """Faixa compacta: identificação à esquerda, três indicadores à direita."""
    if resumo.empty:
        periodo, dias, rotas = "—", "0", "0"
    else:
        periodo = f"{resumo['ROTULO'].iloc[0]} – {resumo['ROTULO'].iloc[-1]}"
        dias = num(len(resumo))
        rotas = num(resumo["ROTAS"].sum())

    titulo = "Indicadores por semana" if por == "Semana" else "Indicadores por dia"
    unidade = "Semanas" if por == "Semana" else "Dias"

    st.markdown(
        f"""
        <div class="cab">
          <div class="cab-id">
            <span class="cab-barra"></span>
            <div>
              <p class="cab-titulo">{titulo}</p>
              <div class="cab-sub">{"Consolidado · Todos os estados" if uf == "TODOS"
                 else f"{uf} · {ESTADOS.get(uf, 'Estado não identificado')}"}
                <span class="cab-sep">·</span> Delly's Food Service</div>
            </div>
          </div>
          <div class="cab-meta">
            <div class="kpi-chip"><div class="kpi-rot">Período</div>
              <div class="kpi-val">{periodo}</div></div>
            <div class="kpi-chip"><div class="kpi-rot">{unidade}</div>
              <div class="kpi-val">{dias}</div></div>
            <div class="kpi-chip destaque"><div class="kpi-rot">Rotas</div>
              <div class="kpi-val">{rotas}</div></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _escolher_semana(semana: str) -> None:
    """Callback dos cards de semana: roda antes do redesenho da página."""
    atual = st.session_state.get("semana_sel", "TODAS")
    st.session_state["semana_sel"] = "TODAS" if atual == semana else semana
    st.session_state["slide"] = 1


def cards_de_semana(df: pd.DataFrame, selecionada: str) -> None:
    """
    Cards de semana no canto inferior direito.

    Clicar numa semana deixa nos gráficos apenas os dias dela; clicar de novo
    (ou em "Período completo") volta ao período inteiro. As semanas já vêm
    separadas por mês no rótulo (JUL/S1, AGO/S1...).
    """
    resumo = (df.groupby("SEMANA")
                .agg(INICIO=("DATA", "min"), FIM=("DATA", "max"),
                     DIAS=("DATA", "nunique"), ROTAS=("ROTA", "count"))
                .sort_values("INICIO"))
    if resumo.empty:
        return

    semanas = list(resumo.index)

    # fila encostada à direita: a primeira coluna serve só de espaço
    largura = 0.62
    colunas = st.columns([max(0.5, 12 - (len(semanas) + 1) * largura)]
                         + [largura] * (len(semanas) + 1), gap="small")

    with colunas[1]:
        st.button("Tudo", key="semana_todas", width="stretch",
                  help="Volta ao período completo",
                  type="primary" if selecionada == "TODAS" else "secondary",
                  on_click=lambda: st.session_state.update(semana_sel="TODAS", slide=1))

    for coluna, semana in zip(colunas[2:], semanas):
        linha = resumo.loc[semana]
        periodo = f"{linha['INICIO'].strftime('%d/%m')}–{linha['FIM'].strftime('%d/%m')}"
        with coluna:
            st.button(semana, key=f"semana_{semana}", width="stretch",
                      help=f"{periodo} · {num(linha['ROTAS'])} rotas",
                      type="primary" if semana == selecionada else "secondary",
                      on_click=_escolher_semana, args=(semana,))


def _mudar_slide(passo: int, total: int) -> None:
    """Callback dos botões: roda antes do redesenho, então nunca fica travado."""
    atual = int(st.session_state.get("slide", 1))
    st.session_state["slide"] = min(max(1, atual + passo), total)


def navegar_slides(resumo: pd.DataFrame, dias_slide: str) -> pd.DataFrame:
    """Divide o período em blocos de dias, como as páginas de um slide."""
    if dias_slide == "Todos" or len(resumo) <= int(dias_slide):
        st.session_state["slide"] = 1
        return resumo

    tamanho = int(dias_slide)
    total = -(-len(resumo) // tamanho)
    atual = min(max(1, int(st.session_state.get("slide", 1))), total)
    st.session_state["slide"] = atual

    esq, meio, dir_, _ = st.columns([0.07, 0.12, 0.07, 0.74])
    esq.button("‹", key="slide_ant", width="stretch", disabled=atual == 1,
               on_click=_mudar_slide, args=(-1, total))
    dir_.button("›", key="slide_prox", width="stretch", disabled=atual == total,
                on_click=_mudar_slide, args=(1, total))
    meio.markdown(f'<div class="paginacao">{atual} / {total}</div>', unsafe_allow_html=True)

    return resumo.iloc[(atual - 1) * tamanho: atual * tamanho].reset_index(drop=True)


def linha_um(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str,
             altura: int = 215, base_estados: pd.DataFrame | None = None) -> None:
    """Bloco operacional: Veículos e Ocupação lado a lado, Drop como card de performance."""
    pesos = [0.95, 0.95, 1.20]
    c1, c2, c3 = st.columns(pesos, gap="small")
    fr1, fr2, fr3 = [0.97 * peso / sum(pesos) for peso in pesos]

    # No consolidado, o maior e o menor do resumo passam a ser estados
    def ranking(coluna: str) -> pd.Series | None:
        if base_estados is None or base_estados["UF"].nunique() < 2:
            return None
        return ranking_estados(base_estados, coluna)

    with c1, st.container(border=True):
        titulo_painel("Veículos", "rotas / dia", linha=1)
        mini_indicadores(resumo, "VEICULOS", por_estado=ranking("VEICULOS"))
        faixa_numeros([num(v) for v in resumo["ROTAS"]], cor="suave", fracao=fr1)
        st.plotly_chart(grafico_barras(resumo, "VEICULOS", altura=altura, fracao=fr1),
                        width="stretch", config=CONFIG_GRAFICO, key="g_veiculos")
        faixa_eixo(list(resumo["EIXO"]), fracao=fr1)

    with c2, st.container(border=True):
        titulo_painel("Ocupação", "% · peso ÷ capacidade", linha=1)
        mini_indicadores(resumo, "OCUPACAO_PCT", sufixo="%",
                         por_estado=ranking("OCUPACAO_PCT"))
        faixa_numeros([num(v * 100) if pd.notna(v) else "—" for v in resumo["OCUPACAO"]],
                      cor="suave", fracao=fr2)
        st.plotly_chart(grafico_barras(resumo, "OCUPACAO", altura=altura, fracao=fr2,
                                       formato="%{y:.0%}"),
                        width="stretch", config=CONFIG_GRAFICO, key="g_ocupacao")
        faixa_eixo(list(resumo["EIXO"]), fracao=fr2)

    with c3, st.container(border=True):
        titulo_painel("Drop", rotulo_drop, linha=1)
        mini_indicadores(resumo, coluna_drop, rotulos=("Média", "Maior", "Menor"),
                         por_estado=ranking(coluna_drop))
        faixa_numeros([num(v) for v in resumo[coluna_drop]], cor="suave", fracao=fr3)
        st.plotly_chart(grafico_drop_por_dia(resumo, coluna_drop, altura=altura, fracao=fr3),
                        width="stretch", config=CONFIG_GRAFICO, key="g_drop_dia")
        faixa_eixo(list(resumo["EIXO"]), fracao=fr3)


def linha_dois(resumo: pd.DataFrame, altura: int = 345) -> None:
    """Bloco analítico: os dois gráficos de maior área da tela."""
    c1, c2 = st.columns(2, gap="small")
    meia_tela = 0.485

    with c1, st.container(border=True):
        titulo_painel("Paradas × entregas", linha=2,
                      series=[("Entregas", COR_AZUL), ("Média de paradas", COR_ESCURA)])
        faixa_numeros([num(v) for v in resumo["ENTREGAS"]], cor="azul", fracao=meia_tela)
        faixa_numeros([num(v) for v in resumo["MEDIA_PARADAS"]], cor="suave", fracao=meia_tela)
        st.plotly_chart(
            grafico_duas_linhas(resumo, "MEDIA_PARADAS", "ENTREGAS", "Média de paradas", "Entregas",
                                altura=altura, fracao=meia_tela),
            width="stretch", config=CONFIG_GRAFICO, key="g_paradas",
        )
        faixa_eixo(list(resumo["EIXO"]), fracao=meia_tela)

    with c2, st.container(border=True):
        titulo_painel("Peso × capacidade por dia", linha=2,
                      series=[("Capacidade (t)", COR_AZUL), ("Peso (t)", COR_ESCURA)])
        faixa_numeros([toneladas(v) for v in resumo["CAPACIDADE"]], cor="azul", fracao=meia_tela)
        faixa_numeros([toneladas(v) for v in resumo["PESO"]], cor="suave", fracao=meia_tela)
        st.plotly_chart(
            grafico_duas_linhas(resumo, "PESO", "CAPACIDADE", "Peso (kg)", "Capacidade (kg)",
                                altura=altura, fracao=meia_tela),
            width="stretch", config=CONFIG_GRAFICO, key="g_peso",
        )
        faixa_eixo(list(resumo["EIXO"]), fracao=meia_tela)


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



# ──────────────────────────────────────────────────────────────────────────────
# VISÃO POR SEMANA
# ──────────────────────────────────────────────────────────────────────────────

def cores_das_semanas(quantidade: int) -> list[str]:
    """Semanas alternam entre o azul cheio e o azul claro."""
    return [COR_AZUL if i % 2 == 0 else COR_AZUL_CLARO for i in range(quantidade)]


def legenda_semanas(resumo: pd.DataFrame) -> None:
    """Faixa do topo: quadradinho da cor, sigla, período e quantidade de dias."""
    cores = cores_das_semanas(len(resumo))
    itens = []
    for cor, (_, linha) in zip(cores, resumo.iterrows()):
        periodo = f"{linha['INICIO'].strftime('%d/%m')}–{linha['FIM'].strftime('%d/%m')}"
        itens.append(
            f'<span class="leg-item"><span class="leg-cor" style="background:{cor}"></span>'
            f'{linha["ROTULO"]} · {periodo} · {num(linha["DIAS"])} dias</span>'
        )
    st.markdown(f'<div class="legenda">{"".join(itens)}</div>', unsafe_allow_html=True)


def grafico_semanal(resumo: pd.DataFrame, coluna: str, altura: int = 190,
                    formato: str = "%{y:,.0f}") -> go.Figure:
    fig = go.Figure(
        go.Bar(
            x=resumo["ROTULO"], y=resumo[coluna], width=0.45,
            marker_color=cores_das_semanas(len(resumo)), marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>" + formato + "<extra></extra>",
        )
    )
    base = {k: v for k, v in LAYOUT_BASE.items() if k != "margin"}
    fig.update_layout(**base, height=altura, bargap=0.5, autosize=True,
                      margin=dict(l=0, r=0, t=4, b=6))
    fig.update_xaxes(**EIXO_X, showticklabels=False, type="category",
                     range=[-0.5, len(resumo) - 0.5])
    fig.update_yaxes(**EIXO_Y, showticklabels=False)
    return fig


def valores_semanais(resumo: pd.DataFrame, valores: list[str], rodapes: list[str],
                     compacta: bool = False) -> None:
    """Números grandes embaixo de cada barra, com a sigla da semana."""
    classe = "sem-cel compacta" if compacta else "sem-cel"
    celulas = "".join(
        f'<div class="{classe}"><div class="sem-val">{valor}</div>'
        f'<div class="sem-rot">{rodape}</div></div>'
        for valor, rodape in zip(valores, rodapes)
    )
    st.markdown(
        f'<div class="sem-faixa" '
        f'style="grid-template-columns: repeat({len(valores)}, minmax(0, 1fr))">'
        f'{celulas}</div>',
        unsafe_allow_html=True,
    )


def painel_semanal(resumo: pd.DataFrame, coluna: str, titulo: str, nota: str,
                   valores: list[str], rodapes: list[str], altura: int, chave: str,
                   compacta: bool = False, formato: str = "%{y:,.0f}") -> None:
    with st.container(border=True):
        titulo_painel(titulo, nota)
        st.plotly_chart(grafico_semanal(resumo, coluna, altura=altura, formato=formato),
                        width="stretch",
                        config={"displayModeBar": False}, key=f"sem_{chave}")
        valores_semanais(resumo, valores, rodapes, compacta=compacta)


def visao_semanal(resumo: pd.DataFrame, coluna_drop: str, rotulo_drop: str,
                  altura: int = 190) -> None:
    """Os seis indicadores da semana, em dois blocos de três painéis."""
    siglas = list(resumo["ROTULO"])
    legenda_semanas(resumo)

    c1, c2, c3 = st.columns(3, gap="small")
    with c1:
        painel_semanal(resumo, "VEICULOS", "Veículos", "rotas",
                       [num(v) for v in resumo["VEICULOS"]], siglas, altura, "veiculos")
    with c2:
        painel_semanal(resumo, "OCUPACAO", "Ocupação", "peso ÷ capacidade",
                       [f"{num(v * 100)}%" for v in resumo["OCUPACAO"]], siglas, altura, "ocupacao",
                       formato="%{y:.0%}")
    with c3:
        painel_semanal(resumo, coluna_drop, "Drop", rotulo_drop,
                       [num(v) for v in resumo[coluna_drop]], siglas, altura, "drop")

    c4, c5, c6 = st.columns(3, gap="small")
    with c4:
        painel_semanal(resumo, "MEDIA_PARADAS", "Média de paradas", "por rota",
                       [num(v) for v in resumo["MEDIA_PARADAS"]], siglas, altura, "paradas")
    with c5:
        painel_semanal(resumo, "ENTREGAS", "Entregas", "ordens",
                       [num(v) for v in resumo["ENTREGAS"]], siglas, altura, "entregas",
                       compacta=len(resumo) > 5)
    with c6:
        rodapes_peso = [f"{s} · cap {num(c / 1000, 1)} t" for s, c in zip(siglas, resumo["CAPACIDADE"])]
        painel_semanal(resumo, "PESO", "Peso", "toneladas · capacidade abaixo",
                       [f"{num(v / 1000, 1)} t" for v in resumo["PESO"]],
                       rodapes_peso, altura, "peso", compacta=True)


def main() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(selo_da_marca(), unsafe_allow_html=True)
    ajustar_graficos_ao_redimensionar()

    df, base_drop, dias_slide, tamanho = barra_lateral()
    proporcao = tamanho["altura"] / 100
    # Alturas pensadas para os cinco painéis caberem numa tela Full HD a 100%.
    altura_cima = int(215 * proporcao)   # bloco operacional (tem o resumo interno)
    altura_baixo = int(345 * proporcao)  # bloco analítico: os maiores da tela

    if df.empty:
        st.markdown(
            '<div class="abertura">'
            '<p class="cab-titulo">Dados e fatos</p>'
            '<div class="cab-sub">Indicadores de roteirização · Delly\'s Food Service</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        st.info(
            "Nenhuma planilha encontrada no repositório. Coloque os relatórios de rotas do "
            "RoadNet na pasta **dados/** (ex.: `dados/AM.xlsx`, `dados/MG.xlsx`) — o código "
            "do estado no nome do arquivo é o que identifica cada operação."
        )
        return

    selecao = df.attrs.get("selecao") or df["UF"].iloc[0]
    coluna_drop = {"Parada": "DROP_PARADA", "Veículo": "DROP_VEICULO", "Rota": "DROP_ROTA"}[base_drop]
    rotulo_drop = {"Parada": "kg / parada", "Veículo": "kg / veículo", "Rota": "kg / rota"}[base_drop]

    aba_dia, aba_semana = st.tabs(["Por dia", "Por semana"])

    with aba_dia:
        semana_sel = st.session_state.get("semana_sel", "TODAS")
        semanas_validas = set(df["SEMANA"].unique())
        if semana_sel not in semanas_validas:
            semana_sel = "TODAS"
        df_dia = df if semana_sel == "TODAS" else df[df["SEMANA"] == semana_sel]

        resumo = indicadores_por_dia(df_dia, por="Dia")
        if resumo.empty:
            st.warning("Nenhuma rota no período selecionado.")
        else:
            pagina = navegar_slides(resumo, dias_slide)
            cabecalho(selecao, pagina, por="Dia")
            cards_de_semana(df, semana_sel)
            if tamanho["modo_painel"] and renderizar_painel is None:
                st.warning(
                    "O modo painel precisa do arquivo **painel_arrastavel.py** na mesma "
                    "pasta do app.py, no repositório. Mostrando o modo apresentação."
                )
            if tamanho["modo_painel"] and renderizar_painel is not None:
                renderizar_painel(
                    figuras_do_painel(pagina, coluna_drop, rotulo_drop),
                    chave=selecao,
                    altura_celula=70,
                    altura_total=int(700 * proporcao),
                )
            else:
                linha_um(pagina, coluna_drop, rotulo_drop, altura=altura_cima,
                         base_estados=df_dia if selecao == "TODOS" else None)
                linha_dois(pagina, altura=altura_baixo)

    with aba_semana:
        semanal = indicadores_por_dia(df, por="Semana")
        if semanal.empty:
            st.warning("Nenhuma rota no período selecionado.")
        else:
            cabecalho(selecao, semanal, por="Semana")
            visao_semanal(semanal, coluna_drop, rotulo_drop,
                          altura=int(220 * proporcao))


if __name__ == "__main__":
    main()
