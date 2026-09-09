"""
DADOS E FATOS — geração da base do site

Lê os relatórios de rotas do RoadNet da pasta dados/ (xlsx, xls, csv), aplica o
mesmo tratamento que o app do Streamlit fazia e grava public/dados.json, que é
o arquivo que o site estático consome.

Rode antes de cada publicação:

    python build_dados.py

O JSON guarda uma linha por estado e por dia, já agregada. Os indicadores
derivados (ocupação, drop, média de paradas) são calculados no navegador, então
trocar o filtro não exige gerar nada de novo.

Por que agregar aqui e não no site: o relatório bruto tem uma linha por rota e
pesa dezenas de MB; agregado por dia ele cabe em poucas centenas de KB e o site
abre instantaneamente.
"""

from __future__ import annotations

import io
import json
import re
import shutil
import sys
import unicodedata
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).parent
PASTA_DADOS = RAIZ / "dados"
PASTA_ASSETS = RAIZ / "assets"
PASTA_PUBLICA = RAIZ / "public"
ARQUIVO_SAIDA = PASTA_PUBLICA / "dados.json"

# Estados atendidos. A chave é o código usado no nome do arquivo.
ESTADOS = {
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

EXTENSOES = {".xlsx", ".xlsm", ".xls", ".csv"}


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
    df["PLACA"] = df["Equipamento"].astype(str).str.strip() if "Equipamento" in df.columns else ""
    df["UF"] = detectar_estado(nome_arquivo, df)
    # A mesma placa em estados diferentes é outra frota: o estado entra na
    # chave para a contagem não juntar veículos distintos.
    df["VEICULO"] = df["UF"] + "·" + df["PLACA"].astype(str)
    df["ARQUIVO"] = Path(nome_arquivo).name

    if "SEMANA" in df.columns:
        df["SEMANA_ARQUIVO"] = df["SEMANA"].astype(str).str.strip()
    else:
        df["SEMANA_ARQUIVO"] = ""

    df = df.dropna(subset=["DATA"])
    df = df[df["CAPACIDADE"].fillna(0) > 0]
    return df.reset_index(drop=True)


def arquivos_da_pasta() -> list[tuple[str, bytes]]:
    """Planilhas da pasta dados/ e também da raiz, se alguém salvou ali."""
    achados: list[tuple[str, bytes]] = []
    vistos: set[str] = set()
    for pasta in (PASTA_DADOS, RAIZ):
        if not pasta.exists():
            continue
        for caminho in sorted(pasta.iterdir()):
            if not caminho.is_file() or caminho.suffix.lower() not in EXTENSOES:
                continue
            if caminho.name.startswith("~$") or caminho.name in vistos:
                continue
            vistos.add(caminho.name)
            achados.append((caminho.name, caminho.read_bytes()))
    return achados


# ──────────────────────────────────────────────────────────────────────────────
# AGREGAÇÃO
# ──────────────────────────────────────────────────────────────────────────────

def agregar(df: pd.DataFrame) -> list[dict]:
    """
    Uma linha por estado e por dia.

    Veículos é contagem de placas distintas dentro do dia e do estado. Como a
    chave do veículo já carrega o estado, somar essas contagens entre estados
    devolve o número certo do consolidado, e somar entre dias devolve o número
    da semana (uma placa que rodou cinco dias conta cinco vezes), exatamente
    como a apresentação sempre calculou.
    """
    varios_meses = df["DATA"].dt.month.nunique() > 1
    df = df.copy()
    calculada = df["DATA"].map(lambda d: rotulo_semana(d, varios_meses))
    df["SEMANA"] = df["SEMANA_ARQUIVO"].where(
        df["SEMANA_ARQUIVO"].astype(bool) & (df["SEMANA_ARQUIVO"] != "nan"), calculada
    )

    agrupado = df.groupby(["UF", "DATA"], as_index=False).agg(
        ROTAS=("ROTA", "count"),
        VEICULOS=("VEICULO", pd.Series.nunique),
        PARADAS=("PARADAS", "sum"),
        ENTREGAS=("ENTREGAS", "sum"),
        PESO=("PESO", "sum"),
        CAPACIDADE=("CAPACIDADE", "sum"),
        VALOR=("VALOR", "sum"),
        DISTANCIA=("DISTANCIA", "sum"),
        SEMANA=("SEMANA", "first"),
    )

    registros = []
    for _, linha in agrupado.sort_values(["DATA", "UF"]).iterrows():
        registros.append({
            "uf": linha["UF"],
            "data": linha["DATA"].strftime("%Y-%m-%d"),
            "semana": linha["SEMANA"],
            "rotas": int(linha["ROTAS"]),
            "veiculos": int(linha["VEICULOS"]),
            "paradas": round(float(linha["PARADAS"] or 0), 2),
            "entregas": round(float(linha["ENTREGAS"] or 0), 2),
            "peso": round(float(linha["PESO"] or 0), 2),
            "capacidade": round(float(linha["CAPACIDADE"] or 0), 2),
            "valor": round(float(linha["VALOR"] or 0), 2),
            "distancia": round(float(linha["DISTANCIA"] or 0), 2),
        })
    return registros


def copiar_assets() -> None:
    """Leva logo e favicon para dentro de public/, que é o que a Vercel publica."""
    if not PASTA_ASSETS.exists():
        return
    destino = PASTA_PUBLICA / "assets"
    destino.mkdir(parents=True, exist_ok=True)
    for caminho in PASTA_ASSETS.iterdir():
        if caminho.is_file():
            shutil.copy2(caminho, destino / caminho.name)


def main() -> int:
    arquivos = arquivos_da_pasta()
    if not arquivos:
        print(f"Nenhuma planilha encontrada em {PASTA_DADOS}. "
              f"Coloque os relatórios do RoadNet ali (ex.: dados/AM.xlsx).")
        return 1

    bases, problemas = [], []
    for nome, conteudo in arquivos:
        try:
            base = tratar(ler_arquivo(nome, conteudo), nome)
            if base.empty:
                problemas.append(f"{nome}: nenhuma rota válida")
                continue
            bases.append(base)
            print(f"  lido  {nome:<28} {len(base):>6} rotas   estado {base['UF'].iloc[0]}")
        except Exception as exc:  # noqa: BLE001
            problemas.append(f"{nome}: {exc}")

    for aviso in problemas:
        print(f"  ERRO  {aviso}")

    if not bases:
        print("Nada foi gerado: nenhum arquivo pôde ser lido.")
        return 1

    df = pd.concat(bases, ignore_index=True)
    registros = agregar(df)

    conteudo = {
        "gerado_em": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
        "estados": ESTADOS,
        "arquivos": sorted(df["ARQUIVO"].unique().tolist()),
        "rotas_processadas": int(len(df)),
        "registros": registros,
    }

    PASTA_PUBLICA.mkdir(parents=True, exist_ok=True)
    ARQUIVO_SAIDA.write_text(
        json.dumps(conteudo, ensure_ascii=False, separators=(",", ":")), encoding="utf-8"
    )
    copiar_assets()

    tamanho = ARQUIVO_SAIDA.stat().st_size / 1024
    print(f"\n{ARQUIVO_SAIDA.relative_to(RAIZ)} gravado — "
          f"{len(registros)} linhas (estado × dia), {tamanho:.0f} KB")
    print(f"Estados: {', '.join(sorted(df['UF'].unique()))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
