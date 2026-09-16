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
import sys
import unicodedata
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).parent
PASTA_DADOS = RAIZ / "dados"
PASTA_PUBLICA = RAIZ / "public"
PASTA_DETALHE = PASTA_PUBLICA / "detalhe"
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

# Nomes que o RoadNet costuma dar à duração da rota. O primeiro que existir na
# planilha é usado; a comparação ignora acento, caixa e espaços.
COLUNAS_TEMPO = [
    "Tempo Total de Operação Planejado",   # nome exato no relatório da Delly's
    "Tempo total de operação",
    "Tempo total de operacao",
    "Tempo total",
    "Tempo Total",
    "Tempo total da rota",
    "Tempo de rota",
    "Tempo planejado",
    "Duração total",
    "Duração",
    "Tempo total de viagem",
    "Tempo total planejado",
]

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
] + COLUNAS_TEMPO

EXTENSOES = {".xlsx", ".xlsm", ".xls", ".csv"}

# São Paulo opera em bases distintas, identificadas pelo começo do ID da rota.
# A ordem importa: "3P" antes de "SP" evita que um ID como "3P12" caia no lugar
# errado, e a lista é percorrida do prefixo mais longo para o mais curto.
UNIDADES_SP = {
    "3P": "3P — Três Passos",
    "BX": "BX — Baixada",
    "HT": "HT — Hortolândia",
    "IT": "IT — Itapeva",
    "JC": "JC — Jacareí",
    "SP": "SP — Capital",
}


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


def horas(valor) -> float:
    """
    Duração da rota em horas, aceitando os formatos que o RoadNet exporta.

    "8:30:00" e "8:30" viram 8,5. Número puro é interpretado como horas quando
    é pequeno e como minutos quando passa de 24 — um valor como 510 só pode ser
    minutos, já que nenhuma rota dura 510 horas.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return float("nan")

    if isinstance(valor, pd.Timedelta):
        return valor.total_seconds() / 3600

    texto = str(valor).strip()
    if not texto or texto in {"-", "--", "nan", "NaT"}:
        return float("nan")

    # rotas que viram o dia saem como "1 day, 9:15:00" ou "2 days, 3:00:00"
    dias = 0.0
    if "day" in texto:
        parte_dias, _, resto = texto.partition(",")
        numero_dias = re.search(r"(\d+)", parte_dias)
        if numero_dias:
            dias = float(numero_dias.group(1)) * 24
        texto = resto.strip() or "0:00"

    if ":" in texto:
        partes = texto.split(":")
        try:
            numeros = [float(parte.replace(",", ".")) for parte in partes]
        except ValueError:
            return float("nan")
        while len(numeros) < 3:
            numeros.append(0.0)
        h, m, s = numeros[0], numeros[1], numeros[2]
        # a planilha mistura "00:54" (hora:minuto) e "00:58:00" (com segundos);
        # os dois caem aqui certos, porque o campo que falta entra como zero
        return dias + h + m / 60 + s / 3600

    numero = br_para_float(texto)
    if pd.isna(numero):
        return float("nan")
    return dias + (numero / 60 if numero > 24 else numero)


def coluna_de_tempo(df: pd.DataFrame) -> str | None:
    """Primeira coluna de duração encontrada, comparando sem acento nem caixa."""
    mapa = {normalizar(c): c for c in df.columns}
    for alvo in COLUNAS_TEMPO:
        achado = mapa.get(normalizar(alvo))
        if achado is not None:
            return achado
    return None


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


def unidade_da_rota(uf: str, rota: str) -> str:
    """
    Base de São Paulo a partir do prefixo do ID da rota.

    Só vale para SP; nos demais estados devolve vazio, e o site continua
    trabalhando com o estado inteiro. IDs que não começam por nenhum dos
    prefixos conhecidos entram como OUTROS, para nenhum número se perder.
    """
    if uf != "SP":
        return ""
    texto = re.sub(r"[^A-Z0-9]", "", str(rota).upper())
    for prefixo in sorted(UNIDADES_SP, key=len, reverse=True):
        if texto.startswith(prefixo):
            return prefixo
    return "OUTROS"


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

    coluna_tempo = coluna_de_tempo(df)
    df["HORAS"] = df[coluna_tempo].map(horas) if coluna_tempo else float("nan")
    df.attrs["coluna_tempo"] = coluna_tempo

    df["DATA"] = df["Sessão de roteirização"].map(extrair_data)
    df["ROTA"] = df["ID"].astype(str) if "ID" in df.columns else ""
    df["PLACA"] = df["Equipamento"].astype(str).str.strip() if "Equipamento" in df.columns else ""
    # O destino da rota vem no nome dela ("AM-ZLESTE 01/07", "MG-ARAXA"); a data
    # no fim é redundante com a coluna DATA e sai para o texto não ficar longo.
    df["DESTINO"] = (
        df["Descrição"].astype(str).str.strip()
          .str.replace(r"\s+\d{2}/\d{2}(/\d{2,4})?$", "", regex=True)
          .replace({"": "—", "nan": "—"})
        if "Descrição" in df.columns else "—"
    )
    df["TIPO_VEICULO"] = (
        df["Tipos de equipamento"].astype(str).str.strip()
          .replace({"": "—", "nan": "—"})
        if "Tipos de equipamento" in df.columns else "—"
    )
    df["UF"] = detectar_estado(nome_arquivo, df)
    # A mesma placa em estados diferentes é outra frota: o estado entra na
    # chave para a contagem não juntar veículos distintos.
    df["VEICULO"] = df["UF"] + "·" + df["PLACA"].astype(str)
    df["ARQUIVO"] = Path(nome_arquivo).name
    df["UNIDADE"] = [unidade_da_rota(uf, rota)
                     for uf, rota in zip(df["UF"], df["ROTA"])]

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

def linhas_agregadas(df: pd.DataFrame, chaves: list[str]) -> list[dict]:
    """Uma linha por combinação das chaves, com os mesmos campos de sempre."""
    agrupado = df.groupby(chaves, as_index=False).agg(
        ROTAS=("ROTA", "count"),
        VEICULOS=("VEICULO", pd.Series.nunique),
        PARADAS=("PARADAS", "sum"),
        ENTREGAS=("ENTREGAS", "sum"),
        PESO=("PESO", "sum"),
        CAPACIDADE=("CAPACIDADE", "sum"),
        VALOR=("VALOR", "sum"),
        DISTANCIA=("DISTANCIA", "sum"),
        HORAS=("HORAS", "sum"),
        ROTAS_COM_HORA=("HORAS", "count"),
        SEMANA=("SEMANA", "first"),
    )

    registros = []
    for _, linha in agrupado.sort_values(["DATA", "UF"]).iterrows():
        registro = {
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
            # horas só entram quando a planilha traz a coluna de duração
            "horas": round(float(linha["HORAS"]), 3) if linha["ROTAS_COM_HORA"] else None,
            "rotasComHora": int(linha["ROTAS_COM_HORA"]),
        }
        if "UNIDADE" in chaves:
            registro["unidade"] = linha["UNIDADE"]
        registros.append(registro)
    return registros


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

    return linhas_agregadas(df, ["UF", "DATA"])


def agregar_unidades(df: pd.DataFrame) -> list[dict]:
    """
    Segunda camada, só para os estados que têm base identificada (hoje, SP).

    Ela convive com a lista principal em vez de substituí-la: assim o número do
    estado inteiro continua saindo de uma agregação única — importante para a
    contagem de veículos, onde somar subgrupos contaria duas vezes a placa que
    rodou em duas bases no mesmo dia.
    """
    com_unidade = df[df["UNIDADE"].astype(bool)]
    if com_unidade.empty:
        return []
    return linhas_agregadas(com_unidade, ["UF", "UNIDADE", "DATA"])


def gravar_detalhe(df: pd.DataFrame) -> tuple[int, float]:
    """
    Um arquivo por dia com as cargas daquele dia, em public/detalhe/.

    O site carrega esses arquivos só quando alguém clica num dia do gráfico —
    por isso eles ficam separados do dados.json, que é lido na abertura. Cada
    linha é uma rota do RoadNet: placa, tipo de veículo, paradas, entregas,
    peso e capacidade.
    """
    if PASTA_DETALHE.exists():
        for antigo in PASTA_DETALHE.glob("*.json"):
            antigo.unlink()
    PASTA_DETALHE.mkdir(parents=True, exist_ok=True)

    def numero(valor, casas=2):
        if pd.isna(valor):
            return None
        return round(float(valor), casas)

    total_bytes = 0
    for data, parte in df.groupby("DATA"):
        cargas = []
        for _, linha in parte.iterrows():
            cargas.append({
                "uf": linha["UF"],
                "rota": str(linha["ROTA"]),
                "placa": linha["PLACA"] or "—",
                "destino": linha.get("DESTINO") or "—",
                "tipo": str(linha.get("TIPO_VEICULO") or "—"),
                "paradas": numero(linha["PARADAS"], 0),
                "entregas": numero(linha["ENTREGAS"], 0),
                "peso": numero(linha["PESO"]),
                "capacidade": numero(linha["CAPACIDADE"]),
                "distancia": numero(linha["DISTANCIA"]),
                "horas": numero(linha["HORAS"], 3),
            })
        cargas.sort(key=lambda c: (c["uf"], c["rota"]))
        arquivo = PASTA_DETALHE / f"{data.strftime('%Y-%m-%d')}.json"
        arquivo.write_text(
            json.dumps({"data": data.strftime("%Y-%m-%d"), "cargas": cargas},
                       ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
        total_bytes += arquivo.stat().st_size

    quantidade = len(list(PASTA_DETALHE.glob("*.json")))
    return quantidade, total_bytes / 1024


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
            coluna_tempo = base.attrs.get("coluna_tempo")
            tempo = f"tempo: {coluna_tempo}" if coluna_tempo else "SEM coluna de tempo"
            print(f"  lido  {nome:<28} {len(base):>6} rotas   "
                  f"estado {base['UF'].iloc[0]:<7} {tempo}")
        except Exception as exc:  # noqa: BLE001
            problemas.append(f"{nome}: {exc}")

    for aviso in problemas:
        print(f"  ERRO  {aviso}")

    if not bases:
        print("Nada foi gerado: nenhum arquivo pôde ser lido.")
        return 1

    df = pd.concat(bases, ignore_index=True)

    if not df["HORAS"].notna().any():
        print("\n  AVISO: nenhuma planilha trouxe coluna de duração da rota.")
        print("  O indicador de tempo de operação vai ficar vazio.")
        print("  Colunas encontradas no primeiro arquivo:")
        for coluna in bases[0].columns:
            if not str(coluna).isupper():
                print(f"    - {coluna}")
        print("  Acrescente o nome certo à lista COLUNAS_TEMPO, no topo deste arquivo.")

    registros = agregar(df)
    unidades = agregar_unidades(df)

    conteudo = {
        "gerado_em": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
        "estados": ESTADOS,
        "arquivos": sorted(df["ARQUIVO"].unique().tolist()),
        "rotas_processadas": int(len(df)),
        "registros": registros,
        # bases de SP: lista paralela, usada só quando o filtro de base é aberto
        "unidades": unidades,
        "nomes_unidades": {**UNIDADES_SP, "OUTROS": "Outros"},
    }

    PASTA_PUBLICA.mkdir(parents=True, exist_ok=True)
    ARQUIVO_SAIDA.write_text(
        json.dumps(conteudo, ensure_ascii=False, separators=(",", ":")), encoding="utf-8"
    )
    dias, kb_detalhe = gravar_detalhe(df)

    tamanho = ARQUIVO_SAIDA.stat().st_size / 1024
    print(f"\n{ARQUIVO_SAIDA.relative_to(RAIZ)} gravado — "
          f"{len(registros)} linhas (estado × dia), {tamanho:.0f} KB")
    print(f"public/detalhe/ — {dias} arquivos de cargas, {kb_detalhe:.0f} KB no total")
    if unidades:
        bases = sorted({r["unidade"] for r in unidades})
        print(f"Bases de SP: {', '.join(bases)} ({len(unidades)} linhas)")
    print(f"Estados: {', '.join(sorted(df['UF'].unique()))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
