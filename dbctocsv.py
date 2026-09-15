# Instalação:
# python -m pip install dbc-to-dbf pandas dbfread

from pathlib import Path

import pandas as pd
from dbctodbf import DBCDecompress
from dbfread import DBF


PASTA_ENTRADA = Path("dados_brutos")
PASTA_SAIDA = Path("dados_csv")
PASTA_DBF = PASTA_SAIDA / "dbf"

PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
PASTA_DBF.mkdir(parents=True, exist_ok=True)


def converter_dbc_para_dbf(
    caminho_dbc: Path,
    pasta_saida: Path
) -> Path:
    """
    Converte um arquivo .dbc do DATASUS para .dbf.
    """

    caminho_dbf = pasta_saida / f"{caminho_dbc.stem}.dbf"

    conversor = DBCDecompress()

    conversor.decompressFile(
        str(caminho_dbc),
        str(caminho_dbf)
    )

    print(
        f"[OK] {caminho_dbc.name} -> "
        f"{caminho_dbf.name}"
    )

    return caminho_dbf


def converter_dbf_para_csv(
    caminho_dbf: Path,
    pasta_saida: Path
) -> Path:
    """
    Lê um arquivo .dbf e salva os dados em .csv.
    """

    caminho_csv = pasta_saida / f"{caminho_dbf.stem}.csv"

    tabela = DBF(
        str(caminho_dbf),
        encoding="latin-1"
    )

    df = pd.DataFrame(iter(tabela))

    df.to_csv(
        caminho_csv,
        sep=";",
        encoding="latin-1",
        index=False
    )

    print(
        f"[OK] {caminho_dbf.name} -> "
        f"{caminho_csv.name}"
    )

    return caminho_csv


def converter_dbc_para_csv(
    caminho_dbc: Path,
    pasta_saida: Path
) -> Path:
    """
    Faz o processo completo:

    .dbc -> .dbf -> .csv
    """

    caminho_dbf = converter_dbc_para_dbf(
        caminho_dbc,
        PASTA_DBF
    )

    caminho_csv = converter_dbf_para_csv(
        caminho_dbf,
        pasta_saida
    )

    return caminho_csv


def converter_todos_os_dbc(
    pasta_entrada: Path,
    pasta_saida: Path
) -> list[Path]:

    arquivos_dbc = sorted(
        pasta_entrada.glob("*.dbc")
    )

    if not arquivos_dbc:
        print(
            f"Nenhum arquivo .dbc encontrado em "
            f"'{pasta_entrada}'."
        )
        return []

    csvs_gerados = []

    for arquivo in arquivos_dbc:

        try:
            csv_path = converter_dbc_para_csv(
                arquivo,
                pasta_saida
            )

            csvs_gerados.append(csv_path)

        except Exception as e:

            print(
                f"[ERRO] Falha ao converter "
                f"{arquivo.name}: {e}"
            )

    return csvs_gerados


def explorar_csv(
    caminho_csv: Path,
    n_linhas_amostra: int = 5
) -> None:

    df = pd.read_csv(
        caminho_csv,
        sep=";",
        encoding="latin-1",
        low_memory=False
    )

    print("=" * 70)
    print(f"Arquivo: {caminho_csv.name}")

    print(
        f"Dimensões: "
        f"{df.shape[0]} linhas x "
        f"{df.shape[1]} colunas"
    )

    print("-" * 70)

    resumo = pd.DataFrame({
        "tipo": df.dtypes,
        "pct_nulos": (
            df.isna().mean() * 100
        ).round(2),
        "n_unicos": df.nunique(),
    })

    print(resumo)

    print("-" * 70)
    print("Amostra de dados:")

    print(
        df.head(n_linhas_amostra)
    )

    print("=" * 70)
    print()


if __name__ == "__main__":

    csvs = converter_todos_os_dbc(
        PASTA_ENTRADA,
        PASTA_SAIDA
    )

    for csv_path in csvs:
        explorar_csv(csv_path)