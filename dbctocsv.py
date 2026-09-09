# pip install readdbc pandas

import os
from pathlib import Path
import pandas as pd
import readdbc

PASTA_ENTRADA = Path("dados_brutos")
PASTA_SAIDA = Path("dados_csv")

PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

def converter_dbc_para_csv(caminho_dbc: Path, pasta_saida: Path) -> Path:
    caminho_csv = pasta_saida / (caminho_dbc.stem + ".csv")
    readdbc.dbc2csv(str(caminho_dbc), str(caminho_csv))
    print(f"[OK] {caminho_dbc.name} -> {caminho_csv.name}")
    return caminho_csv


def converter_todos_os_dbc(pasta_entrada: Path, pasta_saida: Path) -> list[Path]:
    arquivos_dbc = sorted(pasta_entrada.glob("*.dbc"))

    if not arquivos_dbc:
        print(f"Nenhum arquivo .dbc encontrado em '{pasta_entrada}'.")
        return []

    csvs_gerados = []
    for arquivo in arquivos_dbc:
        try:
            csv_path = converter_dbc_para_csv(arquivo, pasta_saida)
            csvs_gerados.append(csv_path)
        except Exception as e:
            print(f"[ERRO] Falha ao converter {arquivo.name}: {e}")

    return csvs_gerados


def explorar_csv(caminho_csv: Path, n_linhas_amostra: int = 5) -> None:
    df = pd.read_csv(caminho_csv, sep=";", encoding="latin-1", low_memory=False)

    print("=" * 70)
    print(f"Arquivo: {caminho_csv.name}")
    print(f"Dimensões: {df.shape[0]} linhas x {df.shape[1]} colunas")
    print("-" * 70)

    resumo = pd.DataFrame({
        "tipo": df.dtypes,
        "pct_nulos": (df.isna().mean() * 100).round(2),
        "n_unicos": df.nunique(),
    })
    print(resumo)

    print("-" * 70)
    print("Amostra de dados:")
    print(df.head(n_linhas_amostra))
    print("=" * 70, "\n")


if __name__ == "__main__":
    csvs = converter_todos_os_dbc(PASTA_ENTRADA, PASTA_SAIDA)

    for csv_path in csvs:
        explorar_csv(csv_path)
