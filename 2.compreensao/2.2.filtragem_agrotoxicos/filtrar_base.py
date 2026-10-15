"""
Filtragem da base e geração de uma cópia filtrada (CSV)
=======================================================

Fluxo do projeto:
    dados_brutos/*.dbc --(dbc_to_csv_datasus.py)--> dados_csv/SEU_ARQUIVO.csv
    dados_csv/SEU_ARQUIVO.csv --(ESTE SCRIPT)--> dados_csv/base_filtrada.csv
    dados_csv/base_filtrada.csv --(analise_alvo_hospital.py)--> relatórios

Este script NÃO altera o CSV original: ele lê a base completa, aplica a lista
FILTROS (em ordem) e grava uma cópia filtrada. Quando descobrir um novo filtro
necessário, basta adicionar uma linha em FILTROS e rodar de novo — os demais
scripts continuam lendo sempre a mesma base_filtrada.csv.

A cópia preserva o conteúdo original das células (tudo é lido como texto), então
códigos com zero à esquerda e valores como "2" não viram "2.0".

Ao final, imprime no console um relatório com quantas linhas cada filtro removeu
e quantas restaram.

Pré-requisitos:
    pip install pandas
"""

from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------
# CONFIGURAÇÃO
# ----------------------------------------------------------------------
CAMINHO_ENTRADA = Path("../../1.conversao_dbc/dados_csv/IEXOBR26.csv")
CAMINHO_SAIDA = Path("dados_csv/base_filtrada.csv")     # cópia filtrada
SEPARADOR = ";"
ENCODING = "latin-1"

# Lista de filtros, aplicados em sequência (um depois do outro).
# Cada filtro: {"coluna": ..., "operador": ..., "valores": ...}
#
# Operadores disponíveis:
#   "in"           valores = [2, 3, 4]            mantém só esses valores
#   "not_in"       valores = [9]                  remove esses valores
#   "==" / "!="    valores = 1                    igual / diferente
#   ">" ">=" "<" "<="   valores = 2020            comparação numérica
#   "entre"        valores = [2020, 2023]         numérico, extremos incluídos
#   "entre_datas"  valores = ["2021-01-01", "2023-12-31"]   datas, extremos incluídos
#   "notna" / "isna"     (sem valores)            mantém preenchidos / vazios
#
# Nulos: "not_in" e "!=" MANTÊM as linhas vazias (só removem o valor citado);
# todos os outros operadores (exceto "isna") removem as linhas vazias.
FILTROS = [
    {"coluna": "AGENTE_TOX", "operador": "in", "valores": [2, 3, 4]},

    # Exemplos para quando precisar de novos filtros (descomente/ajuste):
    # {"coluna": "HOSPITAL", "operador": "in", "valores": [1, 2]},
    # {"coluna": "NU_ANO", "operador": ">=", "valores": 2020},
    # {"coluna": "DT_NOTIFIC", "operador": "entre_datas", "valores": ["2021-01-01", "2023-12-31"]},
]


# ----------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ----------------------------------------------------------------------
def _fmt(n: int) -> str:
    """12345 -> '12.345' (padrão brasileiro)."""
    return f"{n:,}".replace(",", ".")


def _pct(parte: float, total: float) -> str:
    p = parte / total * 100 if total else 0.0
    return f"{p:.1f}%".replace(".", ",")


def _descricao(filtro: dict) -> str:
    valores = filtro.get("valores")
    texto = f"{filtro['coluna']} {filtro['operador']}"
    return texto if valores is None else f"{texto} {valores}"


def _so_numeros(valores) -> bool:
    lista = valores if isinstance(valores, (list, tuple, set)) else [valores]
    return all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in lista)


def construir_mascara(df: pd.DataFrame, filtro: dict) -> pd.Series:
    """Devolve uma máscara booleana (True = linha mantida) para um filtro."""
    col, op, valores = filtro["coluna"], filtro["operador"], filtro.get("valores")

    if col not in df.columns:
        raise KeyError(f"Coluna '{col}' não encontrada. Colunas disponíveis: {list(df.columns)}")
    serie = df[col]

    if op == "notna":
        return serie.notna()
    if op == "isna":
        return serie.isna()

    if op == "entre_datas":
        datas = pd.to_datetime(serie, errors="coerce")
        return datas.between(pd.to_datetime(valores[0]), pd.to_datetime(valores[1])).fillna(False)

    # Comparação numérica quando os valores do filtro são números; texto caso contrário.
    if _so_numeros(valores):
        comparavel = pd.to_numeric(serie, errors="coerce")
    else:
        comparavel = serie.astype("string").str.strip()

    if op == "in":
        mascara = comparavel.isin(valores)
    elif op == "not_in":
        mascara = ~comparavel.isin(valores)
    elif op == "==":
        mascara = comparavel == valores
    elif op == "!=":
        mascara = (comparavel != valores) | comparavel.isna()
    elif op == ">":
        mascara = comparavel > valores
    elif op == ">=":
        mascara = comparavel >= valores
    elif op == "<":
        mascara = comparavel < valores
    elif op == "<=":
        mascara = comparavel <= valores
    elif op == "entre":
        mascara = comparavel.between(valores[0], valores[1])
    else:
        raise ValueError(f"Operador desconhecido: '{op}'")

    return mascara.fillna(False).astype(bool)


def aplicar_filtros(df: pd.DataFrame, filtros: list) -> pd.DataFrame:
    """Aplica os filtros em sequência, imprimindo quantas linhas cada um removeu."""
    n_original = len(df)
    atual = df

    for i, filtro in enumerate(filtros, start=1):
        antes = len(atual)
        atual = atual[construir_mascara(atual, filtro)]
        depois = len(atual)

        print(f"Filtro {i}: {_descricao(filtro)}")
        print(f"   linhas antes: {_fmt(antes)} | mantidas: {_fmt(depois)} "
              f"| removidas: {_fmt(antes - depois)} ({_pct(antes - depois, antes)} do que havia)")
        print(f"   restante em relação à base original: {_pct(depois, n_original)}")
        print()

    return atual


def imprimir_resumo_final(df_original: pd.DataFrame, df_final: pd.DataFrame, filtros: list) -> None:
    n_orig, n_final = len(df_original), len(df_final)
    print("-" * 70)
    print("RESULTADO FINAL")
    print(f"   Linhas na base original : {_fmt(n_orig)}")
    print(f"   Linhas na base filtrada : {_fmt(n_final)} ({_pct(n_final, n_orig)} da original)")
    print(f"   Linhas removidas no total: {_fmt(n_orig - n_final)} ({_pct(n_orig - n_final, n_orig)})")
    print(f"   Colunas: {df_final.shape[1]}")

    colunas_filtradas = list(dict.fromkeys(f["coluna"] for f in filtros))
    for col in colunas_filtradas:
        print(f"\n   Distribuição de {col} na base filtrada:")
        contagem = df_final[col].value_counts(dropna=False).head(10)
        for valor, n in contagem.items():
            rotulo = "(vazio)" if pd.isna(valor) else valor
            print(f"      {rotulo}: {_fmt(n)} ({_pct(n, n_final)})")
    print("-" * 70)


# ----------------------------------------------------------------------
# EXECUÇÃO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    if CAMINHO_SAIDA.resolve() == CAMINHO_ENTRADA.resolve():
        raise ValueError("CAMINHO_SAIDA não pode ser igual a CAMINHO_ENTRADA (o original seria sobrescrito).")

    # Tudo como texto, e só célula vazia vira nulo: a cópia fica fiel ao original.
    df = pd.read_csv(CAMINHO_ENTRADA, sep=SEPARADOR, encoding=ENCODING,
                     dtype=str, keep_default_na=False, na_values=[""], low_memory=False)

    print("=" * 70)
    print("FILTRAGEM DA BASE")
    print("=" * 70)
    print(f"Entrada: {CAMINHO_ENTRADA}")
    print(f"Linhas: {_fmt(len(df))} | Colunas: {df.shape[1]}")
    print(f"Filtros a aplicar: {len(FILTROS)}\n")

    df_filtrado = aplicar_filtros(df, FILTROS)

    if df_filtrado.empty:
        print("ATENÇÃO: nenhuma linha restou após os filtros. Nada foi salvo — revise FILTROS.")
    else:
        CAMINHO_SAIDA.parent.mkdir(parents=True, exist_ok=True)
        df_filtrado.to_csv(CAMINHO_SAIDA, sep=SEPARADOR, encoding=ENCODING, index=False)

        imprimir_resumo_final(df, df_filtrado, FILTROS)
        tamanho_mb = CAMINHO_SAIDA.stat().st_size / 1024 / 1024
        print(f"[OK] Base filtrada salva em: {CAMINHO_SAIDA} ({tamanho_mb:.1f} MB)")