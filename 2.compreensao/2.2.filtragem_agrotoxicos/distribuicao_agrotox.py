"""
Análise aprofundada das colunas (DATASUS) para seleção de variáveis e alvo
============================================================================

Complementa o script anterior (dbc_to_csv_datasus.py). Aqui o objetivo não é
só listar tipo/nulos, mas entender a DISTRIBUIÇÃO de cada coluna — essencial
para decidir o que vira feature, o que vira alvo, e o que deve ser descartado.

Pré-requisitos:
    pip install pandas numpy openpyxl
"""

from pathlib import Path
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font

pd.set_option("display.max_rows", 200)
pd.set_option("display.width", 120)


# ----------------------------------------------------------------------
# 1. PERFIL DETALHADO DE CADA COLUNA
# ----------------------------------------------------------------------
def perfil_colunas(df: pd.DataFrame, max_categorias: int = 10) -> pd.DataFrame:
    """
    Gera uma tabela-resumo por coluna:
    - tipo de dado
    - % de nulos
    - nº de valores únicos
    - se for categórica/baixa cardinalidade: distribuição percentual dos
      valores mais frequentes
    - se for numérica: média, mediana, desvio-padrão, min, max

    'Baixa cardinalidade' = até max_categorias valores distintos, o que
    é comum em colunas codificadas do DATASUS (SEXO, CID, UF, etc.)
    """
    linhas = []
    n = len(df)

    for col in df.columns:
        serie = df[col]
        n_unicos = serie.nunique(dropna=True)
        pct_nulos = round(serie.isna().mean() * 100, 2)

        registro = {
            "coluna": col,
            "tipo": str(serie.dtype),
            "pct_nulos": pct_nulos,
            "n_unicos": n_unicos,
        }

        if pd.api.types.is_numeric_dtype(serie) and n_unicos > max_categorias:
            registro.update({
                "media": round(serie.mean(), 2),
                "mediana": serie.median(),
                "desvio_padrao": round(serie.std(), 2),
                "min": serie.min(),
                "max": serie.max(),
                "distribuicao_top_valores": None,
            })
        else:
            # Categórica (ou numérica com poucos valores, ex: código de CID)
            top = serie.value_counts(dropna=True, normalize=True).head(max_categorias)
            top_formatado = {str(k): f"{v*100:.1f}%" for k, v in top.items()}
            registro.update({
                "media": None, "mediana": None, "desvio_padrao": None,
                "min": None, "max": None,
                "distribuicao_top_valores": top_formatado,
            })

        linhas.append(registro)

    return pd.DataFrame(linhas)


# ----------------------------------------------------------------------
# 2. RELAÇÃO DE CADA VARIÁVEL COM UM ALVO CANDIDATO
# ----------------------------------------------------------------------
def relacao_com_alvo(df: pd.DataFrame, coluna_alvo: str, max_categorias: int = 10) -> None:
    """
    Para cada coluna (exceto o alvo), mostra como ela se comporta por
    classe do alvo. Ajuda a enxergar, visualmente em texto, se a variável
    parece discriminar as classes (candidata forte a feature) ou não.
    """
    if coluna_alvo not in df.columns:
        print(f"Coluna alvo '{coluna_alvo}' não encontrada no DataFrame.")
        return

    print("=" * 70)
    print(f"Distribuição do ALVO: {coluna_alvo}")
    print(df[coluna_alvo].value_counts(dropna=False))
    print(df[coluna_alvo].value_counts(normalize=True, dropna=False).round(3) * 100)
    print("=" * 70, "\n")

    for col in df.columns:
        if col == coluna_alvo:
            continue
        n_unicos = df[col].nunique(dropna=True)

        print(f"--- {col} vs {coluna_alvo} ---")
        if pd.api.types.is_numeric_dtype(df[col]) and n_unicos > max_categorias:
            print(df.groupby(coluna_alvo)[col].agg(["mean", "median", "std"]).round(2))
        else:
            tab = pd.crosstab(df[col], df[coluna_alvo], normalize="index").round(3) * 100
            print(tab.head(max_categorias))
        print()


# ----------------------------------------------------------------------
# 3. SINALIZAÇÃO DE ALERTAS ÚTEIS PARA MODELAGEM
# ----------------------------------------------------------------------
def alertas_qualidade(df: pd.DataFrame, limite_nulos: float = 40.0, limite_constante: float = 99.0) -> None:
    """
    Aponta automaticamente colunas problemáticas:
    - excesso de nulos (acima de limite_nulos %)
    - quase constantes (um único valor domina > limite_constante %) -> pouco
      poder discriminativo
    - colunas com um único valor (sempre descartáveis)
    """
    print("ALERTAS DE QUALIDADE DAS COLUNAS")
    print("-" * 70)
    for col in df.columns:
        serie = df[col]
        n_unicos = serie.nunique(dropna=True)
        pct_nulos = serie.isna().mean() * 100

        # Checagens independentes (não excludentes) — uma coluna pode acumular
        # mais de um alerta, ex: muitos nulos E quase constante nos valores restantes.
        if n_unicos <= 1:
            print(f"[CONSTANTE] '{col}': só tem um valor distinto — descartar")
        if pct_nulos > limite_nulos:
            print(f"[NULOS] '{col}': {pct_nulos:.1f}% de valores ausentes")
        if n_unicos > 1:
            freq_top = serie.value_counts(normalize=True, dropna=True).iloc[0] * 100
            if freq_top > limite_constante:
                print(f"[QUASE CONSTANTE] '{col}': {freq_top:.1f}% dos registros têm o mesmo valor")
    print("-" * 70, "\n")


# ----------------------------------------------------------------------
# 4. EXPORTAÇÃO DO PERFIL PARA PLANILHA (.xlsx)
# ----------------------------------------------------------------------
def _classificar_coluna(col: str, serie: pd.Series, max_categorias: int) -> str:
    """
    Classifica a coluna em 'data', 'numerica' (contínua) ou 'categorica'.

    Colunas de data (prefixo DT/DATA, convenção comum do DATASUS, ou já
    tipadas como datetime) são tratadas à parte: top-N de valores não faz
    sentido para uma data (cada dia tende a ser quase único), então elas
    não entram na aba de distribuição — só no perfil geral, com data
    mínima/máxima.
    """
    nome = col.upper()
    if pd.api.types.is_datetime64_any_dtype(serie) or nome.startswith("DT") or nome.startswith("DATA"):
        return "data"
    if pd.api.types.is_numeric_dtype(serie) and serie.nunique(dropna=True) > max_categorias:
        return "numerica"
    return "categorica"


def gerar_relatorio_excel(
    df: pd.DataFrame,
    caminho_saida: Path,
    max_categorias: int = 15,
    limite_nulos: float = 40.0,
    limite_constante: float = 99.0,
) -> None:
    """
    Gera uma planilha .xlsx com 3 abas:

      - "Perfil Geral": 1 linha por coluna — tipo, % de nulos, nº de
        valores únicos, estatísticas (numéricas) ou intervalo de datas
        (colunas de data).
      - "Distribuicao de Valores": uma MINI-TABELA POR COLUNA categórica
        (não uma tabela única com tudo empilhado), com uma linha "Outros
        valores" cobrindo o restante — a soma de cada tabela sempre
        fecha em 100%. Colunas de data e numéricas contínuas não entram
        aqui (não fazem sentido como top-N).
      - "Alertas de Qualidade": lista só as colunas com algum problema
        (nulos em excesso, constante, quase constante) — uma linha por
        coluna, com todos os motivos combinados, para ficar conciso.
    """
    linhas_perfil = []
    distribuicoes: dict[str, pd.DataFrame] = {}
    linhas_alertas = []

    for col in df.columns:
        serie = df[col]
        n_unicos = serie.nunique(dropna=True)
        pct_nulos = round(serie.isna().mean() * 100, 2)
        categoria = _classificar_coluna(col, serie, max_categorias)

        registro = {
            "coluna": col, "tipo": str(serie.dtype), "categoria": categoria,
            "pct_nulos": pct_nulos, "n_unicos": n_unicos,
            "media": None, "mediana": None, "desvio_padrao": None, "min": None, "max": None,
        }

        if categoria == "numerica":
            registro.update({
                "media": round(serie.mean(), 2),
                "mediana": serie.median(),
                "desvio_padrao": round(serie.std(), 2),
                "min": serie.min(),
                "max": serie.max(),
            })
        elif categoria == "data":
            datas = pd.to_datetime(serie, errors="coerce", dayfirst=True)
            if datas.notna().any():
                registro["min"] = datas.min().date().isoformat()
                registro["max"] = datas.max().date().isoformat()
        else:  # categorica -> monta a mini-tabela de distribuição desta coluna
            contagem = serie.value_counts(dropna=True, normalize=True).head(max_categorias)
            tabela = contagem.reset_index()
            tabela.columns = ["valor", "pct_do_total"]
            tabela["pct_do_total"] = (tabela["pct_do_total"] * 100).round(2)

            pct_mostrado = tabela["pct_do_total"].sum()
            if pct_mostrado < 99.99:  # sobra algo fora do top-N -> soma teria que dar 100%
                outros = pd.DataFrame([{
                    "valor": "Outros valores",
                    "pct_do_total": round(100 - pct_mostrado, 2),
                }])
                tabela = pd.concat([tabela, outros], ignore_index=True)
            distribuicoes[col] = tabela

        linhas_perfil.append(registro)

        # Alertas: checagens independentes, combinadas numa única linha por coluna
        motivos = []
        if n_unicos <= 1:
            motivos.append(("Constante", "Apenas um valor distinto"))
        if pct_nulos > limite_nulos:
            motivos.append(("Excesso de nulos", f"{pct_nulos}% ausentes"))
        if n_unicos > 1:
            freq_top = serie.value_counts(normalize=True, dropna=True).iloc[0] * 100
            if freq_top > limite_constante:
                motivos.append(("Quase constante", f"{freq_top:.1f}% dos registros no mesmo valor"))
        if motivos:
            linhas_alertas.append({
                "coluna": col,
                "alertas": ", ".join(m for m, _ in motivos),
                "detalhe": " | ".join(d for _, d in motivos),
            })

    df_perfil = pd.DataFrame(linhas_perfil)
    df_alertas = pd.DataFrame(linhas_alertas) if linhas_alertas else pd.DataFrame(
        columns=["coluna", "alertas", "detalhe"])

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df_perfil.to_excel(writer, sheet_name="Perfil Geral", index=False)
        df_alertas.to_excel(writer, sheet_name="Alertas de Qualidade", index=False)

    _escrever_distribuicoes(caminho_saida, distribuicoes)
    _formatar_planilha(caminho_saida)
    print(f"[OK] Relatório exportado para: {caminho_saida}")
    print(f"     {len(distribuicoes)} colunas categóricas com tabela de distribuição.")
    print(f"     {len(linhas_alertas)} colunas com algum alerta de qualidade.")


def _escrever_distribuicoes(caminho: Path, distribuicoes: dict) -> None:
    """
    Escreve, na aba 'Distribuicao de Valores', um bloco (tabelinha) por
    coluna categórica: título da coluna em negrito, cabeçalho, linhas de
    valor/percentual, e uma linha em branco separando do próximo bloco.
    """
    wb = load_workbook(caminho)
    aba = wb.create_sheet("Distribuicao de Valores")

    fonte_titulo = Font(name="Arial", size=10, bold=True)
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, italic=True)
    fonte_padrao = Font(name="Arial", size=10)

    linha_atual = 1
    for col, tabela in distribuicoes.items():
        aba.cell(row=linha_atual, column=1, value=f"Coluna: {col}").font = fonte_titulo
        linha_atual += 1

        aba.cell(row=linha_atual, column=1, value="valor").font = fonte_cabecalho
        aba.cell(row=linha_atual, column=2, value="pct_do_total").font = fonte_cabecalho
        linha_atual += 1

        for _, linha in tabela.iterrows():
            aba.cell(row=linha_atual, column=1, value=linha["valor"]).font = fonte_padrao
            aba.cell(row=linha_atual, column=2, value=linha["pct_do_total"]).font = fonte_padrao
            linha_atual += 1

        linha_atual += 2  # linha em branco entre tabelinhas

    aba.column_dimensions["A"].width = 35
    aba.column_dimensions["B"].width = 15
    wb.save(caminho)


def _formatar_planilha(caminho: Path) -> None:
    """Formatação básica (fonte Arial, cabeçalho em negrito, largura de
    coluna, 1ª linha congelada) — aplicada só nas abas 'Perfil Geral' e
    'Alertas de Qualidade', que têm uma única tabela com cabeçalho na
    linha 1. A aba 'Distribuicao de Valores' já é formatada na escrita,
    por ter várias tabelinhas em posições diferentes."""
    wb = load_workbook(caminho)
    fonte_padrao = Font(name="Arial", size=10)
    fonte_cabecalho = Font(name="Arial", size=10, bold=True)

    for nome_aba in ("Perfil Geral", "Alertas de Qualidade"):
        aba = wb[nome_aba]
        aba.freeze_panes = "A2"
        for linha in aba.iter_rows():
            for celula in linha:
                celula.font = fonte_cabecalho if celula.row == 1 else fonte_padrao

        for coluna in aba.columns:
            maior = max((len(str(c.value)) for c in coluna if c.value is not None), default=10)
            letra = coluna[0].column_letter
            aba.column_dimensions[letra].width = min(maior + 2, 50)

    wb.save(caminho)


# ----------------------------------------------------------------------
# 5. EXECUÇÃO DE EXEMPLO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    CAMINHO_CSV = Path("dados_csv/SEU_ARQUIVO.csv")  # ajuste para o seu arquivo
    CAMINHO_EXCEL = Path("relatorio_perfil_colunas.xlsx")

    df = pd.read_csv(CAMINHO_CSV, sep=";", encoding="latin-1", low_memory=False)

    print("\n>>> PERFIL GERAL DAS COLUNAS (console)\n")
    perfil = perfil_colunas(df)
    print(perfil.to_string(index=False))

    print("\n>>> ALERTAS DE QUALIDADE (console)\n")
    alertas_qualidade(df)

    print("\n>>> EXPORTANDO RELATÓRIO PARA PLANILHA\n")
    gerar_relatorio_excel(df, CAMINHO_EXCEL)

    # Depois de definir seu alvo (ex: uma coluna que indica se o CID
    # principal da internação é de intoxicação por agrotóxico, grupo T60
    # da CID-10), descomente e ajuste:
    #
    # relacao_com_alvo(df, coluna_alvo="NOME_DA_SUA_COLUNA_ALVO")