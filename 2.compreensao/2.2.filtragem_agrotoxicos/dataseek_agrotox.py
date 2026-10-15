"""
Associação das colunas com o alvo HOSPITAL (base já filtrada)
=============================================================

O que este script faz:
  1. Carrega a base JÁ FILTRADA (gerada por filtrar_base.py). Os filtros
     (ex: AGENTE_TOX in {2, 3, 4}) ficam todos em filtrar_base.py — aqui não
     se filtra nada, então novos filtros não exigem mexer neste script.
  2. Gera o perfil das demais colunas JÁ NA BASE FILTRADA (tipo, % de nulos,
     distribuição por coluna e alertas de qualidade) em um Excel próprio.
     Este script é autossuficiente: não depende de nenhum outro arquivo.
  3. Prepara o alvo binário a partir de HOSPITAL (1 -> 1, 2 -> 0; demais
     códigos, como "ignorado", e nulos ficam de fora da análise de alvo).
  4. Mede a associação de cada coluna com o alvo e exporta um relatório
     Excel com:
        - Resumo               : funil de linhas, distribuição do filtro e do alvo
        - Ranking Associacao   : V de Cramér, informação mútua, alertas de vazamento
        - Taxa por Categoria   : taxa de hospitalização por valor, nas melhores colunas
        - Redundancia          : associação ENTRE as melhores colunas (evita duplicadas)

Por que não usar correlação de Pearson em tudo?
  Quase todas as colunas do SINAN/DATASUS são CÓDIGOS categóricos (1, 2, 9...).
  Pearson sobre códigos assume uma ordem/escala que não existe (9 = "ignorado"
  não é "maior" que 1). Por isso a medida principal é o V de Cramér (0 a 1),
  complementada pela informação mútua. Para colunas numéricas contínuas, também
  é reportada a correlação ponto-bisserial (Pearson com alvo binário).

Pré-requisitos:
    pip install pandas numpy scikit-learn openpyxl
"""

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font
from sklearn.metrics import mutual_info_score


# ----------------------------------------------------------------------
# CONFIGURAÇÃO
# ----------------------------------------------------------------------
CAMINHO_CSV = Path("dados_csv/base_filtrada.csv")   # saída do filtrar_base.py
PASTA_SAIDA = Path("relatorios")

COLUNA_ALVO = "HOSPITAL"
VALORES_ALVO_POSITIVO = [1]    # confirme no dicionário de dados do seu sistema
VALORES_ALVO_NEGATIVO = [2]    # outros códigos (ex: 9 = ignorado) saem da análise

MAX_CATEGORIAS = 30        # acima disso, o excedente é agrupado em "OUTROS"
LIMITE_NUMERICA = 15       # numérica com mais valores distintos que isso = contínua
BINS_NUMERICAS = 10        # nº de faixas (quantis) para discretizar numéricas
TOP_N_COLUNAS = 15         # quantas colunas entram em "Taxa por Categoria"
TOP_N_REDUNDANCIA = 20     # quantas colunas entram na matriz de redundância

# Colunas numéricas que são códigos (tratadas como categóricas, não como quantidades)
PREFIXOS_CODIGO = ("ID_", "CO_", "COD", "SG_UF", "UF_", "CNES", "CEP")
TRECHOS_CODIGO = ("MUN", "OCUPA")

LIMITE_VAZAMENTO = 0.50          # V acima disso (com o alvo ou com a nulidade) = suspeita
LIMITE_ESPARSA = 90.0            # % de nulos acima disso = coluna muito esparsa

# Campos que, pela natureza, costumam ser preenchidos DEPOIS (ou por causa) da
# internação. Lista inicial: confirme no dicionário de dados e ajuste.
POSSIVEL_VAZAMENTO = {
    "DTINTERNA", "UF_HOSP", "MUN_HOSP", "EVOLUCAO",
    "DT_OBITO", "DT_ENCERRA", "CLASSI_FIN", "CRITERIO",
}


# ----------------------------------------------------------------------
# 1. CARGA E ALVO
# ----------------------------------------------------------------------
def carregar_base(caminho: Path) -> pd.DataFrame:
    """Lê a base já filtrada (saída do filtrar_base.py)."""
    if not caminho.exists():
        raise FileNotFoundError(f"'{caminho}' não existe. Rode antes o filtrar_base.py para gerá-la.")
    df = pd.read_csv(caminho, sep=";", encoding="latin-1", low_memory=False)

    if COLUNA_ALVO not in df.columns:
        raise KeyError(f"Coluna alvo '{COLUNA_ALVO}' não encontrada. Colunas disponíveis: {list(df.columns)}")
    return df


def preparar_alvo(df: pd.DataFrame):
    """Retorna (df apenas com linhas de alvo válido, alvo binário 0/1, contagem bruta do alvo)."""
    codigo = pd.to_numeric(df[COLUNA_ALVO], errors="coerce")
    contagem_bruta = codigo.value_counts(dropna=False).sort_index()

    alvo = pd.Series(np.nan, index=df.index)
    alvo[codigo.isin(VALORES_ALVO_POSITIVO)] = 1
    alvo[codigo.isin(VALORES_ALVO_NEGATIVO)] = 0

    valido = alvo.notna()
    return df[valido].copy(), alvo[valido].astype(int), contagem_bruta


# ----------------------------------------------------------------------
# 2. TRANSFORMAÇÃO DE CADA COLUNA EM CATEGORIAS COMPARÁVEIS
# ----------------------------------------------------------------------
def _eh_data(col: str, serie: pd.Series) -> bool:
    nome = col.upper()
    return pd.api.types.is_datetime64_any_dtype(serie) or nome.startswith("DT") or nome.startswith("DATA")


def _eh_codigo(col: str) -> bool:
    """Colunas numéricas que na verdade são IDENTIFICADORES/códigos (município, UF, ocupação...).
    Não faz sentido fatiá-las em faixas de quantil; são tratadas como categóricas."""
    nome = col.upper()
    return nome.startswith(PREFIXOS_CODIGO) or any(t in nome for t in TRECHOS_CODIGO)


def _agrupar_raras(s: pd.Series) -> pd.Series:
    """Mantém as MAX_CATEGORIAS mais frequentes (e sempre o 'NA'); o resto vira 'OUTROS'."""
    top = s.value_counts().head(MAX_CATEGORIAS).index
    return s.where(s.isin(top) | (s == "NA"), "OUTROS")


def preparar_categoria(col: str, serie: pd.Series):
    """
    Converte a coluna em uma série de rótulos (string) para medir associação.
    Nulos viram a categoria 'NA' — em bases administrativas, o "não preenchido"
    costuma carregar informação e não deve ser descartado.

    Retorna (categorias, ordem_dos_rotulos | None, tipo, corr_pontobiserial | None).
    """
    if _eh_data(col, serie):
        mes = pd.to_datetime(serie, errors="coerce").dt.month.astype("Int64").astype("string")
        return _agrupar_raras(mes.fillna("NA")).astype(str), None, "data (mês)", None

    if pd.api.types.is_numeric_dtype(serie):
        if serie.nunique(dropna=True) > LIMITE_NUMERICA and not _eh_codigo(col):
            faixas = pd.qcut(serie, q=BINS_NUMERICAS, duplicates="drop")
            ordem = [str(i) for i in faixas.cat.categories] + ["NA"]
            cat = faixas.astype(object).where(faixas.notna(), "NA").astype(str)
            return cat, ordem, "numérica (faixas)", None  # corr. calculada fora (precisa do alvo)

        nao_nulo = serie.dropna()
        if len(nao_nulo) and (nao_nulo % 1 == 0).all():
            serie = serie.astype("Int64")  # 2.0 -> "2"

    s = serie.astype("string").fillna("NA")
    return _agrupar_raras(s).astype(str), None, "categórica", None


# ----------------------------------------------------------------------
# 3. MEDIDAS DE ASSOCIAÇÃO
# ----------------------------------------------------------------------
def cramers_v(x: pd.Series, y: pd.Series) -> float:
    """V de Cramér com correção de viés (Bergsma). 0 = sem associação, 1 = associação perfeita."""
    tab = pd.crosstab(x, y)
    if tab.shape[0] < 2 or tab.shape[1] < 2:
        return 0.0
    obs = tab.values.astype(float)
    n = obs.sum()
    esperado = np.outer(obs.sum(axis=1), obs.sum(axis=0)) / n
    chi2 = ((obs - esperado) ** 2 / esperado).sum()
    r, k = obs.shape
    phi2c = max(0.0, chi2 / n - (k - 1) * (r - 1) / (n - 1))
    rc = r - (r - 1) ** 2 / (n - 1)
    kc = k - (k - 1) ** 2 / (n - 1)
    denom = min(kc - 1, rc - 1)
    return float(np.sqrt(phi2c / denom)) if denom > 0 else 0.0


def info_mutua_normalizada(x: pd.Series, y: pd.Series) -> float:
    """Informação mútua dividida pela entropia do alvo: fração da incerteza do alvo explicada pela coluna."""
    p = y.value_counts(normalize=True).values
    entropia = -(p * np.log(p)).sum()
    return float(mutual_info_score(x, y) / entropia) if entropia > 0 else 0.0


def _faixa(v: float) -> str:
    if v >= 0.20:
        return "forte"
    if v >= 0.10:
        return "moderada"
    if v >= 0.05:
        return "fraca"
    return "baixa"


def calcular_ranking(df: pd.DataFrame, alvo: pd.Series):
    """Calcula as medidas para cada coluna (exceto o alvo). Retorna (ranking, cats, ordens)."""
    linhas, cats, ordens = [], {}, {}

    for col in df.columns:
        if col == COLUNA_ALVO:
            continue
        serie = df[col]
        cat, ordem, tipo, _ = preparar_categoria(col, serie)
        cats[col], ordens[col] = cat, ordem

        pct_nulos = serie.isna().mean() * 100
        v = cramers_v(cat, alvo)
        mi = info_mutua_normalizada(cat, alvo)
        v_nulos = cramers_v(serie.isna().astype(int), alvo) if serie.isna().any() else 0.0
        r_pb = None
        if tipo == "numérica (faixas)":
            r_pb = serie.corr(alvo)

        alertas = []
        n_cat = cat.nunique()
        if n_cat <= 1:
            alertas.append("constante no recorte")
        if pct_nulos > LIMITE_ESPARSA:
            alertas.append(f"muito esparsa ({pct_nulos:.1f}% nulos)")
        if serie.nunique(dropna=True) > MAX_CATEGORIAS and tipo == "categórica":
            alertas.append("alta cardinalidade (agrupada em OUTROS)")

        vazamento = []
        if col in POSSIVEL_VAZAMENTO:
            vazamento.append("campo tipicamente pós-desfecho/ligado à internação")
        if v_nulos >= LIMITE_VAZAMENTO:
            vazamento.append(f"nulidade quase determina o alvo (V={v_nulos:.2f})")
        if v >= LIMITE_VAZAMENTO:
            vazamento.append(f"associação suspeitamente alta (V={v:.2f})")

        linhas.append({
            "coluna": col,
            "tipo": tipo,
            "pct_nulos": round(pct_nulos, 2),
            "n_categorias": n_cat,
            "cramers_v": round(v, 3),
            "faixa_associacao": _faixa(v),
            "info_mutua_norm": round(mi, 3),
            "corr_pontobiserial": None if r_pb is None or pd.isna(r_pb) else round(r_pb, 3),
            "assoc_nulidade_alvo": round(v_nulos, 3),
            "suspeita_vazamento": "Sim" if vazamento else "Não",
            "alertas": "; ".join(vazamento + alertas),
        })

    ranking = pd.DataFrame(linhas).sort_values("cramers_v", ascending=False).reset_index(drop=True)
    return ranking, cats, ordens


# ----------------------------------------------------------------------
# 4. TABELAS DE APOIO
# ----------------------------------------------------------------------
def tabela_taxa(cat: pd.Series, alvo: pd.Series, ordem=None) -> pd.DataFrame:
    """Para cada valor da coluna: nº de linhas, % do total e taxa de hospitalização."""
    base = pd.DataFrame({"valor": cat.values, "alvo": alvo.values})
    g = base.groupby("valor")["alvo"].agg(n="size", hospitalizados="sum")
    taxa_geral = alvo.mean() * 100

    g["pct_do_total"] = g["n"] / g["n"].sum() * 100
    g["taxa_hospitalizacao_pct"] = g["hospitalizados"] / g["n"] * 100
    g["dif_vs_geral_pp"] = g["taxa_hospitalizacao_pct"] - taxa_geral

    if ordem:
        g = g.reindex([o for o in ordem if o in g.index])
    else:
        g = g.sort_values("n", ascending=False)

    g = g[["n", "pct_do_total", "hospitalizados", "taxa_hospitalizacao_pct", "dif_vs_geral_pp"]].round(2)
    g.loc["TOTAL"] = [g["n"].sum(), 100.0, g["hospitalizados"].sum(), round(taxa_geral, 2), 0.0]
    return g.reset_index().rename(columns={"index": "valor"})


def matriz_redundancia(colunas: list, cats: dict) -> pd.DataFrame:
    """V de Cramér entre pares de colunas candidatas: valores altos = informação repetida."""
    m = pd.DataFrame(1.0, index=colunas, columns=colunas)
    for i, a in enumerate(colunas):
        for b in colunas[i + 1:]:
            m.loc[a, b] = m.loc[b, a] = round(cramers_v(cats[a], cats[b]), 2)
    return m


def montar_resumo(df_base, contagem_alvo, df_valido, alvo) -> pd.DataFrame:
    n_filt, n_valido = len(df_base), len(df_valido)
    pos, neg = int(alvo.sum()), int((alvo == 0).sum())
    linhas = [
        (f"Linhas na base analisada ({CAMINHO_CSV.name})", n_filt, 100.0),
        (f"Distribuição de {COLUNA_ALVO} na base", None, None),
    ]
    for cod, n in contagem_alvo.items():
        linhas.append((f"   {COLUNA_ALVO} = {cod}", int(n), round(n / n_filt * 100, 2)))

    linhas += [
        (f"Linhas com alvo válido (usadas na análise)", n_valido, round(n_valido / n_filt * 100, 2)),
        (f"   Classe positiva ({COLUNA_ALVO} in {VALORES_ALVO_POSITIVO})", pos, round(pos / n_valido * 100, 2)),
        (f"   Classe negativa ({COLUNA_ALVO} in {VALORES_ALVO_NEGATIVO})", neg, round(neg / n_valido * 100, 2)),
        ("Razão negativa:positiva", round(neg / pos, 2) if pos else None, None),
    ]
    return pd.DataFrame(linhas, columns=["item", "n", "pct"])


# ----------------------------------------------------------------------
# 5. EXPORTAÇÃO PARA EXCEL
# ----------------------------------------------------------------------
def _py(v):
    return v.item() if hasattr(v, "item") else v


def _escrever_taxas(wb, blocos: list) -> None:
    aba = wb.create_sheet("Taxa por Categoria")
    f_titulo = Font(name="Arial", size=10, bold=True)
    f_cab = Font(name="Arial", size=10, bold=True, italic=True)
    f_pad = Font(name="Arial", size=10)

    linha = 1
    for titulo, tabela in blocos:
        aba.cell(row=linha, column=1, value=titulo).font = f_titulo
        linha += 1
        for j, nome in enumerate(tabela.columns, start=1):
            aba.cell(row=linha, column=j, value=nome).font = f_cab
        linha += 1
        for _, reg in tabela.iterrows():
            negrito = reg["valor"] == "TOTAL"
            for j, nome in enumerate(tabela.columns, start=1):
                c = aba.cell(row=linha, column=j, value=_py(reg[nome]))
                c.font = Font(name="Arial", size=10, bold=negrito)
            linha += 1
        linha += 2

    aba.column_dimensions["A"].width = 38
    for letra in "BCDEF":
        aba.column_dimensions[letra].width = 22


def _formatar(wb) -> None:
    f_pad, f_cab = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)

    for nome in ("Resumo", "Ranking Associacao", "Redundancia"):
        aba = wb[nome]
        aba.freeze_panes = "B2" if nome == "Redundancia" else "A2"
        for linha in aba.iter_rows():
            for c in linha:
                c.font = f_cab if (c.row == 1 or (nome == "Redundancia" and c.column == 1)) else f_pad
        for coluna in aba.columns:
            maior = max((len(str(c.value)) for c in coluna if c.value is not None), default=10)
            aba.column_dimensions[coluna[0].column_letter].width = min(maior + 2, 60)

    # escala de cor: quanto mais laranja, maior a associação
    regra = ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                           end_type="num", end_value=1, end_color="F4B183")

    aba = wb["Ranking Associacao"]
    cab = {c.value: c.column_letter for c in aba[1]}
    for nome_col in ("cramers_v", "info_mutua_norm", "assoc_nulidade_alvo"):
        letra = cab[nome_col]
        aba.conditional_formatting.add(f"{letra}2:{letra}{aba.max_row}", regra)

    aba = wb["Redundancia"]
    aba.conditional_formatting.add(f"B2:{aba.cell(row=aba.max_row, column=aba.max_column).coordinate}", regra)


def exportar_relatorio(caminho: Path, resumo, ranking, matriz, blocos) -> None:
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        ranking.to_excel(writer, sheet_name="Ranking Associacao", index=False)
        matriz.to_excel(writer, sheet_name="Redundancia")

    wb = load_workbook(caminho)
    _escrever_taxas(wb, blocos)
    _formatar(wb)
    wb.save(caminho)


# ----------------------------------------------------------------------
# 6. PERFIL DAS COLUNAS DA BASE FILTRADA (Excel)
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
# 7. EXECUÇÃO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    df_base = carregar_base(CAMINHO_CSV)
    print(f"Base analisada: {CAMINHO_CSV} | {len(df_base)} linhas, {df_base.shape[1]} colunas")

    # (a) perfil das colunas da base (abas por coluna, alertas etc.)
    gerar_relatorio_excel(df_base, PASTA_SAIDA / "perfil_base_filtrada.xlsx")

    # (b) alvo e associações
    df_valido, alvo, contagem_alvo = preparar_alvo(df_base)
    print(f"Linhas com {COLUNA_ALVO} válido: {len(df_valido)} | taxa de positivos: {alvo.mean()*100:.2f}%")

    ranking, cats, ordens = calcular_ranking(df_valido, alvo)

    candidatas = ranking[(ranking["suspeita_vazamento"] == "Não") & (ranking["n_categorias"] > 1)]["coluna"].tolist()

    blocos = []
    for col in candidatas[:TOP_N_COLUNAS]:
        v = ranking.loc[ranking["coluna"] == col, "cramers_v"].iloc[0]
        blocos.append((f"Coluna: {col}  (V de Cramér = {v})", tabela_taxa(cats[col], alvo, ordens[col])))

    matriz = matriz_redundancia(candidatas[:TOP_N_REDUNDANCIA], cats)
    resumo = montar_resumo(df_base, contagem_alvo, df_valido, alvo)

    caminho = PASTA_SAIDA / "associacao_com_hospital.xlsx"
    exportar_relatorio(caminho, resumo, ranking, matriz, blocos)

    print(f"\n[OK] Relatórios salvos em '{PASTA_SAIDA}/'")
    print("\nTop 15 colunas por associação com o alvo:")
    print(ranking[["coluna", "cramers_v", "faixa_associacao", "suspeita_vazamento"]].head(15).to_string(index=False))