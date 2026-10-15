"""
Análise dos campos de agente (AGENTE_1, AGENTE_2, AGENTE_3) na base filtrada
=============================================================================

Contexto
--------
AGENTE_TOX é um CÓDIGO (grupo do agente: 02, 03, 04...). AGENTE_1/2/3 são
TEXTO LIVRE digitado na notificação (nome comercial/popular de até 3 produtos).
São campos diferentes: filtrar AGENTE_TOX não restringe o que aparece em
AGENTE_1/2/3 — por isso lá surgem outros produtos (coexposição, erro de
classificação, texto "ignorado", grafias diferentes do mesmo produto).

O que este script faz
---------------------
  1. Normaliza o texto (maiúsculas, sem acento, sem pontuação).
  2. Classifica cada agente em um GRUPO (glifosato, inseticida, medicamento,
     produto domiciliar, ignorado...) por dicionário de palavras-chave.
  3. AUDITORIA do filtro: cruza AGENTE_TOX x grupo de AGENTE_1, e mostra quantos
     casos têm agente aparentemente não-praguicida ou coexposição.
  3b. DISTRIBUIÇÃO POR TIPO DE USO: para CADA valor de AGENTE_TOX (2, 3, 4...),
     mostra como se comportam AGENTE_1, AGENTE_2 e AGENTE_3 — nomes mais
     frequentes e grupos — para comparar o que aparece em uso agrícola x
     doméstico x saúde pública.
  4. Cria variáveis derivadas (nº de agentes, coexposição, grupo do agente 1...)
     e mede a associação de cada uma com o alvo HOSPITAL (V de Cramér, informação
     mútua e taxa de hospitalização por categoria).
  5. Lista os nomes NÃO classificados mais frequentes, para você ampliar o
     dicionário (GRUPOS abaixo) — ele é só um ponto de partida.

Pré-requisitos:
    pip install pandas numpy scikit-learn openpyxl

Entrada: dados_csv/base_filtrada.csv (saída do filtrar_base.py)
Saída:   relatorios/analise_agentes.xlsx + resumo no console (por AGENTE_TOX)
"""

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from sklearn.metrics import mutual_info_score

# ----------------------------------------------------------------------
# CONFIGURAÇÃO
# ----------------------------------------------------------------------
CAMINHO_CSV = Path("dados_csv/base_filtrada.csv")
PASTA_SAIDA = Path("relatorios")

COLUNAS_AGENTE = ["AGENTE_1", "AGENTE_2", "AGENTE_3"]
COLUNA_GRUPO_TOX = "AGENTE_TOX"
COLUNA_ALVO = "HOSPITAL"
VALORES_ALVO_POSITIVO = [1]   # confirme no dicionário de dados
VALORES_ALVO_NEGATIVO = [2]   # outros códigos (ex: 9 = ignorado) saem da parte de associação

TOP_NAO_CLASSIFICADOS = 60    # quantos nomes não classificados listar por coluna
TOP_NOMES_POR_TOX = 15        # quantos nomes mostrar por (AGENTE_TOX x AGENTE_n)

# Rótulos do campo AGENTE_TOX (só para deixar os títulos legíveis)
ROTULOS_TOX = {
    "1": "Medicamento", "2": "Agrotóxico/uso agrícola", "3": "Agrotóxico/uso doméstico",
    "4": "Agrotóxico/uso saúde pública", "5": "Raticida", "6": "Produto veterinário",
    "7": "Produto de uso domiciliar", "8": "Cosmético/higiene pessoal",
    "9": "Produto químico de uso industrial", "10": "Metal", "11": "Drogas de abuso",
    "12": "Planta tóxica", "13": "Alimento e bebida", "14": "Outro", "99": "Ignorado",
}

# Dicionário de classificação: (nome do grupo, regex sobre o texto JÁ normalizado).
# A ORDEM IMPORTA: vale a primeira regra que casar. É um ponto de partida —
# amplie com os nomes da aba "Nao Classificados" e, para nomes comerciais,
# consulte fontes oficiais (ex: AGROFIT/MAPA, monografias da Anvisa).
GRUPOS = [
    ("Chumbinho (raticida clandestino)", r"CHUMBINHO"),
    ("Glifosato", r"GLIFOS|GLYPHOS|ROUNDUP|RANDAP"),
    ("Herbicida (outros)", r"HERBICIDA|TORDON|PARAQUAT|GRAMOXONE|\b2 4 D\b|ATRAZINA|METSULFUR|CLETODIM|DIURON|PICLORAM|GLUFOSINATO"),
    ("Inseticida", r"INSETICIDA|CIPERMETRINA|PERMETRINA|DELTAMETRINA|FENDONA|ACEFATO|CLORPIRIFOS|MALATION|FIPRONIL|IMIDACLOPRID|ORGANOFOSFORAD|CARBAMATO"),
    ("Fungicida", r"FUNGICIDA|MANCOZEB|TEBUCONAZOL|AZOXISTROBINA"),
    ("Raticida (outros)", r"RATICIDA|VENENO DE RATO|RATOX"),
    ("Agrotoxico (generico)", r"AGROTOXICO|DEFENSIVO|VENENO"),
    ("Medicamento", r"CLONAZEPAM|ALPRAZOLAM|DIAZEPAM|FLUOXETINA|AMITRIPTILINA|PARACETAMOL|DIPIRONA|CARBAMAZEPINA|RIVOTRIL|MEDICAMENTO|REMEDIO"),
    ("Alcool/bebida", r"BEBIDA ALCOOLICA|ALCOOL|CERVEJA|CACHACA|VODKA|PINGA"),   # "ALCOOL" pode ser álcool de limpeza: revisar
    ("Droga de abuso", r"COCAINA|CRACK|MACONHA|MERLA|DROGA"),
    ("Produto domiciliar/saneante", r"AGUA SANITARIA|HIPOCLORITO|NAFTALINA|DESINFETANTE|SODA CAUSTICA|AMONIA|ALVEJANTE|DETERGENTE|CLORO"),
]
REGRAS = [(nome, re.compile(rx)) for nome, rx in GRUPOS]

PRAGUICIDAS = {
    "Chumbinho (raticida clandestino)", "Glifosato", "Herbicida (outros)", "Inseticida",
    "Fungicida", "Raticida (outros)", "Agrotoxico (generico)",
}
NAO_PRAGUICIDAS = {"Medicamento", "Alcool/bebida", "Droga de abuso", "Produto domiciliar/saneante"}

GRUPO_VAZIO = "(vazio)"
GRUPO_IGNORADO = "IGNORADO/NAO INFORMADO"
GRUPO_NAO_CLASSIFICADO = "NAO CLASSIFICADO"
PLACEHOLDERS = re.compile(r"^(IGNORADO|IGN|NAO INFORMADO|NI|NAO SABE|NAO SE APLICA|DESCONHECIDO|SEM INFORMACAO)$")


# ----------------------------------------------------------------------
# 1. NORMALIZAÇÃO E CLASSIFICAÇÃO DO TEXTO
# ----------------------------------------------------------------------
def normalizar(texto) -> str:
    """'Água Sanitária!' -> 'AGUA SANITARIA' ; '2,4-D' -> '2 4 D'."""
    if pd.isna(texto):
        return ""
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", t).strip()


def classificar(texto_normalizado: str) -> str:
    if texto_normalizado == "":
        return GRUPO_VAZIO
    if PLACEHOLDERS.match(texto_normalizado):
        return GRUPO_IGNORADO
    for nome, regra in REGRAS:
        if regra.search(texto_normalizado):
            return nome
    return GRUPO_NAO_CLASSIFICADO


def preparar_agentes(df: pd.DataFrame):
    """Retorna (df de nomes normalizados, df de grupos) com uma coluna por AGENTE_n."""
    normalizados = pd.DataFrame({c: df[c].map(normalizar) for c in COLUNAS_AGENTE})
    grupos = normalizados.apply(lambda s: s.map(classificar))
    return normalizados, grupos


def _tipo(grupo: str) -> str:
    if grupo in PRAGUICIDAS:
        return "praguicida (provavel)"
    if grupo in NAO_PRAGUICIDAS:
        return "nao praguicida"
    if grupo in (GRUPO_VAZIO, GRUPO_IGNORADO):
        return "vazio/ignorado"
    return "nao classificado"


# ----------------------------------------------------------------------
# 2. VARIÁVEIS DERIVADAS
# ----------------------------------------------------------------------
def criar_variaveis(grupos: pd.DataFrame) -> pd.DataFrame:
    """Variáveis categóricas derivadas dos agentes, prontas para medir associação com o alvo."""
    g1, g2, g3 = (grupos[c] for c in COLUNAS_AGENTE)
    outros = grupos[COLUNAS_AGENTE[1:]]
    informado = ~grupos.isin([GRUPO_VAZIO, GRUPO_IGNORADO])

    def sim_nao(mascara):
        return mascara.map({True: "Sim", False: "Nao"})

    return pd.DataFrame({
        "ag1_grupo": g1,
        "ag1_tipo": g1.map(_tipo),
        "n_agentes_informados": informado.sum(axis=1).astype(str),
        "coagente_nao_praguicida": sim_nao(outros.isin(NAO_PRAGUICIDAS).any(axis=1)),
        "coagente_praguicida": sim_nao(outros.isin(PRAGUICIDAS).any(axis=1)),
        "envolve_chumbinho": sim_nao((grupos == "Chumbinho (raticida clandestino)").any(axis=1)),
        "envolve_glifosato": sim_nao((grupos == "Glifosato").any(axis=1)),
        "envolve_medicamento_ou_alcool_ou_droga": sim_nao(
            grupos.isin({"Medicamento", "Alcool/bebida", "Droga de abuso"}).any(axis=1)),
    }, index=grupos.index)


# ----------------------------------------------------------------------
# 3. MEDIDAS DE ASSOCIAÇÃO COM O ALVO
# ----------------------------------------------------------------------
def preparar_alvo(df: pd.DataFrame) -> pd.Series:
    """Alvo binário (1/0) apenas nas linhas com HOSPITAL válido; demais ficam NaN."""
    codigo = pd.to_numeric(df[COLUNA_ALVO], errors="coerce")
    alvo = pd.Series(np.nan, index=df.index)
    alvo[codigo.isin(VALORES_ALVO_POSITIVO)] = 1
    alvo[codigo.isin(VALORES_ALVO_NEGATIVO)] = 0
    return alvo


def cramers_v(x: pd.Series, y: pd.Series) -> float:
    """V de Cramér com correção de viés. 0 = sem associação, 1 = perfeita."""
    tab = pd.crosstab(x, y)
    if tab.shape[0] < 2 or tab.shape[1] < 2:
        return 0.0
    obs = tab.values.astype(float)
    n = obs.sum()
    esperado = np.outer(obs.sum(axis=1), obs.sum(axis=0)) / n
    chi2 = ((obs - esperado) ** 2 / esperado).sum()
    r, k = obs.shape
    phi2c = max(0.0, chi2 / n - (k - 1) * (r - 1) / (n - 1))
    rc, kc = r - (r - 1) ** 2 / (n - 1), k - (k - 1) ** 2 / (n - 1)
    denom = min(kc - 1, rc - 1)
    return float(np.sqrt(phi2c / denom)) if denom > 0 else 0.0


def info_mutua_normalizada(x: pd.Series, y: pd.Series) -> float:
    p = y.value_counts(normalize=True).values
    entropia = -(p * np.log(p)).sum()
    return float(mutual_info_score(x, y) / entropia) if entropia > 0 else 0.0


def _faixa(v: float) -> str:
    return "forte" if v >= 0.20 else "moderada" if v >= 0.10 else "fraca" if v >= 0.05 else "baixa"


def ranking_variaveis(feat: pd.DataFrame, alvo: pd.Series) -> pd.DataFrame:
    linhas = []
    for col in feat.columns:
        v = cramers_v(feat[col], alvo)
        linhas.append({
            "variavel": col,
            "n_categorias": feat[col].nunique(),
            "cramers_v": round(v, 3),
            "faixa_associacao": _faixa(v),
            "info_mutua_norm": round(info_mutua_normalizada(feat[col], alvo), 3),
        })
    return pd.DataFrame(linhas).sort_values("cramers_v", ascending=False).reset_index(drop=True)


def tabela_taxa(cat: pd.Series, alvo: pd.Series) -> pd.DataFrame:
    """Por valor: nº de linhas, % do total e taxa de hospitalização (soma fecha em 100%)."""
    base = pd.DataFrame({"valor": cat.values, "alvo": alvo.values})
    g = base.groupby("valor")["alvo"].agg(n="size", hospitalizados="sum")
    taxa_geral = alvo.mean() * 100
    g["pct_do_total"] = g["n"] / g["n"].sum() * 100
    g["taxa_hospitalizacao_pct"] = g["hospitalizados"] / g["n"] * 100
    g["dif_vs_geral_pp"] = g["taxa_hospitalizacao_pct"] - taxa_geral
    g = g.sort_values("n", ascending=False)
    g = g[["n", "pct_do_total", "hospitalizados", "taxa_hospitalizacao_pct", "dif_vs_geral_pp"]].round(2)
    g.loc["TOTAL"] = [g["n"].sum(), 100.0, g["hospitalizados"].sum(), round(taxa_geral, 2), 0.0]
    return g.reset_index().rename(columns={"index": "valor"})


# ----------------------------------------------------------------------
# 4. TABELAS DE AUDITORIA
# ----------------------------------------------------------------------
def tabela_grupos_por_agente(grupos: pd.DataFrame) -> pd.DataFrame:
    """Para cada AGENTE_n: nº e % (sobre o TOTAL de linhas, não só as preenchidas) de cada grupo."""
    n = len(grupos)
    partes = []
    for c in COLUNAS_AGENTE:
        vc = grupos[c].value_counts()
        partes.append(pd.DataFrame({f"{c}_n": vc, f"{c}_pct_linhas": (vc / n * 100).round(2)}))
    return pd.concat(partes, axis=1).fillna(0).sort_values(f"{COLUNAS_AGENTE[0]}_n", ascending=False)\
        .rename_axis("grupo").reset_index()


def tabela_nao_classificados(normalizados: pd.DataFrame, grupos: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    n = len(normalizados)
    for c in COLUNAS_AGENTE:
        nomes = normalizados.loc[grupos[c] == GRUPO_NAO_CLASSIFICADO, c]
        for nome, qtd in nomes.value_counts().head(TOP_NAO_CLASSIFICADOS).items():
            linhas.append({"coluna": c, "nome_normalizado": nome, "n": int(qtd),
                           "pct_das_linhas": round(qtd / n * 100, 2)})
    return pd.DataFrame(linhas, columns=["coluna", "nome_normalizado", "n", "pct_das_linhas"])


def imprimir_resumo(df: pd.DataFrame, grupos: pd.DataFrame, feat: pd.DataFrame) -> None:
    n = len(df)
    pct = lambda x: f"{x / n * 100:.1f}%".replace(".", ",")
    print("=" * 70)
    print(f"AUDITORIA DOS AGENTES  ({n:,} linhas na base filtrada)".replace(",", "."))
    print("=" * 70)
    for c in COLUNAS_AGENTE:
        tipos = grupos[c].map(_tipo).value_counts()
        print(f"\n{c} (% sobre o total de linhas):")
        for tipo in ("praguicida (provavel)", "nao praguicida", "nao classificado", "vazio/ignorado"):
            print(f"   {tipo:<24} {pct(tipos.get(tipo, 0))}")
    print("\nVariáveis derivadas (% de linhas = 'Sim'):")
    for c in ("coagente_nao_praguicida", "coagente_praguicida", "envolve_chumbinho",
              "envolve_glifosato", "envolve_medicamento_ou_alcool_ou_droga"):
        print(f"   {c:<42} {pct((feat[c] == 'Sim').sum())}")
    print("=" * 70)


# ----------------------------------------------------------------------
# 4b. DISTRIBUIÇÃO DOS AGENTES POR VALOR DE AGENTE_TOX
# ----------------------------------------------------------------------
def serie_tox(df: pd.DataFrame) -> pd.Series:
    """AGENTE_TOX como texto limpo ('2', '3', '4'); nulo vira '<NA>'."""
    return pd.to_numeric(df[COLUNA_GRUPO_TOX], errors="coerce").astype("Int64").astype(str)


def codigos_tox(tox: pd.Series) -> list:
    return sorted((c for c in tox.unique() if c != "<NA>"), key=int)


def _rotulo(cod: str) -> str:
    return f"{cod} ({ROTULOS_TOX[cod]})" if cod in ROTULOS_TOX else cod


def bloco_nomes(nomes_do_codigo: pd.Series, top_k: int = TOP_NOMES_POR_TOX) -> pd.DataFrame:
    """Nomes normalizados mais frequentes (entre os PREENCHIDOS) + 'Outros nomes' + TOTAL = 100%."""
    preenchidos = nomes_do_codigo[nomes_do_codigo != ""]
    n_pre = len(preenchidos)
    colunas = ["nome", "grupo", "n", "pct_dos_preenchidos"]
    if n_pre == 0:
        return pd.DataFrame(columns=colunas)

    top = preenchidos.value_counts().head(top_k)
    linhas = [{"nome": nome, "grupo": classificar(nome), "n": int(q),
               "pct_dos_preenchidos": round(q / n_pre * 100, 2)} for nome, q in top.items()]
    resto = n_pre - int(top.sum())
    if resto > 0:
        linhas.append({"nome": "Outros nomes", "grupo": "", "n": resto,
                       "pct_dos_preenchidos": round(resto / n_pre * 100, 2)})
    linhas.append({"nome": "TOTAL", "grupo": "", "n": n_pre, "pct_dos_preenchidos": 100.0})
    return pd.DataFrame(linhas, columns=colunas)


def montar_blocos_nomes(normalizados: pd.DataFrame, tox: pd.Series, cod: str) -> list:
    """Um bloco por AGENTE_n para um valor de AGENTE_TOX. Retorna [(titulo, subtitulo, tabela)]."""
    mascara = tox == cod
    n_cod = int(mascara.sum())
    blocos = []
    for c in COLUNAS_AGENTE:
        nomes = normalizados.loc[mascara, c]
        n_pre = int((nomes != "").sum())
        pct_pre = f"{n_pre / n_cod * 100:.1f}%" if n_cod else "0%"
        blocos.append((f"AGENTE_TOX = {_rotulo(cod)} | {c}",
                       f"linhas do código: {n_cod} | {c} preenchido em {n_pre} ({pct_pre})",
                       bloco_nomes(nomes)))
    return blocos


def tabela_grupos_por_tox(grupos_col: pd.Series, tox: pd.Series, codigos: list) -> pd.DataFrame:
    """Grupo x AGENTE_TOX. O % é sobre TODAS as linhas do código (inclui '(vazio)'), então cada coluna soma 100%."""
    cont = pd.crosstab(grupos_col, tox).reindex(columns=codigos, fill_value=0)
    pct = cont / cont.sum() * 100
    out = pd.DataFrame({"grupo": cont.index})
    for cod in codigos:
        out[f"n_tox{cod}"] = cont[cod].values
        out[f"pct_tox{cod}"] = pct[cod].round(2).values
    out = out.iloc[(-cont.sum(axis=1)).argsort(kind="stable").values].reset_index(drop=True)
    total = {"grupo": "TOTAL"}
    for cod in codigos:
        total[f"n_tox{cod}"], total[f"pct_tox{cod}"] = int(cont[cod].sum()), 100.0
    return pd.concat([out, pd.DataFrame([total])], ignore_index=True)


def imprimir_por_tox(df: pd.DataFrame, normalizados: pd.DataFrame, tox: pd.Series, codigos: list, top: int = 5) -> None:
    n = len(df)
    print("=" * 70)
    print("AGENTES POR TIPO DE USO (valor de AGENTE_TOX)")
    print("=" * 70)
    for cod in codigos:
        mascara = tox == cod
        n_cod = int(mascara.sum())
        print(f"\nAGENTE_TOX = {_rotulo(cod)}: {n_cod:,} linhas ({n_cod / n * 100:.1f}% da base)".replace(",", "."))
        for c in COLUNAS_AGENTE:
            nomes = normalizados.loc[mascara, c]
            preenchidos = nomes[nomes != ""]
            print(f"  {c}: preenchido em {len(preenchidos) / n_cod * 100:.1f}% das linhas")
            for nome, q in preenchidos.value_counts().head(top).items():
                print(f"      {nome:<32} {q / len(preenchidos) * 100:5.1f}% dos preenchidos")
    print("=" * 70)


# ----------------------------------------------------------------------
# 5. EXPORTAÇÃO PARA EXCEL
# ----------------------------------------------------------------------
def _py(v):
    return v.item() if hasattr(v, "item") else v


def _escrever_taxas(wb, blocos: list) -> None:
    aba = wb.create_sheet("Taxa por Categoria")
    linha = 1
    for titulo, tabela in blocos:
        aba.cell(row=linha, column=1, value=titulo).font = Font(name="Arial", size=10, bold=True)
        linha += 1
        for j, nome in enumerate(tabela.columns, start=1):
            aba.cell(row=linha, column=j, value=nome).font = Font(name="Arial", size=10, bold=True, italic=True)
        linha += 1
        for _, reg in tabela.iterrows():
            negrito = reg["valor"] == "TOTAL"
            for j, nome in enumerate(tabela.columns, start=1):
                aba.cell(row=linha, column=j, value=_py(reg[nome])).font = Font(name="Arial", size=10, bold=negrito)
            linha += 1
        linha += 2
    aba.column_dimensions["A"].width = 40
    for letra in "BCDEF":
        aba.column_dimensions[letra].width = 22


def _formatar(wb, abas: list) -> None:
    for nome in abas:
        aba = wb[nome]
        aba.freeze_panes = "A2"
        for linha in aba.iter_rows():
            for c in linha:
                c.font = Font(name="Arial", size=10, bold=(c.row == 1))
        for coluna in aba.columns:
            maior = max((len(str(c.value)) for c in coluna if c.value is not None), default=10)
            aba.column_dimensions[coluna[0].column_letter].width = min(maior + 2, 60)


def _escrever_blocos_horizontais(wb, nome_aba: str, blocos: list) -> None:
    """Blocos (titulo, subtitulo, tabela) lado a lado numa mesma aba, para comparar AGENTE_1/2/3."""
    aba = wb.create_sheet(nome_aba)
    col0 = 1
    for titulo, subtitulo, tabela in blocos:
        aba.cell(row=1, column=col0, value=titulo).font = Font(name="Arial", size=10, bold=True)
        aba.cell(row=2, column=col0, value=subtitulo).font = Font(name="Arial", size=9, italic=True)
        for j, nome in enumerate(tabela.columns):
            aba.cell(row=3, column=col0 + j, value=nome).font = Font(name="Arial", size=10, bold=True, italic=True)
        for i, (_, reg) in enumerate(tabela.iterrows()):
            negrito = reg["nome"] == "TOTAL"
            for j, nome in enumerate(tabela.columns):
                aba.cell(row=4 + i, column=col0 + j, value=_py(reg[nome])).font = Font(name="Arial", size=10, bold=negrito)
        for j, largura in enumerate((34, 28, 8, 20)):
            aba.column_dimensions[get_column_letter(col0 + j)].width = largura
        aba.column_dimensions[get_column_letter(col0 + len(tabela.columns))].width = 4  # espaço entre blocos
        col0 += len(tabela.columns) + 1


def _escrever_blocos_verticais(wb, nome_aba: str, blocos: list) -> None:
    """Blocos (titulo, tabela) empilhados, com linha em branco entre eles."""
    aba = wb.create_sheet(nome_aba)
    linha = 1
    for titulo, tabela in blocos:
        aba.cell(row=linha, column=1, value=titulo).font = Font(name="Arial", size=10, bold=True)
        linha += 1
        for j, nome in enumerate(tabela.columns, start=1):
            aba.cell(row=linha, column=j, value=nome).font = Font(name="Arial", size=10, bold=True, italic=True)
        linha += 1
        for _, reg in tabela.iterrows():
            negrito = reg.iloc[0] == "TOTAL"
            for j, nome in enumerate(tabela.columns, start=1):
                aba.cell(row=linha, column=j, value=_py(reg[nome])).font = Font(name="Arial", size=10, bold=negrito)
            linha += 1
        linha += 2
    aba.column_dimensions["A"].width = 38
    for j in range(2, 12):
        aba.column_dimensions[get_column_letter(j)].width = 14


def exportar(caminho: Path, tabelas: dict, blocos: list, extras=(), ordem_abas=()) -> None:
    """Escreve as tabelas simples, o bloco 'Taxa por Categoria' e as abas extras (funções que recebem o workbook)."""
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome, tabela in tabelas.items():
            tabela.to_excel(writer, sheet_name=nome, index=False)
    wb = load_workbook(caminho)
    _escrever_taxas(wb, blocos)
    for escrever in extras:
        escrever(wb)
    _formatar(wb, list(tabelas))
    for pos, nome in enumerate(ordem_abas):   # abas mais importantes primeiro
        wb.move_sheet(nome, offset=pos - wb.sheetnames.index(nome))
    wb.save(caminho)


# ----------------------------------------------------------------------
# 6. EXECUÇÃO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(CAMINHO_CSV, sep=";", encoding="latin-1", low_memory=False)
    faltando = [c for c in COLUNAS_AGENTE + [COLUNA_GRUPO_TOX, COLUNA_ALVO] if c not in df.columns]
    if faltando:
        raise KeyError(f"Colunas não encontradas: {faltando}")

    normalizados, grupos = preparar_agentes(df)
    tox = serie_tox(df)
    codigos = codigos_tox(tox)

    feat = criar_variaveis(grupos)
    feat["AGENTE_TOX"] = tox                                   # o tipo de uso em si, como variável
    feat["tox_x_ag1_grupo"] = tox + " | " + feat["ag1_grupo"]  # combinação uso x grupo do agente 1

    imprimir_resumo(df, grupos, feat)
    imprimir_por_tox(df, normalizados, tox, codigos)

    # Abas de distribuição por tipo de uso: uma aba por AGENTE_TOX (AGENTE_1/2/3 lado a lado)
    # e uma aba com a distribuição de GRUPOS (AGENTE_TOX nas colunas) para cada AGENTE_n.
    blocos_grupos = [(f"{c}: grupo x AGENTE_TOX (% sobre todas as linhas de cada código)",
                      tabela_grupos_por_tox(grupos[c], tox, codigos)) for c in COLUNAS_AGENTE]
    extras = [lambda wb, cod=cod: _escrever_blocos_horizontais(
                  wb, f"Nomes AGENTE_TOX {cod}", montar_blocos_nomes(normalizados, tox, cod))
              for cod in codigos]
    extras.append(lambda wb: _escrever_blocos_verticais(wb, "Grupos por AGENTE_TOX", blocos_grupos))
    ordem_abas = [f"Nomes AGENTE_TOX {cod}" for cod in codigos] + ["Grupos por AGENTE_TOX"]

    # Associação com o alvo (somente linhas com HOSPITAL válido)
    alvo = preparar_alvo(df)
    valido = alvo.notna()
    alvo_v, feat_v = alvo[valido].astype(int), feat[valido]
    ranking = ranking_variaveis(feat_v, alvo_v)
    blocos = [(f"Variavel: {r.variavel}  (V de Cramer = {r.cramers_v})", tabela_taxa(feat_v[r.variavel], alvo_v))
              for r in ranking.itertuples()]

    tabelas = {
        "Grupos por Agente": tabela_grupos_por_agente(grupos),
        "Associacao HOSPITAL": ranking,
        "Nao Classificados": tabela_nao_classificados(normalizados, grupos),
    }
    caminho = PASTA_SAIDA / "analise_agentes.xlsx"
    exportar(caminho, tabelas, blocos, extras=extras, ordem_abas=ordem_abas)
    print(f"\n[OK] Relatório salvo em: {caminho}")
    print(f"Linhas com {COLUNA_ALVO} válido: {len(alvo_v)} | taxa de positivos: {alvo_v.mean()*100:.2f}%")
    print("\nAssociação das variáveis derivadas com o alvo:")
    print(ranking.to_string(index=False))