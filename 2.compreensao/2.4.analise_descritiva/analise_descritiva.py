"""
    pip install pandas numpy matplotlib openpyxl
"""

import textwrap
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # só salva arquivos; não abre janelas
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIGURAÇÃO
# ----------------------------------------------------------------------
CAMINHO_CSV = Path("../2.2.filtragem_agrotoxicos/dados_csv/base_filtrada.csv")
PASTA_SAIDA = Path("relatorios/descritiva")

COLUNA_ALVO = "HOSPITAL"
VALORES_ALVO_POSITIVO = [1]    # confirme no dicionário de dados
VALORES_ALVO_NEGATIVO = [2]    # outros códigos (ex: 9 = ignorado) ficam fora das taxas

# O script usa a primeira coluna da lista que existir na base. Ajuste aos seus nomes.
CANDIDATAS_IDADE = ["NU_IDADE_N", "IDADE"]
CANDIDATAS_DATA = ["DT_NOTIFIC", "DT_INVEST", "DT_SIN_PRI"]
CANDIDATAS_UF = ["SG_UF_NOT", "SG_UF", "ID_UF"]

# Colunas categóricas mostradas contra o alvo (as que não existirem são avisadas e puladas).
COLUNAS_CATEGORICAS = ["AGENTE_TOX", "CS_SEXO", "CS_RACA", "CS_ESCOL_N",
                       "CIRCUNSTAN", "TPEXP", "VIA_1", "CS_ZONA"]
MIN_CATEGORICAS = 4            # se existirem menos que isso, completa automaticamente
MAX_CATEGORICAS = 8            # máximo de colunas categóricas nos gráficos
MAX_CATEGORIAS_GRAFICO = 8     # categorias por gráfico (as demais viram "Outras categorias")
# Colunas ignoradas na escolha automática: identificadores e campos preenchidos DEPOIS/POR CAUSA da internação.
EXCLUIR_AUTO = {"DTINTERNA", "UF_HOSP", "MUN_HOSP", "EVOLUCAO", "DT_OBITO", "DT_ENCERRA",
                "CLASSI_FIN", "CRITERIO", "DIAG_CONF", "TP_NOT", "ID_AGRAVO"}

MIN_N = 30        # grupos com menos linhas válidas aparecem em cinza (instáveis)
TOP_UF = 15
N_COLUNAS_NULOS = 25
ANO_MIN = 2000    # datas fora de [ANO_MIN, ano atual + 1] são tratadas como erro de digitação

# Rótulos legíveis. Só os códigos que se tem certeza; complete com o dicionário de dados da sua base.
ROTULOS = {
    "HOSPITAL": {"1": "Sim", "2": "Não", "9": "Ignorado"},
    "CS_SEXO": {"M": "Masculino", "F": "Feminino", "I": "Ignorado"},
    "CS_RACA": {"1": "Branca", "2": "Preta", "3": "Amarela", "4": "Parda", "5": "Indígena", "9": "Ignorado"},
    "AGENTE_TOX": {"2": "Agrotóx. agrícola", "3": "Agrotóx. doméstico", "4": "Agrotóx. saúde pública"},
}

UF_IBGE = {11: "RO", 12: "AC", 13: "AM", 14: "RR", 15: "PA", 16: "AP", 17: "TO", 21: "MA", 22: "PI",
           23: "CE", 24: "RN", 25: "PB", 26: "PE", 27: "AL", 28: "SE", 29: "BA", 31: "MG", 32: "ES",
           33: "RJ", 35: "SP", 41: "PR", 42: "SC", 43: "RS", 50: "MS", 51: "MT", 52: "GO", 53: "DF"}

COR_OK, COR_FRACO, COR_REF = "#4C72B0", "#B8C0D0", "#C44E52"


# ----------------------------------------------------------------------
# 1. FUNÇÕES AUXILIARES
# ----------------------------------------------------------------------
def normalizar_codigos(serie: pd.Series) -> pd.Series:
    """2.0 -> '2'; texto sem espaços nas pontas; nulo continua nulo."""
    if pd.api.types.is_numeric_dtype(serie):
        nn = serie.dropna()
        if len(nn) and (nn % 1 == 0).all():
            serie = serie.astype("Int64")
    return serie.astype("string").str.strip()


def rotular(col: str, valor) -> str:
    if pd.isna(valor):
        return "(vazio)"
    v = str(valor)
    nome = ROTULOS.get(col, {}).get(v)
    return f"{v} - {nome}" if nome else v


def categorias(df: pd.DataFrame, col: str) -> pd.Series:
    return normalizar_codigos(df[col]).map(lambda v: rotular(col, v))


def primeira_existente(df: pd.DataFrame, candidatas: list):
    return next((c for c in candidatas if c in df.columns), None)


def preparar_alvo(df: pd.DataFrame) -> pd.Series:
    """1 = hospitalizado, 0 = não; NaN = ignorado/vazio (fica fora das taxas)."""
    codigo = pd.to_numeric(df[COLUNA_ALVO], errors="coerce")
    alvo = pd.Series(np.nan, index=df.index)
    alvo[codigo.isin(VALORES_ALVO_POSITIVO)] = 1
    alvo[codigo.isin(VALORES_ALVO_NEGATIVO)] = 0
    return alvo


def wilson_ic(k: float, n: float, z: float = 1.96):
    """Intervalo de confiança de Wilson (95%) para uma proporção k/n. Retorna (inferior, superior) em 0-1."""
    if n == 0:
        return np.nan, np.nan
    p = k / n
    den = 1 + z ** 2 / n
    centro = (p + z ** 2 / (2 * n)) / den
    meia = z * np.sqrt(p * (1 - p) / n + z ** 2 / (4 * n ** 2)) / den
    return centro - meia, centro + meia


def taxa_por_categoria(cat: pd.Series, alvo: pd.Series, max_cat=None, ordem=None,
                       rotulo_outras="Outras categorias") -> pd.DataFrame:
    """Para cada categoria: n válido, nº de hospitalizados, taxa (%) e IC 95% (Wilson)."""
    base = pd.DataFrame({"cat": cat.values, "alvo": alvo.values}).dropna(subset=["alvo"])
    if ordem is None:
        contagem = base["cat"].value_counts()
        if max_cat is not None and len(contagem) > max_cat:
            manter = list(contagem.head(max_cat).index)
            base["cat"] = np.where(base["cat"].isin(manter), base["cat"], rotulo_outras)
        ordem = list(base["cat"].value_counts().index)
        if rotulo_outras in ordem:
            ordem = [o for o in ordem if o != rotulo_outras] + [rotulo_outras]

    g = base.groupby("cat")["alvo"].agg(n="size", positivos="sum")
    g = g.reindex([o for o in ordem if o in g.index])
    g["taxa"] = g["positivos"] / g["n"] * 100
    ic = [wilson_ic(k, n) for k, n in zip(g["positivos"], g["n"])]
    g["ic_inf"] = [a * 100 for a, _ in ic]
    g["ic_sup"] = [b * 100 for _, b in ic]
    g["pct_do_total"] = g["n"] / g["n"].sum() * 100
    return g.rename_axis("categoria")


def fmt_n(n) -> str:
    return f"{int(n):,}".replace(",", ".")


def barras_taxa(ax, tab: pd.DataFrame, taxa_geral: float, titulo: str) -> None:
    """Barras horizontais da taxa por categoria, com IC 95%, n e linha de referência (taxa geral)."""
    y = np.arange(len(tab))[::-1]
    cores = [COR_OK if n >= MIN_N else COR_FRACO for n in tab["n"]]
    erro = [(tab["taxa"] - tab["ic_inf"]).values, (tab["ic_sup"] - tab["taxa"]).values]
    ax.barh(y, tab["taxa"], color=cores, xerr=erro, ecolor="#333333", capsize=2, height=0.65)
    ax.axvline(taxa_geral, color=COR_REF, ls="--", lw=1)
    ax.set_yticks(y)
    ax.set_yticklabels([str(i)[:30] for i in tab.index], fontsize=8)
    xmax = max(10, tab["ic_sup"].max() * 1.5)
    ax.set_xlim(0, xmax)
    ax.set_xticks([t for t in ax.get_xticks() if 0 <= t <= 100])
    for yi, n, t, ic in zip(y, tab["n"], tab["taxa"], tab["ic_sup"]):
        ax.text(ic + xmax * 0.01, yi, f"{t:.0f}%  (n={fmt_n(n)})", va="center", fontsize=7.5)
    ax.set_title(titulo, fontsize=10, fontweight="bold", loc="left")
    ax.set_xlabel("% hospitalizados", fontsize=8)
    ax.tick_params(axis="x", labelsize=8)


def salvar(fig, nome: str, como_ler: str, pdf: PdfPages) -> None:
    """Rodapé 'Como ler' + PNG + página no PDF."""
    fig.tight_layout(rect=[0, 0.10, 1, 0.97])
    fig.text(0.01, 0.01, textwrap.fill("Como ler: " + como_ler, 150), fontsize=8.5, va="bottom", color="#333333")
    fig.savefig(PASTA_SAIDA / nome, dpi=130, bbox_inches="tight")
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)
    print(f"   [ok] {nome}")


# ----------------------------------------------------------------------
# 2. DECODIFICAÇÕES (idade, data, UF)
# ----------------------------------------------------------------------
def idade_em_anos(serie: pd.Series) -> pd.Series:
    """
    Formato SINAN (4 dígitos): 1º dígito = unidade (1 hora, 2 dia, 3 mês, 4 ano); demais = valor.
    Ex: 4025 = 25 anos; 3008 = 8 meses. Se a coluna já estiver em anos, é usada como está.
    Confirme o formato no dicionário de dados da sua base.
    """
    x = pd.to_numeric(serie, errors="coerce")
    if (x.dropna() >= 1000).mean() > 0.5:
        unidade, valor = x // 1000, x % 1000
        anos = pd.Series(np.nan, index=x.index)
        anos[unidade == 4] = valor[unidade == 4]
        anos[unidade == 3] = valor[unidade == 3] / 12
        anos[unidade == 2] = valor[unidade == 2] / 365
        anos[unidade == 1] = valor[unidade == 1] / 8760
    else:
        anos = x
    anos[(anos < 0) | (anos > 120)] = np.nan
    return anos


def parse_datas(serie: pd.Series) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        d = pd.to_datetime(serie, errors="coerce")
        if d.notna().mean() < 0.5:
            d2 = pd.to_datetime(serie, errors="coerce", dayfirst=True)
            if d2.notna().sum() > d.notna().sum():
                d = d2
    return d


def datas_validas(serie: pd.Series):
    """Converte para data e descarta (vira NaT) o que estiver fora de [ANO_MIN, ano atual + 1]. Retorna (datas, n_descartadas)."""
    d = parse_datas(serie)
    fora = d.notna() & ((d.dt.year < ANO_MIN) | (d.dt.year > pd.Timestamp.today().year + 1))
    return d.where(~fora), int(fora.sum())


def uf_sigla(serie: pd.Series) -> pd.Series:
    s = serie.astype("string").str.strip()
    cod = pd.to_numeric(s, errors="coerce")
    sigla = cod.map(lambda v: UF_IBGE.get(int(v)) if pd.notna(v) else None)
    return sigla.where(sigla.notna(), s.astype(object))


# ----------------------------------------------------------------------
# 3. GRÁFICOS
# ----------------------------------------------------------------------
def grafico_alvo(df, alvo, pdf, tabelas) -> None:
    cod = normalizar_codigos(df[COLUNA_ALVO]).fillna("(vazio)")
    cont = cod.value_counts()
    usados = {str(v) for v in VALORES_ALVO_POSITIVO + VALORES_ALVO_NEGATIVO}
    rotulos = [c if c == "(vazio)" else rotular(COLUNA_ALVO, c) for c in cont.index]

    fig, ax = plt.subplots(figsize=(8, 4.8))
    cores = [COR_OK if c in usados else COR_FRACO for c in cont.index]
    ax.bar(rotulos, cont.values, color=cores)
    for i, n in enumerate(cont.values):
        ax.text(i, n, f"{fmt_n(n)}\n({n / len(df) * 100:.1f}%)", ha="center", va="bottom", fontsize=9)
    ax.set_ylim(0, cont.max() * 1.2)
    ax.set_ylabel("nº de casos")
    validos = int(alvo.notna().sum())
    ax.set_title(f"1. Alvo: {COLUNA_ALVO}  |  taxa de hospitalização (entre válidos): "
                 f"{alvo.mean() * 100:.1f}%  (n válido = {fmt_n(validos)})", fontsize=10, fontweight="bold", loc="left")
    tabelas["1_Alvo"] = pd.DataFrame({"codigo": cont.index, "rotulo": rotulos, "n": cont.values,
                                      "pct": (cont.values / len(df) * 100).round(2),
                                      "usado_nas_taxas": [c in usados for c in cont.index]})
    salvar(fig, "01_alvo.png",
           "cada barra é um código do campo HOSPITAL. Azul = códigos usados nas taxas (Sim/Não); cinza = ignorado/vazio, "
           "que ficam fora de todos os cálculos de taxa. A taxa de hospitalização geral é a referência (linha tracejada) "
           "dos gráficos seguintes.", pdf)


def grafico_nulos(df, pdf, tabelas) -> None:
    pct = (df.isna().mean() * 100).sort_values(ascending=False)
    tabelas["2_Vazios"] = pct.round(2).rename("pct_vazio").rename_axis("coluna").reset_index()
    top = pct.head(N_COLUNAS_NULOS)[::-1]

    fig, ax = plt.subplots(figsize=(9, max(4.5, 0.3 * len(top) + 1.8)))
    cores = [COR_REF if v > 90 else "#DD8452" if v > 40 else COR_OK for v in top.values]
    ax.barh(top.index, top.values, color=cores)
    for i, v in enumerate(top.values):
        ax.text(v + 0.8, i, f"{v:.1f}%", va="center", fontsize=8)
    ax.axvline(40, color="#555555", ls=":", lw=1)
    ax.axvline(90, color="#555555", ls=":", lw=1)
    ax.set_xlim(0, 112)
    ax.set_xlabel("% de linhas vazias")
    ax.tick_params(axis="y", labelsize=8)
    ax.set_title(f"2. Qualidade: colunas com mais valores vazios  "
                 f"({int((pct > 40).sum())} colunas > 40%, {int((pct > 90).sum())} > 90%)",
                 fontsize=10, fontweight="bold", loc="left")
    salvar(fig, "02_vazios.png",
           f"mostra as {len(top)} colunas mais vazias (a tabela completa está no Excel). Laranja > 40%, vermelho > 90%. "
           "Vazio nem sempre é erro: pode ser 'não se aplica' (ex: campos de internação em quem não foi internado). "
           "Pergunte sempre POR QUE está vazio: se o vazio depende do próprio alvo, a coluna não serve para prever.", pdf)


def grafico_idade(df, alvo, pdf, tabelas) -> None:
    col = primeira_existente(df, CANDIDATAS_IDADE)
    if col is None:
        print(f"   [pulado] idade: nenhuma coluna de {CANDIDATAS_IDADE} encontrada")
        return
    anos = idade_em_anos(df[col])
    rotulos = ["0-11", "12-17", "18-29", "30-44", "45-59", "60+"]
    faixa = pd.cut(anos, bins=[0, 12, 18, 30, 45, 60, 121], right=False, labels=rotulos)
    faixa_txt = faixa.astype(object).where(faixa.notna(), "(vazio)")
    tab = taxa_por_categoria(faixa_txt, alvo, ordem=rotulos + ["(vazio)"])

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5))
    ax1.hist(anos.dropna(), bins=np.arange(0, 105, 5), color=COR_OK, edgecolor="white")
    if anos.notna().any():
        ax1.axvline(anos.median(), color=COR_REF, ls="--", lw=1)
        ax1.text(anos.median(), ax1.get_ylim()[1] * 0.95, f" mediana = {anos.median():.0f} anos",
                 color=COR_REF, fontsize=9, va="top")
    ax1.set_xlabel("idade (anos)")
    ax1.set_ylabel("nº de casos")
    ax1.set_title(f"3a. Distribuição da idade ({col}; {int(anos.isna().sum())} sem idade válida)",
                  fontsize=10, fontweight="bold", loc="left")
    barras_taxa(ax2, tab, alvo.mean() * 100, "3b. Taxa de hospitalização por faixa etária")
    tabelas["3_Idade"] = tab.reset_index().round(2)
    salvar(fig, "03_idade.png",
           "à esquerda, quantos casos há em cada idade (barras de 5 anos). À direita, a % de hospitalizados em cada faixa; "
           "o traço preto é o intervalo de confiança de 95% e a linha vermelha tracejada é a taxa geral. "
           "Faixa acima/abaixo da linha com intervalo que não a cruza difere do geral. Barras cinza têm poucos casos.", pdf)


def grafico_tempo(df, alvo, pdf, tabelas) -> None:
    col = primeira_existente(df, CANDIDATAS_DATA)
    if col is None:
        print(f"   [pulado] tempo: nenhuma coluna de {CANDIDATAS_DATA} encontrada")
        return
    datas, invalidas = datas_validas(df[col])

    t = pd.DataFrame({"ym": datas.dt.to_period("M"), "mes": datas.dt.month, "ano": datas.dt.year, "alvo": alvo})
    t = t.dropna(subset=["ym"])
    if t.empty:
        print(f"   [pulado] tempo: nenhuma data válida em {col}")
        return

    def agregar(chave):
        g = t.groupby(chave).agg(n=("alvo", "size"), validos=("alvo", "count"), positivos=("alvo", "sum"))
        g["taxa"] = np.where(g["validos"] >= MIN_N, g["positivos"] / g["validos"] * 100, np.nan)
        return g

    mensal = agregar("ym")
    sazonal = None
    if t["ano"].nunique() >= 2:
        sazonal = agregar("mes")
        # meses do calendário aparecem em quantidades diferentes de anos (ex: jan-mai em 2 anos, jun-dez em 1):
        # por isso a barra é a MÉDIA de casos por ano, e não a soma.
        sazonal["media_casos_por_ano"] = sazonal["n"] / t.groupby("mes")["ano"].nunique()

    def painel(ax, g, rotulos, titulo, xlabel, coluna="n", ylabel="nº de casos"):
        ax.bar(range(len(g)), g[coluna], color=COR_OK)
        ax.set_xticks(range(len(g)))
        passo = max(1, len(g) // 18)
        ax.set_xticklabels([r if i % passo == 0 else "" for i, r in enumerate(rotulos)], rotation=45, ha="right", fontsize=8)
        ax.set_ylabel(ylabel)
        ax.set_xlabel(xlabel, fontsize=8)
        ax2 = ax.twinx()
        ax2.plot(range(len(g)), g["taxa"], color=COR_REF, marker="o", ms=4)
        ax2.set_ylim(0, 100)
        ax2.set_ylabel("% hospitalizados", color=COR_REF)
        ax.set_title(titulo, fontsize=10, fontweight="bold", loc="left")

    ncols = 2 if sazonal is not None else 1
    fig, axes = plt.subplots(1, ncols, figsize=(13 if ncols == 2 else 9, 5))
    axes = np.atleast_1d(axes)
    painel(axes[0], mensal, [str(p) for p in mensal.index], f"4a. Casos e taxa por mês ({col})", "mês da data")
    if sazonal is not None:
        painel(axes[1], sazonal, [str(int(m)) for m in sazonal.index],
               "4b. Sazonalidade: média de casos por mês do calendário", "mês do ano",
               coluna="media_casos_por_ano", ylabel="média de casos por ano")
    tabelas["4_Tempo_mensal"] = mensal.reset_index().astype({"ym": str}).round(2)
    if sazonal is not None:
        tabelas["4_Tempo_sazonal"] = sazonal.reset_index().round(2)

    nota = "" if sazonal is not None else " Só há um ano de dados, então a sazonalidade (padrão que se repete a cada ano) não pode ser avaliada."
    salvar(fig, "04_tempo.png",
           f"barras azuis = nº de casos (eixo esquerdo); linha vermelha = % hospitalizados (eixo direito, só em meses com "
           f"{MIN_N}+ casos válidos). Em 4b, a barra é a média de casos por ano em cada mês do calendário (um mês pode ter "
           f"mais anos de dados que outro). Datas inválidas descartadas: {invalidas}. O mês mais recente costuma estar "
           f"incompleto (notificações ainda chegando).{nota}", pdf)


def grafico_uf(df, alvo, pdf, tabelas) -> None:
    col = primeira_existente(df, CANDIDATAS_UF)
    if col is None:
        print(f"   [pulado] UF: nenhuma coluna de {CANDIDATAS_UF} encontrada")
        return
    uf = uf_sigla(df[col]).fillna("(vazio)")
    if uf.nunique() < 2:
        print(f"   [pulado] UF: {col} tem um único valor ({uf.iloc[0]})")
        return
    tab = taxa_por_categoria(uf, alvo, max_cat=TOP_UF, rotulo_outras="Outras UFs")
    fig, ax = plt.subplots(figsize=(9, max(4.5, 0.38 * len(tab) + 2)))
    barras_taxa(ax, tab, alvo.mean() * 100, f"5. Taxa de hospitalização por UF ({col}; {TOP_UF} UFs com mais casos)")
    tabelas["5_UF"] = tab.reset_index().round(2)
    salvar(fig, "05_uf.png",
           "cada barra é uma UF, ordenada por nº de casos (n). Compare a barra com a linha vermelha (taxa geral) levando em "
           "conta o intervalo de confiança. Diferenças entre UFs podem refletir a prática de notificar/internar, e não só a gravidade.", pdf)


def escolher_categoricas(df):
    existentes = [c for c in COLUNAS_CATEGORICAS if c in df.columns and df[c].nunique(dropna=True) >= 2]
    faltando = [c for c in COLUNAS_CATEGORICAS if c not in df.columns]
    escolhidas = list(existentes)
    if len(escolhidas) < MIN_CATEGORICAS:
        extras = []
        for c in df.columns:
            if c in escolhidas or c == COLUNA_ALVO or c in EXCLUIR_AUTO:
                continue
            if c.upper().startswith(("DT", "DATA", "ID_", "NU_", "SG_", "CO_")):
                continue
            n_un, nulos = df[c].nunique(dropna=True), df[c].isna().mean()
            if 2 <= n_un <= 15 and nulos < 0.5:
                extras.append((nulos, c))
        escolhidas += [c for _, c in sorted(extras)[:MAX_CATEGORICAS - len(escolhidas)]]
    return escolhidas[:MAX_CATEGORICAS], faltando


def graficos_categoricas(df, alvo, colunas, pdf, tabelas) -> None:
    taxa_geral = alvo.mean() * 100
    for pag, i in enumerate(range(0, len(colunas), 4), start=1):
        grupo = colunas[i:i + 4]
        fig, axes = plt.subplots(2, 2, figsize=(14, 9))
        axes = axes.ravel()
        for ax, col in zip(axes, grupo):
            tab = taxa_por_categoria(categorias(df, col), alvo, max_cat=MAX_CATEGORIAS_GRAFICO)
            barras_taxa(ax, tab, taxa_geral, f"{col}  ({df[col].isna().mean() * 100:.0f}% vazio)")
            tabelas[f"6_{col}"[:31]] = tab.reset_index().round(2)
        for ax in axes[len(grupo):]:
            ax.axis("off")
        fig.suptitle(f"6. Taxa de hospitalização por categoria (página {pag})", fontsize=11, fontweight="bold", x=0.01, ha="left")
        salvar(fig, f"06_categoricas_{pag}.png",
               "uma janela por coluna; cada barra é uma categoria (n = nº de casos válidos). Linha vermelha tracejada = taxa geral; "
               "traço preto = intervalo de confiança de 95%; barras cinza têm menos de "
               f"{MIN_N} casos (instáveis). Os códigos aparecem como no banco; complete ROTULOS com o dicionário de dados.", pdf)


# ----------------------------------------------------------------------
# 4. RESUMO NO CONSOLE E EXPORTAÇÃO DAS TABELAS
# ----------------------------------------------------------------------
def resumo_console(df, alvo) -> None:
    print("=" * 70)
    print("ANÁLISE DESCRITIVA — base filtrada")
    print("=" * 70)
    print(f"Arquivo : {CAMINHO_CSV}")
    print(f"Linhas  : {fmt_n(len(df))} | Colunas: {df.shape[1]}")
    col_data = primeira_existente(df, CANDIDATAS_DATA)
    if col_data:
        d, descartadas = datas_validas(df[col_data])
        d = d.dropna()
        if len(d):
            extra = f" | {descartadas} data(s) inválida(s) descartada(s)" if descartadas else ""
            print(f"Período : {d.min().date()} a {d.max().date()} ({col_data}){extra}")
    validos = int(alvo.notna().sum())
    print(f"Alvo    : {COLUNA_ALVO} válido em {fmt_n(validos)} linhas "
          f"({validos / len(df) * 100:.1f}%); fora das taxas: {fmt_n(len(df) - validos)}")
    print(f"Taxa de hospitalização (entre válidos): {alvo.mean() * 100:.1f}%")
    col_idade = primeira_existente(df, CANDIDATAS_IDADE)
    if col_idade:
        anos = idade_em_anos(df[col_idade]).dropna()
        if len(anos):
            print(f"Idade   : mediana {anos.median():.0f} anos (mín {anos.min():.0f}, máx {anos.max():.0f})")
    print("-" * 70)


def exportar_tabelas(tabelas: dict, caminho: Path) -> None:
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        for nome, tab in tabelas.items():
            nome = nome[:31]
            tab.to_excel(w, sheet_name=nome, index=False)
            ws = w.sheets[nome]
            ws.freeze_panes = "A2"
            for coluna in ws.columns:
                maior = max((len(str(c.value)) for c in coluna if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(coluna[0].column)].width = min(maior + 2, 45)
    print(f"   [ok] {caminho.name}")


# ----------------------------------------------------------------------
# 5. EXECUÇÃO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(CAMINHO_CSV, sep=";", encoding="latin-1", low_memory=False)
    if COLUNA_ALVO not in df.columns:
        raise KeyError(f"Coluna alvo '{COLUNA_ALVO}' não encontrada. Colunas: {list(df.columns)}")

    alvo = preparar_alvo(df)
    resumo_console(df, alvo)

    tabelas: dict = {}
    print("Gerando gráficos:")
    with PdfPages(PASTA_SAIDA / "relatorio_descritivo.pdf") as pdf:
        grafico_alvo(df, alvo, pdf, tabelas)
        grafico_nulos(df, pdf, tabelas)
        grafico_idade(df, alvo, pdf, tabelas)
        grafico_tempo(df, alvo, pdf, tabelas)
        grafico_uf(df, alvo, pdf, tabelas)

        colunas, faltando = escolher_categoricas(df)
        if faltando:
            print(f"   [aviso] colunas de COLUNAS_CATEGORICAS não encontradas na base: {faltando}")
        print(f"   colunas categóricas usadas: {colunas}")
        graficos_categoricas(df, alvo, colunas, pdf, tabelas)

    exportar_tabelas(tabelas, PASTA_SAIDA / "tabelas_descritivas.xlsx")
    print(f"\n[OK] Tudo salvo em: {PASTA_SAIDA}/")