"""
Tabela-dicionário das colunas da base filtrada
==============================================
Gera um Excel com UMA linha por coluna da base: % de vazios, tipo, nº de valores
distintos e 4 colunas para você preencher (3 com lista suspensa).

    pip install pandas openpyxl
"""

import shutil
from pathlib import Path

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------
# CONFIGURAÇÃO
# ----------------------------------------------------------------------
CAMINHO_CSV = Path("../2.2.filtragem_agrotoxicos/dados_csv/base_filtrada.csv")
PASTA_SAIDA = Path("relatorios/descritiva")
CAMINHO_EXCEL = PASTA_SAIDA / "tabelas_descritivas.xlsx"
ABA_TABELA = "2_Vazios"

ORDEM_VAZIOS_CRESCENTE = False   # False = mais vazias primeiro; True = menos vazias primeiro

CATEGORIAS_ARTIGO = [
    "Demographic information (gender, age)",
    "Environmental parameters (temperature, humidity, pesticide concentration)",
    "Symptoms (headache, nausea, dizziness)",
    "Pesticide exposure history (type, duration, frequency)",
]
OPCOES_IMPORTANCIA = ["ALTA", "MEDIA", "BAIXA"]
OPCOES_SIM_NAO = ["SIM", "NAO"]
# Colunas que vêm vazias para você preencher (as 3 primeiras com lista suspensa)
COLUNAS_MANUAIS = ["categoria_artigo", "importancia", "presente_dicionario", "descricao_dicionario"]


# ----------------------------------------------------------------------
# 1. TABELA
# ----------------------------------------------------------------------
def ler_anotacoes_anteriores(caminho: Path) -> pd.DataFrame:
    """
    Se o Excel já existir, recupera o que você preencheu à mão (por nome de coluna),
    para que rodar o script de novo NÃO apague suas anotações.
    """
    if not caminho.exists():
        return pd.DataFrame()
    try:
        antigo = pd.read_excel(caminho, sheet_name=ABA_TABELA)
    except Exception as erro:
        print(f"   [aviso] não consegui ler as anotações anteriores de {caminho.name} ({erro}). "
              f"Uma cópia de segurança será feita antes de sobrescrever.")
        return pd.DataFrame()
    if "coluna" not in antigo.columns:
        return pd.DataFrame()
    manuais = [c for c in COLUNAS_MANUAIS if c in antigo.columns]
    return antigo.drop_duplicates("coluna").set_index("coluna")[manuais]


def montar_tabela(df: pd.DataFrame, anotacoes: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por coluna da base: % de vazios, tipo, nº de valores distintos + colunas manuais (vazias)."""
    tab = pd.DataFrame({
        "coluna": list(df.columns),
        "pct_vazio": (df.isna().mean() * 100).round(2).values,
        "tipo": [str(t) for t in df.dtypes],
        "n_distintos": [int(df[c].nunique(dropna=True)) for c in df.columns],
    })
    tab = tab.sort_values(["pct_vazio", "coluna"], ascending=[ORDEM_VAZIOS_CRESCENTE, True]).reset_index(drop=True)
    for c in COLUNAS_MANUAIS:
        tab[c] = tab["coluna"].map(anotacoes[c]) if c in anotacoes.columns else None
    return tab


# ----------------------------------------------------------------------
# 2. EXPORTAÇÃO PARA EXCEL
# ----------------------------------------------------------------------
def adicionar_dropdowns(wb, ws, tab: pd.DataFrame) -> None:
    """
    Cria a aba 'Listas' (opções) e aplica listas suspensas nas colunas manuais.
    As opções ficam numa aba à parte porque a de categoria_artigo contém vírgulas, o que quebraria
    uma lista escrita direto na validação. Para mudar as opções, edite a aba 'Listas'.
    """
    listas = wb.create_sheet("Listas")
    opcoes = {"categoria_artigo": CATEGORIAS_ARTIGO,
              "importancia": OPCOES_IMPORTANCIA,
              "presente_dicionario": OPCOES_SIM_NAO}
    for j, (nome, itens) in enumerate(opcoes.items(), start=1):
        listas.cell(row=1, column=j, value=nome)
        for i, item in enumerate(itens, start=2):
            listas.cell(row=i, column=j, value=item)
        listas.column_dimensions[get_column_letter(j)].width = 70 if nome == "categoria_artigo" else 22

    colunas = list(tab.columns)
    ultima_linha = max(len(tab) + 1, 2)
    for j, (nome, itens) in enumerate(opcoes.items(), start=1):
        letra_lista = get_column_letter(j)
        dv = DataValidation(type="list", formula1=f"=Listas!${letra_lista}$2:${letra_lista}${len(itens) + 1}",
                            allow_blank=True, showErrorMessage=True,
                            errorTitle="Valor inválido", error="Escolha uma das opções da lista.")
        ws.add_data_validation(dv)
        letra = get_column_letter(colunas.index(nome) + 1)
        dv.add(f"{letra}2:{letra}{ultima_linha}")

    # destaque no cabeçalho das colunas manuais + larguras
    amarelo = PatternFill("solid", start_color="FFF2CC")
    larguras = {"categoria_artigo": 62, "importancia": 14, "presente_dicionario": 20, "descricao_dicionario": 50}
    for nome in COLUNAS_MANUAIS:
        letra = get_column_letter(colunas.index(nome) + 1)
        ws[f"{letra}1"].fill = amarelo
        ws.column_dimensions[letra].width = larguras[nome]


def exportar(tab: pd.DataFrame, caminho: Path) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    if caminho.exists():   # cópia de segurança do que existia (inclui suas anotações manuais)
        shutil.copy2(caminho, caminho.with_suffix(".bak.xlsx"))
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        tab.to_excel(w, sheet_name=ABA_TABELA, index=False)
        ws = w.sheets[ABA_TABELA]
        ws.freeze_panes = "A2"
        for coluna in ws.columns:
            maior = max((len(str(c.value)) for c in coluna if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(coluna[0].column)].width = min(maior + 2, 45)
        adicionar_dropdowns(w.book, ws, tab)


# ----------------------------------------------------------------------
# 3. EXECUÇÃO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    df = pd.read_csv(CAMINHO_CSV, sep=";", encoding="latin-1", low_memory=False)
    linhas = f"{len(df):,}".replace(",", ".")
    print(f"Base: {CAMINHO_CSV} | {linhas} linhas, {df.shape[1]} colunas")

    anotacoes = ler_anotacoes_anteriores(CAMINHO_EXCEL)
    if len(anotacoes):
        preenchidas = int(anotacoes.notna().any(axis=1).sum())
        print(f"Anotações anteriores recuperadas: {preenchidas} coluna(s) com algum campo preenchido")

    tabela = montar_tabela(df, anotacoes)
    exportar(tabela, CAMINHO_EXCEL)
    print(f"[OK] Tabela salva em: {CAMINHO_EXCEL}")