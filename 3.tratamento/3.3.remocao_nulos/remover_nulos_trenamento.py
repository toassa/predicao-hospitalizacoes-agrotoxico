
from pathlib import Path
from datetime import datetime
import pandas as pd


# ==========================================
# CONFIGURAÇÕES
# ==========================================

PASTA_ENTRADA = Path("./3.2.divisao_treino_teste")
PASTA_SAIDA = Path("base_filtrada")

ARQUIVO_TREINO = PASTA_ENTRADA / "divisao_base/base_treino.csv"
ARQUIVO_TESTE = PASTA_ENTRADA / "divisao_base/base_teste.csv"

SEPARADOR = ";"
CODIFICACAO = "utf-8-sig"

LIMITE_NULOS = 10.0  # Percentual máximo permitido


def main():

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    # ==========================================
    # 1. LEITURA DAS BASES
    # ==========================================

    treino = pd.read_csv(
        ARQUIVO_TREINO,
        sep=SEPARADOR,
        encoding=CODIFICACAO,
        low_memory=False
    )

    teste = pd.read_csv(
        ARQUIVO_TESTE,
        sep=SEPARADOR,
        encoding=CODIFICACAO,
        low_memory=False
    )

    total_treino = len(treino)
    total_teste = len(teste)

    if total_treino == 0:
        raise ValueError("A base de treinamento está vazia.")

    if total_teste == 0:
        raise ValueError("A base de teste está vazia.")

    if set(treino.columns) != set(teste.columns):
        raise ValueError(
            "As bases de treino e teste possuem colunas diferentes."
        )

    # ==========================================
    # 2. ANÁLISE DE VALORES NULOS
    # ==========================================

    quantidade_nulos = treino.isna().sum()

    percentual_nulos = (
        quantidade_nulos / total_treino
    ) * 100

    # Criação do relatório de todas as colunas
    relatorio_nulos = pd.DataFrame({
        "coluna": treino.columns,
        "quantidade_nulos": quantidade_nulos.values,
        "percentual_nulos": percentual_nulos.values
    })

    relatorio_nulos = relatorio_nulos.sort_values(
        by="percentual_nulos",
        ascending=False
    )

    # ==========================================
    # 3. IDENTIFICAÇÃO DAS COLUNAS
    # ==========================================

    colunas_removidas = relatorio_nulos.loc[
        relatorio_nulos["percentual_nulos"] > LIMITE_NULOS
    ].copy()

    colunas_manter = relatorio_nulos.loc[
        relatorio_nulos["percentual_nulos"] <= LIMITE_NULOS,
        "coluna"
    ].tolist()

    lista_remocao = colunas_removidas["coluna"].tolist()

    # ==========================================
    # 4. REMOÇÃO DAS COLUNAS
    # ==========================================

    treino_filtrado = treino[colunas_manter].copy()
    teste_filtrado = teste[colunas_manter].copy()

    # ==========================================
    # 5. SALVAMENTO DAS BASES
    # ==========================================

    arquivo_treino_saida = (
        PASTA_SAIDA / "base_treino_filtrada.csv"
    )

    arquivo_teste_saida = (
        PASTA_SAIDA / "base_teste_filtrada.csv"
    )

    treino_filtrado.to_csv(
        arquivo_treino_saida,
        sep=SEPARADOR,
        index=False,
        encoding="utf-8-sig"
    )

    teste_filtrado.to_csv(
        arquivo_teste_saida,
        sep=SEPARADOR,
        index=False,
        encoding="utf-8-sig"
    )

    # ==========================================
    # 6. RELATÓRIO DAS COLUNAS REMOVIDAS
    # ==========================================

    arquivo_colunas_removidas = (
        PASTA_SAIDA / "colunas_removidas.csv"
    )

    colunas_removidas.to_csv(
        arquivo_colunas_removidas,
        sep=SEPARADOR,
        index=False,
        encoding="utf-8-sig"
    )

    # Também salva o percentual de nulos de todas
    # as colunas para fins de documentação.
    relatorio_nulos.to_csv(
        PASTA_SAIDA / "analise_valores_nulos.csv",
        sep=SEPARADOR,
        index=False,
        encoding="utf-8-sig"
    )

    # ==========================================
    # 7. GERAÇÃO DO LOG
    # ==========================================

    qtd_removidas = len(lista_remocao)
    qtd_mantidas = len(colunas_manter)

    log = [
        "=" * 70,
        "LOG DE REMOÇÃO DE COLUNAS POR VALORES NULOS",
        "=" * 70,
        f"Data/hora: {datetime.now():%d/%m/%Y %H:%M:%S}",
        "",
        "CONFIGURAÇÕES",
        "-" * 70,
        f"Limite de valores nulos: {LIMITE_NULOS:.2f}%",
        "Base utilizada para análise: Treinamento",
        "Critério de remoção: percentual de nulos > 10%",
        "",
        "BASES ORIGINAIS",
        "-" * 70,
        f"Registros de treino: {total_treino:,}",
        f"Registros de teste: {total_teste:,}",
        f"Colunas originais: {treino.shape[1]}",
        "",
        "RESULTADO DA ANÁLISE",
        "-" * 70,
        f"Colunas removidas: {qtd_removidas}",
        f"Colunas mantidas: {qtd_mantidas}",
        f"Total de colunas após filtragem: {treino_filtrado.shape[1]}",
        "",
        "COLUNAS REMOVIDAS",
        "-" * 70,
    ]

    if qtd_removidas > 0:

        for _, linha in colunas_removidas.iterrows():

            log.append(
                f"{linha['coluna']}: "
                f"{linha['quantidade_nulos']:,} nulos "
                f"({linha['percentual_nulos']:.2f}%)"
            )

    else:
        log.append("Nenhuma coluna ultrapassou o limite estabelecido.")

    log.extend([
        "",
        "BASES RESULTANTES",
        "-" * 70,
        f"Treino: {treino_filtrado.shape[0]:,} registros | "
        f"{treino_filtrado.shape[1]} colunas",
        f"Teste: {teste_filtrado.shape[0]:,} registros | "
        f"{teste_filtrado.shape[1]} colunas",
        "",
        "ARQUIVOS GERADOS",
        "-" * 70,
        str(arquivo_treino_saida.resolve()),
        str(arquivo_teste_saida.resolve()),
        str(arquivo_colunas_removidas.resolve()),
        str((PASTA_SAIDA / 'analise_valores_nulos.csv').resolve()),
        "",
        "VERIFICAÇÕES",
        "-" * 70,
        f"Mesmas colunas no treino e teste: "
        f"{list(treino_filtrado.columns) == list(teste_filtrado.columns)}",
        f"Registros de treino preservados: "
        f"{len(treino_filtrado) == total_treino}",
        f"Registros de teste preservados: "
        f"{len(teste_filtrado) == total_teste}",
        f"Valores nulos restantes no treino: "
        f"{int(treino_filtrado.isna().sum().sum()):,}",
        f"Valores nulos restantes no teste: "
        f"{int(teste_filtrado.isna().sum().sum()):,}",
        "",
        "Observação: a seleção das colunas foi realizada utilizando",
        "somente a base de treinamento. As mesmas colunas foram",
        "removidas da base de teste, sem imputação de valores.",
        "=" * 70
    ])

    arquivo_log = PASTA_SAIDA / "log_filtragem.txt"

    arquivo_log.write_text(
        "\n".join(log),
        encoding="utf-8"
    )

    print("\n".join(log))


if __name__ == "__main__":
    main()
