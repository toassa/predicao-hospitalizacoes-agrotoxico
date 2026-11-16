
from pathlib import Path
from datetime import datetime
import pandas as pd

ARQUIVO_ENTRADA = Path("../2.compreensao/2.2.filtragem_agrotoxicos/dados_csv/base_filtrada.csv")
PASTA_SAIDA = Path("base_filtrada")

SEPARADOR = ";"
CODIFICACAO = "latin-1"

COLUNA_ALVO = "HOSPITAL"
VALOR_IGNORADO = "9.0"


def main():

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    # Carregamento da base original
    base = pd.read_csv(
        ARQUIVO_ENTRADA,
        sep=SEPARADOR,
        encoding=CODIFICACAO,
        low_memory=False
    )

    if COLUNA_ALVO not in base.columns:
        raise ValueError(
            f"Coluna {COLUNA_ALVO} não encontrada."
        )

    total_original = len(base)

    # Identificação dos valores nulos e ignorados
    hospital = (
        base[COLUNA_ALVO]
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
    )

    mascara_nulos = hospital.isna()
    mascara_ignorados = hospital.eq(VALOR_IGNORADO).fillna(False)

    qtd_nulos = int(mascara_nulos.sum())
    qtd_ignorados = int(mascara_ignorados.sum())

    # Exclusão dos registros
    mascara_remover = mascara_nulos | mascara_ignorados

    base_filtrada = base.loc[~mascara_remover].copy()

    total_final = len(base_filtrada)
    total_removido = total_original - total_final

    # Verificação dos valores restantes
    distribuicao = (
        base_filtrada[COLUNA_ALVO]
        .value_counts(dropna=False)
    )

    # Salvamento da nova base
    arquivo_saida = PASTA_SAIDA / "base_sem_nulos_ignorados.csv"

    base_filtrada.to_csv(
        arquivo_saida,
        sep=SEPARADOR,
        index=False,
        encoding="utf-8-sig"
    )

    # ==========================================
    # GERAÇÃO DO LOG
    # ==========================================

    percentual_removido = (
        total_removido / total_original * 100
    )

    percentual_final = (
        total_final / total_original * 100
    )

    log = [
        "=" * 65,
        "LOG DE FILTRAGEM DA VARIÁVEL HOSPITAL",
        "=" * 65,
        f"Data/hora: {datetime.now():%d/%m/%Y %H:%M:%S}",
        f"Arquivo original: {ARQUIVO_ENTRADA.resolve()}",
        "",
        "RESUMO DA BASE ORIGINAL",
        "-" * 65,
        f"Total de registros: {total_original:,}",
        f"Total de colunas: {base.shape[1]}",
        "",
        "REGISTROS REMOVIDOS",
        "-" * 65,
        f"HOSPITAL nulo: {qtd_nulos:,} "
        f"({qtd_nulos / total_original:.2%})",
        f"HOSPITAL = 9 (Ignorado): {qtd_ignorados:,} "
        f"({qtd_ignorados / total_original:.2%})",
        f"Total removido: {total_removido:,} "
        f"({percentual_removido:.2f}%)",
        "",
        "BASE RESULTANTE",
        "-" * 65,
        f"Registros restantes: {total_final:,}",
        f"Percentual preservado: {percentual_final:.2f}%",
        "",
        "DISTRIBUIÇÃO DA VARIÁVEL HOSPITAL",
        "-" * 65,
    ]

    for classe, quantidade in distribuicao.items():

        percentual = quantidade / total_final * 100

        descricao = {
            "1": "Sim",
            "2": "Não"
        }.get(str(classe), "Outro valor")

        log.append(
            f"{classe} ({descricao}): {quantidade:,} "
            f"({percentual:.2f}%)"
        )

    log.extend([
        "",
        "VERIFICAÇÕES",
        "-" * 65,
        f"Valores nulos restantes em HOSPITAL: "
        f"{base_filtrada[COLUNA_ALVO].isna().sum()}",
        f"Valores 9 restantes em HOSPITAL: "
        f"{(hospital.loc[base_filtrada.index] == VALOR_IGNORADO).sum()}",
        f"Colunas preservadas: "
        f"{base.shape[1] == base_filtrada.shape[1]}",
        f"Valores ausentes nas demais colunas: "
        f"{base_filtrada.drop(columns=[COLUNA_ALVO]).isna().sum().sum():,}",
        "",
        "ARQUIVO GERADO",
        "-" * 65,
        str(arquivo_saida.resolve()),
        "",
        "Observação: nenhuma imputação ou transformação foi realizada.",
        "Os valores ausentes das demais variáveis foram preservados.",
        "=" * 65
    ])

    arquivo_log = PASTA_SAIDA / "log_filtragem.txt"

    arquivo_log.write_text(
        "\n".join(log),
        encoding="utf-8"
    )

    print("\n".join(log))


if __name__ == "__main__":
    main()
