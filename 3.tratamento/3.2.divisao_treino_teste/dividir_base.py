from pathlib import Path
from datetime import datetime
import pandas as pd
from sklearn.model_selection import train_test_split

ARQUIVO_ENTRADA = Path("./3.1.ajustar_alvo/base_filtrada/base_sem_nulos_ignorados.csv")
PASTA_SAIDA = Path("divisao_base")

# Configurações para leitura do CSV.
SEPARADOR = ";"               # Use "," se esse for o separador do arquivo
CODIFICACAO = "latin-1"       # Ex.: "utf-8" ou "utf-8-sig"

COLUNA_ALVO = "HOSPITAL"      # Altere para o nome exato da variável-alvo.
                              # Use None para não estratificar a divisão.

PROPORCAO_TESTE = 0.30
SEMENTE_ALEATORIA = 42        # Garante que a divisão possa ser reproduzida.
ESTRATIFICAR = True           # Mantém, quando possível, a proporção das classes.


def formatar_distribuicao(serie, total):
    """Retorna a distribuição de uma variável em quantidade e percentual."""
    contagens = serie.value_counts(dropna=False)
    linhas = []
    for classe, quantidade in contagens.items():
        percentual = (quantidade / total * 100) if total else 0
        linhas.append(
            f"  - {classe}: {quantidade:,} registros ({percentual:.2f}%)"
        )
    return "\n".join(linhas) if linhas else "  - Sem registros"


def main():
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    if not ARQUIVO_ENTRADA.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {ARQUIVO_ENTRADA.resolve()}\n"
            "Altere ARQUIVO_ENTRADA no início do script."
        )

    # Leitura da base original.
    base = pd.read_csv(
        ARQUIVO_ENTRADA,
        sep=SEPARADOR,
        encoding=CODIFICACAO,
        low_memory=False
    )

    total_original = len(base)

    if total_original == 0:
        raise ValueError("A base de dados está vazia.")

    if ESTRATIFICAR and COLUNA_ALVO is not None:
        if COLUNA_ALVO not in base.columns:
            raise ValueError(
                f"A coluna-alvo '{COLUNA_ALVO}' não foi encontrada. "
                f"Colunas disponíveis: {list(base.columns)}"
            )
        estratificacao = base[COLUNA_ALVO]
    else:
        estratificacao = None

    # Divisão aleatorizada: 70% treino e 30% teste.
    treino, teste = train_test_split(
        base,
        test_size=PROPORCAO_TESTE,
        random_state=SEMENTE_ALEATORIA,
        shuffle=True,
        stratify=estratificacao
    )

    # Salva as duas bases em arquivos separados.
    arquivo_treino = PASTA_SAIDA / "base_treino.csv"
    arquivo_teste = PASTA_SAIDA / "base_teste.csv"

    treino.to_csv(arquivo_treino, sep=SEPARADOR, index=False, encoding="utf-8-sig")
    teste.to_csv(arquivo_teste, sep=SEPARADOR, index=False, encoding="utf-8-sig")

    # Monta o relatório de execução.
    total_treino = len(treino)
    total_teste = len(teste)
    percentual_treino = total_treino / total_original * 100
    percentual_teste = total_teste / total_original * 100

    log = [
        "=" * 65,
        "LOG DA DIVISÃO DA BASE DE DADOS",
        "=" * 65,
        f"Data/hora da execução: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"Arquivo de entrada: {ARQUIVO_ENTRADA.resolve()}",
        f"Total de registros originais: {total_original:,}",
        f"Quantidade de colunas: {base.shape[1]}",
        "",
        "CONFIGURAÇÕES DA DIVISÃO",
        "-" * 65,
        "Método: divisão aleatorizada (shuffle=True)",
        f"Proporção treino: {1 - PROPORCAO_TESTE:.0%}",
        f"Proporção teste: {PROPORCAO_TESTE:.0%}",
        f"Semente aleatória (random_state): {SEMENTE_ALEATORIA}",
        f"Estratificação: {'Sim' if estratificacao is not None else 'Não'}",
        f"Coluna usada na estratificação: {COLUNA_ALVO if estratificacao is not None else 'Não aplicável'}",
        "",
        "RESULTADO DA DIVISÃO",
        "-" * 65,
        f"Treino: {total_treino:,} registros ({percentual_treino:.2f}%)",
        f"Teste:  {total_teste:,} registros ({percentual_teste:.2f}%)",
        f"Total:  {total_treino + total_teste:,} registros",
        "",
        "ARQUIVOS GERADOS",
        "-" * 65,
        f"Base de treino: {arquivo_treino.resolve()}",
        f"Base de teste:  {arquivo_teste.resolve()}",
    ]

    if estratificacao is not None:
        log.extend([
            "",
            f"DISTRIBUIÇÃO DA VARIÁVEL '{COLUNA_ALVO}' NA BASE ORIGINAL",
            "-" * 65,
            formatar_distribuicao(base[COLUNA_ALVO], total_original),
            "",
            f"DISTRIBUIÇÃO DA VARIÁVEL '{COLUNA_ALVO}' NA BASE DE TREINO",
            "-" * 65,
            formatar_distribuicao(treino[COLUNA_ALVO], total_treino),
            "",
            f"DISTRIBUIÇÃO DA VARIÁVEL '{COLUNA_ALVO}' NA BASE DE TESTE",
            "-" * 65,
            formatar_distribuicao(teste[COLUNA_ALVO], total_teste),
        ])

    log.extend([
        "",
        "VERIFICAÇÕES",
        "-" * 65,
        f"A soma de treino e teste corresponde à base original: "
        f"{'Sim' if total_treino + total_teste == total_original else 'Não'}",
        f"Registros duplicados na base de treino: {treino.duplicated().sum():,}",
        f"Registros duplicados na base de teste: {teste.duplicated().sum():,}",
        "",
        "Observação: a semente aleatória permite reproduzir a mesma divisão,",
        "desde que a base e as configurações permaneçam iguais.",
        "=" * 65,
    ])

    arquivo_log = PASTA_SAIDA / "log_divisao.txt"
    arquivo_log.write_text("\n".join(log), encoding="utf-8")

    print("\n".join(log))
    print(f"\nLog salvo em: {arquivo_log.resolve()}")


if __name__ == "__main__":
    try:
        main()
    except ValueError as erro:
        print(f"\nERRO NA DIVISÃO: {erro}")
        print(
            "\nSe estiver usando estratificação, verifique se todas as classes "
            "da coluna-alvo possuem registros suficientes. Se necessário, "
            "revise os dados ou defina ESTRATIFICAR = False."
        )
        raise
