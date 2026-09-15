import pandas as pd

def explorar_csv():
    df = pd.read_csv("dados_csv/IEXOBR26.csv", sep=";", encoding="latin-1", low_memory=False)
    print(df.columns)

if __name__ == "__main__":

    explorar_csv()