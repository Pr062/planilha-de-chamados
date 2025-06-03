#chamados
import pandas as pd
from datetime import datetime
import os
import argparse

def main():
    parser = argparse.ArgumentParser(description="Análise combinada de chamados Excel.")
    parser.add_argument('--pasta', type=str, default=os.path.join(os.path.expanduser("~"), "Downloads"),
                        help='Pasta onde está o arquivo Excel (default: Downloads)')
    parser.add_argument('--arquivo', type=str, default="chamado_de_1º_nivel3.xlsx",
                        help='Nome do arquivo Excel (default: chamado_de_1º_nivel3.xlsx)')
    parser.add_argument('--inicio', type=str, default="2025-05-12",
                        help='Data de início para filtro (YYYY-MM-DD), usado no 3º código')
    parser.add_argument('--fim', type=str, default="2025-05-16",
                        help='Data de fim para filtro (YYYY-MM-DD), usado no 3º código')
    args = parser.parse_args()

    caminho_arquivo = os.path.join(args.pasta, args.arquivo)

    if not os.path.exists(caminho_arquivo):
        print(f"Arquivo não encontrado: {caminho_arquivo}")
        return

    # Lê a planilha Excel e limpa colunas
    df = pd.read_excel(caminho_arquivo)
    df.columns = df.columns.str.strip()

    print("\n====== PRIMEIRO CÓDIGO ======")
    try:
        coluna_i = df.iloc[:, 8]  # Coluna 9 (índice 8)
        quantidade_verdadeiros = coluna_i.sum()
        print(f"Quantidade de chamados abertos desde janeiro (coluna 9): {quantidade_verdadeiros}\n")
    except Exception as e:
        print(f"Erro no primeiro código (coluna 9): {e}")

    print("\n====== SEGUNDO CÓDIGO ======")
    col_data = 'Data de resposta'
    col_resolvido = 'Resolvido(a)s'
    if col_data in df.columns and col_resolvido in df.columns:
        try:
            df[col_data] = pd.to_datetime(df[col_data], errors='coerce')
            inicio_maio = datetime(2025, 5, 1)
            fim_maio = datetime(2025, 5, 31, 23, 59, 59)
            df_maio = df[(df[col_data] >= inicio_maio) & (df[col_data] <= fim_maio)]

            resolvidos = df_maio[df_maio[col_resolvido] == True]
            nao_resolvidos = df_maio[df_maio[col_resolvido] != True]

            print(f"MÊS DE MAIO/2025:")
            print(f"Total de chamados em maio: {len(df_maio)}")
            print(f"Resolvidos (verdadeiro): {len(resolvidos)}")
            print(f"Não resolvidos (Falso ou vazio): {len(nao_resolvidos)}\n")
        except Exception as e:
            print(f"Erro no segundo código: {e}")
    else:
        print(f"Colunas '{col_data}' ou '{col_resolvido}' não encontradas para análise do segundo código.\n")

    print("\n====== TERCEIRO CÓDIGO ======")
    try:
        col_data_abertura = df.columns[17]  # coluna 18
        col_resolvido_status = df.columns[8]  # coluna 9
    except IndexError:
        print("Erro: A planilha não possui colunas suficientes para acessar a coluna 18 ou 9.")
        return

    print(f"Usando coluna 18 (abertura): '{col_data_abertura}'")
    print(f"Usando coluna 9 (resolvido): '{col_resolvido_status}'")

    try:
        df[col_data_abertura] = pd.to_datetime(df[col_data_abertura], errors='coerce')
        df[col_resolvido_status] = df[col_resolvido_status].apply(
            lambda x: str(x).strip().lower() in ['true', '1', 'sim', 'yes']
        )
    except Exception as e:
        print(f"Erro ao converter colunas no terceiro código: {e}")
        return

    try:
        inicio = datetime.strptime(args.inicio, "%Y-%m-%d")
        fim = datetime.strptime(args.fim, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    except ValueError:
        print("Erro no formato da data. Use o formato YYYY-MM-DD.")
        return

    chamados_abertos = df[(df[col_data_abertura] >= inicio) & (df[col_data_abertura] <= fim)]
    chamados_resolvidos = chamados_abertos[df[col_resolvido_status] == True]
    chamados_nao_resolvidos = chamados_abertos[df[col_resolvido_status] == False]

    print(f"Período analisado: {inicio.date()} a {fim.date()}")
    print(f"Total de chamados abertos nesse período: {len(chamados_abertos)}")
    print(f"Resolvidos (coluna 9 = True): {len(chamados_resolvidos)}")
    print(f"Não resolvidos (coluna 9 = False): {len(chamados_nao_resolvidos)}")

if __name__ == "__main__":
    main()


