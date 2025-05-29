import pandas as pd
import xlwings as xw
import os
import sys

csv_path = 'processar.csv' #Caminho + nome do documento com sua extensão
excel_path = 'registros.xlsx' #Caminho + nome do documento com sua extensão

# PROCURA A POSIÇÃO DA PRÓXIMA LINHA VAZIA
def get_next_empty_row(ws):
    last = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
    return last + 1 if ws.range(f"A{last}").value is not None else last

# PROCURA O CSV GERADO PELA KRATOS
try:
    df = pd.read_csv(csv_path)
    print(f"Leitura do CSV '{csv_path}' concluída.")
except Exception as e:
    print(f"Erro ao ler o CSV: {e}")
    sys.exit(1)

app = xw.App(visible=False)
wb = None

try:
    # VERIFICA SE O ARQUIVO EXISTE, CASO CONTRÁRIO, CRIA UM NOVO
    if not os.path.exists(excel_path):
        print(f"Arquivo '{excel_path}' não encontrado. Criando novo arquivo.")
        wb = xw.Book()
        wb.save(excel_path)
    wb = app.books.open(excel_path)
    print(f"Abertura da planilha '{excel_path}' concluída.")

    # VERIFICA DE PLANO REGISTROS EXISTE, CASO CONTRÁRIO, CRIA UM PLANO
    if 'Registros' in [s.name for s in wb.sheets]:
        print("Planilha 'Registros' encontrada.")
        ws = wb.sheets['Registros']
    else:
        print("Planilha 'Registros' não encontrada. Criando nova aba.")
        ws = wb.sheets.add('Registros')

    # ENVIA OS DADOS DA LINHA CSV PARA A PRÓXIMA LINHA
    next_row = get_next_empty_row(ws)
    print(f"Próxima linha vazia: {next_row}")

    # FAZ O TRATAMENTO PARA SEPARAR EM COLUNAS A CADA VIRGULA
    ws.range(f'A{next_row}').value = [df.columns.tolist()] + df.values.tolist()
    print("Dados inseridos com sucesso.")

    # SALVA O ARQUIVO
    wb.save()
    print("Arquivo salvo.")

# 
except Exception as e:
    print(f"Ocorreu um erro durante a manipulação do Excel: {e}")
finally:
    if wb:
        wb.close()
    app.quit()
    print("Excel encerrado.")

print(f"Processo finalizado. Dados do CSV foram inseridos na planilha '{excel_path}'.")