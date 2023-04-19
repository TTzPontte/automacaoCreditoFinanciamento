from extrairInfosBacen import extrairDividasJson, extrairNomeCliente
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import numpy as np
import os
import xlwings


def gravarBacenSimulador(pathBacens, pathSimulador):
    #Variáveis de entrada
    files_dir = pathBacens
    planilhaExcel = pathSimulador

    #Criar Lista de Arquivos BACEN
    bacen_files = [f for f in os.listdir(files_dir) if f.endswith(".json")]

    #Variável Auxiliar Vazia
    results = pd.DataFrame()

    #Ler Todos os BACEN da pasta e transformar em DF
    for file in bacen_files:
        
        varTesteCliente = os.path.join(files_dir, file)
        df1 = extrairDividasJson(varTesteCliente)

        #Verificar se BACEN tem conteúdo
        if len(df1) > 0:
            df = extrairDividasJson(varTesteCliente)
            results = pd.concat([results, df], axis=0).reset_index(drop=True)


    #Tratar nan
    results = results.fillna(0)
    colunas = list(df.columns)

    if not 'A vencer, próx 30 dias.'in colunas:
        results['A vencer, próx 30 dias.'] = 0
    if not 'A vencer: 31 a 360 dias.'in colunas:
        results['A vencer: 31 a 360 dias.'] = 0
    if not 'A vencer: acima de 361 dias.'in colunas:
        results['A vencer: acima de 361 dias.'] = 0

    results = results[['nome', 'tipoDivida', 'nomeDivida', 'A vencer, próx 30 dias.', 'A vencer: 31 a 360 dias.', 'A vencer: acima de 361 dias.']]


    wb = xlwings.Book(planilhaExcel)
    ws = wb.sheets['Bacen']

    ws["A1"].value= results

    ### Código antigo 
    # for r_idx, row in enumerate(rows, 1):
    #     for c_idx, value in enumerate(row, 2):
    #         ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(planilhaExcel)
    wb.close()