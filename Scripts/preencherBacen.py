from extrairInfosBacen import extrairDividasJson, extrairNomeCliente
import openpyxl as xl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import numpy as np
import os


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
        df = extrairDividasJson(varTesteCliente)

        #Verificar se BACEN tem conteúdo
        if len(df) > 0:
            df = extrairDividasJson(varTesteCliente)
            results = pd.concat([results, df], axis=0).reset_index(drop=True)

    #Tratar nan
    results = results.fillna(0)

    rows = dataframe_to_rows(results, index=False)
    wb = xl.load_workbook(planilhaExcel, keep_vba=True, data_only=False)
    ws = wb['Bacen']

    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 2):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(planilhaExcel)