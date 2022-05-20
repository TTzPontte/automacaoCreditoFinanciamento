import shutil
from selenium import webdriver
from datetime import date
from selenium.webdriver.chrome.options import Options
from time import sleep
import os
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
from pegarNomeArquivo import nomeDoArquivo
from anexosCliente import listarAnexos
from imageToPdf import transformImageToPDF, mergePDF

#Chromedriver - Google Automatizado
options = webdriver.ChromeOptions() 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#options.add_argument("--headless")
driver = webdriver.Chrome(options=options,executable_path=r'G:\Drives compartilhados\Pontte\Operações\Automações\contaQI\Driver\chromedriver.exe')

#Pegar nome do Usuário
usuario = os.environ.get('USERNAME')

#Caminho Relatório Baixado
pathOrigem = rf"C:\Users\{usuario}\Downloads"

def fazerLogin():
    #Site que irá abrir
    driver.get('https://app.pipefy.com/')
    sleep(10)

    #### Inicio Processo de Login ####
    #Identifica e retorna os elementos
    inserirEmail_element = driver.find_element_by_name('username')
    inseirSenha_element =  driver.find_element_by_name('password')
    entrar_element =  driver.find_element_by_name('submit')

    #InserirValores
    inserirEmail_element.send_keys('matheus.pereira@pontte.com.br')
    inseirSenha_element.send_keys('Pontte2000')
    
    #Clicar
    entrar_element.click()
    sleep(2)

def baixarEMover(pathDestino, dict, nome1, cpf1, nome2, cpf2):
    statusPDF = 0
    try:
        fazerLogin()
    except: 
        print('Login feito')

    sleep(5)
    for nomeChave in dict.keys():
        for nomeValor in dict[nomeChave]:
            driver.get(nomeValor)
            nomeArquivo = nomeDoArquivo(nomeValor)
            origem = pathOrigem + f"\{nomeArquivo}"
            sleep(2)
            if nomeChave == 'Bacen':
                destino = pathDestino + "\Crédito\Bacen" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Certidão de Estado Civil ' or nomeChave == 'Comprovante de Residência' or nomeChave == 'Precisamos de uma foto do seu RG, CNH ou RNE':
                destino = pathDestino + f"\Crédito\Documentos pessoais\{nome1} - {cpf1}" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Foto de RG ou CNH - 2o Pagador(a)':
                destino = pathDestino + f"\Crédito\Documentos pessoais\{nome2} - {cpf2}" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'IRPF':
                destino = pathDestino + "\Crédito\Renda\IR" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Extrato Bancário Últimos 3 meses (Restored)' or nomeChave == 'Extrato bancário últimos 3 meses (2º proponente)' or nomeChave == 'Holerite' or nomeChave == 'Holerite (2º proponente)':
                destino = pathDestino + "\Crédito\Renda" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            #Novos
            elif nomeChave == 'Contrato Social' or nomeChave == 'Extrato Bancário Últimos 6 Meses (PJ) (Restored)':
                destino = pathDestino + "\Crédito\Documentos empresa" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Faturamento 3 anos':
                destino = pathDestino + "\Crédito\Documentos empresa" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Balanço Patrimonial 3 anos':
                destino = pathDestino + "\Crédito\Documentos empresa" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'DRE 3 anos':
                destino = pathDestino + "\Crédito\Documentos empresa" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Matrícula do imóvel':
                destino = pathDestino + "\Imóvel\Documentos do Imóvel" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'CCV':
                destino = pathDestino + "\Imóvel\Documentos do Imóvel" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Capa IPTU':
                destino = pathDestino + "\Imóvel\Documentos do Imóvel" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
            elif nomeChave == 'Fotos do imóvel':
                destino = pathDestino + "\Imóvel\Fotos" + f"\{nomeArquivo}"
                shutil.move(origem,destino)
                if ".png" in destino or ".jpeg" in destino or ".jpg" in destino:
                    transformImageToPDF(destino)
                    statusPDF = 1
    
    destinoImagemImovel = pathDestino + "\Imóvel\Fotos"
    mergePDF(destinoImagemImovel)

                
    


#Final