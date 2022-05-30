from matplotlib import textpath
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup as bs
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning) 
from time import sleep
import json
import pickle

#Caminho para iniciar ChromeDriver
options = webdriver.ChromeOptions() 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#options.add_argument("--headless")
driver = webdriver.Chrome(options=options,executable_path=r'G:\Drives compartilhados\Pontte\Operações\Automações\contaQI\Driver\chromedriver.exe')

def fazerLogin():
    #Site que irá abrir
    driver.get('https://app.pipefy.com/')

    sleep(5)

    #### Inicio Processo de Login ####

    #Identifica e retorna os elementos
    print('Inicio do processo de login')
    inserirEmail_element = driver.find_element_by_name('username')
    inseirSenha_element =  driver.find_element_by_name('password')
    entrar_element =  driver.find_element_by_name('submit')
    #InserirValoresa
    inserirEmail_element.send_keys('opscontrole@pontte.com.br')
    inseirSenha_element.send_keys('Pontteops22')
    print('login realizado')

    #Clicar
    entrar_element.click()
    sleep(2)
def listarAnexos(linkDoPipefy, pathClient):
    #Fazer Login
    try:
        fazerLogin()
    except:
        pass
    sleep(5)
    #Entrar no Card do Cliente
    driver.get(linkDoPipefy)
    sleep(5)
    #Clicar em Anexos
    driver.find_element_by_id("AnexosTab").click()
    sleep(5)
    #Pegar HTML da página
    soup = bs(driver.page_source, 'html.parser')

    ######################################### Iniciar Scraping de Informações #########################################

    keyValue = {'Anexos do Email': [], 'Bacen':[], 'Certidão de Estado Civil ': [], 
            'Comprovante de Residência': [], 'Extrato Bancário Últimos 3 meses (Restored)': [],
            'Precisamos de uma foto do seu RG, CNH ou RNE':[], 'Foto de RG ou CNH - 2o Pagador(a)':[],
            'IRPF': [], 'Extrato bancário últimos 3 meses (2º proponente)': [], 'Holerite':[],
            'Holerite (2º proponente)':[], 'Contrato Social':[], 'Faturamento 3 anos':[],
            'Balanço Patrimonial 3 anos':[], 'DRE 3 anos':[], 'Matrícula do imóvel':[], 'CCV': [],
            'Capa IPTU':[], 'Fotos do imóvel':[]
            }

    #Gerar Listas de Informações que receberão os Anexos
    try:
        listaAnexoBacen = []
        listaAnexoEstadoCivil = []
        listaAnexoComprovanteRes =[]
        listaAnexoExtrato = []
        listaAnexoDocumento = []
        listaAnexoDocumento2 = []
        listaAnexoIRPF = []
        listaAnexoExtrato2 = []
        listaAnexoHolerite = []
        listaAnexoHolerite2 = []
        #Lista PJ
        listaAnexoContratoSocial = []
        listaAnexoFaturamento3Anos = []
        listaAnexoBalanco = []
        listaAnexoDRE = []
        listaAnexoMatricula = []
        listaAnexoCCV = []
        listaAnexoCapaIPTU = []
        listaAnexoFotoImovel = []
    except:
        print('Erro ao gerar as listas')

    #Scraping dos Anexos
    html1 = soup.find('div', {'class':'pp-box-content'})
    for divPrincipal in html1:
        titulo = divPrincipal.find('span')
        if titulo is not None:
            if titulo.text == 'Bacen' or titulo.text == 'BACEN':
                conteudo2 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoBacen in conteudo2:
                    listaAnexoBacen.append(anexoBacen['href'])
            if titulo.text == 'Certidão de Estado Civil ':
                conteudo3 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoEstado in conteudo3:
                    listaAnexoEstadoCivil.append(anexoEstado['href'])
            if titulo.text == 'Comprovante de Residência':
                conteudo4 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoComprovante in conteudo4:
                    listaAnexoComprovanteRes.append(anexoComprovante['href'])
            if titulo.text == 'Extrato Bancário Últimos 3 meses (Restored)':
                conteudo5 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoExtrato in conteudo5:
                    listaAnexoExtrato.append(anexoExtrato['href'])
            if titulo.text == 'Precisamos de uma foto do seu RG, CNH ou RNE':
                conteudo6 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoDocumento in conteudo6:
                    listaAnexoDocumento.append(anexoDocumento['href'])
            if titulo.text == 'Foto de RG ou CNH - 2o Pagador(a)':
                conteudo7 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoDocumento2 in conteudo7:
                    listaAnexoDocumento2.append(anexoDocumento2['href'])
            if titulo.text == 'IRPF':
                conteudo8 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoIRPF in conteudo8:
                    listaAnexoIRPF.append(anexoIRPF['href'])
            if titulo.text == 'Extrato bancário últimos 3 meses (2º proponente)':
                conteudo9 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoExtrato2 in conteudo9:
                    listaAnexoExtrato2.append(anexoExtrato2['href'])
            if titulo.text == 'Holerite':
                conteudo10 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoHolerite in conteudo10:
                    listaAnexoHolerite.append(anexoHolerite['href'])
            if titulo.text == 'Holerite (2º proponente)':
                conteudo11 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoHolerite2 in conteudo11:
                    listaAnexoHolerite2.append(anexoHolerite2['href'])
            if titulo.text == 'Contrato Social':
                conteudo12 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoContratoSocial in conteudo12:
                    listaAnexoContratoSocial.append(anexoContratoSocial['href'])
            if titulo.text == 'Faturamento 3 anos':
                conteudo13 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoFaturamento3Anos in conteudo13:
                    listaAnexoFaturamento3Anos.append(anexoFaturamento3Anos['href'])
            if titulo.text == 'Balanço Patrimonial 3 anos':
                conteudo14 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoBalanco in conteudo14:
                    listaAnexoBalanco.append(anexoBalanco['href'])
            if titulo.text == 'DRE 3 anos':
                conteudo15 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoDRE in conteudo15:
                    listaAnexoDRE.append(anexoDRE['href'])
            if titulo.text == 'Matrícula do imóvel':
                conteudo16 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoMatricula in conteudo16:
                    listaAnexoMatricula.append(anexoMatricula['href'])
            if titulo.text == 'CCV':
                conteudo17 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoCCV in conteudo17:
                    listaAnexoCCV.append(anexoCCV['href'])
            if titulo.text == 'Capa IPTU':
                conteudo18 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoCapaIPTU in conteudo18:
                    listaAnexoCapaIPTU.append(anexoCapaIPTU['href'])
            if titulo.text == 'Fotos do imóvel':
                conteudo19 = divPrincipal.find_all('a', {'class' : 'StyledAnchor-pstyle__sc-112a9fu-0 ccMLjg pp-margin-right-16 pp-no-margin pp-anchor'})
                for anexoFotos in conteudo19:
                    listaAnexoFotoImovel.append(anexoFotos['href'])
                    
    #Popular Keyvalue
    keyValue['Bacen'] = listaAnexoBacen    
    keyValue['Certidão de Estado Civil '] = listaAnexoEstadoCivil    
    keyValue['Comprovante de Residência'] = listaAnexoComprovanteRes    
    keyValue['Extrato Bancário Últimos 3 meses (Restored)'] = listaAnexoExtrato    
    keyValue['Precisamos de uma foto do seu RG, CNH ou RNE'] = listaAnexoDocumento    
    keyValue['Foto de RG ou CNH - 2o Pagador(a)'] = listaAnexoDocumento2    
    keyValue['IRPF'] = listaAnexoIRPF
    keyValue['Extrato bancário últimos 3 meses (2º proponente)'] = listaAnexoExtrato2
    keyValue['Holerite'] = listaAnexoHolerite
    keyValue['Holerite (2º proponente)'] = listaAnexoHolerite2
    keyValue['Contrato Social'] = listaAnexoContratoSocial
    keyValue['Faturamento 3 anos'] = listaAnexoFaturamento3Anos
    keyValue['Balanço Patrimonial 3 anos'] = listaAnexoBalanco
    keyValue['DRE 3 anos'] = listaAnexoDRE
    keyValue['Matrícula do imóvel'] = listaAnexoMatricula
    keyValue['CCV'] = listaAnexoCCV
    keyValue['Capa IPTU'] = listaAnexoCapaIPTU
    keyValue['Fotos do imóvel'] = listaAnexoFotoImovel
        
    #Retornar dicionáiro com os anexos
    return keyValue
