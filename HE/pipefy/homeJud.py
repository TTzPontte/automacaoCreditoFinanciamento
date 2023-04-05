from downloadAttachaments import makeDownload #downloadAttachaments --> Dont have return
from getAttchamentsJud import getAttachamentsCard #getAttachemnts --> Return: Key Value and ListNames
from convertToPdf import checkFolder 
from time import sleep
import os

def executeAll(cardNumber, pathdirName):
    chaveValor = {}
    listaNomes = []

    #Definir Produto
    if 'Simulação FI' in pathdirName or 'FINANCIAMENTO' in pathdirName:
        produto = "FI"
    elif 'Simulação HE' in pathdirName or 'HOME EQUITY' in pathdirName:
        produto = "HE"
    
    chaveValor, listaNomes = getAttachamentsCard(cardNumber, produto)
    
    #Pegar nome do Usuário
    usuario = os.environ.get('USERNAME')

    #Aguardar 2 segundos
    sleep(2)

    for item in chaveValor:
        dirName = pathdirName + "\DD"
            
        urlDownload = chaveValor[item]
        if urlDownload != []:
            for link in urlDownload:
                for value in listaNomes:
                    if value in link:
                        path = os.path.join(dirName, value)
                        makeDownload(link, path)
    
    #Verificar se existe arquivo em imagem
    checkFolder(dirName)


