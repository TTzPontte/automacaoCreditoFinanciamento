from pipefy.downloadAttachaments import makeDownload #downloadAttachaments --> Dont have return
from pipefy.getAttchaments import getAttachamentsCard #getAttachemnts --> Return: Key Value and ListNames
from time import sleep
from convertToPdf import checkFolder
import os

def executeAll(cardNumber, pathdirName):
    chaveValor = {}
    listaNomes = []

    chaveValor, listaNomes = getAttachamentsCard(cardNumber)

    #Pegar nome do Usuário
    usuario = os.environ.get('USERNAME')

    #Aguardar 2 segundos
    sleep(2)

    for item in chaveValor:
        nomeChave = item
        print(item)
        if nomeChave == 'BACEN' or nomeChave == 'BACENs':
            dirName = pathdirName + "\Crédito\Bacen"

        elif nomeChave == 'Certidão de Estado Civil ' or nomeChave == 'Comprovante de Residência' or nomeChave == 'Precisamos de uma foto do seu RG, CNH ou RNE' or nomeChave == 'Documentos Pessoais':
            dirName = pathdirName + f"\Crédito\Documentos pessoais"

        elif nomeChave == 'Foto de RG ou CNH - 2o Pagador(a)':
            dirName = pathdirName + f"\Crédito\Documentos pessoais"

        elif nomeChave == 'IRPF':
            dirName = pathdirName + "\Crédito\Renda"

        elif nomeChave == 'Extrato Bancário Últimos 3 meses (Restored)' or nomeChave == 'Extrato bancário últimos 3 meses (2º proponente)' or nomeChave == 'Holerite' or nomeChave == 'Holerite (2º proponente)' or nomeChave =='Documentos Financeiros':
            dirName = pathdirName + "\Crédito\Renda"

        elif nomeChave == 'Contrato Social' or nomeChave == 'Extrato Bancário Últimos 6 Meses (PJ) (Restored)':
            dirName = pathdirName + "\Crédito\Renda"
        
        elif nomeChave == 'SERASAS':
            dirName = pathdirName + "\Crédito\Serasa"

        elif nomeChave == 'Faturamento 3 anos':
            dirName = pathdirName + "\Crédito\Renda"

        elif nomeChave == 'Balanço Patrimonial 3 anos':
            dirName = pathdirName + "\Crédito\Renda"
            
        elif nomeChave == 'DRE 3 anos':
            dirName = pathdirName + "\Crédito\Renda"
            
        elif nomeChave == 'Matrícula do imóvel' or nomeChave == 'Documentos Imóvel':
            dirName = pathdirName + "\Imóvel"
            
        elif nomeChave == 'CCV':
            dirName = pathdirName + "\Imóvel"
            
        elif nomeChave == 'Capa IPTU':
            dirName = pathdirName + "\Imóvel"

        elif nomeChave == 'Fotos Do Imóvel':
            dirName = pathdirName + "\Imóvel"

        elif nomeChave == 'Documentos Vendedores':
          dirName = pathdirName + "\DD\\1. Vendedores"
        
        else:
            dirName = rf"C:\Users\{usuario}\Downloads"
            
        urlDownload = chaveValor[item]
        if urlDownload != []:
            for link in urlDownload:
                for value in listaNomes:
                    if value in link:
                        path = os.path.join(dirName, value)
                        makeDownload(link, path)
        # Passar de Imagem para PDF
        checkFolder(dirName)

#Teste
# numeroCard = 627681545
# pathA = r'C:\Users\MatheusPereira\Documents\automacaoCreditoFinanciamento\pastaTeste\Ghislain Assis De Souza- ID 630009805'
# executeAll(numeroCard, pathA)


