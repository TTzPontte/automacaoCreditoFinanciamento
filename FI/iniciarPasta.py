# Bibliotecas
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
import os
import glob
import json
from datetime import datetime
import re
import xlwings as xw
from preencherBacen import gravarBacenSimulador
from pipefy.home import executeAll
from colorama import Fore, Back, Style, init

init()

### Validar qual função o analista quer que rode ###
valid = False
while valid == False:
    rodar = input(Fore.YELLOW+"""Escolha a opção que deseja executar: 
    - Tudo (1)
    - Download Docs (2)
    - Preencher Bacen (3)
    Digitar apenas o número da opção desejada --> """+Style.RESET_ALL)
    if rodar == "1" or rodar == "2" or rodar == "3":
       rodar = int(rodar)
       valid = True
    else:
        print(Fore.RED+"Escolha uma opção valida!"+Style.RESET_ALL)

### Funçoes ###

def deParaimovel(varImovel):
    if varImovel == "Casa Padrão" or varImovel == "Casa de Condomínio":
        varImovel = "Casa"
    return varImovel
def deParaOps(varOps):
    if varOps == "Pessoa Física":
        varOps = "PF"
    elif varOps == "Pessoa Jurídica":
        varOps = "PJ"
    return varOps
def converterDatas(date):
    try:
        date =  date[:10]
        date_obj = datetime.strptime(str(date), '%Y-%m-%d')
        date = date_obj.strftime("%d/%m/%Y")
    except:
        pass
    return date


with open("dadosOp-v1.json", encoding='utf-8') as meu_json:
    dados = json.load(meu_json)
try:
    pasta_cliente = rf"G:\Drives compartilhados\Pontte\Operações\Projetos\Automação de Crédito\Simulação FI - V2\{dados['nomeCompleto']}- ID {dados['idCliente']}"
    pasta_bacen = rf"G:\Drives compartilhados\Pontte\Operações\Projetos\Automação de Crédito\Simulação FI - V2\{dados['nomeCompleto']}- ID {dados['idCliente']}\Crédito\Bacen"
except:
    print(Fore.RED+"!! Erro no Diretório ou Nome da Pasta !!"+Style.RESET_ALL)

id = dados['idCliente']
linkPipefy = "https://app.pipefy.com/open-cards/" + str(id)
nomeCompleto = dados['nomeCompleto']
tipoOps = dados['tipoOperação']
tipoOps = deParaOps(tipoOps)
cpf = dados['cpf']
dataNascimento = dados['dataNascimento']
dataNascimento = converterDatas(dataNascimento)
comporRenda = dados['comporRenda']
nomeComporRenda = dados['nomeCompleto2']
cpfComporRenda = dados['cpf2']
dataNascimentoComporRenda = dados['dataNascimento2']
dataNascimentoComporRenda = converterDatas(dataNascimentoComporRenda)
valorImovel = dados['valorImovel']
valorLiquido = dados['valorImovelFinanciar']
qtdParcelas = dados['quantidadeParcelas']
qtdParcelas = re.sub('[^0-9]', '', qtdParcelas)
carencia = dados['carencia']
carencia = re.sub('[^0-9]', '', carencia)
#-------- new ----------
tipoImovel = dados['tipoImovel']
tipoImovel = deParaimovel(tipoImovel)
Socio = dados['socio']
razaoSocial = dados['razaoSocial']
cnpjEmpresa = dados['cnpj']
areaImovel = dados['areaImovel']
enderecoImovel = dados['enderecoImovel']
enderecoImovel =  dados['enderecoImovel']
itbi = dados['incluirItbiRegistro']
rendaMedia = dados['rendaMedia']
rendaMedia2 = dados['rendaMedia2']
parceiro = dados['empresaParceira']


# Pegar o Simulador
pathSpreadSheet = glob.glob(pasta_cliente+"\*xlsm") #Pegar Caminho do Simulador dentro da pasta do cliente
pathSpreadSheet = pathSpreadSheet[0] #Transformar de Lista ---> STR
#print(pathSpreadSheet)

# Baixando Documentos
if rodar == 1 or rodar == 2:
    try:
        executeAll(dados['idCliente'],pasta_cliente )
        print(Fore.GREEN +'Download de Documento Finalizado!'+Style.RESET_ALL)
    except:
        print(Fore.RED+"!! Erro no Download dos documento anexados no card !!"+Style.RESET_ALL)




# Preencher Simulador com as Informações do Card
if rodar == 1:
    # Manipulação do simulador do Excel
    try:
        wb = xw.Book(pathSpreadSheet, read_only=False) #Carregar a planilha
        # Abrir Abas
        ws = wb.sheets['Inputs']


        # Escrever nas planilhas
        ws["F11"].value = int(id)
        ws["D12"].value = int(qtdParcelas)
        if carencia == "Não informado" or carencia == 'Não tenho interesse' or carencia == '':
            ws["D16"].value = 0
        else:    
            ws["D16"].value = float(carencia)
        if comporRenda == "Sim":
            somaRenda = float(rendaMedia.replace(",","")) + float(rendaMedia2.replace(",",""))
        else:
            somaRenda = float(rendaMedia.replace(",",""))


        ws["D18"].value = somaRenda
        ws["F13"].value = parceiro
        ws["D11"].value = valorLiquido
        ws["D45"].value = valorImovel
        # ws["C13"].value = valorImovel
        ws["F17"].value = itbi
        ws["D14"].value = tipoOps
        # ws["D15"].value =     ## Colocar se é residencial ou comercial   
        # Nomes
        ws["C23"].value = nomeCompleto
        ws["C24"].value = nomeComporRenda
        ws["C25"].value = razaoSocial
        # DataNascimento
        ws["F23"].value = dataNascimento
        ws["F24"].value = dataNascimentoComporRenda
        # Link
        ws["E8"].value = linkPipefy
        # #CPF e CNPJ
        ws["E23"].value = cpf
        ws["E24"].value = cpfComporRenda
        ws["E25"].value = cnpjEmpresa

        # #Avaliacao
        # ws2["F9"].value = areaImovel

        #Salvar e Fechar
        wb.save(pathSpreadSheet)
        wb.close()
        print(Fore.GREEN +'Preenchimento do Simulador Finalizado!'+Style.RESET_ALL)
    except:
        print(Fore.RED + '!! Erro ao preencher Excel! !!'+Style.RESET_ALL)

if rodar == 1 or rodar == 3:        
    try:
        gravarBacenSimulador(pasta_bacen, pathSpreadSheet)
        print(Fore.GREEN +'Preenchimeto do Bacen Finalizado! '+Style.RESET_ALL)
    except:
        print(Fore.RED + '!! Erro ao gravar BACEN !!'+Style.RESET_ALL)

