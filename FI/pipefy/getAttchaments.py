import requests
import json

def getAttachamentsCard(cardID):
    cardID = str(cardID)
    
    #Preparar variaveis para requisição da API
    url = "https://api.pipefy.com/graphql"
    payload = {"query": "{ card (id: " + cardID + ") {attachments{field{label} path url} }}"}
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjozMDEyNjA0NzYsImVtYWlsIjoiZGV2QHBvbnR0ZS5jb20uYnIiLCJhcHBsaWNhdGlvbiI6MzAwMjI0MjU1fX0.mETDV7VXfKgr7ubBcEqtf1IyJ2OHbOjgUFKF3Bk7J2We_UUXNh0oq0N6ZEmVsLYaqPyQR2qx7yn7KfpztPoqcg'
    }

    #Requisição da API
    response = requests.post(url, json=payload, headers=headers)

    #Guardando retorno da API em uma variável
    textoAnexo = response.text

    #Criar chave valor com a saída da API
    dictAnexo = json.loads(textoAnexo)

    #KeyValue FI
    keyValue = {'Anexos do Email': [], 'Bacen':[], 'Certidão de Estado Civil ': [], 
        'Comprovante de Residência': [], 'Extrato Bancário Últimos 3 meses (Restored)': [],
        'Precisamos de uma foto do seu RG, CNH ou RNE':[], 'Foto de RG ou CNH - 2o Pagador(a)':[],
        'IRPF': [], 'Extrato bancário últimos 3 meses (2º proponente)': [], 'Holerite':[],
        'Holerite (2º proponente)':[], 'Contrato Social':[], 'Faturamento 3 anos':[],
        'Balanço Patrimonial 3 anos':[], 'DRE 3 anos':[], 'Matrícula do imóvel':[], 'CCV': [],
        'Capa IPTU':[], 'Fotos do imóvel':[],'Extrato Bancário Últimos 6 meses (PJ) (Restored)':[], 'Documentos Financeiros':[],
        'Documentos Pessoais': [], 'Documentos Imóvel':[],'Documentos Vendedores':[],'SERASAS':[]
        }

    #Criar lista vazia para receber valores com nme dos anexos
    listNames = []

    #Loop para extração da lista de nome dos arquivos e preencher o valor do keyValue
    for item in dictAnexo['data']['card']['attachments']:
        try:
            kv = keyValue[item['field']['label']]
            kv.append(item['url'])
        except:
            pass

        #Criar Lista com nome dos anexos
        newValue = str(item['path'])
        inicio = newValue.find("/", 15) +1
        listNames.append(newValue[inicio:])
    
    return keyValue, listNames, 