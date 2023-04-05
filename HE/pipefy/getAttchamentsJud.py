import requests
import json

def getAttachamentsCard(cardID, produto):
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
    if produto == "FI":
        keyValue = {'Bacen': [], 'Precisamos de uma foto do seu RG, CNH ou RNE':[], 'Foto de RG ou CNH - 2o Pagador(a)': [], 'Documentos Pessoais': [],
            'Documentos Imóvel': [], 'Documentos Vendedores':[]
            }
    else:
        keyValue = {'Documentos Pessoais': [], 'SERASAS':[], 'Documentos Imóvel': []
            }
    

    #Criar lista vazia para receber valores com nme dos anexos
    listNames = []

    #Loop para extração da lista de nome dos arquivos e preencher o valor do keyValue
    for item in dictAnexo['data']['card']['attachments']:
        try:
            #print(item['field']['label'])
            kv = keyValue[item['field']['label']]
            if item['field']['label'] == 'Bacen':
                if 'Serasa' in item['url'] or 'serasa' in item['url']:
                    kv.append(item['url'])
            else:
                kv.append(item['url'])
        except:
            pass

        #Criar Lista com nome dos anexos
        newValue = str(item['path'])
        inicio = newValue.find("/", 15) +1
        listNames.append(newValue[inicio:])
    
    return keyValue, listNames 
