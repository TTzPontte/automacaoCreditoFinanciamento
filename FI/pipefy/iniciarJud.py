# Bibliotecas
import json
import os
from homeJud import executeAll


# Pegando Id do Card
with open(r"..\Scripts\dadosOp-v1.json", encoding='utf-8') as meu_json:
    dados = json.load(meu_json)

idCard = dados['idCliente']

# Verificando de é  HE ou FI
path = os.getcwd()

# Chamando função para baixar os docs de jud
executeAll(idCard,path)

print("BERTANHOLA")