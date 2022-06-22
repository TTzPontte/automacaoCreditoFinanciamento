import pandas as pd
import json

#arquivoJson = r"G:\Drives compartilhados\Pontte\Operações\Projetos\Automação de Crédito\Json\FERNANDOMAGNODEOLIVEIRACASTRO-28142050803-2022-05-27.json"
#arquivoJson = r"G:\Drives compartilhados\Pontte\Operações\Projetos\Automação de Crédito\Json\AdrianaChagasAlao-93165790663-2022-03-29.json"

def extrairNomeCliente(arquivo):
    with open(arquivo, encoding='utf-8') as file:
        data = json.load(file)
        data3 = data['signers']
    
    for x in data3:
        nome = x['name']
    
    return nome

def extrairDividasJson(arquivo):
    with open(arquivo, encoding='utf-8') as file:
        data = json.load(file)
        data2 = data['scr_data']
        data3 = data['signers']
    
    #Listar todas as datas mostradas no Json
    listaData = []
    for line in data2:
        if 'reference_date' in line:
            listaData.append(line['reference_date'])
    
    #Selecionar a última data observada no json
    tamanhoLista=  len(listaData)
    extList = listaData[-2]
    
    #Extrair Bloco que contém a última data observada no json
    teste = ''
    for line in data2:
        if 'reference_date' in line:
            if line['reference_date'] == extList:
                teste = line['reference_date']
            elif teste == extList:
                teste = 1
            elif teste == 1:
                teste = 0
        if teste == 1:
            conteudo = line
    
    #Criar Listas Vazias para receber valores do Json
    nomeCliente = []
    listaCategoria = []
    listaTipo = []
    listaGrupo = []
    listaValor = []
    listaDivida = []

    #Extrair nome do cliente
    for x in data3:
        nome = x['name']
        nomeCliente.append(nome)

    #Extrair e preencher nas listas
    for i in conteudo['operation_items']:
        listaCategoria.append(i['category_sub']['category']['category_description'])
        listaTipo.append(i['due_type']['description'])
        listaGrupo.append(i['due_type']['due_type_group'])
        listaDivida.append(i['category_sub']['description'])
        listaValor.append(i['due_value'])
    
    #Criar listas aplicando os DE-PARA's
    vencimentoDivida = []
    nomeDividas = []
    for itemList in listaTipo:
        vencimentoDivida.append(tempoDivida(itemList))
    for nomeDivida in listaDivida:
        nomeDividas.append(deParaDividas(nomeDivida))

    #Criar uma lista única (Pré Criação do DF)
    listaUnica = list(zip(listaCategoria, vencimentoDivida, listaGrupo, nomeDividas, listaValor))

    #Criar Data Frame
    df = pd.DataFrame(listaUnica, columns=[['Categoria', 'Tipo', 'Grupo', 'Nome da Dívida', 'Valor']], index=None)
    df.insert(0,"Nome do Cliente", nomeCliente[0])
    df.columns = ["nome", 'categoria', 'vencimento', 'tipoDivida', 'nomeDivida', 'valor']

    #Criar Novo DF
    newDF = df.groupby(['tipoDivida','nomeDivida'])['valor'].sum().unstack('tipoDivida').reset_index().fillna(0)
    newDF = pd.DataFrame(newDF)

    #Contruir DataFrame de Saída
    newDF2 = pd.DataFrame(newDF['nomeDivida'])
    try:
        newDF2['Vencido'] = newDF['Vencido']
    except:
        newDF2['Vencido'] = 0
    try:
        newDF2['Prejuizo'] = newDF['Prejuizo']
    except:
        newDF2['Prejuizo'] = 0
    try:
        newDF2['Em dia'] = newDF['A vencer']
    except:
        newDF2['Em dia'] = 0
    try:
        newDF2['Limite'] = newDF['Limite de Credito']
    except:
        newDF2['Limite'] = 0
    
    return newDF2

#Funções DE-PARA

def deParaDividas(x):
    if 'Cartão' in x or 'cartao' in x or 'cartão' in x or 'Cartao' in x:
        x = 'Cartão de Crédito'
    elif 'capital' in x or 'Capital' in x:
        x = 'Capital de Giro'
    elif 'cheque' in x or 'Cheque' in x:
        x = 'Cheque Especial'
    elif 'Conta garantida' in x or 'conta garantida' in x:
        x = 'Cheque Especial'
    elif 'sem consignação' in x:
        x = 'Crédito Pessoal s/ Consignação'
    elif 'Crédito pessoal' in x or 'crédito pessoal' in x:
        x = 'Crédito Pessoal'
    elif 'adiantamentos a depositantes' in x or 'desconto de duplicatas' in x:
        x = 'Desconto de duplicatas'
    elif 'aquisição de bens – outros bens' in x:
        x = 'Financiamento Outros Bens'
    elif 'aquisição de bens – veículos automotores' in x:
        x = 'Financiamento de Automóvel'
    elif 'financiamento habitacional – SFH' in x:
        x = 'Financiamento de Imóvel - SFH'
    elif 'financiamento habitacional – exceto SFH' in x:
        x = 'Financiamento de Imóvel'
    elif 'contratado e não utilizado' in x or 'custeio e pré-custeio' in x or 'microcrédito' in x or 'outros empréstimos' in x or 'recebíveis adquiridos' in x:
        x = 'Outros empréstimos'
    elif 'financiamento de projeto' in x or 'outros financiamentos' in x or 'Outros financiamentos' in x:
        x = 'Outros financiamentos'
    elif 'Descontos' in x or 'Outros Limites' in x or 'avais e fianças honrados' in x or 'garantias prestadas' in x or 'desconto de cheques' in x:
        x = 'Outros Recebíveis Adquiridos'
    
    #Retornar valor de X
    return x

def tempoDivida(y):
    if 'a vencer até 30 dias' in y:
        y = 'A vencer, próx 30 dias.'
    elif 'a vencer de 31 a 60 dias' in y or 'a vencer de 61 a 90 dias' in y or 'a vencer de 91 a 180 dias' in y or 'Créditos a vencer de 181 a 360 dias' in y:
        y = 'A vencer: 31 a 360 dias.'
    else:
        y = 'A vencer: acima de 361 dias.'    
    #Retornar Valor
    return y
