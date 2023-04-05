import pandas as pd
import json


def extrairNomeCliente(arquivo):
    try:
        with open(arquivo, encoding='utf-8') as file:
            data = json.load(file)
            nome = data['analysis_output']['scr']['name']
        
        return nome
    except:
        with open(arquivo, encoding='utf-8') as file:
            data = json.load(file)
            nome = data['subject_name']
    
        return nome

def extrairDividasJson(arquivo):
    try:
        with open(arquivo, encoding='utf-8') as file:
            data = json.load(file)
            data2 = data['analysis_output']['scr']['scr_data']
            data3 = data['analysis_output']['scr']['signers']
        
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

        nome = data['analysis_output']['scr']['name']
        nomeCliente.append(nome)

        #Extrair e preencher nas listas
        for i in conteudo['operation_items']:
            listaCategoria.append(i['modality_description'])
            listaTipo.append(i['domain_description'])
            listaGrupo.append(i['domain_group'])
            listaDivida.append(i['submodality_description'])
            listaValor.append(i['value'])
        
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
        newDF = df.groupby(['nome','tipoDivida','nomeDivida', 'vencimento'])['valor'].sum().unstack('vencimento').reset_index().fillna(0)
        newDF = pd.DataFrame(newDF)
        #Gerar Excel - Bacen
        createDF(data2, nome)
        return newDF

    except:
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

        nome = data['subject_name']
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
        newDF = df.groupby(['nome','tipoDivida','nomeDivida', 'vencimento'])['valor'].sum().unstack('vencimento').reset_index().fillna(0)
        newDF = pd.DataFrame(newDF)

        #Gerar Excel - Bacen
        createDF(data2, nome)
        return newDF

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

def createDF(scr_data_teste, name):
    import pandas as pd

    # Coloque os dados fornecidos em uma variável chamada scr_data
    scr_data = scr_data_teste

    # Crie uma lista para guardar os DataFrames
    dfs = []

    # Percorra a lista de dados em scr_data
    for data in scr_data:
        reference_date = data['reference_date']
        operation_items = data['operation_items']
        
        # Crie um DataFrame temporário para a iteração atual
        df_temp = pd.DataFrame(columns=['reference_date', 'modality_description', 'submodality_description', 'domain_group', 'domain_description', 'value'])

        # Percorra a lista de itens em operation_items
        for item in operation_items:
            
            try:
                modality_description = item['modality_description']
                submodality_description = item['submodality_description']
                domain_group = item['domain_group']
                domain_description = item['domain_description']
                value = item['value']
            except:
                modality_description = item['category_sub']['category']['category_description']
                submodality_description = item['category_sub']['description']
                domain_group = item['due_type']['due_type_group']
                domain_description = item['due_type']['description']
                value = item['due_value']
            
            # Adicione uma nova linha ao DataFrame temporário com os dados obtidos
            df_temp = df_temp.append({
                'reference_date': reference_date,
                'modality_description': modality_description,
                'submodality_description': submodality_description,
                'domain_group': domain_group,
                'domain_description': domain_description,
                'value': value
            }, ignore_index=True)

        # Adicione o DataFrame temporário à lista de DataFrames
        dfs.append(df_temp)

    # Concatene todos os DataFrames da lista com pd.concat()
    df = pd.concat(dfs)

    df.to_excel(f'..\Crédito\Bacen\Consulta Bacen - {name}.xlsx', sheet_name='Teste Crédito', index=False)