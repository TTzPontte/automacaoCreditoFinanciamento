def nomeDoArquivo(stringAnexo):
    status = 0 
    numeroString = -1
    while status == 0:
        
        if stringAnexo[numeroString:].find("/") == -1:
            numeroString+= -1
        else:
            status = 1

    nomeArquivo = stringAnexo[numeroString+1:]
    return nomeArquivo
