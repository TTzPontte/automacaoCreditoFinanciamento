import requests

# Sua chave de API do Pipefy
api_key = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjozMDEyNjA0NzYsImVtYWlsIjoiZGV2QHBvbnR0ZS5jb20uYnIiLCJhcHBsaWNhdGlvbiI6MzAwMjI0MjU1fX0.mETDV7VXfKgr7ubBcEqtf1IyJ2OHbOjgUFKF3Bk7J2We_UUXNh0oq0N6ZEmVsLYaqPyQR2qx7yn7KfpztPoqcg"

# ID do relatório que você deseja baixar
report_id = "REPORT_ID"

# Cabeçalhos da requisição
headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json"
}

# Fazendo a requisição
response = requests.get(f"https://api.pipefy.com/reports/{report_id}/download", headers=headers)

# Verificando se a requisição foi bem sucedida
if response.status_code == 200:
    # Salvando o arquivo
    open("report.xlsx", "wb").write(response.content)
    print("Arquivo salvo com sucesso!")
else:
    # Exibindo o erro
    print(response.json())
