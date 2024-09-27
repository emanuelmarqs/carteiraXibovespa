import pandas as pd
import requests
import openpyxl
# Definindo o token de autenticação
token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxNzI5MDc3NjQ5LCJqdGkiOiJkYmY2ZmQyNzcxMjk0ZDg0YWViMWEzY2RjY2Y0M2IyOCIsInVzZXJfaWQiOjQ1fQ.sYGlN9LERroRQSqrWuKDQMj4WfLOSI7OF2x1LFZbg0o"  
headers = {'Authorization': 'JWT {}'.format(token)}

# Parâmetros da requisição
params = {'data_base': '2023-04-03'}
planilhao = requests.get('https://laboratoriodefinancas.com/api/v1/planilhao', params=params, headers=headers)
planilhao = planilhao.json()['dados']

# Transformando os dados em um DataFrame
df_planilhao = pd.DataFrame.from_dict(planilhao)
colunas_interesse = ['ticker', 'roe']
df_planilhao = df_planilhao[colunas_interesse]
df_planilhao = df_planilhao[df_planilhao['ticker'].isin(['RNEW3', 'UNIP6', 'RNEW4', 'PETR3'])]
df_ticker_ordenado = df_planilhao.sort_values(by='roe', ascending=False)
df_top10 = df_ticker_ordenado.head(10)

if len(df_top10) < 10:
    restantes = df_ticker_ordenado.tail(10 - len(df_top10))
    df_top10 = pd.concat([df_top10, restantes]).reset_index(drop=True)

acoes = {}
resultados = []

# Obtenção dos preços das ações
for ticker in df_top10['ticker']:
    params = {'ticker': ticker, 'data_ini': '2023-04-01', 'data_fim': '2024-04-01'}
    r = requests.get('https://laboratoriodefinancas.com/api/v1/preco-corrigido', params=params, headers=headers)

    if r.status_code == 200:
        preco_corrigido = r.json()['dados']
        df_preco = pd.DataFrame.from_dict(preco_corrigido)
        colunas_desejadas = ['ticker', 'data', 'abertura', 'fechamento']
        df_filtropreco = df_preco[colunas_desejadas]
        acoes[ticker] = df_filtropreco
        
        preco_inicial = df_filtropreco['abertura'].iloc[0]
        preco_final = df_filtropreco['fechamento'].iloc[-1]
        variacao_percentual = (preco_final - preco_inicial) / preco_inicial * 100
        resultados.append({
            'ticker': ticker,
            'preco_inicial': preco_inicial,
            'preco_final': preco_final,
            'variacao_percentual': variacao_percentual
        })

df_resultados = pd.DataFrame(resultados)

# Resultado da carteira de investimentos
carteira = (df_resultados['variacao_percentual'] * 0.10).sum()

# Obtenção do IBOVESPA
params = {'ticker': 'ibov', 'data_ini': '2023-04-03', 'data_fim': '2024-04-01'}
response = requests.get('https://laboratoriodefinancas.com/api/v1/preco-diversos', headers=headers, params=params)

if response.status_code == 200:
    data = response.json()['dados']
    df_ibovespa = pd.DataFrame(data)

    primeira_linha_ibovespa = df_ibovespa.iloc[[0]]
    ultima_linha_ibovespa = df_ibovespa.iloc[[-1]]
    df_ibovespa_filtrado = pd.concat([primeira_linha_ibovespa, ultima_linha_ibovespa])

    valor_fechamento2023_ibovespa = df_ibovespa_filtrado.iloc[0, df_ibovespa_filtrado.columns.get_loc('abertura')]
    valor_fechamento2024_ibovespa = df_ibovespa_filtrado.iloc[1, df_ibovespa_filtrado.columns.get_loc('fechamento')]
    calculo_de_rendimento_ibovespa = (valor_fechamento2024_ibovespa - valor_fechamento2023_ibovespa) / valor_fechamento2023_ibovespa * 100

    # Criando DataFrame para o IBOVESPA
    df_ibovespa_resultado = pd.DataFrame({
        'data': ['2023-04-03', '2024-04-01'],
        'valor': [valor_fechamento2023_ibovespa, valor_fechamento2024_ibovespa],
        'variacao_percentual': [0, calculo_de_rendimento_ibovespa]
    })

    # Comparação entre a carteira e o IBOVESPA
    comparacao = {
        'Descrição': ['Carteira', 'IBOVESPA', 'Diferença'],
        'Valor (%)': [carteira, calculo_de_rendimento_ibovespa, calculo_de_rendimento_ibovespa - carteira]
    }
    resultado_final = pd.DataFrame(comparacao)

    # Salvando em um arquivo Excel
    with pd.ExcelWriter('resultado.xlsx') as writer:
        df_resultados.to_excel(writer, sheet_name='Acoes', index=False)
        df_ibovespa_resultado.to_excel(writer, sheet_name='IBOVESPA', index=False)
        resultado_final.to_excel(writer, sheet_name='Comparacao', index=False)

    print("Resultados salvos em 'resultado.xlsx'.")
else:
    print(f"Erro no IBOVESPA: {response.status_code}")
