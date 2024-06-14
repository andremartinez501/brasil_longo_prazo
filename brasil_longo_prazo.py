from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import pandas as pd
from sklearn.linear_model import LinearRegression
import numpy as np


#Caminho para o diretório de download 
download_dir = r'C:\Users\andre\Downloads\downloads_automatizados'

options = webdriver.ChromeOptions()
prefs = {"download.default_directory": download_dir}
options.add_experimental_option("prefs", prefs)

#Inicializar o ChromeDriver usando o WebDriver Manager
driver = webdriver.Chrome(service=webdriver.chrome.service.Service(ChromeDriverManager().install()), options=options)

try:
    #Abrir o site do Itaú
    driver.get('https://www.itau.com.br/itaubba-pt/analises-economicas/projecoes')

    time.sleep(5)

    time.sleep(2)
    
    download_link = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="csce364dd9e2f707f8"]/div[3]/ul/li[3]/div/div/div[2]/a'))
)
   

    #Rolar até o link de download
    driver.execute_script("arguments[0].scrollIntoView();", download_link)

    #Clicar no link de download
    driver.execute_script("arguments[0].click();", download_link)

    time.sleep(10)

finally:
    driver.quit()

#Buscar o arquivo baixado mais recente no diretório de download
def ultimo_arquivo(diretorio):
    files = [os.path.join(diretorio, f) for f in os.listdir(diretorio)]
    return max(files, key=os.path.getctime)

file_path = ultimo_arquivo(download_dir)

#Carregar o arquivo Excel em um DataFrame do pandas tendo a segunda linha como cabeçalho
df = pd.read_excel(file_path, header=1)

#Remover as colunas A e B (as duas primeiras colunas)
df = df.iloc[:, 2:]

#Remover a linha 3 (a linha com índice 2)
df = df.drop(index=2)

#Identificar a linha onde está "Taxa de Juros"
taxa_juros_idx = df[df.iloc[:, 0].str.contains('Taxa de Juros', na=False, case=False)].index[0]

#Identificar a última linha onde está "Inflação" antes de "Taxa de Juros"
inflacao_mask = df.iloc[:taxa_juros_idx, 0].str.contains('Inflação', na=False, case=False)
inflacao_idx = inflacao_mask[inflacao_mask].index[-1]

#Identificar a linha onde começa "Finanças Públicas"
financas_publicas_idx = df[df.iloc[:, 0].str.contains('Finanças Públicas', na=False, case=False)].index[0]

#Separar as seções "Inflação" e "Taxa de Juros"
df_inflacao = df.loc[inflacao_idx+1:taxa_juros_idx-1].reset_index(drop=True)
df_taxa_juros = df.loc[taxa_juros_idx+1:financas_publicas_idx-1].reset_index(drop=True)

#Remover as linhas de índice "Selic - média do ano", "TJLP (taxa nominal) - fim de periodo" e "TLP (taxa real)" na seção "Taxa de Juros"
df_taxa_juros = df_taxa_juros.drop(index=[1, 5, 6]).reset_index(drop=True)

#Identificar dinamicamente as colunas que contêm os anos desejados
colunas_desejadas = [col for col in df.columns if isinstance(col, str) and col.endswith('P') and col[:-1].isdigit()]

df_inflacao_focado = df_inflacao.loc[:, colunas_desejadas]
df_taxa_juros_focado = df_taxa_juros.loc[:, colunas_desejadas]

#Função para aplicar regressão linear e fazer a previsão para 2028
def linear_regression_forecast(series):
    X = np.array([2024, 2025, 2026, 2027]).reshape(-1, 1)
    y = series.values
    model = LinearRegression()
    model.fit(X, y)
    return model.predict(np.array([[2028]]))[0]

#Adicionar a coluna 2028P com valores previstos usando regressão linear
df_inflacao_focado['2028P'] = df_inflacao_focado.apply(linear_regression_forecast, axis=1)
df_taxa_juros_focado['2028P'] = df_taxa_juros_focado.apply(linear_regression_forecast, axis=1)

#Adicionar a coluna de indicadores
df_inflacao_focado.insert(0, 'Indicador', df_inflacao.iloc[:, 0])
df_taxa_juros_focado.insert(0, 'Indicador', df_taxa_juros.iloc[:, 0])

#Função para calcular a taxa mensal a partir da taxa anual
def calcular_taxa_mensal(taxa_anual):
    return (1 + taxa_anual)**(1/12) - 1

#Função para exibir as taxas mensais para todos os meses de cada ano em porcentagem
def exibir_taxas_mensais(df_focado):
    
    meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    
    for _, row in df_focado.iterrows():
        indicador = row['Indicador']
        for ano in colunas_desejadas + ['2028P']:
            taxa_anual = row[ano]
            taxa_mensal = calcular_taxa_mensal(taxa_anual) * 100 
            for mes in range(1, 13):
                mes_str = meses[mes - 1]
                print(f"{mes_str}/{ano[:4]} || {indicador} || Taxa: {taxa_mensal:.2f}%")
            print('-' * 50)

#Exibir as taxas mensais para cada indicador e cada ano
print("Inflacao - Taxas mensais:")
exibir_taxas_mensais(df_inflacao_focado)

print("\nTaxa de Juros - Taxas mensais:")
exibir_taxas_mensais(df_taxa_juros_focado)





