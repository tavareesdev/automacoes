from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import time
import os
import glob
import re
import pandas as pd
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import Border, Side
from openpyxl import Workbook

edge_driver_path = 'caminho_para_o_executavel_do_edgedriver.exe'

edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True

navegador = webdriver.Edge(options=edge_options)

navegador.get("https://administracao.evoluaprofissional.com.br/")

time.sleep(2)
campo_email = navegador.find_elements(By.XPATH, '//*[@id="Login"]')
campo_email[0].send_keys("easytech.osasco@gmail.com")
time.sleep(2)

campo_senha = navegador.find_elements(By.XPATH, '//*[@id="Senha"]')
campo_senha[0].send_keys("$ucesso3331")
campo_senha[0].send_keys(Keys.ENTER)
time.sleep(2)

navegador.find_element(By.XPATH, '//*[@id="accordionMenu"]/a[6]').click()
time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="pedagogico"]/ul/li[11]/a').click()
time.sleep(2)

data1 = navegador.find_elements(By.ID, "dt_inicial")
data_antiga = '01/01/2020'
data1[0].send_keys(data_antiga)
time.sleep(2)

data_de_hoje = datetime.now().date()
data_de_ontem = data_de_hoje - timedelta(days=1)
data_de_ontem = data_de_ontem.strftime("%d/%m/%Y")
data2 = navegador.find_elements(By.ID, "dt_final")
data2[0].send_keys(data_de_ontem)
navegador.find_element(By.XPATH, '/html').send_keys(Keys.ESCAPE)
time.sleep(3)

navegador.find_element(By.XPATH, '//*[@id="data"]/div[2]/div/button').click()

# Página Situação
time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="situacao"]/div/div/div[1]/div/div/div/div[2]/a').click()

time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="situacao"]/div/div/div[2]/div/button').click()

# Página Organograma
time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="organograma"]/button').click()

time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="organograma"]/div[2]/div/button').click()

time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="heading-0"]/div').click()

time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="alunos"]/button').click()

time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="alunos"]/div[2]/div[1]/button').click()

time.sleep(90)

# Obtém o diretório de download
diretorio_download = "C:\\Users\\User\\Downloads"

arquivos = os.listdir(diretorio_download)

# Seleciona o arquivo mais recente com extensão .xlsx
arquivo_excel = max([f for f in arquivos if f.endswith('.xlsx')], key=lambda x: os.path.getmtime(os.path.join(diretorio_download, x)))

# Caminho completo do arquivo Excel
caminho_arquivo_excel = os.path.join(diretorio_download, arquivo_excel)

# Lê o arquivo Excel usando o pandas
df = pd.read_excel(caminho_arquivo_excel)

def trim_values(value):
    if isinstance(value, str):
        # Substituir dois espaços juntos por um espaço
        return re.sub('  +', ' ', value.strip())
    return value

# Ler os nomes da coluna D e armazenar em variáveis
variables = {}
for i, value in enumerate(df['Aluno']):
    if value not in variables.values():
        variables[f'var{i+1}'] = value

variables = {key: trim_values(value) for key, value in variables.items()}

variables = list(variables.values())

os.remove(caminho_arquivo_excel)

data_faltantes = {}

cont = 1


for item in variables:
    if cont < 11: 

        navegador.find_element(By.XPATH, '//*[@id="pedagogico"]/ul/li[2]/a').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="filtrar_por_aluno noPrint"]/div/div/button').click()
        time.sleep(2)
        nome_aluno = navegador.find_element(By.XPATH, '//*[@id="filtrar_por_aluno noPrint"]/div/div/div/div[1]/input')
        nome_aluno.send_keys('Fernanda Alves de Oliveira')
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="filtrar_por_aluno noPrint"]/div/div/div/div[2]/ul/li/a').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="btn_filtrar"]').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="btn_imprimir"]').click()
        time.sleep(2)
        navegador.send_keys(Keys.ENTER)
        time.sleep(2)
        navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/div/div[3]/div/button[1]').click()


        

        



        # Obtém o diretório de download
        diretorio_download = "C:\\Users\\User\\Downloads"

        arquivos = os.listdir(diretorio_download)

        # Seleciona o arquivo mais recente com extensão .xlsx
        arquivo_excel = max([f for f in arquivos if f.endswith('.xlsx')], key=lambda x: os.path.getmtime(os.path.join(diretorio_download, x)))

        # Caminho completo do arquivo Excel
        caminho_arquivo_excel = os.path.join(diretorio_download, arquivo_excel)

        # Lê o arquivo Excel usando o pandas
        df_aluno = pd.read_excel(caminho_arquivo_excel)
        print(df_aluno)
        
        ultimo_valor = df_aluno['Data'].iloc[-1]

        data_faltantes[f'var{i+1}'] = ultimo_valor + ' ' + item

        os.remove(caminho_arquivo_excel)

        cont = cont + 1

a=1