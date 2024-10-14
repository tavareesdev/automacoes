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

def trim_values(value):
    if isinstance(value, str):
        # Substituir dois espaços juntos por um espaço
        return re.sub('  +', ' ', value.strip())
    return value

arquivo = 'C:\\Users\\Ped\\Documents\\Relatórios\\Relatório Coordenação Julho.xlsx'

df = pd.read_excel(arquivo)

alunos_faltantes = df[(df.iloc[:, 6] == 'BLOQUEADO')].iloc[:, 0].drop_duplicates().tolist()

Inativos = {f'var{i+1}': aluno for i, aluno in enumerate(alunos_faltantes)}
Inativos = {key: trim_values(value) for key, value in Inativos.items()}
Inativos = list(Inativos.values())

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

# iframe = WebDriverWait(navegador, 10).until(
#     EC.presence_of_element_located((By.ID, "huggy-trigger-27512"))
#     and EC.visibility_of_element_located((By.ID, "huggy-trigger-27512"))
# )

# navegador.switch_to.frame(iframe)

# navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/header/div[1]/div').click()

# navegador.switch_to.default_content()
# time.sleep(2)

# iframe = WebDriverWait(navegador, 10).until(
#     EC.presence_of_element_located((By.ID, "huggy-trigger-27202"))
#     and EC.visibility_of_element_located((By.ID, "huggy-trigger-27202"))
# )

# navegador.switch_to.frame(iframe)

# navegador.find_element(By.XPATH, '/html/body/div[2]/div/div/header/div[1]/div').click()

# navegador.switch_to.default_content()
# time.sleep(2)

navegador.find_element(By.XPATH, '//*[@id="accordionMenu"]/a[1]').click()
time.sleep(2)

for item in Inativos:
    
    navegador.find_element(By.XPATH, '//*[@id="alunos"]/ul/li/ul/li[1]/a').click()
    time.sleep(3)

    campo_email = navegador.find_elements(By.XPATH, '//*[@id="search"]')
    campo_email[0].send_keys(item)
    time.sleep(4)
    try:
        navegador.find_element(By.XPATH, '//*[@id="dropdownMenuButton"]').click()
        time.sleep(3)
        navegador.find_element(By.XPATH, '//*[@id="lista-usuario"]/tr/td[5]/div/div/a').click()
        time.sleep(4)
        navegador.find_element(By.XPATH, '//*[@id="formPerfil"]/span/div/div/label[1]').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="btn_salvar_dados"]').click()
        time.sleep(2)
    except:
        pass