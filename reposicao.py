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

# Carregar o arquivo do Excel
excel_file2 = 'C:\\Users\\User\\Documents\\_LIGAÇÕES FALTANTES NOVO 2024.xlsx'

# Ler a planilha do Excel
df_excel2 = pd.read_excel(excel_file2, sheet_name='ABRIL')

# Lista para armazenar os resultados da comparação
resultados2 = {}

# Selecionar os nomes da coluna A onde o valor da coluna E é "NÃO"
alunos_faltantes = df_excel2[(df_excel2.iloc[:, 4] == 'NÃO') & (df_excel2.iloc[:, 1] == 'ATIVO')].iloc[:, 0].drop_duplicates().tolist()

# Criar a estrutura de dados desejada
variables3 = {f'var{i+1}': aluno for i, aluno in enumerate(alunos_faltantes)}

variables3 = list(variables3.values())
variables3.sort(key=str)

for item in variables3:
    print(item)

a = True