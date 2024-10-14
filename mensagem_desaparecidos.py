from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
import time
import os
import tabula
import pandas as pd
from datetime import datetime, timedelta
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from pdf2image import convert_from_path
import pytesseract
import re
from dateutil.relativedelta import relativedelta
import urllib
import xml.etree.ElementTree as ET
from pathlib import Path
import glob
from selenium.common.exceptions import NoSuchElementException
pyautogui.FAILSAFE = False

DOWNLOAD_DIR = str(Path.home() / "Documents")

def get_latest_downloaded_pdf(download_folder):
    # Usando glob para listar todos os arquivos PDF no diretório de downloads
    list_of_files = glob.glob(os.path.join(download_folder, '*.pdf'))
    
    if not list_of_files:
        raise FileNotFoundError("Nenhum arquivo PDF encontrado no diretório de downloads.")

    # Obtendo o arquivo mais recente com base no tempo de modificação
    latest_file = max(list_of_files, key=os.path.getmtime)
    
    return latest_file

def find_latest_pdf(directory):
    files = list(Path(directory).glob("*.pdf"))
    if not files:
        return None
    latest_file = max(files, key=lambda f: f.stat().st_mtime)
    return latest_file

def ler_xlsx_para_dataframe(caminho_xlsx):
    try:
        df = pd.read_excel(caminho_xlsx)
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo XLSX: {e}")
        return None

def trim_values(value):
    if isinstance(value, str):
        return value.strip()  # Remove espaços no início e no final, mas preserva espaços entre palavras
    return value  # Retorna o valor original se não for uma string

# Configurações do Webdriver
service = webdriver.edge.service.Service('C:\\Users\\Ped\\anaconda3\\msedgedriver.exe')  # Substitua pelo caminho correto
options = webdriver.EdgeOptions()
options.add_argument('--kiosk-printing')  # Permite a impressão automática
options.add_argument('--print-to-pdf')  # Configura o Edge para imprimir em PDF automaticamente
options.add_argument('--disable-gpu')  # Desativa a GPU para evitar problemas gráficos
options.add_argument('--no-sandbox')  # Desativa o sandbox para evitar problemas de segurança

# Inicializa o Webnavegador
navegador = webdriver.Edge(service=service, options=options)

# Acesse o site
navegador.get("https://administracao.evoluaprofissional.com.br/")

# Login
time.sleep(2)
campo_email = navegador.find_element(By.XPATH, '//*[@id="Login"]')
campo_email.send_keys("easytech.osasco@gmail.com")
time.sleep(2)

campo_senha = navegador.find_element(By.XPATH, '//*[@id="Senha"]')
campo_senha.send_keys("$ucesso3331")
campo_senha.send_keys(Keys.ENTER)
time.sleep(2)

caminho_arquivo_excel = "C:\\Users\\Ped\\Documents\\Relatórios\\Relatório Coordenação Setembro - Copia.xlsx"

# Lê o arquivo Excel usando o pandas
df = pd.read_excel(caminho_arquivo_excel)

# Filtra os alunos cuja situação é "ATIVO"
df_ativos = df[(df['Situação'] == 'DESAPARECIDO') & (df['Meses desde primeiro acesso'] < 19)]

# Armazena os nomes dos alunos ativos em variáveis
variables = {}
for i, value in enumerate(df_ativos['Aluno']):
    if value not in variables.values():
        variables[f'var{i+1}'] = value

variables = {key: trim_values(value) for key, value in variables.items()}

variables = list(variables.values())

df_acessos = ler_xlsx_para_dataframe(caminho_arquivo_excel)

if df_acessos is not None:
    # Remove espaços em branco extras dos nomes dos alunos e substitui múltiplos espaços por um único espaço
    df_acessos['Aluno'] = df_acessos['Aluno'].str.strip()

    # Processa os dados para encontrar a primeira data de acesso para cada aluno
    df_acessos['Data de Primeiro Acesso'] = pd.to_datetime(df_acessos['Data de Primeiro Acesso'], errors='coerce')  # Garante que a coluna 'Data de Primeiro Acesso' seja do tipo datetime
    primeiras_datas_acesso = df_acessos.groupby('Aluno')['Data de Primeiro Acesso'].min().reset_index()
    data_primeiro_acesso = dict(zip(primeiras_datas_acesso['Aluno'], primeiras_datas_acesso['Data de Primeiro Acesso']))
else:
    print("Não foi possível ler o arquivo XLSX.")
    data_primeiro_acesso = {}

cont = 1

alunos_processados = []
alunos_enviados = []

for item in variables:
    if item not in alunos_enviados:

        alunos_processados.append(item)
        print(alunos_processados)

        primeira_data = data_primeiro_acesso[item]

        if cont > 1: 
            navegador.switch_to.window(navegador.window_handles[0])
        else:
            navegador.find_element(By.XPATH, '//*[@id="accordionMenu"]/a[7]').click()   
            time.sleep(2)
        # Navegação até a página desejada
        navegador.find_element(By.XPATH, '//*[@id="pedagogico"]/ul/li[2]/a').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="filtrar_por_aluno noPrint"]/div/div/button').click()
        time.sleep(2)
        nome_aluno = navegador.find_element(By.XPATH, '//*[@id="filtrar_por_aluno noPrint"]/div/div/div/div[1]/input')
        nome_aluno.send_keys(item)
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="filtrar_por_aluno noPrint"]/div/div/div/div[2]/ul/li/a').click()
        time.sleep(10)
        navegador.find_element(By.XPATH, '//*[@id="btn_filtrar"]').click()
        time.sleep(7)
        navegador.find_element(By.XPATH, '//*[@id="btn_imprimir"]').click()
        time.sleep(2)

        # Simula Ctrl+P para abrir o diálogo de impressão
        action = ActionChains(navegador)
        action.key_down(Keys.CONTROL).send_keys('p').key_up(Keys.CONTROL).perform()
        time.sleep(7)  # Aumente o tempo se necessário

        # Use pyautogui para interagir com o diálogo de impressão
        pyautogui.write(f'Historico {item}.pdf')  # Nome do arquivo com a primeira letra maiúscula
        time.sleep(2)

        pyautogui.press('enter')
        time.sleep(15)

        diretorio_download = "C:\\Users\\Ped\\Documents"

        # Caminho do arquivo PDF
        download_path = get_latest_downloaded_pdf(diretorio_download)
        # Converter apenas a primeira página do PDF para imagem
        try:
            # Convertendo apenas a primeira página (páginas[0])
            paginas = convert_from_path(download_path, first_page=1, last_page=1)
        except Exception as e:
            raise RuntimeError(f"Erro ao converter o PDF: {e}")

        texto_extraido = ""
        # Extrair texto da primeira página
        if paginas:
            texto = pytesseract.image_to_string(paginas[0])  # Extrai apenas da primeira página
            texto_extraido += texto

            # Expressão regular para encontrar números de telefone no formato (11)9200-69221 e (11)9767-96364
            phone_pattern = re.compile(r'\(\d{2}\)\d{4,5}-\d{5}')

            # Encontra todos os números de telefone no texto
            phones = phone_pattern.findall(texto_extraido)

            # Verifica se encontramos dois números e os armazena em variáveis separadas
            if len(phones) >= 2:
                phone1 = phones[0]
                phone2 = phones[1]
            elif len(phones) >= 1:
                phone1 = phones[0]
            else:
                print("Nenhum número de telefone foi encontrado.")

            envio = f"Olá {item},\n\nEspero que esteja tudo bem com você. Notei que faz um tempo que não temos notícias suas, e estou preocupado com sua ausência. Por favor, me avise como está e se há algo com o qual eu possa ajudar. Seu retorno é muito importante para nós, e ficaria aliviado ao saber como você está.\n\nAguardo sua resposta o mais breve possível.\n\nAtenciosamente,\nEquipe Pedagógica."
           
            if cont > 1:
                navegador.switch_to.window(navegador.window_handles[1])
            else:
                navegador.execute_script("window.open('https://web.whatsapp.com/', '_blank');")
                navegador.switch_to.window(navegador.window_handles[1])

            pdf_file = find_latest_pdf(DOWNLOAD_DIR)
            if not pdf_file:
                print("Nenhum arquivo PDF encontrado.")
                exit()

            while len(navegador.find_elements(By.ID, "side")) < 1:
                time.sleep(1)

            # já estamos com o login feito no whatsapp web
            if len(phones) >= 1:
                for phone in phones:
                    texto = urllib.parse.quote(envio)
                    link = f"https://web.whatsapp.com/send?phone={phone}&text={texto}"
                    navegador.get(link)
                    time.sleep(20)
                    try:
                        elemento = navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[2]/button')
                        elemento[0].click()  
                        time.sleep(2)
                    except:
                        while len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button')) < 1:
                            time.sleep(1)
                        navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button').send_keys(Keys.ENTER)
                    time.sleep(5)
            cont = cont + 1

# Fechar o navegador
navegador.quit()