from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import urllib
import pandas as pd
import random

contatos_df = pd.read_excel("Faltantes.xlsx")
print(contatos_df)

# Caminho para o Microsoft Edge WebDriver
edge_driver_path = 'caminho_para_o_executavel_do_edgedriver.exe'

# Configuração do driver do Edge
edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True

# Inicializa o driver do Edge
navegador = webdriver.Edge(options=edge_options)

navegador.get("https://web.whatsapp.com/")

while len(navegador.find_elements(By.ID, "side")) < 1:
    time.sleep(1)

# já estamos com o login feito no whatsapp web
for i, mensagem in enumerate(contatos_df['Pessoa']):
    cont = 1
    pessoa = contatos_df.loc[i, "Pessoa"]
    numero = contatos_df.loc[i, "Número"]
    nome = contatos_df.loc[i, "Nome"]
    genero = contatos_df.loc[i, "Genero"]
    data = contatos_df.loc[i, "Data"]
    nome_completo = contatos_df.loc[i, "Nome Completo"]

    lista_saudacao = ['Bom diaa', 'Eaii', 'Oii', 'Oiee']
    saudacao = random.choice(lista_saudacao)

    lista_pergunta = ['tudo bem?', 'tudo certo?', 'beleza?', 'tranquilo?']
    pergunta = random.choice(lista_pergunta)

    lista_aula = ['reposição', 'aula para repor']
    aula = random.choice(lista_aula)

    lista_humilhacao = ['Nos de um retorno por favor!!', 'Me de um retorno por favor :)', 'Fico no aguardo :)']
    humilhacao = random.choice(lista_humilhacao)
    
    if genero == 'Feminino':
        lista_apelido = ['princesa', 'jovem', 'minha amiga']
        apelido = random.choice(lista_apelido)
    else:
        lista_apelido = ['meu amigo', 'jovem', 'amigo']
        apelido = random.choice(lista_apelido)

    if "Resp" in pessoa or "RESP" in pessoa:
        primeira_mensagem = f'Boa tarde, tudo bem?'
    else:
        primeira_mensagem = f'Boa tarde {nome_completo}, {pergunta}'

    primeira_mensagem = f'{saudacao} {nome}, {pergunta}'
    segunda_mensagem = f'Você está com uma {aula} pendente conosco {apelido}, devido a sua falta dia {data}'
    confirmacao = 'Que dia você conseguiria repor? Pode ser online ou presencial'

    mensagem = urllib.parse.quote(f'{primeira_mensagem}\n\n{segunda_mensagem}\n\n{confirmacao}\n\n{humilhacao}')

    # mensagens = [
    #     primeira_mensagem,
    #     prof,
    #     segunda_mensagem,
    #     confirmacao,
    #     terceira_mensagem,
    #     aguardo
    # ]
    # for x in mensagens:

    link = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem}"
    navegador.get(link)
    time.sleep(10)
    try:
        elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button')
        elemento[0].click()
    except:
        while len(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span')) < 1:
            time.sleep(1)
        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
    time.sleep(5)