from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import urllib
import pandas as pd

contatos_df = pd.read_excel("mensagens_terca.xlsx")
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
for i, mensagem in enumerate(contatos_df['Mensagem']):
    # if "Resp" in contatos_df.loc[i, "Pessoa"] or "RESP" in contatos_df.loc[i, "Pessoa"]:
    # Bom diaa {nome}, tudo bem?\n\n
    cont = 1
    pessoa = contatos_df.loc[i, "Pessoa"]
    numero = contatos_df.loc[i, "Número"]
    nome = contatos_df.loc[i, "Nome"]
    texto = urllib.parse.quote(f"Bom diaa {nome}, tudo bem?\n\n{mensagem}")
    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
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