from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import urllib
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service as EdgeService

edge_driver_path = 'caminho_para_o_executavel_do_edgedriver.exe'

edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True

navegador = webdriver.Edge(options=edge_options)

navegador.get("https://web.whatsapp.com/")

while len(navegador.find_elements(By.ID, "side")) < 1:
    time.sleep(1)

contatos_df = pd.read_excel("mensagens_segunda.xlsx")

cont = 100
for i, mensagem in enumerate(contatos_df['Número']):
    if cont < 6:
        texto = contatos_df.loc[i, "Número"]
        editable_div.send_keys(texto)

        time.sleep(2)

        try:
            elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div')
            elemento[0].click()
        except:
            try:
                elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div[1]/div')
                elemento[0].click()
            except:
                try:
                    elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div/span/div/div/div')
                    elemento[0].click()
                    cont = 100

                    time.sleep(2)
                    pass
                except:
                    elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div/header/div/div[1]/div/span')
                    elemento[0].click()
                    cont = 100

                    time.sleep(2)
                    pass

        time.sleep(2)

        cont = cont + 1

    elif cont == 6:
        elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/div/div/div/div/div/div/span/div/div/div')
        elemento[0].click()
        cont = 100

        time.sleep(2)
    else:
        link = f"https://web.whatsapp.com/send?phone=+55 11 94518-9755"
        navegador.get(link)
        time.sleep(30)

        try:                                                    
            elemento_hover = navegador.find_element(By.CSS_SELECTOR, 'div.x9f619.x1d8l2j3.xm7lytj.x150jy0e.xs9asl8.x16hj40l > div._akbu > span._ao3e.selectable-text.copyable-text')
            
            actions = ActionChains(navegador)

            actions.move_to_element(elemento_hover).perform()
        except:
            try:
                elemento_hover = navegador.find_elements(By.XPATH, '//*[@id="main"]/div[3]/div/div[2]/div[3]/div[7]/div/div/div[1]/div[1]/div[1]/div[2]')

                actions = ActionChains(navegador)

                actions.move_to_element(elemento_hover[0]).perform()
            except:
                elemento_hover = navegador.find_element(By.CSS_SELECTOR, 'button[aria-label="Reproduzir mensagem de voz"]')

                actions = ActionChains(navegador)

                actions.move_to_element(elemento_hover).perform()

        navegador.implicitly_wait(5)

        elemento = navegador.find_element(By.CSS_SELECTOR, 'span[data-icon="down-context"]')
        elemento.click()

        time.sleep(2)

        elemento = navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/li[4]/div')
        elemento[0].click()

        time.sleep(2)

        try:
            elemento = navegador.find_elements(By.XPATH, '//*[@id="main"]/span[2]/div/button[5]')
            elemento[0].click()
        except:
            try:
                elemento = navegador.find_elements(By.XPATH, '//*[@id="main"]/span[2]/div/button[5]/span')
                elemento[0].click()
            except:
                elemento = navegador.find_element(By.CSS_SELECTOR, 'span[data-icon="forward"]')
                elemento.click()
                
        time.sleep(2)

        editable_div = navegador.find_element(By.CSS_SELECTOR, 'div[contenteditable="true"]')

        editable_div.click()

        time.sleep(2)

        cont = 1

time.sleep(5)

navegador.quit()