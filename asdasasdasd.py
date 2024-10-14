from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

# Configuração do WebDriver para o Microsoft Edge
options = webdriver.EdgeOptions()
driver = webdriver.Edge(options=options)

# Abrir a primeira guia com o YouTube
driver.get('https://www.youtube.com')
time.sleep(3)  # Espera 3 segundos para a página carregar

# Abrir uma nova guia com o Google
driver.execute_script("window.open('https://www.google.com', '_blank');")
time.sleep(3)  # Espera 3 segundos para a página carregar

# Alternar para a segunda guia (Google)
driver.switch_to.window(driver.window_handles[1])
print("Agora na guia do Google")
time.sleep(3)  # Espera 3 segundos

# Identificar o campo de pesquisa pelo ID na guia do Google
search_box = driver.find_element(By.NAME, 'q')  # ID do campo de pesquisa do Google
search_box.send_keys('Selenium WebDriver')  # Digitar 'Selenium WebDriver' na caixa de pesquisa
search_box.send_keys(Keys.RETURN)  # Pressionar Enter para pesquisar

time.sleep(3)  # Espera 3 segundos para ver o resultado da pesquisa

# Alternar para a primeira guia (YouTube)
driver.switch_to.window(driver.window_handles[0])
print("Agora na guia do YouTube")
time.sleep(3)  # Espera 3 segundos

# Fechar o navegador
driver.quit()
