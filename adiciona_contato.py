from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options

# Configure o caminho do geckodriver
geckodriver_path = '/caminho/para/geckodriver'

# Configurações do Firefox
options = Options()
options.set_preference('profile', '/caminho/para/seu/perfil')

# Inicializa o WebDriver
service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=options)

driver.get('https://contacts.google.com/')
