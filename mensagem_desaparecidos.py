import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import urllib
import pandas as pd

contatos_df = pd.read_excel("C:\\Users\\gtava\\OneDrive\\Documentos\\Automacao\\Desaparecidos.xlsx")
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
    # if "Resp" in contatos_df.loc[i, "Pessoa"] or "RESP" in contatos_df.loc[i, "Pessoa"]:
    # Bom diaa {nome}, tudo bem?\n\n
    envio = "Olá, tudo bem?\n\nNotamos que o(a) " + contatos_df.loc[i, "Pessoa"] + " está ausente das aulas há algum tempo e gostaríamos de saber se está tudo bem. Sentimos a falta dele(a) e estamos preocupados.\n\nSe estiverem enfrentando alguma dificuldade ou precisarem de suporte, estamos aqui para ajudar. Queremos garantir que o(a) aluno(a) esteja participando das aulas e aproveitando ao máximo o conteúdo.\n\nPodem nos dar um retorno para que possamos entender melhor a situação e oferecer o apoio necessário? Contamos com vocês!\n\nAbraços,Equipe Pedagógica."
    pessoa = contatos_df.loc[i, "Pessoa"]
    numero = contatos_df.loc[i, "Número"]
    nome = contatos_df.loc[i, "Nome"]
    texto = urllib.parse.quote(envio)
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

    def enviar_email():
        # Configurações do servidor SMTP
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        smtp_user = "gtavares744@gmail.com"
        smtp_password = "owmt bryn qcgq wnpr"

        # Informações do email
        remetente = "gtavares744@gmail.com"
        destinatario = contatos_df.loc[i, "Email"]
        assunto = "Retorno do Aluno"

        # Criação da mensagem
        mensagem = MIMEMultipart()
        mensagem['From'] = remetente
        mensagem['To'] = destinatario
        mensagem['Subject'] = assunto

        # Anexar o corpo do email
        mensagem.attach(MIMEText(envio, 'plain'))

        try:
            # Conectar ao servidor SMTP
            servidor = smtplib.SMTP(smtp_server, smtp_port)
            servidor.starttls()  # Iniciar TLS para segurança
            servidor.login(smtp_user, smtp_password)  # Logar no servidor

            # Enviar email
            servidor.sendmail(remetente, destinatario, mensagem.as_string())

            print("Email enviado com sucesso!")
        except Exception as e:
            print(f"Erro ao enviar email: {e}")
        finally:
            servidor.quit()  # Fechar a conexão com o servidor

    if __name__ == "__main__":
        enviar_email()