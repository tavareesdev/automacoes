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

caminho_arquivo_excel = "C:\\Users\\Ped\\Documents\\Relatórios\\Relatório Coordenação Julho.xlsx"

# Lê o arquivo Excel usando o pandas
df = pd.read_excel(caminho_arquivo_excel)

# Filtra os alunos cuja situação é "ATIVO"
df_ativos = df[(df['Situação'] == 'ATIVO') & (df['Meses desde primeiro acesso'] < 17)]

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
alunos_enviados = ['Adailson Marques Correa', 'Adila Vitória Xavier Gonçalves', 'Adriano Gustavo da Silva Ribeiro','Adrielle Nunes Da Cruz', 'Adrielly Vitoria  Goulart Dos Santos', 'Aelton Pereira de Sousa', 'Agatha  Sophia Dos Santos Nascimento', 'Alexandre  Augusto Pessoa Rosa', 'Alexandre da Conceição', 'Alexsander Viera Santos', 'Alfredo  Wallyson Schiffner', 'Alice Alves Ferreira',
'Alice santos da silva', 'Alicia Silva Nascimento', 'Aline Rocha Soares', 'Allan Salvador Pereira da Silva', 'Allifer  Kauã Rodrigues Costa', 'Alvaro  Bezerra Mendes', 'Alycya Nunes dos santos', 'Amanda Menezes De Almeida', 'Ana  Beatriz Lescano Dos Santos', 'Ana  Vitoria Paulino Dos Santos', 'Ana Alice da Silva Santos', 'Ana Beatriz Araujo Torquato', 'Ana Beatriz Silveira Santos', 'Ana Beatriz de Sousa Marquês', 'Ana Caroliny Silva de Souza', 'Ana Clara  Jesus dos Santos', 'Ana Clara  Joana do Nascimento', 'Ana Julia  Santos Souza', 'Ana Julia Dos santos Silva', 'Ana Julia da Silva Cruz', 'Ana Kesia Santos Silva Herculano', 'Ana Luisa Neves Pereira', 'Ana Paula  Corte Madeiro', 'Ana Paula Dos Santos Barbosa', 'Ana Victoria Oliveira de jesus', 'Ana Vitoria dos Santos',
'Anderson Messias de Sousa Moura', 'Andressa Katlin da Silva da Pereira','Andrey de Souza  Santos', 'Andrielle dos santos Alves', 'Anne Victoria  Alves Rocha', 'Antonio  Everton Pereira Da Silva', 'Antonio Gabriel Holanda Nobre De Oliveira',
'Any Caroline  Alves dos Santos', 'Arthur  de souza cruz', 'Arthur Ricardo  Felisberto De Souza', 'Arthur Santos da Costa', 'Augusto  da Silva','Beatriz  Vitoria Santos Da Silva','Beatriz Nunes Dos Santos','Beatriz Santana Amorim dos Santos', 'Bianca Pereira leal Vitório', 'Bianca dos Santos Oliveira', 'Breno Cardoso de Britto Filho', 'Brunna Eduarda Da Silva Herrera', 'Bruno  Ferreira Rodrigues', 'Bruno Santos da Cruz', 'Bruno Willian Dos Santos Silva', 'Bryan  Feliz De Almeida', 'Bryan  Nicolas Silva Santos', 'Bryan Lucas Lira Galvão', 'Bryan Xavier do Nascimento', 'Budelyne Pierre', 'Cailane  Silva De Souza', 'Caio  Fagundes Alquimim', 'Caio Santos Ramires', 'Camila Ribeiro De Jesus Silva', 'Camilly Silva Lopes', 'Carlos  Eduardo Rodrigues', 'Carlos Daniel Da Silva', 'Carlos Eduardo  Santos Mamede', 'Carlos Eduardo Viana Rocha da Silva', 'Carlos Gabriel Amaral de Oliveira', 'Carolina Lopes Melo', 'Catarina  Moreira Conceição',
'Catarina Santana de Camargo', 'Cecilia De Santana Carvalho', 'Cinthya Gabrielly Prates De Azevedo', 'Clara Cerqueira', 'Claudio  Victor Menezes Antunes','Claudio Prazeres da Silva', 'Daniel Costa Da Silva',
'Daniel Rodrigues Marques de Lima', 'Danilo Barreto de Souza', 'Danilo Santos Moreno', 'Davi  Azevedo da Silva', 'Davi  De Albuquerque Lopes', 'Davi  Souza Nascimento', 'Davi Luiz Santos Bertoncini', 'David  de Souza correia', 'Dayron  Pereira dos Santos', 'Diego Leandro  de Lima', 'Diego Melo da Silva', 'Diego Rufino da Silva', 'Diogo Albiero', 'Débora Assunção Zeferino', 
'Edicassio  Nery Marques','Eduardo  Furtado De Lima', 'Eduardo Gomes Ferreira', 'Eduardo Henrique  Dos Santos', 'Eloa Camargo de souza', 'Eloane  Aparecida De Lima Ramos', 'Eloisa Tereza De Lima Ramos', 'Emanuella  Amaral de Souza', 'Emily Santos Araújo','Enzo Cardoso de Oliveira', 'Enzo Melo Dos Santos', 'Enzo Mendes  Pereira', 'Erick  Rian Romão Ferreira da Silva', 'Erick Demian Arnaut Sinsuk', 'Erick Marques Francisco', 'Esther  Lima Andrade', 'Evellyn gomes Silva Souza', 'Felipe  Henrique  Wutenrg Leite', 'Felipe Carvalho', 'Felipe Gomes Nascimento', 'Felipe Reis  de Lima', 'Felipe Rodrigues Nascimento', 'Felippe Lessa dos Santos', 'Fernanda Alves dos Santos', 'Fernando  Angelo Cunha Symphoroso', 'Flavia  keitty', 'Francisco Alessandro Mendes Bringel', 'Francisco Lucas  Vicente Alves', 'Francisco Riquelme Duarte Silva', 'Gabriel  Roque dos Santos', 'Gabriel  Santos Pereira', 'Gabriel  Silva Rodrigues Custódio', 'Gabriel Aleixo Rodrigues', 'Gabriel Antonio  Aparecido Silva', 'Gabriel Barbosa Boer', 'Gabriel Francisco  Almeida Dos Santos', 'Gabriel Leal Dos Santos', 'Gabriel Linhares Magalhães', 'Gabriel Xavier Da Silva', 'Gabriela  Marcelino', 'Gabriela de Jesus Sena', 'Gabrielle  Sales do Couto'
,'Gabrielle Conceição dos Santos', 'Gabrielly  Rodrigues Gouvea', 'Gabrielly Silva Lopes', 'Geovana  Oliveira Da Silva', 'Geovana Antunes Dos Santos', 'Geovanna  Souza de Jesus', 'Geovanna Félix Monteiro', 'Geovanna Souza Santos', 'Geovanna de Freitas Fernandes', 'Geyse  de Amorim Rodrigues', 'Ghiovanna Gonçalves Da Silva', 'Giovana  santos de souza', 'Giovanna  Aparecida Dos Santos', 'Giovanna Andrielly Pereira Dos Santos', 'Giovanna Freitas de Sousa', 'Giovanna Gabriel Rodrigues', 'Giovanni Aureliano de Assis', 'Gislaine Vitória Da Silva', 'Glenda medeiros rocha dos santos', 'Graziele  dos Santos', 'Guilherme  Augusto Morais leite', 'Guilherme  Lima Menezes', 'Guilherme Campos Mascarenhas Soares', 'Guilherme Henrique Candido Oliveira', 'Guilherme Miguel Rodrigues da Silva', 'Guilherme Pereira de Souza', 'Guillyane Santos  Silva', 'Gustavo  Da Silva Sizisnande', 'Gustavo  Henrique Viera De Sousa', 'Gustavo  Vital Santana', 'Gustavo Aparecido Sales', 'Gustavo Reis  de Lima', 
'Gyovana Do Nascimento Costa', 'Hebert Conceição Reis da Cruz', 'Heloisa  Barbosa Bonfim', 'Hemilly  Rodrigues dos Santos', 'Henrique  Alves de Araújo', 'Higor  Pires Silva', 'Hudson Machado Oliveira', 'Hugo Lima Gomes', 'Hugo Xavier Silva Fidelis do Nascimento', 'Iago Rafael Fagundes Pinto', 'Iara francisco haddad', 'Iasmym Vitoria Barroso Ferreira', 'Ingrid  Vitoria Bento Oliveira', 'Iris Maria  da Silva', 'Isabella  Cristina de Jesus', 'Isabelli Martins de Oliveira', 'Isabelly Barreto Lima', 'Isabelly Lopes Piris',
'Isabely Menegasso Ribeiro', 'Isabely Pereira Dos Santos', 'Isabely Rodrigues  Medeiros', 'Isaias Piaulino Rocha', 'Ivana Gabrielly Alves de Sousa', 
'Jady Clara Urbano de Oliveira', 'Jamile dos Santos  Vieira', 'Jamily  De Lima', 'Jamily Barbosa Rodrigues', 'Jamily Vitoria Oliveira Silva', 'Jefferson Raonny Pereira Veloso', 'Jefferson Santos de Carvalho', 'Jenifer  das Neves Costa', 'Jheniffer Sousa Dantas', 'Jhonata Gabryell Aquino Xavier', 'Jhonatan  Santos Santana', 'Jhonny Gabriel da silva Santos', 'Jonas Souza da Silva', 'Jonatas  de Souza Santos', 'José Aldevir Teixeira Paz', 'José Emanuel Pereira dos Santos', 'José Henrique dos Santos Luiz', 'José Lucas  Anacleto da Silva', 'Joyce de Oliveira Moraes', 'João  Victor Leite',
'João  Victor Sousa Santos', 'João  Vittor Santos de Sousa', 'João Pedro  Ferreira Frankin Alves', 'João Victor  Da Silva Holanda', 'João Vitor Batista da Silva', 'João Vitor Moitinho dos Santos','Juan Oliveira de Medeiros',
'Julia  Aparecida Da Silva Nascimento', 'Julia Gabriele Martins da Silva', 'Julia Vitorino Marcilio', 'Juliana Nunes da Silva', 'Julio César  Machado Ramos', 'Julio Vitor Juvenal Barbosa', 'Jussara de Almeida Santos', 'Júlia  da Silva Quero', 'Júlia Augusta  Rodrigues dos Santos', 'Kaik Cartaxo Romano', 'Kaio Henrique araujo dos santos', 'Kaique  de Souza Godinho', 'Kamilly  Vitoria Firmino da Silva', 'Karina Da silva Gonçalves', 'Kauan Oliveira Aragão', 'Kauan Ronaldo Domingos Vasconcelos', 'Kauane  de Araújo Ricarte', 'Kauane Santos Pereira', 'Kauã Alves  Clementino', 'Kauã De Oliveira Santos', 'Kauã Silva Souza', 'Kauã almeida tavares', 'Kauã da silva milani', 'Kayke Vieira da Silva Sá', 'Kaylane da Silva', 'Keilla Vitoria Nascimento Dos Santos', 'Kelvis  Da Silva Santos', 'Kelvyn Henrique Neri Marciano', 'Kemilly  Marangoni', 'Ketelin Lorrany dos Santos', 'Kethellen Isabella Marangoni dos Santos', 'Ketlyn  Milena Macedo Nascimento', 'Kevin De Araujo Santos', 'Kezia De Oliveira Da Silva', 'Laisná de Sousa  santos', 'Lana Hikari Mori','Lara Midori Oshiro Moreiro',
'Larissa Barbosa Bacelar', 'Larissa Rosseti Rodrigues','Laura  Cordeiro Guerreiro', 'Laura  Victoria Santos Pereira', 'Laura Aparecida  Torres Pereira', 'Leandro  de Jesus dos Santos','Leticia  Lopes dos Santos', 'Leticia Alanis Araujo de Brito', 'Leticia Nunes De Sousa', 'Leticia castro Guimarães','Letycia  Helena Alves de Godói', 'Letícia  Maciel Roberti De Matos Gordo', 'Lohan Andrade Chaves', 'Lorane Gonçalves de Santana', 'Luana  Aparecida Alves Moisinho', 'Luara Azevedo Heubene Fernandes', 'Lucas  Alves Dos Santos', 'Lucas  Gabriel Lima Ribeiro', 'Lucas  Genofre Tavares', 'Lucas  Ribeiro De Sousa', 'Lucas  Santana Gomes', 'Lucas  Silva de Mendonça', 'Lucas Domingos Silva', 'Lucas Lima Da Silva', 'Lucas Ryan Araujo Venceslau', 'Lucas da Silva Barbosa', 'Luis Eduardo  Rangel Matos', 'Luiz Fernando Alves de Oliveira', 'Lyvia  Araujo De Barros', 'Lyvia  Lima Da Silva', 'Manuela  Nunes da Silva', 'Marcella Vitória Correa', 'Marco Antônio Alves Vieira', 'Marcus Vinicius Barbosa Diogo', 'Marcus Vinicius Santos Lima', 'Maria  Clara Cavalcante Martins', 'Maria  Clara Ferreira De Oliveira', 'Maria  Eduarda Andrade De Sousa', 'Maria  Eduarda Lelis Marques'
, 'Maria  Fernanda Lima Da Cruz', 'Maria  Vitoria Ribeiro Santos', 'Maria Alice Almeida Prado', 'Maria Antonia Ferreira dos Santos', 'Maria Candida Marques Da Gama', 
'Maria Clara  Da Silva Feitosa', 'Maria Clara Nunes da Silva', 'Maria Clara de Paula Sales', 'Maria Eduarda  Alves oliveira', 'Maria Eduarda  França de Sousa', 'Maria Eduarda  de Souza Aguiar', 'Maria Eduarda Bernardo Carvalho', 'Maria Eduarda Buriti Ferreira', 'Maria Eduarda Cavalcanti De Oliveira', 'Maria Eduarda Ferreira Gonçalves', 'Maria Eduarda Moreira Mergulhão', 'Maria Eduarda Satú dos Santos',
'Maria Eduarda Vieira Claudio', 'Maria Hevelly Batista Vasconcelos', 'Maria Isabela Rodrigues dos Santos', 'Mariana  Santos Bispo', 'Mariana Janice da Silva', 'Mariana Rodrigues Da Silva', 'Marianna Vitória dos Santos', 'Mateus  Pereira Dos Santos', 'Mateus Pedrosa da Silva', 'Matheus  Danilo Conceição Moreira', 'Matheus  De Luchio Custodio Rodrigues Oliveira', 'Matheus  da Silva Alves', 'Matheus Gabriel  Silva de Oliveira', 'Matheus Lopes  Gomes', 'Matheus Paulino Rodrigues Da Silva', 'Matheus Ramos Miranda', 'Matheus Santos de Oliveira', 'Matilde Antonia Firmino Silva', 'Maxwell Souza Cruz', 'Maysa Lorraine Izebel de Oliveira', 'Maysa de Souza Oliveira', 'Micael  Ortiz Botte Gomes', 'Michel Adilson  Dos Santos', 'Michelly Silva de Santana', 'Miguel  Feitosa Da Silva', 'Miguel Andrade de Lima', 'Miguel Assis Ferreira', 'Miguel Dias  Alves Da Silva', 'Miguel Prado Moreira', 'Millena Ramos de Freitas', 'Monica  Araceli Mujica Avalo', 'Monique  Alves Da Silva', 'Naiara vieira oliveira', 'Natan Gabriel Santos Amador', 'Nathalia  de Oliveira Sales', 'Nathalia Beatriz Mendes de Oliveira', 'Nathan De Medeiros Martins', 'Nathanael  Lins Cavalcante', 'Natã Lucas De Lima', 'Natãn  Ferreira De Souza', 'Nayane Barbosa Silva', 'Nayara Da Silva Carvalho', 'Nicolas  Eduardo Mendes dos Santos', 'Nicolas Araújo Martins', 'Nicollas Dias Santana', 'Nicollas Ribeiro da silva', 'Nicolly Emilly Machado', 'Nicolly Raphaela Cardoso Gonçalves', 'Nikolas Alexandre de Souza santos', 'Noelma Vitória Vieira Nogueira', 'Nycolas Soares Melo de Souza', 'Nycolas dos Santos Silva', 'Otavio Eduardo Almeida Farias Correia', 'Paloma Farias Alencar', 'Pamela sandreia  cordeiro de oliveira', 'Paula Dos Santos   Damascena', 'Paulo Henrique Dias Lima', 'Pedro  Lucca Lopes Reis', 'Pedro  Paixão Alves', 'Pedro Henrique  Costa Massaro', 'Pedro Henrique  Pereira Sampaio', 'Pedro Henrique Paiva Barbosa santos', 'Pedro Henrique de Souza Barbosa', 'Pedro Henrique fernandes marques', 'Pedro Raphael de Souza Alves', 'Pedro de Araujo Veras', 'Poliana  Lago Pereira', 'Rafael  Fernandes Pinto', 'Rafaela  Soares Souza', 'Rafaella  Vitória Raimundo', 'Raffaela Vitoria de Souza Nascimento', 'Raul Ravy lopes da silva', 'Rayssa de Jesus Santos', 'Rebeca Natyelli Correa', 'Ricardo  Araujo De Andrade Lima', 'Richard  henrique saraiva marioto', 'Richard Ferreira Andrade'
,'Robert  Willians Andrade Domingos', 'Roberto  Silva de Almeida', 'Rogério Costa Goulart', 'Ruthe Emanuelly de Assis Santos', 'Ruân da Silva Alves', 
'Ryan  Souza Dos Santos', 'Ryan De Souza Felix', 'Ryan Martins da Silva Tobias', 'Ryan Oliveira dos Santos', 'Samara Vick Da Silva Mendonça', 'Samuel  De Souza Tomé'
, 'Samuel  Fernandes De Santana', 'Samuel Silva Carvalho', 'Samuel Silva De Oliveira Moraes', 'Samuel Victor da Silva melo', 'Sara  de Oliveira Lima', 
'Sarah Rocha Vieira Dos Santos', 'Shirley Alves Santos', 'Shirley heloysa lopes de souto', 'Sofia Santos Rodrigues', 'Sonia  Luiza de Oliveira Damaceno', 'Sophia  Cirilo De Souza', 'Sophia Elem Pereira Da Silva', 'Stefany  Martins De Lima', 'Stefany dos Santos Sena', 'Taina  Vitória Bomfim Silveira', 'Tatiana Alves Da Silva', 'Thainan santos de jesus', 'Tharyk Augusto Pereira De Oliveira', 'Thauany Da silva Carvalho', 'Thaueemily Leal Martins', 'Thaylla Miranda Amorim', 'Thayna Amaral da Slva', 'Thiffany Da Silva Santos da Conceição', 'Valeria  Cauany Santos Souza', 'Victor  Araújo Silvestre', 'Victor Hugo Gomes Lisboa', 'Victor Nogueira de Araújo', 'Victtoria Xavier Spadacini', 'Vinicius  Eduardo Flor Dos Santos', 'Vinicius  Galdino De Oliveira', 'Vinicius  Mota Alencar', 'Vinicius  Rick Nascimento Silva', 'Vinicius  Tavares Ferreira', 'Vinicius Alexandre Lia', 'Vinicius Miguel Marques de Oliveira Ribeiro', 'Vinícius Matheus da Silva Lima', 'Vitor  Zamorano Silva Braz', 'Vitoria  Arruda Gomes Martins', 'Vitoria Bela da Silva', 'Vitoria Duran Da Silva', 'Vitoria pereira martins', 'Vitória   Borges Pereira', 'Vitória Alves dos Santos', 'Viviane Souza  Barbosa', 'Vyctor Alexandre  Araújo', 'Wellington Bezerra Dos Reis', 'Wellington Silva Oliveira', 'Wendel Alexandrino da Silva', 'Wendy Angelica Costa Oliveira'
]

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

        # Extrair a primeira data e curso
        match = re.search(r'(\d{2}/\d{2}/\d{2}) \d{2}:\d{2} (.+?)\n', texto_extraido)
        cursos = re.findall(r'Curso\s*\n\n(.*?)\n', texto_extraido, re.DOTALL)

        if match:
            formato_data = '%d/%m/%Y'
            curso = match.group(2)
            data = primeira_data  # Converte a string da data em objeto datetime
            
            # Calcula a diferença em meses
            hoje = datetime.now()
            diff = relativedelta(hoje, data)
            diff_meses = diff.years * 12 + diff.months

            print(f"Curso: {curso}")
            print(f"Data: {data}")
            print(f"Diferença em meses: {diff_meses}")
        else:
            data = primeira_data
            
            # Calcula a diferença em meses
            hoje = datetime.now()
            diff = relativedelta(hoje, data)
            diff_meses = diff.years * 12 + diff.months
            curso = cursos[0]
            print("Curso e data não encontrados.")

        if diff_meses < 18:
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
            
            if "Word" in curso:
                id_curso = 1
                id_curso2 = None
            elif "Excel" in curso:
                id_curso = 2
                id_curso2 = 3
            elif "PowerPoint" in curso:
                id_curso = 4
                id_curso2 = 5
            elif "Outlook" in curso:
                id_curso = 6
                id_curso2 = None
            elif "Photoshop" in curso:
                id_curso = 7
                id_curso2 = None
            elif "Illustrator" in curso:
                id_curso = 8
                id_curso2 = None 
            elif "Rotinas" in curso:
                id_curso = 9
                id_curso2 = 10
            elif "Recursos Humanos" in curso:
                id_curso = 11
                id_curso2 = 12
            elif "Assistente Contábil" in curso:
                id_curso = 13
                id_curso2 = 14
            elif "Marketing Digital" in curso:
                id_curso = 15
                id_curso2 = 16
            elif "Gestão de Pessoas" in curso:
                id_curso = 17
                id_curso2 = 18
            else:
                id_curso = 19
                id_curso2 = None 
            
            if diff_meses == id_curso or (id_curso2 is not None and diff_meses == id_curso2):
                print("ok")

                situacao = "ok"
                envio = f"Olá, tudo bem?\n\nEspero que esta mensagem os encontre bem!\n\nGostaria de informar que estamos reenviando o boletim do(a) aluno(a) devido a uma falha no sistema, que afetou o envio correto de alguns boletins. O problema foi identificado e corrigido ontem, e agora todas as informações estão atualizadas e corretas.\n\nAproveito para compartilhar que {item} está fazendo um excelente progresso nas aulas, mantendo-se em dia com as atividades e demonstrando grande dedicação. O comprometimento dele(a) tem sido exemplar, e isso certamente o(a) coloca em uma ótima posição para alcançar seus objetivos acadêmicos.\n\nCaso tenham qualquer dúvida ou precisem de mais informações, por favor, fiquem à vontade para entrar em contato. Estamos à disposição.\n\nUm grande abraço,\n\nAtenciosamente,\nEquipe Pedagógica."
            else:

                atraso = diff_meses - id_curso

                if diff_meses > id_curso:
                    print("not")

                    situacao = "not"
                    envio = f"Olá, tudo bem?\n\nGostaria de informar que estamos reenviando esta mensagem devido a uma falha no sistema, que impactou o envio correto de algumas notificações. O problema foi corrigido ontem, e agora todas as informações estão atualizadas.\n\nNo momento, {item} está atrasado(a) em {atraso} módulo(s) em relação às aulas e atividades recentes. Entendemos que podem haver desafios, e estamos aqui para ajudar no que for necessário. No entanto, é importante que ele(a) recupere o ritmo para garantir que consiga concluir o curso dentro do tempo previsto.\n\nEstamos à disposição para discutir maneiras de apoiar o(a) aluno(a) nesse processo e ajudá-lo(a) a colocar tudo em dia.\n\nContem conosco para o que precisarem.\n\nUm grande abraço,\n\nAtenciosamente,\nEquipe Pedagógica."
                else:
                    print("adiantado")
                    
                    situacao = "not"
                    envio = f"Olá, tudo bem?\n\nEspero que esta mensagem os encontre bem!\n\nGostaria de informar que estamos reenviando o boletim do(a) aluno(a) devido a uma falha no sistema, que afetou o envio correto de alguns boletins. O problema foi identificado e corrigido ontem, e agora todas as informações estão atualizadas e corretas.\n\nAproveito para compartilhar que {item} está fazendo um excelente progresso nas aulas, mantendo-se em dia com as atividades e demonstrando grande dedicação. O comprometimento dele(a) tem sido exemplar, e isso certamente o(a) coloca em uma ótima posição para alcançar seus objetivos acadêmicos.\n\nCaso tenham qualquer dúvida ou precisem de mais informações, por favor, fiquem à vontade para entrar em contato. Estamos à disposição.\n\nUm grande abraço,\n\nAtenciosamente,\nEquipe Pedagógica."

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
                        
                        attach_button = navegador.find_element(By.XPATH, '//div[@title="Anexar"]')
                        attach_button.click()
                        time.sleep(1)

                        file_input = navegador.find_element(By.XPATH, '//input[@type="file"]')
                        file_input.send_keys(str(pdf_file))
                        time.sleep(10)

                        send_button = navegador.find_element(By.XPATH, '//span[@data-icon="send"]')
                        send_button.click()
                        print("Arquivo enviado com sucesso!")
                        time.sleep(15)
                    except:
                        while len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button')) < 1:
                            time.sleep(1)
                        navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button').send_keys(Keys.ENTER)
                    time.sleep(5)
            cont = cont + 1

# Fechar o navegador
navegador.quit()