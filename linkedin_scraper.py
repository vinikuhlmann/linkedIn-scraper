import pandas as pd
import logging
import time
from collections import namedtuple
from getpass import getpass
from os import getcwd
from relatorio_individual import gerar_relatorio_individual
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from sys import exit as sys_exit

from relatorio_individual import gerar_relatorio_individual 

Index = namedtuple('Index', 'empresa nome_perfil_busca cargo_perfil_busca linkedin_perfil_busca is_conexao_direta nome_conexao')
Data = namedtuple('Data', 'cargo_conexao linkedin_conexao')
tabela_dados = {'index': [],
                'columns': ['Cargo da Conexão', 'LinkedIn da Conexão'],
                'data': [],
                'index_names': ['Empresa', 'Nome', 'Cargo', 'LinkedIn', 'Conexão direta?', 'Nome da Conexão'],
                'column_names': ['']}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    datefmt="%H:%M:%S"
)

def init_driver() -> webdriver:

    options = Options()
    options.add_argument("start-maximized")
    options.add_argument("enable-automation")
    options.add_argument("--headless") # Comente para visualizar o navegador
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-browser-side-navigation")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-crash-reporter")
    options.add_argument("--disable-in-process-stack-traces")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")
    options.add_argument("--output=/dev/null")
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(5)
    driver.maximize_window()

    return driver

# * Faz o login do usuário
def linkedin_login(driver: webdriver, usuario, senha):

    driver.get('https://www.linkedIn.com')

    login_linkedIn = driver.find_element(By.ID, 'session_key')
    login_linkedIn.send_keys(usuario)

    password_linkedIn = driver.find_element(By.ID, 'session_password')
    password_linkedIn.send_keys(senha)
    password_linkedIn.send_keys(Keys.ENTER)

    time.sleep(2)

# * Coleta e retorna o nome de uma empresa, usando o link do perfil dela
def extrair_nome_empresa(driver: webdriver, linkedin_empresa):

    driver.get(linkedin_empresa)
    nome_empresa = driver.find_element(By.XPATH, "//div[@class='block mt2']").find_element(By.TAG_NAME, 'h1').text

    return nome_empresa

# * Verificar conexão direta ao perfil de um diretor
def extrair_is_conexao_direta(driver: webdriver, linkedin_perfil_de_busca):
    
    driver.get(linkedin_perfil_de_busca)
    is_conexao_direta = True

    try:
        driver.find_element(By.XPATH, """//div[@class='message-anywhere-button pvs-profile-actions__action artdeco-button artdeco-button--secondary 
        artdeco-button--muted']""")
    except NoSuchElementException:
        is_conexao_direta = False
    

    return is_conexao_direta

# * Coleta conexões de um perfil de diretor
def extrair_dados_perfil_de_busca(driver: webdriver, linkedin_perfil_de_busca):

    driver.get(linkedin_perfil_de_busca)
    
    try:
        element_link_conexoes = driver.find_element(By.XPATH, "//a[@class='app-aware-link display-flex pv1 align-items-center link-without-hover-visited']")
        link_conexoes = element_link_conexoes.get_attribute('href')
    except NoSuchElementException:
        return ([], [])

    lista_nomes = []
    lista_dados = []
    
    # Alguns perfis com muitos seguidores fornecem uma página de "insights" ao invés de uma página de conexões
    # Estes perfis não são usáveis pelo programa
    if link_conexoes.endswith("insights"):
        logging.info('A página deste diretor não fornece as conexões em comum')
        return ([], [])

    pagina = 1

    while True:

        # Insere atributo da página entre o último e o penúltimo atributo do link e busca
        link_conexoes = '&'.join(link_conexoes.split('&')[:-1])+f"&page={pagina}&"+'&'.join(link_conexoes.split('&')[-1:])
        
        driver.get(link_conexoes)

        # Avança a página de busca
        pagina += 1

        time.sleep(2)

        # Encontra os elementos da conexão
        lista_element_conexao = driver.find_elements(By.XPATH, "//li[@class='reusable-search__result-container ']")

        if len(lista_element_conexao) == 0:
            break

        for element_conexao in lista_element_conexao:
            
            # Extrai todo o texto do elemento da conexão o divide por linhas
            conexao = element_conexao.text.split('\n')

            if len(conexao) == 0:
                break
            
            # A posição do nome e do cargo da conexão pode variar, por existir o elemento "Status" ou sem motivo aparente
            if (conexao[0].startswith("Status") or conexao[0] == conexao[1]):
                nome_conexao = conexao[1]
                cargo_conexao = conexao[5]
            else:
                nome_conexao = conexao[0]
                cargo_conexao = conexao[4]

            # Extrai o link do elemento
            linkedin_conexao = element_conexao.find_element(By.CLASS_NAME, "app-aware-link").get_attribute('href').split('?')[0]

            lista_nomes.append(nome_conexao)
            lista_dados.append(Data(cargo_conexao, linkedin_conexao))

    return (lista_nomes, lista_dados)

# * Insere uma linha em uma tabela de dicionário
def inserir_data(index: Index, data: Data):
    tabela_dados['index'].append(tuple(index))
    tabela_dados['data'].append(list(data))

# * Função principal
def extrair_conexoes(nome, usuario, senha, perfis_busca_xlsx) -> tuple[str, pd.DataFrame]:

    # Abre o arquivo de perfis de busca e armazena em um DataFrame
    try:
        perfis_busca_cols = ['Nome', 'Cargo', 'LinkedIn', 'LinkedIn Empresa']
        perfis_busca_df = pd.read_excel(perfis_busca_xlsx, engine='openpyxl', usecols=perfis_busca_cols)
    except FileNotFoundError:
        logging.error(f"{perfis_busca_xlsx} não foi encontrado")
        sys_exit()
    except KeyError:
        logging.error(f"Nomes de colunas de {perfis_busca_xlsx} não correspondem com o esperado \n Colunas esperadas: {str(perfis_busca_cols)}")
        sys_exit()

    try:
        driver = init_driver()
    except SessionNotCreatedException as e:
        logging.error("O chromedriver.exe está desatualizado. Baixe a versão correta em https://chromedriver.chromium.org/downloads")
        sys_exit()

    logging.info(f"Logando no perfil de {nome}") 
    linkedin_login(driver, usuario, senha)
    if driver.current_url == 'https://www.linkedin.com/uas/login-submit':
        logging.error(f"A senha do perfil de {nome} está incorreta")
        driver.quit()

    for idx, row in perfis_busca_df.iterrows():
        
        try:
            empresa = extrair_nome_empresa(driver, row['LinkedIn Empresa'])
        except:
            logging.info("Não foi possível extrair o nome da empresa")
            empresa = None
        
        # Armazena as informações atuais nas namedtuples
        index = Index(empresa=empresa, nome_perfil_busca=row['Nome'], linkedin_perfil_busca=row['LinkedIn'], cargo_perfil_busca=row['Cargo'], is_conexao_direta='X', 
        nome_conexao=None)
        data = Data(cargo_conexao=pd.NA, linkedin_conexao=pd.NA)

        logging.info(f"Extraindo conexões entre {index.nome_perfil_busca} e {nome}")

        try:
            is_conexao_direta = extrair_is_conexao_direta(driver, index.linkedin_perfil_busca)
        except:
            logging.info(f"Não foi possível acessar o perfil de {index.nome_perfil_busca}")
            
        if is_conexao_direta:
            logging.info(f"{index.nome_perfil_busca} tem conexão direta com {nome}")
            inserir_data(index, data)
        else:
            logging.info(f"{index.nome_perfil_busca} não tem conexão direta com {nome}")

        scrape_return = extrair_dados_perfil_de_busca(driver, index.linkedin_perfil_busca)

        if len(scrape_return[0]) == 0: # Checa se algum resultado foi retornado
            logging.info(f"{index.nome_perfil_busca} não tem conexões em comum com {nome}")
        else:
            logging.info(f"{index.nome_perfil_busca} tem {len(scrape_return[0])} conexões com {nome}")
            for nome_conexao, data in zip(*scrape_return):
                index = Index(empresa=empresa, nome_perfil_busca=row['Nome'], linkedin_perfil_busca=row['LinkedIn'], cargo_perfil_busca=row['Cargo'], 
                is_conexao_direta=pd.NA, nome_conexao=nome_conexao)
                inserir_data(index, data)

    driver.quit()

    conexoes_df = pd.DataFrame.from_dict(tabela_dados, orient='tight')

    logging.info("Organizando dados")
    conexoes_df.reset_index(inplace=True)
    conexoes_df.sort_values(by=['Empresa', 'Nome', 'Nome da Conexão'], key=lambda col: col.str.lower(), inplace=True)
    conexoes_df.reset_index(drop=True, inplace=True)

    return nome, conexoes_df

if __name__ == '__main__':
    nome, usuario, senha = None, None, None
    while nome == None:
        nome = input("Digite seu nome e aperte ENTER\n")
    while usuario == None:
        usuario = input("Digite seu login do LinkedIn e aperte ENTER\n")
    while senha == None:
        senha = getpass("Digite sua senha do LinkedIn e aperte ENTER\n")
    perfis_busca_xlsx = getcwd()+"/Perfis de busca.xlsx"
    nome, conexoes_df = extrair_conexoes(nome, usuario, senha, perfis_busca_xlsx)
    gerar_relatorio_individual(nome, conexoes_df)