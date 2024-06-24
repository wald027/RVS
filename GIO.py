from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from readConfig import readConfig,queryByNameDict
from customLogging import setup_logging
import databaseSQLExpress
import logging

dictConfig = readConfig()

def OpenGIO(logger):
    #Path = queryByNameDict('PathDriverEdge',dictConfig)
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    Path= r"realvidaseguros\lib\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(Path),options=Browser_options)
    LinkGIO = queryByNameDict('LinkGIO',dictConfig)
    driver.get(LinkGIO)
    time.sleep(10)
    try:
        driver.find_element_by_id("otherTileText")
        driver.maximize_window()
        logger.info("Website disponível!")
        return driver
    except Exception as e:
        logger.error(f"Website indisponível {e}")


def loginGIO(driver):
    search = driver.find_element_by_id("otherTileText")
    search.click()
    time.sleep(5)
    search = driver.find_element(By.NAME,'loginfmt')
    search.send_keys(queryByNameDict('EmailGIO',dictConfig))
    search.send_keys(Keys.RETURN)
    time.sleep(5)
    search = driver.find_element(By.NAME,'passwd')
    search.send_keys(queryByNameDict('PasswordGIO',dictConfig))
    time.sleep(5)
    search.send_keys(Keys.RETURN)
    time.sleep(15)
    SMS=input("Carregar Enter Após Login Efetuado: ")

def navegarGIO(driver):
    print(driver.title)
    search = driver.find_element(By.XPATH,'/html/body/div[2]/nav/div/ul/li[3]/a')
    search.click()

def pesquisarGIO(driver):
    print(driver.title)
    
    #serch = 

if __name__ == '__main__':
    server = queryByNameDict('SQLExpressServer',dictConfig)
    database = queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    setup_logging(db)
    logger = logging.getLogger(__name__)
    logger.info("A Iniciar Performer")
    #driver = OpenGIO(logger)
    #attach a sessão já aberta
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    Path= r"realvidaseguros\lib\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(Path),options=Browser_options)
    #loginGIO(driver)
    if not (driver):
        logger.error("Erro a Abrir Website")
    #    raise SystemError('Erro ao Abrir Website')
    navegarGIO(driver)

