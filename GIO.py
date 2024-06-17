from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
from readConfig import readConfig,queryByNameDict

dictConfig = readConfig()

def OpenGIO():
    Path = queryByNameDict('PathDriverEdge',dictConfig)
    #Path= r"realvidaseguros\lib\msedgedriver.exe"
    driver = webdriver.Edge(Path)
    LinkGIO = queryByNameDict('LinkGIO',dictConfig)
    driver.get(LinkGIO)
    time.sleep(10)
    try:
        driver.find_element_by_id("otherTileText")
        driver.maximize_window()
        print("[+]Website disponível!")
        return driver
    except:
        print("[-]Website indisponível!")


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
    SMS=input("Insere o Codigo SMS: ")


def main():
    driver = OpenGIO()
    if (driver):
        loginGIO(driver)
    else:
        print("[-]Erro a Abrir Website....")

main()