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
import pandas as pd 
import re
from BusinessRuleExceptions import *
#cd diretorio chrome
#chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\chrome-profile"

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
 

def pesquisarGIO(driver,search,pesquisa):
    print(driver.title)
    searchButton = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[6]/div/button[1]')
    driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[6]/div/button[2]').click()
    time.sleep(2)
    search.clear()
    search.send_keys(pesquisa)
    searchButton.click()
    time.sleep(4)
    
    searchNumEntries=driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[1]/div/label/select')
    searchNumEntries.click()
    time.sleep(1)
    searchNumPlus=driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[1]/div/label/select/option[4]')
    searchNumPlus.click()
    time.sleep(3)
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[1]/div/h2').click
    time.sleep(10)
 

def ScrapTableGIO(driver):
    pattern = r'\d+'
    NumRegistos = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[3]/div[1]/div').text
    NumRegistos = re.findall(pattern,NumRegistos)
    table = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/table/tbody')
    
    headers=['Nome','TipoEntidade','NIF','Phone','Email','DOB',' ']
    table_data = []
    #Extair Info 
    while True:
        rows = table.find_elements(By.TAG_NAME,'tr')
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, 'td')
            col_data = [col.text for col in cols]
            table_data.append(col_data)
        if not len(table_data) >= max(list(map(int, NumRegistos))):
            driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[3]/div[2]/div/ul/li[3]').click()
            time.sleep(2)
        else:
            break
    #Converter para dataframe
    try:
        df=pd.DataFrame(table_data,columns=headers)
    except:
        df = pd.DataFrame
    print(df)
    return df


def ScrapDetalhesEntidadeGIO(driver):
    Nome = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div/div/div/form/div/fieldset[1]/div[1]/div/input').get_attribute('value')
    NumIF = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[2]/div/div/div/div/form/div/fieldset[2]/div/div[4]/input').get_attribute('value')
    headers=['Nome','NIF']
    df = pd.DataFrame([[Nome,NumIF]],columns=headers)
    print(df)    
    return df

def ScrapApoliceGIO(driver):
    pattern = r'\d+'
    NumRegistos = driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[3]/div[1]/div').text
    NumRegistos = re.findall(pattern,NumRegistos)
    print(max(list(map(int, NumRegistos))))
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[1]/div[1]/div/label/select').click()
    time.sleep(2)
    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[1]/div[1]/div/label/select/option[4]').click()
    time.sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(10)
    #Extrair Info
    table = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/table/tbody')
    rows = table.find_elements(By.TAG_NAME,'tr')
    headers=['Nome','TipoEntidade','NIF','Phone','Email','DOB',' ']
    table_data = []
    #Extair Info 
    while True:
        rows = table.find_elements(By.TAG_NAME,'tr')
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, 'td')
            col_data = [col.text for col in cols]
            table_data.append(col_data)
        if not len(table_data) == max(list(map(int, NumRegistos))):
            driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/div/div/div[3]/div[2]/div/ul/li[5]/a').click()
            time.sleep(2)
        else:
            break
    
    #Converter para dataframe
    #df=pd.DataFrame(table_data,columns=headers)
    #print(df)
    pass

def idAlertas(driver,df:pd.DataFrame,regras):
    
    searchEmail = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[4]/input')
    searchName = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[1]/input')
    searchNIF = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[2]/input')
    searchApolc = driver.find_element(By.XPATH,'/html/body/div[2]/div/form/div/div/div/div[5]/input')
    for i, row in df.iterrows():
        match row['IDIntencao']:
            case 0:
                pesquisarGIO(driver,searchEmail,row['EmailRemetente'])
                df = ScrapTableGIO(driver)
                print(df)
            case 1:
                pesquisarGIO(driver,searchEmail,row['EmailRemetente'])
                df = ScrapTableGIO(driver)
                if not df.empty:
                    #raise bre
                    break
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                pesquisarGIO(driver,searchNIF,row['NIF'])
                df = ScrapTableGIO(driver)
                print(df['TipoEntidade'].str.contains('T'))


                print(df)
    print("?")            
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    
 