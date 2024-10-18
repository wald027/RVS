import csv
import os.path
import pandas as pd 


def readConfig(path):
    configPath = path
    if os.path.isfile(configPath):
        dictConfig:pd.DataFrame
        dictConfig =pd.read_excel(configPath,keep_default_na=False,sheet_name='Sheet1') 
        return dictConfig
    else:
        print("Ficheiro Config não encontrado!")

def queryByNameDict(name,dictConfig:pd.DataFrame):
    for index,row in dictConfig.iterrows():
        if row['name'] == name:
            return row['value']
    return None


""""
def readConfig():
    # Path da Config
    csv_file_path = r'realvidaseguros\Config.csv'
    if os.path.isfile(csv_file_path):
        print("Ficheiro Config Existe")
        #Init Dictionary vazio
        dictConfig = []
        inf = 1
        # Abrir CSV com csv.DictReader
        with open(csv_file_path, mode='r', newline='') as csv_file:
            #next(csv_file,inf)  #ignorar primeira linha de config
            csv_reader = csv.DictReader(csv_file,delimiter=";")
            
            # Ir a Cada Row e Guardar no DictConfig
            for row in csv_reader:
                dictConfig.append(row)
        #for row in dictConfig: #debug
        #    print(row)
        return dictConfig
    else:
        print("Ficheiro de Configuração não encontrado!")
"""

""""
def queryByNameDict(name,dictConfig):
    intCounter = 0
    for row in dictConfig:
        intCounter=intCounter+1
        if row['name'] == name:
            return row['value']
        else:
            if intCounter == len(dictConfig):
                return None
"""            
#Apagar
def readRegrasApolices():
    file_path = r'C:\Users\brunofilipe.lobo\Documents\Code\realvidaseguros\intencoes.xlsx'
    if file_path:
        dfRegras = pd.read_excel(file_path,keep_default_na=False)
        return dfRegras
    else:
        print("Ficheiro de Regras de Apólice, não encontrado!")
        return None