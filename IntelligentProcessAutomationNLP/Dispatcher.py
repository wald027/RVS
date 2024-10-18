from customScripts import customLogging 
from customScripts import databaseSQLExpress
from Automation import MailboxRVS
import logging
from customScripts import readConfig
import time
from ModelNLP.NLP import EmailClassifier
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String
import os
from pywinauto import Application
import pandas as pd
#NUM_LABELS = 11
#BASE_DIR = 'realvidaseguros/'
#TOKENIZER_PATH = BASE_DIR + "tokenizer"
#MODEL_PATH = BASE_DIR + "model"
#DATABASE = "RealVidaSeguros"
#STATUS_TABLE = "Queue_Items"
#EMAIL_TABLE = "Pedidos"

PATH_CONFIG = r'Config.xlsx'

COLUMN_NAMES = [
    'EmailRemetente','DataEmail', 'EmailID','Subject', 'Body', 'Anexos',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao', 'Score', 'IDTermosExpressoes',
    'DetalheMensagem', 'Mensagem', 'Estado','EmailCurto','NomeIntencao'
]

def main():
    #iniciar database, custom logger
    dictConfig = readConfig.readConfig(PATH_CONFIG)
    Debug = readConfig.queryByNameDict('Teste_Dispatcher_NLP',dictConfig)
    server = readConfig.queryByNameDict('SQLExpressServer',dictConfig)
    database = readConfig.queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    driver = readConfig.queryByNameDict('SQLDriver',dictConfig)
    if Debug:#Exec de Teste 
        databaseLogsTable=readConfig.queryByNameDict('LogsTableName_Teste',dictConfig)
        STATUS_TABLE = readConfig.queryByNameDict('QueueTableName_Teste',dictConfig)
        EMAIL_TABLE = readConfig.queryByNameDict('TableName_Teste',dictConfig)
    else:#Exec Normal
        databaseLogsTable=readConfig.queryByNameDict('LogsTableName',dictConfig)
        STATUS_TABLE = readConfig.queryByNameDict('QueueTableName',dictConfig)
        EMAIL_TABLE = readConfig.queryByNameDict('TableName',dictConfig)

    nomeprocesso = readConfig.queryByNameDict('NomeProcesso',dictConfig)
    BASE_DIR = readConfig.queryByNameDict('Base_Dir',dictConfig)
    NUM_LABELS = readConfig.queryByNameDict('NumLabelsNLP',dictConfig)
    #TOKENIZER_PATH = readConfig.queryByNameDict('TokenizerPath',dictConfig)
    intencoes_filepath = readConfig.queryByNameDict('PathConfigIntencoes',dictConfig)
    dfDict = pd.read_excel(intencoes_filepath,sheet_name='LabelMap')
    label_map = dict(zip(dfDict['Key'].astype(str), dfDict['Value'].astype(str)))

    logger = customLogging.setup_logging(db,databaseLogsTable,nomeprocesso)
    try:
        #logger = logging.getLogger(__name__)
        if Debug:
            logger.warning('O DISPATCHER ESTÁ A EXECUTAR EM MODO TESTE')
            logger.warning('OS DADOS VÃO SER ADICIONADOS ÀS TABELAS DE TESTE')
        logger.info(f"A Iniciar o Dispatcher do Processo {nomeprocesso}....")
        time.sleep(1)
        logger.info("Config Lida Com Sucesso!")
        #for app in readConfig.queryByNameDict('AplicacoesDisp',dictConfig).split(','):
        #    KillAllApplication(app+'.exe',logger)
        #InitApplications(readConfig.queryByNameDict('outlookPath',dictConfig),logger)
        mailcount = MailboxRVS.NewGetEmailsInbox(logger,db,dictConfig,nomeprocesso,EMAIL_TABLE,STATUS_TABLE)
        if mailcount > 0:
            logger.info("Emails Extraídos com Sucesso!")
            ENGINE = create_engine(f"mssql+pyodbc://@{server}/{database}?driver={driver}&Trusted_Connection=yes")
            CONN = ENGINE.connect()
            EmailClassifier(BASE_DIR,NUM_LABELS,STATUS_TABLE,EMAIL_TABLE,COLUMN_NAMES,ENGINE,label_map,logger,db,Debug).run()
        else:
            logger.warning('Sem Emails para Tratamento!')  
        time.sleep(5)
        logger.info("Dispatcher Terminado")
    except Exception as e:
        logger.error(f"Erro no Dispatcher {e}")
    db.close()


    
def InitApplications(outlook_path,logger:logging.Logger):
    #outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    logger.info(f'A tentar abrir a aplicação {outlook_path}')
    Application().start(outlook_path)
    time.sleep(10)
    logger.info('Aplicação iniciada com sucesso!')

def KillAllApplication(processname,logger:logging.Logger):
    try:
        logger.info(f'A Forçar o Fecho da Aplicação {processname}...')
        os.system(f'taskkill /f /im {processname}')
        logger.info(f'{processname} fechado com Sucesso!')
    except Exception as e:
        logger.warning(e)


if __name__ == '__main__':
    main()