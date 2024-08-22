from customScripts import customLogging 
from customScripts import databaseSQLExpress
from Automation import MailboxRVS
import logging
from customScripts import readConfig
import time
from ModelNLP.NLP import EmailClassifier
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String

NUM_LABELS = 11
BASE_DIR = 'realvidaseguros/'
TOKENIZER_PATH = BASE_DIR + "tokenizer"
MODEL_PATH = BASE_DIR + "model"
DATABASE = "RealVidaSeguros"
STATUS_TABLE = "QueueItem"
EMAIL_TABLE = "Emails_IPA_NLP"

COLUMN_NAMES = [
    'EmailRemetente','DataEmail', 'EmailID','Subject', 'Body', 'Anexos',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao', 'Score', 'IDTermosExpressoes',
    'DetalheMensagem', 'Mensagem', 'Estado'
]


def main():
    #iniciar database, custom logger
    dictConfig = readConfig.readConfig()
    server = readConfig.queryByNameDict('SQLExpressServer',dictConfig)
    database = readConfig.queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    driver = readConfig.queryByNameDict('SQLDriver',dictConfig)
    databaseLogsTable=readConfig.queryByNameDict('LogsTableName',dictConfig)


    customLogging.setup_logging(db,databaseLogsTable)
    try:
        logger = logging.getLogger(__name__)
        logger.info("A Iniciar o Dispatcher do Processo RVS IPA NLP....")
        time.sleep(1)
        logger.info("Config Lida Com Sucesso!")
        mailcount = MailboxRVS.GetEmailsInbox(logger,db,dictConfig)
        if mailcount > 0:
            logger.info("Emails Extraídos com Sucesso!")
            ENGINE = create_engine(f"mssql+pyodbc://@{server}/{DATABASE}?driver={driver}&Trusted_Connection=yes")
            CONN = ENGINE.connect()
            EmailClassifier(BASE_DIR,NUM_LABELS,STATUS_TABLE,EMAIL_TABLE,COLUMN_NAMES,ENGINE,logger,db).run()
        else:
            logger.warning('Sem Emails para Tratamento!')  
        time.sleep(5)
        logger.info("Dispatcher Terminado")
    except Exception as e:
        logger.error(f"Erro Dispatcher {e}")
    db.close()

if __name__ == '__main__':
    main()

    