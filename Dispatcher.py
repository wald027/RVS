from customLogging import setup_logging
import databaseSQLExpress
import MailboxRVS
import logging
import readConfig
import time

if __name__ == '__main__':
    #iniciar database, custom logger
    dictConfig = readConfig.readConfig()
    server = readConfig.queryByNameDict('SQLExpressServer',dictConfig)
    database = readConfig.queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    setup_logging(db)
    logger = logging.getLogger(__name__)
    logger.info("A Iniciar o Dispatcher do Processo RVS IPA NLP....")
    time.sleep(1)
    logger.info("Config Lida Com Sucesso!")
    MailboxRVS.GetEmailsInbox(logger)
    logger.info("Emails Extraídos com Sucesso!")
    #Invoke NLP
    time.sleep(5)
    logger.info("Dispatcher Terminado")
    db.close()