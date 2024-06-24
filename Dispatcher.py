from customLogging import setup_logging
import databaseSQLExpress
import MailboxRVS
import logging
import readConfig
import time

def main():
    #iniciar database, custom logger
    dictConfig = readConfig.readConfig()
    server = readConfig.queryByNameDict('SQLExpressServer',dictConfig)
    database = readConfig.queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    setup_logging(db)
    try:
        logger = logging.getLogger(__name__)
        logger.info("A Iniciar o Dispatcher do Processo RVS IPA NLP....")
        time.sleep(1)
        logger.info("Config Lida Com Sucesso!")
        MailboxRVS.GetEmailsInbox(logger)
        logger.info("Emails Extraídos com Sucesso!")
        #Invoke NLP
        time.sleep(5)
        logger.info("Dispatcher Terminado")
    except Exception as e:
        logger.error(f"Erro Dispatcher {e}")
        db.close()

if __name__ == '__main__':
    main()