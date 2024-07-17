from GIO import *
from readConfig import *
import databaseSQLExpress
from BusinessRuleExceptions import *

COLUMN_NAMES = [
    'EmailRemetente','Anexos','EmailID',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao'
]

def main():
    server = queryByNameDict('SQLExpressServer',dictConfig)
    database = queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    setup_logging(db)
    logger = logging.getLogger(__name__)
    logger.info("A Iniciar Performer.....")
    #driver = OpenGIO(logger)
    #attach a sessão já aberta
    try:
        driver = InitApplications()
        logger.info("Aplicações Iniciadas com Sucesso!")
    except Exception as e:
        logger.error(f"Erro ao iniciar Aplicações {e}")
        raise e
    while True:
        dfQueueItem = pd.DataFrame
        try:
            dfQueueItem = databaseSQLExpress.GetQueueItem(db,COLUMN_NAMES,"QueueItem","Emails_IPA_NLP")
        except Exception as e:
            logger.error(f"Erro ao tentar is buscar QueueItem {e}")
        
        if not dfQueueItem.empty:
            print(dfQueueItem)
            try:
                idAlertas(driver,dfQueueItem,"")
            except BusinessRuleException as e:
                logger.error(f"{e}")
            except Exception as e:
                logger.error(f"SystemError no processamento do registo: {e}")
            
        else:
            logger.info("Sem QueueItems para tratar.")    
            break
    #loginGIO(driver)
    #navegarGIO(driver)
    #pesquisarGIO(driver)
    #ScrapTableGIO(driver)
    #ScrapDetalhesEntidadeGIO(driver)
    #ScrapApoliceGIO(driver)

def InitApplications():
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    Path= r"realvidaseguros\lib\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(Path),options=Browser_options)
    return driver

if __name__ == '__main__':
    main()