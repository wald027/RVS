from GIO import *
from readConfig import *
import databaseSQLExpress
from BusinessRuleExceptions import *

COLUMN_NAMES = [
    'EmailRemetente','Anexos','EmailID','Subject','Body',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao','Score'
]

def main():
    dictConfig = readConfig()
    server = queryByNameDict('SQLExpressServer',dictConfig)
    database = queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    setup_logging(db,'LOGS')
    logger = logging.getLogger(__name__)
    logger.info("A Iniciar Performer.....")
    #driver = OpenGIO(logger)
    #attach a sessão já aberta
    dfResultadosOutput=pd.DataFrame(columns=['DataProcessamento','EmailID','EmailRemetente','NIF','Apolice','Nome','DataEmail','TemaIdentificado','ViaTratamento','MensagemOutput','Estado'])
    try:
        driver = InitApplications()
        logger.info("Aplicações Iniciadas com Sucesso!")
    except Exception as e:
        logger.error(f"Erro ao iniciar Aplicações {e}")
        raise e
    while True:
        dfQueueItem:pd.DataFrame
        try:
            dfQueueItem = databaseSQLExpress.GetQueueItem(db,COLUMN_NAMES,"QueueItem","Emails_IPA_NLP")
            print(dfQueueItem)
        except Exception as e:
            logger.error(f"Erro ao tentar is buscar QueueItem {e}")
        
        if not dfQueueItem.empty:
            logger.info(f'A tratar o registo com o EmailID/Reference {dfQueueItem["EmailID"].to_string().replace("0","").replace(" ","")} e com Intenção Identificada pelo NLP de {dfQueueItem["IDIntencao"].to_string().replace("0","").replace(" ","")}')
            try:
                if dfQueueItem.loc[0,'Score'] < float(queryByNameDict('TrustScore',dictConfig).replace(',','.')):
                    raise BusinessRuleException(f'Score de {dfQueueItem.loc[0,"Score"]*100}% de Confiança Abaixo do Permitido')
                IDbd = dfQueueItem.loc[0,'IDIntencao']
                idAlertas(driver,dfQueueItem,dictConfig,logger)
                dfQueueItem.loc[0,'IDIntencao'] = IDbd #solucao temporaria
                dfQueueItem['DetalheMensagem'] = 'Tratamento realizado com Sucesso'
                dfQueueItem['Mensagem'] = 'Sucesso no tratamento'
                dfQueueItem['Estado'] = 'Sucesso'
            except BusinessRuleException as e:
                logger.error(f"{e}")
                dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,"",'QueueItem',"","Failed",'Definição do Negócio',e)
            except Exception as e:
                logger.error(f'SystemError no processamento do registo: {e}')
                dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,"",'QueueItem',"","Failed","Erro de Sistema",e)
            #adicionar à dataframe de report?
            dfResultadosOutput
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