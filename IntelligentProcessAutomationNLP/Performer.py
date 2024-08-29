from Automation.GIO import *
from customScripts.readConfig import *
import customScripts.databaseSQLExpress as databaseSQLExpress
from Automation.BusinessRuleExceptions import BusinessRuleException
from customScripts.customLogging import setup_logging
from datetime import datetime

COLUMN_NAMES = [
    'EmailRemetente','DataEmail','Anexos','EmailID','Subject','Body',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao','Score','Mensagem','DetalheMensagem','Estado'
]

def main():
    dictConfig = readConfig()
    server = queryByNameDict('SQLExpressServer',dictConfig)
    database = queryByNameDict('Database',dictConfig)
    db = databaseSQLExpress.ConnectToBD(server,database)
    nomeprocesso = queryByNameDict('NomeProcesso',dictConfig)
    databaseLogsTable=queryByNameDict('LogsTableName',dictConfig)
    tabelaPedidos =queryByNameDict('TableName',dictConfig)
    queueItem =queryByNameDict('QueueTableName',dictConfig)
    setup_logging(db,databaseLogsTable,nomeprocesso)
    logger = logging.getLogger(__name__)
    logger.info("A Iniciar Performer.....")
    #driver = OpenGIO(logger)
    #attach a sessão já aberta
    try:
        driver = InitApplications()
        logger.info("Aplicações Iniciadas com Sucesso!")
        dfReportOutput=pd.DataFrame(columns=['Data Processamento','Reference','EmailRemetente','NIF','Apolice','Nome','DataEmail','TemaIdentificado','ViaTratamento','MensagemOutput','Estado'])

    except Exception as e:
        logger.error(f"Erro ao iniciar Aplicações {e}")
        raise e
    while True:
        dfQueueItem:pd.DataFrame
        try:
            dfQueueItem = databaseSQLExpress.GetQueueItem(db,COLUMN_NAMES,queueItem,tabelaPedidos)
            #print(dfQueueItem)
        except Exception as e:
            logger.error(f"Erro ao tentar is buscar QueueItem {e}")
        
        if not dfQueueItem.empty:
            logger.info(f'A tratar o registo com o EmailID/Reference {dfQueueItem["EmailID"].to_string().replace("0","").replace(" ","")} e com Intenção Identificada pelo NLP de {dfQueueItem["IDIntencao"].to_string().replace("0","").replace(" ","")}')
            try:
                IDbd = dfQueueItem.loc[0,'IDIntencao']
                if dfQueueItem.loc[0,'Score'] < queryByNameDict('TrustScore',dictConfig):
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade do NLP'
                    raise BusinessRuleException(f'Score de {dfQueueItem.loc[0,"Score"]*100}% de Confiança Abaixo do Permitido')
                idAlertas(driver,dfQueueItem,dictConfig,logger)
                dfQueueItem.loc[0,'IDIntencao'] = IDbd #solucao temporaria
                dfQueueItem['DetalheMensagem'] = 'Tratamento realizado com Sucesso'
                dfQueueItem['Mensagem'] = 'Sucesso no tratamento'
                dfQueueItem['Estado'] = 'Processado'
                databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Sucesso",'Processado','Tratamento realizado com Sucesso')
            except BusinessRuleException as e:
                logger.error(f"{e}")
                dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip()
                dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
            except Exception as e:
                logger.error(f'SystemError no processamento do registo: {e}')
                dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                dfQueueItem.loc[0,"DetalheMensagem"] = ''
                dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed","Erro de Sistema",e)
            #adicionar à dataframe de report?
            dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem)
            logger.info('Tratamento do registo finalizado e adicionado à tabela de Report Output')
        else:
            logger.info("Sem QueueItems para tratar.")    
            break
    #print(dfReportOutput)
    file_path = f'Output_Pedidos_de_Clientes_{datetime.now().strftime("%d%m%Y_%H%M%S")}.xlsx'
    dfReportOutput.to_excel(file_path, index=False, header=True)
    
    #loginGIO(driver)
    #navegarGIO(driver)
    #pesquisarGIO(driver)
    #ScrapTableGIO(driver)
    #ScrapDetalhesEntidadeGIO(driver)
    #ScrapApoliceGIO(driver)

def InitApplications():
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    Path= r"C:\Users\brunofilipe.lobo\OneDrive - CGI\Code\realvidaseguros\IntelligentProcessAutomationNLP\Automation\lib\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(Path),options=Browser_options)
    return driver

def prepararOutput(df:pd.DataFrame):
    for index,row in df.iterrows():
        new_row = {
            'Data Processamento': f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}',
            'Reference': f"{row['EmailID']}",
            'EmailRemetente': f"{row['EmailRemetente']}",
            'NIF': f"{row['NIF']}",
            'Apolice': f"{row['Apolice']}",
            'Nome': f"{row['Nome']}",
            'DataEmail': f"{row['DataEmail']}",
            'TemaIdentificado': f"{row['IDIntencao']}",
            'ViaTratamento': f"Teste",
            'MensagemOutput': f"{row['Mensagem']}",
            'Estado': f"{row['Estado']}"
        }
    return new_row

if __name__ == '__main__':
    main()