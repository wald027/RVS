from Automation.GIO import *
from customScripts.readConfig import *
import customScripts.databaseSQLExpress as databaseSQLExpress
from Automation.BusinessRuleExceptions import BusinessRuleException
from customScripts.customLogging import setup_logging
from datetime import datetime

COLUMN_NAMES = [
    'EmailRemetente','DataEmail','Anexos','EmailID','Subject','Body',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao','Score','Mensagem','DetalheMensagem','Estado','Status'
]

def main():
    dictConfig = readConfig(r'Config.xlsx')
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
    mailboxname = queryByNameDict('MailboxName',dictConfig)
    pastaEmailsTratamento = queryByNameDict('EmailsToMove',dictConfig)
    pastaEmailsTratamentoManual = queryByNameDict('EmailTratamentoManualMove',dictConfig) 
    pastaEmailsSucesso = queryByNameDict('EmailSucessoMove',dictConfig)
    #driver = OpenGIO(logger)
    #attach a sessão já aberta
    try:
        #for app in queryByNameDict('AplicacoesPerf',dictConfig).split(','):
            #KillAllApplication(app+'.exe',logger)
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
                #Atividade NLP (Verificar Dados providenciados pelo mesmo)
                try:
                    logger.info('A Verificar Dados Vindos do NLP....')
                    #IDbd = dfQueueItem.loc[0,'IDIntencao']
                    if dfQueueItem.loc[0,'Status'] == 'NLP FAILED':
                        raise Exception(f'Falha do NLP ao Processar Registo')
                    if dfQueueItem.loc[0,'Score'] < queryByNameDict('TrustScore',dictConfig):
                        raise BusinessRuleException(f'Score de {dfQueueItem.loc[0,"Score"]*100}% de Confiança Abaixo do Permitido')
                    elif queryByNameDict('TratarEmailsHistorico',dictConfig) == 'Não' and dfQueueItem.loc[0,'HistoricoEmails'] == True:
                        raise BusinessRuleException(f'O Processamento de Emails com Histórico está Desabilitado')
                    logger.info('Dados Verificados com Sucesso!')
                except BusinessRuleException as e:
                    logger.error(f"{e}")
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip()
                    dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade do NLP'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA')
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
                    raise BusinessRuleException(e)
                except Exception as e:
                    logger.error(f'SystemError no processamento do registo: {e}')
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = ''
                    dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                    dfQueueItem.loc[0,"Mensagem"] = 'Indisponilibidade do NLP'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA')
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed","Erro de Sistema",e)
                    raise Exception(e) 
                #Atividade GIO
                try:
                    logger.info('A Processar a Atividade de Obtenção de Dados do GIO.....')
                    IDbd = dfQueueItem.loc[0,'IDIntencao']
                    idAlertas(driver,dfQueueItem,dictConfig,logger)
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd #solucao temporaria
                    logger.info('Atividade Processada com Sucesso!')
                except BusinessRuleException as e:
                    logger.error(f"{e}")
                    dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip()
                    dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade de obter dados no GIO'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA')
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
                    raise BusinessRuleException(e)
                except Exception as e:
                    logger.error(f'SystemError no processamento do registo: {e}')
                    dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = ''
                    dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                    dfQueueItem.loc[0,"Mensagem"] = 'Indisponilibidade em obter dados no GIO'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA')
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed","Erro de Sistema",e)
                    raise Exception(e) 
                #Atividade EMail
                try:
                    logger.info('A Processar a Atividade de Obtenção Envio de Email.....')
                    EnviarEmail(dfQueueItem,dictConfig,logger)
                    dfQueueItem['DetalheMensagem'] = 'Tratamento realizado com Sucesso'
                    dfQueueItem['Mensagem'] = 'Sucesso no tratamento'
                    dfQueueItem['Estado'] = 'Processado'
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Sucesso",'Processado','Tratamento realizado com Sucesso')
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via Email')
                    mail = SearchMailInbox(logger,pastaEmailsTratamento,mailboxname,dfQueueItem.loc[0,"EmailID"])
                    MoveEmailToFolder(logger,pastaEmailsSucesso,mailboxname,mail)
                    logger.info('Registo Processado com Sucesso!')
                except BusinessRuleException as e:
                    logger.error(f"{e}")
                    dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip()
                    dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade em enviar Email pelo Outlook'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA')
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
                    raise BusinessRuleException(e)
                except Exception as e:
                    logger.error(f'SystemError no processamento do registo: {e}')
                    dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = ''
                    dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                    dfQueueItem.loc[0,"Mensagem"] = 'Indisponibilidade em enviar Email pelo Outlook'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA')
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed","Erro de Sistema",e)
                    raise Exception(e)  
            except BusinessRuleException as e:
                mail = SearchMailInbox(logger,pastaEmailsTratamento,mailboxname,dfQueueItem.loc[0,"EmailID"])
                MoveEmailToFolder(logger,pastaEmailsTratamentoManual,mailboxname,mail)
                logger.error(f'{(str(e).split(":")[1]+":"+str(e).split(":")[2]).strip()}') #Fix para não aparecer "Definição de Negocio:" duas vezes seguidas
            except Exception as e:
                mail = SearchMailInbox(logger,pastaEmailsTratamento,mailboxname,dfQueueItem.loc[0,"EmailID"])
                MoveEmailToFolder(logger,pastaEmailsTratamentoManual,mailboxname,mail)
                logger.error(f'Erro de Sistema no Processamento do Registo - {e}')
        else:
            logger.info("Sem QueueItems para tratar.")    
            break
    #print(dfReportOutput)
    file_path = f'Output\Output_Pedidos_de_Clientes_{datetime.now().strftime("%d%m%Y_%H%M%S")}.xlsx'
    dfReportOutput.to_excel(file_path, index=False, header=True)
    databaseSQLExpress.SetReportOutput(db,'Report_Output',dfReportOutput)
    
    #loginGIO(driver)
    #navegarGIO(driver)
    #pesquisarGIO(driver)
    #ScrapTableGIO(driver)
    #ScrapDetalhesEntidadeGIO(driver)
    #ScrapApoliceGIO(driver)

def InitApplications():
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    Application().start(outlook_path)
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    Path= r"C:\Users\brunofilipe.lobo\OneDrive - CGI\Code\realvidaseguros\IntelligentProcessAutomationNLP\Automation\lib\chromedriver.exe"
    driver = webdriver.Chrome(service=Service(Path),options=Browser_options)
    return driver

def KillAllApplication(processname,logger:logging.Logger):
    try:
        logger.info(f'A Forçar o Fecho da Aplicação {processname}...')
        os.system(f'taskkill /f /im {processname}')
        logger.info(f'{processname} fechado com Sucesso!')
    except Exception as e:
        logger.warning(e)

def prepararOutput(df:pd.DataFrame,viatratamento):
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
            'ViaTratamento': f"{viatratamento}",
            'MensagemOutput': f"{row['Mensagem']}",
            'Estado': f"{row['Estado']}"
        }
    return new_row

if __name__ == '__main__':
    main()