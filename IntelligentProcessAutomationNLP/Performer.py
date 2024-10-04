from Automation.GIO import *
from customScripts.readConfig import *
import customScripts.databaseSQLExpress as databaseSQLExpress
from Automation.BusinessRuleExceptions import BusinessRuleException
from customScripts.customLogging import setup_logging
from datetime import datetime
from Automation.MailboxRVS import SendEmail as SendOutlookMail
import subprocess

COLUMN_NAMES = [
    'EmailRemetente','DataEmail','Anexos','EmailID','Subject','Body',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao','Score','Mensagem','DetalheMensagem','Estado','Status'
]

def main():
    try:
        dictConfig = readConfig(r'Config.xlsx')
    except Exception as e:
        print(f'Erro ao tentar ler a Config.xlsx, {e}')
        #logger = logging.getLogger(__name__)
        #logger.info(f'Erro ao tentar ler a Config.xlsx, {e}')
        raise e
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
    try:
        for app in queryByNameDict('AplicacoesPerf',dictConfig).split(','):
            KillAllApplication(app+'.exe',logger)
        driver = InitApplications(dictConfig)
        logger.info("Aplicações Iniciadas com Sucesso!")
        dfReportOutput=pd.DataFrame(columns=['Data Processamento','Reference','EmailRemetente','NIF','Apolice','Nome','DataEmail','TemaIdentificado','ViaTratamento','Estado','DetalheMensagem','MensagemOutput','TemaReal'])
    except Exception as e:
        logger.error(f"Erro ao iniciar Aplicações {e}")
        SendOutlookMail(logger,(queryByNameDict('EM01_Body',dictConfig)).replace('[E]',str(e)),queryByNameDict('EM01_Subject',dictConfig),queryByNameDict('EM01_To',dictConfig))
        raise e
    while True:
        dfQueueItem:pd.DataFrame
        try:
            dfQueueItem = databaseSQLExpress.GetQueueItem(db,COLUMN_NAMES,queueItem,tabelaPedidos)
            #print(dfQueueItem)
        except Exception as e:
            logger.error(f"Erro ao tentar is buscar QueueItem {e}")
        
        if not dfQueueItem.empty:
            logger.info(f'A tratar o registo com o EmailID/Reference {dfQueueItem.loc[0,"EmailID"]} e com Intenção Identificada pelo NLP de {dfQueueItem.loc[0,"IDIntencao"]}')
            try:
                #Atividade NLP (Verificar Dados providenciados pelo mesmo)
                try:
                    logger.info('A Verificar Dados Vindos do NLP....')
                    IDbd = dfQueueItem.loc[0,'IDIntencao']
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
                    dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip().casefold().capitalize()
                    dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade do NLP'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA',IDbd)
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
                    raise BusinessRuleException(e)
                except Exception as e:
                    logger.error(f'SystemError no processamento do registo: {e}')
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = ''
                    dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                    dfQueueItem.loc[0,"Mensagem"] = 'Indisponilibidade do NLP'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA',IDbd)
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
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip().casefold().capitalize()
                    dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade de obter dados no GIO'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA',IDbd)
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
                    raise BusinessRuleException(e)
                except Exception as e:
                    logger.error(f'SystemError no processamento do registo: {e}')
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = ''
                    dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                    dfQueueItem.loc[0,"Mensagem"] = 'Indisponilibidade em obter dados no GIO'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA',IDbd)
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed","Erro de Sistema",e)
                    raise Exception(e) 
                #Atividade EMail
                try:
                    logger.info('A Processar a Atividade de Envio de Email.....')
                    EnviarEmail(dfQueueItem,dictConfig,logger)
                    dfQueueItem['DetalheMensagem'] = 'Tratamento realizado com Sucesso'
                    dfQueueItem['Mensagem'] = 'Sucesso no tratamento'
                    dfQueueItem['Estado'] = 'Processado'
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Sucesso",'Processado','Tratamento realizado com Sucesso')
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via Email',IDbd)
                    mail = SearchMailInbox(logger,pastaEmailsTratamento,mailboxname,dfQueueItem.loc[0,"EmailID"])
                    MoveEmailToFolder(logger,pastaEmailsSucesso,mailboxname,mail)
                    logger.info('Registo Processado com Sucesso!')
                except BusinessRuleException as e:
                    logger.error(f"{e}")
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = str(e).split(':')[1].strip().casefold().capitalize()
                    dfQueueItem.loc[0,"Estado"] = 'Definição do Negócio'
                    dfQueueItem.loc[0,"Mensagem"] = 'Impossibilidade em enviar email pelo outlook'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA',IDbd)
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed",'Definição do Negócio',str(e).split(':')[1].strip())
                    raise BusinessRuleException(e)
                except Exception as e:
                    logger.error(f'SystemError no processamento do registo: {e}')
                    #dfQueueItem.loc[0,'IDIntencao'] = IDbd#solucao temporaria
                    dfQueueItem.loc[0,"DetalheMensagem"] = ''
                    dfQueueItem.loc[0,"Estado"] = 'Erro de Sistema'
                    dfQueueItem.loc[0,"Mensagem"] = 'Indisponibilidade em enviar email pelo outlook'
                    dfReportOutput.loc[len(dfReportOutput)] = prepararOutput(dfQueueItem,'Via RNA',IDbd)
                    databaseSQLExpress.UpdateQueueItem(db,dfQueueItem,dfQueueItem.loc[0,"Mensagem"],queueItem,tabelaPedidos,"Failed","Erro de Sistema",e)
                    raise Exception(e)  
            except BusinessRuleException as e:
                mail = SearchMailInbox(logger,pastaEmailsTratamento,mailboxname,dfQueueItem.loc[0,"EmailID"])
                mail.Unread=True
                MoveEmailToFolder(logger,pastaEmailsTratamentoManual,mailboxname,mail)
                logger.error(f'{(str(e).split(":")[1]+":"+str(e).split(":")[2]).strip()}') #Fix para não aparecer "Definição de Negocio:" duas vezes seguidas
            except Exception as e:
                mail = SearchMailInbox(logger,pastaEmailsTratamento,mailboxname,dfQueueItem.loc[0,"EmailID"])
                mail.Unread=True
                MoveEmailToFolder(logger,pastaEmailsTratamentoManual,mailboxname,mail)
                for app in queryByNameDict('AplicacoesPerf',dictConfig).split(','):
                   KillAllApplication(app+'.exe',logger)
                logger.error(f'Erro de Sistema no Processamento do Registo - {e}')
                try:
                    driver = InitApplications(dictConfig)
                except Exception as e:
                    logger.error(f"Erro ao reiniciar Aplicações {e}")
                    SendOutlookMail(logger,(queryByNameDict('EM01_Body',dictConfig)).replace('[E]',str(e)),queryByNameDict('EM01_Subject',dictConfig),queryByNameDict('EM01_To',dictConfig))
                    raise e
        else:
            logger.info("Sem QueueItems para tratar.")    
            break
    #print(dfReportOutput)
    file_path = f'Output\Output_Pedidos_de_Clientes_{datetime.now().strftime("%d%m%Y_%H%M%S")}.xlsx'
    dfReportOutput.drop(columns='TemaReal').to_excel(file_path, index=False, header=True)
    databaseSQLExpress.SetReportOutput(db,'Report_Output',dfReportOutput.drop(columns='TemaReal'))
    #for app in queryByNameDict('AplicacoesPerf',dictConfig).split(','):
    #    KillAllApplication(app+'.exe',logger)
    body = queryByNameDict('EM02_Body',dictConfig).replace('[A]',str(len(dfReportOutput))).replace('[B]',str((dfReportOutput[dfReportOutput['Estado']=='Processado'].shape[0]))).replace('[C]',str((dfReportOutput[dfReportOutput['Estado'] != 'Processado'].shape[0]))).replace('[D]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '0')].shape[0]))).replace('[E]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '1')].shape[0]))).replace('[F]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '2')].shape[0]))).replace('[G]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'].isin(['3', '3NA']))].shape[0]))).replace('[H]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '5')].shape[0]))).replace('[I]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '6')].shape[0]))).replace('[J]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '7')].shape[0]))).replace('[K]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '8')].shape[0]))).replace('[L]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '9')].shape[0]))).replace('[M]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == '10')].shape[0]))).replace('[N]',str((dfReportOutput[(dfReportOutput['Estado'] == 'Processado') & (dfReportOutput['TemaReal'] == 'NA')].shape[0])))
    SendOutlookMail(logger,body,queryByNameDict('EM02_Subject',dictConfig),queryByNameDict('EM02_To',dictConfig))
    logger.info('Performer Terminado!')


def InitApplications(dictConfig):
    outlook_path = queryByNameDict('outlookPath',dictConfig)
    outlook = Application().start(outlook_path)
    time.sleep(5)
    outlook.window(title_re=".*Inbox.*").minimize()
    chrome_path = queryByNameDict('ChromePath',dictConfig)
    urlGIO = queryByNameDict('LinkGIO',dictConfig)
    arguments = fr"--remote-debugging-port=9222 --user-data-dir=C:\selenium\chrome-profile --url {urlGIO} --headless"
    subprocess.Popen([chrome_path] + arguments.split())
    #os.system(r"startChrome.bat")
    Browser_options = Options()
    Browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    browserdriver= queryByNameDict('PathDriverBrowser',dictConfig)
    driver = webdriver.Chrome(service=Service(browserdriver),options=Browser_options)
    driver.maximize_window()
    driver.implicitly_wait(10)
    navegarGIO(driver)
    return driver

def KillAllApplication(processname,logger:logging.Logger):
    try:
        logger.info(f'A Forçar o Fecho da Aplicação {processname}...')
        os.system(f'taskkill /f /im {processname}')
        logger.info(f'{processname} fechado com Sucesso!')
    except Exception as e:
        logger.warning(e)

def prepararOutput(df:pd.DataFrame,viatratamento,idBD):
    for index,row in df.iterrows():
        new_row = {
            'Data Processamento': f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}',
            'Reference': f"{row['EmailID']}",
            'EmailRemetente': f"{row['EmailRemetente']}",
            'NIF': f"{row['NIF']}",
            'Apolice': f"{row['Apolice']}",
            'Nome': f"{row['Nome']}",
            'DataEmail': f"{row['DataEmail']}",
            'TemaIdentificado': f"{idBD}",
            'ViaTratamento': f"{viatratamento}",
            'MensagemOutput': f"{row['Mensagem']}",
            'DetalheMensagem': f'{row["DetalheMensagem"]}',
            'Estado': f"{row['Estado']}",
            'TemaReal':f"{row['IDIntencao']}"
        }
    return new_row

if __name__ == '__main__':
    main()