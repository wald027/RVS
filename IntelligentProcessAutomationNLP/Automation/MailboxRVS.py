import win32com.client
from datetime import datetime, timedelta
from customScripts.readConfig import queryByNameDict
from customScripts.databaseSQLExpress import *
import pandas as pd
from pywinauto import Application
import time
#eventualmente retirar isto e usar a db connection vinda do dispacther
#server = 'PT-L162219\SQLEXPRESS'
#database = 'RealVidaSeguros'
#dictConfig = readConfig()

#conn = ConnectToBD(server,database)


def InitEmailConn(logger,mailbox_name):
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')

    root_folder = mapi.Folders.Item(mailbox_name)  # Colocar nome mailbox
    logger.info(f"Mailbox {mailbox_name} encontrada!") 
    return root_folder

#Encontrar pasta na mailbox
def find_folder(parent_folder, folder_name):
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return None


def EmailRegraDiscard(mail,logger,current_Mailbox,dfRegrasEmail):
    #dfRegrasEmail = pd.read_excel(r"C:\Users\brunofilipe.lobo\OneDrive - CGI\Code\realvidaseguros\Config.xlsx",sheet_name='RegrasEmailDiscard',keep_default_na=False)

    for index, row in dfRegrasEmail.iterrows():
        for col in dfRegrasEmail.columns:
            if col == 'ExtrairInfo'and row[col] == 'Não' and row['Remetente'] ==mail.SenderName + f' <{mail.SenderEmailAddress}>':
                logger.info(f"Detetado Email com Regra para ignorar vindo de: {mail.SenderName + f' <{mail.SenderEmailAddress}>'}")
                for column in dfRegrasEmail.drop(columns=['Remetente','ExtrairInfo','PastaMover']).columns:
                    if not row[column] == 'NA':
                        if column == 'Subject' and mail.Subject.find(row[column])>-1:
                            logger.info(f'Match com a Regra de {column}, contendo {row[column]}')
                            mail.move(find_folder(current_Mailbox,row['PastaMover']))
                            logger.info(f'Email movido para a pasta {row["PastaMover"]}')
                            boolDiscard = True
                            return boolDiscard
                        elif column == 'Body' and mail.Body.find(row[column])>-1:
                            logger.info(f'Match com a Regra de {column}, contendo {row[column]}')
                            mail.move(find_folder(current_Mailbox,row['PastaMover']))
                            logger.info(f'Email movido para a pasta {row["PastaMover"]}')
                            boolDiscard = True
                            return boolDiscard
    logger.info(f"Sem Match com nenhuma Regra!")

def EmailRegraPreTratamento(mail,logger,dfRegrasEmail):
    #dfRegrasEmail = pd.read_excel(r"C:\Users\brunofilipe.lobo\OneDrive - CGI\Code\realvidaseguros\Config.xlsx",sheet_name='RegrasEmailsPreTratamento',keep_default_na=False)

    for index, row in dfRegrasEmail.iterrows():
        if row['ExtrairInfo'] == 'Sim' and row['Remetente'] == mail.SenderName + f' <{mail.SenderEmailAddress}>':
            logger.info(f"Detetado Email com Regra para pré-tratamento vindo de: {mail.SenderName + f' <{mail.SenderEmailAddress}>'}")
            #Tratar NIF
            try:
                NumIF = (mail.Subject.split(row['NIF'].split('|')[0])[1].split(row['NIF'].split('|')[1])[0]).strip()
                if not NumIF == "" and len(NumIF) > 8 and len(NumIF) < 10 and NumIF.isnumeric() == True:
                    logger.info(f"NIF extraído com sucesso: {NumIF}")
                else:
                    NumIF = ''
                if NumIF == '':
                    NumIF = (mail.Body.split(row['NIF'].split('|')[0])[1].split(row['NIF'].split('|')[1])[0]).strip()
                    if not NumIF == "" and len(NumIF) > 8 and len(NumIF) < 10 and NumIF.isnumeric() == True:
                        logger.info(f"NIF extraído com sucesso: {NumIF}")
                    else:
                        NumIF = ''
            except Exception:
                NumIF =''
            #Tratar Restantes Informações
            for col in dfRegrasEmail.drop(columns=['Remetente','ExtrairInfo','NIF']).columns:
                if not row[col] == 'NA':
                    if mail.Body.find(row[col].split('|')[0].strip()) >-1:
                        try:
                            info = mail.Body.split(row[col].split('|')[0])[1].split(row[col].split('|')[1])[0]
                        except ValueError:
                            info = mail.Body.split(row[col].split('|')[0])[1]
                        except Exception:
                            match col:
                                case 'Body':
                                    info = mail.Body
                                case 'Subject':
                                    info = mail.Subject
                                case _:
                                    info = ""
                        match col:
                            case 'Body':
                                Body = info.strip()
                            case 'Subject':
                                Subject = info.strip()
                            case 'Nome':
                                Nome = info.strip().lower().title()
                            case 'Email':
                                Email = info.strip()
                            case 'Apolice':
                                Apolice = info.strip()
                        logger.info(f'{col} extraído com sucesso: {info.strip()}')
                    elif mail.Body.find(row[col].split('|')[0])>1:
                        try:
                            info = mail.Subject.split(row[col].split('|')[0])[1].split(row[col].split('|')[1])[0]
                        except ValueError:
                            info = mail.Subject.split(row[col].split('|')[0])[1]
                        except Exception:
                            match col:
                                case 'Body':
                                    info = mail.Body
                                case 'Subject':
                                    info = mail.Subject
                                case _:
                                    info = ""
                        match col:
                            case 'Body':
                                Body = info.strip()
                            case 'Subject':
                                Subject = info.strip()
                            case 'Nome':
                                Nome = info.strip().lower().title()
                            case 'Email':
                                Email = info.strip()
                            case 'Apolice':
                                Apolice = info.strip()
                        logger.info(f'{col} extraído com sucesso: {info.strip()}')
                    else:
                        match col:
                            case 'Body':
                                Body = mail.Body
                            case 'Subject':
                                Subject = mail.Subject
                            case 'Nome':
                                Nome = ""
                            case 'Email':
                                Email = ""
                            case 'Apolice':
                                Apolice = ""
                        logger.info(f'Sem Match Com Regra Definida para {col}')
    #return valores extraidos, Apolice não vai na tupple uma vez que nao foi detetado nenhum caso, no entanto está pronto para tal
    return Body, NumIF, Nome, Subject, Email

#Regra Antiga - Funcional mas não usar
def EmailWithRegra(mail,logger):
    Body= ''
    Subject = ''
    logger.info(f"Detetado Email com Regra vindo de: {mail.SenderName + f' <{mail.SenderEmailAddress}>'}")
    NumIF = mail.Subject.split("-")[0].replace("NIF","").replace(" ","")
    if not NumIF == "" and len(NumIF) > 8 and len(NumIF) < 10 and NumIF.isnumeric() == True:
        logger.info(f"NIF extraído com sucesso: {NumIF}")
    else:
        NumIF = ""
    if mail.Body.find('Notas:')>-1:
        Body = mail.body.split('Notas:')[1]
        
    for line in mail.Body.splitlines():
        if line.find('Tipo Assunto:')>-1 or line.find('Assunto:')>-1:
            Subject = line.split(':')[1]
            logger.info(f'Assunto extraído com sucesso: {Subject}')
        if line.find('Nome:')>-1:
            if not '@' in line.split(':')[1]:
                Nome = line.split(':')[1].lower().title().strip()
                Nome = Nome.replace("'",' ')
                logger.info(f"Nome extraído com sucesso: {Nome}")
            else:
                Nome = ''
        if line.find('Email:')>-1:
            Email = line.split(':')[1].lower()
            logger.info(f'Email extraído com sucesso: {Email}')
        if line.find('Mensagem:')>-1 :
            Body = line.split(':')[1]
            logger.info(f'Body extraído com sucesso: {Body}')

    return Body, NumIF, Nome, Subject, Email

            

def GetEmailsInbox(logger,conn,dictConfig,nomeprocesso,tablename,queuetablename):
    #nomeprocesso = queryByNameDict('NomeProcesso',dictConfig)
    #tablename = queryByNameDict("TableName", dictConfig)
    #queuetablename=queryByNameDict('QueueTableName',dictConfig)
    mailbox_name =  queryByNameDict("MailboxName",dictConfig)
    inbox_name= queryByNameDict("InboxFolder",dictConfig)
    folder_toreview = queryByNameDict("EmailsToMove",dictConfig)
    current_Mailbox = InitEmailConn(logger,mailbox_name) #Aceder Diretório Raiz do Email
    current_folder = find_folder(current_Mailbox, inbox_name) #Procurar a inbox do Email
    folder_toMove=find_folder(current_Mailbox,folder_toreview)#Procurar Pasta para onde os emails vão apos lidos
    dfRegrasEmailDiscard = pd.read_excel(queryByNameDict('PathConfigRegrasEmails',dictConfig),sheet_name=queryByNameDict('SheetRegrasEmailDiscard',dictConfig),keep_default_na=False)
    dfRegrasEmailPreTratamento = pd.read_excel(queryByNameDict('PathConfigRegrasEmails',dictConfig),sheet_name=queryByNameDict('SheetRegrasPreTratamento',dictConfig),keep_default_na=False)
    if current_folder:
        logger.info(f"Pasta Encontada: {current_folder.Name}")
        # Aceder aos emails
        messages = current_folder.Items

        #Sample de Filtros
        #received_dt = datetime.now() - timedelta(days=1)
        #received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        #messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
        #messages = messages.Restrict("[SenderEmailAddress] = 'brun0l0b0@outlook.com'")
        #messages = messages.Restrict("[Subject] = 'Sample Report'")
        numEmails = messages.count 
        logger.info(f'Existem {messages.count} emails na pasta {current_folder.Name}') #nome da pasta
        #https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia conteudo email
        for mail in list(messages):
            Attachments='False'
            html_body=mail.HTMLBody
            for attachment in mail.attachments:
                print(attachment.Filename)
                if attachment.Filename not in html_body:
                    Attachments='True'
                    break
            property_accessor = mail.PropertyAccessor
            message_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
            boolDiscard = False
            for emailAddrDiscard in queryByNameDict('SenderEmailDiscard',dictConfig).split('|'):
                if emailAddrDiscard == (mail.SenderName + f' <{mail.SenderEmailAddress}>'):
                    boolDiscard = EmailRegraDiscard(mail,logger,current_Mailbox,dfRegrasEmailDiscard)
                    if boolDiscard:
                        numEmails = numEmails - 1
                        break
            if not boolDiscard:           
                for emailAddr in queryByNameDict("SenderEmailExtract",dictConfig).split('|'):
                    if emailAddr == (mail.SenderName + f' <{mail.SenderEmailAddress}>'):
                        Body, NumIF, Nome, Subject, Email =EmailRegraPreTratamento(mail,logger,dfRegrasEmailPreTratamento)
                        #Body, NumIF, Nome, Subject, Email = EmailWithRegra(mail,logger)
                        columns =['EmailRemetente','DataEmail','EmailID','Subject','Body','Anexos','NIF','Nome']
                        data = [(Email,mail.SentOn,message_id,Subject,Body,Attachments,NumIF,Nome)]    
                        #mail.Subject =Subject
                        break
                    else:
                        data = [(mail.SenderEmailAddress,mail.SentOn,message_id,mail.Subject,mail.Body,Attachments)]
                        columns =['EmailRemetente','DataEmail','EmailID','Subject','Body','Anexos']
                logger.info(f"Sender: {mail.SenderEmailAddress} Subject:{mail.Subject} Recebido: {mail.senton} Message-ID: {message_id} Attachments:{Attachments}")#Enviar BD e Logs
                try:
                    InsertDataBD(conn,tablename,columns,data)
                    logger.info("Email Enviado com Sucesso para a Base de Dados!")
                    columns =['Status','Reference','SpecificContent','Process']
                    data = [('NLP',message_id,''.join(map(str, data)),nomeprocesso)]
                    InsertDataBD(conn,queuetablename,columns,data)
                    if mail.Unread:
                        mail.Unread=False
                        mail.save()
                    mail.move(folder_toMove)
                    logger.info(f"Email Movido para a Pasta {folder_toMove}")
                    time.sleep(3)
                except Exception as e:
                    logger.error(f"Erro ao tentar inserir Info na Base de Dados: {e}")
                    numEmails = numEmails -1
        return numEmails
    else:
        logger.warn(f"Pasta: {inbox_name} não encontrada!")
    #if conn:
    #    conn.close

def SearchMailInbox(logger,pastapesquisar,mailbox,emailID):
    root_folder = InitEmailConn(logger,mailbox)
    foldertratamento = find_folder(root_folder,pastapesquisar)
    messages = foldertratamento.Items

    for mail in list(messages):
        property_accessor = mail.PropertyAccessor
        message_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
        if message_id == emailID:
            logger.info('Email Encontrado!')
            return mail

def MoveEmailToFolder(logger,pastatomove,mailbox,mail):
    time.sleep(5)
    root_folder = InitEmailConn(logger,mailbox)
    foldertomove = find_folder(root_folder,pastatomove)
    try:
        mail.Move(foldertomove)
        time.sleep(5)
        logger.info(f'Email Movido Para a Pasta {pastatomove} com Sucesso!')
    except Exception as e:
        logger.error(f'Impossibilidade em Mover Email para a Pasta {pastatomove}')

def SendEmail(logger,body,subject,To):

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0) 

    mail.Subject = subject
    mail.Body = body
    mail.To = To 

    # Send the email
    mail.Display()

    time.sleep(15)
    app = Application(backend='uia').connect(title_re='.*Message.*')
    main_window = app.window(title_re='.*Message.*')
    main_window.set_focus()
    try:
        main_window.child_window(title="Non-Business", control_type="ListItem").click_input()
    except:
        logger.info('Sem Label de Classificação')
    try:
        main_window.child_window(title="Send", control_type="Button").click_input()
        logger.info(f'Email para {To} , Enviado com Sucesso!')
    except Exception as e:
        logger.error(f'Impossibilidade em enviar o Email: {e}')
        raise Exception('Impossibilidade em enviar o Email')


    logger.info(f'Email enviado para {To}')