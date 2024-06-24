import win32com.client
from datetime import datetime, timedelta
from readConfig import readConfig, queryByNameDict
from databaseSQLExpress import *

#eventualmente retirar isto e usar a db connection vinda do dispacther
server = 'PT-L162219\SQLEXPRESS'
database = 'RealVidaSeguros'
dictConfig = readConfig()
mailbox_name =  queryByNameDict("MailboxName",dictConfig)
inbox_name= queryByNameDict("InboxFolder",dictConfig)
folder_toreview = queryByNameDict("EmailsToMove",dictConfig)
conn = ConnectToBD(server,database)
tablename = queryByNameDict("TableName", dictConfig)

def InitEmailConn(logger):
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

def EmailWithRegra(mail,logger):
    logger.info(f"Detetado Email com Regra vindo de: {mail.SenderEmailAddress}")
    NumIF = mail.Subject.split("-")[0].replace("NIF","").replace(" ","")
    if not NumIF == "" and len(NumIF) > 8 and len(NumIF) < 10:
        logger.info(f"NIF extraído com sucesso: {NumIF}")
    for line in mail.Body.splitlines():
        if line.find('Nome:')>-1:
            Nome = line.split(':')[1].lower().title()
            logger.info(f"Nome extraído com sucesso: {Nome}")
            break
    Body = mail.Body.split('Notas:')[1]
    return Body, NumIF, Nome

            

def GetEmailsInbox(logger):
    current_Mailbox = InitEmailConn(logger)
    current_folder = find_folder(current_Mailbox, inbox_name)
    folder_toMove=find_folder(current_Mailbox,folder_toreview)
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

        logger.info(f'Existem {messages.count} emails na pasta {current_folder.Name}') #nome da pasta
        Attachments='False'
        #https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia conteudo email
        for mail in list(messages):
            html_body=mail.HTMLBody
            for attachment in mail.attachments:
                if attachment.Filename not in html_body:
                    Attachments='True'
                    break
            property_accessor = mail.PropertyAccessor
            message_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
            for emailAddr in queryByNameDict("SenderEmailException",dictConfig).split('|'):
                if emailAddr == mail.SenderEmailAddress:
                    Body, NumIF, Nome = EmailWithRegra(mail,logger)
                    columns =['EmailRemetente','DataEmail','EmailID','Subject','Body','Anexos','NIF','Nome']
                    data = [(mail.SenderEmailAddress,mail.SentOn,message_id,mail.Subject,Body,Attachments,NumIF,Nome)]    
                    break
                else:
                    data = [(mail.SenderEmailAddress,mail.SentOn,message_id,mail.Subject,mail.Body,Attachments)]
                    columns =['EmailRemetente','DataEmail','EmailID','Subject','Body','Anexos']

            logger.info(f"Sender: {mail.SenderEmailAddress} Subject:{mail.Subject} Recebido: {mail.senton} Message-ID: {message_id} Attachments:{Attachments}")#Enviar BD e Logs
            try:
                InsertDataBD(conn,tablename,columns,data)
                logger.info("Email Enviado com Sucesso para a Base de Dados!")
                columns =['Status','Reference','SpecificContent','Process']
                data = [('NLP',message_id,'Teste','RVSIPA2024')]
                InsertDataBD(conn,"QueueItem",columns,data)
                if mail.Unread:
                    mail.Unread=False
                    mail.save()
                mail.move(folder_toMove)
            except Exception as e:
                logger.error(f"Erro ao tentar inserir Info na Base de Dados: {e}")
#            if mail.attachments.Count > 0:
#                print("Attachments: True")
#            else:
#                print("Attachments: False")
    else:
        logger.warn(f"Pasta: {inbox_name} não encontrada!")
    if conn:
        conn.close

