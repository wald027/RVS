import win32com.client
import pandas as pd
import logging
from customScripts.readConfig import readConfig, queryByNameDict
from customScripts.customLogging import setup_logging
from customScripts.databaseSQLExpress import ConnectToBD


# Reading the configuration and extracting mailbox details
dictConfig = readConfig(r'C:\Users\nayara.rodrigues\Documents\RVS-main\Config_TreinoModelo.xlsx')
mailbox_name = queryByNameDict("MailboxName", dictConfig)
inbox_name = queryByNameDict("InboxFolder", dictConfig)
folder_toreview = queryByNameDict("EmailsToMove", dictConfig)

DATABASE = "RealVidaSeguros" # Nome do banco de dados
server = 'PT-L163255\SQLEXPRESS01'
#Initialize logging (optional, for debugging purposes)
#logging.basicConfig(level=logging.INFO)
#logger = logging.getLogger()
db = ConnectToBD(server,DATABASE)
setup_logging(db,'LOGS_TREINO') # Configura o logger para registrar logs no banco de dados
logger = logging.getLogger(__name__) # Cria uma instância do logger

df_regras_emails = pd.read_excel(r"C:\Users\nayara.rodrigues\Documents\RVS-main\IntelligentProcessAutomationNLP\ModelNLP\RegrasEmails.xlsx",sheet_name='RegrasEmailsPreTratamento',keep_default_na=False)
df_regras_emails_ignorar = pd.read_excel(r"C:\Users\nayara.rodrigues\Documents\RVS-main\IntelligentProcessAutomationNLP\ModelNLP\RegrasEmails.xlsx",sheet_name='RegrasEmailDiscard',keep_default_na=False)
# Initialize Outlook connection and root mailbox
def InitEmailConn(logger, mailbox_name):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    try:
        root_folder = mapi.Folders.Item(mailbox_name)  # Replace with mailbox name
        logger.info(f"Mailbox '{mailbox_name}' found!") 
        return root_folder
    except Exception as e:
        logger.error(f"Error connecting to mailbox: {str(e)}")
        return None

# Function to find a specific folder within the mailbox
'''
def find_folder(parent_folder, folder_name):
    for folder in parent_folder.Folders:
        if folder.Name == folder_name: # validação
            return folder
    return None
'''
def find_folder(parent_folder, folder_name):
    for folder in parent_folder.Folders:
        if folder.Name == folder_name: # validação folder_name = 'Modelo de Dados NLP'  Name= 'Deleted Items'
            return folder
    return None
    

def extract_emails_from_folder(main_folder,dictConfig,logger, labelled=False):
    """
    Extract emails from all folders and subfolders within the main folder, 
    and return a pandas DataFrame containing all email data.
    """
    email_list = []

    def process_folder(folder):
        """Recursive function to process a folder and its subfolders."""
        # Process the current folder
        #print(f"Nome folder {folder.Name}")
        label_names = set()
        label_names.add(folder.Name)
        #print(f"Messagem here: {label_names}")
        teste_folder = []
        messages = folder.Items
        
        for message in messages:
            #print(f"Messagem here: {message}")
            try:
                teste_folder.append(folder.name)                
                label_names.add(folder.Name)
                #print(f"label_names: {teste_folder}")
                subject = message.Subject
                sender = message.SenderEmailAddress
                #to = message.To
                boolDiscard=False
                for emailaddrReject in queryByNameDict('SenderEmailDiscard',dictConfig).split('|'):
                    if emailaddrReject == (message.SenderName + f' <{message.SenderEmailAddress}>'):
                        boolDiscard = EmailRegraDiscard(message,logger,df_regras_emails_ignorar)
                        if boolDiscard:
                            break
                if not boolDiscard:
                    for emailaddr in queryByNameDict('SenderEmailExtract',dictConfig).split('|'): 
                        if emailaddr == (message.SenderName + f' <{message.SenderEmailAddress}>'):
                            #body, subject = EmailWithRegraTreino(message,logger)
                            body, subject = EmailRegraPreTratamento(message,logger, df_regras_emails)
                            break
                        else:
                            body = message.Body

                    # Handle the case where ReceivedTime might be None or invalid
                    date = message.ReceivedTime if hasattr(message, 'ReceivedTime') and message.ReceivedTime else None
                    date = str(date).split('+')[0]
                    #message_id = message.EntryID
                    property_accessor = message.PropertyAccessor
                    message_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                    # Customize labeling logic if needed
                    if labelled:
                        label = "some_label"  # Placeholder for custom logic
                        template = "some_template"  # Placeholder for custom logic
                        amostragem = "some_amostragem"  # Placeholder for custom logic
                    else:
                        label = None
                        template = None
                        amostragem = None

                    # Append the extracted email data to the list
                    #email_list.append([message_id, sender, subject, to, date, body, label, template, amostragem])
                    #email_list.append([message_id, sender, subject, to, date, body, folder.name, template, amostragem])
                    email_list.append([sender, date, message_id, subject, body, folder.name.split()[1], teste_folder])
                else:
                    print(f'Email para ignorar {message.Subject}')


            except Exception as e:
                logger.error(f"Error processing email: {str(e)}")
        print(teste_folder)
        print("quantidade folder", len(teste_folder))
        # Recursively process subfolders
        for subfolder in folder.Folders:
            process_folder(subfolder)

    # Start processing the main folder and its subfolders
    process_folder(main_folder)

    # Create a pandas DataFrame from the list of email data
    df = pd.DataFrame(email_list, columns=['Email Remetente', "Data Email", "Email ID", 'Subject', 'Body', 'Label', "Name Folders"])
    #print(f"{df['Name Folders'][0]}")
    
    # Clean the Date column to ensure valid datetime formats
    #df['Date'] = df['Date'].apply(lambda x: pd.NaT if pd.isnull(x) or x is None else x)

    # Apply safe conversion to datetime, handling invalid formats and coercing errors
    try:
        print('HI')
        #df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    except Exception as e:
        logger.error(f"Error converting Date column to datetime: {str(e)}")
    '''
    for index, row in df.iterrows():
        for emailaddr in queryByNameDict('SenderEmailException',dictConfig).split('|'):
            if row['From'] == emailaddr:
                row = EmailWithRegraTreino(row,logger)
                break
    '''           
    return df

def dataframe(logger,dictConfig):
    mailbox_name = queryByNameDict("MailboxName", dictConfig)
    inbox_name = queryByNameDict("InboxFolder", dictConfig)
    #inbox_name = 'Modelo de Dados'
    root_folder = InitEmailConn(logger,mailbox_name)
    folder = find_folder(root_folder,inbox_name)
    df = extract_emails_from_folder(folder,dictConfig,logger, labelled=False)
    return df

def EmailWithRegraTreino(mail,logger):
    logger.info(f"Detetado Email com Regra vindo de: {mail.SenderEmailAddress}")

    print(mail.SenderName + f' <{mail.SenderEmailAddress}>')

    if mail.Body.find('Notas:')>-1:
        Body = mail.body.split('Notas:')[1]
    for line in mail.Body.splitlines():
        if line.find('Tipo Assunto:')>-1 or line.find('Assunto:')>-1:
            Subject = line.split(':')[1]
            logger.info(f'Assunto extraído com sucesso: {Subject}')
        if line.find('Mensagem:')>-1:
            Body = line.split(':')[1]
            logger.info(f'Body extraído com sucesso: {Body}')

    return Body, Subject

# Main function to initialize and extract emails

def EmailRegraPreTratamento(mail,logger,dfRegrasEmail):
    #dfRegrasEmail = pd.read_excel(r"C:\Users\brunofilipe.lobo\OneDrive - CGI\Code\realvidaseguros\Config.xlsx",sheet_name='RegrasEmailsPreTratamento',keep_default_na=False)
    print(dfRegrasEmail)
    for index, row in dfRegrasEmail.iterrows():
        if row['ExtrairInfo'] == 'Sim' and row['Remetente'] == mail.SenderName + f' <{mail.SenderEmailAddress}>':
            logger.info(f"Detetado Email com Regra para pré-tratamento vindo de: {mail.SenderName + f' <{mail.SenderEmailAddress}>'}")
            #Tratar Restantes Informações
            for col in dfRegrasEmail.drop(columns=['Remetente','ExtrairInfo']).columns:
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
                        logger.info(f'{col} extraído com sucesso: {info.strip()}')
                    else:
                        match col:
                            case 'Body':
                                Body = mail.Body
                            case 'Subject':
                                Subject = mail.Subject
                        logger.info(f'Sem Match Com Regra Definida para {col}')

    return Body, Subject

def EmailRegraDiscard(mail,logger,dfRegrasEmail):
    #dfRegrasEmail = pd.read_excel(r"C:\Users\brunofilipe.lobo\OneDrive - CGI\Code\realvidaseguros\Config.xlsx",sheet_name='RegrasEmailDiscard',keep_default_na=False)

    for index, row in dfRegrasEmail.iterrows():
        for col in dfRegrasEmail.columns:
            if col == 'ExtrairInfo'and row[col] == 'Não' and row['Remetente'] ==mail.SenderName + f' <{mail.SenderEmailAddress}>':
                logger.info(f"Detetado Email com Regra para ignorar vindo de: {mail.SenderName + f' <{mail.SenderEmailAddress}>'}")
                for column in dfRegrasEmail.drop(columns=['Remetente','ExtrairInfo']).columns:
                    if not row[column] == 'NA':
                        if column == 'Subject' and mail.Subject.find(row[column])>-1:
                            logger.info(f'Match com a Regra de {column}, contendo {row[column]}')
                            boolDiscard = True
                            return boolDiscard
                        elif column == 'Body' and mail.Body.find(row[column])>-1:
                            logger.info(f'Match com a Regra de {column}, contendo {row[column]}')
                            boolDiscard = True
                            return boolDiscard
    logger.info(f"Sem Match com nenhuma Regra!")

def main():
    root_folder = InitEmailConn(logger, mailbox_name)

    if root_folder:
        # Locate the inbox and folder to review
        print(f'Aqui está meu root_folder {root_folder}')
        inbox_folder = find_folder(root_folder, inbox_name)
        review_folder = find_folder(root_folder, folder_toreview)

        if inbox_folder:
            print('ola1231')
            logger.info(f"Folder '{inbox_name}' found!")
            df_inbox = extract_emails_from_folder(inbox_folder,dictConfig,logger)

            logger.info(f"Extracted {len(df_inbox)} emails from folder '{inbox_name}'.")

            # Optionally save the DataFrame to a CSV file
            #df_inbox.to_excel(f"{inbox_name}_emails15.xlsx", index=False)
            logger.info(f"Saved inbox emails to '{inbox_name}_emails.xlsx'.")
     
# Execute the main function
if __name__ == "__main__":
    main()
