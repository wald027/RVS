import win32com.client
import pandas as pd
import re
import logging
import os
from customScripts.readConfig import readConfig, queryByNameDict
from customScripts.customLogging import setup_logging
from customScripts.databaseSQLExpress import ConnectToBD


# Reading the configuration and extracting mailbox details
dictConfig = readConfig(r'IntelligentProcessAutomationNLP\ModelNLP\Config_TreinoModelo.xlsx')
mailbox_name = queryByNameDict("MailboxName", dictConfig)
inbox_name = queryByNameDict("InboxFolder", dictConfig)
base_dir = queryByNameDict("Base_Dir", dictConfig)
file_name = "RegrasEmails.xlsx"
regras_emails_path = os.path.join(base_dir, "ModeloNLP", file_name)

# Nome DataBase
DATABASE = queryByNameDict("Database", dictConfig) 
# Server
server = queryByNameDict("SQLExpressServer", dictConfig) 
db = ConnectToBD(server,DATABASE)
setup_logging(db,'LOGS_TREINO') # logger configuration to register
logger = logging.getLogger(__name__) # Cria uma instância do logger


# Initialize Outlook connection and root mailbox
def InitEmailConn(logger, mailbox_name):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace('MAPI')

    try:
        root_folder = mapi.Folders.Item(mailbox_name)
        logger.info(f"Mailbox '{mailbox_name}' found!") 
        return root_folder
    except Exception as e:
        logger.error(f"Error connecting to mailbox: {str(e)}")
        return None

# Function to find a specific folder within the mailbox
def find_folder(parent_folder, folder_name):
    for folder in parent_folder.Folders:
        if folder.Name == folder_name: # validação --> folder_name = 'Modelo de Dados NLP'
            return folder
    return None

def count_non_empty_folders():
    # Initialize the connection to the root mailbox and find the target inbox folder
    root_folder = InitEmailConn(logger, mailbox_name)
    folder = find_folder(root_folder, inbox_name)    
    
    # Count non-empty folders
    count = 0
    value = []
    for subfolder in folder.Folders:
        if subfolder:#.Items or subfolder.Folders:
            print("here subfolder", subfolder)
            count += 1
    return count


# Function to return a dictionary where keys are the label numbers and values are just the folder names
def get_non_empty_folder_labels(logger, mailbox_name):
    # Initialize the connection to the root mailbox and find the target inbox folder
    root_folder = InitEmailConn(logger, mailbox_name)
    folder = find_folder(root_folder, inbox_name)
    
    # Build dictionary of non-empty folder names using label numbers as keys
    label_map = {}
    for subfolder in folder.Folders:
            # Extract the label number and folder name separately
            match = re.match(r"Label (\d+) - (.+)", subfolder.Name)
            if match:
                label_number = int(match.group(1))  # Extract the number and convert it to an integer
                folder_name = match.group(2)        # Extract the folder name part after "Label <number> - "
                label_map[label_number] = folder_name
            else:                
                logger.warning(f"Folder '{subfolder.Name}' does not match the expected 'Label <number> - <name>' format.")
    return label_map



def extract_emails_from_folder(main_folder,dictConfig,logger, labelled=False):
   
    email_list = []

    def process_folder(folder):
        """Recursive function to process a folder and its subfolders."""
        # Process the current folder
        label_names = set()
        label_names.add(folder.Name)
        teste_folder = []
        messages = folder.Items
        
        for message in messages:
            teste_folder = ''
            try:
                #teste_folder.append(folder.name)   
                teste_folder = folder.name             
                label_names.add(folder.Name)
                subject = message.Subject
                sender = message.SenderEmailAddress
                boolDiscard=False
                for emailaddrReject in queryByNameDict('SenderEmailDiscard',dictConfig).split('|'):
                    if emailaddrReject == (message.SenderName + f' <{message.SenderEmailAddress}>'):
                        boolDiscard = EmailRegraDiscard(message,logger,df_regras_emails_ignorar)
                        if boolDiscard:
                            break
                if not boolDiscard:
                    for emailaddr in queryByNameDict('SenderEmailExtract',dictConfig).split('|'): 
                        if emailaddr == (message.SenderName + f' <{message.SenderEmailAddress}>'):
                            body, subject = EmailWithRegraTreino(message,logger)
                            body, subject = EmailRegraPreTratamento(message, df_regras_emails,logger)
                            break
                        else:
                            body = message.Body

                    # Handle the case where ReceivedTime might be None or invalid
                    date = message.ReceivedTime if hasattr(message, 'ReceivedTime') and message.ReceivedTime else None
                    date = str(date).split('+')[0]                    
                    #message_id = message.EntryID
                    property_accessor = message.PropertyAccessor
                    message_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                  
                    # Append the extracted email data to the list                    
                    email_list.append([sender, date, message_id, subject, body, folder.name.split()[1], teste_folder[10:]])                    
                else:
                    print(f'Email para ignorar {message.Subject}')


            except Exception as e:
                logger.error(f"Error processing email: {str(e)}")
        
        # Recursively process subfolders
        for subfolder in folder.Folders:
            process_folder(subfolder)

    # Start processing the main folder and its subfolders
    process_folder(main_folder)

    # Create a pandas DataFrame from the list of email data
    df = pd.DataFrame(email_list, columns=['Email Remetente', "Data Email", "Email ID", 'Subject', 'Body', 'Label', "Nome Label"])    
    # Apply safe conversion to datetime, handling invalid formats and coercing errors
    try:
        pass        # df['data']
    except Exception as e:
        logger.error(f"Error converting Date column to datetime: {str(e)}")
        
    return df

def dataframe(dictConfig):
    mailbox_name = queryByNameDict("MailboxName", dictConfig)
    inbox_name = queryByNameDict("InboxFolder", dictConfig)
    root_folder = InitEmailConn(logger,mailbox_name)
    folder = find_folder(root_folder,inbox_name)
    df = extract_emails_from_folder(folder,dictConfig,logger, labelled=False)
    return df

def EmailWithRegraTreino(mail,logger):

    logger.info(f"Detetado Email com Regra vindo de: {mail.SenderEmailAddress}")
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
        inbox_folder = find_folder(root_folder, inbox_name)

        if inbox_folder:
            logger.info(f"Folder '{inbox_name}' found!")
            df_inbox = extract_emails_from_folder(inbox_folder,dictConfig,logger)
            logger.info(f"Extracted {len(df_inbox)} emails from folder '{inbox_name}'.")
            # Optionally save the DataFrame to a CSV file
            df_inbox.to_excel(f"{inbox_name}_todos_os_dados_teste.xlsx", index=False)
            logger.info(f"Saved inbox emails to '{inbox_name}_emails.xlsx'.")
     
# Execute the main function
if __name__ == "__main__":
    main()
