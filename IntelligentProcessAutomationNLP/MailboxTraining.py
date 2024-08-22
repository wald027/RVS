import win32com.client
import pandas as pd
import logging
from customScripts.readConfig import readConfig, queryByNameDict
from customScripts.customLogging import setup_logging
from customScripts.databaseSQLExpress import ConnectToBD


# Reading the configuration and extracting mailbox details
#dictConfig = readConfig()
#mailbox_name = queryByNameDict("MailboxName", dictConfig)
#inbox_name = queryByNameDict("InboxFolder", dictConfig)
#folder_toreview = queryByNameDict("EmailsToMove", dictConfig)
"""
DATABASE = "RealVidaSeguros" # Nome do banco de dados
server = 'PT-L164962\SQLEXPRESS'
# Initialize logging (optional, for debugging purposes)
#logging.basicConfig(level=logging.INFO)
#logger = logging.getLogger()
db = ConnectToBD(server,DATABASE)
setup_logging(db,'LOGS') # Configura o logger para registrar logs no banco de dados
logger = logging.getLogger(__name__) # Cria uma instância do logger
"""

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
def find_folder(parent_folder, folder_name):
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
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
        teste_folder = []
        messages = folder.Items
        for message in messages:
            try:
                teste_folder.append(folder.name)
                subject = message.Subject
                sender = message.SenderEmailAddress
                to = message.To
                for emailaddr in queryByNameDict('SenderEmailException',dictConfig).split('|'): 
                    if emailaddr == sender:
                        body, subject = EmailWithRegraTreino(message,logger)
                        break
                    else:
                        body = message.Body

                # Handle the case where ReceivedTime might be None or invalid
                date = message.ReceivedTime if hasattr(message, 'ReceivedTime') and message.ReceivedTime else None
                date = str(date).split('+')[0]
                message_id = message.EntryID

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
                email_list.append([message_id, sender, subject, to, date, body, folder, template, amostragem])

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
    df = pd.DataFrame(email_list, columns=['Message ID', 'From', 'Subject', 'To', 'Date', 'Body', 'Label', 'Label Template', 'Amostragem'])
    #df = pd.DataFrame(email_list, columns=['Message ID', 'From', 'Subject', 'To', 'Body', 'Body', 'Label', 'Label Template', 'Amostragem'])
    #df.to_excel('Mudanças.xlsx')
    print(df)
    
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
    #inbox_name = queryByNameDict("InboxFolder", dictConfig)
    inbox_name = 'Modelo de Dados'
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
""""
def main():
    root_folder = InitEmailConn(logger)

    if root_folder:
        # Locate the inbox and folder to review
        inbox_folder = find_folder(root_folder, inbox_name)
        review_folder = find_folder(root_folder, folder_toreview)

        if inbox_folder:
            logger.info(f"Folder '{inbox_name}' found!")
            df_inbox = extract_emails_from_folder(inbox_folder)
            logger.info(f"Extracted {len(df_inbox)} emails from folder '{inbox_name}'.")

            # Optionally save the DataFrame to a CSV file
            df_inbox.to_excel(f"{inbox_name}_emails10.xlsx", index=False)
            logger.info(f"Saved inbox emails to '{inbox_name}_emails.xlsx'.")
        '''
        if review_folder:
            logger.info(f"Folder '{folder_toreview}' found!")
            df_review = extract_emails_from_folder(review_folder)
            logger.info(f"Extracted {len(df_review)} emails from folder '{folder_toreview}'.")

            # Optionally save the DataFrame to a CSV file
            df_review.to_excel(f"{folder_toreview}_emails.xlsx", index=False)
            logger.info(f"Saved reviewed emails to '{folder_toreview}_emails.xlsx'.")
        '''
# Execute the main function
if __name__ == "__main__":
    main()
"""