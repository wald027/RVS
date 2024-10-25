import logging      
from customScripts import databaseSQLExpress 
import datetime
import os 

class CustomHandler(logging.StreamHandler):
    def __init__(self,db,table,nomeprocesso):
        super().__init__()
        self.db = db
        self.table=table
        self.nomeprocesso=nomeprocesso
    #juntar file_name, func_name e message
    def emit(self,record):
        columns =['Process','Robot','Time','Level','[User]','Message']
        msg = "{} - "+record.msg.replace('{','').replace('}','')
        if '{' in msg and '}' in msg:
            msg = msg.format(record.filename + " | " + record.funcName)
       #record.msg ="{} - " +record.msg
        data = [(self.nomeprocesso,f'{os.environ["COMPUTERNAME"]}_{os.getlogin()}',datetime.datetime.now(),record.levelname,os.getlogin(),msg)]
        #print(data)#debug
        if record:
            databaseSQLExpress.InsertDataBD(self.db,self.table,columns,data)

#class provavelmente pode ser retirada
class CustomFilter(logging.Filter):
    def filter(self, record):
        if not hasattr(record, 'user'):
            record.user = os.getlogin()  # Hardcoded user
        if not hasattr(record, 'robot'):
            record.robot = f'{os.environ["COMPUTERNAME"]}_{os.getlogin()}'  # Hardcoded robot
        return True
    
class CustomFormatter(logging.Formatter):
    def format(self, record):
        if not hasattr(record, 'user'):
            record.user = os.getlogin()  # Hardcoded user
        if not hasattr(record, 'robot'):
            record.robot = f'{os.environ["COMPUTERNAME"]}_{os.getlogin()}'  # Hardcoded robot
        return super().format(record)
    
def setup_logging() -> logging.Logger:

    os.makedirs('Logs', exist_ok=True)

    logger = logging.getLogger('RealVidaSeguros')
    logger.setLevel(logging.DEBUG)
    
    #Handler para os logs aparecerem na consola (talvez desativar em producao)
    formatter = CustomFormatter('%(asctime)s | %(levelname)s | %(robot)s | %(filename)s | %(funcName)s - %(message)s')
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    #Handler de INFO level logs
    file_handler_info = logging.FileHandler(f'Logs\Info_Logs_ObterEmails_{datetime.datetime.now().strftime("%d%m%Y_%H%M%S")}.txt', encoding='utf-8')
    file_handler_info.setLevel(logging.INFO)
    file_handler_info.setFormatter(formatter)
    logger.addHandler(file_handler_info)

    #Handler de Debug level logs
    file_handler_debug = logging.FileHandler(f'Logs\Debug_Logs_ObterEmails_{datetime.datetime.now().strftime("%d%m%Y_%H%M%S")}.txt',encoding='utf-8')
    file_handler_debug.setLevel(logging.DEBUG)
    file_handler_debug.setFormatter(formatter)
    logger.addHandler(file_handler_debug)

    logger.addFilter(CustomFilter())

    #logger.debug("Logging para consola e ficheiros txt inicializado.")

    return logger

def setup_logging_db(db,table,nomeprocesso)-> logging.Logger:
    #Inicializa os logs para a base dados
    logger = logging.getLogger('RealVidaSeguros')
    logger.setLevel(logging.DEBUG)
    Customhandler = CustomHandler(db,table,nomeprocesso)
    logger.addHandler(Customhandler)

    #logger.debug("Logging connectado com a base de dados!")

    return logger