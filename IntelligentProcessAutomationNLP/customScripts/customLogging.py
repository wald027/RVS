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

def setup_logging(db,table,nomeprocesso):
    logger = logging.Logger('RealVidaSeguros')
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    Customhandler = CustomHandler(db,table,nomeprocesso)
    logger.addHandler(Customhandler)
    #Handler para os logs aparecerem na consola (talvez desativar em producao)
    formatter = logging.Formatter('%(asctime)s | %(filename)s | %(funcName)s | %(levelname)s | %(message)s')
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
