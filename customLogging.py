import logging      
import databaseSQLExpress
import datetime
import os 

class CustomHandler(logging.StreamHandler):
    def __init__(self,db):
        super().__init__()
        self.db = db
    #juntar file_name, func_name e message
    def emit(self,record):
        columns =['Process','Robot','Time','Level','[User]','Message']
        record.msg ="{} - " +record.msg 
        data = [('RVSIPA2024','RVSIPA2024_{}'.format(os.getlogin()),datetime.datetime.now(),record.levelname,os.getlogin(),record.msg.format(record.filename+" | "+record.funcName))]
        #print(data)#debug
        if record:
            databaseSQLExpress.InsertDataBD(self.db,'LOGS',columns,data)

def setup_logging(db):
    logger = logging.Logger('RealVidaSeguros')
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    Customhandler = CustomHandler(db)
    logger.addHandler(Customhandler)
    #Handler para os logs aparecerem na consola (talvez desativar em producao)
    formatter = logging.Formatter('%(asctime)s | %(filename)s | %(funcName)s | %(levelname)s | %(message)s')
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
