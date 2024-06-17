import logging      
import databaseSQLExpress
import datetime

class CustomHandler(logging.StreamHandler):
    def __init__(self,db):
        super().__init__()
        self.db = db

    def emit(self,record):
        columns =['File_Name','Func_Name','Log_Level','Message','Load_Time']
        data = [(record.filename,record.funcName,record.levelname,record.msg,datetime.datetime.now())]
        if record:
            databaseSQLExpress.InsertDataBD(self.db,'LOGS',columns,data)

def setup_logging(db):
    logger = logging.Logger('RealVidaSeguros')
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    Customhandler = CustomHandler(db)
    logger.addHandler(Customhandler)
    #Handler para os logs aparecerem na consola (talvez desativar em producao)
    formatter = logging.Formatter('%(asctime)s | %(filename)s | %(funcName)s | %(levelname)s | %(message)s')
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
