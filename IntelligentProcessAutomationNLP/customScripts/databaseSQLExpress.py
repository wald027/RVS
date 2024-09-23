import pyodbc
import pandas as pd
import datetime
import os 

def ConnectToBD(server, database):
    connection_string = (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={server};'
        f'DATABASE={database};'
        f'Trusted_Connection=yes;'
    )
    try:
        conn = pyodbc.connect(connection_string)
        print("Ligação bem Sucedida")
        return conn
    except Exception as e:
        print(f"Erro a Ligar à BD {e}")
        return None

def InsertDataBD(conn:pyodbc.Connection, table_name,columns,data):
    cursor = conn.cursor()
    placeholders = ', '.join(['?']*len(columns))
    columns_str =', '.join(columns)
    sql_insert=f"INSERT INTO {table_name} ({columns_str}) values ({placeholders})"

    for row in data:
        if len(row) != len(columns):
            print(len(row),' ',len(columns))
            raise ValueError("Numero de rows não é igual às columns")
        cursor.execute(sql_insert, row)

    conn.commit()
    #print("Info Inserida!")
    cursor.close()

def GetQueueItem(conn:pyodbc.Connection,column_names,QueueTable,InfoTable):
    cursor = conn.cursor()
    query = f"""
            SELECT TOP(1) {', '.join(column_names)}
            FROM {InfoTable} et
            JOIN {QueueTable} st ON et.EmailID = st.Reference
            WHERE st.Status = 'NLP FINISHED' OR st.Status = 'NLP FAILED';
        """
    cursor.execute(query)
    results = [list(row) for row in cursor.fetchall()]
    df = pd.DataFrame(results,columns=column_names)

    for i in df['EmailID']:
        query = f"""
                    UPDATE {QueueTable}
                    Set Status = 'In Progress' , [Started Performer] = GETDATE(),Robot = '{os.environ["COMPUTERNAME"]}_{os.getlogin()}'
                    WHERE Reference = '{i}';
                """ 
        cursor.execute(query)
    cursor.close()
    return df

def UpdateQueueItem(conn:pyodbc.Connection, df:pd.DataFrame,mensagem,QueueTable,InfoTable,estado,exception,exception_message):
    cursor = conn.cursor()
    for i in df['EmailID']:
        query = f"""
                    Update {QueueTable}
                    Set Status = '{estado}', [Ended Performer] = GETDATE(),Exception='{exception}',Exception_Message='{exception_message}'
                    WHERE Reference = '{i}';
                """
        cursor.execute(query)
        query = f"""
                    Update {InfoTable}
                    Set DetalheMensagem = '{exception_message}', Mensagem ='{mensagem}' , Estado='{exception}'
                    Where EmailID = '{i}';
                """
        cursor.execute(query)
    cursor.close()



def SetReportOutput(conn:pyodbc.Connection, table_name,df:pd.DataFrame):
    cursor = conn.cursor()
    placeholders = ', '.join(['?']*len(df.columns))
    columns_str =', '.join([f'[{col}]' for col in df.columns])
    sql_insert=f"INSERT INTO {table_name} ({columns_str}) values ({placeholders})"

    for i, row in df.iterrows():
        row_as_str = [str(value).replace(",","") for value in row]
        row_as_str = [None if value.strip() == "" else value for value in row_as_str]
        cursor.execute(sql_insert,tuple(row_as_str))
    conn.commit()
    cursor.close()