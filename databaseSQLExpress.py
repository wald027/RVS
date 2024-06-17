import pyodbc

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

def InsertDataBD(conn, table_name,columns,data):
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
