# import sqlalchemy as db
# from sqlalchemy import text
import pandas as pd
import time
from transformers import BertTokenizer, BertForSequenceClassification, pipeline
from ModelNLP import helpers



class EmailClassifier:
    def __init__(self, base_dir: str, num_labels: int, status_table: str, email_table: str, column_names: list[str], engine, logger,db) -> None:
        # Inicializa a classe com os parâmetros necessários
        self.base_dir = base_dir # Diretório base onde estão armazenados o modelo e o tokenizador
        self.num_labels = num_labels # Número de rótulos de classificação
        self.status_table = status_table # Nome da tabela de status no banco de dados
        self.email_table = email_table # Nome da tabela de e-mails no banco de dados
        self.column_names = column_names # Lista de nomes das colunas a serem extraídas
        self.engine = engine # Conexão com o banco de dados
        self.cursor = db.cursor() # Cursor para executar comandos SQL
        self.logger = logger # Logger para registrar informações e erros

    def get_emails(self) -> pd.DataFrame:
        # Carrega dados do banco de dados
        query = f"""
            SELECT {', '.join(self.column_names)}
            FROM {self.email_table} et
            JOIN {self.status_table} st ON et.EmailID = st.Reference
            WHERE st.Status = 'NLP';
        """
        # Query para selecionar os e-mails com status 'NLP'
        df = pd.read_sql_query(query, con=self.engine) # Executa a consulta SQL e carrega os resultados em um DataFrame
        self.logger.debug(df) # Log dos dados carregados
        self.logger.info("LOADED - emails data.") # Log informando que os dados foram carregados


        # Atualiza o status dos e-mails para 'NLP IN PROGRESS'
        for i in df["EmailID"].tolist():   
            query = f"""
                UPDATE {self.status_table}
                SET Status = 'NLP IN PROGRESS'
                WHERE Reference = '{i}';
            """  # text() is needed when using sqlalchemy
            self.cursor.execute(query) # Executa a atualização para cada EmailID
        self.logger.info("UPDATED - NLP in progress.") # Log informando que o status foi atualizado
        return df
    
    def get_predictions(self, df: pd.DataFrame) -> pd.DataFrame:
        # Carrega o tokenizador e o modelo BERT
        model = BertForSequenceClassification.from_pretrained(self.base_dir.strip('/')+'/IntelligentProcessAutomationNLP/ModelNLP/model', num_labels=self.num_labels)
        tokenizer = BertTokenizer.from_pretrained(self.base_dir.strip('/')+'/IntelligentProcessAutomationNLP/ModelNLP/tokenizer', truncation=True, padding="max_length", max_length=128)
        self.logger.info("LOADED - model and tokenizer.") # Informa que o modelo e o tokenizador foram carregados

        # Obter predições
        clf = pipeline("text-classification", model=model, tokenizer=tokenizer, truncation=True, padding="max_length", max_length=128)
        predictions = clf(df["Text"].to_list()) # Gera predições para a lista de textos
        labels = [str(int(pred["label"].split('_')[1])) for pred in predictions] # Extrai os rótulos das predições
        scores = [str(round(float(pred["score"]), 2)) for pred in predictions] # Extrai as pontuações de confiança das predições
        self.logger.info("PREDICTIONS - generated.") # Informa que as predições foram geradas

        df["IDIntencao"] = labels # Adiciona os rótulos previstos ao DataFrame
        df["Score"] = scores # Adiciona as pontuações de confiança ao DataFrame
        self.logger.info("PREDICTIONS - inserted on dataframe.") # Informa que as predições foram inseridas no DataFrame
        return df # Retorna o DataFrame com as predições

    def update_row(self, row):
            # Obtém informações adicionais (Nome, Apólice) do banco de dados para um email específico
            query = f"""
                SELECT Nome, Apolice
                From {self.email_table}
                WHERE EmailID = '{row['EmailID']}';
                """ 

            self.cursor.execute(query)
            results = [list(x) for x in self.cursor.fetchall()]  # Extrai os resultados da consulta      
            # Atualiza os campos 'Nome' e 'NIF' com as informações adicionais se disponíveis    
            if results[0][0] is not None:
                if row['Nome']:
                    row['Nome'] = results[0][0] + "|" + row['Nome']
                else:
                    row['Nome'] = results[0][0]
            if results[0][1] is not None:
                if row['NIF']:
                    row['NIF'] = results[0][1] + "|" + row['NIF']
                else:
                    row['NIF'] = results[0][1]
            # Prepara a query para atualizar o banco de dados com as novas informações
            data = ', '.join([f"{col} = '{row[col]}'" for col in
                              ['NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao', 'Score', 'IDTermosExpressoes']])
                                   
            query = f"""
                UPDATE {self.email_table}
                SET {data}
                WHERE EmailID = '{row['EmailID']}';
            """
            self.cursor.execute(query)  # Executa a query de atualização

    def update_database(self, df) -> None:
        # Salva as atualizações no banco de dados
        df = df[self.column_names] # Seleciona as colunas relevantes
        df.apply(self.update_row, axis=1) # Aplica a função de atualização para cada linha do DataFrame

        # Atualiza o status dos emails para "NLP FINISHED"
        for i in df['EmailID'].tolist():
            query = f"""
                UPDATE {self.status_table}
                SET Status = 'NLP FINISHED'
                WHERE Reference = '{i}';
            """
            self.cursor.execute(query)

        time.sleep(3) # Espera 3 segundos
        # Atualiza o status para "NLP FAILED" para os emails que ainda estão em progresso
        query = f"""
            UPDATE {self.status_table}
            SET Status = 'NLP FAILED'
            WHERE Status = 'NLP IN PROGRESS';
        """
        self.cursor.execute(query)
        self.logger.info("UPDATED - NLP finished.") # Informa que o processo de NLP foi finalizado

    def run(self):
        # Executa o processo de classificação de emails
        df = self.get_emails() # Obtém os dados dos emails
        print(df)
        # Limpa o corpo do email, removendo sentenças não-portuguesas, caracteres não-imprimíveis e espaços extras
        df["Text"] = df["Body"].copy()
        
        # Identifica emails respondidos ou com histórico de conversas
        df["HistoricoEmails"] = df.apply(lambda x: helpers.get_historico(x["Body"], x['EmailID'], self.logger), axis=1)
        # Identifica textos vazios e tenta completar com o assunto do email se o texto for muito curto
        df["Text"] = df.apply(lambda x: x["Text"] if len(x["Text"].split()) >= 5 else helpers.clean(x["Subject"]) + " " + x["Text"], axis=1)

        # Obtém as predições
        df = self.get_predictions(df)

        # Geração de características adicionais
        try:
            df["Apolice"] = df.apply(lambda x: helpers.get_apolice(x["Subject"].strip(".") + ". " + x["Body"], x['EmailID'], self.logger), axis=1)
        except Exception as e:
            self.logger.error(f"Erro NLP Apólice: {e}")
            raise e
        try:
            df["Nome"] = df.apply(lambda x: helpers.get_names(x["Subject"].strip(".") + ". " + x["Body"], self.logger), axis=1)
        except Exception as e:
            self.logger.error(f"Erro NLP Nome: {e}")
            raise e
        try:
            df["NIF"] = df.apply(lambda x: helpers.get_nif(x["Subject"].strip(".") + ". " + x["Body"], self.logger), axis=1)
        except Exception as e:
            self.logger.error(f"Erro NLP NIF: {e}")
            raise e
        try:
            df["IDTermosExpressoes"] = df["Body"].apply(helpers.get_top_three_keywords_counts)
        except Exception as e:
            self.logger.error(f"Erro NLP ID Termos Expressões: {e}")
            raise e
        self.logger.info("FEATURES - generated.") # Informa que as características foram geradas
        df.to_excel('output_integration.xlsx') # Salva os resultados em um arquivo Excel
        self.update_database(df) # Atualiza o banco de dados com as novas informações


'''
## CONFIG variables
NUM_LABELS = 11
BASE_DIR = "/Users/feliperocha/Documents/CGI/Email Answering/"
TOKENIZER_PATH = BASE_DIR + "tokenizer"
MODEL_PATH = BASE_DIR + "model"
DATABASE = "RealVidaSeguros"
STATUS_TABLE = "QueueItem"
EMAIL_TABLE = "Emails_IPA_NLP"

COLUMN_NAMES = [
    'EmailRemetente','DataEmail', 'EmailID','Subject', 'Body', 'Anexos',
    'NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao', 'Score', 'IDTermosExpressoes',
    'DetalheMensagem', 'Mensagem', 'Estado'
]
ENGINE = db.create_engine(f'mysql+mysqlconnector://root:@localhost:3306/{DATABASE}')
CONN = ENGINE.connect()

# Load data from database
query = f"""
    SELECT {', '.join(COLUMN_NAMES)}
    FROM {EMAIL_TABLE} et
    JOIN {STATUS_TABLE} st ON et.EmailID = st.Reference
    WHERE st.Status = "NLP";
"""
df = pd.read_sql_query(query, con=ENGINE)

# Update Status
query = text(f"""
    UPDATE {STATUS_TABLE}
    SET Status = 'NLP IN PROGRESS'
    WHERE Reference IN {tuple(df['EmailID'].to_list())};
""")  # text() is needed when using sqlalchemy
CONN.execute(query)

# Clean e-mail body (remove non-portuguese sentences, non-printable character and extra spaces)
df["Text"] = df["Body"].copy()

# Identify e-mails replied / e-mail chain
df["HistoricoEmails"] = df["Body"].apply(lambda x: ((x.count("From:") > 1) and (x.count("To:") > 1)) | ((x.count("De:") > 1) and (x.count("Para:") > 1)))

# Identify empty texts (this may happen if the e-mail is not i english or if it's too short)
df["Text"] = df.apply(lambda x: x["Text"] if len(x["Text"].split()) >= 5 else helpers.clean(x["Subject"]) + " " + x["Text"], axis=1)

# Load the BERT tokenizer and model
model = BertForSequenceClassification.from_pretrained(MODEL_PATH, num_labels=NUM_LABELS)
tokenizer = BertTokenizer.from_pretrained(TOKENIZER_PATH, truncation=True, padding="max_length", max_length=128)

# Get predictions
clf = pipeline("text-classification", model=model, tokenizer=tokenizer, truncation=True, padding="max_length", max_length=128)
predictions = clf(df["Text"].to_list())
labels = [str(int(pred["label"].split('_')[1])) for pred in predictions]
scores = [str(round(float(pred["score"]), 2)) for pred in predictions]

df["IDIntencao"] = labels
df["Score"] = scores

# Aditional Features
df["Apolice"] = df.apply(lambda x: helpers.get_apolice(x["Subject"].strip(".") + ". " + x["Body"]), axis=1)
df["Nome"] = df.apply(lambda x: helpers.get_names((x["Subject"].strip(".") + ". " + x["Body"])), axis=1)
df["NIF"] = df.apply(lambda x: helpers.get_nif(x["Subject"].strip(".") + ". " + x["Body"]), axis=1)
df["IDTermosExpressoes"] = df["Body"].apply(helpers.get_top_three_keywords_counts)

# Save to database
df = df[COLUMN_NAMES]

def update_row(row):
    data = ', '.join([f"{col} = '{row[col]}'" for col in ['NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao', 'Score', 'IDTermosExpressoes']])
    query = text(f"""
        UPDATE {EMAIL_TABLE}
        SET {data}
        WHERE EmailID = '{row['EmailID']}';
    """)
    CONN.execute(query)
df.apply(update_row, axis=1)

# Update Status
query = text(f"""
    UPDATE {STATUS_TABLE}
    SET Status = 'NLP FINISHED'
    WHERE Reference IN {tuple(df['EmailID'].to_list())};
""")
CONN.execute(query)

time.sleep(3)

query = text(f"""
    UPDATE {STATUS_TABLE}
    SET Status = 'NLP FAILED'
    WHERE Status = 'NLP IN PROGRESS';
""")
CONN.execute(query)
'''