# import sqlalchemy as db
# from sqlalchemy import text
import pandas as pd
import time
from transformers import BertTokenizer, BertForSequenceClassification, pipeline

import helpers


class EmailClassifier:
    def __init__(self, base_dir: str, num_labels: int, status_table: str, email_table: str, column_names: list[str], engine, logger,db) -> None:
        self.base_dir = base_dir
        self.num_labels = num_labels
        self.status_table = status_table
        self.email_table = email_table
        self.column_names = column_names
        self.engine = engine
        self.cursor = db.cursor()
        self.logger = logger

    def get_emails(self) -> pd.DataFrame:
        # Load data from database
        query = f"""
            SELECT {', '.join(self.column_names)}
            FROM {self.email_table} et
            JOIN {self.status_table} st ON et.EmailID = st.Reference
            WHERE st.Status = 'NLP';
        """
        df = pd.read_sql_query(query, con=self.engine)
        self.logger.debug(df)
        self.logger.info("LOADED - emails data.")

        # Update Status
        for i in df["EmailID"].tolist():   
            query = f"""
                UPDATE {self.status_table}
                SET Status = 'NLP IN PROGRESS'
                WHERE Reference = '{i}';
            """  # text() is needed when using sqlalchemy
            self.cursor.execute(query)
        self.logger.info("UPDATED - NLP in progress.")
        return df
    
    def get_predictions(self, df: pd.DataFrame) -> pd.DataFrame:
        # Load the BERT tokenizer and model
        model = BertForSequenceClassification.from_pretrained(self.base_dir.strip('/')+'/model', num_labels=self.num_labels)
        tokenizer = BertTokenizer.from_pretrained(self.base_dir.strip('/')+'/tokenizer', truncation=True, padding="max_length", max_length=128)
        self.logger.info("LOADED - model and tokenizer.")

        # Get predictions
        clf = pipeline("text-classification", model=model, tokenizer=tokenizer, truncation=True, padding="max_length", max_length=128)
        predictions = clf(df["Text"].to_list())
        labels = [str(int(pred["label"].split('_')[1])) for pred in predictions]
        scores = [str(round(float(pred["score"]), 2)) for pred in predictions]
        self.logger.info("PREDICTIONS - generated.")

        df["IDIntencao"] = labels
        df["Score"] = scores
        self.logger.info("PREDICTIONS - inserted on dataframe.")
        return df
    
    def update_row(self, row):
            data = ', '.join([f"{col} = '{row[col]}'" for col in
                              ['NIF', 'Apolice', 'Nome', 'HistoricoEmails', 'IDIntencao', 'Score', 'IDTermosExpressoes']])
            query = f"""
                UPDATE {self.email_table}
                SET {data}
                WHERE EmailID = '{row['EmailID']}';
            """
            self.cursor.execute(query)

    def update_database(self, df) -> None:
        # Save to database
        df = df[self.column_names]
        df.apply(self.update_row, axis=1)

        # Update Status
        for i in df['EmailID'].tolist():
            query = f"""
                UPDATE {self.status_table}
                SET Status = 'NLP FINISHED'
                WHERE Reference = '{i}';
            """
            self.cursor.execute(query)

        time.sleep(3)

        query = f"""
            UPDATE {self.status_table}
            SET Status = 'NLP FAILED'
            WHERE Status = 'NLP IN PROGRESS';
        """
        self.cursor.execute(query)
        self.logger.info("UPDATED - NLP finished.")

    def run(self):
        # Get data
        df = self.get_emails()

        # Clean e-mail body (remove non-portuguese sentences, non-printable character and extra spaces)
        df["Text"] = df["Body"].copy()

        # Identify e-mails replied / e-mail chain
        df["HistoricoEmails"] = df["Body"].apply(lambda x: ((x.count("From:") > 1) and (x.count("To:") > 1)) | ((x.count("De:") > 1) and (x.count("Para:") > 1)))

        # Identify empty texts (this may happen if the e-mail is not i english or if it's too short)
        df["Text"] = df.apply(lambda x: x["Text"] if len(x["Text"].split()) >= 5 else helpers.clean(x["Subject"]) + " " + x["Text"], axis=1)

        # Get predictions
        df = self.get_predictions(df)

        # Aditional Features
        df["Apolice"] = df.apply(lambda x: helpers.get_apolice(x["Subject"].strip(".") + ". " + x["Body"]), axis=1)
        df["Nome"] = df.apply(lambda x: helpers.get_names((x["Subject"].strip(".") + ". " + x["Body"])), axis=1)
        df["NIF"] = df.apply(lambda x: helpers.get_nif(x["Subject"].strip(".") + ". " + x["Body"]), axis=1)
        df["IDTermosExpressoes"] = df["Body"].apply(helpers.get_top_three_keywords_counts)
        self.logger.info("FEATURES - generated.")

        self.update_database(df)


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