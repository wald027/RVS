import extract_msg
import pandas as pd
import os
import re
import spacy
import langid
from datetime import datetime
import logging
import sklearn
print(sklearn.__version__)

from sklearn.model_selection import train_test_split
from imblearn.over_sampling import RandomOverSampler
from sklearn.metrics import classification_report, confusion_matrix
import numpy as np

from transformers import BertTokenizer, BertForSequenceClassification, EarlyStoppingCallback, Trainer, TrainingArguments, DataCollatorWithPadding, pipeline
from datasets import Dataset, DatasetDict
import evaluate
import torch
from customScripts.customLogging import setup_logging
from customScripts.databaseSQLExpress import ConnectToBD
from customScripts.readConfig import *


#import sys
#sys.path.append('/content/drive/MyDrive/CGI/Email Answering')

import ModelNLP.helpers as helpers
from MailboxTraining import extract_emails_from_folder, dataframe

# Variables
# Define se o conjunto de dados é rotulado ou não
LABELLED = True
# Define se o modelo será treinado ou não
TRAIN = False
# Define o número de rótulos (categorias) para a classificação
NUM_LABELS = 11
# Caminho do modelo BERT pré-treinado em português, cased (sensível a maiúsculas/minúsculas)
MODEL_PATH = "neuralmind/bert-base-portuguese-cased"
# Diretório base para armazenar modelos e dados
BASE_DIR = 'realvidaseguros/'
# Caminho para salvar o modelo treinado
TRAINED_MODEL_PATH = BASE_DIR + "model_2"
TRAINED_MODEL_PATH_2 = BASE_DIR + "model_3"
# Diretório onde estão armazenados os arquivos de emails
EMAILS_DIRECTORY = BASE_DIR + "Dados Classificados/" 
# Nome do servidor SQL
server = 'PT-L164962\SQLEXPRESS'
# Nome da base de dados
DATABASE = "RealVidaSeguros"
# Conecta ao banco de dados usando as credenciais fornecidas
db = ConnectToBD(server,DATABASE)
# Configura o sistema de logs para registrar eventos e erros
setup_logging(db,'LOGS')
# Cria um objeto logger para registrar eventos durante a execução do código
logger = logging.getLogger(__name__)

'''
#Define uma função para importar emails de arquivos .msg para um DataFrame do pandas
def import_msg_to_df(directory, extension='.msg', labelled=False):
    """
    Import .msg files from an organized directory into a pandas DataFrame.
    """
    # Lista para armazenar os dados extraídos dos emails
    my_list = []
    # Dicionário para mapear rótulos (labels) para suas descrições
    label_map = {}
    # Se os dados forem rotulados
    if labelled:
        # Para cada amostra no diretório
        for dir_amostra in os.listdir(directory):
            # Define o caminho do diretório da amostra
            dir = directory+dir_amostra+'/'
            # Se o caminho é um diretório
            if os.path.isdir(dir):
                print("Hello")
                amostragem = dir_amostra
                # Para cada diretório de rótulo na amostra
                for dir_label in os.listdir(dir):
                    # Define o caminho completo para o rótulo
                    full_dir = dir+dir_label+'/'
                    # Se o caminho é um diretório
                    print(full_dir)
                    if os.path.isdir(full_dir):
                        print("Hello2")
                        # Extrai o número do rótulo do nome do diretório
                        label = dir_label.split(' - ')[0]#.strip("Label "))
                        # Extrai a descrição do rótulo do nome do diretório
                        template = dir_label.split(' - ')[1]
                        # Adiciona o rótulo e sua descrição ao mapa de rótulos
                        label_map[label] = template
                        # Para cada arquivo no diretório de rótulo
                        for file in os.listdir(full_dir):
                            # Se o arquivo tem a extensão correta
                            if file.endswith(extension):
                                # Extrai os dados do arquivo .msg
                                msg = extract_msg.Message(full_dir+file)
                                # Adiciona os dados do email à lista
                                my_list.append([msg.filename, msg.sender, msg.messageId, msg.to, msg.date, msg.subject, msg.body, label, template, amostragem])
    # Se os dados não forem rotulados
    else:
        # Para cada amostra no diretório
        for dir_amostra in os.listdir(directory):
             # Define o caminho completo para a amostra
            full_dir = directory+dir_amostra+'/'
             # Se o caminho é um diretório
            if os.path.isdir(full_dir):
                amostragem = dir_amostra
                # Para cada arquivo no diretório de amostra
                for file in os.listdir(full_dir):
                    # Se o arquivo tem a extensão correta
                    if file.endswith(extension):
                        # Extrai os dados do arquivo .msg
                        msg = extract_msg.Message(full_dir+file)
                        # Adiciona os dados do email à lista
                        my_list.append([msg.filename, msg.sender, msg.messageId, msg.to, msg.date, msg.subject, msg.body, None, None, amostragem])
    # Cria um DataFrame do pandas com os dados dos emails
    df = pd.DataFrame(my_list, columns=['File Name', 'From', 'Email ID', 'To', 'Date', 'Subject', 'Body', 'Label', 'Label Template', 'Amostragem'])
    #print(df)
    # Retorna o DataFrame e o mapa de rótulos
    return df, label_map
'''

#df_original = extract_emails_from_folder(EMAILS_DIRECTORY, labelled=LABELLED)
dictConfig = readConfig()
df_original = dataframe(logger,dictConfig)
print(df_original)
# Chama a função para importar os emails e obtém o DataFrame e o mapa de rótulos
#df_original, label_map = import_msg_to_df(EMAILS_DIRECTORY, labelled=LABELLED)
#print(df_original)
#Id das Intenções
label_map = {
    0: 'Reforços apólices financeiras',
    1: 'Atualização dados pessoais A',
    2: 'Atualização dados pessoais B',
    3: 'Atualização de capital da apólice',
    4: 'Pedido de informação da apólice',
    5: 'Pedido de resgate de apólice financeira',
    6: 'Acesso à Área Reservada de Clientes (MyRealVida)',
    7: 'Alteração de IBAN de débito',
    8: 'Pedido de anulação de apólices Universo',
    9: 'Participação de sinistros Acidentes Pessoais',
    10: 'Participação de sinistros Vida Risco'
}
print("Sizes: ", df_original.size)
# Adiciona uma coluna ao DataFrame indicando se o email contém histórico de mensagens anteriores
df_original["Histórico Emails"] = df_original.apply(lambda x: helpers.get_historico(x["Body"], x['Email ID'], logger), axis=1)

# Cria uma cópia do corpo do email na coluna "Text" : Limpeza do dados  -> retira palavras que não são em lingua portuguesa, caracteres não imprimivéis e extra espaços. 
df_original["Text"] = df_original["Body"].copy() #.apply(helpers.clean)

# Se o texto do email for muito curto, adiciona o assunto ao texto para complementar
# Identificar textos vazios -> Emails não vazios = not em Inglês ou muito pequenos) 
df_original["Text"] = df_original.apply(lambda x: x["Text"] if len(x["Text"].split()) >= 5 else helpers.clean(x["Subject"]) + " " + x["Text"], axis=1)

#print(df_original.apply(lambda x: x["Text"] if len(x["Text"].split()) >= 5 else helpers.clean(x["Subject"]) + " " + x["Text"], axis=1).head())

# Adiciona uma coluna indicando se o email é curto (menos de 4 palavras)
df_original["Email Curto"] = df_original["Text"].apply(lambda x: len(x.split()) <= 3)

# Adiciona uma coluna indicando se o email é duplicado, com base no texto
df_original['Duplicado'] = df_original.duplicated('Text')

print("Preprocessed dataframe: ", df_original)
# Cria uma amostra aleatória do DataFrame e embaralha os dados
#df = df_original.sample(frac=1, random_state=0)
#df_original = df_original.sample(frac=1, random_state=0)
print("After sampling:", df_original.shape)
# Filtra emails para remover aqueles com histórico de mensagens anteriores
#df = df[df["Histórico Emails"] == False]
#df_original = df_original[df_original["Histórico Emails"] == False]
print("After filtering 'Histórico Emails':", df_original.shape)
# Filtra emails duplicados
#df = df[df['Duplicado'] == False]
#df_original = df_original[df_original['Duplicado'] == False]
print("After filtering 'Duplicado':", df_original.shape)
# Filtra emails muito curtos
#df = df[df["Email Curto"] == False]
#df_original = df_original[df_original["Email Curto"] == False]
print("After filtering 'Email Curto':", df_original.shape)

# Se os dados são rotulados, divide em conjuntos de treino e teste

if LABELLED:
    # Separa os recursos (X) e os rótulos (y)
    print("Dataframe here")
    #print(df_original)
    X = df_original.drop("Label", axis=1)
    print(X)
    y = df_original["Label"]

    
    # Divide os dados em conjuntos de treino e teste, mantendo a proporção das classes
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)

    train = pd.concat([pd.DataFrame(X_train), pd.Series(y_train)], axis=1)
    train.columns = list(X.columns) + ["Label"]

    test = pd.concat([pd.DataFrame(X_test), pd.Series(y_test)], axis=1)
    test.columns = list(X.columns) + ["Label"]

# TREINO MODELO --> Se o treinamento do modelo estiver ativado

if LABELLED:
    # Preparar o dataset
    train["Mode"] = "TRAIN" # Marca os dados de treino com a label "TRAIN"
    test["Mode"] = "TEST" # Marca os dados de teste com a label "TEST"
    # Concatenar o Treino e o Teste
    df = pd.concat([train, test]) # Une os datasets de treino e teste em um único DataFrame

# Carrega o modelo e o tokenizador BERT para classificação de sequência
model = BertForSequenceClassification.from_pretrained(MODEL_PATH, num_labels=NUM_LABELS)
tokenizer = BertTokenizer.from_pretrained(MODEL_PATH, truncation=True, padding="max_length", max_length=128)

#device = torch.device("cuda")
#model.to(device)

# PREDIÇÕES

# Criação do pipeline para classificação de texto usando o modelo e tokenizador BERT 
clf = pipeline("text-classification", model=model, tokenizer=tokenizer, truncation=True, padding="max_length", max_length=128)
# Gera as predições para a lista de textos
predictions = clf(df["Text"].to_list())
# Para cada predição, extrai o rótulo previsto
labels = [int(pred["label"].split('_')[1]) for pred in predictions]
# Para cada predição, extrai a pontuação de confiança
scores = [round(float(pred["score"]), 2) for pred in predictions]

# Adicionar as predições ao DataFrame 
df["Prediction"] = labels # Rótulos previstos
df["Prediction Template"] = [label_map[label] for label in labels] # Descrição dos rótulos previstos
df["Score"] = scores # Pontuações de confiança das predições

# Adiciona informações sobre e-mails não usados no treinamento-teste do modelo
df = pd.merge(df, df_original, on=[col for col in df_original.columns if col in df.columns], how="right").fillna("")

# Recursos adicionais
df["Apólice"] = df.apply(lambda x: helpers.get_apolice((x["Subject"].strip(".") + ". " + x["Body"]),x['Email ID'], logger), axis=1) # Extrai número de apólice
df["Nome"] = df.apply(lambda x: helpers.get_names((x["Subject"].strip(".") + ". " + x["Body"]), logger), axis=1) # Extrai nomes de pessoas
df["NIF"] = df.apply(lambda x: helpers.get_nif(x["Subject"].strip(".") + ". " + x["Body"], logger), axis=1) # Extrai número de identificação fiscal (NIF)
df["ID Termos Expressões"] = df["Body"].apply(helpers.get_top_three_keywords_counts) # Extrai as três palavras-chave mais frequentes

# Salvar resultados
df["Date"] = df["Date"].astype(str) # Converte a coluna de data para string
df = df.drop(columns=["File Name", "To"]) # Remove colunas desnecessárias
df = df.rename(columns={"From": "Email Remetente", "Date": "Data Email", }) # Renomeia colunas para nomes mais claros

# Validação rápida
if LABELLED:
    # Seleciona colunas para o arquivo de saída e salva os resultados em um arquivo Excel
    df = df[["Email Remetente", "Data Email", "Email ID", "Subject", "Body", "NIF", "Apólice", "Nome", "Amostragem", "Mode", "Histórico Emails", "Duplicado", "Email Curto", "Label", "Prediction", "Label Template", "Prediction Template", "Score", "ID Termos Expressões"]]
    #df.to_excel(BASE_DIR+f"resultados_{str(datetime.now()).split()[0].replace('-', '')}.xlsx", index=False)
    df.to_excel("resultados_20240812_2.xlsx")
    # Filtra os dados de teste e imprime o relatório de classificação
    df_val = df[(df["Mode"] == "TEST")]
    print(classification_report(df_val["Label"].astype(int), df_val["Prediction"].astype(int), target_names=[name for i, name in sorted(label_map.items())]))
    print(confusion_matrix(df_val["Label"].astype(int), df_val["Prediction"].astype(int)))

else:
    # Para dados não rotulados, marca como "NEW" e salva em um arquivo Excel
    df["Mode"] = "NEW"
    df = df[["Email Remetente", "Data Email", "Email ID", "Subject", "Body", "NIF", "Apólice", "Nome", "Amostragem", "Mode", "Histórico Emails", "Duplicado", "Email Curto", "Label", "Prediction", "Label Template", "Prediction Template", "Score", "ID Termos Expressões"]]
    df.to_excel(BASE_DIR+f"resultados_não_classificados_{str(datetime.now()).split()[0].replace('-', '')}.xlsx", index=False)
    # Imprime o relatório de classificação detalhado para os dados de teste

if LABELLED:
    print(classification_report(df_val["Label"].astype(int), df_val["Prediction"].astype(int), target_names=[
        'Reforços apólices financeiras',
        'Atualização dados pessoais A',
        'Atualização dados pessoais B',
        'Atualização de capital da apólice',
        'Pedido de informação da apólice',
        'Pedido de resgate de apólice financeira',
        'Acesso à Área Reservada de Clientes (MyRealVida)',
        'Alteração de IBAN de débito',
        'Pedido de anulação de apólices Universo',
        'Participação de sinistros Acidentes Pessoais',
        'Participação de sinistros Vida Risco']
    ))


#df.groupby("Mode").count()

#df[["Histórico Emails", "Duplicado", "Email Curto"]].sum()