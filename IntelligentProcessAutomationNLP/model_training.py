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
import langid
import spacy
from customScripts.customLogging import setup_logging
from customScripts.databaseSQLExpress import ConnectToBD
from customScripts.readConfig import *
from MailboxTraining import extract_emails_from_folder, dataframe, dictConfig

#import sys
#sys.path.append('/content/drive/MyDrive/CGI/Email Answering')

#import ModelNLP.helpers as helpers
nlp = spacy.load("pt_core_news_sm")
# Variables
# Define se o conjunto de dados é rotulado ou não
LABELLED = True
# Define se o modelo será treinado ou não
TRAIN = True
# Define o número de rótulos (categorias) para a classificação
NUM_LABELS = 11
# Caminho do modelo BERT pré-treinado em português, cased (sensível a maiúsculas/minúsculas)
MODEL_PATH = "neuralmind/bert-base-portuguese-cased"
# Diretório base para armazenar modelos e dados
BASE_DIR = 'realvidaseguros/'
# Caminho para salvar o modelo treinado
#TRAINED_MODEL_PATH = BASE_DIR + "model_2_0402"
TRAINED_MODEL_PATH = BASE_DIR + "model_3"
# Diretório onde estão armazenados os arquivos de emails
EMAILS_DIRECTORY = BASE_DIR + "Dados Classificados/" 
# Nome do servidor SQL
server = 'PT-L163255\SQLEXPRESS01'
# Nome da base de dados
DATABASE = "RealVidaSeguros"
# Conecta ao banco de dados usando as credenciais fornecidas
db = ConnectToBD(server,DATABASE)
# Configura o sistema de logs para registrar eventos e erros
setup_logging(db,'LOGS_TREINO')
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


filepath = r'C:\Users\nayara.rodrigues\Documents\RVS-main\RVS_emails_20240823 (1) - Copy.xlsx'
#filepath = "c:\\Users\\nayara.rodrigues\\Documents\\RVS-main\\IntelligentProcessAutomationNLP\\ModelNLP\\helpers.py"
#filepath = "C:\\Users\\nayara.rodrigues\\Documents\\RVS-main\\Modelo de Dados NLP_emails15.xlsx"
#df_original = pd.read_excel(filepath)
#print(f"dataframe {df_original}")

#dictConfig = readConfig(r'C:\Users\nayara.rodrigues\Documents\RVS-main\Config_TreinoModelo.xlsx')
#df_original = pd.read_excel(filepath)
#df_original = dataframe(logger,dictConfig)
#print(f"here: {df_original}")

#df_original = dataframe(logger, dictConfig)
df_original = pd.read_excel(filepath)
print(df_original)
#print(f"Data frame teste here: {df_original['Name Folders'].iloc[0][0]}")


label_map = {
    0: 'Reforços apólices financeiras',
    1: 'Atualização dados pessoais A',
    2: 'Atualização dados pessoais B',
    3: 'Atualização de capital da apólice',
    
    4: 'Pedido de resgate de apólice financeira',
    5: 'Acesso à Área Reservada de Clientes (MyRealVida)',
    6: 'Alteração de IBAN de débito',
    7: 'Pedido de anulação de apólices Universo',
    8: 'Participação de sinistros Acidentes Pessoais',
    9: 'Participação de sinistros Vida Risco'
}

label_map = {
    0: 'Reforços apólices financeiras',
    1: 'Atualização dados pessoais A',
    2: 'Atualização dados pessoais B',
    3: 'Atualização de capital da apólice',
    #4: 'Pedido de informação da apólice',
    5: 'Pedido de resgate de apólice financeira',
    6: 'Acesso à Área Reservada de Clientes (MyRealVida)',
    7: 'Alteração de IBAN de débito',
    8: 'Pedido de anulação de apólices Universo',
    9: 'Participação de sinistros Acidentes Pessoais',
    10: 'Participação de sinistros Vida Risco'
}

convert_map = {
    0: 0,
    1: 1,
    2: 2,
    3: 3,
    5: 4,
    6: 5,
    7: 6,
    8: 7,
    9: 8,
    10: 9
}

convert_back_map = {
    0: 0,
    1: 1,
    2: 2,
    3: 3,
    4: 5,
    5: 6,
    6: 7,
    7: 8,
    8: 9,
    9: 10
    
}

#df_original["Label Template"] = df_original["Label"].map(label_map)

# PREPROCESSING

def clean(text):
    """
    Limpeza dos dados
    """
    # remove frases que não são em português
    for sentence in text.split("\n"):
        for sent in sentence.split("."):
            if langid.classify(sent)[0] != "pt" and len(sent) > 1:
                text = text.replace(sent, " ")

    # remove caracteres que não são aceitos.
    text = text.replace("\n", " ").replace("\r", " ")

    # remove espaços duplos
    text = re.sub(' +', ' ', text)
    text = re.sub(' \.', '.', text)
    text = re.sub('\.+', '.', text)

    return text.lstrip(". ")

#Validação da apólice
def validar_apolice2(modalidade, apolice):

  current_date = datetime.now()
  current_year = current_date.year

  # Todas são válidas, onde a modalidade for acima de 32/XXXXXX (Não existem meses com 32 dias, logo, não é data);
  if int(modalidade) > 32 and not (1900 <= int(apolice) <= current_year):
      return f"{modalidade}/{apolice}"

  # Todas são válidas, onde a apólice tiver mais de 4 dígitos ou 3 dígitos. Ex: 01/053501 ou 25/21979 ou 31/535.
  if len(apolice) == 3 or len(apolice) > 4:
      return f"{modalidade}/{apolice}"

  # regra apenas quando a apólice tiver 1 ou 2 dígitos.
  # Todas são válidas, onde a modalidade for de 04 a 31 e a apólice NÃO for um número entre 01 (Com ou sem zero a esquerda) e 12.
  # NOTA 1: Apólice entre 01 e 12, pois isso representa o número de meses do ano, ou seja, de 13 em diante já consideramos válida,
  # salvo à indicação de que podemos considerar os dois últimos dígitos do ano. Se isso acontecer, temos de formatar esta regra;
  if len(apolice) == 1 or len(apolice) == 2:
    if (4 <= int(modalidade) > 32) and (int(apolice) > 13):
      return f"{modalidade}/{apolice}"

  # Aplicar essa regra apenas quando a apólice tiver 4 dígitos. NÃO são válidas, onde a modalidade for de 04 a 12 e a apólice
  # for um número entre 1980 e o ano atual, no caso 2024. NOTA 1: Modalidade entre 04 e 12, pois isso representa o número de
  # meses do ano em que existem modalidades na lista de produtos da RVS.
  if len(apolice) == 4:
     if (4 >= int(modalidade) >= 12) and 1900 <= int(apolice) <= current_year:
         return f"{modalidade}/{apolice}"
  return None

#print(df_original)


# Limpeza da apólice
def cleaner(validar):
    # se o valor não for nulo
    if validar != None:
        # split do valor
        split_value = validar.split("-") if "-" in validar else (validar.split("/") if "/" in validar else (validar.split(".") if "." in validar else [validar]))
        # se a modalidade for menor ou igual e a apólice for menor ou igual a 6
        if (len(split_value[0]) <= 2) and (3 < len(split_value[1]) <=6):
            # retorna apólice
            return f"{split_value[0]}/{split_value[1]}"
        else:
            return None
    else:
        return None

# Extrair as apólices do email
def get_apolice(text) -> str:
    # Logger para a Base de Dados
    # se o texto for string
    if isinstance(text, str):
        # retira os espaços e as quebras de linhas
        without_spaces = text.replace(" ", "").replace("\n", "")
        # encontrar os números no texto: modalidade entre 1 a 2 digitos, apolice entre 1 a 6
        number = re.findall(r'\d{1,2}/\d{1,6}|\d+\.\d+|\d{1,2}-\d{1,6}|\d+/\d+', without_spaces)
        apolice_valida = []
        # Processar e validar cada número encontrado seguindo as regras do regex
        for i in number:
            validar = ""
            # numero com '/'
            if "/" in i:
                #split pela '/'
                temp = i.split("/")
                #modalidade e apolice
                modalidade, apolice = temp[0], temp[1]
                #validar modalidade/apolice
                validar = validar_apolice2(modalidade, apolice)
            # numero com '-'
            elif "-" in i:
                #split pela '-'
                temp = i.split("-")
                #modalidade e apolice
                modalidade, apolice = temp[0], temp[1]
                #validar modalidade/apolice
                validar = validar_apolice2(modalidade, apolice)
            # numero com '.'
            elif "." in i:
                #split pela '.'
                temp = i.split(".")
                #modalidade e apolice
                modalidade, apolice = temp[0], temp[1]
                #validar modalidade/apolice
                validar = validar_apolice2(modalidade, apolice)
            # se a validação da apólice não retorna None
            if validar is not None:
                #Limpeza da apólice
                validar = cleaner(validar)
                apolice_valida.append(validar)
        #validar se não há valores em None
        apolice_valida = [value for value in apolice_valida if value is not None]
        #havendo valores na lista de apolices, retira-se os duplicados.
        if len(apolice_valida) > 0:
            # retorna uma string com os valores separados por "|"
            return "|".join(set(apolice_valida))
    #logger
    return " "

#Validação do NIF
def validar_nif(nif):
    # str nif contendo '-' ou '/'
    if '-' in nif  or '/' in nif:
        return False

    # comprimento da string nif superior ou inferior a 9
    if len(nif) != 9:
        return False

    # nif iniciado com 0
    if int(nif[0]) == 0:
        return False

    # calculo nif
    s = 9 * int(nif[0]) + 8 * int(nif[1]) + 7 * int(nif[2]) + 6 * int(nif[3]) + 5 * int(nif[4]) + 4 * int(nif[5]) + 3 * int(nif[6]) + 2 * int(nif[7]) + 1 * int(nif[8])

    # resto da divisão
    resto = s % 11

    # se o resto for 1
    if resto == 1:
        # retorna na ultima posição o valor 0
        return nif[8] == '0'

    # se o resto for 0
    elif resto == 0:
        return True
    else:
        return False

# Extrai os telefones para não se confundir com os NIFs
def is_telephone(text):

   pattern = r'\(\+351\)\s*(\d+)|\+351\s*(\d+)|Tel[:.]?\s*(\d+)|Cel[:.]?\s*(\d+)|Telefone[:.]?\s*(\d+)|Telemóvel[:.]?\s*(\d+)|Apólices|apolice|Apolice|apólice'

   res = []

   # Para cada email
   for i in text:
      # se há números de telefone
      matches = re.findall(pattern, i)
      # se existe
      if matches:
         # criar uma lista
         res.append(matches)

   value_return = []
   for i in res:
      for j in i[0]:  # Todo: rever it
         if j:
            value_return.append(j)
   return value_return


# Encontrar os números com até 9 digitos
def find_numbers(text):
    numbers = []

    # Primeiro padrão
    pattern = r'\b\d{9}\b'
    # encontrar o padrão no texto
    nine_digit_numbers = re.findall(pattern, text)

    # se encontrar valores com 9 digitos
    if nine_digit_numbers:
        # loop pela lista de números
        for number in nine_digit_numbers:
            # se houver algum valor
            if number:
                # limpar números onde houver '/' ou '-'
                cleaned_number = re.sub(r'[\s\-/]', '', number)
                # criar uma lista de números
                numbers.append(cleaned_number)
        return numbers
    else:
        # Segundo padrão
        pattern = r'^\d+(\s\d+|\-\d+|/\d+)*$'
        # encontrar o segundo padrão
        nine_digit_numbers = re.findall(pattern, text)
        # loop pela lista de números
        for number in nine_digit_numbers:
            # se houver algum valor
            if number:
                # limpar números onde houver '/' ou '-'
                cleaned_number = re.sub(r'[\s\-/]', '', number)
                # criar uma lista
                numbers.append(cleaned_number)
        # se a lisata de números tiver algum valor
        if len(numbers) > 0:
            return numbers
    return None


# Extrair os NIFs
def get_nif(text):
    # logger
    #logger.info("A tratar NIF")
    nif_values = []
    numbers = find_numbers(text)

    # se não for null
    if numbers is not None:
        # loop pela lista
        for j in numbers:
            # se o elemento tiver comprimento 9
            if len(str(j)) == 9:
                # validar o nif
                #logger.info("Validar nif")
                value = validar_nif(j)
                # não for vazia
                if value:
                    #logger.info("NIF valido")
                    # criar uma lista de valores
                    nif_values.append(str(j))
        # não for vazia
        if nif_values:
            # loop pela lista de nifs validos
            for i in nif_values:
                # se as 2 primeiras posições contiver algum elemento da lista
                if i[:2] in ['92', '93', '94', '95', '96', '97']:
                    # remover valores
                    nif_values.remove(i)
            # logger
            #logger.info("Fim do tratamento do NIF")
            # retornar uma string de valores
            return '|'.join(set(nif_values))
        else:
            # logger
            #logger.info("Fim do tratamento do NIF")
            return " "
    else:
        # logger
        #logger.info("Fim do tratamento do NIF")
        return " "


# Extrair os Nomes
def get_names(text):
    # logger
    #logger.info("A tratar Nome")
    # valores presentes nos emails que devem ser excluidos
    noise = ["pt50", "diretor", "subdiretor", "professor", "Suplementar"]  # ["agradecia", "obrigada", "pt", "rua", "lisboa", "nif", "cumprimentos"]
    names = []
    # chamar o modelo
    docs = nlp(" ".join([t.capitalize() for t in str(text).split()]))
    # encontrar as entities no corpo do texto
    entities = docs.ents
    # percorrer as entities
    for entity in entities:
        # em cada entitite selecionar o .label onde for 'PER' e onde o comprimento do .text for maior ou igual à 4
        if entity.label_ == "PER" and len(entity.text) >= 4:
            # aplicar o .text no modelo
            doc = nlp(entity.text)
            names_partial = []
            # loop pelos
            for token in doc:
                if token.ent_type_ == "PER" and 4 <= len(token.text) <= 12 and token.pos_ == "PROPN" and token.text.lower() not in noise:
                    txt = re.sub(" +", " ", token.text.replace('\n', " ")).strip()
                    names_partial.append(txt)
            names_concat = " ".join(names_partial).strip()
            if len(names_concat) > 0 and names_concat not in names:
                names.append(names_concat)
    #logger.info("Fim do tratamento do Nome")
    return "|".join(list(set(names)))

def get_historico(i):
    res = []
    #print(text)
    #for i in text:#.iterrows():
     #   print(i)
    # se contiver algum dos termos abaixo
    if  "De:" in i and "Enviado:" in i and "Para:" in i and "Assunto:" in i :
        return True
    elif  "De:" in i and "Enviado:" in i and "Para:" in i and "Cc:" in i and "Assunto:" in i:
        return True
    elif "De:" in i and "Data:" in i and "Assunto:" in i and "Para:" in i and "Cc:" in i:
        return True
    elif "De:" in i and "Enviada:" in i and "Para:" in i and "Assunto:" in i:
        return True
    elif "De:" in i and "Enviado:" in i and "Cc:" in i and "Assunto:" in i:
        return True
    elif "De:" in i and "Assunto:" in i and "Data:" in i and "Para:" in i:
        return True
    elif "De:" in i and "Date:" in i and "Subject:" in i and "To:" in i:
        return True
    elif "De:" in i and "Data:" in i and "Para:" in i and "Assunto:" in i:
        return True
    elif "De:" in i and "Data:" in i and "Assunto:" in i and "Para:" in i:
        return True
    elif "Data:" in i and "De:" in i and "Assunto:" in i and "Cc:" in i and "Para:" in i :
        return True
    elif "From:" in i and "Sent:" in i and "To:" in i and "Cc:" in i and "Subject:" in i:
        return True
    elif "From:" in i and "Sent on:" in i and "CC:" in i and "Subject:" in i :
        return True
    elif "From:" in i and "Sent:" in i and "To:" in i and "Subject:" in i :
        return True
    elif "From:" in i and "Data:" in i and "Assunto:" in i and "Para:" in i:
        return True        
    elif "Clientes Real Vida <info.clientes@realvidaseguros.pt" in i:
        return True
    elif "Real Vida Seguros <documentoseletronicos@realvidaseguros.pt" in i :
        return True        
    elif "Real Vida Seguros <noreply@realvidaseguros.pt>" in i:  
        return True
    elif "Real Vida Seguros <documentoseletronicos@realvidaseguros.pt> " in i:
        return True
    elif '----- Mensagem de Real Vida Seguros <noreply@realvidaseguros.pt <mailto:noreply@realvidaseguros.pt> > ---------' in i:
        return True
    elif "---------- Forwarded message ---------" in i:
        return True
    elif "---------- Mensagem encaminhada ---------" in i:
        return True
    elif "----- Mensagem de Real Vida Seguros <digital@cert.realvidaseguros.pt> ---------" in i:
        return True
    elif "-------- Mensagem original --------" in i:
        return True
    else:        
        return False
    

def get_top_three_keywords_counts(text):
    """
    Counts the number of times each keyword appears in the text.
    Returns the three templates with the highest keywords count.
    """
    # Numeração de cada template
    template_to_standard = {
        "Reforços Apólices Financeiras": 0,
        "Atualização Dados Pessoais A":  1,
        "Atualização Dados Pessoais B":  2,
        "Atualização capital apólice": 3,
        "Pedido de informação da apólice": 4,
        "Pedido de resgate de apólice financeira": 5,
        "Acesso Área Reservada": 6,
        "Alteração IBAN": 7,
        "Anulação de Apólices": 8,
        "Participação de sinistro - Acidente": 9,
        "Participação de sinistro - Vida": 10,
    }
    # termos para cada template
    template_keywords = {
            "Reforços Apólices Financeiras": ["reforço", "entrega suplementar", "entrega extraordinária", "financeiro", "ppr"],
            "Atualização Dados Pessoais A": ["morada", "telefone", "telemóvel", "e-mail"],
            "Atualização Dados Pessoais B": ["data de nascimento", "nome", "nif", "sexo"],
            "Atualização capital apólice": ["atualização de capital", "capital em dívida", "banco"],
            "Pedido de informação da apólice": ["informação", "documentos", "dados"],
            "Pedido de resgate de apólice financeira": ["resgate", "resgate parcial", "resgate total", "levantar valor", "financeir[ao]", "ppr"],
            "Acesso Área Reservada": ["myrealvida", ".rea reservada", "código de acesso", "password", "recupera(r|ção)"],
            "Alteração IBAN": ["iban", "nib", "altera.+número de conta"],
            "Anulação de Apólices": ["universo", "anulação", "cancelamento", "seguro saúde", "seguro saúde star", "seguro dentista"],
            "Participação de sinistro - Acidente": ["acidente", "incapacidade", "invalidez", "falecimento"],
            "Participação de sinistro - Vida": ["doença", "morte", "óbito", "falecimento"]
    }
    # inicializa um dicionário chamado counts, onde cada chave é um modelo do dicionário template_keywords, e cada valor é definido como 0.
    counts = {k:0 for k in template_keywords.keys()}
    #converte a variável text para uma string, caso ainda não seja.
    text = str(text)
    # bloco itera sobre cada modelo e sua lista associada de palavras-chave no dicionário template_keywords.
    # Para cada palavra-chave na lista, ele verifica se a palavra-chave (independentemente de maiúsculas ou minúsculas)
    # está presente no text usando re.findall(). Se a palavra-chave for encontrada, incrementa a contagem para aquele modelo no dicionário counts.
    for template, keywords in template_keywords.items():
        for keyword in keywords:
            if re.findall(keyword, text.lower()):
                counts[template] += 1
    # classifica os itens do dicionário counts (que são pares de modelo-contagem) pelo valor da contagem em ordem decrescente,
    # usando uma função lambda como chave de classificação.
    top_three = sorted(counts.items(), key=lambda item: item[1], reverse=True)[:3]
    # mapeia os três principais modelos para seus rótulos padrão usando o dicionário template_to_standard.
    top_three_labels = [template_to_standard[label] for label, count in top_three]
    # retorna a lista dos três principais rótulos padrão correspondentes aos modelos com o maior número de correspondências de palavras-chave
    return top_three_labels

#df_original = df_original[["Subject", 'Body', 'Label']]


#df_original["Label"].value_counts()

#df_original['Label']

#df_original.isna().sum()

#df_original.head()


#df_original = dataframe(logger, dictConfig)
#df_original["Label"] = df_original["Label"].map(convert_map)
#print(df_original["Label"].value_counts())
#display(df_original[df_original["Label"].isna()])

#df_original["Text"] = df_original.apply(lambda x: x["Text"] if len(str(x["Text"]).split()) >= 5 else helpers.clean(x["Subject"]) + " " + str(x["Text"]), axis=1)
#historico = get_historico(df_original['Body'])
# Adiciona uma coluna ao DataFrame indicando se o email contém histórico de mensagens anteriores
df_original["NIF"] = df_original.apply(lambda x: get_nif(str(x["Subject"]) + " " + str(x['Body'])), axis=1)
df_original["Apolice"] = df_original.apply(lambda x: get_apolice(str(x["Subject"]) + " " + str(x['Body'])), axis=1)
df_original["Nome"] = df_original.apply(lambda x: get_names(str(x["Subject"]) + " " + str(x['Body'])), axis=1)
df_original['Mode'] = ''
#historico = get_historico(df_original['Body'])  #df_original.apply(lambda x: get_historico(str(x['Body'])), axis=1)
#print(historico)
#df_original["Histórico Emails"] = historico          #df_original.apply(lambda x: get_historico(str(x["Subject"]) + " " + str(x['Body'])), axis=1)
df_original["Histórico Emails"] = df_original.apply(lambda x: get_historico(str(x['Body'])), axis=1)
df_original['Duplicado'] = df_original['Duplicado'] = df_original.duplicated('Body')
#Drop column Concatenated
df_original["Concatenated"] = df_original["Subject"] + " " + df_original["Body"]
df_original["Email Curto"] = df_original["Concatenated"].apply(lambda x: isinstance(x, str) and len(x.split()) <= 5) #df_original["Body"].apply(lambda x: len(str(x).split()) <= 5)
df_original['Label']
df_original['ID Intenção'] = ''
#df_original['Label Intenção'] = df_original["Label"].map(label_map)
df_original['Nome Intenção'] = df_original["Label"].map(label_map)
df_original['Score'] = ''
df_original['ID Termos Expressões'] = df_original.apply(lambda x: get_top_three_keywords_counts(str(x["Subject"]) + " " + str(x['Body'])), axis=1)

df_original = df_original.drop('Concatenated', axis=1)
df_original = df_original[['Email Remetente', 'Data Email', 'Email ID', 'Subject', 'Body', 'NIF', 'Apolice', 'Nome', 'Mode', 'Histórico Emails', 'Duplicado', 'Email Curto',
                            'Label','ID Intenção', 'Nome Intenção', 'Score', 'ID Termos Expressões']]

#print(f"Formatado: {df_original}")
#df_original.to_excel("df_original_preprocessamento_10.xlsx")

# Cria uma cópia do corpo do email na coluna "Text" : Limpeza do dados  -> retira palavras que não são em lingua portuguesa, caracteres não imprimivéis e extra espaços. 
df_original["Text"] = df_original["Body"].copy() #.apply(helpers.clean)

# Se o texto do email for muito curto, adiciona o assunto ao texto para complementar
# Identificar textos vazios -> Emails não vazios = not em Inglês ou muito pequenos) 
#df_original["Text"] = df_original.apply(lambda x: x["Text"] if len(str(x["Text"]).split()) >= 5 else helpers.clean(x["Subject"]) + " " + str(x["Text"]), axis=1)

#print(df_original.apply(lambda x: x["Text"] if len(x["Text"].split()) >= 5 else helpers.clean(x["Subject"]) + " " + x["Text"], axis=1).head())

# Adiciona uma coluna indicando se o email é curto (menos de 4 palavras)
#df_original["Email Curto"] = df_original["Body"].apply(lambda x: len(str(x).split()) <= 5)

# Adiciona uma coluna indicando se o email é duplicado, com base no texto
#df_original['Duplicado'] = df_original.duplicated('Text')

df_filtered = df_original.sample(frac=1, random_state=42)


# FILTROS

# Remove  e-mails replied / e-mail chain
df_filtered = df_filtered[df_filtered["Histórico Emails"] == False]

# Remove duplicated emails
df_filtered = df_filtered[df_filtered['Duplicado'] == False]

# Remove short texts
df_filtered = df_filtered[df_filtered["Email Curto"] == False]
#df_filtered.to_excel("resultados_preprocessamento_2.xlsx")
print(df_filtered["Label"])

print(f" size df {df_filtered.shape}")

X = df_filtered["Subject"].fillna("") + " - " + df_filtered["Body"].fillna("")
y = df_filtered["Label"].rename("label")

# Train-test split
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)

# Over sampling
ros = RandomOverSampler(random_state=42)
X_train_, y_train_ = ros.fit_resample(np.asarray(X_train).reshape(-1, 1), y_train)

train_ = pd.concat([pd.DataFrame(X_train_, columns=["text"]), pd.Series(y_train_)], axis=1)
test_ = pd.concat([pd.DataFrame(X_test, columns=["text"]), pd.Series(y_test)], axis=1)

train_.shape, test_.shape

#df_original["Mode"] = list(map(lambda x:"TEST" if x in list (X_test.index) else "TRAIN", df_original.index))
df_filtered["Mode"] = list(map(
    lambda x: "TEST" if x in list(X_test.index) else "TRAIN",
    df_filtered.index
))

#print(f"DF_Original: {df_original['Mode']}")
#df_filtered.to_excel("mode_4.xlsx")
print(train_.shape)
print(test_.shape)

###########


# Load the BERT tokenizer and model
tokenizer = BertTokenizer.from_pretrained(MODEL_PATH, truncation=True, padding="max_length", max_length=128, return_tensors='pt')
model = BertForSequenceClassification.from_pretrained(
    MODEL_PATH,
    num_labels=NUM_LABELS,
    hidden_dropout_prob=0.2,
    attention_probs_dropout_prob=0.1
)#.to(DEVICE)

# Preprocessing and Preparing Dataset
def preprocess_function(df):
    return tokenizer(df["text"], truncation=True, padding=True, max_length=128)

tokenized_train = Dataset.from_pandas(train_).map(preprocess_function, batched=True)
tokenized_test = Dataset.from_pandas(test_).map(preprocess_function, batched=True)

ds = DatasetDict()

ds['train'] = tokenized_train
ds['validation'] = tokenized_test

data_collator = DataCollatorWithPadding(tokenizer)

# Training arguments
accuracy = evaluate.load("accuracy")



def compute_metrics(eval_pred):
    predictions, labels = eval_pred
    predictions = np.argmax(predictions, axis=1)
    return accuracy.compute(predictions=predictions, references=labels)

training_args = TrainingArguments(
    output_dir=TRAINED_MODEL_PATH,
    learning_rate=1.5e-5,
    per_device_train_batch_size=8,
    per_device_eval_batch_size=8,
    warmup_steps=100,
    num_train_epochs=32,
    weight_decay=0.01,
    logging_strategy="epoch",
    evaluation_strategy="epoch",
    save_strategy="epoch",
    load_best_model_at_end=True,
    save_total_limit=3,
    greater_is_better=True,
    metric_for_best_model="accuracy",
    group_by_length=True,
    #fp16=GPU,

)

# Train Model
trainer = Trainer(
    model=model,
    args=training_args,
    train_dataset=ds["train"],
    eval_dataset=ds["validation"],
    tokenizer=tokenizer,
    data_collator=data_collator,
    compute_metrics=compute_metrics,
    callbacks=[EarlyStoppingCallback(early_stopping_patience=5)]
)

trainer.train() #

# Save model
trainer.model.save_pretrained(TRAINED_MODEL_PATH)



# Quick validation

model = BertForSequenceClassification.from_pretrained(TRAINED_MODEL_PATH, num_labels=NUM_LABELS)
tokenizer = BertTokenizer.from_pretrained(MODEL_PATH, truncation=True, padding="max_length", max_length=128)

#model.to(DEVICE)

clf = pipeline("text-classification", model=model, tokenizer=tokenizer, truncation=True, padding="max_length", max_length=128)#, device=DEVICE)
preds = clf(test_["text"].to_list())
preds = [int(pred["label"].split('_')[1]) for pred in preds]

#print(classification_report(test_["label"], preds, target_names=[name for i, name in sorted(label_map.items())]))
#print(confusion_matrix(test_["label"].astype(int), preds))

df_final = df_original.copy()

df_final["Mode"] = list(map(lambda x:"TEST" if x in list (X_test.index) else "TRAIN", df_final.index))

# Get predictions
model = BertForSequenceClassification.from_pretrained(TRAINED_MODEL_PATH, num_labels=NUM_LABELS)
tokenizer = BertTokenizer.from_pretrained(MODEL_PATH, truncation=True, padding="max_length", max_length=128)

#model.to(DEVICE)

clf = pipeline("text-classification", model=model, tokenizer=tokenizer, truncation=True, padding="max_length", max_length=128)#, device=DEVICE)

# Ensure 'Body' column is strings
df_final["Body"] = df_final["Body"].astype(str)

# Convert to list and pass to the model
text_list = df_final["Body"].to_list()
predictions = clf(text_list)
#predictions = clf(df_final["Body"].to_list())
labels = [int(pred["label"].split('_')[1]) for pred in predictions]
scores = [round(float(pred["score"]), 2) for pred in predictions]

now = datetime.now()
formatted_time = now.strftime("%d%m%Y_%H%M%S")
print(formatted_time)
df_final = df_final.drop('Text', axis=1)

df_final["Label"] = df_original["Label"].map(convert_back_map)  # if no topic 4 convert_back_map
df_final['ID Intenção'] = [convert_back_map[label] for label in labels]
#df_final['ID Intenção'] = labels
df_final['Nome Intenção'] = [label_map[label] for label in labels]
df_final["Score"] = scores

# Save results
df_final.to_excel(BASE_DIR+f"Resultados_{str(formatted_time).split()[0].replace('-', '')}.xlsx", index=False)
