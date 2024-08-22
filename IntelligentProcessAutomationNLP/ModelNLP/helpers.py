
import re
import spacy
import langid
import pandas as pd
from datetime import datetime
import nltk


nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')

from nltk.tokenize import word_tokenize
from nltk import pos_tag



nlp = spacy.load("pt_core_news_lg")


## PREPROCESSAMENTO

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


## POSPROCESSAMENTO

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
def get_apolice(text, emailID, logger) -> str:
    # Logger para a Base de Dados
    logger.info(f"A tratar apólice {emailID}")   
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
    logger.info(f"Fim do tratamento da apólice {emailID}")
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
def get_nif(text, logger):  
    # logger
    logger.info("A tratar NIF")
    nif_values = []
    numbers = find_numbers(text)
    
    # se não for null
    if numbers is not None:
        # loop pela lista
        for j in numbers:
            # se o elemento tiver comprimento 9
            if len(str(j)) == 9: 
                # validar o nif
                logger.info("Validar nif")        
                value = validar_nif(j)                              
                # não for vazia
                if value:
                    logger.info("NIF valido")    
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
            logger.info("Fim do tratamento do NIF")   
            # retornar uma string de valores   
            return '|'.join(set(nif_values))
        else: 
            # logger
            logger.info("Fim do tratamento do NIF")
            return " "
    else:
        # logger
        logger.info("Fim do tratamento do NIF")
        return " "


# Extrair os Nomes
def get_names(text, logger):
    # logger
    logger.info("A tratar Nome")
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
    logger.info("Fim do tratamento do Nome")
    return "|".join(list(set(names)))


def get_historico(text, emailID, logger):
    #logger
    logger.info(f"A tratar Nome {emailID}")
    # loop pelo texto
    for i in text:
        # se contiver algum dos termos abaixo
        if ("De:" in i and "Enviado:" in i and "Para:" in i and "Assunto:" in i or
            "De:" in i and "Enviado:" in i and "Para:" in i and "Cc:" in i and "Assunto:" in i or
            "De:" in i and "Data:" in i and "Assunto:" in i and "Para:" in i and "Cc:" in i or
            "De:" in i and "Enviada:" in i and "Para:" in i and "Assunto:" in i or
            "De:" in i and "Enviado:" in i and "Cc:" in i and "Assunto:" in i or
            "De:" in i and "Assunto:" in i and "Data:" in i and "Para:" in i or
            "De:" in i and "Date:" in i and "Subject:" in i and "To:" in i or
            "De:" in i and "Data:" in i and "Para:" in i and "Assunto:" in i or
            "De:" in i and "Data:" in i and "Assunto:" in i and "Para:" in i or
            "Data:" in i and "De:" in i and "Assunto:" in i and "Cc:" in i and "Para:" in i or
            "From:" in i and "Sent:" in i and "To:" in i and "Cc:" in i and "Subject:" in i or
            "From:" in i and "Sent on:" in i and "CC:" in i and "Subject:" in i or
            "From:" in i and "Sent:" in i and "To:" in i and "Subject:" in i or
            "From:" in i and "Data:" in i and "Assunto:" in i and "Para:" in i):
            #logger
            logger.info("Fim do tratamento do Histórico")
            #retorna "Tem histórico"
            return "TRUE"
        # se contiver algum dos termos abaixo
        elif ("Clientes Real Vida <info.clientes@realvidaseguros.pt <mailto:info.clientes@realvidaseguros.pt>" in i or
              "Real Vida Seguros <noreply@realvidaseguros.pt <mailto:noreply@realvidaseguros.pt>" in i or
              "Real Vida Seguros <documentoseletronicos@realvidaseguros.pt <mailto:documentoseletronicos@realvidaseguros.pt>" in i or
              "----- Mensagem de Real Vida Seguros <digital@cert.realvidaseguros.pt <mailto:digital@cert.realvidaseguros.pt> > ---------" in i or
              '----- Mensagem de Real Vida Seguros <noreply@realvidaseguros.pt <mailto:noreply@realvidaseguros.pt> > ---------' in i or
              "---------- Forwarded message ---------" in i or
              "-------- Mensagem original --------" in i):
            #logger
            logger.info(f"Fim do tratamento do Histórico {emailID}")
            return "TRUE" 
        #logger
        logger.info(f"Fim do tratamento do Histórico {emailID}")           
        return 'FALSE'



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

def df():
    #envia a dataframe para os demais ficheiros
    return df
