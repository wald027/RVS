
import re
import spacy
import langid
import pandas as pd
import datetime



nlp = spacy.load("pt_core_news_lg")


## PREPROCESSING

def clean(text):
    """
    Some text cleaning functions.
    """
    # remove non portuguese sentences:
    for sentence in text.split("\n"):
        for sent in sentence.split("."):
            if langid.classify(sent)[0] != "pt" and len(sent) > 1:
                text = text.replace(sent, " ")

    # remove non printtable characters
    text = text.replace("\n", " ").replace("\r", " ")

    # remove double spaces
    text = re.sub(' +', ' ', text)
    text = re.sub(' \.', '.', text)
    text = re.sub('\.+', '.', text)

    return text.lstrip(". ")


## POST PROCESSING

def validate_apolice(modalidade, apolice):
    """
    Validates if an apolice numbers follows specific rules.
    """
    current_year = datetime.datetime.now().year
    cond1 = int(modalidade) > 32
    cond2 = len(apolice) == 3 or len(apolice) > 4
    cond3 = (len(apolice) == 1 or len(apolice) == 2) and (4 <= int(modalidade) > 32 and int(apolice) > 13)
    cond4 = not (len(apolice) == 4 and 4 <= int(modalidade) <= 12 and 1900 <= int(apolice) <= current_year)
    return cond1 or cond2 or cond3 or cond4


def get_apolice(text):
    """
    Extracts valid apolices from email body.
    """
    numbers = re.findall(r'\d{1,2}[-/ ]?\d{1,6}', text)
    valid_numbers = []
    for number in numbers:
        split_char = "/" if "/" in number else "-" if "-" in number else " " if " " in number else None
        if split_char:
            modalidade, apolice = number.split(split_char)
            if validate_apolice(modalidade, apolice):
                valid_numbers.append(number)
    return " | ".join(list(set(valid_numbers))) if valid_numbers else ""


# TODO: alguns NIFs estão vindo fora de ordem, além de se repetirem, por isso a dataframe possui mais linha que o normal. Também há situações que eles não são selecionados
def validate_nif(nif: str|int) -> bool:
    """
    Validates if a nif is valid.
    """
    if len(str(nif)) != 9:
        return False
    if int(nif[0]) == 0:
        return False

    s = 9 * int(nif[0]) + 8 * int(nif[1]) + 7 * int(nif[2]) + 6 * int(nif[3]) + 5 * int(nif[4]) + 4 * int(nif[5]) + 3 * int(nif[6]) + 2 * int(nif[7]) + 1 * int(nif[8])
    resto = s % 11

    if resto == 1:
        return int(nif[8]) == 0
    else:
        return resto == 0


def get_nif(text: str) -> str:
    """
    Extracts valid NIFs from email body.
    """
    numbers = re.findall(r'\d{3}[ -]?\d{3}[ -]?\d{3}', text)
    valid_nifs = [number for number in numbers if validate_nif(number)]
    return " | ".join(list(set(valid_nifs))) if valid_nifs else ""


def get_names(text):
    noise = ["pt50", "diretor", "subdiretor", "professor"]  # ["agradecia", "obrigada", "pt", "rua", "lisboa", "nif", "cumprimentos"]
    names = []
    docs = nlp(" ".join([t.capitalize() for t in str(text).split()]))
    entities = docs.ents
    for entity in entities:
        if entity.label_ == "PER" and len(entity.text) >= 4:
            doc = nlp(entity.text)
            names_partial = []
            for token in doc:
                if token.ent_type_ == "PER" and 4 <= len(token.text) <= 12 and token.pos_ == "PROPN" and token.text.lower() not in noise:
                    txt = re.sub(" +", " ", token.text.replace('\n', " ")).strip()
                    names_partial.append(txt)
            names_concat = " ".join(names_partial).strip()
            if len(names_concat) > 0 and names_concat not in names:
                names.append(names_concat)
    return " | ".join(list(set(names)))


def get_top_three_keywords_counts(text):
    """
    Counts the number of times each keyword appears in the text.
    Returns the three templates with the highest keywords count.
    """
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

    counts = {k:0 for k in template_keywords.keys()}
    text = str(text)

    for template, keywords in template_keywords.items():
        for keyword in keywords:
            if re.findall(keyword, text.lower()):
                counts[template] += 1

    top_three = sorted(counts.items(), key=lambda item: item[1], reverse=True)[:3]
    top_three_labels = [template_to_standard[label] for label, count in top_three]

    return top_three_labels
