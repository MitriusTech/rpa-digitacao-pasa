import pandas as pd
import re

def normalize_hyphens(s):
    # Substitui diferentes tipos de traços por '-'
    return re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2015]', '-', s)


def normalize_dict_hyphens(data):
    if isinstance(data, dict):
        # Se for um dicionário, processa recursivamente as chaves e valores
        return {key: normalize_dict_hyphens(value) for key, value in data.items()}
    elif isinstance(data, list):
        # Se for uma lista, processa cada elemento
        return [normalize_dict_hyphens(item) for item in data]
    elif isinstance(data, str):
        # Se for uma string, aplica a normalização de traços
        return normalize_hyphens(data)
    else:
        # Se não for string, lista ou dicionário, retorna o valor original
        return data


def get_parameters(report_name=None):
    path = './data' if not report_name else f'./data/{report_name}'
    return __get_parameters(f'{path}/parameters.xlsx')


def __get_parameters(xlsx="./data/parameters.xlsx", sheet_name="values", key="key", value="value"):
    df = pd.read_excel(xlsx, engine='openpyxl', sheet_name=sheet_name)
    return normalize_dict_hyphens(dict(zip(df[key], df[value])))

global_parameters = get_parameters()