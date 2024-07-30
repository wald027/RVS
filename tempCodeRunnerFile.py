import pandas as pd
from realvidaseguros import BusinessRuleExceptions

file_path = r'C:\Users\brunofilipe.lobo\Documents\Code\realvidaseguros\intencoes.xlsx'
dfRegrasApoliceAtivas = pd.read_excel(file_path,keep_default_na=False,sheet_name='ApolAtivas')

print(dfRegrasApoliceAtivas)