class BusinessRuleException(Exception):
    def __init__(self,message):
        self.message = f"Definição do Negócio: {message}" 
        super().__init__(self.message)

    def __str__(self):
        return f'{self.message}'
    pass

"""
try:
    raise Exception('Ola')
    raise BusinessRuleException('Teste')
except BusinessRuleException as e:
    print(e)
except Exception as e :
    print(e)
"""