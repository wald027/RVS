class BusinessRuleException(Exception):
    """Base class for business rule exceptions"""
    pass

class BusinessRuleExpection(BusinessRuleException):
    """Raised when an invalid order amount is detected"""
    def __init__(self,message):
        self.message = f"Definição do Negócio: {message}" 
        super().__init__(self.message)

    def __str__(self):
        return f'{self.message}'



raise BusinessRuleException(BusinessRuleException('Teste'))