import logging

class erros():
    def __init__(self,quant = 0):
        self.quant=quant       

class Config_class():
    def __init__(self):
        # Definindo as variáveis de configuração
        self.intervalo_min            =  3.5
        self.intervalo_max            =  7.5
        self.EMAIL_HOST               =  "smtp.hostinger.com"
        self.EMAIL_PORT               =  587
        self.EMAIL_USER               =  "no-reply@worldopportunity.com.br"
        self.EMAIL_PASSWORD           =  "#Curso123456"
        self.FROM_EMAIL               =  "no-reply@worldopportunity.com.br"
        self.LOG_FILE                 =  "email_log.txt"
        self.LIMITE_DIARIO            =  100  # Número máximo de e-mails a serem enviados por dia
        self.emails_enviados          =  []
        self.nome_planilha            =  'ECONODATA.xlsx'
        self.erro_limite              =  3
        self.erros_consecutivos       = erros()
        self.tentativas_enviar_email  =  3
        ########## assuntos dos emails
        self.assunto_primeiro_email   =  "Como o domínio de idiomas pode transformar seu negócio?"
        self.assunto_segundo_email    =  "assunto do segundo email"
        self.assunto_terceiro_email   =  "assunto do terceiro e-mail"
        self.contador_emails_enviados = 0
        
        self.logging = logging
        self.logging.basicConfig(filename=self.LOG_FILE, level=logging.INFO, format='%(asctime)s - %(message)s')


class Config_instance():
    config=Config_class()
        
