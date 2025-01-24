import logging
from informacoes_sensiveis import informacoes_sensiveis as info_sensiveis

class erros():
    def __init__(self,quant = 0):
        self.quant=quant       

class Config_class():
##    def __init__(self):
    # Definindo as variáveis de configuração
    intervalo_min            =  3.5
    intervalo_max            =  7.5
    EMAIL_HOST               =  "smtp.hostinger.com"
    EMAIL_PORT               =  587
    EMAIL_USER               =  info_sensiveis.email
    EMAIL_PASSWORD           =  info_sensiveis.senha
    FROM_EMAIL               =  info_sensiveis.email
    LOG_FILE                 =  "email_log.txt"
    LIMITE_DIARIO            =  100  # Número máximo de e-mails a serem enviados por dia
    emails_enviados          =  set()
    emails_tentando_enviar    = set()
    nome_planilha            =  'ECONODATA.xlsx'
    erro_limite              =  3
    erros_consecutivos       = erros()
    tentativas_enviar_email  =  3
    filtro_de_cargos         =  ''

    ########## assuntos dos emails

    assunto_primeiro_email   =  "Como o domínio de idiomas pode transformar seu negócio?"
    assunto_segundo_email    =  "assunto do segundo email"
    assunto_terceiro_email   =  "assunto do terceiro e-mail"
    contador_emails_enviados = 0
    intervalo_entre_emails   = 2
    ####### objeto de rastreio dos loggs
    logging = logging
    logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s - %(message)s')
##    class error_messages():
##        
##        @classmethod
##        def autenticacao_ao_enviar_email(cls,email,erro):
##            return 
    @classmethod
    def atualiza_Emails_enviados(cls,email):
##        quant=len(cls.emails_enviados)
        cls.emails_enviados.add(email)
##        if quant == len(self.emails_enviados):
            #### futuramente posso colocar alguma logica aqui se achar necessário
##            pass #self.loggin()

##class Config_instance():
##    config=Config_class()
##        
