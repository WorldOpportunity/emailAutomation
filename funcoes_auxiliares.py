import re 
from datetime import datetime

""" aqui colocarei funcoes auxiliares """


def eh_data_valida(string):
    """
    Testa se uma string é um formato de data válido comum.
    Suporta formatos como: dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd, etc.
    """
    if string is None:
        return False
    
    # Se a string tiver data e hora, separa apenas a parte da data
    if ' ' in string:
        string = string.split(' ')[0]
    formatos_comuns = [
        "%d/%m/%Y",  # Dia/Mês/Ano (31/12/2025)
        "%m/%d/%Y",  # Mês/Dia/Ano (12/31/2025)
        "%Y-%m-%d",  # Ano-Mês-Dia (2025-12-31)
        "%d-%m-%Y",  # Dia-Mês-Ano (31-12-2025)
        "%m-%d-%Y",  # Mês-Dia-Ano (12-31-2025)
        "%d/%m/%y",  # Dia/Mês/Ano com dois dígitos (31/12/25)
        "%m/%d/%y",  # Mês/Dia/Ano com dois dígitos (12/31/25)
        "%Y/%m/%d",  # Ano/Mês/Dia (2025/12/31)
        "%d-%b-%Y",  # Dia-Mês com abreviação-Ano (31-Dec-2025)
        "%d %b %Y",  # Dia Mês com abreviação Ano (31 Dec 2025)
        "%d %B %Y",  # Dia Nome completo do Mês Ano (31 Dezembro 2025)
        "%Y.%m.%d",  # Ano.Mês.Dia (2025.12.31)
    ]
    
    for formato in formatos_comuns:
        try:
            # Tenta fazer o parsing da string usando o formato atual
            datetime.strptime(string, formato)
            return True
        except ValueError:
            pass  # Se der erro, continua com o próximo formato
    return False

def numero_para_letra_coluna(numero):
    """Converte um número de coluna em sua letra correspondente no Excel."""
    letras = ''
    while numero > 0:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(resto + 65) + letras
    return letras
def get_cargo(row,colunas_existentes):
    try:
        cargo=row[colunas_existentes["CARGO"]]
        return cargo
    except :
        print(f'erro ao procurar o cargo: {err}')
        return ''
def extrair_email(string,filtro = ''):
    # Expressão regular para detectar um e-mail válido
    padrao_email = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    # Procurar por um e-mail na string
    try:
        emails = re.findall(padrao_email, string)
        if emails:
            if filtro:
                # Retorna o primeiro e-mail que contém o filtro, se houver
                for email in emails:
                    if filtro.lower() in email.lower():
                        return email
                return None  # Nenhum e-mail atende ao filtro
            else:
                return emails[0]  # Retorna o primeiro e-mail encontrado
        return None # não encontrou e-mail
    except:
        return None

# Função para obter o índice de uma coluna pelo nome
def obter_indice_coluna(colunas_existentes, nome_coluna):
    try:# Procurar a coluna pelo nome        
        return colunas_existentes[nome_coluna]
    except ValueError:
        logging.error(f"Coluna '{nome_coluna}' não encontrada na planilha.")
        return None

######## deixarei aqui as funções que não uso mais porem que não quero perder o codigo, mas de forma comentada
##
##def verificar_emails_enviados_no_dia():
##    global config
##    hoje = datetime.now().strftime('%Y-%m-%d')
##    
##    # Verificar no log se há e-mails enviados hoje
##    try:
##        with open(config.LOG_FILE, 'r') as log_file:
##            for linha in log_file:
##                if hoje in linha and "E-mail enviado com sucesso" in linha:
##                    pass ## tenho que implementar isso aqui direito
##                ##config.contador_emails_enviados += 1
##    except FileNotFoundError:
##        config.logging.info("Arquivo de log não encontrado. Contagem de e-mails iniciada.")
