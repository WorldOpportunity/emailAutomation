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

def extrair_email(string):
    # Expressão regular para detectar um e-mail válido
    padrao_email = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    # Procurar por um e-mail na string
    try:
        resultado = re.findall(padrao_email, string)
        
        # Se encontrar algum e-mail, retorna o primeiro encontrado, caso contrário, retorna None
        if resultado:
            return resultado[0]
        else:
            return None
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


# Na função enviar_email
##async def enviar_email(row, assunto, corpo, email, nome):
##    global config
##    corpo = corpo.replace("{nome}", nome)
##    tentativas = 3  # Número máximo de tentativas
##    resultado = False
##    msg = MIMEMultipart()
##    msg['From'] = config.FROM_EMAIL
##    msg['To'] = email
##    msg['Subject'] = assunto
##    msg.attach(MIMEText(corpo, 'html'))
##    for tentativa in range(tentativas):
##        try:
##            if tentar_conectar_com_reconexao(EMAIL_HOST = config.EMAIL_HOST,EMAIL_PORT = config.EMAIL_PORT,EMAIL_USER = config.EMAIL_USER,EMAIL_PASSWORD = config.EMAIL_PASSWORD):
##                print('conectei ao server com o email_user=  ',config.EMAIL_USER)          
##                try:  # Enviando e-mail
##                    server = smtplib.SMTP(config.EMAIL_HOST, config.EMAIL_PORT)
##                    server.starttls()
##                    server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
##                    server.sendmail(config.FROM_EMAIL, email, msg.as_string())
##                    config.contador_emails_enviados += 1  # Incrementando o contador
##                    config.emails_enviados.append(email)
##                    print(f'E-mail enviado com sucesso para {email}. E-mails enviados hoje: {config.contador_emails_enviados}/{config.LIMITE_DIARIO}')
##                    resultado = True
##                except smtplib.SMTPConnectError as e:
##                    config.logging.error(f"Erro de conexão com o servidor SMTP: {e}")
##                    print(f"Erro de conexão com o servidor SMTP: {e}")
##                except smtplib.SMTPAuthenticationError as e:
##                    config.logging.error(f"Erro de autenticação no servidor SMTP: {e}")
##                    print(f"Erro de autenticação no servidor SMTP: {e}")
##                except smtplib.SMTPException as e:
##                    config.logging.error(f"Erro ao enviar e-mail: {e}")
##                    print(f"Erro ao enviar e-mail: {e}")
##                finally:
##                    server.quit()
##                
##            else:
##                print('a verificação de conexão com o servidor retornou falso')
##                
##        except smtplib.SMTPAuthenticationError as e:
##            config.logging.error(f"Erro de autenticação ao enviar e-mail para {email}: {str(e)}")
##            print(f'Erro de autenticação ao enviar e-mail para {email}')
##            resultado = 'Erro de autenticação'
##        except smtplib.SMTPRecipientsRefused as e:
##            config.logging.error(f"Destinatário recusado ao enviar e-mail para {email}: {str(e)}")
##            print(f'Destinatário recusado ao enviar e-mail para {email}')
##            resultado =  'Destinatário recusado'
##        except smtplib.SMTPConnectError as e:
##            config.logging.error(f"Erro de conexão ao enviar e-mail para {email}: {str(e)}")
##            if tentativa < tentativas - 1:  # Tentar novamente até o limite de tentativas
##                print(f"Tentando novamente... {tentativa + 1}/{tentativas}")
##                await asyncio.sleep(5)  # Espera de 5 segundos antes de tentar novamente
##                continue
##            else:
##                config.logging.error(f"Falha ao conectar após {tentativas} tentativas para {email}.")
##            resultado =  'Erro de conexão'
##        except Exception as e:
##            config.logging.error(f"Erro inesperado ao enviar e-mail para {email}: {str(e)}")
##            print(f'Erro inesperado ao enviar e-mail para {email}: {str(e)}')
##            resultado =  'Erro inesperado'
##    return resultado
##
##async def altera_tabela(row_index,column_index,value,sheet,workbook):
####    sheet.cell(row  =  row_index, column  =  column_index).value  =  value
##    sheet[f"{FAs.numero_para_letra_coluna(column_index)}{row_index}"] = value
##    print('alterei a tabela e o valor da celula ',row_index,'  ',column_index, '  é ' )
##    print(sheet[f"{FAs.numero_para_letra_coluna(column_index)}{row_index}"].value)
##    await salvar_planilha(workbook,sheet)
##    return sheet

##def verificar_conexao_smtp(EMAIL_HOST,EMAIL_PORT,EMAIL_USER,EMAIL_PASSWORD):
##    try:
##        # Criando o objeto SMTP
##        with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as server:
##            # Tentar conectar ao servidor SMTP
####            server.set_debuglevel(1)  # Ativa a depuração para ver a comunicação com o servidor SMTP
##            server.starttls()  # Inicia a criptografia
##            server.login(EMAIL_USER, EMAIL_PASSWORD)  # Tenta fazer o login com as credenciais fornecidas
##            print("Conexão estabelecida com sucesso ao servidor SMTP.")
##            return True
##    except smtplib.SMTPException as e:
##        print(f"Falha na conexão com o servidor SMTP: {str(e)}")
##        return False
##
##def tentar_conectar_com_reconexao(EMAIL_HOST,EMAIL_PORT,EMAIL_USER,EMAIL_PASSWORD):
##    tentativas_max = 5
##    for tentativa in range(tentativas_max):
##        if verificar_conexao_smtp(EMAIL_HOST,EMAIL_PORT,EMAIL_USER,EMAIL_PASSWORD):
##            return True
##        else:
##            print(f"Tentando novamente a conexão... Tentativa {tentativa + 1}/{tentativas_max}")
##            time.sleep(5)  # Espera 5 segundos antes de tentar novamente
##    logging.error("Falha ao conectar ao servidor SMTP após várias tentativas.")
##    return False
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
