# Arquivo: email_manager.py
from Config import Config_class as conf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import funcoes_auxiliares as FAs
import asyncio

class EmailManager:
    def __init__(self, host = conf.EMAIL_HOST, port = conf.EMAIL_PORT, user = conf.EMAIL_USER, password = conf.EMAIL_PASSWORD, from_email = conf.FROM_EMAIL):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.from_email = from_email

    def conectar(self):
        "essa função retorna o objeto 'server' ou retorna falso"
##        print('entre no conectar do email manager')
        resultado=False
        for tentativa in range(conf.tentativas_enviar_email):
            try:
##                print('primeira linha ...')
##                print(f'o self.host é {self.host} e o self.port é: {self.port}')
                server = smtplib.SMTP(self.host, self.port)
##                print('segunda linha ...')
##                server.set_debuglevel(1)
##                print('terceira linha ...')
                server.starttls()
##                print('quarta linha ...')
                server.login(self.user, self.password)
##                print('quinta linha ...')
                resultado =  server
##                print('consegui conectar ao server')
                break
            except Exception as e:
                print( f"Erro ao conectar no servidor SMTP: {e}" )
                conf.logging.error(f"Erro ao conectar no servidor SMTP: {e}")
        return resultado

    async def enviar_email(self, to_email, subject, body):
        """Envia o e-mail para o destinatário com a tentativa de reconexão."""
        resultado = False
        conf.emails_tentando_enviar.add(to_email)
##        print('entrei na funcao do email manager de enviar email')
##        print(f"o email é: {to_email} o assunto é: {subject} e o texto é: {body}")
        for tentativa in range(conf.tentativas_enviar_email):
##            print('entrei no for do enviar email do manager')
            server = self.conectar()
            if not server:
                conf.logging.error("Falha na conexão, tentando novamente...")
                await asyncio.sleep(5)  # Espera antes de tentar novamente
                continue

            try:
                msg = MIMEMultipart()
                msg['From'] = self.from_email
                msg['To'] = to_email
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'html'))
                server.sendmail(self.from_email, to_email, msg.as_string())
                resultado = True
                print(f"E-mail enviado com sucesso para {to_email}")
                conf.logging.info(f"E-mail enviado com sucesso para {to_email}")
                conf.contador_emails_enviados += 1
                conf.emails_tentando_enviar.discard(to_email)
                conf.emails_enviados.add(to_email)
##                print('acabei de mandar o email e aqui dentro do email manager o contador é: ',conf.contador_emails_enviados)
                break  # Sai do loop caso o e-mail seja enviado com sucesso
            except smtplib.SMTPAuthenticationError as e:
                conf.logging.error(f"Erro de autenticação ao enviar e-mail para {to_email}: {e}")
            except smtplib.SMTPRecipientsRefused as e:
                conf.logging.error(f"Destinatário recusado ao enviar e-mail para {to_email}: {e}")
            except smtplib.SMTPException as e:
                conf.logging.error(f"Erro ao enviar e-mail para {to_email}: {e}")
            finally:
                server.quit()  # Garantir que o servidor seja fechado no final

            if not resultado and tentativa < conf.tentativas_enviar_email - 1:
                conf.logging.info(f"Tentando novamente... Tentativa {tentativa + 1}/{conf.tentativas_enviar_email}")
                await asyncio.sleep(5)  # Espera antes de tentar novamente

        conf.emails_tentando_enviar.discard(to_email)
        
        return resultado
    
##    async def enviar_email(self, to_email, subject, body):
##        resultado = False
##        for tentativa in range(conf.tentativas_enviar_email):
##            try:
##                with self.conectar() as server:
##                    msg = MIMEMultipart()
##                    msg['From'] = self.from_email
##                    msg['To'] = to_email
##                    msg['Subject'] = subject
##                    msg.attach(MIMEText(body, 'html'))
##                    server.sendmail(self.from_email, to_email, msg.as_string())
##                    resultado =  True
##                    break
##            except Exception as e:
##                conf.logging.error(f"Erro ao enviar e-mail para {to_email}: {e}")
##                resultado = False
##        return resultado
