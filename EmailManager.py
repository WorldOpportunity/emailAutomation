# Arquivo: email_manager.py
from Config import Config_class as conf
import smtplib
import imaplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import funcoes_auxiliares as FAs
import asyncio

class EmailManager:
    def __init__(self, host = conf.EMAIL_HOST, port = conf.EMAIL_PORT, user = conf.EMAIL_USER, password = conf.EMAIL_PASSWORD, from_email = conf.FROM_EMAIL, imap_host = conf.IMAP_HOST,imap_port = conf.IMAP_PORT):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.from_email = from_email
        self.imap_host= imap_host
        self.imap_port = imap_port

    def conectar(self):
        "essa função retorna o objeto 'server' ou retorna falso"
##        print('entre no conectar do email manager')
        resultado=False
        for tentativa in range(conf.tentativas_enviar_email):
            try:
                server = smtplib.SMTP(self.host, self.port)
                server.starttls()
                server.login(self.user, self.password)
                resultado =  server
                break
            except Exception as e:
                print( f"Erro ao conectar no servidor SMTP: {e}" )
                conf.logging.error(f"Erro ao conectar no servidor SMTP: {e}")
        return resultado

    async def enviar_email(self, to_email, subject, body):
        """Envia o e-mail para o destinatário com a tentativa de reconexão."""
        resultado = False
        conf.emails_tentando_enviar.add(to_email)
        for tentativa in range(conf.tentativas_enviar_email):
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
                
                # salvar o email na pasta enviados
                self.salvar_em_enviados(msg)
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

    def salvar_em_enviados(self, msg):
        """Salva uma cópia do e-mail na pasta 'Enviados' usando IMAP."""
        try:
            with imaplib.IMAP4_SSL(self.imap_host, self.imap_port) as imap:
                imap.login(self.user, self.password)

                # Formata a mensagem para salvar
                raw_message = msg.as_string().encode("utf-8")

                # Seleciona a pasta "Enviados" (ou cria, se não existir)
                status, _ = imap.select('"Sent"')  # Padrão internacional para pasta de enviados
                if status != "OK":
                    conf.logging.warning("Pasta 'Sent' não encontrada. Tentando criar.")
                    imap.create('"Sent"')
                    imap.select('"Sent"')

                # Adiciona a mensagem na pasta "Enviados"
                imap.append('"Sent"', "\\Seen", imaplib.Time2Internaldate(), raw_message)
                print("E-mail salvo na pasta 'Enviados' com sucesso.")
                conf.logging.info("E-mail salvo na pasta 'Enviados' com sucesso.")

        except Exception as e:
            conf.logging.error(f"Erro ao salvar e-mail na pasta 'Enviados': {e}") 

