import asyncio , time , smtplib ,  openpyxl , random ,  os , copy 
from datetime import datetime
from Config import Config_instance
config = Config_instance().config
from loggerManager import LoggerManager
loggerProgressManager = LoggerManager()
from EmailManager import EmailManager
EmailManager = EmailManager()
import funcoes_auxiliares as FAs



def atualiza_planilha_com_logs(planilha, loggerProgressManager):
    """
    Carrega as alterações persistidas anteriormente no LoggerManager para a planilha.

    Args:
        planilha: Objeto representando a planilha atual.
        loggerProgressManager: Instância do LoggerManager responsável pelo rastreamento de alterações.

    Returns:
        None
    """
##    try:
        # Obtém o estado de todas as alterações registradas
    estados_salvos = loggerProgressManager.get_all_state()
        
    if not isinstance(estados_salvos, dict):
        raise ValueError("O estado retornado por get_all_state não é um dicionário válido.")
        
        # Itera pelas coordenadas e valores salvos
    print(f'começando a atualizaçao de valores ( {len(estados_salvos)} ) na tabela de acordo com os dados dos logs')
    for coordenadas, valor in estados_salvos.items():
        if (
                not isinstance(coordenadas, tuple) 
                or len(coordenadas) != 2 
                or not all(isinstance(c, int) for c in coordenadas)
            ):
                raise ValueError(f"Coordenadas inválidas detectadas: {coordenadas}")
            
        linha, coluna = coordenadas
        letra_da_coordenada= f"{FAs.numero_para_letra_coluna(coluna)}{linha}"
        planilha[letra_da_coordenada] = valor  # Atualiza a célula com o valor salvo
        print(f' a coordenada {letra_da_coordenada} foi atualizada para: {planilha[letra_da_coordenada].value}')
##    except Exception as e:
####        loggerProgressManager.log_error(f"Erro ao carregar planilha: {str(e)}")
##        raise


def carregar_planilha():
    global config
##    try:
    if not os.path.exists(config.nome_planilha):
        raise FileNotFoundError("O arquivo " + config.nome_planilha + " não foi encontrado.")
    workbook = openpyxl.load_workbook(config.nome_planilha)
    sheet = workbook.active
        ## com sorte, esse codigo vai carregar as mudanças registradas no logger dentro da versão na memória da planilha
    atualiza_planilha_com_logs(sheet, loggerProgressManager)
    return sheet, workbook

##    except FileNotFoundError as e:
##        config.logging.error(f"Erro ao carregar o arquivo: {str(e)}")
##        print(str(e))
##        return None, None  # Retorna None para indicar falha, mas não encerra o programa
##    except Exception as e:
##        config.logging.error(f"Erro ao carregar a planilha: {str(e)}")
##        print(f"Erro ao carregar a planilha: {str(e)}")
##        return None, None  # Retorna None para indicar falha, mas não encerra o programa
sheet, workbook = carregar_planilha()


def altera_e_salva(row_index,column_index,value,sheet,workbook):
    letra_da_tabela=FAs.numero_para_letra_coluna(column_index)
##    print(f'a sigla usada para acesar a celula na hora de alterar é: {letra_da_tabela}{row_index}')
    coordenada = f"{letra_da_tabela}{row_index}"
    sheet[coordenada].value = value
    loggerProgressManager.update(row_index, column_index = column_index + 1, new_value = value)
##    print(f'alterei a tabela e o valor da celula {coordenada}  é ' )
##    print(sheet[coordenada].value)
    #salvar_planilha_sem_continuar(workbook,sheet)
    return sheet, coordenada

def dias_passados(data_str = datetime.now()):
    try:
        if data_str  ==  'sim':
            return 0
        
        data_email = datetime.strptime(data_str, '%Y-%m-%d %H:%M:%S')  # Converte string de data para objeto datetime
        data_atual = datetime.now()  # Obtém a data e hora atuais
        delta = data_atual - data_email  # Calcula a diferença de tempo
        return delta.days  # Retorna o número de dias passados
    except Exception as e:
        config.logging.error(f"Erro ao calcular dias passados: {str(e)}")
        return 0  # Retorna 0 em caso de erro, evitando que o erro cause falha no processamento



def ajustar_colunas(sheet):
    colunas_necessarias = ['EMAIL', 'Nome Completo', 'Primeiro E-MAIL ENVIADO?', 'Segundo email enviado?', 'Terceiro email enviado?']
    header = sheet[1]
####    print('o heather é:   ',header)
##    print('o tipo do header é:  ',type(header))
    # Verificando as colunas existentes
    colunas_existentes = [cell.value for cell in header]
    novasColunas={}
    for index,value in enumerate(colunas_existentes):
        novasColunas[value]  =  index 

    colunas_existentes = novasColunas
##    print("colunas existentes é:",colunas_existentes)
    # Se alguma coluna necessária não existir, criá-la
    tiveQueCriar = False
    for coluna in colunas_necessarias:
        if coluna not in colunas_existentes:
            sheet.cell(row=1, column=len(colunas_existentes) + 1, value=coluna)
            colunas_existentes[coluna] = len(colunas_existentes) + 1
            print(f"Coluna '{coluna}' criada.")
            tiveQueCriar = True        

    return tiveQueCriar,colunas_existentes


# Função para salvar a planilha com re-tentativas
def salvar_planilha_sem_continuar(workbook,sheet):
    global config
    tentativas_max = 5
    for tentativa in range(tentativas_max):
        try:
            workbook.active  =  sheet  
            workbook.save(config.nome_planilha)
            print("Planilha salva com sucesso.")
            return
        except Exception as e:
            config.logging.error(f"Erro ao salvar a planilha (Tentativa {tentativa + 1}): {str(e)}")
            print(f'Erro ao salvar a planilha: {str(e)}')
            if tentativa == tentativas_max - 1:
                config.logging.error(f"Falha ao salvar a planilha após {tentativas_max} tentativas.")
                print(f"Falha ao salvar a planilha após {tentativas_max} tentativas.")
                break
            print('esperando 5 segundos...')
            time.sleep(5)  # Esperar 30 segundos antes de tentar novamente
            print('foi')
async def salvar_planilha(workbook,sheet):
    global config
    tentativas_max = 5
    for tentativa in range(tentativas_max):
        try:
            workbook.active  =  sheet  
            workbook.save(config.nome_planilha)
            print("Planilha salva com sucesso.")
            return
        except Exception as e:
            config.logging.error(f"Erro ao salvar a planilha (Tentativa {tentativa + 1}): {str(e)}")
            print(f'Erro ao salvar a planilha: {str(e)}')
            if tentativa == tentativas_max - 1:
                config.logging.error(f"Falha ao salvar a planilha após {tentativas_max} tentativas.")
                print(f"Falha ao salvar a planilha após {tentativas_max} tentativas.")
                break
            await asyncio.sleep(5)  # Esperar 30 segundos antes de tentar novamente

# Função para enviar e-mails com limite de concorrência
async def enviar_email_com_concorrencia(row, assunto, corpo, email, nome, semaphore):
    # Usar o semaphore para garantir que no máximo 2 e-mails sejam enviados ao mesmo tempo
    async with semaphore:
        corpo = corpo.replace("{nome}", nome)
        sucesso = await EmailManager.enviar_email(subject = assunto, body = corpo, to_email  =   email)
        return sucesso
def trata_erros_nos_emails_e_salva_planilha(resultado,logging,erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes,indice_email):
    if resultado in ['Erro de autenticação', 'Erro inesperado']:
        erros_consecutivos.quant += 1
        if erros_consecutivos.quant >= config.erro_limite:
            logging.error(f"Falha contínua ao enviar e-mails. Pausando o processo.")
            return True
    else:
        erros_consecutivos.quant = 0  # Resetar se o envio foi bem-sucedido

        if resultado == 'Erro de autenticação':
            logging.error(f'Falha na autenticação ao tentar enviar para {email}')
            return True  # Se falhar na autenticação, para o processo
        
        elif resultado == 'Destinatário recusado': # Marcar como inválido
            ########################### preciso implementar a mudança na planilha em si.... não só na linha 'in memory'
##            row[colunas_existentes['Primeiro E-MAIL ENVIADO?']]  = 'Email inválido'
            sheet, coordenada  = altera_e_salva(row_index,indice_email,'Email inválido',sheet,workbook)
##            print('alterei nos logs o registro de um e-mail inválido')

        elif resultado == 'Erro inesperado':
            pass    # Manter tentativa de envio
            
        else:
##                    indiceDoEmail  =  colunas_existentes['Primeiro EMAIL ENVIADO?']
##            print('alterando tabela devido a email enviado com sucesso')
##                    sheet  = await altera_tabela(row_index,indice_primeiro_email,time.strftime('%Y-%m-%d %H:%M:%S'),sheet,workbook)
            sheet, coordenada  = altera_e_salva(row_index,indice_email,time.strftime('%Y-%m-%d %H:%M:%S'),sheet,workbook)
            
##            print("depois de assar pela função que altera o valor da celula é: ",sheet[coordenada].value)
##            row[colunas_existentes[valor_colunas_existentes]]  = time.strftime('%Y-%m-%d %H:%M:%S')


        
# Função para processar os e-mails
async def processar_emails():
    global config
    global sheet, workbook 
    

    # Verifica se a planilha foi carregada corretamente
    if sheet is None or workbook is None:
        config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
        print("Não foi possível carregar a planilha. Encerrando o programa.")
        return  # Não prosseguir se a planilha não foi carregada corretamente

    if config.contador_emails_enviados >= config.LIMITE_DIARIO:
##        config.logging.info(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado. Encerrando o programa.")
##        print(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado. ")
        print('vou descansar por uma hora')
        await asyncio.sleep(60*60)        
        config.contador_emails_enviados=0
    else:
        pass
##        print(f'o contador de emails enviados no config é: {config.contador_emails_enviados} e o limite de emails diario é: {config.LIMITE_DIARIO}')
    
    tiveQueCriar,colunas_existentes = ajustar_colunas(sheet)
    if tiveQueCriar:
        print('tive que criar coluna nova')
        sheet, workbook = carregar_planilha()
        if sheet is None or workbook is None:
            config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
            print("Não foi possível carregar a planilha. Encerrando o programa.")
            return  # Não prosseguir se a planilha não foi carregada corretamente
            
    row_index = 2  # Começando da segunda linha para evitar o cabeçalho

    semaphore = asyncio.Semaphore(2)  # Limita para 2 e-mails sendo enviados ao mesmo tempo

    # Loop para percorrer as linhas da planilha
    while row_index <= sheet.max_row:
##            print(row_index)
        try:
            if config.contador_emails_enviados >= config.LIMITE_DIARIO:
                config.logging.info(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado. Encerrando o programa.")
                print(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado. descansando...")
                asyncio.sleep(60*60)
                config.contador_emails_enviados = 0
                print('reiniciando o trabalho')
##                break  # Interrompe o loop se o limite de e-mails foi alcançado
            else:
                pass
##                print(f'o contador de emails enviados no config é: {config.contador_emails_enviados} e o limite de emails diario é: {config.LIMITE_DIARIO}')
            row   = [ x.value for x in sheet[row_index]]
            nome  = str( row[colunas_existentes['Nome Completo']] or '')
            email = FAs.extrair_email(row[colunas_existentes['EMAIL']])
##            print(f"o numero da colçuna de emails é: {colunas_existentes['EMAIL']}")

            
            corpo_primeiro_email = str([x.value for x in sheet[2]][colunas_existentes['Corpo primeiro e-mail']] or "Mensagem padrão caso o corpo não esteja definido.")
            corpo_segundo_email  = str([x.value for x in sheet[2]][colunas_existentes['Corpo segundo e-mail']] or "Mensagem padrão caso o corpo não esteja definido.")         
            corpo_terceiro_email = str([x.value for x in sheet[2]][colunas_existentes['Terceiro email enviado?']] or "Mensagem padrão caso o corpo não esteja definido.")
                
            indice_primeiro_email  = FAs.obter_indice_coluna(colunas_existentes, 'Primeiro E-MAIL ENVIADO?')
            indice_segundo_email   = FAs.obter_indice_coluna(colunas_existentes, 'Segundo email enviado?')
            indice_terceiro_email  = FAs.obter_indice_coluna(colunas_existentes, 'Terceiro email enviado?')

            primeiro_email  ,  segundo_email  ,  terceiro_email = row[indice_primeiro_email]  ,  row[indice_segundo_email]  ,  row[indice_terceiro_email]
##            print('cheguei aqui')
            if email in config.emails_enviados :
                print('e-mail repetido') 
                row_index += 1
                continue
            if not email:
                pass
##                print('email vazio')
##                print('a row é: ',row)
##                print('a colunas_existentes é: ',colunas_existentes)
##                
                row_index  +=  1
                continue
            
            # Verifica se o e-mail ainda não foi enviado
            if email and primeiro_email != 'Email inválido' and not FAs.eh_data_valida(primeiro_email):
                print('o email é:', email)
##                print('a coluna primeiro email é: ', primeiro_email)
##                print('a row que entrou pra enviar no primeiro email é: ',row)
                resultado = await enviar_email_com_concorrencia(row, config.assunto_primeiro_email, corpo_primeiro_email, email, nome,  semaphore)
                parar  =  trata_erros_nos_emails_e_salva_planilha(resultado,config.logging,config.erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes='Primeiro E-MAIL ENVIADO?',indice_email = indice_primeiro_email)
                if parar:
                    break

            # Verifica se já passou o tempo para enviar o segundo e-mail
##            print('o valor do primeiro_email é: ',primeiro_email)
            if segundo_email is None and primeiro_email and dias_passados(primeiro_email) >= 5:
                print('a coluna segundo email é: ', segundo_email)    
                resultado = await enviar_email_com_concorrencia(row, config.assunto_segundo_email, corpo_segundo_email, email, nome, semaphore)
                parar=trata_erros_nos_emails_e_salva_planilha(resultado,config.logging,config.erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes='Segundo email enviado?', indice_email = indice_segundo_email)
                if parar:
                    break

            # Verifica se já passou o tempo para enviar o terceiro e-mail
##            print('o valor do segundo_email é: ',segundo_email)
            if terceiro_email is None and segundo_email and dias_passados(segundo_email) >= 7:
                print('a coluna terceiro email é: ', terceiro_email)
                resultado = await enviar_email_com_concorrencia(row, config.assunto_terceiro_email, corpo_terceiro_email, email, nome, semaphore)
                parar = trata_erros_nos_emails_e_salva_planilha(resultado,config.logging,config.erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes='Terceiro email enviado?',indice_email = indice_terceiro_email)
                if parar:
                    break

            row_index += 1
            # Esperar antes de enviar o próximo e-mail
            intervalo = random.uniform(config.intervalo_min, config.intervalo_max)
            await asyncio.sleep(intervalo)

        except Exception as e:
            config.logging.error(f"Erro ao processar a linha {row_index}: {str(e)}")
            print(f"Erro ao processar a linha {row_index}. Erro: {str(e)}")
            row_index += 1  # Continuação do processamento, mas registrando o erro
    print('Processamento de e-mails concluído!')


# Função principal para rodar o código
async def main():
    await processar_emails()

if __name__ == '__main__':
    asyncio.run(main())
