import asyncio , time , smtplib ,  openpyxl , random ,  os , copy 
from datetime import datetime
from Config import Config_class as config
from loggerManager import LoggerManager
loggerProgressManager = LoggerManager()
from EmailManager import EmailManager
EmailManager = EmailManager()
import funcoes_auxiliares as FAs


def condicao_enviar_email(email,primeiro_email,segundo_email,terceiro_email):
    """essa função retorna falso se não for pra enviar e-mail e retorna o numero do email que é para ser enviado caso seja para enviar
        elsa assume que todos os valores são validos 
"""
##    print('')
##    print(f'o email: {email} com as datas {primeiro_email}  {segundo_email}  {terceiro_email}')
    if not email or email in config.emails_enviados: # valida se é um email e ve se esse email ja consta nos enviados
##        print('retornou False')
        return False
    if FAs.eh_data_valida(primeiro_email):
        if dias_passados(primeiro_email)  <  config.intervalo_entre_emails:
##            print('retornou False')
            return False
        else:
            if FAs.eh_data_valida(segundo_email):
                if dias_passados(segundo_email)  <  config.intervalo_entre_emails:
##                    print('retornou False')
                    return False
                else:
                    if FAs.eh_data_valida(terceiro_email):
##                        print('retornou False')
                        return False
                    else:
##                        print('retornou 3')
                        return 3
            else:
##                print('retornou 2')
                return 2
    else:
##        print('retornou 1')
        return 1  #### para enviar o primeiro email
    

def pega_email_e_datas_da_linha(linha , colunas_existentes):
    valores=[]
    try:
        valores=[x.value for x in linha]
        linha = valores
    except:
        pass
    email = FAs.extrair_email(linha[colunas_existentes["EMAIL"]],filtro='rh')
    primeiro_email = linha[colunas_existentes['Primeiro E-MAIL ENVIADO?']]
    segundo_email  = linha[colunas_existentes['Segundo email enviado?']]
    terceiro_email = linha[colunas_existentes['Terceiro email enviado?']]
    return email , primeiro_email  ,  segundo_email  ,  terceiro_email

def atualiza_planilha_com_logs(planilha, loggerProgressManager, colunas_existentes):
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
    global sheet
    estados_salvos = loggerProgressManager.get_all_state()
    
    if not isinstance(estados_salvos, dict):
        raise ValueError("O estado retornado por get_all_state não é um dicionário válido.")
        
        # Itera pelas coordenadas e valores salvos
    print(f'começando a atualizaçao de valores ( {len(estados_salvos)} ) na tabela de acordo com os dados dos logs')
    linhas_para_verificar=set()
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
        linhas_para_verificar.add(linha)

    for linha in linhas_para_verificar:
        linha  =  planilha[linha]
        email,primeiro_email,segundo_email,terceiro_email = pega_email_e_datas_da_linha(linha , colunas_existentes)
        if not condicao_enviar_email(email,primeiro_email,segundo_email,terceiro_email):
            config.atualiza_Emails_enviados(linha[colunas_existentes["EMAIL"]]) #adiciona o email referente ao valor ja salvo aos emails enviados
    print('numero de emails que ja enviei e que não é para enviar ainda: ',len(config.emails_enviados))
                                                                   
                
    print('terminei de atualizar os valores')                
##        print(f' a coordenada {letra_da_coordenada} foi atualizada para: {planilha[letra_da_coordenada].value}')
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
##    print('o tipo do sheet antes da ajustar colunas é: ',sheet)
    tiveQueCriar,colunas_existentes = ajustar_colunas(sheet)
##    print('o tipo do sheet depois da ajustar_colunas e antes da atualiza_planilha_com_logs é: ',sheet)
    atualiza_planilha_com_logs(sheet, loggerProgressManager , colunas_existentes)
##    print('o tipo do sheet depois do atualiza_planilha_com_logs é: ',sheet)
    
    return tiveQueCriar , colunas_existentes , sheet , workbook

def ajustar_colunas(sheet):
    colunas_necessarias = ['EMAIL', 'Nome Completo', 'Primeiro E-MAIL ENVIADO?', 'Segundo email enviado?', 'Terceiro email enviado?']
    header = sheet[1]
    colunas_existentes = [cell.value for cell in header]
    novasColunas={}
    for index,value in enumerate(colunas_existentes):
        novasColunas[value]  =  index 

    colunas_existentes = novasColunas
    # Se alguma coluna necessária não existir, criá-la
    tiveQueCriar = False
    for coluna in colunas_necessarias:
        if coluna not in colunas_existentes:
            sheet.cell(row=1, column=len(colunas_existentes) + 1, value=coluna)
            colunas_existentes[coluna] = len(colunas_existentes) + 1
            print(f"Coluna '{coluna}' criada.")
            tiveQueCriar = True
            
    if tiveQueCriar:
        print('tive que criar coluna nova dentro do ajustar coluna')
        sheet, workbook = carregar_planilha()
        if sheet is None or workbook is None:
            config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
            print("Não foi possível carregar a planilha. Encerrando o programa.")
            return  # Não prosseguir se a planilha não foi carregada corretamente
    return tiveQueCriar,colunas_existentes

def altera_e_salva(row_index,column_index,value,sheet,workbook):
    coluna_na_tabela  =  column_index + 1
    letra_da_tabela   =  FAs.numero_para_letra_coluna(coluna_na_tabela)
    linha_da_tabela   =  row_index  +  1
##    print(f'a sigla usada para acesar a celula na hora de alterar é: {letra_da_tabela}{row_index}')
    coordenada = f"{letra_da_tabela}{linha_da_tabela}"
    sheet[coordenada].value = value
    print(f'no altera e salva temos o row_index = {row_index} o column_index = {column_index} e a coordenada mudada na tabela é: {coordenada} porem o index salvo nos logs para ser pego acrescenta + 1')
    loggerProgressManager.update(row_index = linha_da_tabela, column_index = coluna_na_tabela , new_value = value)
##    print(f'alterei a tabela e o valor da celula {coordenada}  é ' )
##    print(sheet[coordenada].value)
    #salvar_planilha_sem_continuar(workbook,sheet)
    return sheet, coordenada

def dias_passados(data_str = datetime.now()):
    try:
        if data_str  ==  'sim':
            return 0
        
        data_email = datetime.strptime(data_str, '%Y-%m-%d %H:%M:%S').replace(hour=0, minute=0, second=0, microsecond=0)  # Converte string de data para objeto datetime
        data_atual = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)  # Obtém a data e hora atuais
        delta = data_atual - data_email  # Calcula a diferença de tempo
##        print(f'diferença de dias foi:{delta}')
        return delta.days  # Retorna o número de dias passados
    except Exception as e:
        config.logging.error(f"Erro ao calcular dias passados: {str(e)}")
        return 0  # Retorna 0 em caso de erro, evitando que o erro cause falha no processamento


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
            sheet, coordenada  = altera_e_salva(row_index,indice_email,'Email inválido',sheet,workbook)

        elif resultado == 'Erro inesperado':
            pass    # Manter tentativa de envio
            
        else:
            sheet, coordenada  = altera_e_salva(row_index,indice_email,time.strftime('%Y-%m-%d %H:%M:%S'),sheet,workbook)
        
# Função para processar os e-mails
async def processar_emails(config,sheet,workbook,fechar_programa):
##    print('o tipo do sheet no inicio do processar_emails  é: ',type(sheet))
##    print('o sheet é:',sheet)

    # Verifica se a planilha foi carregada corretamente
    if sheet is None or workbook is None or fechar_programa:
        config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
        print("Não foi possível carregar a planilha. Encerrando o programa.")
        return  # Não prosseguir se a planilha não foi carregada corretamente

    if config.contador_emails_enviados >= config.LIMITE_DIARIO:
        config.logging.info(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado às {datetime.now()}")
##        print(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado. ")
        print('vou descansar por uma hora')
        await asyncio.sleep(60*60)        
        config.contador_emails_enviados=0  ## reiniciando a contagem de emails ate o próximo intervalo
        config.logging.info(f"voltei a enviar e-mail às {datetime.now()}")
    else:
        pass
    row_index = 2  # Começando da segunda linha para evitar o cabeçalho

    semaphore = asyncio.Semaphore(2)  # Limita para 2 e-mails sendo enviados ao mesmo tempo

    # Loop para percorrer as linhas da planilha
##    print('o tipo do sheet antes da linha do erro é: ',type(sheet))
##    print('o sheet antes da linha do erro é: ',sheet)
    corpo_primeiro_email = str(sheet[2][colunas_existentes['Corpo primeiro e-mail'  ]].value   or "Mensagem padrão caso o corpo não esteja definido.") ## aqui retorna o valor do email como string
    corpo_segundo_email  = str(sheet[2][colunas_existentes['Corpo segundo e-mail'   ]].value   or "Mensagem padrão caso o corpo não esteja definido.") ## aqui retorna o valor do email como string         
    corpo_terceiro_email = str(sheet[2][colunas_existentes['Terceiro email enviado?']].value   or "Mensagem padrão caso o corpo não esteja definido.") ## aqui retorna o valor do email como string

    indice_primeiro_email  = FAs.obter_indice_coluna(colunas_existentes, 'Primeiro E-MAIL ENVIADO?') ## aqui retorna o valor do indice como int
    letra_indice_primeiro_email = FAs.numero_para_letra_coluna(indice_primeiro_email)
    
    indice_segundo_email   = FAs.obter_indice_coluna(colunas_existentes, 'Segundo email enviado?')## aqui retorna o valor do indice como int
    letra_segundo_primeiro_email = FAs.numero_para_letra_coluna(indice_segundo_email)
    
    indice_terceiro_email  = FAs.obter_indice_coluna(colunas_existentes, 'Terceiro email enviado?')## aqui retorna o valor do indice como int
    letra_terceiro_primeiro_email  = FAs.numero_para_letra_coluna(indice_terceiro_email)                         

    while row_index <= sheet.max_row:
##        try:
        if config.contador_emails_enviados >= config.LIMITE_DIARIO:
            config.logging.info(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado às {datetime.now()}")
        ##        print(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado. ")
            print('vou descansar por uma hora')
            await asyncio.sleep(60*60)        
            config.contador_emails_enviados=0  ## reiniciando a contagem de emails ate o próximo intervalo
            config.logging.info(f"voltei a enviar e-mail às {datetime.now()}")
            print('reiniciando o trabalho')
##                break  # Interrompe o loop se o limite de e-mails foi alcançado
        else:
            pass
        row   = [ x.value for x in sheet[row_index]]
        nome  = str( row[colunas_existentes['Nome Completo']] or '')
        email, primeiro_email, segundo_email, terceiro_email = pega_email_e_datas_da_linha(row,colunas_existentes)
        vai_enviar = False
        if not email:
            row_index  +=  1
            continue
        if email in config.emails_enviados or email in config.emails_tentando_enviar:
            print('e-mail repetido') 
            row_index += 1
            continue
        
        if config.filtro_de_cargos:
            cargo = FAs.get_cargo(row,colunas_existentes)
            pass
                
            
            # Verifica se o e-mail ainda não foi enviado
        condicao= condicao_enviar_email(email,primeiro_email,segundo_email,terceiro_email)
        if not condicao:
##                print(f'não é para enviar email para: {email} com os valores primeiro_email: {primeiro_email}, segundo_email: {segundo_email}, terceiro_email: {terceiro_email}')
            row_index  +=  1
            continue
        if condicao == 1:
            print(f'vou mandar o email: {email} como primeiro email')
            vai_enviar = True
            resultado = await enviar_email_com_concorrencia(row, config.assunto_primeiro_email, corpo_primeiro_email, email, nome,  semaphore)
            parar  =  trata_erros_nos_emails_e_salva_planilha(resultado,config.logging,config.erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes='Primeiro E-MAIL ENVIADO?',indice_email = indice_primeiro_email)
            if parar:
                break
        if condicao == 2:
            print(f'email: {email} como segundo email operação suspensa')
            vai_enviar = False
##                resultado = await enviar_email_com_concorrencia(row, config.assunto_segundo_email, corpo_segundo_email, email, nome, semaphore)
##                parar=trata_erros_nos_emails_e_salva_planilha(resultado,config.logging,config.erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes='Segundo email enviado?', indice_email = indice_segundo_email)
##                if parar:
##                    break
        if condicao == 3:
            print(f'email: {email} como terceiro email operação suspensa')
            vai_enviar  =  False
##                resultado = await enviar_email_com_concorrencia(row, config.assunto_terceiro_email, corpo_terceiro_email, email, nome, semaphore)
##                parar = trata_erros_nos_emails_e_salva_planilha(resultado,config.logging,config.erros_consecutivos,row_index,row,sheet,workbook,email,colunas_existentes,valor_colunas_existentes='Terceiro email enviado?',indice_email = indice_terceiro_email)
##                if parar:
##                    break

        row_index += 1
            # Esperar antes de enviar o próximo e-mail
        intervalo = random.uniform(config.intervalo_min, config.intervalo_max)
        if vai_enviar:
            await asyncio.sleep(intervalo)

##        except Exception as e:
##            config.logging.error(f"Erro ao processar a linha {row_index}: {str(e)}")
##            print(f"Erro ao processar a linha {row_index}. Erro: {str(e)}")
##            row_index += 1  # Continuação do processamento, mas registrando o erro
    print('Processamento de e-mails concluído!')

tiveQueCriar , colunas_existentes , sheet , workbook = carregar_planilha()
fechar_programa = False
if tiveQueCriar:
    print('tive que criar coluna nova la no fim do codigo')
    tiveQueCriar , colunas_existentes , sheet , workbook = carregar_planilha()
    if sheet is None or workbook is None:
        fechar_programa = True
        config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
        print("Não foi possível carregar a planilha. Encerrando o programa.")
##async def iniciar():
##    await processar_emails()

# Função principal para rodar o código
async def main():
    await processar_emails(config,sheet,workbook,fechar_programa)

if __name__ == '__main__':
    asyncio.run(main())
