import asyncio , time , smtplib ,  openpyxl , random ,  os , copy 
from datetime import datetime
from Config import Config_class as config
from loggerManager import LoggerManager
loggerProgressManager = LoggerManager()
from EmailManager import EmailManager
EmailManager = EmailManager()
import funcoes_auxiliares as FAs

class sheet_info():
    def __init__(self,sheet,colunas_existentes,tive_que_criar):
        self.sheet = sheet
        self.colunas_existentes=colunas_existentes
        self.tive_que_criar = tive_que_criar
        

def condicao_enviar_email(email,primeiro_email,segundo_email,terceiro_email):
    """essa função retorna falso se não for pra enviar e-mail e retorna o numero do email que é para ser enviado caso seja para enviar
        elsa assume que todos os valores são validos 
"""
    if not email or email in config.emails_enviados or email in config.emails_tentando_enviar: # valida se é um email e ve se esse email ja consta nos enviados
        return False
    if FAs.eh_data_valida(primeiro_email):
        if dias_passados(primeiro_email)  <  config.intervalo_entre_emails:
            return False
        else:
            if FAs.eh_data_valida(segundo_email):
                if dias_passados(segundo_email)  <  config.intervalo_entre_emails:
                    return False
                else:
                    if FAs.eh_data_valida(terceiro_email):
                        return False
                    else:
                        return 3
            else:
                return 2
    else:
        return 1  #### para enviar o primeiro email
    

def pega_email_e_datas_da_linha(linha , colunas_existentes, filtrar_email = False):
    valores=[]
    try:
        valores=[x.value for x in linha]
        linha = valores
    except:
        pass
    
    email = FAs.extrair_email(linha[colunas_existentes["EMAIL"]],filtro=config.filtro_email if filtrar_email else '')
    
    try:
        primeiro_email = linha[colunas_existentes['Primeiro E-MAIL ENVIADO?']]
        segundo_email  = linha[colunas_existentes['Segundo email enviado?']]
        terceiro_email = linha[colunas_existentes['Terceiro email enviado?']]
        return email , primeiro_email  ,  segundo_email  ,  terceiro_email
    except:
        print('a colunas existentes quando deu errro é:',colunas_existentes)
        print('a linha é: ',linha)

def atualiza_planilha_com_logs(sheet_list, loggerProgressManager, workbook):
    """
    Carrega as alterações persistidas anteriormente no LoggerManager para a planilha.

    Args:
        planilha: Objeto representando a planilha atual.
        loggerProgressManager: Instância do LoggerManager responsável pelo rastreamento de alterações.

    Returns:
        None
    """
    estados_salvos = loggerProgressManager.get_all_state()
    
    if not isinstance(estados_salvos, dict):
        raise ValueError("O estado retornado por get_all_state não é um dicionário válido.")
        
        # Itera pelas coordenadas e valores salvos
    tamanho=0
    for planilha_in_memory in estados_salvos:
        tamanho += len(estados_salvos[planilha_in_memory])
    print(f'começando a atualizaçao de valores ( {tamanho} ) na tabela de acordo com os dados dos logs')
    linhas_para_verificar  =  set()
    print('o estados_salvos é:'  ,  estados_salvos)
    for nome_planilha , alteracoes_planilha in estados_salvos.items():
        planilha_atual  =  sheet_list[0].sheet
        linhas_para_verificar=set()
        if nome_planilha  !=  planilha_atual.title:
            for  x in workbook.worksheets:
                if x.title  ==  nome_planilha:
                    planilha_atual  =  x   ##### aqui coloca a planilha certa com o nome que veio do logger
        for sheet_info in sheet_list:
            if sheet_info.sheet.title == nome_planilha:
                colunas_existentes  =  sheet_info.colunas_existentes  ##### aqui coloca a colunas_existentes certa com o nome que veio do logger
        for coordenadas, valor in alteracoes_planilha.items():
            if (
                    not isinstance(coordenadas, tuple) 
                    or len(coordenadas) != 2 
                    or not all(isinstance(c, int) for c in coordenadas)
                ):
                    raise ValueError(f"Coordenadas inválidas detectadas: {coordenadas}")
                
            linha, coluna = coordenadas
            letra_da_coordenada= f"{FAs.numero_para_letra_coluna(coluna)}{str(linha)}"
                  
            planilha_atual[letra_da_coordenada] = valor  # Atualiza a célula com o valor salvo
            linhas_para_verificar.add(linha)

        for linha in linhas_para_verificar:
            linha  =  planilha_atual[linha]
            email,primeiro_email,segundo_email,terceiro_email = pega_email_e_datas_da_linha(linha , colunas_existentes)
            if not condicao_enviar_email(email,primeiro_email,segundo_email,terceiro_email):
                config.atualiza_Emails_enviados(linha[colunas_existentes["EMAIL"]]) #adiciona o email referente ao valor ja salvo aos emails enviados
    print('numero de emails que ja enviei e que não é para enviar ainda: ',len(config.emails_enviados))                
    print('terminei de atualizar os valores')


def carregar_planilha():
    global config

    if not os.path.exists(config.nome_planilha):
        raise FileNotFoundError("O arquivo " + config.nome_planilha + " não foi encontrado.")
    workbook = openpyxl.load_workbook(config.nome_planilha)
    sheet_list=[]
    for sheet in workbook.worksheets:
        tiveQueCriar,colunas_existentes = ajustar_colunas(sheet)
        sheet_list.append(sheet_info(sheet,colunas_existentes,tiveQueCriar))
    
    atualiza_planilha_com_logs(sheet_list, loggerProgressManager ,workbook)
    
    return sheet_list , workbook

def ajustar_colunas(sheet):
    colunas_necessarias = ['EMAIL', 'Nome Completo', 'Primeiro E-MAIL ENVIADO?', 'Segundo email enviado?', 'Terceiro email enviado?']
    header = sheet[1]
    print(f'o header na planilha {sheet.title} é: {header}')
    colunas_existentes = [cell.value for cell in header]
    novasColunas=copy.deepcopy({})
    for index,value in enumerate(colunas_existentes):
        novasColunas[value]  =  index 

    colunas_existentes = copy.deepcopy(novasColunas)
    print(f'as colunas existentes na planilha {sheet.title} são: {colunas_existentes}')
    # Se alguma coluna necessária não existir, criá-la
    tiveQueCriar = False
    for coluna in colunas_necessarias:
        if coluna not in colunas_existentes:
            sheet.cell(row=1, column=len(colunas_existentes) + 1, value=coluna)  ##### esse +1 aqiu só pode ser usado quando for atras de coordenada na planilha do modo que o excel entende
            colunas_existentes[coluna] = len(colunas_existentes) # esse aqui não é o +1 pq é usado como indice de uma lista!
            print(f"Coluna '{coluna}' criada.")
            tiveQueCriar = True
            
    if tiveQueCriar:
        print(f'tive que criar coluna nova dentro do ajustar coluna com a sheet: {sheet.title}')
    return tiveQueCriar,colunas_existentes

def altera_e_salva(row_index,column_index,value,sheet,workbook):
    coluna_na_tabela  =  column_index + 1
    letra_da_tabela   =  FAs.numero_para_letra_coluna(coluna_na_tabela)
    linha_da_tabela   =  row_index  +  1
##    print(f'a sigla usada para acesar a celula na hora de alterar é: {letra_da_tabela}{row_index}')
    coordenada = f"{letra_da_tabela}{linha_da_tabela}"
    sheet[coordenada].value = value
    print(f'no altera e salva temos o row_index = {row_index} o column_index = {column_index} e a coordenada mudada na tabela é: {coordenada} porem o index salvo nos logs para ser pego acrescenta + 1')
    loggerProgressManager.update(row_index = linha_da_tabela, column_index = coluna_na_tabela , new_value = value,nome_planilha = sheet.title)
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
async def processar_emails(config,sheet_list,workbook,fechar_programa):

    # Verifica se a planilha foi carregada corretamente
    if sheet_list is None or workbook is None or fechar_programa:
        config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
        print("Não foi possível carregar a planilha. Encerrando o programa.")
        return  # Não prosseguir se a planilha não foi carregada corretamente

    if config.contador_emails_enviados >= config.LIMITE_DIARIO:
        config.logging.info(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado às {datetime.now()}")
        print('vou descansar por uma hora')
        await asyncio.sleep(60*60)        
        config.contador_emails_enviados=0  ## reiniciando a contagem de emails ate o próximo intervalo
        config.logging.info(f"voltei a enviar e-mail às {datetime.now()}")
    else:
        pass

    semaphore = asyncio.Semaphore(2)  # Limita para 2 e-mails sendo enviados ao mesmo tempo

    # Loop para percorrer as linhas da planilha
    for sheet_info in sheet_list:
        sheet  =  sheet_info.sheet
        sera_filtrada  =  sheet.title in config.lista_planilhas_filtrar
        print(f'a planilha {sheet.title} sofrerá filtro:{sera_filtrada}')
        colunas_existentes  =  sheet_info.colunas_existentes
        #######  os corpos dos emails vão sempre estar na primeira planilha ###########
        corpo_primeiro_email = str(sheet_list[0].sheet[2][sheet_list[0].colunas_existentes['Corpo primeiro e-mail'  ]].value   or "Mensagem padrão caso o corpo não esteja definido.") ## aqui retorna o valor do email como string
        corpo_segundo_email  = str(sheet_list[0].sheet[2][sheet_list[0].colunas_existentes['Corpo segundo e-mail'   ]].value   or "Mensagem padrão caso o corpo não esteja definido.") ## aqui retorna o valor do email como string         
        corpo_terceiro_email = str(sheet_list[0].sheet[2][sheet_list[0].colunas_existentes['Terceiro email enviado?']].value   or "Mensagem padrão caso o corpo não esteja definido.") ## aqui retorna o valor do email como string

        indice_primeiro_email        = FAs.obter_indice_coluna(colunas_existentes, 'Primeiro E-MAIL ENVIADO?') ## aqui retorna o valor do indice como int
        letra_indice_primeiro_email  = FAs.numero_para_letra_coluna(indice_primeiro_email)
        
        indice_segundo_email         = FAs.obter_indice_coluna(colunas_existentes, 'Segundo email enviado?')## aqui retorna o valor do indice como int
        letra_segundo_primeiro_email = FAs.numero_para_letra_coluna(indice_segundo_email)
        
        indice_terceiro_email  = FAs.obter_indice_coluna(colunas_existentes, 'Terceiro email enviado?')## aqui retorna o valor do indice como int
        letra_terceiro_primeiro_email  = FAs.numero_para_letra_coluna(indice_terceiro_email)                         
    
            
        row_index = 2       ###### começando na segunda linha...
        while row_index <= sheet.max_row:
    ##        try:
            if config.contador_emails_enviados >= config.LIMITE_DIARIO:
                config.logging.info(f"Limite de {config.LIMITE_DIARIO} e-mails enviados no dia alcançado às {datetime.now()}")
                print('vou descansar por uma hora')
                workbook.save(f"planilha_atualizada_{sheet.title}.xlsx")
                await asyncio.sleep(60*60)        
                config.contador_emails_enviados=0  ## reiniciando a contagem de emails ate o próximo intervalo
                config.logging.info(f"voltei a enviar e-mail às {datetime.now()}")
                print('reiniciando o trabalho')
            else:
                pass
            try:
                row   = [ x.value for x in sheet[row_index]]
            except Exception as e:
                # Mensagem personalizada
                print("Ocorreu um erro ao executar o programa! na linha 'row   = [ x.value for x in sheet[row_index]]'")
                print(f"a sheet é: {sheet}")
                print(f'o row_index é: {row_index}')
                print("Erro capturado:", e)  # Exibe o erro original
                raise 
                
            nome  = str( row[colunas_existentes['Nome Completo']] or '')
            email, primeiro_email, segundo_email, terceiro_email = pega_email_e_datas_da_linha(row,colunas_existentes,filtrar_email = sera_filtrada)
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
                
                    
                
                # Verifica se o e-mail ainda não foi enviado
            condicao= condicao_enviar_email(email,primeiro_email,segundo_email,terceiro_email)
            if not condicao:
    ##                print(f'não é para enviar email para: {email} com os valores primeiro_email: {primeiro_email}, segundo_email: {segundo_email}, terceiro_email: {terceiro_email}')
                row_index  +=  1
                continue
            if condicao == 1:
                print(f'vou mandar o email: {email} como primeiro email, mas suspendi')
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
        print('planilha concluida')
        workbook.save(f"planilha_atualizada_{sheet.title}.xlsx")
    print('Processamento de e-mails concluído!')

sheet_list , workbook = carregar_planilha()
fechar_programa = False
##if tiveQueCriar:
##    print('tive que criar coluna nova la no fim do codigo')
##    tiveQueCriar , colunas_existentes , sheet , workbook = carregar_planilha()
##    if sheet is None or workbook is None:
##        fechar_programa = True
##        config.logging.error("Não foi possível carregar a planilha, o processo será encerrado.")
##        print("Não foi possível carregar a planilha. Encerrando o programa.")
##async def iniciar():
##    await processar_emails()

# Função principal para rodar o código
async def main():
    await processar_emails(config,sheet_list,workbook,fechar_programa)

if __name__ == '__main__':
    asyncio.run(main())
