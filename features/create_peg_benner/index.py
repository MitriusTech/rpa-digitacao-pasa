import logging
from utils.global_parameters import  global_parameters, __get_parameters
from utils.config import config
from bot_base import *
from commons import *
from playwright.sync_api import sync_playwright
import traceback

@timeit
def create_peg_benner(worksheet) -> list:
 
    logging.info(f'Lançando PEGs no Benner...')

    protocols_total = 0
    protocols_count = 0
    protocols_success = 0          

    #
    refund_types = __get_parameters(sheet_name="refund_types", key="from", value="to")

    with sync_playwright() as playwright:

        @timeit
        @handle_exceptions(default_return=False)
        @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
        def exists(sc, label, value, dependValueList="") -> bool:  
            return len(search(sc, label, value, dependValueList)) > 0
    
        @timeit
        @handle_exceptions(default_return=[])
        @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
        def search(sc, label, value, dependValueList="") -> list:  

            logging.info(f'Pesquisando {value} em {label}...')       

            payload = {
                "query": str(value),
                "sc": sc,
                "dependValueList": dependValueList,
                "maxRows": 1,
                "startRow": 0,
                "fieldsJson": ""
            }

            # Usar o contexto com sessão ativa
            request_context = context.request

            response = request_context.post(
                global_parameters[f'benner.api_search_{global_parameters["env"]}'],
                form=payload,
                ignore_https_errors=True
            )                

            # check if the request was successful
            if response.status != HTTPStatus.OK:
                logging.error(f"{response.text()}")  
                return []

            result = json.loads(response.text())

            if len(result) == 0:
                logging.error(f'{value} não encontrado')  
                return []

            return result
    
        @timeit
        def add_peg() -> bool:     

            @timeit
            def fill_guide() -> str:

                nonlocal guide_number, executor_sc, endereco_executor_sc

                @timeit
                def fill_principal_guide():

                    nonlocal guide_number

                    logging.info(f'Preenchendo guia principal...')
                    
                    # preencher o número da Ordem
                    safe_locator(page,"xpath=//*[@data-label='Ordem']//input[@type='text']").first.fill(row[colunas.get("refund_qty")].value)

                    guide_number = safe_locator(page,"xpath=//*[@data-label='N&#250;mero da guia']//input[@type='text']").first.input_value().strip()
                    
                    # copiar o número da guia para Número da Guia Principal
                    safe_locator(page,"xpath=//*[@data-label='N&#250;mero da Guia Principal']//input[@type='text']").first.fill(guide_number)

                    # copiar o número da guia para Nr Guia Solicitação
                    safe_locator(page,"xpath=//*[@data-label='Nr Guia Solicita&#231;&#227;o']//input[@type='text']").first.fill(guide_number)

                    # copiar o número da guia para No Guia Prestador
                    safe_locator(page,"xpath=//*[@data-label='N&#186; Guia Prestador']//input[@type='text']").first.fill(guide_number)

                @timeit
                def fill_data_guide() -> str:

                    @timeit                 
                    def set_executor_address() -> str:        

                        element = page.locator('select[data-fieldname="ENDERECOEXECUTOR"]') 
                        
                        page.evaluate("""
                            (id) => {
                                Benner.Apps.CustomLookup.showDialog(id);return false;
                            }  
                        """, element.get_attribute("id"))   

                        wait_for_load_state(page, 1000)               
                        
                        # Aguardar carregamento do iframe (opcional, se necessário)
                        page.wait_for_selector("iframe")

                        # Acessar o primeiro iframe da página
                        frames = page.frames
                        frame = frames[1]

                        if not is_element_ready(frame, "#Resultado_SimpleGrid > tbody > tr", timeout=int(global_parameters["timeout"])*2):
                            return f'Nenhum endereço encontrado para o prestador'                        

                        first_row = frame.locator("#Resultado_SimpleGrid > tbody > tr").first
                        first_row.click()

                        return None

                    nonlocal executor_sc, endereco_executor_sc

                    logging.info(f'Criando guia dados da guia...')

                    supplier_id = formatar_cpf_cnpj(row[colunas.get("supplier_id")].value)

                    # clicar na aba Dados da Guia
                    safe_locator(page,'//a[@data-toggle="tab" and normalize-space(text())="Dados da Guia"]').click()

                    wait_for_load_state(page)

                    # recupera o data-searchcontext do EXECUTOR para facilitar consultas futuras
                    if not executor_sc:
                        locator = safe_locator(page, 'select[data-fieldname="EXECUTOR"]')
                        if locator.count():
                            executor_sc = locator.get_attribute('data-searchcontext')       

                    if not endereco_executor_sc:
                        locator = safe_locator(page, 'select[data-fieldname="ENDERECOEXECUTOR"]')
                        if locator.count():
                            endereco_executor_sc = locator.get_attribute('data-searchcontext')                                                     

                    # Geral > Beneficiário
                    if not fill_select2(page, 'select[data-fieldname="BENEFICIARIO"]', row[colunas.get("user")].value):
                        return "Beneficiário não encontrado ou inválido"

                    # Geral > Executor
                    if not fill_select2(page, 'select[data-fieldname="EXECUTOR"]', supplier_id):
                        return "Prestador não encontrado ou inválido"
                    
                    # Geral > Endereço executor
                    if not fill_select2_by_index(page, 'select[data-fieldname="ENDERECOEXECUTOR"]', 1):
                        if not fill_select2(page, 'select[data-fieldname="ENDERECOEXECUTOR"]', ""):
                            if (error_message := set_executor_address()):
                                return error_message                    

                    return None

                @timeit
                def fill_service_characteristics_guide():

                    logging.info(f'Preenchendo guia características do atendimento...')

                    # clicar na aba Dados da Guia
                    safe_locator(page,'//a[@data-toggle="tab" and normalize-space(text())="Características do Atendimento"]').click()

                    wait_for_load_state(page)

                    # Finalidade do atendimento = Rotina
                    fill_select2(page, 'select[data-fieldname="FINALIDADEATENDIMENTO"]', global_parameters["benner.finalidade_atendimento"]) 

                    # Local atendimento = Rede Livre Escolha
                    selector = f'input[value="{global_parameters["benner.local_atendimento"]}"]'
                    if not is_element_ready(page, selector, printscreen=False):
                        fill_select2(page, 'select[data-fieldname="LOCALATENDIMENTO"]', global_parameters["benner.local_atendimento"]) 

                    # selector = f'input[value="{global_parameters["benner.local_atendimento"]}"]'
                    # input_element = safe_locator(page,selector)

                    # if not input_element.count() > 0 and input_element.first.is_visible():
                    #     fill_select2(page, 'select[data-fieldname="LOCALATENDIMENTO"]', global_parameters["benner.local_atendimento"]) 

                    # Regime de atendimento = Ambulatorial
                    fill_select2(page, 'select[data-fieldname="REGIMEATENDIMENTO"]', global_parameters["benner.regime_atendimento"]) 

                    # Condição de atendimento = Eletivo
                    fill_select2(page, 'select[data-fieldname="CONDICAOATENDIMENTO"]', global_parameters["benner.condicao_atendimento"]) 

                    # Tipo do tratamento
                    fill_select2(page, 'select[data-fieldname="TIPOTRATAMENTO"]', global_parameters["benner.tipo_tratamento"]) 

                    # Objetivo do tratamento
                    fill_select2(page, 'select[data-fieldname="OBJETIVOTRATAMENTO"]', global_parameters["benner.objetivo_tratamento"]) 
               
                logging.info(f'Criando guia...')
                
                # cliar em GUIAS
                safe_locator(page,f'xpath=//div/span[normalize-space(text())="Guias"]').first.click()   

                wait_for_load_state(page)   

                # clica em +Novo
                with page.expect_navigation():
                    safe_locator(page,"//div[contains(@class, 'portlet light')][.//span[normalize-space(text())='Guias']]//a[normalize-space(text())='Novo']").click()

                error_message = get_error_message()
                if error_message:
                    return f'Não foi possível criar uma nova guia. ' + re.sub('\s+',' ', error_message)

                steps = [
                    (fill_principal_guide, ()),
                    (fill_data_guide, ()),
                    (fill_service_characteristics_guide, ())
                ]

                for step_func, args in steps:
                    
                    error_message = step_func(*args)
                    
                    if error_message is not None:
                        return error_message 

                # Salvar
                page.keyboard.press("Control+Enter")

                wait_for_load_state(page, 1000)

                error_message = get_error_message()
                if error_message:
                    return f'Não foi possível salvar a guia. ' + re.sub('\s+',' ', error_message)

                logging.info(f"Guia {guide_number} lançada com sucesso")

                return None

            @timeit
            def fill_peg() -> str:

                logging.info(f'Preenchendo PEG...')

                nonlocal peg, client_id
                expense_nf = ''.join(filter(str.isdigit, row[colunas.get("expense_nf")].value or ""))
                expense_date = formatar_data(row[colunas.get("expense_date")].value or "")

                result = navigate_postback_using_form_data(page, form_data, action_url, "ctl00$Main$WDG_V_FILIAIS_779_LNK1", "New")

                if not result["success"]:
                    return f'Falha ao tentar criar o PEG: {result.get("error")}'

                # Selecione "Em Digitação"

                # xpath = "//a[@title='Ctrl + Insert']"

                # if not safe_locator(page, xpath).is_visible():
                #     safe_locator(page,f'//div/span[normalize-space(text())="Em Digitação"]').click()
                #     wait_for_load_state(page)               

                # if not is_element_ready(page, xpath):
                #     return f'Botão "Ctrl + Insert" não está visível'             
                
                # with page.expect_navigation():
                #     safe_locator(page, xpath).first.click()      

                client_id = get_client_id(page)

                if not client_id:
                    return f'client_id não encontrado no corpo da página'   

                # Tipo de PEG = Reembolso
                safe_locator(page,"xpath=//label[normalize-space(text())='Reembolso']").first.click() 

                # Data de recebimento - protocol_date
                safe_locator(page,"xpath=//div[@data-label='Recebimento']//input[@type='text']").first.fill(formatar_data(row[colunas.get("protocol_date")].value))
                
                # Despesa - expense_nf
                if expense_nf:
                    input_locator = safe_locator(page,"xpath=//span[@data-label='N&#250;mero da nota']//input[@type='text']") 
                    input_locator.focus()
                    for letra in expense_nf:
                        input_locator.press(letra)
                    input_locator.press("Tab")

                    # Data de emissão da NF = expense_date
                    safe_locator(page,"xpath=//div[@data-label='Data Emiss&#227;o Nota']//input[@type='text']").first.fill(expense_date)

                # Quantidade de Guias – Apresentada = refund_qty
                page.evaluate('''
                    (value) => {
                        $("span[data-field='QTDGUIA'] input[type='text']").val(value).trigger("change");
                    }
                ''', row[colunas.get("refund_qty")].value)

                # Tipo do PEG
                if not row[colunas.get("refund_type")].value in refund_types:
                    peg_type = global_parameters["benner.default_peg_type"]
                else:
                    peg_type = refund_types[row[colunas.get("refund_type")].value]

                fill_select2(page, 'select[data-fieldname="TIPOPEG"]', peg_type) 

                # Beneficiário Titular 
                fill_select2(page, 'select[data-fieldname="BENEFICIARIO"]', row[colunas.get( global_parameters["benner.chave_beneficiario"] )].value) 

                # Gravando o protocolo na observação
                safe_locator(page,'//a[@data-toggle="tab" and normalize-space(text())="Observações"]').click()
                observacao_interna = safe_locator(page,'div:has(.label-title:has-text("Observação")) textarea').first
                observacao_interna.fill(f'Protocolo mobilesaude:[{row[colunas.get("protocol_id")].value}]\n{observacao_interna.input_value()}')

                # volta para a 1a aba
                safe_locator(page,'//a[@data-toggle="tab" and normalize-space(text())="Principal"]').click()
                
                # Salvar
                page.keyboard.press("Control+Enter")

                wait_for_load_state(page, 1000)

                error_message = get_error_message()
                if error_message:
                    return f'Não foi possível salvar o formulário. ' + re.sub('\s+',' ', error_message)

                peg = get_peg()

                logging.info(f'PEG {peg} lançado com sucesso')
                
                return None

            def valid_required_fields() -> str:

                logging.info(f'Validando campos obrigatórios...')

                supplier_id = formatar_cpf_cnpj(row[colunas.get("supplier_id")].value)
                isCPF = is_CPF(supplier_id)
                isCNPJ = is_CNPJ(supplier_id)
                expense_nf = ''.join(filter(str.isdigit, row[colunas.get("expense_nf")].value or ""))
                expense_date = formatar_data(row[colunas.get("expense_date")].value or "")

                if not row[colunas.get("user")].value:
                    return "Nome do beneficiário não informado"

                # código do prestador deve ser um CPF ou CNPJ válido
                if not isCPF and not isCNPJ:
                    return "Código do prestador inválido"
                
                # Data da NF é obrigatório para prestador PJ
                if isCNPJ and not expense_date:
                    return "Data da NF não informada ou inválida"                  

                # Número da NF é obrigatório para prestador PJ
                if isCNPJ and not expense_nf:
                    return "NF não informada ou inválida"             
                
                # Número da NF deve estar acompanhada da data
                if (expense_nf and not expense_date):
                    return "Número da NF deve estar acompanhada da data"  
                
                # Verifica se o prestador existe na base
                if executor_sc:

                    result = search(executor_sc, "prestador", supplier_id)

                    if not result:
                        return 'Prestador não encontrado ou inválido'
                    else:
                        if endereco_executor_sc:
                            if not (
                                    exists(endereco_executor_sc, "endereço prestador", "", f'EXECUTOR={result[0]["id"]}') or
                                    exists(endereco_executor_sc, "endereço prestador", " ", f'EXECUTOR={result[0]["id"]}')
                                ):
                                return 'Endereço do prestador não encontrado ou inválido'

                return None                   

            def get_peg() -> str:

                locator = safe_locator(page,"//span[@data-field='PEG']//xmp")
                if locator.count() == 0:
                    return None
                    
                return locator.first.inner_text().strip()
                
            def get_error_message() -> str:

                locator = safe_locator(page,"xpath=//div[contains(@class, 'message-panel')]//span[last()]")
                if locator.count() == 0:
                    return None
                
                error_message = str(re.sub('\s+',' ',locator.first.inner_text()).strip()) 

                if "porém a gravação do registro será permitida" in error_message:
                    logging.info(error_message)
                    return None 
                
                if "O Beneficiário Titular do PEG é também o titular da família do beneficiário desta guia!" in error_message:
                    logging.info(error_message)
                    return None                 

                return error_message            
    
            @timeit
            @handle_exceptions(default_return=False)
            @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)            
            def delete_peg(peg) -> bool:

                if not global_parameters["benner.delete_peg"] or not peg:
                    return False

                logging.info(f'Deletando o peg {peg}...')

                if not is_element_ready(page,'#breadcrumbUpdatePanel'):
                    logging.error(f'breadcrumb não encontrado')            
                    return False

                with page.expect_navigation():
                    safe_locator(page,'#breadcrumbUpdatePanel a', has_text=peg.replace(".","")).click()

                page.evaluate("""
                    (client_id) => {
                        __doPostBack('ctl00$Main$WDG_V_SAM_PEG_'+client_id+'_PRINCIPAL','CMD_EXCLUIRPEG');
                    }
                """, client_id)   

                wait_for_load_state(page)      # javascript:__doPostBack('ctl00$Main$WDG_V_SAM_PEG_1465_PRINCIPAL$RQTrue','')

                if not is_element_ready(page,'input[type="button"][value="Continuar"]'):
                    logging.error(f'botão "Continuar" não encontrado')
                    return False          
                
                with page.expect_navigation():
                    safe_locator(page,'input[type="button"][value="Continuar"]').click()

                wait_for_load_state(page)

                message = None
                xpath = '#ctl00_TopBarMainContent_globalMessagePanel_message'

                if is_element_ready(page, xpath):
                    message = safe_locator(page, xpath).inner_text()
                    logging.info(message)                 
                
                return message == "Processo enviado para execução no servidor! Peg sendo excluído!"

            nonlocal executor_sc, endereco_executor_sc
            error_message = None
            guide_number = None
            peg = None
            client_id = None
            
            try:

                logging.info(f'Criando PEG...')    

                # preenche os dados do PEG
                guide_number = None
                peg = None
                client_id = None
                steps = [valid_required_fields, fill_peg, fill_guide]

                for step in steps:
                    error_message = step()
                    if error_message:
                        logging.error(f"Erro na etapa '{step.__name__}': {error_message}")
                        break

            except Exception as e:

                page.screenshot(path=f'{log}\\screenshot_{today}_{today_.strftime("%H%M%S")}.png', full_page=True)
                error_message = str(e)
                logging.error(traceback.format_exc())
            
            finally:

                # atualizar Excel
                logging.info("atualizando excel...")

                row[colunas.get("PEG")].value = None if error_message else peg
                row[colunas.get("comment")].value = error_message
                row[colunas.get("guide_number")].value = guide_number

                wb.save(worksheet)     

                # exclui o peg neste caso
                if error_message and peg:
                    delete_peg(peg)             

                # # Selecione a filial
                # with page.expect_navigation():
                #     safe_locator(page,'#breadcrumbUpdatePanel a', has_text=global_parameters["benner.filial"]).click()

            return not error_message

        @timeit
        @handle_exceptions(default_return=False)
        @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
        def login() -> bool:

            with page.expect_navigation():
                page.goto(global_parameters[f'benner.url_{global_parameters["env"]}'])      

            safe_locator(page,"input[name='wesLogin$loginWes$UserName']").first.fill(config["benner"]["username"])
            safe_locator(page,"input[name='wesLogin$loginWes$Password']").first.fill(config["benner"]["password"])
            
            with page.expect_navigation():
                safe_locator(page,"//*[@id='LoginButton']").click()

            logging.info(f'Login benner com sucesso {config["benner"]["username"]}...')

            # Competências de PEG
            with page.expect_navigation():
                page.goto(global_parameters[f'benner.competenciasdepeg_{global_parameters["env"]}'])

            # Selecione a competência
            with page.expect_navigation():
                safe_locator(page,f'//a[text()={global_parameters["benner.competencia"]}]').click()

            # cliar em FILIAIS
            safe_locator(page,f'//div/span[@title="FILIAIS"]').first.click()
            
            # Selecione a filial
            with page.expect_navigation():
                safe_locator(page,f'//a[text()="{global_parameters["benner.filial"]}"]').click()  

            logging.info(f'Seleção da filial {global_parameters["benner.filial"]} com sucesso')

            return True          

        @timeit
        @handle_exceptions(default_return=None)
        @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
        def goto_filial() -> tuple[dict, str]:

            # Competências de PEG
            with page.expect_navigation():
                page.goto(global_parameters[f'benner.competenciasdepeg_{global_parameters["env"]}'])

            # Selecione a competência
            with page.expect_navigation():
                safe_locator(page,f'//a[text()={global_parameters["benner.competencia"]}]').click()

            # cliar em FILIAIS
            safe_locator(page,f'//div/span[@title="FILIAIS"]').first.click()
            
            # Selecione a filial
            with page.expect_navigation():
                safe_locator(page,f'//a[text()="{global_parameters["benner.filial"]}"]').click()  

            # 1. Captura o form e a URL correta
            form_data, action_url = capture_aspnet_form(page)       

            logging.info(f'Seleção da filial {global_parameters["benner.filial"]} com sucesso')

            return form_data, action_url               
        
        logging.info(f'Acessando o site...')
        logging.info(global_parameters[f'benner.url_{global_parameters["env"]}'])
        logging.info(f'Usuário {config["benner"]["username"]}')

        browser = playwright.chromium.launch(
            headless=bool(global_parameters["benner.headless"]), 
            args=["--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage"]
        )

        if bool(global_parameters["benner.record_video"]):
            context = browser.new_context(record_video_dir=log)
        else:
            context = browser.new_context()

        page = context.new_page()

        # Define o timeout padrão (em milissegundos)
        page.set_default_timeout(int(global_parameters["timeout"]))
        page.set_default_navigation_timeout(int(global_parameters["timeout"]))        
        
        page.route("**/*", lambda route: route.abort() if route.request.resource_type in ["image", "font"] else route.continue_()) #, "stylesheet", "font"
         
        #
        wb = safely_load_workbook(worksheet)

        if not wb:
            return 0, 0 

        ws = wb.active  # Pega a primeira planilha ativa    

        # Criar um dicionário para mapear os nomes das colunas aos seus índices
        #colunas = {cell.value: cell.column for cell in ws[1]}  # Pega a linha 1 como cabeçalho
        protocols_count = 0
        protocols_total = ws.max_row -1
        executor_sc = None
        endereco_executor_sc = None

        if not login():
            return protocols_total, 0 
        
        form_data, action_url = goto_filial()

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):

            protocols_count += 1

            if protocols_count % int(global_parameters["benner.bloco_login"]) == 0:
                if not login():
                    logging.error(f'Criação dos PEGs interrompida no protocolo {protocols_count}/{protocols_total}')
                    break            
                form_data, action_url = goto_filial()
        
            logging.info(f'Processando protocolo {row[colunas.get("protocol_id")].value} {protocols_count}/{protocols_total}...')
            
            if row[colunas.get("comment")].value:
                logging.warning(row[colunas.get("comment")].value)   

            if not row[colunas.get("PEG")].value:
                protocols_success += (1 if add_peg() else 0)
            else:
                logging.info(f'PEG {row[colunas.get("PEG")].value} criado anteriormente')     

        # 
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
           
            if not row[colunas.get("comment")].value and not row[colunas.get("PEG")].value:
                row[colunas.get("comment")].value = f'Processamento interrompido'  
                logging.warning(f'Processamento interrompido do protocolo {row[colunas.get("protocol_id")].value}')

        if bool(global_parameters["benner.record_video"]):
            logging.info(f'vídeo gravado em {page.video.path()}')

        context.close()

        wb.save(worksheet)
    
    return protocols_total, protocols_success   
        