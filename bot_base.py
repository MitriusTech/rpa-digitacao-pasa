import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import logging
import sys
from utils.config import config
from utils.global_parameters import global_parameters
from utils.globals import *
import commons
import urllib3
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
from mss import mss
import shutil
from wrapt_timeout_decorator import *

commons.pd

colunas = {valor: indice for indice, valor in enumerate(commons.pd.read_excel(
    "./data/parameters.xlsx", engine='openpyxl', sheet_name="ids", header=None).iloc[:, 0].tolist())}


def screenShot() -> str:
    screenshot = f'{log}\\{datetime.today().strftime("%H%M%S")}.png'

    with mss() as sct:
        sct.shot(output=screenshot)

    return screenshot

def remove_process_folder():
    
    target = today_ - relativedelta(days=int(config["commons"]["retention"]))
    targetFormatted = target.strftime('%Y%m%d')
    folders = [ f.path for f in os.scandir(f'{os.getcwd()}/log') if f.is_dir() ]

    logging.warning("removendo pastas antigas...")

    for folder in folders:
        folder = os.path.basename(folder)
        targetFolder = folder[:8]

        if targetFolder.isnumeric() > 0:
            if targetFolder < targetFormatted:
                logging.warning(f'removendo pasta {targetFolder}...')
                shutil.rmtree(f'{os.getcwd()}/log/{folder}')

def show_exception_and_exit(exc_type, exc_value, tb):
   
    logging.error(exc_value, exc_info=(exc_type, exc_value, tb))
    
    if not bool(config["smtp"]["enabled"]):
        return
    
    screenShot()
    
    with open(config["smtp"]["template"], 'r', encoding=config["commons"]["encoding"]) as template:
        
        html = template.read().format(
            "Ocorreram problemas durante o lançamento dos PEGs",
            "Ocorreram problemas durante o lançamento dos PEGs. Todos os detalhes do processamento estão no log em anexo.",
            exc_value,
            "",
            "lançamento PEG"
            )
    
    logging.info("Enviando e-mail de erro...")
    logging.info(config["emails"]["error"] + "," + str(global_parameters["emails_error"]))
    logging.info("#### Finalizado ####")   
    logging.shutdown()

    commons.sendemail_postmarkapp(config["smtp"]["host"], 
              config["smtp"]["port"], 
              config["smtp"]["username"], 
              config["smtp"]["password"], 
              config["smtp"]["headers"], 
              f'{config["smtp"]["subject"]} - AGENDA {todayFormatted} {today_.strftime("%H:%M")}',
              config["smtp"]["from"], 
              config["emails"]["error"] + "," + str(global_parameters["emails_error"]),
              html,
              [config["smtp"]["logo"]],
              [f"{log}/{arquivo}" for arquivo in os.listdir(log)])
    
def bot_base():

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    sys.excepthook = show_exception_and_exit

    if os.path.exists(path) == False:
        os.makedirs(path)
    
    if os.path.exists(temp) == False:
        os.makedirs(temp)

    if os.path.exists(log) == False:
        os.makedirs(log)    

    # Log
    logging.root.handlers = []
    logging.basicConfig(level=int(config['commons']['log_level']), format="%(asctime)s; %(levelname)s; %(module)s.%(funcName)s.%(lineno)d; %(message)s", 
        handlers=[ 
            logging.FileHandler(logFileFullPath, mode='w', encoding = config['commons']['encoding'], delay=False), 
            logging.StreamHandler(sys.stdout) 
        ])
    logging.getLogger().setLevel(int(config['commons']['log_level']))

    commons.close_excel() 

    remove_process_folder()

    commons.delete_temp_files()

def getConfig():
    return config

def formatar_data(valor):
    if isinstance(valor, datetime):
        return valor.strftime('%d/%m/%Y')
    return valor  # Retorna o valor original se não for datetime

def get_client_id(page) -> str:
    return page.evaluate("""
        () => {
        const el = document.querySelector('[id^="ctl00_Main_WDG_V_SAM_PEG_"]');
        if (!el) return null;
        const match = el.id.match(/^ctl00_Main_WDG_V_SAM_PEG_(\\d+)_/);
        return match ? match[1] : null;
        }
        """)

def safe_locator(page, *args, **kwargs):

    try:
        locator = page.locator(*args, **kwargs)
        return locator
    except (TimeoutError, Exception) as e:
        logging.error(f"erro ao localizar elemento: {e}")
        page.screenshot(path=f'{log}\\screenshot_{today}_{today_.strftime("%H%M%S")}.png', full_page=True)                
    return None    

def is_element_ready(page, selector, check="visible", timeout=6000, printscreen=True) -> bool:
    """
    Verifica se o elemento está visível, anexado ou habilitado.

    :param page: Objeto da página do Playwright.
    :param selector: Seletor CSS do elemento (string).
    :param check: Tipo de verificação: "visible", "attached", "enabled".
    :param timeout: Tempo de espera em milissegundos.
    :return: True se o elemento está conforme esperado, senão uma string com a mensagem de erro.
    """

    try:
        locator = page.locator(selector)

        if check in ["visible", "attached", "hidden", "detached"]:
            page.wait_for_selector(selector, state=check, timeout=timeout)
            return True

        elif check == "enabled":
            # Verifica se está visível primeiro (evita erro de lookup invisível)
            page.wait_for_selector(selector, state="visible", timeout=timeout)
            if locator.is_enabled():
                return True
            else:
                logging.error(f"Elemento '{selector}' está desabilitado.")
        else:
            logging.error(f"Tipo de verificação '{check}' não é suportado.")

    except Exception as e:
        logging.error(f"Erro ao verificar '{selector}' com check='{check}': {str(e)}")

    if printscreen:
        page.screenshot(path=f'{log}\\screenshot_{today}_{today_.strftime("%H%M%S")}.png', full_page=True)                

    return False      


def send_emails(sumary):
        
        if not bool(config["smtp"]["enabled"]):
            return

        with open(config["smtp"]["template"], 'r', encoding=config["commons"]["encoding"]) as template:
            
            html = template.read().format(
                "A tarefa de lançamento de PEG foi realizada com sucesso",
                "A tarefa de lançamento de PEG foi realizada com sucesso. Todos os detalhes do processamento estão no log em anexo.",
                "<p style=\"font-family:'Courier New'\">" + "<br/>".join(sumary) + "</p>",
                "",
                "lançamento PEG"
                )
        
        logging.info("Enviando e-mail...")
        logging.info(config["emails"]["success"] + "," + str(global_parameters["emails_success"]))
        logging.shutdown()

        commons.sendemail_postmarkapp(config["smtp"]["host"], 
                config["smtp"]["port"], 
                config["smtp"]["username"], 
                config["smtp"]["password"], 
                config["smtp"]["headers"], 
                f'{config["smtp"]["subject"]} - AGENDA {todayFormatted} {today_.strftime("%H:%M")}', 
                config["smtp"]["from"], 
                config["emails"]["success"] + "," + str(global_parameters["emails_success"]), 
                html,
                [config["smtp"]["logo"]],
                [f"{log}/{arquivo}" for arquivo in os.listdir(log)]) 
        