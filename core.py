import warnings
from cryptography.utils import CryptographyDeprecationWarning
from commons import __get_parameters
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore", category=CryptographyDeprecationWarning)
import logging
from commons import *
from bot_base import *
import pandas as pd
from retrying import retry
from wrapt_timeout_decorator import *
from datetime import timedelta
from playwright.sync_api import sync_playwright, expect
from bs4 import BeautifulSoup
import urllib.parse
from lxml import etree
import pandas as pd
from openpyxl import load_workbook
from http import HTTPStatus
from tinydb import TinyDB, Query
import humanize
import base64
import ssl
import certifi
from collections import Counter
from features.update_protocols_mobile_saude.index import update_protocols_mobile_saude
from features.read_protocols_mobile_saude.index import read_protocols_mobile_saude, export_backlog
from features.create_peg_benner.index import create_peg_benner

colunas = {valor: indice for indice, valor in enumerate(pd.read_excel("./data/parameters.xlsx", engine='openpyxl', sheet_name="ids", header=None).iloc[:, 0].tolist())}


@retry(retry_on_exception=retry_if_stop_exception, 
       wait_fixed=int(config["restart"]["wait_fixed"]), 
       stop_max_attempt_number=int(config["restart"]["stop_max_attempt_number"]))  
@timeout(int(config["restart"]["max_execution_time"]), timeout_exception=StopIteration, dec_poll_subprocess=2)
@timeit
def run() -> int: 

  
    try:

        # Criar um contexto SSL que ignora a verificação de certificado
        ssl_context = ssl.create_default_context()
        ssl_context.check_hostname = False
        ssl_context.verify_mode = ssl.CERT_NONE

        os.environ['SSL_CERT_FILE'] = certifi.where()            

        bot_base()

        banner()

        logging.info('Iniciando a execução da automação...')

        for key, value in global_parameters.items(): logging.info(f"{key}: {value}")            

        delete_temp_files()
                
        reload_old_files()

        logging.info(f'log {log}')

        protocols_total = 0
        protocols_success = 0        

        if bool(global_parameters["mobilesaude.read_protocols_mobile_saude"]):
            worksheet = read_protocols_mobile_saude() 
        else:
            worksheet = global_parameters["mobilesaude.worksheet"]
                
        if bool(global_parameters["benner.create_peg_benner"]) and worksheet:
            protocols_total, protocols_success = create_peg_benner(worksheet)

        if bool(global_parameters["mobilesaude.update_protocols_mobile_saude"]) and worksheet:
            protocols_total, protocols_success = update_protocols_mobile_saude(worksheet)

        export_backlog()

        endTime = timer()
        tempo_medio = ((endTime-startTime)/protocols_total) if protocols_total else 0
        saving = float(global_parameters["saving"])

        sumary = [
            "Protocolos processados..............: " + str(protocols_total),
            "Protocolos processados com sucesso..: " + str(protocols_success),
            "Protocolos processados com erro.....: " + str(protocols_total-protocols_success),
            "Tempo médio por protocolo...........: " + humanize.precisedelta(tempo_medio),
            "Tempo de processamento..............: " + humanize.precisedelta(timedelta(seconds=endTime-startTime)),
            "Tempo economizado...................: " + humanize.precisedelta(protocols_total * (saving - tempo_medio))
        ]

        for x in sumary: logging.info(x)

        finish()

        logging.info("#### Finalizado ####")     

        send_emails(sumary)

        return 0

    except:
        show_exception_and_exit(*sys.exc_info())    