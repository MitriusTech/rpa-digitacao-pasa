import logging
from packages.global_parameters import global_parameters
from packages.config import config
from packages.commons import *
from packages.bot_base import *
import ast


@timeit
def update_protocols_mobile_saude(worksheet) -> list:

    logging.info(f'Atualizando protocolos no mobilesaude...')

    protocols_total = 0
    protocols_count = 0
    protocols_success = 0

    env = global_parameters["env"]
    bot_username = global_parameters["mobilesaude.bot_user_name"]
    session = requests.session()

    @timeit
    @handle_exceptions(default_return=False)
    @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
    def assign(new_attendant: str, occurrence: dict, idStatus: int) -> bool:
        if not bool(global_parameters["mobilesaude.assign"]):
            logging.info("Atribuição de usuário está desativada.")
            return None

        logging.info(
            f"Atribuindo protocolo para o(a) atendente {new_attendant}.")

        attendants = get_attendance_data(session)

        if not attendants:
            logging.error("Não foi possível buscar lista de atendentes")
            return None

        found_user = False

        for attendant in attendants:
            name = attendant.get("nome")
            if name == new_attendant:
                idAttendant = attendant.get("id_atendente")
                url = global_parameters["mobilesaude.change_occurrence"]

                payload = {
                    "ocorrencias": [occurrence],
                    "alteraAtendente": {
                        "idAtendente": idAttendant
                    },
                    "updateStatus": {
                        "desfecho": None,
                        "motivo": None,
                        "arquivos": [],
                        "idStatus": idStatus
                    }
                }
                response = safe_post(session, url, json=payload)

                if not response.ok or not response.json().get("status"):
                    return False

                logging.info(f"Protocolo {protocol_id} atribuído a {name}.")
                found_user = True
                break

        if not found_user:
            logging.error(f"Usuário {new_attendant} não encontrado.")
            return None
        
        return True

    @timeit
    @handle_exceptions(default_return=None)
    @retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
    def assign_to_bot(idStatus:int) -> bool:
        url = global_parameters["mobilesaude.assigned_to_user"] % refund_id
        payload = {
            "idStatus": idStatus
        }

        response = safe_put(session, url, json=payload)

        if not response.ok or not response.json().get("status"):
            return False

        logging.info(
            f'Protocolo {refund_id} atribuído com sucesso ao bot {bot_username}.')
        return True

    def update_protocol(change_status):
        logging.info(f'Atualizando protocolo {protocol_id}')

        buffer = ""
        PEG = ""
        
        if PEG or status_id != global_parameters["mobilesaude.requested_status"]:
            row[colunas.get("assigned")].value = assigned
            row[colunas.get(
                "complement")].value = f'Protocolo já estava atualizado para {status_desc}{" por " + assigned if assigned else ""}{" com o PEG " + PEG if PEG else ""}'

            logging.info(row[colunas.get("complement")].value)
            return None


        if assigned and assigned != bot_username or not assigned:
            if not assign_to_bot(global_parameters["mobilesaude.requested_status"]):
                return f"Não foi possível atribuir o protocolo {protocol_id} ao bot {bot_username}."
            
        row[colunas.get("assigned")].value = bot_username
            
        service_tabs = occurrence_data.get("abasAtendimento", [])
        internal_observation_url = {
            "url": "",
            "exists": False
        }

        if not len(service_tabs):
            return f'Não foi possível obter as abas de atendimento da ocorrência {refund_id}'

        for tab in service_tabs:
            tab_name = tab.get("nome", "").lower()
            if tab_name == "pagamento":
                id_form = tab.get("id_form")
                id_form_data = tab.get("formDatas")

                if not id_form_data:
                    return f'Não foi possível obter os dados do formulário de pagamento da ocorrência {refund_id}'

                url = global_parameters["mobilesaude.form_data"] % (
                    id_form, id_form_data[0])
                response = safe_get(session, url)

                if not response.ok or not response.json().get("status"):
                    return f'Não foi possível obter os dados do pagamento da ocorrência {refund_id}'

                response_json = response.json().get("data").get("data")

                PEG = response_json.get("numeroLote")
                row[colunas.get("PEG")].value = PEG

            if tab_name == "observações internas":
                id_form = tab.get("id_form")
                id_form_data = tab.get("formDatas")

                if not id_form_data:
                    internal_observation_url[
                        "url"] = global_parameters["mobilesaude.create_payment_data"] % id_form
                    internal_observation_url["exists"] = False
                else:
                    current_id_form_data = id_form_data[0]
                    internal_observation_url["url"] = global_parameters["mobilesaude.form_data"] % (
                        id_form, current_id_form_data)
                    internal_observation_url["exists"] = True

        if row[colunas.get("PEG")].value:
            buffer += f'PEG:{row[colunas.get("PEG")].value}'

        if row[colunas.get("comment")].value and global_parameters["mobilesaude.update_protocol_with_error"]:

            if buffer:
                buffer += '\n'

            buffer += f'{row[colunas.get("comment")].value}'

        if buffer:
            if not internal_observation_url["exists"]:
                payload = {"observacao_interna": buffer}
                create_internal_observation_req = safe_post(
                    session,
                    internal_observation_url["url"],
                    json=payload
                )
                if not create_internal_observation_req.ok:
                    error_message = f'Falha ao criar observação interna do protocolo: {protocol_id}'
                    return error_message

            payload = {"observacoes-internas": buffer}

            response = safe_put(
                session, internal_observation_url["url"], json=payload)

            if not response.ok or response.json().get("status"):
                error_message = f'Falha ao atualizar a observação interna do protocolo {protocol_id}: {response.status} {response.text()}'
                return error_message

        internal_observation_url["exists"] = False
        internal_observation_url["url"] = None

        if change_status and PEG:
            analise = global_parameters["mobilesaude.status_analysis"]
            url = global_parameters["mobilesaude.update_protocol_status"] % (
                refund_id, analise)
            req_change_status = safe_put(session, url)

            if not req_change_status.ok:
                error_message = f'Não foi possível atualizar o status do protocolo {protocol_id}'
                return error_message
            
            return None

    logging.info(f'Acessando o site...')
    logging.info(global_parameters[f"mobilesaude.url_{env}"])
    logging.info(f'Usuário {config["mobilesaude"]["username"]}')

    wb = safely_load_workbook(worksheet)

    if not wb:
        return 0, 0

    ws = wb.active

    protocols_count = 0
    protocols_total = ws.max_row - 1

    if not login_mobile_saude(session):
        return protocols_total, 0

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):

        protocols_count += 1
        error_message = ""
        protocol_id = row[colunas.get("protocol_id")].value
        PEG = row[colunas.get("PEG")].value
        refund_id = row[colunas.get("refund_id")].value
        comment = row[colunas.get("comment")].value
        

        logging.info(
            f'Consultando o protocolo {protocol_id} {protocols_count}/{protocols_total}...')

        change_status = (not comment) and (PEG)

        if change_status or global_parameters["mobilesaude.update_protocol_with_error"]:
            error_message = None
            
            protocol_data = get_occurrence_by_protocol(session, protocol_id)
            occurrence_data = get_occurrence_data(session, refund_id)

            if not protocol_data or not occurrence_data:
                error_message = f"Não foi possível obter os dados do protocolo {protocol_id}"
                logging.error(error_message)
                continue

            assigned = safe_get_text(protocol_data, "nome_atendente")
            status_id = safe_get_text(protocol_data, "id_status")
            status_desc = safe_get_text(protocol_data, "status_label")

            error_message = update_protocol(change_status)
            if error_message:
                if isinstance(error_message, Exception):
                    error_message = str(error_message)
                logging.error(error_message)

            # se não tinham comentário de erro, verifica se agora tem
            if not comment and error_message:
                row[colunas.get("comment")].value = error_message

            # se tem comentário de erro, checa se precisa atribuir para usuário humano
            if comment and global_parameters["mobilesaude.assign_alternative"]:
                alternative_user = global_parameters["mobilesaude.alternative_user_name"]
                
                if assigned == alternative_user:
                    logging.info(f"O protocolo {protocol_id} está atribuído ao usuário alternativo {alternative_user}.")
                
                else:
                    if assign(alternative_user, occurrence_data, global_parameters["mobilesaude.status_analysis"]):
                        row[colunas.get("assigned")].value = alternative_user
                    else:
                        logging.error(f"Não foi possível atribuir o protocolo {protocol_id} ao usuário alternativo {alternative_user}.")

            # se terminou o processo sem comentário de erro, setar "OK"
            if not comment:
                row[colunas.get("comment")].value = "OK"

            wb.save(worksheet)  # atualiza o Excel

        else:
            logging.warning(
                f'Nada para atualizar no prococolo {protocol_id}...')

        protocols_success += 1 if comment == "OK" else 0

    wb.save(worksheet)

    return protocols_total, protocols_success
