import logging
from packages.global_parameters import *
from packages.config import *
from packages.commons import *
import ast


@timeit
@handle_exceptions(default_return=[])
@retry(retry_on_exception=try_again_on_any_exception, wait_fixed=10000, stop_max_attempt_number=5)
def get_protocols(session) -> list:

    payload = json.dumps({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "timesync"
    })

    response = safe_post(session,global_parameters["mobilesaude.timestamp"], headers={'Content-Type': 'application/json'}, data=payload, verify=False)

    if not response.ok:
        logging.error(f"Erro ao obter timestamp: {response.status_code}")
        return None

    custom_headers = session.headers.copy()

    subject_filters = global_parameters["mobilesaude.subject_filters"]

    if isinstance(subject_filters, str) and subject_filters.startswith("[") and subject_filters.endswith("]"):
        subject_filters = ast.literal_eval(subject_filters)

    filter_value = json.dumps({
        "dataRegistro": ["", ""],
        "mostrar_encerrados": False,
        "id_status": [global_parameters["mobilesaude.requested_status"]],  
        "atendimento":1,
        "id_tipo_ocorrencia": global_parameters["mobilesaude.occurrence_type_filter"]
    })

    payload = {
        "limit": 9999,
        "order": "DESC",
        "Order_field": "data_registro",
        "Filter": filter_value
    }

    for key, value in payload.items():
        custom_headers[key] = value if isinstance(value, str) else str(value)

    response_occurrences = safe_get(
        session,
        global_parameters["mobilesaude.query_protocolos"], headers=custom_headers, verify=False
    )

    if not response_occurrences.status_code == HTTPStatus.OK or not response_occurrences.json().get("status"):
        logging.error(f"Erro ao obter protocolos")
        return None

    protocols = json.loads(response_occurrences.text)["data"]

    with open(f'{log}/{today}{today_.strftime("%H%M%S")}.json', "w") as file:
        json.dump(protocols, file, indent=4)

    return protocols


@timeit
def export_backlog():

    # Create a session object to maintain cookies and headers across requests
    session = requests.Session()

    if not login_mobile_saude(session):
        return None

    protocols = get_protocols(session)

    # Extração das datas (coluna índice 4)
    # datas = [linha[4] for linha in protocols]
    datas = [
        transform_timestamp_to_datetime(linha.get("data_registro"))
        if isinstance(linha, dict) else None
        for linha in protocols
    ]

    # Contar ocorrências por data
    contagem = Counter(datas)
    total = sum(contagem.values())

    # Exibir resultados
    registros = []
    logging.info(f"{'Data':<15} {'Qtd':<5} {'%':<6}")
    for data, qtd in sorted(contagem.items()):
        percentual = round((qtd / total) * 100, 2)
        registros.append({
            "Data": data,
            "Quantidade": qtd,
            "Percentual (%)": percentual
        })
        logging.info(f"{data:<15} {qtd:<5} {percentual:>5.1f}%")

    # Exportar para Excel
    df = pd.DataFrame(registros)
    df.to_excel(
        f'{log}/backlog_por_data_{today}{today_.strftime("%H%M%S")}.xlsx', index=False)

    return None


@timeit
def read_protocols_mobile_saude() -> str:

    @timeit
    def process_protocol() -> list:

        file_id = os.path.basename(file)
        refund_id = protocol["id_ocorrencia"]
        protocol_id = protocol["protocolo"]
        status_id = protocol["id_status"]
        refund_type = protocol["assunto"]
        refund_qty = 1
        protocol_date = transform_timestamp_to_datetime(protocol["data_registro"])
        protocol_date = protocol_date
        refund_value = ""
        status_desc = protocol["status_label"]
        PEG = ""
        holder_name = ""
        phone_number = ""
        holder_cpf = ""
        plan = ""
        payment_day = ""
        payment_type = ""
        lot = ""
        notes = ""
        expense_id = ""
        card = ""
        user = ""
        expense_status = ""
        supplier_id = ""
        supplier_name = ""
        supplier_state = ""
        supplier_city = ""
        expense_date = ""
        expense_nf = ""
        guide_number = ""
        assigned = ""
        peg_situation = ""
        comment = ""
        peg_occurrence = ""
        refunded_value = ""
        complement = ""

        logging.info(f'Extraindo dados do protocolo {protocol_id}...')

        occurrence_url = global_parameters["mobilesaude.occurrence_details"] % refund_id

        response = safe_get(session, occurrence_url)

        if response.status_code == HTTPStatus.OK:
            response_json = response.json()
            data = response_json.get("data")

            form_data = data["formData"]
            id_form = form_data[0]["id_form"]
            id_form_data = form_data[0]["id_form_data"]

            form_data_url = global_parameters["mobilesaude.form_data"] % (id_form, id_form_data)

            request_form_data = safe_get(session, form_data_url)

            if not request_form_data.status_code == HTTPStatus.OK:
                logging.warning(f'Não foi possível obter os dados do formulário para o protocolo {protocol_id}')
                return None

            request_form_data_json = request_form_data.json().get("data")["data"]

            assigned = safe_get_text(data, "atendente")
            plan = safe_get_text(request_form_data_json, "plano")
            holder_name = safe_get_text(data["solicitante"], "nome")
            holder_cpf = safe_get_text(request_form_data_json, "beneficiario_cpf")
            phone_number = safe_get_phone_number(data, request_form_data_json)
            refund_value = safe_get_text(request_form_data_json, "valor-da-despesa")
            user = safe_get_text(request_form_data_json, "beneficiario") or safe_get_text(data["beneficiario"], "nome")
            supplier_name = safe_get_text(request_form_data_json, "nome-fantasia")
            expense_nf = safe_get_text(request_form_data_json, "numero-da-nota-fiscal-recibo")
            expense_date = convert_date(safe_get_text(request_form_data_json, "data-da-despesa"))
            expense_status = safe_get_text(data, "status_label_interna")
            payment_type = safe_get_label_in_array(request_form_data_json, "tipo-de-reembolso")
            supplier_city = safe_get_label_in_array(request_form_data_json, "cidade")
            supplier_state = safe_get_label_in_array(request_form_data_json, "estado")
            supplier_id = safe_supplier_document(request_form_data_json)
            card = safe_get_text(request_form_data_json, "matricula")

            service_tabs = data.get("abasAtendimento")

            payment_data_url = get_url_payment_data(service_tabs)

            if payment_data_url:
                payment_data = get_payment_data(payment_data_url, session).get("data")

                get_payment_day = safe_get_text(payment_data, "dataPagamento")
                payment_day = convert_date(get_payment_day) 
                lot = safe_get_text(payment_data, "numeroLote")
                PEG = lot
                notes = safe_get_text(payment_data, "observacoes")
            else:
                logging.warning(f'Não foi possível obter os dados de pagamento do protocolo {protocol_id}')
                
            comment = "PEG previamente associado" if PEG else ""
        else:
            comment = f"Erro ao abrir detalhes do protocolo {protocol_id} code: {response.status_code}"

        if comment:
            logging.error(comment)

        return [
            file_id,
            env,
            refund_id,
            protocol_id,
            protocol_date,
            status_id,
            status_desc,
            refund_type,
            refund_qty,
            refund_value,
            card,
            user,
            holder_name,
            holder_cpf,
            phone_number,
            plan,
            payment_day,
            payment_type,
            lot,
            expense_id,
            expense_status,
            supplier_id,
            supplier_name,
            supplier_state,
            supplier_city,
            expense_date,
            expense_nf,
            PEG,
            notes,
            assigned,
            guide_number,
            refunded_value,
            peg_situation,
            peg_occurrence,
            comment,
            complement,
        ]

    
    @timeit
    def get_protocols_by_env(env) -> pd.DataFrame:

        logging.info(f'Selecionando protocolos do ambiente {env}...')

        db = LMDBWrapper()

        # Converter para DataFrame
        df = pd.DataFrame(db.all())
        
        if not df.empty:

            # Ordenar por datetime decrescente
            df = df.sort_values(by='file_id', ascending=False)

            # Pegar a primeira ocorrência de cada protocol_id (a mais recente)
            df = df.drop_duplicates(subset='protocol_id', keep='first')

            # Filtrar apenas registros do ambiente
            df = df[df['env'] == env]

        else:

            df = pd.DataFrame(columns=column_ids)

        logging.info(
            f'Selecionando {len(df)} protocolos do ambiente {env}... OK')

        return df

    @timeit
    def get_last_protocols_with_error() -> pd.DataFrame:

        def extrair_datetime(file_id):
            try:
                # Procura padrão de data + hora no formato 20250413_1652
                match = re.search(r'(\d{8}_\d{4})', file_id)
                if match:
                    return datetime.strptime(match.group(1), "%Y%m%d_%H%M")
            except:
                pass
            return None

        logging.info(
            f'Selecionando protocolos com erro nas últimas {global_parameters["sla_tratamento_erros"]} horas...')

        df = get_protocols_by_env(global_parameters["env"])

        # cria a coluna com datetime
        df['file_datetime'] = df['file_id'].apply(extrair_datetime)

        # Limites de tempo
        agora = datetime.now()
        limite_inferior = agora - \
            timedelta(hours=int(global_parameters["sla_tratamento_erros"]))

        # Filtro
        filtro = (
            df['PEG'].isnull() &
            df['file_datetime'].notnull() &
            (df['file_datetime'] >= limite_inferior) &
            (df['file_datetime'] <= agora) &
            (~(df['comment'] == "OK")) &
            df['comment'].notnull()
        )

        df = df[filtro]

        # dropa as colunas que não estão em column_names
        df = df[["protocol_id"]]

        # Ordenar do mais antigo para o mais novo
        df = df.sort_values(by='protocol_id', ascending=True)

        total = len(df)

        for i, row in enumerate(df.itertuples(index=False), start=1):
            logging.info(f'Protocolo {row.protocol_id} com erro {i}/{total}')

        logging.info(f'{total} protocolo(s) selecionado(s) com erro.')

        return df

    @timeit
    def get_reprocess_protocols() -> pd.DataFrame:

        logging.info(f'Selecionando protocolos para reprocessamento...')

        df = get_protocols_by_env(global_parameters["env"])

        df = df[
            (df['PEG'].notnull()) &
            (~(df['comment'] == "OK"))
        ]

        # Limpa as mensagens de erro para poder reprocessar o protocolo
        df['comment'] = None


        #
        df.rename(columns=dict(zip(column_ids, column_names)), inplace=True)


        # dropa as colunas que não estão em column_names
        df = df[column_names]

        # Ordenar do mais antigo para o mais novo
        df = df.sort_values(by='protocolo', ascending=True)

        total = len(df)

        for i, row in enumerate(df.itertuples(index=False), start=1):
            logging.info(
                f'Protocolo {row.protocolo} do arquivo {row.arquivo} selecionado para reprocessamento {i}/{total}')

        logging.info(
            f'{total} protocolo(s) selecionado(s) para reprocessamento.')

        return df

    env = global_parameters["env"]

    logging.info(f'Extraindo protocolos do Mobilesaude...')
    logging.info(f'Acessando o site...')
    logging.info(global_parameters[f"mobilesaude.url_{env}"])
    logging.info(f'Usuário {config["mobilesaude"]["username"]}')

    session = requests.Session()

    if not login_mobile_saude(session):
        return None

    file = f'{log}/protocolos_{today}_{today_.strftime("%H%M")}.xlsx'
    array = [pd.read_excel("./data/parameters.xlsx", engine='openpyxl', sheet_name="labels", header=None).iloc[:, 0].tolist()]
    column_names = pd.read_excel("./data/parameters.xlsx", engine='openpyxl',sheet_name="labels", header=None).iloc[:, 0].tolist()
    column_ids = pd.read_excel("./data/parameters.xlsx", engine='openpyxl',sheet_name="ids", header=None).iloc[:, 0].tolist()

    df_reprocess_protocols = get_reprocess_protocols()

    # somente pesquisar protocolos se a lista de reprocessamento for menor que a quantidade esperada para processamento
    if len(df_reprocess_protocols) <= int(global_parameters["mobilesaude.total_records"]):

        # lista de protocolos que apresentaram erro nas últimas x horas
        df_protocols_with_error = get_last_protocols_with_error()
        protocols_with_error = set(
            df_protocols_with_error['protocol_id'].values)

        #
        total_records = int(
            global_parameters["mobilesaude.total_records"]) - len(df_reprocess_protocols)

        protocols = get_protocols(session)

        # Filtra o array removendo os que já existem no protocols_with_error
        protocols = [
            linha for linha in protocols if linha["protocolo"] not in protocols_with_error]

        protocols = protocols[:total_records]

        protocols_count = 0

        for protocol in protocols:

            protocols_count += 1

            logging.info(
                f'Lendo protocolo {protocol["protocolo"]} {protocols_count}/{len(protocols)}')

            # verifica se o protocolo já está na lista de reprocessamento
            if not protocol["protocolo"] in df_reprocess_protocols['protocolo'].values:
                array.append(process_protocol())
            else:
                logging.info(
                    f'Protocolo {protocol["protocolo"]} já estava na lista de reprocessamento')

    logging.info(f'Exportando {file}...')

    # Criar um DataFrame com os dados
    df = pd.DataFrame(array[1:], columns=array[0]).astype(str)

    df_reprocess_protocols['arquivo'] = os.path.basename(file)

    df = pd.concat([df_reprocess_protocols, df], ignore_index=True)

    # Limitar a quantidade de protocolos
    df = df.head(int(global_parameters["mobilesaude.total_records"]))

    # Ordenar do mais antigo para o mais novo
    df = df.sort_values(by='protocolo', ascending=True)

    # Salvar corretamente em formato aberto para `openpyxl`
    df.to_excel(file, index=False, engine="openpyxl")
    
    file

    return file
