import requests, json, configparser, csv, os, datetime, ctypes, logging, sys, limit_timer, openpyxl


class send_requests:
    def __init__(self, token, formato, headers): # Cria os atributos necessários para o inicializador
        self.token = token
        self.formato = formato
        self.headers = headers
        date_directory_main = f'C:\\Tiny_Orders\\Tiny_Pedidos' # Cria o diretório principal caso não exista
        os.makedirs(date_directory_main, exist_ok=True)

    def fetch_orders_id(self, url, d_inicial, d_final): # Função para pegar ID's dos pedidos
        global resposta
        data = f"token={self.token}&formato={self.formato}&dataInicial={d_inicial}&dataFinal={d_final}"
        response = requests.post(url, headers=self.headers, data=data)
        resposta = response.text

    def get_orders(self, url, id): # Função para pegar os pedidos
        global resposta2
        data = f"token={self.token}&formato={self.formato}&id={id}"
        response = requests.post(url, headers=self.headers, data=data)
        resposta2 = response.text
        
class error_treatment:
    def __init__(self):
        self.MB_OK = 0x0
        self.ICON_INFO = 0x40
        self.ICON_ERROR = 0x10
        self.logger = logging.getLogger('my_logger') # Cria o objeto de log
        self.logger.setLevel(logging.INFO) # Configura o nível de log
        handler = logging.FileHandler('C:\\Tiny_Orders\\log_file.log') # Cria arquivo handler
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%Y-%m-%d %H:%M:%S') # Define o formato do log
        handler.setFormatter(formatter) # Define o formato do log para o handler
        self.logger.addHandler(handler) # Adiciona o handler ao logger

    def pop_up_erro(self, erro):
        ctypes.windll.user32.MessageBoxW(0, f"{erro}", "ERRO:", self.MB_OK | self.ICON_ERROR)
        
    def pop_up_info(self, info):
        ctypes.windll.user32.MessageBoxW(0, f"{info}", "INFO:", self.MB_OK | self.ICON_INFO)
    
    def log_erro(self, erro):
        self.logger.error(erro)

    def log_info(self, info):
        self.logger.info(info)


# Função para formatar à uma variável um objeto JSON
def jsonfy(directory, variavel):
    with open(directory,"w+") as file:
        response_parsed = json.loads(variavel)
        file.write(json.dumps(response_parsed, indent=2))
        file.close()
        return response_parsed

# Cria uma instância da class para tratamento de erros
error_log = error_treatment()

# Armazena o valor limite de requests
request_limit = 29

# Pega o token e a data do arquivo .cfg
config = configparser.ConfigParser()
config.read("C:\\TinyAPI\\token.cfg")
if "KEY" in config:
    token = config.get("KEY", "token")
    data = config.get("KEY", "data")
else:
    error_log.log_erro("O token e/ou data não foram encontrados no arquivo de configuração ou o mesmo não foi encontrado.")
    error_log.pop_up_erro("Valores Token e/ou Data não encontrados. \nVerifique que o arquivo .cfg existe.")
    sys.exit()

# Cria uma instância da class e atribui os valores: token, format e headers
api_link = send_requests(token, "JSON", {"Content-Type": "application/x-www-form-urlencoded"})

# Envia uma request para pegar os pedidos vendidos no dia x
api_link.fetch_orders_id("https://api.tiny.com.br/api2/pedidos.pesquisa.php", data, data)

# Extrai e armazena em uma lista os id's de todos os pedidos
res_parsed = jsonfy("C:\\Tiny_Orders\\id_orders.json", resposta)

# Converte res_parsed a uma string
res_parsed_text = json.dumps(res_parsed)

# Verifica se não houve pedidos na data
if "Erro" in res_parsed_text:
    error_log.log_erro("A pesquisa não retornou resultados.")
    error_log.pop_up_erro(f"Não foram encontrados pedidos no dia {data}.")
    sys.exit()

# Armazena em uma lista os ID's dos pedidos
orders = res_parsed["retorno"]["pedidos"]
orders_id = list()

for order in orders:
    orders_id.append(order["pedido"]["id"])

# Tratamento de erro ao exceder o limite de pedidos diarios
if len(orders_id) > 59:
    error_log.pop_up_erro("Limite de pedidos excedidos.")
    error_log.log_info("O limite de pedidos para um dia é de 58. Ao exceder, um erro foi gerado.")
    sys.exit()

# Verifica se a quantidade de ID's irá exceder o limite de chamadas da API
if len(orders_id) > request_limit:
    API_limit = True
else:
    API_limit = False

# Formata a data do pedido
date_string = data
parsed_date = datetime.datetime.strptime(date_string, '%d/%m/%Y')
formatted_date = parsed_date.strftime('%Y-%m-%d')

# Cria o diretório (caso não exista) para armazenar os pedidos 
date_directory = f'C:\\Tiny_Orders\\Tiny_Pedidos\\{formatted_date}'
os.makedirs(date_directory, exist_ok=True)

# Carrega a planilha xlsx e define as fileiras
try:
    workbook = openpyxl.load_workbook('C:\\Tiny_Orders\\tabela.xlsx')
    worksheet = workbook['Planilha1']
except:
    error_log.pop_up_erro("Não foi possivel abrir a planilha, verifique o log para mais informações")
    error_log.log_erro("Não foi possivel abrir a planilha, verifique que a planilha está no local correto e que o nome da folha é Planilha1")
next_row = worksheet.max_row + 1
next_row += 1

# Processa os pedidos
c = 0
for pedido in orders_id:
    # Manda a request para obter os pedidos
    api_link.get_orders("https://api.tiny.com.br/api2/pedido.obter.php", pedido)

    # Extrai os campos necessários do pedido
    res_parsed = jsonfy("C:\\Tiny_Orders\\order_fields.json", resposta2)

    # Define se o vendedor é um vendedor em específico
    id_vendedor = res_parsed["retorno"]["pedido"]["id_vendedor"]
    if id_vendedor == "768676428":
        vendedor = True
    else:
        vendedor = False

    # Filtra os pedidos que tem CFe
    if res_parsed["retorno"]["pedido"]["obs_interna"] == "CFe":
        continue
    
    # Pega as informações necessárias do pedido
    try:
        client_name = res_parsed["retorno"]["pedido"]["cliente"]["nome"]

        if "-" in client_name:
            client_name = client_name.replace("-", "_")

        final_price = res_parsed["retorno"]["pedido"]["total_pedido"]
        final_price_replaced = final_price.replace(".", ",")
        desconto = res_parsed["retorno"]["pedido"]["valor_desconto"]
        n_ecommerce = res_parsed["retorno"]["pedido"]["numero_ecommerce"]
        data_pedido = res_parsed["retorno"]["pedido"]["data_pedido"]
        total_pedido = res_parsed["retorno"]["pedido"]["total_pedido"]

        if type(n_ecommerce) != str:
            n_ecommerce = res_parsed["retorno"]["pedido"]["numero"]

    except Exception as E:
        error_log.log_erro(E)
        error_log.pop_up_erro("Houve um erro ao ler parâmetros do arquivo JSON. \n Verifique o log para mais detalhes.")

    # Cálculo de desconto
    original_price = float(final_price) + float(desconto)
    discount_percentage = (desconto / original_price) * 100

    # Arredonda para o valor mais perto
    possiveis_descontos = [5, 7, 10, 12, 15]
    rounded_percentage = min(possiveis_descontos, key=lambda x: abs(x - discount_percentage))

    # Faz a conversão para int
    porcentagem_converted = int(rounded_percentage)
    
    if vendedor:

        # Lê o dicionário que armazena os ID's e os preços
        with open('dict.txt', 'r') as file:
            content = file.read()
        id_prices = eval(content)

        diff = 0

        # Grava em um CSV os dados formatados do JSON em cada pedido
        with open(f'{date_directory}\\#{n_ecommerce}-{client_name}-{final_price_replaced}.csv', 'w', newline='', encoding='utf-8') as file:
            
            csv_writer = csv.writer(file, delimiter=";")
            
            try:
                for venda in res_parsed["retorno"]["pedido"]["itens"]:
                    id_produto = venda["item"]["id_produto"]
                    bar_code = venda["item"]["codigo"]
                    name = venda["item"]["descricao"]
                    quantity = venda["item"]["quantidade"]
                    price = venda["item"]["valor_unitario"] # Preço colocado no produto
                    try:
                        price2 = id_prices.get(id_produto) # Preço de tabela
                    except Exception as E:
                        error_log.log_erro(E)
                        error_log.pop_up_erro(f"O produto '{id_produto}' não foi encontrado na base de dados.")
                        sys.exit()

                    if float(price) > float(price2): # Detecta se há diferença de preço e armazena a diferença
                        diff = diff + (round((float(price) - float(price2)),2) * float(quantity))

                    if desconto > 0: # Detecta se existe desconto no pedido
                        price = round(float(price) * (1 - porcentagem_converted / 100), 4)
                        csv_writer.writerow([bar_code, name, quantity, price])
                    else:
                        csv_writer.writerow([bar_code, name, quantity, price2])

                # Escreve no xlsx as informações de valores cobrados a mais no pedido
                worksheet.cell(row=next_row, column=1, value=data_pedido)
                worksheet.cell(row=next_row, column=2, value=n_ecommerce)
                worksheet.cell(row=next_row, column=3, value=client_name)
                worksheet.cell(row=next_row, column=4, value=total_pedido)
                worksheet.cell(row=next_row, column=5, value=float(total_pedido) - round(float(diff), 2))
                worksheet.cell(row=next_row, column=6, value=round(diff, 2))

                next_row += 1
            except Exception as E:
                error_log.log_erro(E)
                error_log.pop_up_erro("Houve um erro ao criar os arquivos CSV. \n Verifique o log para mais detalhes.")

        # Detecta se o limite da API foi ultrapasado
        if c == (request_limit - 1) and API_limit == True:
            error_log.log_info("Limite de requests atingido.")
            limit_timer.create_timer_window() # Chama a função de timer
        c += 1
    else:

        # Grava em um CSV os dados formatados do JSON em cada pedido
        with open(f'{date_directory}\\#{n_ecommerce}-{client_name}-{final_price_replaced}.csv', 'w', newline='', encoding='utf-8') as file:
            csv_writer = csv.writer(file, delimiter=";")
            try:
                for venda in res_parsed["retorno"]["pedido"]["itens"]:
                    bar_code = venda["item"]["codigo"]
                    name = venda["item"]["descricao"]
                    quantity = venda["item"]["quantidade"]
                    price = venda["item"]["valor_unitario"]
                    if desconto > 0:
                        price = round(float(price) * (1 - porcentagem_converted / 100), 4)
                    else:
                        pass

                    csv_writer.writerow([bar_code, name, quantity, price])
            except Exception as E:
                error_log.log_erro(E)
                error_log.pop_up_erro("Houve um erro ao criar os arquivos CSV. \n Verifique o log para mais detalhes.")
        if c == (request_limit - 1) and API_limit == True:
            error_log.log_info("Limite de requests atingido.")
            limit_timer.create_timer_window() # Chama a função de timer
        c += 1

# Salva o arquivo xlsx
workbook.save('C:\\Tiny_Orders\\tabela.xlsx')

# Mostra informações finais sobre a operação
error_log.pop_up_info(f"Foram recebidos: {c} pedidos do dia {data}")
error_log.log_info(f"Foram recebidos: {c} pedidos do dia {data}")
sys.exit()