import requests, json, configparser, csv, os, datetime, ctypes, logging, sys, time, socket


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

if "Erro" in res_parsed_text:
    error_log.log_erro("A pesquisa não retornou resultados.")
    error_log.pop_up_erro(f"Não foram encontrados pedidos no dia {data}.")
    sys.exit()

orders = res_parsed["retorno"]["pedidos"]
orders_id = list()

for order in orders:
    orders_id.append(order["pedido"]["id"])

# Formata a data do pedido
date_string = data
parsed_date = datetime.datetime.strptime(date_string, '%d/%m/%Y')
formatted_date = parsed_date.strftime('%Y-%m-%d')

# Cria o diretório (caso não exista) para armazenar os pedidos 
date_directory = f'C:\\Tiny_Orders\\Tiny_Pedidos\\{formatted_date}'
os.makedirs(date_directory, exist_ok=True)

# Processa os pedidos
c = 0
for pedido in orders_id:
    # Manda a request para obter os pedidos
    api_link.get_orders("https://api.tiny.com.br/api2/pedido.obter.php", pedido)

    # Extrai os campos necessários do pedido
    res_parsed = jsonfy("C:\\Tiny_Orders\\order_fields.json", resposta2)

    if res_parsed["retorno"]["pedido"]["obs_interna"] == "CFe":
        break
    
    try:
        client_name = res_parsed["retorno"]["pedido"]["cliente"]["nome"]
        final_price = res_parsed["retorno"]["pedido"]["total_pedido"]
        n_ecommerce = res_parsed["retorno"]["pedido"]["numero_ecommerce"]
        if type(n_ecommerce) != str:
            n_ecommerce = res_parsed["retorno"]["pedido"]["numero"]
    except Exception as E:
        error_log.log_erro(E)
        error_log.pop_up_erro("Houve um erro ao ler parâmetros do arquivo JSON. \n Verifique o log para mais detalhes.")

    # Grava em um CSV os dados formatados do JSON em cada pedido
    with open(f'{date_directory}\\#{n_ecommerce}-{client_name}-{final_price}.csv', 'w', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file)
        try:
            for venda in res_parsed["retorno"]["pedido"]["itens"]:
                bar_code = venda["item"]["codigo"]
                name = venda["item"]["descricao"]
                quantity = venda["item"]["quantidade"]
                price = venda["item"]["valor_unitario"]
                csv_writer.writerow([bar_code, name, quantity, price])
        except Exception as E:
            error_log.log_erro(E)
            error_log.pop_up_erro("Houve um erro ao criar os arquivos CSV. \n Verifique o log para mais detalhes.")
    c += 1

error_log.pop_up_info(f"Foram recebidos: {c} pedidos do dia {data}")
error_log.log_info(f"Foram recebidos: {c} pedidos do dia {data}")
sys.exit()