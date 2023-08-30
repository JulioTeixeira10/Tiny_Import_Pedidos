import requests, json, configparser, csv


class send_requests:
    def __init__(self, token, formato, headers):
        self.token = token
        self.formato = formato
        self.headers = headers

    def fetch_orders_id(self, url, d_inicial, d_final):
        global resposta
        data = f"token={self.token}&formato={self.formato}&dataInicial={d_inicial}&dataFinal={d_final}"
        response = requests.post(url, headers=self.headers, data=data)
        resposta = response.text

    def get_orders(self, url, id):
        global resposta2
        data = f"token={self.token}&formato={self.formato}&id={id}"
        response = requests.post(url, headers=self.headers, data=data)
        resposta2 = response.text
        

# Função para formatar à uma variável um objeto JSON
def jsonfy(directory, variavel):
    with open(directory,"w+") as file:
        response_parsed = json.loads(variavel)
        file.write(json.dumps(response_parsed, indent=2))
        file.close()
        return response_parsed


# Pega o token e a data do arquivo .cfg
config = configparser.ConfigParser()
config.read("C:\\TinyAPI\\token.cfg")
token = config.get("KEY", "token")
data = config.get("KEY", "data")

# Cria uma instância da class e atribui os valores: token, format e headers
api_link = send_requests(token, "JSON", {"Content-Type": "application/x-www-form-urlencoded"})

# Envia uma request para pegar os pedidos vendidos no dia x
api_link.fetch_orders_id("https://api.tiny.com.br/api2/pedidos.pesquisa.php", data, data)

# Extrai e armazena em uma lista os id's de todos os pedidos
res_parsed = jsonfy("C:\\Tiny_Orders\\id_orders.json", resposta)
orders = res_parsed["retorno"]["pedidos"]
orders_id = list()
for order in orders:
    orders_id.append(order["pedido"]["id"])




api_link.get_orders("https://api.tiny.com.br/api2/pedido.obter.php", "761440137")

# Extrai os campos necessários do pedido
res_parsed = jsonfy("C:\\Tiny_Orders\\order_fields.json", resposta2)

client_name = res_parsed["retorno"]["pedido"]["cliente"]["nome"]
final_price = res_parsed["retorno"]["pedido"]["total_pedido"]

for venda in res_parsed["retorno"]["pedido"]["itens"]:
    bar_code = venda["item"]["codigo"]
    name = venda["item"]["descricao"]
    quantity = venda["item"]["quantidade"]
    price = venda["item"]["valor_unitario"]
