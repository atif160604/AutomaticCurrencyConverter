import json
import urllib.parse
import urllib3
from urllib3.poolmanager import PoolManager
import openpyxl as xl

excel = xl.load_workbook("/Users/atifagboatwala/Desktop/CurrencyExchange/Daily Exchange Rate 22-08-2023.xlsx")

exchange_rate_sheet = excel["22-08-2023"]

print(exchange_rate_sheet['C9'].value)







#
# API_BASE_ADDRESS = 'https://fx-api.fluentax.com'
#
# testEndpoint = "https://fx-api.fluentax.com/v1/Currencies"
# endpoint = "https://sso.fluentax.com/auth/realms/fluentax/protocol/openid-connect/token"
# client_id = "53d8d57a-8fce-45df-b7f2-e7d20eefe548"
# client_secret = "ApvcYs8npXVC09CVzfuXFt9gLNuZHkk0"
#
# dictionary_of_currency = {
#     "USD": 5
# }
#
# parameters = {
#     'grant_type': "client_credentials",
#     'client_id': client_id,
#     'client_secret': client_secret,
#     'scope': 'fx_api',
# }
#
#
# def get_access_token(http: PoolManager) -> str:
#     fields: dict[str, str] = {'grant_type': 'client_credentials', 'client_id': client_id,
#                               'client_secret': client_secret, 'scope': 'fx_api'}
#
#     token_response = http.request(
#         'POST',
#         endpoint,
#         fields=fields,
#         encode_multipart=False
#     )
#     token_response_data = json.loads(token_response.data.decode('utf-8'))
#     access_token: str = token_response_data['access_token']
#
#     return access_token
#
#
# def retrieve_latest_exchange_rates():
#     http = urllib3.PoolManager()
#
#     access_token = get_access_token(http)
#
#     bank_id = 'AECB'
#     format = 'json'
#     latest_rates_request_url = urllib.parse.urljoin(
#         API_BASE_ADDRESS, f'v1/Banks/{bank_id}/DailyRates/Latest?format={format}')
#
#     latest_rates_response = http.request(
#         'GET',
#         latest_rates_request_url,
#         headers={'Authorization': f'bearer {access_token}'}
#     )
#     rates = latest_rates_response.json()
#     convert = rates["rates"]["2024-02-06"]["USD"]["rate"]
#     print(convert)
#     print(latest_rates_response.json())


# retrieve_latest_exchange_rates()
