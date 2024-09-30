import requests

# validating ičo's using ares api

response = requests.get("https://ares.gov.cz/ekonomicke-subjekty-v-be/rest/ekonomicke-subjekty/27468801")


print(response.json())