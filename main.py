import os.path
import requests
import openpyxl
from config.settings import KEY

def fetch_page(page_request):
    return requests.get("https://arc.agidesk.com/api/v1/datasets/serviceissues",
                        params={
                            "app_key": KEY,
                            "page": page_request,
                            "forecast": "teams"
                        }).json()

wb = openpyxl.Workbook()
ws = wb.active

page = 1
row = 1

while True:
    register = fetch_page(page)
    print("Página", page, "→", len(register), "registros")

    if not register:
        break

    for reg in register:
        ws.append(list(reg.values()))

    if len(register) < 1000:
        break

    page += 1

try:
    if os.path.isfile("resultado.xlsx"):
        os.remove("resultado_old.xlsx")
        os.rename("resultado.xlsx", "resultado_old.xlsx")
        print("Arquivo substituído!")
    else:
        print("Arquivo salvo!")

    wb.save("resultado.xlsx")
except Exception as e:
    print("Erro ao salvar o arquivo! ->", e)
