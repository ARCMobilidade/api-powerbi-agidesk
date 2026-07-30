import os.path
import requests
import openpyxl
from config.settings import KEY

def fetch_page(page):
    return requests.get(
        url="https://arc.agidesk.com/api/v1/datasets/serviceissues", 
        params={
            "app_key": KEY,
            "page": page,
            "forecast": "teams"
        }).json()


wb = openpyxl.Workbook()
ws = wb.active

page = 1
row = 1
count_registers = 0
headers = False

print("\033[92;1;6mMontando novo arquivo. Aguarde...\033[0m")
while True:
    register = fetch_page(page)

    print(f"\033[90mPagina {page}")
    if not register:
        break

    count_registers += len(register)

    for reg in register:
        if not headers:
            ws.append(list(reg.keys()))
            headers = True

        ws.append(list(reg.values()))

    if len(register) < 1000:
        print(f"\033[90;1mFinalizando...\033[0m")
        break

    page += 1


try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    old_file = os.path.join(BASE_DIR, "resultado_old.xlsx")
    new_file = os.path.join(BASE_DIR, "resultado.xlsx")

    if os.path.isfile(new_file):
        if os.path.isfile(old_file):
            os.remove(old_file)
        os.rename(new_file, old_file)
    
    wb.save(new_file)

except Exception as e:
    print("Erro ao salvar o arquivo! ->", e)
