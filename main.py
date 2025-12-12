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
headers = [
    "id", "prefix", "title", "content", "created_at", "updated_at", "creator_id",
    "creator", "customer_id", "customer", "customercode", "customertype_id",
    "customerexternalcode", "customertype", "contact_id", "contact", "contactemail",
    "contactphone", "contactcellphone", "status_id", "status", "suspended",
    "priority_id", "priority", "type_id", "type", "source_id", "source", "internal",
    "service_id", "service", "servicetopic_id", "servicetopic", "servicecategory_id",
    "servicecategory", "servicecatalog_id", "servicecatalog", "department_id",
    "department", "costcenter_id", "costcenter", "businessunit_id", "businessunit",
    "tag_id", "tag", "closure_id", "closure", "fact_id", "fact", "factdescription",
    "action_id", "action", "actiondescription", "cause_id", "cause", "product_id",
    "product", "causedescription", "list_id", "list", "board_id", "board", "project_id",
    "project", "workflowstep_id", "workflowstep", "workflow_id", "workflow",
    "workflowcategory_id", "workflowcategory", "responsible_id", "responsible",
    "team_id", "team", "responsedate", "started_at", "duedate", "finished_at",
    "startduration", "finishduration", "totalduration", "absstartduration",
    "absfinishduration", "abstotalduration", "effort", "amount", "houramount",
    "timesheet", "module_id", "cost", "teamgroup_id", "teamgroup", "template_id",
    "template", "cronjob_id", "queuedjob_id", "maincustomer_id", "maincustomer",
    "teamresponsible_id", "teamresponsible", "solutioncomment", "contactcostcenter_id",
    "contactcostcenter", "contactdepartment_id", "contactdepartment",
    "contactbusinessunit_id", "contactbusinessunit", "customertag_id", "customertag",
    "customfield_id", "searchid", "scheduleddate"
]

ws.append(headers)
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
