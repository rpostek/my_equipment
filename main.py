from fastapi import FastAPI, Request, Query, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request
from starlette.responses import RedirectResponse
from starlette import status
import win32api
import win32net
import win32security
from dotenv import load_dotenv
import os
import pyodbc
from typing import Annotated

load_dotenv()
db_string = os.getenv('db_string')
env_domain = os.getenv('domain').lower()
app = FastAPI()

def get_logged_ad_user(request):
    try:
        if 'x-iis-windowsauthtoken' in request.headers.keys():
            handle_str = request.headers['x-iis-windowsauthtoken']
            handle = int(handle_str, 16) # need to convert from Hex / base 16
            win32security.ImpersonateLoggedOnUser(handle)
            sid = win32security.GetTokenInformation(handle, 1)[0]
            username, domain, account_type = win32security.LookupAccountSid(None, sid)
            user_info = win32net.NetUserGetInfo(win32net.NetGetAnyDCName(), username, 2)
            full_name = user_info['full_name']
            win32security.RevertToSelf() # undo impersonation
            win32api.CloseHandle(handle) # don't leak resources, need to close the handle!
            if domain.lower() == env_domain:
                return full_name
            else:
                return None
        else:
            return None
    except:
        return None

class Optiest:
    def __init__(self, odbc_string):
        self.odbc_string = odbc_string
        self.hosts = []

    def get_data(self, person, dep):
        with (pyodbc.connect(self.odbc_string) as cnxn):
            sql_query = "select EAN, nazwa, NR_FABRYCZNY, NR_INWENTARZOWY, OSOBA_ODP_FULL, LOK_KOD, WARTOSC_AKT_RAZEM, JOR_KOD \
            from dbo.v_lista_obiekt \
            where STATUS_FULL = 'Oznakowany' and OSOBA_ODP_FULL = ?"
            if dep:
                sql_query += f" and JOR_KOD='{dep}'"
            sql_query += " order by EAN"
            cursor = cnxn.cursor()
            cursor.execute(sql_query, person)
            rows = cursor.fetchall()
            for r in rows:
                self.hosts.append({   "EAN": r[0],
                                      "name": r[1],
                                      "number": r[3],
                                      "serial": r[2] if r[2] else '',
                                      "location": r[5],
                                      "person": r[4],
                                      "value": r[6],
                                      "department": r[7]}
                                     )

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/")
async def root(request: Request, dep: Annotated[str | None, Query()] = None):
    if dep:
        dep = dep.upper()
        if dep not in ['WAG', 'WKS', 'UDW', 'WOS', 'ZIN', 'WOW', 'WIR', 'WGN', 'ZPKS', 'WSSL']:
            dep = None
    full_name = get_logged_ad_user(request)
    full_name = full_name.strip()
    opti = Optiest(db_string)
    opti.get_data(full_name, dep)
    return templates.TemplateResponse(
        request=request, name="items.html", context={"items": opti.hosts, "user": full_name, "dep": dep}
    )

@app.exception_handler(404)
def not_found_exception_handler(request: Request, exc: HTTPException):
    return RedirectResponse(url="/", status_code=status.HTTP_302_FOUND)
