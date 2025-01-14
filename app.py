from main import *
from api import *
import uvicorn
# from uvicorn import Config, Server
from fastapi import FastAPI, Cookie
from fastapi.middleware.cors import CORSMiddleware
import logging, logging.config
import sqlite3
import jwt
import os
import traceback
from fastapi import FastAPI, Depends
from config import Config

from starlette_context import context, request_cycle_context, _request_scope_context_storage

async def my_context_dependency(token = Cookie(default=None)):
    try:
        decoded_cookie = jwt.decode(token, options={"verify_signature": False})
        data = {"username": decoded_cookie.get('username'), "role": decoded_cookie.get('aud')}
    except Exception:
        data = {"username": None}

    with request_cycle_context(data):
        # yield allows it to pass along to the rest of the request
        yield

app = FastAPI(debug=False, dependencies=[Depends(my_context_dependency)])
app.include_router(router, prefix="/api")
app.include_router(api, prefix="/api")

# with open('config.json', 'r') as config_file:
#     contents = config_file.read()
#     config = json.loads(contents)

origins = Config.WHITELISTED_LOCAL_IPS

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

@app.on_event('startup')
def before_first_request():
    try:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        client_db_path = os.path.join(dir_path, 'client.db')
        client_db = sqlite3.connect(client_db_path, timeout=10)
        c = client_db.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS client (Name TEXT PRIMARY KEY, PAN TEXT, Email TEXT,role TEXT, Created_At TIMESTAMP)")
        c.execute("CREATE TABLE IF NOT EXISTS newrequest (username TEXT, password TEXT, nameofuser TEXT,role TEXT, email TEXT, activated INTEGER)")
        # c.execute("CREATE TABLE IF NOT EXISTS users (username TEXT, password TEXT, nameofuser TEXT, email TEXT, created_at TIMESTAMP, activated INTEGER, date_modified TIMESTAMP)")
        client_db.commit()
        c.close()
        client_db.close()
            
    except: 
        traceback.print_exc()

class ContextFilter(logging.Filter):
    """
    This is a filter which injects contextual information into the log.

    Rather than use actual contextual information, we just use random
    data in this demo.
    """

    USERS = ['samir', 'mayur', 'ronit']

    def filter(self, record):
        try:
            data = _request_scope_context_storage.get()
        except:
            data = {}
        record.user = data.get('username')
        return True

@app.get('/')
def health():
    return "Working"

import concurrent_log_handler

LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': { 
        'logformatter': { 
            'format': '[%(asctime)s.%(msecs)03d] %(levelname)s [%(thread)d] - %(user)s - %(message)s'
        },
    },
    'handlers': {
        'logconsole': {
            'level': 'INFO',
            'formatter': 'logformatter',
            'class': 'logging.StreamHandler',
            'stream': 'ext://sys.stdout',
            'filters'  : ['customfilter']
        },
        'hand01': {
            'level': 'NOTSET',
            'formatter': 'logformatter',
            'class': 'concurrent_log_handler.ConcurrentRotatingFileHandler',
            'filename': 'logs/logfile.log',
            'mode': 'a',
            'maxBytes': 10*1024*1024,
            'backupCount': 9999,
            'filters'  : ['customfilter']
        }
    },
    'root': { 
        'handlers': ['logconsole', 'hand01'],
        'level': 'DEBUG',
        'propagate': False
    },
    'filters': {
        'customfilter': {
            '()' : ContextFilter,
        }
    },
}


if __name__ == "__main__":
    # uvicorn.run("app:app", host="0.0.0.0", port=3001, reload=True, log_config=LOGGING_CONFIG)
    # For Hosting remove reload
    print(Config.BACKEND_APP_PORT)
    uvicorn.run("app:app", host="0.0.0.0", port=Config.BACKEND_APP_PORT, workers=Config.BACKEND_APP_WORKERS, log_config=LOGGING_CONFIG)