import os
from sys import platform
from dotenv import load_dotenv

# Load Environment Variables
if platform == "win32":
    load_dotenv()

# Config Variables from Environment Variables
class Config:
    BACKEND_APP_PORT = int(os.getenv('BACKEND_APP_PORT'))
    BACKEND_APP_WORKERS = int(os.getenv('BACKEND_APP_WORKERS'))
    ACCESS_TOKEN_EXPIRE_MINUTES =  int(os.getenv('ACCESS_TOKEN_EXPIRE_MINUTES'))
    REFRESH_TOKEN_EXPIRE_MINUTES =  int(os.getenv('REFRESH_TOKEN_EXPIRE_MINUTES'))
    COOKIE_EXPIRE_SECONDS = int(os.getenv('COOKIE_EXPIRE_SECONDS'))
    SAP_USERNAME = os.getenv('SAP_USERNAME')
    SAP_PASSWORD = os.getenv('SAP_PASSWORD')
    WHITELISTED_LOCAL_IPS = os.getenv('WHITELISTED_LOCAL_IPS').split(',')
    SECRET_KEY = os.getenv('SECRET_KEY')
    SECRET_KEY_SAP = os.getenv('SECRET_KEY_SAP')
    SFTP_HOST_NAME = os.getenv('SFTP_HOST_NAME')
    SFTP_PORT = int(os.getenv('SFTP_PORT'))
    ACCEPT_ZIP_URL = os.getenv('ACCEPT_ZIP_URL')
    AZURE_SUMMARY_URL = os.getenv('AZURE_SUMMARY_URL')
    AZURE_STATUS_URL = os.getenv('AZURE_STATUS_URL')
    STORAGE_PATH = os.getenv('STORAGE_PATH')


