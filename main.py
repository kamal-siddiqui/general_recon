from fastapi import BackgroundTasks, APIRouter, Body
from fastapi.middleware.cors import CORSMiddleware
import hashlib
from Crypto import Random
from Crypto.Cipher import AES
from base64 import b64encode, b64decode
from fastapi.responses import JSONResponse
import zipfile
import logging
import requests
from shutil import copyfile, rmtree
from PIL import Image
import sqlite3
from datetime import datetime, timedelta
import base64
import json
import re
import hashlib
import os
import pandas as pd
from pandas import options
from dateutil.relativedelta import relativedelta, MO
import traceback
import uuid
import main_onlyexecution
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions, BlobClient
import paramiko
from config import Config
# from dotenv import load_dotenv
# load_dotenv(os.path.dirname(os.path.realpath(__file__)) + "/.env")

options.io.excel.xlsx.writer = 'xlsxwriter'
Image.MAX_IMAGE_PIXELS = None

router = APIRouter()
router.message = ""

# with open('config.json', 'r') as config_file:
#     contents = config_file.read()
#     config = json.loads(contents)

dir_path = os.path.dirname(os.path.realpath(__file__))

# --------------------------- AUTH APIS START --------------------------- #

# Class For AES Encryption / Decryption

class AESCipher(object):
    def __init__(self, key=None):
        self.block_size = AES.block_size
        self.key = hashlib.sha256(key.encode()).digest()

    def encrypt(self, plain_text):
        plain_text = self.__pad(plain_text)
        iv = Random.new().read(self.block_size)
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        encrypted_text = cipher.encrypt(plain_text.encode())
        return b64encode(iv + encrypted_text).decode("utf-8")

    def decrypt(self, encrypted_text):
        encrypted_text = b64decode(encrypted_text)
        iv = encrypted_text[:self.block_size]
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        plain_text = cipher.decrypt(encrypted_text[self.block_size:]).decode("utf-8")
        return self.__unpad(plain_text)

    def decrypt_sap(self, encrypted_text, secret):
        try:
            cipher = AES.new(secret, AES.MODE_ECB)
            encrypted_text_bytes = base64.b64decode(encrypted_text)
            decrypted_text_bytes = cipher.decrypt(encrypted_text_bytes)
            decrypted_text_bytes = decrypted_text_bytes.decode('utf-8')
            decrypted_text_bytes = self.__unpad(decrypted_text_bytes)
            return decrypted_text_bytes
        except Exception as ex:
            msg = "Exception while decrypting data"
            raise Exception(msg) from ex

    def __pad(self, plain_text):
        number_of_bytes_to_pad = self.block_size - \
            len(plain_text) % self.block_size
        ascii_string = chr(number_of_bytes_to_pad)
        padding_str = number_of_bytes_to_pad * ascii_string
        padded_plain_text = plain_text + padding_str
        return padded_plain_text

    @staticmethod
    def __unpad(plain_text):
        last_character = plain_text[len(plain_text) - 1:]
        return plain_text[:-ord(last_character)]

# Auth Constants

ACCESS_TOKEN_EXPIRE_MINUTES = Config.ACCESS_TOKEN_EXPIRE_MINUTES  #10
REFRESH_TOKEN_EXPIRE_MINUTES = Config.REFRESH_TOKEN_EXPIRE_MINUTES  #30
COOKIE_EXPIRE_SECONDS = Config.COOKIE_EXPIRE_SECONDS
SECRET_KEY = Config.SECRET_KEY
SECRET_KEY_SAP = Config.SECRET_KEY_SAP
STORAGE_PATH = Config.STORAGE_PATH
cipher = AESCipher(SECRET_KEY)


def execute_query(db_path, execution_query):
    db = sqlite3.connect(db_path, timeout=100, check_same_thread=True)
    c = db.cursor()
    c.execute(execution_query)
    c.close()
    db.commit()
    db.close()

def run_gl_recon_azure(account_name, account_key, container_name, input_files, master_files, fromdate, todate, gstinlist, requestId, groupCode):
    try:
        temp_db_name = str(uuid.uuid4())
        start_time = datetime.now()
        # STORAGE_PATH = os.getenv('STORAGE_PATH')

        # execute_query('execution.db', f"UPDATE filestatus SET status = 'Processing Complete' where request_id = '{requestId}'")
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"INSERT INTO filestatus (request_id, temp_db_name,recd_time,status) VALUES ('{requestId}', '{temp_db_name}','{start_time}','')")

        # create a client to interact with blob storage
        connect_str = 'DefaultEndpointsProtocol=https;AccountName=' + account_name + ';AccountKey=' + account_key + ';EndpointSuffix=core.windows.net'

        blob_service_client = BlobServiceClient.from_connection_string(connect_str)

        # use the client to connect to the container
        container_client = blob_service_client.get_container_client(container_name)
        blob_list = list(map(lambda x: x.name, container_client.list_blobs()))

        # get a list of all blob files in the container
        blob_list = list(filter(lambda x: x in input_files + master_files, blob_list))

        # Sort the filtered list to ensure master_files come before input_files
        blob_list = sorted(blob_list, key=lambda x: (x not in master_files, x not in input_files))


        dir_path = os.path.dirname(os.path.realpath(__file__))
        path = dir_path + '/Azure Blob GL Recon'
        panwisedb_path = f'{path}/{temp_db_name}.db'

        temp_db = sqlite3.connect(panwisedb_path)
        c = temp_db.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS masters (Particulars TEXT PRIMARY KEY, Status TEXT)")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('GSTIN','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('GL Code','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('Supply Type','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('Doc Type','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('Tax Code','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('TB','Pending')")
        c.close()
        temp_db.commit()

        # generate a shared access signiture for files and load them into Python
        for blob_i in blob_list:
            # generate a shared access signature for each blob file
            sas_i = generate_blob_sas(account_name = account_name,
                                    container_name = container_name,
                                    blob_name = blob_i,
                                    account_key=account_key,
                                    permission=BlobSasPermissions(read=True),
                                    expiry=datetime.utcnow() + timedelta(hours=1))
            sas_url = 'https://' + account_name+'.blob.core.windows.net/' + container_name + '/' + blob_i + '?' + sas_i
            main_onlyexecution.process_file(False, [temp_db_name, blob_i, sas_url, '/Azure Blob GL Recon'])
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"UPDATE filestatus SET status = 'Processing Complete' where request_id = '{requestId}'")

        main_onlyexecution.valgldump(False, [temp_db_name, 'Azure Blob GL Recon'])
        # main_onlyexecution.seqRun(False, [temp_db_name, ["False", "False", "False", "False", "False", "False", "False"], 'Azure Blob GL Recon'])
        main_onlyexecution.valsr(False, [temp_db_name, 'Azure Blob GL Recon'])
        main_onlyexecution.uploadsummary(False, [temp_db_name, 'Azure Blob GL Recon'])
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"UPDATE filestatus SET status = 'Validation Complete' where request_id = '{requestId}'")

        # call exportgaps from main and srg1 in reports to export srvsgl report
        # fromdate and todate in documentdate format
        output_file = main_onlyexecution.exportgaps(False, [temp_db_name, ['srgl'], fromdate, todate, gstinlist, 'Azure Blob GL Recon'])

        # saveGlReconSummary
        recon_summary = pd.read_sql_query('''select * from Summary_Totals where Particulars IN ('Sales Register','GL Dump - Output');''', temp_db)
        gl_dump_summary = recon_summary[recon_summary['Particulars'] == 'GL Dump - Output'].reset_index(drop=True)
        sr_summary = recon_summary[recon_summary['Particulars'] == 'Sales Register'].reset_index(drop=True)
        summary = {
            "requestId" : str(requestId),
            "groupCode": str(groupCode),
            "summary" : 
            {
            "srInvCount" : sr_summary['Count'][0],
            "srTaxval" : sr_summary['Taxable_Value'][0],
            "srIgst" : sr_summary['IGST'][0],
            "srSgst" : sr_summary['SGST'][0],
            "srCgst" : sr_summary['CGST'][0],
            "srCess" : sr_summary['Cess'][0],
            "srTotTaxAmt" : sr_summary['Total_Tax_Amount'][0],
            "glDumpInvCount" : gl_dump_summary['Count'][0],
            "glDumpTaxVal" : gl_dump_summary['Taxable_Value'][0],
            "glDumpIgst" : gl_dump_summary['IGST'][0],
            "glDumpCgst" : gl_dump_summary['CGST'][0],
            "glDumpSgst" : gl_dump_summary['SGST'][0],
            "glDumpCess" : gl_dump_summary['Cess'][0],
            "glDumpTotTaxAmt" : gl_dump_summary['Total_Tax_Amount'][0]
            }
        }

        requests.post(Config.AZURE_SUMMARY_URL, json=json.loads(json.dumps(summary, default=str)), verify=False, timeout=100)

        output_file = output_file.to_csv(index=False).encode('utf-8')

        # Open the CSV file and create a ZIP archive
        csv_filename = f"{requestId}"
        zip_filename = csv_filename + ".zip"

        timestamp = datetime.utcnow()+ relativedelta(hours = 5.5)
        timestamp = timestamp.strftime("%d_%m_%Y_%H_%M_%f")
        with zipfile.ZipFile(zip_filename, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr(f'SRvsGL_{requestId}_{timestamp}.csv', output_file)

        # Create a folder (a.k.a. blob directory) and upload zip in it
        folder_name = f"{requestId}/"
        blob_name = f'{requestId}/{zip_filename}'
        blob_client = blob_service_client.get_blob_client(container=container_name, blob=blob_name)
        with open(zip_filename, "rb") as data:
            blob_client.upload_blob(data, overwrite=True)
            
        time_taken = datetime.now() - start_time
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"UPDATE filestatus SET status = 'Report Generated',time_taken = '{time_taken}' where request_id = '{requestId}'")

        # saveGlReconStatus
        requests.post(Config.AZURE_STATUS_URL, json={"requestId" : str(requestId), "groupCode": str(groupCode), "reconStatus" : "Report Generated"}, verify=False, timeout=100)

        try:
            for file in input_files:
                blob_client = container_client.get_blob_client(file)
                blob_client.delete_blob()
        except Exception:
            logging.error("run_gl_recon_azure",exc_info=True)
            traceback.print_exc()

    except Exception:
        logging.error("run_gl_recon_azure",exc_info=True)
        requests.post(Config.AZURE_STATUS_URL, json={"requestId" : str(requestId), "groupCode": str(groupCode), "reconStatus" : "Failed"}, verify=False, timeout=100)
        traceback.print_exc()

@router.post('/initiateGlRecon')
async def initiateGlReconAzure(background_tasks: BackgroundTasks, payload = Body(...)):
    try:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # STORAGE_PATH = os.getenv('STORAGE_PATH')
        if 'execution.db' not in os.listdir(STORAGE_PATH):
            # create filestatus table
            execute_query(os.path.join(STORAGE_PATH, 'execution.db'), "CREATE TABLE IF NOT EXISTS filestatus (request_id TEXT PRIMARY KEY, temp_db_name TEXT, recd_time TEXT, time_taken TEXT, status TEXT)")

        if 'Azure Blob GL Recon' not in os.listdir(dir_path):
            try:
                os.mkdir('Azure Blob GL Recon')
            except:
                pass

        # initiate the gl recon
        glReconDetails = payload.get("glReconDetails")
        requestId = glReconDetails.get("requestId")
        gstinlist = list(map(lambda x: eval(x).strip(), glReconDetails.get("gstin").split(',')))
        fromTaxperiod = glReconDetails.get("fromTaxperiod")
        toTaxperiod = glReconDetails.get("toTaxperiod")
        fromdate = f"{fromTaxperiod[2:]}-{fromTaxperiod[:2]}-01" # 2021-04-23
        todate = f"{toTaxperiod[2:]}-{toTaxperiod[:2]}-01" # 2021-04-23
        gl_dump_files = list(map(lambda x: x.strip(), glReconDetails.get("glDumpFileDetails").get("glDumpFileName").split(',')))
        sales_register_files = list(map(lambda x: x.strip(), glReconDetails.get("slaesregFile").get("salesRegFileName").split(',')))
        # master_files = ['BusinessUnitcode.csv', 'Documenttype.csv', 'GLCodeMappingMasterGL1.csv', 'supplytypeMaster.csv', 'Taxcode.csv', 'TBxlutf8.csv','GLDumpMappingFileMaster.csv']
        master_files = glReconDetails.get("glMasterFileName").split(',')
        groupCode = glReconDetails.get("groupCode")

        # blob credentials
        container_name = glReconDetails.get("contaonerName")
        account_name = glReconDetails.get("storageAccountName")
        account_key = glReconDetails.get("StorageKey")

        # initiate gl recon process
        background_tasks.add_task(run_gl_recon_azure, account_name, account_key, container_name, gl_dump_files + sales_register_files, master_files, fromdate, todate, gstinlist, requestId, groupCode)

        response_json = {
        "status" : "Sucess",
        "error" : 
        {
            "errorCode" : "",
            "errorMsg" : ""
        }
        }
        status_code = 200
    except Exception as e:
        logging.error("initiateGlRecon",exc_info=True)
        response_json = {
        "status" : "Failed",
        "error" : 
        {
            "errorCode" : "",
            "errorMsg" : e.args
        }
        }
        status_code = 500

    return JSONResponse(content=response_json, status_code=status_code)

# GL Recon SAP-SFTP
def run_gl_recon_sap(request_id, group_code, sftp_username, sftp_password, fromTaxPeriod, toTaxPeriod, gstins):
    try:
        sap_username = Config.SAP_USERNAME
        sap_password = Config.SAP_PASSWORD
        
        authorization_header = f"{sap_username}:{sap_password}"
        authorization_header = b64encode(authorization_header.encode())
        authorization_header = authorization_header.decode('utf-8')
        authorization_header = f"Basic {authorization_header}"

        payload = {
            "req": {
                "requestId": request_id,
                "status": "Completed",
                "errMsg": ""
                    }
        }
        headers = {
            "Authorization": authorization_header,
            "groupCode": group_code
        }
        accept_zip_url = Config.ACCEPT_ZIP_URL  # Replace with your actual SAP API URL

        temp_db_name = str(uuid.uuid4())
        start_time = datetime.now()
        # STORAGE_PATH = os.getenv('STORAGE_PATH')
        # execute_query('execution.db', f"UPDATE filestatus SET status = 'Processing Complete' where request_id = '{requestId}'")
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"INSERT INTO filestatus (request_id, temp_db_name,recd_time,status) VALUES ('{request_id}', '{temp_db_name}','{start_time}','')")

        # SFTP connection information
        hostname = Config.SFTP_HOST_NAME
        port = Config.SFTP_PORT
        username = cipher.decrypt_sap(sftp_username, b64decode(SECRET_KEY_SAP))
        password = cipher.decrypt_sap(sftp_password, b64decode(SECRET_KEY_SAP))

        # Create an SFTP client
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        # Connect to the SFTP server
        client.connect(hostname, port, username, password)

        # Open an SFTP session
        sftp = client.open_sftp()

        stfp_inbound_path = os.path.join('SFTP_ROOT','INTG_INBOUND','GLRECON')
        stfp_inbound_path = os.path.join(stfp_inbound_path, str(request_id))
        stfp_outbound_path = os.path.join('SFTP_ROOT','INTG_OUTBOUND','GLRECON')
        files_list = sftp.listdir(stfp_inbound_path)

        business_unit_list = list()
        document_type_list = list()
        gl_code_master_list = list()
        gl_dump_master = list()
        supply_type_list = list()
        tax_code_list = list()
        trial_balance_list = list()
        remaining_files = list()
        for index, element in enumerate(list(map(lambda x: x.lower(), files_list))):
            if re.search('gl', element) and re.search('code', element) and re.search('master', element):
                gl_code_master_list.append(files_list[index])
            elif re.search('business', element) and re.search('unit', element):
                business_unit_list.append(files_list[index])
            elif re.search('document', element) and re.search('type', element):
                document_type_list.append(files_list[index])
            elif re.search('supply', element) and re.search('type', element):
                supply_type_list.append(files_list[index])
            elif re.search('tax', element) and re.search('code', element):
                tax_code_list.append(files_list[index])
            elif re.search('tb', element) or (re.search('trial', element) and re.search('balance', element)):
                trial_balance_list.append(files_list[index])
            elif re.search('gl', element) and re.search('dump', element) and re.search('mapping', element):
                gl_dump_master.append(files_list[index])
            else:
                remaining_files.append(files_list[index])
        files_list = gl_dump_master + business_unit_list + document_type_list + gl_code_master_list + supply_type_list + tax_code_list + trial_balance_list + remaining_files
        
        sftp_archive_path = os.path.join('SFTP_ROOT','INTG_INBOUND','ARCHIVE')

        dir_path = os.path.dirname(os.path.realpath(__file__))
        local_path = dir_path + '/GL Recon SAP'
        panwisedb_path = f'{local_path}/{temp_db_name}.db'

        if str(request_id) not in os.listdir(local_path):
            try:
                os.mkdir(os.path.join(local_path, str(request_id)))
            except:
                logging.error("run_gl_recon_sap",exc_info=True)
                traceback.print_exc()

        local_path = os.path.join(local_path, str(request_id))

        temp_db = sqlite3.connect(panwisedb_path)
        c = temp_db.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS masters (Particulars TEXT PRIMARY KEY, Status TEXT)")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('GSTIN','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('GL Code','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('Supply Type','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('Doc Type','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('Tax Code','Pending')")
        c.execute("INSERT INTO masters (Particulars,Status) VALUES ('TB','Pending')")
        c.close()
        temp_db.commit()

        # generate a shared access signiture for files and load them into Python
        for file in files_list:
            remote_download_path = os.path.join(stfp_inbound_path, file)
            
            # Create the local destination directory if it doesn't exist
            file_download_path = os.path.join(local_path, file)

            # os.makedirs(file_download_path, exist_ok=True)
            with open(os.path.join(file_download_path), "w") as file_temp:
                logging.error("run_gl_recon_sap",exc_info=True)
                traceback.print_exc()

            sftp.get(remote_download_path, file_download_path)

            main_onlyexecution.process_file(False, [temp_db_name, file, file_download_path, '/GL Recon SAP'])
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"UPDATE filestatus SET status = 'Processing Complete' where request_id = '{request_id}'")

        main_onlyexecution.valgldump(False, [temp_db_name, 'GL Recon SAP'])
        # main_onlyexecution.seqRun(False, [temp_db_name, ["False", "False", "False", "False", "False", "False", "False"], 'GL Recon SAP'])
        main_onlyexecution.valsr(False, [temp_db_name, 'GL Recon SAP'])
        main_onlyexecution.uploadsummary(False, [temp_db_name, 'GL Recon SAP'])

        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"UPDATE filestatus SET status = 'Validation Complete' where request_id = '{request_id}'")

        # call exportgaps from main and srg1 in reports to export srvsgl report
        # fromdate and todate in documentdate format
        output_file = main_onlyexecution.exportgaps(False, [temp_db_name, ['srgl'], fromTaxPeriod, toTaxPeriod, gstins, 'GL Recon SAP'])
        
        timestamp = datetime.utcnow()+ relativedelta(hours = 5.5)
        timestamp = timestamp.strftime("%d_%m_%Y_%H_%M_%f")
        csv_filename = f'SRvsGL_{request_id}_{timestamp}.csv'
        output_file.to_csv(os.path.join(local_path, csv_filename), index=False)

        sftp_file_path = os.path.join(stfp_outbound_path, request_id)
        if request_id not in sftp.listdir(stfp_outbound_path):
            sftp.mkdir(sftp_file_path)

        sftp.put(os.path.join(local_path, csv_filename), os.path.join(sftp_file_path, csv_filename))

        time_taken = datetime.now() - start_time
        execute_query(os.path.join(STORAGE_PATH, 'execution.db'), f"UPDATE filestatus SET status = 'Report Generated',time_taken = '{time_taken}' where request_id = '{request_id}'")

        response = requests.post(accept_zip_url, json=json.loads(json.dumps(payload, default=str)), headers=headers, stream=True)
        try:
            if os.path.exists(panwisedb_path):
                temp_db.close()
                os.remove(panwisedb_path)
        except:
            logging.error("run_gl_recon_sap",exc_info=True)
            traceback.print_exc()
        try:
            if os.path.exists(local_path):
                rmtree(local_path)
        except:
            logging.error("run_gl_recon_sap",exc_info=True)
            traceback.print_exc()

        # Create request_id folder in Archive if not present
        if request_id not in sftp.listdir(sftp_archive_path):
            sftp.mkdir(os.path.join(sftp_archive_path, str(request_id)))

        # Move files from the source directory to the destination directory
        for file in files_list:
            try:
                source_file = os.path.join(stfp_inbound_path, file)
                destination_path = os.path.join(sftp_archive_path, str(request_id), file)
                sftp.rename(source_file, destination_path)
            except:
                logging.error(f"run_gl_recon_sap-{source_file}-{destination_path}",exc_info=True)
                traceback.print_exc()

        # Delete request_id from Inbound GL Recon directory
        if request_id in sftp.listdir(os.path.join('SFTP_ROOT','INTG_INBOUND','GLRECON')):
            sftp.rmdir(os.path.join(stfp_inbound_path))

        # Close the SFTP session
        sftp.close()

        # Close the SSH connection
        client.close()

    except Exception as e:
        logging.error("run_gl_recon_sap",exc_info=True)
        payload = {
            "req": {
                "requestId": request_id,
                "status": "Failed",
                "errMsg": e.args
                    }
        }
        response = requests.post(accept_zip_url, json=json.loads(json.dumps(payload, default=str)), headers=headers, stream=True)
        traceback.print_exc()

@router.post('/initiateGlRecon.do')
async def initiateGlReconSAP(background_tasks: BackgroundTasks, payload = Body(...)):
    try:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # STORAGE_PATH = os.getenv('STORAGE_PATH')
        if 'execution.db' not in os.listdir(STORAGE_PATH):
            # create filestatus table
            execute_query(os.path.join(STORAGE_PATH, 'execution.db'), "CREATE TABLE IF NOT EXISTS filestatus (request_id TEXT PRIMARY KEY, temp_db_name TEXT, recd_time TEXT, time_taken TEXT, status TEXT)")

        if 'GL Recon SAP' not in os.listdir(dir_path):
            try:
                os.mkdir('GL Recon SAP')
            except:
                pass

        # initiate the gl recon
        request = payload.get("req")
        request_id = request.get("requestId")
        group_code = request.get("groupCode")
        sftp_username = request.get("sftpUsername")
        sftp_password = request.get("sftpPassword")
        fromTaxPeriod = request.get("fromTaxPeriod")
        toTaxPeriod = request.get("toTaxPeriod")
        gstins = list(map(lambda x: x.strip(), request.get("gstins").split(',')))

        # initiate gl recon process
        background_tasks.add_task(run_gl_recon_sap, request_id, group_code, sftp_username, sftp_password, fromTaxPeriod, toTaxPeriod, gstins)

        response_json = {
            "hdr": {"status": "S"},
            "resp" : "Received Succesfully"
        }
        status_code = 200
    except Exception as e:
        logging.error("initiateGlReconSAP",exc_info=True)
        response_json = {
            "hdr": {
                "status": "E"
            },
            "errMsg": e.args
        }
        status_code = 500

    return JSONResponse(content=response_json, status_code=status_code)
