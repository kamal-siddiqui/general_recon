# Inbuilt Libs
import hashlib
import uuid
import json
from datetime import datetime
from typing import Any, Optional, List
from enum import Enum
import re
# Third Party Libs
from fastapi import FastAPI, Request,APIRouter
import jwt
from Crypto.PublicKey import RSA
from Crypto.Random import get_random_bytes
from Crypto.Cipher import AES, PKCS1_OAEP
from base64 import b64encode
import os
# from jsonschema import validate, Draft7Validator,SchemaError,ValidationError
from pydantic import BaseModel, ValidationError, validator, conint, constr, confloat

jwt.api_jws.PyJWS.header_typ = False
# jwt.encode(..., headers={"alg": "HS512"})
api = APIRouter()
def get_current_folder():
    return os.path.dirname(os.path.abspath(__file__))

# with open(get_current_folder() + '/gstin_site_schema.json') as file:
#     schema_json = json.load(file)


private_key = """-----BEGIN RSA PRIVATE KEY-----
MIIEpAIBAAKCAQEAqK3lX9bk6hsG1BlxmCNfiABSMYN53LFNNDq7Svo4q23xpLgd
ExlxNo3Bz+uJeSyPxmHQ6mS9+zU8awyiOm9RfhSUKvzvEKmiC+RWGbZT+LG8V3pf
Fr7n/A0SeqkMcngVG6g5SSV9PSMzaUg6qssN7oaN9jWRVr8j8ccDbfMz2DsuamT1
6Qd7cNXTqMisozHCk1/QO3LWJjw86X3WCXDWtJEstV7mhySbeIXVeKDv2ppiZZKe
wt06i6XBjt283LC11J7dNaYz7DvJAhVbaEggpi4EnW5RdN8NQ209t3nNtFd//BsM
/n2Q23cfCa0/2sTEizYvI68E+nXcNuWOAj3wAQIDAQABAoIBAB3TwDFyEcEO3eaA
CEz3hlNJgT9FqTUb+hOwBgH0XLR9qMLwlp2TvGgB6aFvXDwb4+GX1uY3wbtr1r70
OTSdjhq+H1Q5rTl0UZYKPqplyhP6M9yBJFLkl6eDlT6w6WVNdCgTl2umkC6RLKxj
jBI6/T/uNaixoL1m579/ak0VOhf0V5qPBmznR505/m/Nw6oWGjrNdrw3dBLZvTYP
ALARUtWDYBQa7Cxl1AqZyG0p1D0GZ5EaPDc/LOsv05+YDrNLi6UXl24XWu8eOjCT
VCIZzeO6yZwpm0OVGtBBamobsNvXZKubIuwrMjucq82O3EnQZseAyvbVUv7Lphs5
+rbzTgUCgYEAzgy5/yNbsvkIsw5efsjznBoZo9KFj7GVu8IqV2p6lM6EuEodT8Qv
v6CGiQPTV/nJDOI2DEaZQM79PPW0L9lWHVNXurOB7O66nRyC084TksUjONo4Znwm
b4Jh3bRF7poq6O8ut8Jt3GgzY2JG7uQRsTtIqJf/Kj9dUG3nE4v5dbUCgYEA0ZH9
2PI5Cii8xXZz43dni9UBdeJQnWQlwqaLpgGtmUlKZlNjNE8D2TMyH12Nhx1kVR7x
bUILDlVw53GInxHZzfdZ7ugdndKYkFzZWja0vrWJaRo5s9RsjN5692cnXnrw6wnE
vaYuUwfE89/ZGfAsUm6dhTp+K+CTFtBSFhW1wJ0CgYEAzLngZBxhlCXT+vSf9yD6
y3MzXo5hnjA4MeHt6AUn3oqDXAhnr0Wim6eHhMOETbklheONCA1tX/NJsjP+4Dv1
UBXq6NpKkXtxd4FIi+IJmJ6/LFHGEC3ykoDddEcV5MjRMbfUl0hbl88AoBKZn+qD
mbDptHmxUey7bpqEKeu95LECgYEAwFwr0AVNm+iWlP1MFD8WeUBT7duEuWMiUc/D
IOYrbSbbtp7V0T6xvp0CZc3eSWYOIR+c5PeY5FhCoP4SNEgTTr26+9Js1N9oECJZ
kzfhoadJ8IIU8t6JoKfZ4Nr7RPq9xk+aGaW+oZHhEySlxuwwEp3b0l1FUIr7GBax
MfpNcPUCgYB8biYln3skRKKx7zkV+nQfQVhNIAsyFISYVteOdl5PK13bIZJgBC+t
2NG2CBpJkst73kwBOTeq2BqNPoHtx+JzndplrXWiEp2GOyB6dX7pJaGS3O2PFUyP
Rz1VVhK7PwbCuQ4t6+P8GpdTcMI/UzeAt4+a1ITaCUYcjh79sMPQ/A==
-----END RSA PRIVATE KEY-----"""
# private_key = open('private.pem').read()
# public_key = open('receiver.pem').read()

# IRN = SHA-256 HASH (Supplier GSTIN + Fin. Year + Doc Type + Doc Number)
# hashlib.sha256(b"29AAAPH9357H0002021-22INVSAM/011").hexdigest()

def get_fiscal_yr(dt_str: str) -> str:
    date_obj = datetime.strptime(dt_str, "%d/%m/%Y")
    if date_obj.month <= 3:
        return f"{date_obj.year - 1}-{date_obj.year % 100}"
    return f"{date_obj.year}-{(date_obj.year % 100) + 1}"

# Transaction Details - key -> `TranDtls`
class TaxScheme(str, Enum):
    GST = "GST"

class SupplyType(str, Enum):
    B2B = "B2B"
    SEZWP = "SEZWP"
    SEZWOP = "SEZWOP"
    EXPWP = "EXPWP"
    EXPWOP = "EXPWOP"
    DEXP = "DEXP"

class ReverseChargeYN(str, Enum):
    Y = "Y"
    N = "N"

class IntraIgstYN(str, Enum):
    Y = "Y"
    N = "N"

class TransactionDetails(BaseModel):
    TaxSch: TaxScheme
    SupTyp: SupplyType
    RegRev: Optional[ReverseChargeYN]
    EcmGstin: Optional[constr(regex="^([0-9]{2}[0-9A-Z]{13})$")]
    IgstOnIntra: Optional[IntraIgstYN]

# Document Details - key -> `DocDtls`
class DocumentTypeEnum(str, Enum):
    INV = "INV"
    CRN = "CRN"
    DBN = "DBN"

class DocumentDetails(BaseModel):
    Typ: DocumentTypeEnum
    No: constr(regex="^([a-zA-Z1-9]{1}[a-zA-Z0-9/-]{0,15})$")
    Dt: str
    
    @validator('Dt')
    def doc_dt_format_check(cls, value):
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`Dt` has invalid format')
        return value

# Seller Details - key -> `SellerDtls`
class SellerDetails(BaseModel):
    Gstin: constr(regex="^([0-9]{2}[0-9A-Z]{13})$")
    LglNm: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    TrdNm: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Addr1: constr(min_length=1, max_length=100, regex="^([^\\\"])*$")
    Addr2: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Loc: constr(min_length=3, max_length=50, regex="^([^\\\"])*$")
    Pin: conint(gt=100000, lt=999999)
    Stcd: constr(regex="^(?!0+$)([0-9]{1,2})$")
    Ph: constr(regex="^([0-9]{6,12})$")
    Em: Optional[constr(regex="^[a-zA-Z0-9+_.-]+@[a-zA-Z0-9.-]+$")]

# Buyer Details - key -> `BuyerDtls`
class BuyerDetails(BaseModel):
    Gstin: constr(regex="^(([0-9]{2}[0-9A-Z]{13})|URP)$")
    LglNm: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    TrdNm: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Pos: constr(regex="^(?!0+$)([0-9]{1,2})$")
    Addr1: constr(min_length=1, max_length=100, regex="^([^\\\"])*$")
    Addr2: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Loc: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    Pin: Optional[conint(gt=100000, lt=999999)]
    Stcd: constr(regex="^(?!0+$)([0-9]{1,2})$")
    Ph: Optional[constr(regex="^([0-9]{6,12})$")]
    Em: Optional[constr(regex="^[a-zA-Z0-9+_.-]+@[a-zA-Z0-9.-]+$")]

# Dispatch Details - key -> `DispDtls`
class DispatchDetails(BaseModel):
    Nm: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    Addr1: constr(min_length=1, max_length=100, regex="^([^\\\"])*$")
    Addr2: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Loc: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    Pin: conint(gt=100000, lt=999999)
    Stcd: constr(regex="^(?!0+$)([0-9]{1,2})$")

# Shipping Details - key -> `ShipDtls`
class ShippingDetails(BaseModel):
    Gstin: Optional[constr(regex="^(([0-9]{2}[0-9A-Z]{13})|URP)$")]
    LglNm: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    TrdNm: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Addr1: constr(min_length=1, max_length=100, regex="^([^\\\"])*$")
    Addr2: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Loc: constr(min_length=3, max_length=100, regex="^([^\\\"])*$")
    Pin: conint(gt=100000, lt=999999)
    Stcd: constr(regex="^(?!0+$)([0-9]{1,2})$")

# Item List Details - key -> `ItemList`
class IsService(str, Enum):
    Y = "Y"
    N = "N"

class BatchDetails(BaseModel):
    Nm: constr(min_length=3, max_length=20, regex="^([^\\\"])*$")
    ExpDt: Optional[str]
    WrDt: Optional[str]

    @validator('ExpDt', 'WrDt')
    def doc_dt_format_check(cls, value):
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`ExpDt` or `WrDt` has invalid format')
        return value

class AttributeDetails(BaseModel):
    Nm: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]
    Val: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]

class LineItem(BaseModel):
    SlNo: constr(min_length=1, max_length=6, regex="^([0-9]{1,6})$") # mandatory
    PrdDesc: Optional[constr(min_length=3, max_length=300, regex="^([^\\\"])*$")]
    IsServc: IsService # mandatory
    HsnCd: constr(min_length=4, max_length=8, regex="^(?!0+$)([0-9]{4}|[0-9]{6}|[0-9]{8})$") # mandatory
    Barcde: Optional[constr(min_length=3, max_length=30, regex="^([^\\\"])*$")]
    Qty: Optional[confloat(ge=0, le=9999999999.999)]
    FreeQty: Optional[confloat(ge=0, le=9999999999.999)]
    Unit: Optional[constr(regex="^([A-Z|a-z]{3,8})$")]
    UnitPrice: confloat(ge=0, le=999999999999.999) # mandatory
    TotAmt: confloat(ge=0, le=999999999999.99) # mandatory
    Discount: Optional[confloat(ge=0, le=999999999999.99)]
    PreTaxVal: Optional[confloat(ge=0, le=999999999999.99)]
    AssAmt: confloat(ge=0, le=999999999999.99) # mandatory
    GstRt: confloat(ge=0, le=999.999) # mandatory
    IgstAmt: Optional[confloat(ge=0, le=999999999999.99)]
    CgstAmt: Optional[confloat(ge=0, le=999999999999.99)]
    SgstAmt: Optional[confloat(ge=0, le=999999999999.99)]
    CesRt: Optional[confloat(ge=0, le=999.999)]
    CesAmt: Optional[confloat(ge=0, le=999999999999.99)]
    CesNonAdvlAmt: Optional[confloat(ge=0, le=999999999999.99)]
    StateCesRt: Optional[confloat(ge=0, le=999.999)]
    StateCesAmt: Optional[confloat(ge=0, le=999999999999.99)]
    StateCesNonAdvlAmt: Optional[confloat(ge=0, le=999999999999.99)]
    OthChrg: Optional[confloat(ge=0, le=999999999999.99)]
    TotItemVal: confloat(ge=0, le=999999999999.99) # mandatory
    OrdLineRef: Optional[constr(min_length=1, max_length=50, regex="^([^\\\"])*$")]
    OrgCntry: Optional[constr(regex="^([A-Z|a-z]{2})$")]
    PrdSlNo: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    BchDtls: Optional[BatchDetails]
    AttribDtls: Optional[List[AttributeDetails]]

# Value Details - key -> `ValDtls`
class ValueDetails(BaseModel):
    AssVal: confloat(ge=0, le=99999999999999.99)
    CgstVal: Optional[confloat(ge=0, le=99999999999999.99)]
    SgstVal: Optional[confloat(ge=0, le=99999999999999.99)]
    IgstVal: Optional[confloat(ge=0, le=99999999999999.99)]
    CesVal: Optional[confloat(ge=0, le=99999999999999.99)]
    StCesVal: Optional[confloat(ge=0, le=99999999999999.99)]
    Discount: Optional[confloat(ge=0, le=99999999999999.99)]
    OthChrg: Optional[confloat(ge=0, le=99999999999999.99)]
    RndOffAmt: Optional[confloat(ge=-99.99, le=99.99)]
    TotInvVal: confloat(ge=0, le=99999999999999.99)
    TotInvValFc: Optional[confloat(ge=0, le=99999999999999.99)]

# Payment Details - key -> `PayDtls`
class PaymentDetails(BaseModel):
    Nm: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]
    AccDet: Optional[constr(min_length=1, max_length=18, regex="^([^\\\"])*$")]
    Mode: Optional[constr(min_length=1, max_length=18, regex="^([^\\\"])*$")]
    FinInsBr: Optional[constr(min_length=1, max_length=11, regex="^([^\\\"])*$")]
    PayTerm: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]
    PayInstr: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]
    CrTrn: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]
    DirDr: Optional[constr(min_length=1, max_length=100, regex="^([^\\\"])*$")]
    CrDay: Optional[conint(ge=0, le=9999)]
    PaidAmt: Optional[confloat(ge=0, le=99999999999999.99)]
    PaymtDue: Optional[confloat(ge=0, le=99999999999999.99)]

# Ref Details - key -> `RefDtls`
class DocPerdDetails(BaseModel):
    InvStDt: str
    InvEndDt: str

    @validator('InvStDt', 'InvEndDt')
    def dt_format_check(cls, value):
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`InvStDt` or `InvEndDt` has invalid format')
        return value

class PrecDocDetails(BaseModel):
    InvNo: constr(regex="^[1-9a-zA-Z]{1}[0-9a-zA-Z/-]{1,15}$")
    InvDt: str
    OthRefNo: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]

    @validator('InvDt')
    def dt_format_check(cls, value):
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`InvDt` has invalid format')
        return value

class ControlDetails(BaseModel):
    RecAdvRefr: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    RecAdvDt: Optional[str]
    TendRefr: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    ContrRefr: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    ExtRefr: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    ProjRefr: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    PORefr: Optional[constr(min_length=1, max_length=16, regex="^([^\\\"])*$")]
    PORefDt: Optional[str]

    @validator('RecAdvDt', 'PORefDt')
    def dt_format_check(cls, value):
        if value is None:
            return value
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`RecAdvDt` or `PORefDt` has invalid format')
        return value

class RefDetails(BaseModel):
    InvRm: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    DocPerdDtls: Optional[DocPerdDetails]
    PrecDocDtls: Optional[List[PrecDocDetails]]
    ContrDtls: Optional[List[ControlDetails]]



# Additional Document Details - key -> `AddlDocDtls`
class AddDocDetails(BaseModel):
    Url: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    Docs: Optional[constr(min_length=3, max_length=1000, regex="^([^\\\"])*$")]
    Info: Optional[constr(min_length=3, max_length=1000, regex="^([^\\\"])*$")]

# Exp Details - key -> `ExpDtls`
class ExpDetails(BaseModel):
    ShipBNo: Optional[constr(min_length=1, max_length=20, regex="^([^\\\"])*$")]
    ShipBDt: Optional[str]
    Port: Optional[constr(regex="^[0-9|A-Z|a-z]{2,10}$")]
    RefClm: Optional[constr(regex="^([Y|N]{1})$")]
    ForCur: Optional[constr(regex="^[A-Z|a-z]{3,16}$")]
    CntCode: Optional[constr(regex="^([A-Z]{2})$")]
    ExpDuty: Optional[confloat(ge=0, le=999999999999.99)]

    @validator('ShipBDt')
    def trans_dt_format_check(cls, value):
        if value is None:
            return value
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`ShipBDt` has invalid format')
        return value

# E-way Bill Details - key -> `EwbDtls`
class TransportMode(str, Enum):
    road = "1"
    rail = "2"
    air = "3"
    ship = "4"

class VehicleType(str, Enum):
    O = "O" # ODC
    R = "R" # Regular

class EwbDetails(BaseModel):
    TransId: Optional[constr(regex="^([0-9]{2}[0-9A-Z]{13})$")]
    TransName: Optional[constr(min_length=3, max_length=100, regex="^([^\\\"])*$")]
    TransMode: Optional[TransportMode]
    Distance:  conint(ge=0, le=4000)
    TransDocNo: Optional[constr(regex="^([a-zA-Z0-9/-]{1,15})$")]
    TransDocDt: Optional[str]
    VehNo: Optional[constr(min_length=4, max_length=20, regex="^([A-Z|a-z|0-9]{4,20})$")]
    VehType: Optional[VehicleType]

    @validator('TransDocDt')
    def trans_dt_format_check(cls, value):
        if value is None:
            return value
        try:
            datetime.strptime(value, "%d/%m/%Y")
        except:
            raise ValueError('`TransDocDt` has invalid format')
        return value

# Main Json
class Invoice(BaseModel):
    Version: constr(min_length=1, max_length=6) # 1
    Irn: Optional[constr(min_length=64, max_length=64)] # 2
    TranDtls: TransactionDetails # 3
    DocDtls: DocumentDetails # 4
    SellerDtls: SellerDetails # 5
    BuyerDtls: BuyerDetails # 6
    DispDtls: Optional[DispatchDetails] # 7
    ShipDtls: Optional[ShippingDetails] # 8
    ItemList: List[LineItem] # 9
    ValDtls: ValueDetails # 10
    PayDtls: Optional[PaymentDetails] # 11
    RefDtls: Optional[RefDetails] # 12
    AddlDocDtls: Optional[List[AddDocDetails]] # 13
    ExpDtls: Optional[ExpDetails] # 14
    EwbDtls: Optional[EwbDetails] # 15


@api.post('/generate')
def generate_irn(invoice: Invoice):
    # print(invoice)
    # invoice = await request.json()
    # try:
    #     validate(invoice, schema_json, cls=Draft7Validator)
    #     print("LINE-44-Schema is valid JSON")
    # except (SchemaError, ValidationError) as e:
    #     print(e.json_path)
    #     # print(e.args)
    #     return {"message": "Invalid Json", "message": e.message, "path": e.json_path}
    # except Exception as e:
    #     print("Invalid Json")
    #     return {"message": "Invalid Json"}
    # except Exception as exc:
    #     print(exc.InvalidIndicatorValue(name=schema_json.name, value=schema_json.value, spec_type=type(schema_json)) )
    # validation_errors = sorted(validator.iter_errors(invoice), key=lambda e: e.path)

    # errors = []

    # for error in validation_errors:
    #     message = error.message
    #     if error.path:
    #         message = "[{}] {}".format(
    #             ".".join(str(x) for x in error.absolute_path), message
    #         )

    #     errors.append(message)
    # print(errors)
    # newrequest = json.dumps({})
    
    # Document Details
    doc_details = invoice.DocDtls
    doc_dt = doc_details.Dt
    doc_type = doc_details.Typ
    doc_no = doc_details.No
    fin_yr = get_fiscal_yr(doc_dt)
    
    # Seller/Buyer Details
    sup_gstin = invoice.SellerDtls.Gstin
    buyer_gstin = invoice.BuyerDtls.Gstin
    
    # Generating IRN/other details with existing details
    irn = hashlib.sha256((sup_gstin + fin_yr + doc_type + doc_no).encode('utf-8')).hexdigest()
    print("IRN", irn)
    ack_no = str(uuid.uuid4())
    act_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Invoice Line Items/Values
    doc_value = invoice.ValDtls.TotInvVal
    item_list = invoice.ItemList
    item_count = len(item_list)
    main_hsn = item_list[0].HsnCd if item_count > 0 else None

    # Signed Invoice Data
    invoice_to_be_signed = json.loads(invoice.json())
    invoice_to_be_signed['AckNo'] = ack_no
    invoice_to_be_signed['AckDt'] = act_dt
    invoice_to_be_signed['Irn'] = irn

    # QR Code Data
    qr_data = json.dumps({
        "SellerGstin": sup_gstin,
        "BuyerGstin": buyer_gstin,
        "DocNo": doc_no,
        "DocTyp": doc_type,
        "DocDt": doc_dt,
        "TotInvVal": doc_value,
        "ItemCnt": item_count,
        "MainHsnCode": main_hsn,
        "Irn": irn,
        "IrnDt": act_dt
    })

    signed_invoice = jwt.encode({'data': json.dumps(invoice_to_be_signed)}, private_key, algorithm="RS256",headers={"alg": "RS256","kid": "115F4426617A7938BE1BA06DBEE91A427584EDAB","typ": "JWT","x5t": "EV9EJmF6eTi-G6BtvukaQnWE7as"})
    signed_qr = jwt.encode({'data': qr_data, 'iss': 'ey'}, private_key, algorithm="RS256",headers={"alg": "RS256","kid": "115F4426617A7938BE1BA06DBEE91A427584EDAB","typ": "JWT","x5t": "EV9EJmF6eTi-G6BtvukaQnWE7as"})
    response_data = {'AckNo': ack_no, 'AckDt': act_dt, 'Irn': irn, 'SignedInvoice': signed_invoice, 'SignedQRCode': signed_qr, 'Status': 'ACT', 'EwbNo': None, 'EwbDt': None, 'EwbValidTill': None, 'Remarks': None}

    return response_data



# Response From API:
# {'Status': 1, 'ErrorDetails': None, 'Data': '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', 'InfoDtls': [{'InfCd': 'EWBERR', 'Desc': [{'ErrorCode': '4021', 'ErrorMessage': 'Transporter document date cannot be earlier than the invoice date.'}]}]}

# Decrypted Response
# {'AckNo': 212210144634396, 'AckDt': '2022-11-06 01:46:00', 'Irn': 'd0c091f072efe8ef71fe2a648690f053cf60b4cc160bc001b530058ad252c0f5', 'SignedInvoice': 'eyJhbGciOiJSUzI1NiIsImtpZCI6IkVEQzU3REUxMzU4QjMwMEJBOUY3OTM0MEE2Njk2ODMxRjNDODUwNDciLCJ0eXAiOiJKV1QiLCJ4NXQiOiI3Y1Y5NFRXTE1BdXA5NU5BcG1sb01mUElVRWMifQ.eyJkYXRhIjoie1wiQWNrTm9cIjoyMTIyMTAxNDQ2MzQzOTYsXCJBY2tEdFwiOlwiMjAyMi0xMS0wNiAwMTo0NjowMFwiLFwiSXJuXCI6XCJkMGMwOTFmMDcyZWZlOGVmNzFmZTJhNjQ4NjkwZjA1M2NmNjBiNGNjMTYwYmMwMDFiNTMwMDU4YWQyNTJjMGY1XCIsXCJWZXJzaW9uXCI6XCIxLjFcIixcIlRyYW5EdGxzXCI6e1wiVGF4U2NoXCI6XCJHU1RcIixcIlN1cFR5cFwiOlwiQjJCXCIsXCJSZWdSZXZcIjpcIllcIixcIklnc3RPbkludHJhXCI6XCJOXCJ9LFwiRG9jRHRsc1wiOntcIlR5cFwiOlwiSU5WXCIsXCJOb1wiOlwiU0FNLzAxMVwiLFwiRHRcIjpcIjE5LzA4LzIwMjFcIn0sXCJTZWxsZXJEdGxzXCI6e1wiR3N0aW5cIjpcIjI5QUFBUEg5MzU3SDAwMFwiLFwiTGdsTm1cIjpcIkVZIGNvbXBhbnkgcHZ0IGx0ZFwiLFwiVHJkTm1cIjpcIkVZIEluZHVzdHJpZXNcIixcIkFkZHIxXCI6XCJEaXZ5YXNyZWUgQ2hhbWJlcnNcIixcIkFkZHIyXCI6XCJMYW5nZm9yZCBHYXJkZW5zXCIsXCJMb2NcIjpcIkJlbmdhbHVydVwiLFwiUGluXCI6NTYwMDI1LFwiU3RjZFwiOlwiMjlcIixcIlBoXCI6XCI5MDAwMDAwMDAwXCIsXCJFbVwiOlwiYWJjQGdtYWlsLmNvbVwifSxcIkJ1eWVyRHRsc1wiOntcIkdzdGluXCI6XCIyOUFXR1BWNzEwN0IxWjFcIixcIkxnbE5tXCI6XCJYWVogY29tcGFueSBwdnQgbHRkXCIsXCJUcmRObVwiOlwiWFlaIEluZHVzdHJpZXNcIixcIlBvc1wiOlwiMTJcIixcIkFkZHIxXCI6XCI3dGggYmxvY2ssIGt1dmVtcHUgbGF5b3V0XCIsXCJBZGRyMlwiOlwia3V2ZW1wdSBsYXlvdXRcIixcIkxvY1wiOlwiR0FOREhJTkFHQVJcIixcIlBpblwiOjU2MjE2MCxcIlBoXCI6XCI5MTExMTExMTExMVwiLFwiRW1cIjpcInh5ekB5YWhvby5jb21cIixcIlN0Y2RcIjpcIjI5XCJ9LFwiRGlzcER0bHNcIjp7XCJObVwiOlwiQUJDIGNvbXBhbnkgcHZ0IGx0ZFwiLFwiQWRkcjFcIjpcIjd0aCBibG9jaywga3V2ZW1wdSBsYXlvdXRcIixcIkFkZHIyXCI6XCJrdXZlbXB1IGxheW91dFwiLFwiTG9jXCI6XCJCYW5hZ2Fsb3JlXCIsXCJQaW5cIjo1NjIxNjAsXCJTdGNkXCI6XCIyOVwifSxcIlNoaXBEdGxzXCI6e1wiR3N0aW5cIjpcIjI5QVdHUFY3MTA3QjFaMVwiLFwiTGdsTm1cIjpcIkNCRSBjb21wYW55IHB2dCBsdGRcIixcIlRyZE5tXCI6XCJrdXZlbXB1IGxheW91dFwiLFwiQWRkcjFcIjpcIjd0aCBibG9jaywga3V2ZW1wdSBsYXlvdXRcIixcIkFkZHIyXCI6XCJrdXZlbXB1IGxheW91dFwiLFwiTG9jXCI6XCJCYW5hZ2Fsb3JlXCIsXCJQaW5cIjo1NjIxNjAsXCJTdGNkXCI6XCIyOVwifSxcIkl0ZW1MaXN0XCI6W3tcIkl0ZW1Ob1wiOjAsXCJTbE5vXCI6XCIxXCIsXCJJc1NlcnZjXCI6XCJOXCIsXCJQcmREZXNjXCI6XCJSaWNlXCIsXCJIc25DZFwiOlwiMTAwMVwiLFwiQmFyY2RlXCI6XCIxMjM0NTZcIixcIlF0eVwiOjEwMC4zNDUsXCJGcmVlUXR5XCI6MTAsXCJVbml0XCI6XCJCQUdcIixcIlVuaXRQcmljZVwiOjk5LjU0NSxcIlRvdEFtdFwiOjk5ODguODQsXCJEaXNjb3VudFwiOjEwLFwiUHJlVGF4VmFsXCI6MSxcIkFzc0FtdFwiOjk5NzguODQsXCJHc3RSdFwiOjEyLjAsXCJJZ3N0QW10XCI6MTE5Ny40NixcIkNnc3RBbXRcIjowLFwiU2dzdEFtdFwiOjAsXCJDZXNSdFwiOjUsXCJDZXNBbXRcIjo0OTguOTQsXCJDZXNOb25BZHZsQW10XCI6MTAsXCJTdGF0ZUNlc1J0XCI6MTIsXCJTdGF0ZUNlc0FtdFwiOjExOTcuNDYsXCJTdGF0ZUNlc05vbkFkdmxBbXRcIjo1LFwiT3RoQ2hyZ1wiOjEwLFwiVG90SXRlbVZhbFwiOjEyODk3LjcsXCJPcmRMaW5lUmVmXCI6XCIzMjU2XCIsXCJPcmdDbnRyeVwiOlwiQUdcIixcIlByZFNsTm9cIjpcIjEyMzQ1XCIsXCJCY2hEdGxzXCI6e1wiTm1cIjpcIjEyMzQ1NlwiLFwiRXhwRHRcIjpcIjAxLzA4LzIwMjFcIixcIldyRHRcIjpcIjAxLzA5LzIwMjFcIn0sXCJBdHRyaWJEdGxzXCI6W3tcIk5tXCI6XCJSaWNlXCIsXCJWYWxcIjpcIjEwMDAwXCJ9XX1dLFwiVmFsRHRsc1wiOntcIkFzc1ZhbFwiOjk5NzguODQsXCJDZ3N0VmFsXCI6MCxcIlNnc3RWYWxcIjowLFwiSWdzdFZhbFwiOjExOTcuNDYsXCJDZXNWYWxcIjo1MDguOTQsXCJTdENlc1ZhbFwiOjEyMDIuNDYsXCJEaXNjb3VudFwiOjEwLFwiT3RoQ2hyZ1wiOjIwLFwiUm5kT2ZmQW10XCI6MC4zLFwiVG90SW52VmFsXCI6MTI5MDgsXCJUb3RJbnZWYWxGY1wiOjEyODk3Ljd9LFwiUGF5RHRsc1wiOntcIk5tXCI6XCJBQkNERVwiLFwiQWNjRGV0XCI6XCI1Njk3Mzg5NzEzMjEwXCIsXCJNb2RlXCI6XCJDYXNoXCIsXCJGaW5JbnNCclwiOlwiU0JJTjExMDAwXCIsXCJQYXlUZXJtXCI6XCIxMDBcIixcIlBheUluc3RyXCI6XCJHaWZ0XCIsXCJDclRyblwiOlwidGVzdFwiLFwiRGlyRHJcIjpcInRlc3RcIixcIkNyRGF5XCI6MTAwLFwiUGFpZEFtdFwiOjEwMDAwLFwiUGF5bXREdWVcIjo1MDAwfSxcIlJlZkR0bHNcIjp7XCJJbnZSbVwiOlwiVEVTVFwiLFwiRG9jUGVyZER0bHNcIjp7XCJJbnZTdER0XCI6XCIwMS8wOC8yMDIxXCIsXCJJbnZFbmREdFwiOlwiMDEvMDkvMjAyMVwifSxcIlByZWNEb2NEdGxzXCI6W3tcIkludk5vXCI6XCJET0MvMDAyXCIsXCJJbnZEdFwiOlwiMDEvMDgvMjAyMVwiLFwiT3RoUmVmTm9cIjpcIjEyMzQ1NlwifV0sXCJDb250ckR0bHNcIjpbe1wiUmVjQWR2UmVmclwiOlwiRG9jLzAwM1wiLFwiUmVjQWR2RHRcIjpcIjAxLzA4LzIwMjFcIixcIlRlbmRSZWZyXCI6XCJBYmMwMDFcIixcIkNvbnRyUmVmclwiOlwiQ28xMjNcIixcIkV4dFJlZnJcIjpcIllvNDU2XCIsXCJQcm9qUmVmclwiOlwiRG9jLTQ1NlwiLFwiUE9SZWZyXCI6XCJEb2MtNzg5XCIsXCJQT1JlZkR0XCI6XCIwMS8wOC8yMDIxXCJ9XX0sXCJBZGRsRG9jRHRsc1wiOlt7XCJVcmxcIjpcImh0dHBzOi8vZWludi1hcGlzYW5kYm94Lm5pYy5pblwiLFwiRG9jc1wiOlwiVGVzdCBEb2NcIixcIkluZm9cIjpcIkRvY3VtZW50IFRlc3RcIn1dLFwiRXhwRHRsc1wiOntcIlNoaXBCTm9cIjpcIkEtMjQ4XCIsXCJTaGlwQkR0XCI6XCIwMS8wOC8yMDIxXCIsXCJQb3J0XCI6XCJJTkFCRzFcIixcIlJlZkNsbVwiOlwiTlwiLFwiRm9yQ3VyXCI6XCJBRURcIixcIkNudENvZGVcIjpcIkFFXCJ9LFwiRXdiRHRsc1wiOntcIlRyYW5zSWRcIjpcIjEyQVdHUFY3MTA3QjFaMVwiLFwiVHJhbnNOYW1lXCI6XCJYWVogRVhQT1JUU1wiLFwiVHJhbnNNb2RlXCI6XCIxXCIsXCJEaXN0YW5jZVwiOjEwMCxcIlRyYW5zRG9jTm9cIjpcIkRPQzAxXCIsXCJUcmFuc0RvY0R0XCI6XCIxOC8wOC8yMDIxXCIsXCJWZWhOb1wiOlwia2ExMjM0NTZcIixcIlZlaFR5cGVcIjpcIlJcIn19IiwiaXNzIjoiTklDIn0.p2LPDZO5kbvpfbFXIvXSErKZgKZfz8EWEHbO1DFvogCWRFNaGrNN_mx1sVOeiT7BBBVX1Eytgb3nD0yL34B96hmWQ377c2P_SoMRRv4DcUtyBATSVrC9q1eKKqOUfm8yVqX5UJe9lkqCuSXcTwr35NkekN2-e0KmZHHnxxTYhcmO8IOPP14crpnFYov7nPznfijhdXLZd63cPxieeuStjWubVXXF7ZdJjQXbTtKP2irzm0EM75oZ2OLJMbyE3qXdaoFjk1TdBaeiJM1VuCGLVwhgUVAyG1zfrpxV3xHxGiMRgbvy5eYrixXco_AgxlWdZwgq-yjG-O6QC_EYXtZsvQ', 'SignedQRCode': 'eyJhbGciOiJSUzI1NiIsImtpZCI6IkVEQzU3REUxMzU4QjMwMEJBOUY3OTM0MEE2Njk2ODMxRjNDODUwNDciLCJ0eXAiOiJKV1QiLCJ4NXQiOiI3Y1Y5NFRXTE1BdXA5NU5BcG1sb01mUElVRWMifQ.eyJkYXRhIjoie1wiU2VsbGVyR3N0aW5cIjpcIjI5QUFBUEg5MzU3SDAwMFwiLFwiQnV5ZXJHc3RpblwiOlwiMjlBV0dQVjcxMDdCMVoxXCIsXCJEb2NOb1wiOlwiU0FNLzAxMVwiLFwiRG9jVHlwXCI6XCJJTlZcIixcIkRvY0R0XCI6XCIxOS8wOC8yMDIxXCIsXCJUb3RJbnZWYWxcIjoxMjkwOCxcIkl0ZW1DbnRcIjoxLFwiTWFpbkhzbkNvZGVcIjpcIjEwMDFcIixcIklyblwiOlwiZDBjMDkxZjA3MmVmZThlZjcxZmUyYTY0ODY5MGYwNTNjZjYwYjRjYzE2MGJjMDAxYjUzMDA1OGFkMjUyYzBmNVwiLFwiSXJuRHRcIjpcIjIwMjItMTEtMDYgMDE6NDY6MDBcIn0iLCJpc3MiOiJOSUMifQ.gdg7mQeyTa4W4wOpUxMQKpsbGinWT2D3iqRolpwD_s8CFXsoJVQskndkXVw4d3vuHh3K7MmdKCRvFTH5xyENuIFLWV3e115lIdrZRZ9fCUPlrZGrPXRolZyhEPQ68Pb0_AboS089pBYdasi0-WjCVMXC-uyW9-xh9DzsFenPSYS2IelD4Dc8MelBQtaDv-a9H1AmMIAtVp-5--mbyTStfHGVhBkGBDP9IEcpT6dSVM077eTMpufTqrx3OyQDUmA7KlfmxniIN0jkRuRCrmZkCr4T0Z6RU_AI9ULmuje0zCn6qe0gtHFnsisknHg1SBYAJlUdNJMpJoirlESg7HlN1w', 'Status': 'ACT', 'EwbNo': None, 'EwbDt': None, 'EwbValidTill': None, 'Remarks': None}