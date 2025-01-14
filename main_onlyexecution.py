from fastapi import Request, UploadFile, File, Response, BackgroundTasks, APIRouter
# from fastapi.middleware.cors import CORSMiddleware
# from fastapi.responses import JSONResponse
# from fastapi.concurrency import run_in_threadpool
from typing import List
import csv
import zipfile
# import string
# import tabula
import urllib
import logging
# import random
import time
# import requests
from openpyxl.utils.cell import get_column_letter
# from flask_paginate import Pagination, get_page_args
# from functools import wraps
# from googletrans import Translator
# from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from io import StringIO
from io import BytesIO
# from fuzzywuzzy import process
# from fuzzywuzzy import fuzz
import shutil as sh
# import PIL.ImageOps
from PIL import Image
# import qrcode
# from win32com import client
from openpyxl.utils.dataframe import dataframe_to_rows
# from openpyxl import Workbook, load_workbook
import openpyxl
import sqlite3
from datetime import date, datetime, timedelta
# import calendar
# import win32com.client
# from werkzeug.utils import secure_filename
# import xlsxwriter
# import concurrent.futures
# import pytesseract
# from pyzbar import pyzbar
# from shutil import copyfile
# import cv2
import fitz
import time
# import glob
# import threading
import base64
# import ast
# import jwt
import json
import re
# import copy
# import socket
# import hashlib
import os
import numpy as np
import pandas as pd
# from pandas.io.json import json_normalize
from pandas import options
# from os.path import exists
import sys
# from pydantic import BaseModel
# from dateutil.relativedelta import relativedelta, MO
from collections import OrderedDict
# from num2words import num2words
# import shutil
# from pypika import Query, Table,  Order, Column, Tables
# from pypika.terms import Star, CustomFunction, BasicCriterion, Field
import traceback
# from functools import reduce
# import operator
from base64 import b64encode, b64decode
# import urllib.parse

options.io.excel.xlsx.writer = 'xlsxwriter'
Image.MAX_IMAGE_PIXELS = None

router = APIRouter()

license_validity = "invalid"
router.message = ""  
clientPAN = ""
clientName = ""

dir_path = os.path.dirname(os.path.realpath(__file__))

def clientPAN_check(df, gstin_column):
    n = 't'
    if n == 't':
        drop_index_list = []
        return drop_index_list
    else:
        drop_index_list = df.iloc[np.where(np.logical_or(df[gstin_column] == None, np.logical_or(df[gstin_column] == '', df[gstin_column].str[2:12] != clientPAN)))].index
        # print(f'Length of GSTIN not matching with clientPAN : {len(drop_index_list)}')
        return drop_index_list

def checkInputFileColumns(lst):
    if len(lst) != 59:
        return 0
    else:
        return 0

def checkInputFileColumns1(lst):
    if len(lst) != 65:
        return 0
    else:
        return 0

def calc_inv_value(itms):
    total = {
        'txval': 0.0,
        'iamt': 0.0,
        'camt': 0.0,
        'samt': 0.0,
        'csamt': 0.0,
        'val': 0.0,
        'tax_rate': []
    }

    for itm in itms:
        for key in total.keys():
            try:
                if key in itm['itm_det']:
                    total[key] += itm['itm_det'][key]
                    total['val'] += itm['itm_det'][key]
                if 'rt' in itm['itm_det']:
                    if itm['itm_det']['rt'] not in total['tax_rate']:
                        total['tax_rate'].append(itm['itm_det']['rt'])
            except:
                if key in itm:
                    total[key] = itm[key]

                if 'rt' in itms:
                    if itms['rt'] not in total['tax_rate']:
                        total['tax_rate'].append(itms['rt'])
        # # print(total)
    return total

def calc_inv_value1(itms):
    total = {
        'txval': 0.0,
        'igst': 0.0,
        'cgst': 0.0,
        'sgst': 0.0,
        'cess': 0.0,
        'val': 0.0,
        'tax_rate': []
    }

    for itm in itms:
        for key in total.keys():
            try:
                if key in itm['itm_det']:
                    total[key] += itm['itm_det'][key]
                    total['val'] += itm['itm_det'][key]
                if 'rt' in itm['itm_det']:
                    if itm['itm_det']['rt'] not in total['tax_rate']:
                        total['tax_rate'].append(itm['itm_det']['rt'])
            except:
                if key in itm:
                    total[key] = itm[key]

                if 'rt' in itms:
                    if itms['rt'] not in total['tax_rate']:
                        total['tax_rate'].append(itms['rt'])
        # # print(total)
    return total

def calc_rating_PR(row):
    if row['TaxVal_PR']!=0:
        try:
            rating = (row['AbsoluteCount'] * 1 + row['MisMatchDocCount'] * 0.9 + row['MismatchCount'] * 0.75 + row['PotentialCount'] * 0.5 + row['Add2ACount'] * 0.25 - row['AddPRCount'] * 0.25) / row['TaxVal_PR']
            return rating
        except:
            rating = (row['AbsoluteCount'] * 1 + row['MisMatchDocCount'] * 0.9 + row['MismatchCount'] * 0.75 + row['PotentialCount'] * 0.5 + row['Add2BCount'] * 0.25 - row['AddPRCount'] * 0.25) / row['TaxVal_PR']
            return rating
    else:
        return 0.0

def calc_rating_2A(row):
    if row['TaxVal_2A']!=0:
        rating = (row['AbsoluteCount'] * 1 + row['MisMatchDocCount'] * 0.9 + row['MismatchCount'] * 0.75 + row['PotentialCount'] * 0.5 + row['Add2ACount'] * 0.25 - row['AddPRCount'] * 0.25) / row['TaxVal_2A']
        return rating
    else:
        return 0.0

def calc_rating_2B(row):
    if row['TaxVal_2B']!=0:
        rating = (row['AbsoluteCount'] * 1 + row['MisMatchDocCount'] * 0.9 + row['MismatchCount'] * 0.75 + row['PotentialCount'] * 0.5 + row['Add2BCount'] * 0.25 - row['AddPRCount'] * 0.25) / row['TaxVal_2B']
        return rating
    else:
        return 0.0
    
def calc_diff(row):
    if(row['AddPR_GST_PR'] is None): row['AddPR_GST_PR'] = 0
    if(row['Add2A_GST_2A'] is None): row['Add2A_GST_2A'] = 0
    diff = row['AddPR_GST_PR'] - row['Add2A_GST_2A']
    return diff 

def save_duplicate_files(df, filename, current_user): # this function will save all the duplicate files
    dir_path = os.path.dirname(os.path.realpath(__file__))
    time1 = datetime.now()
    timestamp = time1.strftime("%d-%m-%Y-%H-%M")
    save_path = f'''{dir_path}\\Client-Details\\{current_user}\\Duplicates dropped'''
    try:
        os.mkdir(save_path)
    except:
        pass
    df.to_csv(save_path +'\\'+ filename +' _'+ timestamp +'.csv', index=False, date_format='%d/%m/%Y')

# def save_duplicate_files_pr2b(df, filename, current_user): # this function will save all the duplicate files
#     dir_path = os.path.dirname(os.path.realpath(__file__))
#     folder_name = 'Duplicates dropped'

def updatecounters(clientPAN, username):
    dummydate = datetime(2021, 4, 1)
    path = dir_path + '/Client-Details'
    panwisedb = sqlite3.connect(f'{path}/{username}/{clientPAN}/{clientPAN}.db', timeout=10)
    try:
        query = "select RecipientGSTIN,DocumentDate,TaxableValue,GST from PurchaseRegisterDigi"
        prdata = pd.read_sql_query(query, panwisedb)
        prdata['DocumentDate'] = pd.to_datetime(
            prdata['DocumentDate'], dayfirst=True)
        prdata = prdata.groupby(
            [pd.Grouper(key='DocumentDate', freq='M'), "RecipientGSTIN"]).sum()
        prdata = prdata.reset_index()
        prdata = prdata.rename(columns={
                               'TaxableValue': 'PRTAXVAL', 'GST': 'PRITC', 'RecipientGSTIN': 'SupplierGSTIN'})
        prdata['category'] = "Purchase Register"
        if len(prdata) == 0:
            neworder = ['SupplierGSTIN', 'DocumentDate',
                        'PRTAXVAL', 'PRITC', 'category']

            prdata = pd.DataFrame(columns=neworder, data=[
                                  ['', dummydate, 0, 0, "Purchase Register"]])

    except:
        neworder = ['SupplierGSTIN', 'DocumentDate',
                    'PRTAXVAL', 'PRITC', 'category']

        prdata = pd.DataFrame(columns=neworder, data=[
                              ['', dummydate, 0, 0, "Purchase Register"]])
    # # print(prdata)
    # prdata.to_csv("GAPS\PR1.csv",index=False)
    try:
        query = "select CustomerGSTIN,DocumentDate,`Taxable Value`,Tax from GSTR_2A"
        twoadata = pd.read_sql_query(query, panwisedb)
        twoadata['DocumentDate'] = pd.to_datetime(
            twoadata['DocumentDate'], dayfirst=True)
        twoadata['Tax'] = pd.to_numeric(twoadata['Tax'], errors='coerce')
        twoadata['Taxable Value'] = pd.to_numeric(
            twoadata['Taxable Value'], errors='coerce')
        twoadata = twoadata.groupby(
            [pd.Grouper(key='DocumentDate', freq='M'), "CustomerGSTIN"]).sum()
        twoadata = twoadata.reset_index()
        twoadata = twoadata.rename(columns={
                                   'Taxable Value': '2ATAXVAL', 'Tax': '2AITC', 'CustomerGSTIN': 'SupplierGSTIN'})
        twoadata['category'] = "GSTR 2A"
    except:
        neworder = ['SupplierGSTIN', 'DocumentDate',
                    '2ATAXVAL', '2AITC', 'category']

        twoadata = pd.DataFrame(columns=neworder, data=[
                                ['', dummydate, 0, 0, "GSTR 2A"]])
    # # print(twoadata)
    try:
        query = "select CustomerGSTIN,DocumentDate,`Taxable Value`,Tax from GSTR_2B"
        twobdata = pd.read_sql_query(query, panwisedb)
        # # print(twobdata)
        twobdata['DocumentDate'] = pd.to_datetime(
            twobdata['DocumentDate'], dayfirst=True)
        twobdata = twobdata.groupby(
            [pd.Grouper(key='DocumentDate', freq='M'), "CustomerGSTIN"]).sum()
        twobdata = twobdata.reset_index()
        twobdata = twobdata.rename(columns={
                                   'Taxable Value': '2BTAXVAL', 'Tax': '2BITC', 'CustomerGSTIN': 'SupplierGSTIN'})
        twobdata['category'] = "GSTR 2B"
    except:
        neworder = ['SupplierGSTIN', 'DocumentDate',
                    '2BTAXVAL', '2BITC', 'category']

        twobdata = pd.DataFrame(columns=neworder, data=[
                                ['', dummydate, 0, 0, "GSTR 2B"]])

    try:
        query = "select SupplierGSTIN,DocumentDate,`Taxable Value`,Tax from GSTR_1"
        onedata = pd.read_sql_query(query, panwisedb)
        # # print(onedata)
        onedata['DocumentDate'] = pd.to_datetime(
            onedata['DocumentDate'], dayfirst=True)
        onedata = onedata.groupby(
            [pd.Grouper(key='DocumentDate', freq='M'), "SupplierGSTIN"]).sum()
        onedata = onedata.reset_index()
        onedata = onedata.rename(
            columns={'Taxable Value': '1TAXVAL', 'Tax': '1LIAB'})
        onedata['category'] = "GSTR 1"
    except:
        neworder = ['SupplierGSTIN', 'DocumentDate',
                    '1TAXVAL', '1LIAB', 'category']

        onedata = pd.DataFrame(columns=neworder, data=[
                               ['', dummydate, 0, 0, "GSTR 1"]])

    try:
        query1 = "select SupplierGSTIN,DocumentDate,TaxableValue,GST from SalesRegisterDigi"
        sales = pd.read_sql_query(query1, panwisedb)
        # # print(sales)
        sales['DocumentDate'] = pd.to_datetime(
            sales['DocumentDate'], dayfirst=True)
        sales = sales.groupby(
            [pd.Grouper(key='DocumentDate', freq='M'), "SupplierGSTIN"]).sum()
        sales = sales.reset_index()
        sales = sales.rename(
            columns={'TaxableValue': 'SRTAXVAL', 'GST': 'OUTLIAB'})
        sales['category'] = "Sales Register"
        #sales = sales.resample('MS',on='DocumentDate').sum()
    except:
        neworder = ['SupplierGSTIN', 'DocumentDate',
                    'SRTAXVAL', 'OUTLIAB', 'category']

        sales = pd.DataFrame(columns=neworder, data=[
                             ['', dummydate, 0, 0, "Sales Register"]])
    try:
        query1 = "select GSTIN,Document_Date_GL,Amount_GL from GLDump where GL_Type == 'Revenue'"
        revenuegl = pd.read_sql_query(query1, panwisedb)
        revenuegl['Document_Date_GL'] = pd.to_datetime(
            revenuegl['Document_Date_GL'], dayfirst=True)
        revenuegl = revenuegl.groupby(
            [pd.Grouper(key='Document_Date_GL', freq='M'), "GSTIN"]).sum()
        revenuegl = revenuegl.reset_index()
        revenuegl = revenuegl.rename(columns={
                                     'Document_Date_GL': 'DocumentDate', 'Amount_GL': 'GLTAXVAL', 'GSTIN': 'SupplierGSTIN'})
        revenuegl['GLTAXVAL'] = revenuegl['GLTAXVAL']*-1
        revenuegl['category'] = "Revenue GL"
    except:
        neworder = ['SupplierGSTIN', 'DocumentDate', 'GLTAXVAL', 'category']

        revenuegl = pd.DataFrame(columns=neworder, data=[
                                 ['', dummydate, 0, "Revenue GL"]])
    try:
        query1 = "select GSTIN,Document_Date_GL,Amount_GL from GLDump where GL_Type == 'Input'"
        itcgl = pd.read_sql_query(query1, panwisedb)

        itcgl['Document_Date_GL'] = pd.to_datetime(
            itcgl['Document_Date_GL'], dayfirst=True)
        itcgl = itcgl.groupby(
            [pd.Grouper(key='Document_Date_GL', freq='M'), "GSTIN"]).sum()
        itcgl = itcgl.reset_index()
        itcgl = itcgl.rename(columns={
                             'Document_Date_GL': 'DocumentDate', 'Amount_GL': 'ITCGL', 'GSTIN': 'SupplierGSTIN'})
        itcgl['ITCGL'] = itcgl['ITCGL']*1
        itcgl['category'] = "ITC GL"

    except:
        neworder = ['SupplierGSTIN', 'DocumentDate', 'ITCGL', 'category']

        itcgl = pd.DataFrame(columns=neworder, data=[
                             ['', dummydate, 0, "ITC GL"]])

    try:
        query1 = "select GSTIN,Document_Date_GL,Amount_GL from GLDump where GL_Type == 'CGST_GL' or GL_Type == 'SGST_GL' or GL_Type == 'IGST_GL' or GL_Type == 'UGST_GL' "
        outputtaxgl = pd.read_sql_query(query1, panwisedb)
        outputtaxgl['Document_Date_GL'] = pd.to_datetime(
            outputtaxgl['Document_Date_GL'], dayfirst=True)
        outputtaxgl = outputtaxgl.groupby(
            [pd.Grouper(key='Document_Date_GL', freq='M'), "GSTIN"]).sum()
        outputtaxgl = outputtaxgl.reset_index()
        outputtaxgl = outputtaxgl.rename(columns={
                                         'Document_Date_GL': 'DocumentDate', 'Amount_GL': 'OUTTAXGL', 'GSTIN': 'SupplierGSTIN'})
        outputtaxgl['OUTTAXGL'] = outputtaxgl['OUTTAXGL']*-1
        outputtaxgl['category'] = "Liability GL"
    except:
        neworder = ['SupplierGSTIN', 'DocumentDate', 'OUTTAXGL', 'category']

        outputtaxgl = pd.DataFrame(columns=neworder, data=[
                                   ['', dummydate, 0, "Liability GL"]])
    # # print(prdata)
    # # print(sales)
    merge = pd.DataFrame(columns=['SupplierGSTIN', 'DocumentDate'])
    try:
        merge = pd.merge(merge, sales, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = prdata.append(sales, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    try:
        merge = pd.merge(merge, prdata, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = prdata.append(sales, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    # # print(merge)
    try:
        merge = pd.merge(merge, revenuegl, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = merge1.append(revenuegl, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    try:
        merge = pd.merge(merge, itcgl, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = merge1.append(itcgl, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    try:
        merge = pd.merge(merge, outputtaxgl, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = merge1.append(outputtaxgl, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    # # print(merge)
    try:
        merge = pd.merge(merge, twoadata, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = merge1.append(twoadata, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    # # print(merge)
    try:
        merge = pd.merge(merge, twobdata, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = merge1.append(twobdata, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass
    try:
        merge = pd.merge(merge, onedata, on=[
                         'SupplierGSTIN', 'DocumentDate'], how='outer')
        merge = merge.fillna(0)
        merge1 = merge1.append(onedata, ignore_index=True, sort=False)
        merge1 = merge1.fillna(0)
    except:
        pass

    # # print(merge)
    try:
        merge.reset_index(inplace=True)
        merge1.reset_index(inplace=True)
    except:
        pass
    try:
        form_data = request.form()
        form_data = dict(form_data)
        merge.drop(merge[merge['SupplierGSTIN'] != str(
            form_data['chartType'])].index, inplace=True)
    except:
        merge = merge.resample('MS', on='DocumentDate').sum()
    try:
        merge.reset_index(inplace=True)
        merge1.reset_index(inplace=True)
    except:
        pass
    # # print(merge1)
    try:
        merge['DocumentDate'] = pd.to_datetime(merge['DocumentDate'], dayfirst=True).apply(
            lambda x: x.strftime('%b-%Y') if x else "")
        merge1['DocumentDate'] = pd.to_datetime(merge1['DocumentDate'], dayfirst=True).apply(
            lambda x: x.strftime('%b-%Y') if x else "")

    except:
        pass
    merge = merge.fillna(0)
    merge1 = merge1.fillna(0)
    merge1['Total'] = merge1.sum(axis=1)
    #merge = merge.div(1000000)
    # # print(merge1)
    # # print(merge)
    merge = merge.set_index('DocumentDate')
    try:
        merge['PRTAXVAL'] = merge['PRTAXVAL']/1000000
    except:
        pass
    try:
        merge['PRITC'] = merge['PRITC']/1000000
    except:
        pass
    try:
        merge['SRTAXVAL'] = merge['SRTAXVAL']/1000000
    except:
        pass
    try:
        merge['OUTLIAB'] = merge['OUTLIAB']/1000000
    except:
        pass
    try:
        merge['GLTAXVAL'] = merge['GLTAXVAL']/1000000
    except:
        pass
    try:
        merge['ITCGL'] = merge['ITCGL']/1000000
    except:
        pass
    try:
        merge['OUTTAXGL'] = merge['OUTTAXGL']/1000000
    except:
        pass
    try:
        merge['2ATAXVAL'] = merge['2ATAXVAL']/1000000
    except:
        pass
    try:
        merge['2AITC'] = merge['2AITC']/1000000
    except:
        pass
    try:
        merge['2BTAXVAL'] = merge['2BTAXVAL']/1000000
    except:
        pass
    try:
        merge['2BITC'] = merge['2BITC']/1000000
    except:
        pass
    try:
        merge['1TAXVAL'] = merge['1TAXVAL']/1000000
    except:
        pass
    try:
        merge['1LIAB'] = merge['1LIAB']/1000000
    except:
        pass

    merge = merge.round(2)
    merge.reset_index(inplace=True)
    # # print(merge)
    merge.to_sql("chart1", panwisedb, if_exists="replace", index=False)
    merge1.to_sql("chart2", panwisedb, if_exists="replace", index=False)

    return {"success": True}

# def uploadsummary(clientPAN, username):
def uploadsummary(is_local, parameters):
    # is_local true if localhost else false if Azure Blob Storage
    # parameters will be clientPAN and current_user if localhost else temp_db_name
    dir_path = os.path.dirname(os.path.realpath(__file__))
    folder_name = '' if is_local else parameters[1]
    # path = dir_path + '/Client-Details' if is_local else dir_path + folder_name
    path = os.path.join(dir_path, 'Client-Details') if is_local else os.path.join(dir_path, folder_name)

    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]

    # panwisedb_path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'
    panwisedb_path = os.path.join(path, current_user, clientPAN, f'{clientPAN}.db') if is_local else os.path.join(path, f'{temp_db_name}.db')

    panwisedb = sqlite3.connect(panwisedb_path, timeout=10)

    current = pd.DataFrame()
    try:
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM SalesRegisterDigi"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "Sales Register"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
        query = "select SUM(TaxableValue),COUNT(RecipientGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigi"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "Purchase Register"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(RecipientGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:
    
        query = "select SUM(`Taxable Value`),COUNT(SupplierGSTIN),SUM(`Central Tax Amount`),SUM(`State/UT Tax Amount`),SUM(`IGST Amount`), SUM(CessAmount) FROM GSTR_2A"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "GSTR 2A"
        upd=upd.rename(columns={'SUM(`Taxable Value`)':'Taxable_Value','SUM(`Central Tax Amount`)':'CGST','SUM(`State/UT Tax Amount`)':'SGST','SUM(`IGST Amount`)':'IGST','COUNT(SupplierGSTIN)':'Count','SUM(CessAmount)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
    
        query = "select SUM(`Taxable Value`),COUNT(SupplierGSTIN),SUM(`Central Tax Amount`),SUM(`State/UT Tax Amount`),SUM(`IGST Amount`), SUM(CessAmount) FROM GSTR_1"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "GSTR 1"
        upd=upd.rename(columns={'SUM(`Taxable Value`)':'Taxable_Value','SUM(`Central Tax Amount`)':'CGST','SUM(`State/UT Tax Amount`)':'SGST','SUM(`IGST Amount`)':'IGST','COUNT(SupplierGSTIN)':'Count','SUM(CessAmount)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:
        query = "select SUM(`Taxable Value`),COUNT(SupplierGSTIN),SUM(`Central Tax Amount`),SUM(`State/UT Tax Amount`),SUM(`IGST Amount`), SUM(CessAmount) FROM GSTR_2B"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "GSTR 2B"
        upd=upd.rename(columns={'SUM(`Taxable Value`)':'Taxable_Value','SUM(`Central Tax Amount`)':'CGST','SUM(`State/UT Tax Amount`)':'SGST','SUM(`IGST Amount`)':'IGST','COUNT(SupplierGSTIN)':'Count','SUM(CessAmount)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigiConso"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "Purchase Register - Conso"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
        
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigi"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "Purchase Register - Processed"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM SalesRegisterFlash"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "SR - Flash"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM SalesRegisterFlashError"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "SR - Flash Error"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:    
    
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterFlash"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "PR - Flash"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterFlashError"
        upd = pd.read_sql_query(query, panwisedb)
        
        upd['Particulars'] = "PR - Flash Error"
        upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
        neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
        upd=upd.reindex(columns=neworder)        
        upd=upd.fillna(0)
        current = upd.append(current)
    except:
        pass
    try:
        query = "select GL_Type,COUNT(Amount_GL),SUM(Amount_GL) FROM GLDump GROUP BY GL_Type"
        upd = pd.read_sql_query(query,panwisedb)
        upd["SUM(Amount_GL)"]= upd["SUM(Amount_GL)"].apply(pd.to_numeric)
        upd["COUNT(Amount_GL)"]= upd["COUNT(Amount_GL)"].apply(pd.to_numeric)
        #upd.drop(upd[upd['GL_Type']=="Input"].index, inplace=True)
        
        #upd = upd.transpose()

        upd1 = pd.DataFrame(columns=['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess'],data=[['GL Dump - Output',upd['COUNT(Amount_GL)'].sum(),upd[upd['GL_Type']=="Revenue"]['SUM(Amount_GL)'].sum(),upd[upd['GL_Type']=="CGST_GL"]['SUM(Amount_GL)'].sum(),upd[np.logical_or(upd['GL_Type']=="SGST_GL",upd['GL_Type']=="UGST_GL")]['SUM(Amount_GL)'].sum(),upd[upd['GL_Type']=="IGST_GL"]['SUM(Amount_GL)'].sum(),0]])

        
        current = upd1.append(current)
    except:
        pass
    try:
        query = "select GL_Type,COUNT(Amount_GL),SUM(Amount_GL) FROM GLDump GROUP BY GL_Type"
        upd = pd.read_sql_query(query,panwisedb)
        upd["SUM(Amount_GL)"]= upd["SUM(Amount_GL)"].apply(pd.to_numeric)
        upd["COUNT(Amount_GL)"]= upd["COUNT(Amount_GL)"].apply(pd.to_numeric)
        #upd.drop(upd[upd['GL_Type']=="Input"].index, inplace=True)
        
        #upd = upd.transpose()

        upd1 = pd.DataFrame(columns=['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess'],data=[['GL Dump - Input',upd['COUNT(Amount_GL)'].sum(),upd[upd['GL_Type']=="Expenses"]['SUM(Amount_GL)'].sum(),upd[upd['GL_Type']=="CGST_GL_Input"]['SUM(Amount_GL)'].sum(),upd[np.logical_or(upd['GL_Type']=="SGST_GL_Input",upd['GL_Type']=="UGST_GL_Input")]['SUM(Amount_GL)'].sum(),upd[upd['GL_Type']=="IGST_GL_Input"]['SUM(Amount_GL)'].sum(),0]])

        
        current = upd1.append(current)
    except:
        pass
    try:
        current['Total_Tax_Amount'] = current['Taxable_Value'] + current['CGST'] + current['SGST'] + current['IGST'] + current['Cess']
        current[['Count', 'Taxable_Value', 'CGST', 'SGST', 'IGST', 'Cess', 'Total_Tax_Amount']] = current[['Count', 'Taxable_Value', 'CGST', 'SGST', 'IGST', 'Cess', 'Total_Tax_Amount']].applymap(lambda x: round(x, 2))
        current.to_sql("Summary_Totals", panwisedb, if_exists="replace",index=False)
    except:
        pass
    return "OK"

# def seqRun(option_list, clientPAN):
def seqRun(is_local, parameters):
    # ---------------------Reading Input Inward SAP Dump File------------------------------
    # print("Reading Database File...")
    dir_path = os.path.dirname(os.path.realpath(__file__))
    MASTER_path = dir_path+"/"+"Flash"+"/"+"Master"+"/"
    Annexure_path = dir_path+"/"+"Flash"+"/"+"Annexure"+"/"

    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]
    option_list = '' if is_local else parameters[1]

    folder_name = '' if is_local else parameters[2]
    # path = dir_path + '/Client-Details' if is_local else dir_path + folder_name
    path = os.path.join(dir_path, 'Client-Details') if is_local else os.path.join(dir_path, folder_name)
    # path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'
    path = os.path.join(path, current_user, clientPAN, f'{clientPAN}.db') if is_local else os.path.join(path, f'{temp_db_name}.db')

    panwisedb = sqlite3.connect(path, timeout=10)

    try:
        query = "select DocKey from SalesRegisterFlash"
        doclist = pd.read_sql_query(query, panwisedb)
    except:
        doclist = pd.DataFrame(columns=['DocKey'])
    query = "select * FROM SalesRegisterDigi"
    #df = pd.read_sql_query(query, panwisedb)
    #df.drop(columns=['GST'], axis =1, inplace = True)
    hsn_goods = pd.read_excel(
        r""+MASTER_path+"Master file - HSN Code for goods.xlsx")
    hsn_services = pd.read_excel(
        r""+MASTER_path+"Master file - HSN codes for services.xlsx")
    uqc = pd.read_excel(r""+MASTER_path+"UOM Master.xlsx")
    port_code = pd.read_excel(
        r""+MASTER_path+"Service codes_state code and port code master.xlsx", sheet_name="Port Code")
    outward_doc_type = pd.read_excel(
        r""+MASTER_path+"Document type Supply type master (Outward).xlsx", sheet_name="Document type master")
    outward_supp_type = pd.read_excel(
        r""+MASTER_path+"Document type Supply type master (Outward).xlsx", sheet_name="Supply type master")
    outward_supp_doctype_combo = pd.read_excel(
        r""+MASTER_path+"Supplytype-DOCTYPE-Combination.xlsx", sheet_name="Outward")

    annexure = pd.read_excel(r""+Annexure_path+"DigiGST_Outward_format.xlsx")
    # print("Reading of Database File Completed...")
    for df in pd.read_sql_query(query, panwisedb, chunksize=10000):
        # print(df)
    # ------------------Checking Format and changing accordingly-----------------
        try:
            df.drop(df[df['DocKey'].isin(doclist['DocKey'])].index, inplace=True)
        except:
            pass
        try:
            if df.columns[0] != "SourceIdentifier":
                for i in range(10):

                    if df.iloc[0, 0] == "SourceIdentifier":
                        break
                    df.drop(index=i, inplace=True)
                df.columns = df.iloc[0, :]
                df.drop(index=i, inplace=True)
                df.reset_index(inplace=True)
                df.drop("index", axis=1, inplace=True)
        except:
            #messagebox.showwarning("warning","Error: File not in DigiGST Format",icon="warning")
            pass
            # print("Process Terminated...Please upload again")

        #progress['value'] = 20

        ListOfFilecolumns = df.columns

        flag = checkInputFileColumns(ListOfFilecolumns)

        if(flag == 1):
            #messagebox.showwarning("warning","Error: File not in DigiGST Format",icon="warning")
            # print("Process Terminated...Please upload again")
            return False
        else:
            # ------------------------Adding required columns -----------------
            # df.columns=annexure.columns
            ## print("Applying Validations...")
            df["EY Remark"] = "NA"
            df["Unaltered_Document_Number"] = df["DocumentNumber"]

            df.replace("NA", "NN", inplace=True)

            df["Flash Remark"] = "NA"

            def Return_period_change(df, key):
                if key:
                    # ----------------Changing Return Period-----------------
                    def checkReturnPeriodFormat(x):
                        regex = '(0[1-9]|10|11|12)20[0-9]{2}$'
                        try:
                            if(re.search(regex, str(x))):
                                return "Valid"
                            else:
                                return "Invalid"
                        except:
                            return "Invalid"
                    try:
                        df["ReturnPeriod"] = df["ReturnPeriod"].apply(
                            lambda x: "0"+str(x) if len(x) == 5 else str(x))
                    except:
                        pass
                    df["Temp"] = df["ReturnPeriod"].apply(
                        checkReturnPeriodFormat)
                    #df.loc[df["Temp"]=="Invalid","Flash Remark"] = df["Flash Remark"] + "| Invalid_Ret_Pd"

                    try:
                        df.loc[(df["ReturnPeriod"].str.len() != 6) & (
                            df["Temp"] == "Invalid"), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Ret_Pd"
                    except:
                        pass
                    df.drop('Temp', axis=1, inplace=True)
                else:
                    # -----------------Invalid ReturnPeriod--------------------------
                    def checkReturnPeriodFormat(x):
                        regex = '(0[1-9]|10|11|12)20[0-9]{2}$'
                        try:
                            if(re.search(regex, str(x))):
                                return "Valid"
                            else:
                                return "Invalid"
                        except:
                            return "Invalid"

                    df["Temp"] = df["ReturnPeriod"].apply(
                        checkReturnPeriodFormat)
                    try:
                        df.loc[(df["ReturnPeriod"].str.len() != 6) & (
                            df["Temp"] == "Invalid"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Ret_Pd"
                    except:
                        pass
                    df.drop('Temp', axis=1, inplace=True)

                return df

            def valid_date(datestring):

                if str(datestring).find("-") != -1:

                    try:

                        if len(str(datestring).split("-")[0]) == 4:

                            try:
                                int(str(datestring).split("-")[0])
                                if int(str(datestring).split("-")[1]) <= 12:
                                    return "ValidDate"
                                else:
                                    return "InvalidDate"
                            except:
                                return "InvalidDate"
                        else:

                            return "InvalidDate"
                    except:
                        return "InvalidDate"
                if str(datestring).find("//") != -1:

                    try:

                        if len(str(datestring).split("//")[0]) == 4:

                            try:
                                int(str(datestring).split("//")[0])
                                if int(str(datestring).split("//")[1]) <= 12:
                                    return "ValidDate"
                                else:
                                    return "InvalidDate"
                            except:
                                return "InvalidDate"
                        else:

                            return "InvalidDate"
                    except:
                        return "InvalidDate"

                if str(datestring).find(".") != -1:

                    try:

                        if len(str(datestring).split(".")[0]) == 4:

                            try:
                                int(str(datestring).split(".")[0])
                                if int(str(datestring).split(".")[1]) <= 12:
                                    return "ValidDate"
                                else:
                                    return "InvalidDate"
                            except:
                                return "InvalidDate"
                        else:

                            return "InvalidDate"
                    except:
                        return "InvalidDate"

            def Invalid_Document_Number(df, key):
                lst = ["!", "@", "#", ".", "$", "^",
                       "&", "*", "(", ")", "\\", " "]
                try:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x))
                except:
                    try:
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: int(x))
                    except:
                        pass

                lst1 = ["CR", "DR", "RNV"]
                if key:
                    try:
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace(".", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("&", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace(" ", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("^", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("@", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("!", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("#", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("$", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("%", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("*", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("\\", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace(")", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("(", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("?", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("_", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("<", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace(">", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("{", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("}", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("[", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("]", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("~", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace(":", ""))
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: str(x).replace("=", ""))
                    except:
                        pass
                    df.loc[(df["DocumentNumber"].str.len() > 16) | (
                        df["DocumentNumber"] == "NA"), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Doc_No"
                else:
                    def checkSpecialCharacter(x):
                        regex = re.compile('[@_!#$%^&*()<>?\|}{~:]= .')
                        try:
                            if(regex.search(str(x))):

                                return "Invalid"

                            if (str(x).find("?") != -1):
                                return "Invalid"
                            else:
                                return "Valid"
                        except:
                            return "Invalid"

                    df["Temp"] = df["DocumentNumber"].apply(
                        checkSpecialCharacter)
                    try:
                        df.loc[(df["Temp"] == "Invalid") | (df["DocumentNumber"].str.len() > 16) | (
                            df["DocumentNumber"] == "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Doc_No"
                    except:
                        pass
                    df.drop('Temp', axis=1, inplace=True)

                return df

            def Invalid_Original_document_Number(df, key):
                lst = ["!", "@", "#", ".", "$", "^",
                       "&", "*", "(", ")", "\\", " "]
                lst1 = ["CR", "DR", "RNV"]
                try:
                    df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                        lambda x: str(x))
                except:
                    try:
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: int(x))
                    except:
                        pass

                if key:
                    try:
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(".", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("&", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(" ", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("!", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("@", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("#", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("$", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("%", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("^", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("*", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("(", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(")", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("\\", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("?", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("_", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("<", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(">", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("{", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("}", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("[", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("]", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("~", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(":", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("=", ""))

                    except:
                        pass

                    df.loc[(df["OriginalDocumentNumber"] == "NA") & (df["DocumentType"].isin(
                        lst1)), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    df.loc[(df["OriginalDocumentNumber"].str.len() > 16) & (df["DocumentType"].isin(
                        lst1)), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Org_Doc_No"

                else:

                    lst1 = ["CR", "DR", "RNV"]

                    def checkSpecialCharacter(x):
                        regex = re.compile('[@_!#$%^&*()<>?\|}{~:]= .')
                        if(regex.search(str(x))):

                            return "Invalid"
                        else:

                            return "Valid"

                    df["Temp"] = df["OriginalDocumentNumber"].apply(
                        checkSpecialCharacter)
                    #df.loc[(df["Temp"]=="Invalid") & (df["DocumentType"].isin(lst1))  ,"Flash Remark"] = df["Flash Remark"] + "| Invalid_Org_Doc_No"
                    #df.loc[(df["OriginalDocumentNumber"]=="NA") & (df["DocumentType"].isin(lst1)),"Flash Remark"]=df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    #df.loc[(df["OriginalDocumentNumber"].str.len()>16) & (df["DocumentType"].isin(lst1)),"Flash Remark"]=df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    try:
                        df.loc[((df["OriginalDocumentNumber"].str.len() > 16) | (df["Temp"] == "Invalid")) & (
                            df["DocumentType"].isin(lst1)), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    except:
                        pass
                    df.drop('Temp', axis=1, inplace=True)

                return df

            # _______________Line Number Check nad change______________________

            def Line_number_change(df, key):
                if key:

                    df["LineNumber"] = df.index+1

                else:
                    df.loc[(df["LineNumber"] == "NA"),
                           "Flash Remark"] = df["Flash Remark"]+"| Invalid_Line_No"

                    LineSeries = df.groupby('DocumentNumber')[
                        'LineNumber'].nunique()
                    newdf = LineSeries.to_frame()
                    newdf.reset_index(inplace=True)

                    DocumentSeries = df.groupby('DocumentNumber')[
                        'DocumentNumber'].count()
                    newdf1 = DocumentSeries.to_frame()
                    newdf1.columns = ["DocumentNumber1"]

                    newdf1.reset_index(inplace=True)

                    for i in range(len(newdf)):
                        if newdf["LineNumber"][i] != newdf1["DocumentNumber1"][i]:
                            df.loc[(df["DocumentNumber"] == newdf["DocumentNumber"][i]),
                                   'Flash Remark'] = df['Flash Remark'] + "| Invalid_Line_No"

                    df.reset_index(drop=True, inplace=True)

                return df

            # ------------------------Invalid POS check and chnage----------------

            def POS_change(df, key):
                # ----------------Invalid POS------------------------------
                if key:

                    def int_change(x):
                        try:
                            if(len(str(int(x))) == 1):
                                return("0"+str(int(x)))
                            else:
                                return(str(int(x)))
                        except:
                            return x

                    df["POS"] = df["POS"].apply(lambda x: int_change(x))

                    df["POS"] = df["POS"].apply(
                        lambda x: "0"+str(x) if len(str(x)) == 1 else str(x))
                    df["BillToState"] = df["BillToState"].apply(
                        lambda x: int_change(x))
                    df["ShipToState"] = df["ShipToState"].apply(
                        lambda x: int_change(x))
                    #df["BillToState"]=df["BillToState"].apply(lambda x:str(int(x)) if x!="NA" else x)
                    #df["ShipToState"]=df["ShipToState"].apply(lambda x:str(int(x)) if x!="NA" else x)

                    listof_POS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["POS"].isin(
                        listof_POS), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"

                    def int_change(x):
                        try:

                            return(str(x).split(".")[0])
                        except:
                            return str(x)

                    lst = ["EXPWT", "EXPT"]
                    df["POS"] = df["POS"].apply(lambda x: int_change(x))
                    df.loc[(df["POS"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"

                    def int_change(x):
                        try:

                            return(str(x).split(".")[0])

                        except:
                            return str(x)

                    df["BillToState"] = df["BillToState"].apply(
                        lambda x: int_change(x))
                    listof_POS = ["NA", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["BillToState"].isin(listof_POS) & (
                        df["CustomerGSTIN"].str[0:2] != df["BillToState"]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_BillTo"
                    lst = ["EXPWT", "EXPT"]

                    df.loc[(df["BillToState"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_BillTo"
                    #df.loc[(df["BillToState"].str != "NA") & (df["BillToState"].str.len() != 2) & (df["SupplierGSTIN"].str[0:2] != df["BillToState"]) ,"Flash Remark"] = df["Flash Remark"] + "| Invalid_BillTo"

                    # ----------------Invalid ShipToState------------------------
                    df["ShipToState"] = df["ShipToState"].apply(
                        lambda x: int_change(x))
                    listof_POS = ["NA", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["ShipToState"].isin(
                        listof_POS), "Flash Remark"] = df["Flash Remark"] + "| Invalid_ShipTo"
                    #df.loc[(df["ShipToState"].str != "NA") & (df["ShipToState"].str.len() != 2) ,"Flash Remark"] = df["Flash Remark"] + "| Invalid_ShipTo"
                    lst = ["EXPWT", "EXPT"]

                    df.loc[(df["ShipToState"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_ShipTo"

                else:
                    listof_POS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df["POS"] = df["POS"].apply(
                        lambda x: "0"+str(x) if len(str(x)) == 1 else str(x))

                    df.loc[~df["POS"].isin(
                        listof_POS), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"

                    def int_change(x):
                        try:

                            return(str(x).split(".")[0])
                        except:
                            return str(x)

                    lst = ["EXPWT", "EXPT"]
                    df["POS"] = df["POS"].apply(lambda x: int_change(x))
                    df.loc[(df["POS"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"

                    def int_change(x):
                        try:

                            return(str(x).split(".")[0])

                        except:
                            return str(x)

                    df["BillToState"] = df["BillToState"].apply(
                        lambda x: int_change(x))
                    listof_POS = ["NA", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["BillToState"].isin(listof_POS) & (
                        df["CustomerGSTIN"].str[0:2] != df["BillToState"]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_BillTo"
                    lst = ["EXPWT", "EXPT"]

                    df.loc[(df["BillToState"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_BillTo"
                    #df.loc[(df["BillToState"].str != "NA") & (df["BillToState"].str.len() != 2) & (df["SupplierGSTIN"].str[0:2] != df["BillToState"]) ,"Flash Remark"] = df["Flash Remark"] + "| Invalid_BillTo"

                    # ----------------Invalid ShipToState------------------------
                    df["ShipToState"] = df["ShipToState"].apply(
                        lambda x: int_change(x))
                    listof_POS = ["NA", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["ShipToState"].isin(
                        listof_POS), "Flash Remark"] = df["Flash Remark"] + "| Invalid_ShipTo"
                    #df.loc[(df["ShipToState"].str != "NA") & (df["ShipToState"].str.len() != 2) ,"Flash Remark"] = df["Flash Remark"] + "| Invalid_ShipTo"
                    lst = ["EXPWT", "EXPT"]

                    df.loc[(df["ShipToState"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_ShipTo"

                return df

            def Date_check_change(df, key):
                # -----------------------------------Invalid Document Date------------------------------

                if key:

                    def all_format_change(x):

                        try:
                            if (str(x).split()[0].find(".") != -1):

                                if (len(str(x).split()[0].split(".")[0]) == 2) & (len(str(x).split()[0].split(".")[2]) == 4):

                                    return (str(x).split()[0].split(".")[2]+"-"+str(x).split()[0].split(".")[1]+"-"+str(x).split()[0].split(".")[0])
                                else:
                                    return (str(x).split()[0].split(".")[0]+"-"+str(x).split()[0].split(".")[1]+"-"+str(x).split()[0].split(".")[2])

                            if (str(x).split()[0].find("-") != -1):

                                if (len(str(x).split()[0].split("-")[0]) == 2) & (len(str(x).split()[0].split("-")[1]) == 2):

                                    return (str(x).split()[0].split("-")[2]+"-"+str(x).split()[0].split("-")[1]+"-"+str(x).split()[0].split("-")[0])
                                else:
                                    return (str(x).split()[0].split("-")[0]+"-"+str(x).split()[0].split("-")[1]+"-"+str(x).split()[0].split("-")[2])

                            else:
                                return x
                        except:
                            return x

                    def Chnage_Date_format(x):
                        try:
                            date = datetime.strptime(x, '%Y-%m-%d')

                            return(str(date).split()[0])
                        except:
                            return all_format_change(x)

                    df["DocumentDate"] = df["DocumentDate"].apply(
                        lambda x: Chnage_Date_format(x))
                    df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(
                        lambda x: Chnage_Date_format(x))
                    #df["PurchaseVoucherDate"]=df["PurchaseVoucherDate"].apply(lambda x:Chnage_Date_format(x))
                    #df["PaymentDate"]=df["PaymentDate"].apply(lambda x:Chnage_Date_format(x))
                    #df["BillOfEntryDate"]=df["BillOfEntryDate"].apply(lambda x:Chnage_Date_format(x))
                    df["ShippingBillDate"] = df["ShippingBillDate"].apply(
                        lambda x: Chnage_Date_format(x))
                    df["AccountingVoucherDate"] = df["AccountingVoucherDate"].apply(
                        lambda x: Chnage_Date_format(x))

                else:

                    df["Temp"] = df["DocumentDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[(df["Temp"] == "InvalidDate") | (df["Temp"] == "NA"),
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Doc_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["DocumentDate"] = df["DocumentDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)
                    # # ---------------------------------End Invalid Document Dates---------------------------

                    # -----------------------------------Invalid OriginalDocumentDate------------------------------
                    df["Temp"] = df["OriginalDocumentDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Org_Doc_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid OriginalDocument Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)
                    # # ---------------------------------End Invalid OriginalDocument Dates---------------------------

                    # -----------------------------------Invalid ShippingBillDate------------------------------
                    df["Temp"] = df["ShippingBillDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_SB_Date"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["ShippingBillDate"] = df["ShippingBillDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)
                    # # ---------------------------------End Invalid ShippingBillDate Dates---------------------------

                    # -----------------------------------Invalid AccountingVoucherDate------------------------------
                    df["Temp"] = df["AccountingVoucherDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Ac_Voucher_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["AccountingVoucherDate"] = df["AccountingVoucherDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)

                    # # ---------------------------------End Invalid ContractDate ---------------------------
                    '''
                    df["Temp"] = df["ContractDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
                    df.loc[(df["Temp"] == "InvalidDate"),"Flash Remark"] = df["Flash Remark"] + "| Invalid ContractDate"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate") | (df["Temp"] == "NA")]


                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True,inplace=True)
                    df["ContractDate"]=df["ContractDate"].apply(lambda x:str(x).split()[0] if x!="NA" else x)
                    '''

    #                #-----------------------------------Invalid PurchaseVoucherDate------------------------------
    #                 df["Temp"] = df["PurchaseVoucherDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
    #                 df.loc[df["Temp"] == "InvalidDate","Flash Remark"] = df["Flash Remark"] + "| Invalid PurchaseVoucherDate"
    #                 df2 = df[df["Temp"] == "InvalidDate"]
    #                 df1 = df[(df["Temp"] == "ValidDate") | (df["Temp"] == "NA")]

    #                 # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
    #                 frames = [df1, df2]
    #                 df = pd.concat(frames)
    #                 df.drop('Temp', axis=1, inplace=True)
    #                 df.reset_index(drop=True,inplace=True)
    #                 df["PurchaseVoucherDate"]=df["PurchaseVoucherDate"].apply(lambda x:str(x).split()[0] if x!="NA" else x)

    #                 # # ---------------------------------End Invalid PurchaseVoucherDate Dates---------------------------

    #                 #-----------------------------------Invalid BillOfEntryDate------------------------------
    #                 df["Temp"] = df["BillOfEntryDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
    #                 df.loc[df["Temp"] == "InvalidDate","Flash Remark"] = df["Flash Remark"] + "| Invalid BillOfEntryDate"
    #                 df2 = df[df["Temp"] == "InvalidDate"]
    #                 df1 = df[(df["Temp"] == "ValidDate") | (df["Temp"] == "NA")]

    #                 # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
    #                 frames = [df1, df2]
    #                 df = pd.concat(frames)
    #                 df.drop('Temp', axis=1, inplace=True)
    #                 df.reset_index(drop=True,inplace=True)

    #                 df["BillOfEntryDate"]=df["BillOfEntryDate"].apply(lambda x:str(x).split()[0] if x!="NA" else x)
    #                 # # ---------------------------------End Invalid BillOfEntryDate Dates---------------------------
    #                 #-----------------------------------Invalid PaymentDate------------------------------
    #                 df["Temp"] = df["PaymentDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
    #                 df.loc[df["Temp"] == "InvalidDate","Flash Remark"] = df["Flash Remark"] + "| Invalid PaymentDate"
    #                 df2 = df[df["Temp"] == "InvalidDate"]
    #                 df1 = df[(df["Temp"] == "ValidDate") | (df["Temp"] == "NA")]

    #                 # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
    #                 frames = [df1, df2]
    #                 df = pd.concat(frames)
    #                 df.drop('Temp', axis=1, inplace=True)
    #                 df.reset_index(drop=True,inplace=True)
    #                 df["PaymentDate"]=df["PaymentDate"].apply(lambda x:str(x).split()[0] if x!="NA" else x)
    #             # # ---------------------------------End Invalid PaymentDate Dates---------------------------

                return df

            # -------------------------itc VALIDATIONS---------------------------

            def ITC_validations(df, key):
                if key:
                    def checkPositiveandNegativeValuesFormat(x):
                        regex = '^-?(0|[1-9]\d*)(\.\d+)?$'
                        if(re.search(regex, str(x)) or x == "NA"):

                            return "Valid"
                        else:

                            return "Invalid"

                    df["Temp"] = df["AvailableIGST"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "|Invalid Available IGST"
                    df.drop('Temp', axis=1, inplace=True)
                    df["Temp"] = df["AvailableCGST"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "|Invalid Available CGST"
                    df.drop('Temp', axis=1, inplace=True)
                    df["Temp"] = df["AvailableSGST"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "|Invalid Available SGST"
                    df.drop('Temp', axis=1, inplace=True)
                    df["Temp"] = df["AvailableCess"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "|Invalid Available Cess"
                    df.drop('Temp', axis=1, inplace=True)

                    df.loc[(((df["AvailableIGST"].abs) > 0) | ((df["AvailableCGST"].abs) > 0) | ((df["AvailableSGST"].abs) > 0) | ((df["AvailableCess"].abs) > 0)) & (df["EligibilityIndicator"]
                                                                                                                                                                      == "NO"), "Flash Remark"] = df["Flash Remark"]+"| The derived value of the Total tax available as ITC from purchase register needs to be Nil if eligibility flag is None"
                    df["Temp"] = df["AvailableIGST"] + \
                        df["AvailableCGST"]+df["AvailableSGST"]
                    df.loc[df["Temp"] > df["TaxableValue"], "Flash Remark"] = df["Flash Remark"] + \
                        "Total tax available as ITC is more than the Total Tax Amount of this invoice"
                    df.loc[(df["Temp"].abs == 0) & (df["EligibilityIndicator"] != "NO"),
                           "Flash Remark"] = df["Flash Remark"]+"Available ITC amount is not as per Eligibility Indicator"
                    df.drop("Temp", inplace=True, axis=1)
                return df

            # ---------------------------Invoice value Validations----------------------------
            def Invoice_value_validaions(df, key):
                if key:
                    try:

                        df["InvoiceValue"] = df["TaxableValue"] + df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + \
                            df["StateUTTaxAmount"]+df["CessAmountAdvalorem"] + \
                            df["CessAmountSpecific"]
                    except:
                        pass
                '''
                else:
                    #--------------Line Number Check--------------------

                    df.loc[(df["InvoiceValue"]!=(df["TaxableValue"] + df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"]+df["CessAmountAdvalorem"]+df["CessAmountSpecific"])),"Flash Remark"]=df["Flash Remark"]+"| Invalid Invoice Value"

                    #------------Invoice Level Check------------------

                    df["Temp"]=df["TaxableValue"]+df["IntegratedTaxAmount"]+df["CentralTaxAmount"]+df["StateUTTaxAmount"]+df["CessAmountAdvalorem"]+df["CessAmountSpecific"]
                    LineSeries = df.groupby('DocumentNumber')['InvoiceValue','Temp'].sum()
                    newdf = LineSeries
                    newdf.reset_index(inplace=True)
                    for i in range(len(newdf)):


                        if newdf["InvoiceValue"][i]!=newdf["Temp"][i]:

                            df.loc[(df["DocumentNumber"] == newdf["DocumentNumber"][i]),'Flash Remark'] = df['Flash Remark'] + "| Invalid Invoice Value"


                    df.drop('Temp', axis=1, inplace=True)
                '''

                return df
            # ----------------------Replacing all values where blank is there with 0 -------------------------
            df["Quantity"].fillna(0, inplace=True)

            df["IntegratedTaxAmount"].fillna(0, inplace=True)
            df["StateUTTaxAmount"].fillna(0, inplace=True)
            df["CentralTaxAmount"].fillna(0, inplace=True)

            df["StateUTTaxRate"].fillna(0, inplace=True)
            df["IntegratedTaxRate"].fillna(0, inplace=True)
            df["CentralTaxRate"].fillna(0, inplace=True)

            # df["AvailableCGST"].fillna(0,inplace=True)
            # df["AvailableIGST"].fillna(0,inplace=True)
            # df["AvailableSGST"].fillna(0,inplace=True)
            # df["AvailableCess"].fillna(0,inplace=True)

            df["InvoiceValue"].fillna(0, inplace=True)
            df["TaxableValue"].fillna(0, inplace=True)

            df["CessRateAdvalorem"].fillna(0, inplace=True)
            df["CessAmountAdvalorem"].fillna(0, inplace=True)
            df["CessRateSpecific"].fillna(0, inplace=True)
            df["CessAmountSpecific"].fillna(0, inplace=True)

            df.fillna("NA", inplace=True)
            try:
                df["DocumentType"] = df["DocumentType"].apply(
                    lambda x: str(x).upper())
                df["SupplyType"] = df["SupplyType"].apply(
                    lambda x: str(x).upper())
            except:
                pass

            # ---------------Calling Option Functions----------------------

            zero_supply_type = ["NON", "EXT",
                                "NIL", "EXPWT", "NSY", "SEZ", "DXP"]
            Non_zero_supply_type = ["TAX", "EXPT", "ISD",
                                    "ISDIE", "ISD8", "ISIE8", "SOA", "LGAS", "DXP"]

            df = Return_period_change(df, option_list[4])
            df = Invalid_Document_Number(df, option_list[0])

            df = Invalid_Original_document_Number(df, option_list[0])
            df = Line_number_change(df, option_list[2])
            df = POS_change(df, option_list[3])

            df = Date_check_change(df, option_list[1])

            df = Invoice_value_validaions(df, option_list[5])

            df = Date_check_change(df, 0)
            # ----------------Changing Date columns to required format-----------------
            #df["DocumentDate"] = df["DocumentDate"].apply(lambda x:str(x).split()[0])
            #df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(lambda x:str(x).split()[0])
            #df["PurchaseVoucherDate"] = df["PurchaseVoucherDate"].apply(lambda x:str(x).split()[0])
            #df["BillOfEntryDate"] = df["BillOfEntryDate"].apply(lambda x:str(x).split()[0])
            #df["PaymentDate"] = df["PaymentDate"].apply(lambda x:str(x).split()[0])

            # ----------------Changing DocumentNumber,OriginaldocumentNumber to required Text format-----------------
            #df["DocumentNumber"] = df["DocumentNumber"].astype(str)
            #df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].astype(str)
            df["POS"] = df["POS"].astype(str)

            # df["EY Comments"] = np.nan

            # ------------------------Similar document number having multiple ShippingBillDate------------------------------------------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['ShippingBillDate'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ShippingBillDate"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"),
                       'Flash Remark'] = df["Flash Remark"] + "| Single_doc-Multiple_SB_Dt"

            df.drop('Temp', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple Shipping BillNo------------------------------------------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby(
                'Key_ID')['ShippingBillNumber'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ShippingBillNumber"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Multiple_SB_No"

            df.drop('Temp', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple return period------------------------------------------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['ReturnPeriod'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ReturnPeriod"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Multiple_Ret_Pd"

            df.drop('Temp', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple document dates------------------------------------------------

            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['DocumentDate'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["DocumentDate"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Multiple_Doc_Dt"

            df.drop('Temp', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple customer GSTIN------------------------------------------------
            #df["Temp"] = df["DocumentDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['SupplierGSTIN'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["SupplierGSTIN"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(df["Key_ID"] == i),
                                                                     'Flash Remark'] + "| Similar DocumentNumber cannot have multiple SupplierGSTINs"

            #df.drop('Temp', axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple Original Customer,Supplier GSTIN------------------------------------------------
            '''
            DocumentSeries = df.groupby('DocumentNumber')['OriginalSupplierGSTIN'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["OriginalSupplierGSTIN"] > 1]["DocumentNumber"].tolist()

            for i in ListOfDocumentNos:
                df.loc[df["DocumentNumber"] == i ,'Flash Remark'] = df.loc[df["DocumentNumber"] == i ,'Flash Remark'] + "| Similar DocumentNumber cannot have multiple SupplierGSTINs"
            '''
            # ------------------------Similar document number having multiple customer GSTIN------------------------------------------------
            #df["Temp"] = df["DocumentDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['CustomerGSTIN'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["CustomerGSTIN"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "| Single_doc-Multiple_Rec_GSTIN"

            #df.drop('Temp', axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple PORT Codes------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)

            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['PortCode'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["PortCode"] > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "| Single_doc-Multiple_Portcode"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple  BillOFEntry------------------------------------------------
            '''
            DocumentSeries = df.groupby('DocumentNumber')['BillOfEntry'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["BillOfEntry"] > 1]["DocumentNumber"].tolist()

            for i in ListOfDocumentNos:
                df.loc[df["DocumentNumber"] == i ,'Flash Remark'] = df.loc[df["DocumentNumber"] == i ,'Flash Remark'] + "| Similar DocumentNumber cannot have multiple BillOfEntry"


            # ------------------------Similar document number having multiple  BillOFEntry------------------------------------------------
            df["FY"]=df["DocumentDate"].apply(lambda x : x[0:4] if x!= "NA" else x)

            df["Key_ID"]=df["DocumentNumber"].map(str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['BillOfEntry'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)


            ListOfDocumentNos = newdf[newdf["BillOfEntry"] > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) ,'Flash Remark'] = df.loc[(df["Key_ID"] == i),'Flash Remark'] + "|  Similar DocumentNumber cannot have multiple BillOfEntry"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True,inplace=True)
            # ------------------------Similar document number having multiple  BillOFEntryDate------------------------------------------------
            df["FY"]=df["DocumentDate"].apply(lambda x : x[0:4] if x!= "NA" else x)

            df["Key_ID"]=df["DocumentNumber"].map(str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['BillOfEntryDate'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["BillOfEntryDate"] > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) ,'Flash Remark'] = df.loc[(df["Key_ID"] == i),'Flash Remark'] + "|  Similar DocumentNumber cannot have multiple BillOfEntryDate"


            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True,inplace=True)
            '''
            # ------------------------Similar document number having multiple POS------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['POS'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["POS"] > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "| Single_doc-Different_POS"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple ReverseChargeFlag------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby(
                'Key_ID')['ReverseChargeFlag'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ReverseChargeFlag"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(df["Key_ID"] == i),
                                                                     'Flash Remark'] + "|  Similar DocumentNumber cannot have multiple ReverseChargeFlag"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # --------------------------Document Numbers (CR or DR) having multiple OriginalReferences--------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)

            ListOfDocNos = df[((df["DocumentType"] == "CR") | (
                df["DocumentType"] == "DR"))]["Key_ID"].unique()
            for i in ListOfDocNos:
                if(df[df["Key_ID"] == i].count()[0] > 1):
                    df2 = df[df["Key_ID"] == i]
                    if(len(df2["OriginalDocumentNumber"].unique()) > 1):
                        df.loc[df["Key_ID"] == i, 'Flash Remark'] = df.loc[df["Key_ID"]
                                                                           == i, 'Flash Remark'] + "| Single_doc-Multiple_OrgDoc"
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            # --------------------------Document Numbers (CR or DR) having multiple OriginalReferences--------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            ListOfDocNos = df[((df["DocumentType"] == "CR") | (
                df["DocumentType"] == "DR"))]["Key_ID"].unique()
            for i in ListOfDocNos:
                if(df[df["Key_ID"] == i].count()[0] > 1):
                    df2 = df[df["Key_ID"] == i]
                    if(len(df2["OriginalDocumentDate"].unique()) > 1):
                        df.loc[df["Key_ID"] == i, 'Flash Remark'] = df.loc[df["Key_ID"]
                                                                           == i, 'Flash Remark'] + "| Single_doc-Multiple_Org_date"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)

            # --------------------------Document Numbers (CR or DR) having multiple OriginalCustomerGSTIN--------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)

            ListOfDocNos = df[((df["DocumentType"] == "CR") | (
                df["DocumentType"] == "DR"))]["Key_ID"].unique()
            for i in ListOfDocNos:
                if(df[df["Key_ID"] == i].count()[0] > 1):
                    df2 = df[df["Key_ID"] == i]
                    if(len(df2["OriginalCustomerGSTIN"].unique()) > 1):
                        df.loc[df["Key_ID"] == i, 'Flash Remark'] = df['Flash Remark'] + \
                            "| Single_doc-Multiple_Org_Cus_GSTIN"
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            # -----------------Applying Validations-------------------------

            def checkPositiveandNegativeValuesFormat(x):
                regex = '^-?(0|[1-9]\d*)(\.\d+)?$'
                if(re.search(regex, str(x)) or x == "NA"):

                    return "Valid"
                else:

                    return "Invalid"

            def checkPositiveValuesFormat(x):
                regex = '^(0|[1-9]\d*)(\.\d+)?$'
                if(re.search(regex, str(x)) or x == "NA"):
                    return "Valid"
                else:
                    return "Invalid"
            # -----------------Invalid TaxableValue-------------------------
            df["Temp"] = df["TaxableValue"].apply(
                checkPositiveandNegativeValuesFormat)
            #df.loc[(df["TaxableValue"]<=0) & (df["DocumentType"]!="CR"),"Flash Remark"] = df["Flash Remark"] + "| Invalid_Taxable_Val"

            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Taxable_Val"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid IntegratedTaxAmount-------------------
            df["Temp"] = df["IntegratedTaxAmount"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_IGST_Amt"
            df["Information Error"] = ""
            try:
                df["Temp"] = df["TaxableValue"]*df["IntegratedTaxRate"]/100
                df.loc[abs(df["Temp"]-df["IntegratedTaxAmount"]) > 0.1,
                       "Information Error"] = df["Information Error"]+"| Invalid ISGTAmount"
                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    pass

            # -----------------Invalid CentralTaxAmount-------------------
            df["Temp"] = df["CentralTaxAmount"].apply(
                checkPositiveandNegativeValuesFormat)

            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_CGST_Amt"
            try:
                df["Temp"] = df["TaxableValue"]*df["CentralTaxRate"]/100

                df.loc[abs(df["Temp"]-df["CentralTaxAmount"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"|    Invalid CGSTAmount"

                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    pass

            # -----------------Invalid StateUTTaxAmount-------------------
            df["Temp"] = df["StateUTTaxAmount"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_SGST/UTGST_Amt"
            try:
                df["Temp"] = df["TaxableValue"]*df["StateUTTaxRate"]/100

                df.loc[abs(df["Temp"]-df["StateUTTaxAmount"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"|    Invalid SGSTAmount"
                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    pass

            # -----------------Invalid CessAmountAdvalorem-------------------
            try:
                df["Temp"] = df["CessAmountAdvalorem"].apply(
                    checkPositiveandNegativeValuesFormat)
                df.loc[df["Temp"] == "Invalid",
                       "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cess_Amt"
                df["Temp"] = df["TaxableValue"]*df["CessRateAdvalorem"]/100

                df.loc[abs(df["Temp"]-df["CessAmountAdvalorem"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"|    Invalid CessAmountAdvalorem"
                df.drop('Temp', axis=1, inplace=True)
            except:
                pass
            # -----------------Invalid CessAmountSpecific--------------------
            try:
                df["Temp"] = df["CessAmountSpecific"].apply(
                    checkPositiveandNegativeValuesFormat)
                df.loc[df["Temp"] == "Invalid",
                       "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cess_Amt"
                df["Temp"] = df["TaxableValue"]*df["CessRateSpecific"]/100

                df.loc[abs(df["Temp"]-df["CessAmountSpecific"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"|    Invalid CessAmountSpecific"
                df.drop('Temp', axis=1, inplace=True)
            except:
                pass
            # -----------------Invalid Invoice Value--------------------
            df["Temp"] = df["InvoiceValue"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Invoice_Val"
            df.drop('Temp', axis=1, inplace=True)
            # -----------------Invalid ContractValue--------------------------

            '''

            df["Temp"] = df["ContractValue"].apply(checkPositiveandNegativeValuesFormat)
            #df.loc[(df["TaxableValue"]<=0) & (df["DocumentType"]!="CR"),"Flash Remark"] = df["Flash Remark"] + "| Invalid TaxableValue"

            df.loc[df["Temp"]=="Invalid","Flash Remark"] = df["Flash Remark"] + "| Invalid ContractValue"
            df.drop('Temp', axis=1, inplace=True)
            '''

    #         #-----------------Invalid CIFValue--------------------------
    #         df["Temp"] = df["CIFValue"].apply(checkPositiveandNegativeValuesFormat)
    #         df.loc[df["Temp"]=="Invalid","Flash Remark"] = df["Flash Remark"] + "|Invalid CIFValue"
    #         lst=["IMPG","SEZG"]
    #         df.loc[(df["CIFValue"]=="NA") & (df["DocumentType"].isin(lst)),"Flash Remark"] = df["Flash Remark"] + "|Invalid CIFValue"
    #         df.drop('Temp', axis=1, inplace=True)

    #         #-----------------Invalid CustomDuty--------------------------
    #         df["Temp"] = df["CustomDuty"].apply(checkPositiveandNegativeValuesFormat)
    #         df.loc[df["Temp"]=="Invalid","Flash Remark"] = df["Flash Remark"] + "|Invalid CustomDuty"
    #         df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid Quantity------------------------------
            df["Temp"] = df["Quantity"].apply(checkPositiveValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Qty"
            df.drop('Temp', axis=1, inplace=True)

            # ---------For Invalid IGSTRate------------------------
            '''

            try:
                df["Temp"] =  (df["IntegratedTaxAmount"]) / ((df["TaxableValue"]))*100
            except:
                df["Temp"] = 0
            df["Temp"] = df["Temp"].apply(lambda x: round(x,2))
            df["IntegratedTaxRate"] = df["IntegratedTaxRate"].apply(lambda x: round(x,2))
            df.loc[df["IntegratedTaxRate"]-df["Temp"]>0.1,"Flash Remark"] = df["Flash Remark"] + "| Invalid IntegratedTaxRate"
            '''
            df.loc[~df["IntegratedTaxRate"].isin(
                [0, 0.1, 3, 5, 12, 18, 28, 0.0, 3.0, 5.0, 12.0, 18.0, 28.0]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_IGST_Rt"
            #df.drop('Temp', axis=1, inplace=True)

            # ---------For Invalid SGSTRate--------------------
            '''
            try:
                df["Temp"] =  (df["StateUTTaxAmount"]) / ((df["TaxableValue"]))*100
            except:
                df["Temp"] = 0
            df["Temp"] = df["Temp"].apply(lambda x: round(x,2))
            df["StateUTTaxRate"] = df["StateUTTaxRate"].apply(lambda x: round(x,2))
            df.loc[(df["StateUTTaxRate"]-df["Temp"]>0.1),"Flash Remark"] = df["Flash Remark"] + "| Invalid StateUTTaxRate"
            '''
            df.loc[~df["StateUTTaxRate"].isin(
                [0, 0.05, 1.5, 2.5, 6, 9, 14, 6.0, 9.0, 14.0]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_SGST/UT_Rt"
            #df.drop('Temp', axis=1, inplace=True)
            # ---------For Invalid CGSTRate-------------------
            '''

            try:
                df["Temp"] =  (df["CentralTaxAmount"]) / ((df["TaxableValue"]))*100
            except:
                df["Temp"] = 0
            df["Temp"] = df["Temp"].apply(lambda x: round(x,2))
            df["CentralTaxRate"] = df["CentralTaxRate"].apply(lambda x: round(x,2))

            df.loc[(df["CentralTaxRate"]-df["Temp"]>0.1),"Flash Remark"] = df["Flash Remark"] + "| Invalid CentralTaxRate"
            '''
            df.loc[~df["CentralTaxRate"].isin(
                [0, 0.05, 1.5, 2.5, 6, 9, 14, 6.0, 9.0, 14.0]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_CGST_Rt"
            #df.drop('Temp', axis=1, inplace=True)

            # ----------------Invalid Pre-GST Flag-----------------

            lst = ["Y", "N", "NA", "Yes", "No", "NO", "YES"]
            df.loc[~df["CRDRPreGST"].isin(lst), "Flash Remark"] = df.loc[~df["CRDRPreGST"].isin(
                lst), "Flash Remark"]+"| Invalid_Pre_GST_flag"
            # ----------------Invalid Supplier GSTIN-----------------

            def GSTIN_check(x):
                regex = "(\d{2})(\D{5})(\d{4})(\D{1})(\w{3})"
                if re.search(regex, str(x)):
                    listof_POS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37"]
                    gstin_2 = str(x)[0:2]
                    if ((gstin_2) not in listof_POS):
                        return "InvalidGSTIN"
                    else:
                        return "ValidGSTIN"

                else:
                    return "InvalidGSTIN"

            df["Temp"] = df["SupplierGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)

            df.loc[(df["SupplierGSTIN"].str.len() != 15) | (df["Temp"] == "InvalidGSTIN"),
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Supplier GSTIN"

            # ----------------Invalid Customer or CustomerGSTIN GSTIN-----------------

            df["Temp"] = df["CustomerGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            #df.loc[df["Temp"]=="InvalidGSTIN","Flash Remark"]=df["Flash Remark"]+"| Invalid CustomerGSTIN"

            df.loc[((df["CustomerGSTIN"].str.len() != 15) | (df["Temp"] == "InvalidGSTIN")) & (
                df["CustomerGSTIN"] != "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid CustomerGSTIN"

            df.loc[((df["CustomerGSTIN"].str.len() != 15) | (df["Temp"] == "InvalidGSTIN")) & (df["CustomerGSTIN"] == "NA") & (
                df["SupplyType"] == "SEZ"), "Flash Remark"] = df["Flash Remark"] + "| Invalid CustomerGSTIN"

            # ----------------Invalid OriginalCustomerGSTIN-----------------
            df["Temp"] = df["OriginalCustomerGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            df.loc[df["Temp"] == "InvalidGSTIN",
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Org_CustGSTIN"
            df.loc[((df["OriginalCustomerGSTIN"].str.len() != 15) | (df["Temp"] == "InvalidGSTIN")) & (
                df["OriginalCustomerGSTIN"] != "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Org_CustGSTIN"

            # ----------------Invalid eComGSTIN-----------------
            df["Temp"] = df["eComGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            #df.loc[df["Temp"]=="InvalidGSTIN","Flash Remark"]=df["Flash Remark"]+"| Invalid_eComGSTIN"
            df.loc[((df["eComGSTIN"].str.len() != 15) | (df["Temp"] == "InvalidGSTIN")) & (df["eComGSTIN"] != "NA") & (
                df["eComGSTIN"] != "N"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_eComGSTIN"

            df.drop("Temp", axis=1, inplace=True)

            # ----------------Invalid CessRate------------------------------
            listof_CessRate = ["NA", "0", "60", "12", "71", "65", "61", "21", "5", "12.5", "72",
                               "17", "11", "290", "49", "160", "142", "20", "204", "96", "89", "15", "1", "3", "15"]
            df["CessRateSpecific"] = df["CessRateSpecific"].apply(
                lambda x: str(x).split(".")[0] if x != "NA" else x)
            df.loc[~df["CessRateSpecific"].isin(
                listof_CessRate), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cess_Rt"

    #         #----------------Invalid ReverseChargeFlag-----------------
    #         lst=["Y","N","NA","y","n"]
    #         df.loc[~df["ReverseChargeFlag"].isin(lst),"Flash Remark"] = df["Flash Remark"] + "| Invalid ReverseChargeFlag"

    #         #----------------Invalid ITC Flag-----------------
    #         df.loc[((df["ITCReversalIdentifier"] != "NA") & ((df["ITCReversalIdentifier"] != "T1") | (df["ITCReversalIdentifier"] != "T3"))),"Flash Remark"] = df["Flash Remark"] + "| Invalid ITCFlag"

    #         #----------------Invalid Bill Of Entry-----------------
    #         lst=["IMPG","SEZG"]
    #         df.loc[(df["BillOfEntry"] != "NA") & (df["BillOfEntry"].str.len() != 7),"Flash Remark"] = df["Flash Remark"] + "| Invalid Length of BillOfEntry Number "
    #         df.loc[(df["BillOfEntry"] == "NA") & (df["DocumentType"].isin(lst)),"Flash Remark"] = df["Flash Remark"] + "| Invalid BillOfEntry"

            # --------------------------------Supplier and Customer GSTIN are same---------------------------------------------------------
            try:
                df.loc[df["SupplierGSTIN"] == df["CustomerGSTIN"],
                       'Flash Remark'] = df["Flash Remark"]+"| Supplier and Recipient GSTIN are same"
            except:
                pass

    #         #----------------Invalid ITC Eligibility Indicator-----------------
    #         lst=["IS","IG","NO","CG"]
    #         df.loc[~df["EligibilityIndicator"].isin(lst),"Flash Remark"] = df["Flash Remark"] + "| Invalid ITC Eligibility Indicator"

            # -------------CGST and SGST Rates cannot be different-------------
            df.loc[(df["CentralTaxRate"] != df["StateUTTaxRate"]),
                   "Flash Remark"] = df["Flash Remark"] + "| CGST_rate≠SGST_rate"

            # -------------CGST and SGST Amounts cannot be different-------------
            df.loc[(df["CentralTaxAmount"] != df["StateUTTaxAmount"]),
                   "Flash Remark"] = df["Flash Remark"] + "| CGST≠SGST"

            # ------------Invalid document Number length-------------------------
            try:
                df.loc[df["DocumentNumber"].str.len(
                ) > 16, "Flash Remark"] = df["Flash Remark"] + "| Invalid_Doc_No"
            except:
                df["DocumentNumber"] = df["DocumentNumber"].apply(
                    lambda x: str(x))
                df.loc[df["DocumentNumber"].str.len(
                ) > 16, "Flash Remark"] = df["Flash Remark"] + "| Invalid_Doc_No"

            # ------------Invalid document Number length-------------------------
            try:
                df.loc[(df["OriginalDocumentNumber"].str.len() > 16) & ((df["OriginalDocumentNumber"] != "NA")),
                       "Flash Remark"] = df["Flash Remark"] + "| Invalid OriginalDocumentNumber-OriginalDocumentNumberDocumentNumber exceeds 16 digits"
            except:
                df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                    lambda x: str(x))
                df.loc[df["OriginalDocumentNumber"].str.len() > 16, "Flash Remark"] = df["Flash Remark"] + \
                    "| Invalid OriginalDocumentNumber-OriginalDocumentNumber exceeds 16 digits"
            # ------------Original Document number or Original Document Date is missing-------------
            lst1 = ["CR", "DR"]
            df.loc[(df["DocumentType"].isin(lst1)) & (df["OriginalDocumentNumber"]
                                                      == "NA"), 'Flash Remark'] = df["Flash Remark"]+"| Invalid_Org_Doc_No"
            df.loc[(df["DocumentType"].isin(lst1)) & (df["OriginalDocumentDate"]
                                                      == "NA"), 'Flash Remark'] = df["Flash Remark"]+"| Invalid_Org_Doc_Dt"

            # ------------Taxable Value should be 0 or BLANK for Exempt, Non-GST or NIL rated supply------------
            #df.loc[(((df["SupplyType"] =="EXT") | (df["SupplyType"] =="NON") | (df["SupplyType"] =="NSY")) & (df["TaxableValue"]!=0)),"Flash Remark"] = df["Flash Remark"] + "| TaxableValue should be 0 for EXT,NON,NSY supply"

            # ------------Taxable Value should not be 0 or BLANK for NIL rated supply-------------
            df.loc[((df["SupplyType"] == "NIL") & (df["TaxableValue"] == 0)),
                   "Flash Remark"] = df["Flash Remark"] + "| NIL rated cannot have taxable value"

            # ------------Invoice Value should not be 0 or BLANK for Exempt, Non-GST or NIL rated supply-------------
            df.loc[(((df["SupplyType"] == "EXT") | (df["SupplyType"] == "NON") | (df["SupplyType"] == "NSY") | (df["SupplyType"] == "NIL")) & (
                df["InvoiceValue"] == 0)), "Flash Remark"] = df["Flash Remark"] + "| Invoice Value should not be 0 for EXT,Non-GST,NIL supply"

    #         # ---------For Invalid CommonSupplyIndicator-------------
    #         lst=["YL","NL","Y","N","NA"]
    #         df.loc[~df["CommonSupplyIndicator"].isin(lst) ,"Flash Remark"] = df["Flash Remark"] + "| Invalid CommonSupplyIndicator"

            # ----------8 digit HSN is required for export/import------------------------

            # ----------In case of SEZ, CGST and SGST cannot be applied-----------------------------
            try:
                df.loc[(((df["SupplyType"] == "SEZG") | (df["SupplyType"] == "SEZS")) & (
                    (df["CentralTaxAmount"]+df["StateUTTaxAmount"]) != 0)), "Flash Remark"] = df["Flash Remark"] + "| SEZ_Nil CGST/SGST"
            except:
                pass

    #         #---------------------------------------
    #         df.loc[(((df["SupplyType"]=="NIL") | (df["SupplyType"]=="EXT") | (df["SupplyType"]=="NON") | (df["SupplyType"]=="NSY")) & ((df["CentralTaxAmount"]+df["StateUTTaxAmount"])!=0)) ,"Flash Remark"] = df["Flash Remark"] + "| Capital Goods flag cannot be there for services"

    #         #----------In case of Imports, CGST and SGST cannot be applied-----------------------------
    #         df.loc[(((df["SupplyType"]=="IMPG") | (df["SupplyType"]=="IMPS")) & ((df["CentralTaxAmount"]+df["StateUTTaxAmount"])!=0)) ,"Flash Remark"] = df["Flash Remark"] + "| In case of Imports, CGST and SGST cannot be applied"

            # ----------Capital Goods flag cannot be there for services-----------------------------
            #df["HSNorSAC"]=df["HSNorSAC"].apply(lambda x:str(x).split(".")[0] if x!="NA" else x)
            # # print((df["HSNorSAC"].str[0:2]))
    #         df.loc[((df["HSNorSAC"].str[0:2]=="99") & (df["EligibilityIndicator"]=="CG")) ,"Flash Remark"] = df["Flash Remark"] + "| Capital Goods flag cannot be there for services"

            # ----------Input flag cannot be there for services-----------------------------Input services flag cannot be there for goods
    #         df.loc[((df["HSNorSAC"].str[0:2]!="99") & (df["EligibilityIndicator"]=="IS")) ,"Flash Remark"] = df["Flash Remark"] + "| Input services flag cannot be there for goods"

            # ----------Input flag cannot be there for services-----------------------------
    #         df.loc[(df["HSNorSAC"].str[0:2]=="99") & (df["EligibilityIndicator"]=="IG") ,"Flash Remark"] = df["Flash Remark"] + "| Input flag cannot be there for services"

            # ----------Tax Rate and Tax Amount can not to be 0 or BLANK in case Supply Type is TAX or SEZ or DTA or DXP--------------
            try:

                df.loc[(df["SupplyType"].isin(Non_zero_supply_type)) & ((df["StateUTTaxAmount"]+df["CentralTaxAmount"]+df["IntegratedTaxAmount"] == 0) & (df["StateUTTaxRate"] +
                                                                                                                                                          df["CentralTaxRate"]+df["IntegratedTaxRate"] == 0)), "Flash Remark"] = df["Flash Remark"] + "| TaxRate/Tax Amt is NIL_SupplyType_TAXorSEZorDTA"

                df.loc[(df["SupplyType"].isin(zero_supply_type)) & ((df["StateUTTaxAmount"]+df["CentralTaxAmount"]+df["IntegratedTaxAmount"] != 0)
                                                                    & (df["StateUTTaxRate"]+df["CentralTaxRate"]+df["IntegratedTaxRate"] != 0)), "Flash Remark"] = df["Flash Remark"] + "| TaxAmt≠0"
            except:
                pass

            # ----------TaxableValue/TaxAmount cannot be negative as document types other than CR or RCR or RFV--------------
            try:
                df.loc[(((df["DocumentType"] != "CR") | (df["DocumentType"] != "RCR") | (df["DocumentType"] != "RFV")) & ((df["TaxableValue"] < 0) | (df["StateUTTaxAmount"] < 0) | (df["CentralTaxAmount"] < 0) | (
                    df["IntegratedTaxAmount"] < 0) | (df["InvoiceValue"] < 0) | (df["CessAmountSpecific"] < 0) | (df["CessAmountAdvalorem"] < 0))), "Flash Remark"] = df["Flash Remark"] + "| Values cannot be negative"
            except:
                pass

    #         #-------------------Original document to be reported in case of RFV---------------
    #         df.loc[(df["DocumentType"]=="RFV") & (df["OriginalDocumentNumber"]=="NA"),"Flash Remark"]=df["Flash Remark"] + "| Original document to be reported in case of RFV"

    #         #----------Reverse charge flag cannot be Y, when services are procured from SEZ--------------
    #         df.loc[(df["SupplyType"] == "SEZS") & (df["ReverseChargeFlag"] == "Y"),"Flash Remark"] = df["Flash Remark"] + "| Tax amount cannot be blank in case value in Cess amount is available"

            # ----------Tax amount cannot be blank in case value in Cess amount is available--------------
            df.loc[(df["CessAmountSpecific"]+df["CessRateAdvalorem"] != 0) & (df["TaxableValue"]
                                                                              == 0), "Flash Remark"] = df["Flash Remark"] + "| Invalid TaxAmt_CessAmt available"

            #df["Flash Remark"]=df["Flash Remark"].apply(lambda x:x.replace("NA|",""))

            # df.replace("NA","",inplace=True)

            # _____________________MAster Validations__________________________

            # -------------Unit of Measurement-----------------------
            # uqc["UQC1"]=""
            # uqc["UQC1"]=uqc["UQC"].copy()

            #uqc["UQC"]=uqc["UQC"].apply(lambda x:str(x).split("-")[0])
            # uqc["UQC"]=uqc["UQC"].map(str)+uqc["UQC1"].map(str)
            UQC = []
            for x in list(uqc["UQC"]):
                UQC.append(x)
                UQC.append(str(x).split("-")[0])

            df.loc[~df["UnitOfMeasurement"].isin(
                UQC), "Flash Remark"] = df["Flash Remark"]+"| Invalid_UOM"

            #df["UnitOfMeasurement"]=df["UnitOfMeasurement"].apply(lambda x:"OTH" if x=="NA" else x)
            # # print(UQC)

            if len(df[df["PortCode"] == "NA"]) != len(df):
                PCODE = list(port_code["Code"])
                df.loc[(~df["PortCode"].isin(PCODE)) & (df["PortCode"] != "NA"),
                       "Flash Remark"] = df["Flash Remark"]+"| Invalid_Portcode"

            doc_type = outward_doc_type["Corresponding Document type"]
            supp_type = outward_supp_type["Corresponding supply type"]

            # ---------------------Document Type and Supply Type validations from Masters------------------------------

            df.loc[(~df["DocumentType"].isin(doc_type)),
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Doc_Type"

            df.loc[(~df["SupplyType"].isin(supp_type)),
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Supply_Type"

            # ---------------------HSN Validation from Masters------------------------------

            hsn_g = hsn_goods["HSN Code"]

            hsn_s = hsn_services["Service Code (Tariff)"]
            hsn_list = []

            for x in hsn_g:
                hsn_list.append(str(x))
                hsn_list.append(str(x)[0:6])
                hsn_list.append(str(x)[0:4])
                hsn_list.append(str(x)[0:2])
            for x in hsn_s:
                hsn_list.append(str(x))
                hsn_list.append(str(x)[0:4])
                hsn_list.append(str(x)[0:2])

            # ------------8 digit HSN is required for export/import of Goods----------
            # ------------------  Invalid HSN or SAC--------------------------------------

            if(len(df[df["HSNorSAC"] == "NA"]) != len(df)):
                #df.loc[(~df["HSNorSAC"].isin(hsn_list)) & (df["HSNorSAC"]!="NA") ,"Flash Remark"]=df["Flash Remark"]+"| Invalid_HSN"
                df.loc[(~df["HSNorSAC"].isin(hsn_list)),
                       "Flash Remark"] = df["Flash Remark"]+"| Invalid_HSN"

                lst = ["EXPWT", "EXPT"]
                try:
                    df.loc[(df["SupplyType"].isin(lst)) & (df["HSNorSAC"].str[0:2] != "99") & (df["HSNorSAC"].str.len(
                    ) != 8), "Information Error"] = df["Information Error"] + "| Export_Import-8 digit HSN"
                except:
                    df["HSNorSAC"] = df["HSNorSAC"].map(str)
                    df.loc[(df["SupplyType"].isin(lst)) & (df["HSNorSAC"].str[0:2] != "99") & (df["HSNorSAC"].str.len(
                    ) != 8), "Information Error"] = df["Information Error"] + "| Export_Import-8 digit HSN"

            # ----------------------Replacing all values where blank is there with 0 -------------------------
            df["Quantity"].fillna(0, inplace=True)

            df["IntegratedTaxAmount"].fillna(0, inplace=True)
            df["StateUTTaxAmount"].fillna(0, inplace=True)
            df["CentralTaxAmount"].fillna(0, inplace=True)

            df["StateUTTaxRate"].fillna(0, inplace=True)
            df["IntegratedTaxRate"].fillna(0, inplace=True)
            df["CentralTaxRate"].fillna(0, inplace=True)

            # df["AvailableCGST"].fillna(0,inplace=True)
            # df["AvailableIGST"].fillna(0,inplace=True)
            # df["AvailableSGST"].fillna(0,inplace=True)

            df["InvoiceValue"].fillna(0, inplace=True)
            df["TaxableValue"].fillna(0, inplace=True)

            #progress['value'] = 30

            df.fillna("NA", inplace=True)

            # ----------------Changing Date columns to required format-----------------
            #df["DocumentDate"] = df["DocumentDate"].apply(lambda x:str(x).split()[0])
            #df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(lambda x:str(x).split()[0])
            #df["PurchaseVoucherDate"] = df["PurchaseVoucherDate"].apply(lambda x:str(x).split()[0])
            #df["BillOfEntryDate"] = df["BillOfEntryDate"].apply(lambda x:str(x).split()[0])
            #df["PaymentDate"] = df["PaymentDate"].apply(lambda x:str(x).split()[0])

            # ----------------Changing DocumentNumber,OriginaldocumentNumber to required Text format-----------------
            #df["DocumentNumber"] = df["DocumentNumber"].astype(str)
            #df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].astype(str)
            df["POS"] = df["POS"].astype(str)
            #progress['value'] = 40

            # ----------------Original invoice prior to FY 19-20-----------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["DocumentDate"].fillna("NA", inplace=True)
            #AfterMonthYearCompare = InputFinancialYear*100+9
            #PriorMonthYearToCompare = InputFinancialYear*100+4
            # PreviousMonthYearToCompare=(InputFinancialYear-1)*100+4

            df["OriginalDocumentDate"].fillna("NA", inplace=True)

            def FY1(x):
                try:
                    if x != "NA":
                        try:
                            return(int(int(x)/10000))
                        except:
                            return 0
                except:
                    return 0

            df["temp1"] = df["ReturnPeriod"].apply(lambda x: FY1(x))
            df["temp"] = ""
            for i in range(len(df)):
                if(df["DocumentType"][i] == "CR") & (df["Temp"][i] == "ValidDate"):
                    if(df["OriginalDocumentDate"][i] != "NA"):
                        try:
                            df["temp"][i] = int(df["OriginalDocumentDate"][i].split(
                                "-")[0] + df["OriginalDocumentDate"][i].split("-")[1])
                        except:
                            try:
                                df["temp"][i] = int(df["OriginalDocumentDate"][i].split(
                                    "//")[0] + df["OriginalDocumentDate"][i].split("//")[1])
                            except:
                                try:
                                    df["temp"][i] = int(df["OriginalDocumentDate"][i].split(
                                        ".")[0] + df["OriginalDocumentDate"][i].split(".")[1])
                                except:
                                    df["temp"][i] = 0
                        # if((df["temp"][i] > AfterMonthYearCompare) & (df["temp"][i] <PriorMonthYearToCompare) & (df["temp1"][i]>9)):

                            #df["Flash Remark"] = df["Flash Remark"][i]+ "| CN reorted after sep of next FY"
                        # if((df["temp"][i] > AfterMonthYearCompare) & (df["temp"][i] < PreviousMonthYearToCompare) & (df["temp1"][i]<=9)):

                            #df["Flash Remark"] = df["Flash Remark"][i]+ "| CN reorted after sep of next FY"
            df.drop('temp', axis=1, inplace=True)
            df.drop('Temp', axis=1, inplace=True)

            # ----------------2.7.Invoice are dated prior FY 19-20----------------------------------------------------------------------
            def transform(x):
                if str(x).find("-") != -1:
                    try:
                        return(int(str(x).split("-")[0]+str(x).split("-")[1]))
                    except:
                        return(0)
                else:
                    try:
                        return(int(str(x).split("/")[0]+str(x).split("/")[1]))
                    except:
                        return(0)

            df["temp"] = df["DocumentDate"].apply(lambda x: transform(x))

            #df["EY Remark"] = df["EY Remark"] + df["temp"].apply(lambda x:"| Pending" if x < PriorMonthYearToCompare else "")

            #df.loc[(df["temp"]<PriorMonthYearToCompare) & (df["temp1"]>9),"Flash Remark"]=df["Flash Remark"]+"| Invoice pretaining to Previous FY"

            #df.loc[(df["temp"]<PreviousMonthYearToCompare) & (df["temp1"]<=9),"Flash Remark"]=df["Flash Remark"]+"| Invoice pretaining to Previous FY"

            #df["Flash Remark"] = df["Flash Remark"]+ df["temp"].apply(lambda x:"| Invoice pretaining to Previous FY" if x < PriorMonthYearToCompare and x!=0 else "")
            df.drop('temp', axis=1, inplace=True)
            df.drop('temp1', axis=1, inplace=True)
            #progress['value'] = 50

            # ---------2.8.Invoice are dated further to the return period---------------------------------------------------------------

            def transform(x):

                if str(x).find("-") != -1:
                    try:
                        return(int(str(x).split("-")[0]+str(x).split("-")[1]))
                    except:
                        return(0)
                if str(x).find("//") != -1:
                    try:
                        return(int(str(x).split("/")[0]+str(x).split("/")[1]))
                    except:
                        return(0)
                if str(x).find(".") != -1:
                    try:
                        return(int(str(x).split(".")[0]+str(x).split(".")[1]))
                    except:
                        return(0)
                else:
                    return(0)

            def transform1(x):
                try:
                    if len(str(x)) == 5:
                        return("0"+str(x))
                    else:
                        return(str(x))
                except:
                    return(0)
            try:
                df["ReturnPeriodTemp"] = df["ReturnPeriod"].apply(
                    lambda x: transform1(x))
                df["ReturnPeriodTemp"] = df["ReturnPeriodTemp"].apply(lambda x: str(
                    x)[2:6]) + df["ReturnPeriodTemp"].apply(lambda x: str(x)[0:2])
                df["ReturnPeriodTemp"] = df["ReturnPeriodTemp"].apply(
                    lambda x: int(x) if x != "NA" else 0)
                df["DocumentDateTemp"] = df["DocumentDate"].apply(
                    lambda x: transform(x))
                current_date = transform(str(datetime.now()).split()[0])
                #df.loc[df["ReturnPeriodTemp"] < df["DocumentDateTemp"],'EY Remark'] = df["EY Remark"]+""
                df.loc[(df["ReturnPeriodTemp"] < df["DocumentDateTemp"]) & (df["ReturnPeriodTemp"] != 0) & (
                    df["DocumentDateTemp"] != 0), 'Flash Remark'] = df["Flash Remark"]+"| Inv beyond Ret_Pd"
                df.loc[(df["ReturnPeriodTemp"] > current_date) & (df["ReturnPeriodTemp"] != 0),
                       'Flash Remark'] = df["Flash Remark"]+"| TaxPeriod_beyond current date"

                df.drop("ReturnPeriodTemp", axis=1, inplace=True)
                df.drop("DocumentDateTemp", axis=1, inplace=True)
            except:
                try:
                    df.drop("ReturnPeriodTemp", axis=1, inplace=True)
                    df.drop("DocumentDateTemp", axis=1, inplace=True)
                except:
                    pass

            # -------------------------Cancelled Documents--------------------------
            dff1 = df[df["SupplyType"] == "CAN"]
            l2 = dff1['DocumentNumber'].unique().tolist()

            df["TaxableValue"] = pd.to_numeric(
                (df['TaxableValue']), errors='coerce')

            for j in l2:
                aa = df.loc[(df['DocumentNumber'] == j) & (
                    df['TaxableValue'] < 0), 'TaxableValue'].unique()
                for ii in aa:
                    count1 = df.loc[(df['DocumentNumber'] == j) & (
                        df['TaxableValue'] == abs(ii)), 'Flash Remark'].count()
                    count = df.loc[(df['DocumentNumber'] == j) & (
                        df['TaxableValue'] == ii), 'Flash Remark'].count()
                    if count1 >= count:
                        df.loc[(df['DocumentNumber'] == j) & (
                            df['TaxableValue'] == ii), 'Flash Remark'] = 'Cancelled Documents'
                        df.loc[(df['DocumentNumber'] == j) & (
                            df['TaxableValue'] == ii), 'EY Remark'] = 'Pending'
                        for i in range(len(df)):
                            if (df['DocumentNumber'].iloc[i] == j) & (df['TaxableValue'].iloc[i] == abs(ii)) & (count != 0):
                                df['Flash Remark'].iloc[i] = 'Cancelled Documents'
                                df['EY Remark'].iloc[i] = 'Pending'
                                count = count-1
                    if ((count1 < count) & (count1 != 0)):
                        df.loc[(df['DocumentNumber'] == j) & (df['TaxableValue'] == abs(
                            ii)), 'Flash Remark'] = 'Cancelled Documents'
                        df.loc[(df['DocumentNumber'] == j) & (
                            df['TaxableValue'] == abs(ii)), 'EY Remark'] = 'Pending'
                        for i in range(len(df)):
                            if (df['DocumentNumber'].iloc[i] == j) & (df['TaxableValue'].iloc[i] == (ii)) & (count1 != 0):
                                df['Flash Remark'].iloc[i] = 'Cancelled Documents'
                                df['EY Remark'].iloc[i] = 'Pending'
                                count1 = count1-1

            #      # print("Done with Cancelled Document Validation.....")
            # ------------------------Similar document number having multiple dates------------------------------------------------

            # ----------------Similar Original Document number having multiple original Document Dates------------------------------------------------
            OriginalDocSeries = df.groupby('OriginalDocumentNumber')[
                'OriginalDocumentDate'].nunique()
            newdf = OriginalDocSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfOriginalDocNo = newdf[newdf["OriginalDocumentDate"]
                                        > 1]["OriginalDocumentNumber"].tolist()

            for i in ListOfOriginalDocNo:
                df.loc[df["OriginalDocumentNumber"] == i,
                       'EY Remark'] = df.loc[df["OriginalDocumentNumber"] == i, 'EY Remark'] + ""
                df.loc[df["OriginalDocumentNumber"] == i, 'Flash Remark'] = df.loc[df["OriginalDocumentNumber"]
                                                                                   == i, 'Flash Remark'] + "| Similar OriginalDocumentNumber having multiple dates"

        #     # print("Done with Original Document having Similar Original DocDates ...Validation.....")

            # ----------------Adding initial zero if length of the below columns is 1 and to required Text format: POS, ReturnPeriod-----------------
            #df["POS"] = df["POS"].apply(lambda x:"0"+x if len(x) == 1 else x)
            #df["ReturnPeriod"] = df["ReturnPeriod"].apply(lambda x:"0"+x if len(x) == 5 else x)

            # -------------------------------------Originial invoice date and document number is missing-------------
            #df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & ((df["OriginalDocumentNumber"] == "NA") | (df["OriginalDocumentDate"] == "NA"))) ,'EY Remark'] = df["EY Remark"]+""
            #df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & ((df["OriginalDocumentNumber"] == "NA") | (df["OriginalDocumentDate"] == "NA"))) ,'Flash Remark'] = df["Flash Remark"]+"| Original document number/date is missing"

            df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & (
                (df["DocumentNumber"] == "NA") | (df["DocumentDate"] == "NA"))), 'EY Remark'] = df["EY Remark"]+""
            df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & ((df["DocumentNumber"] == "NA") | (
                df["DocumentDate"] == "NA"))), 'Flash Remark'] = df["Flash Remark"]+"| Document number/date is missing"

            # --------------------------------Supplier and Customer GSTIN are same---------------------------------------------------------
            # try:
            #    df.loc[df["SupplierGSTIN"] == df["CustomerGSTIN"] ,'Flash Remark'] = df["Flash Remark"]+"| Supplier and Recipient GSTIN are same"
            # except:
            #    pass

            # --------------------------------RCM Flag-------------------------------------------

            # -----------------------------POS is different from CUSTOMER GSTIN--------------------------------------------------------------------------
            #df.loc[(df["POS"] != "NA") & (df["POS"] != df["CustomerGSTIN"].apply(lambda x:str(x)[0:2])), "EY Remark"] = df["EY Remark"] + ""
            df.loc[(df["POS"] != "NA") & (df["POS"] != df["CustomerGSTIN"].apply(lambda x:str(x)[0:2])) & (df["CustomerGSTIN"] != "NA"),
                   "Information Error"] = df["Information Error"] + "| POS is different from CUSTOMER GSTIN - please confirm the CUSTOMER state"
            #df.loc[df["POS"] == "NA", "Information Error"] = df["Information Error"] + "| POS is blank- Kindly confirm"

            # ---------------------------Remaining Transactions:Stock Transfer--------------------------------------------------------------------------

            # --------------------------SupplyType1==EXT--------------------------------------------------------------------------
            #df.loc[((df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"] == 0)),"EY Remark"] = df["EY Remark"] + "| Ignore"
            #df.loc[((df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"] == 0)),"Flash Remark"] = df["Flash Remark"] + "| Exempt"
            #progress['value'] = 80

            try:

                df.loc[(df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"]
                                                      != 0), "Flash Remark"] = df["Flash Remark"] + "| Tax charged in exempt category supplies"

                # --------------------------SupplyType=NIL----------------------------

                df.loc[(df["SupplyType"] == "NIL") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                                                      df["StateUTTaxAmount"] == 0), "Flash Remark"] = df["Flash Remark"] + "| Nil rated"

            except:
                pass

            # ---------3.4.Eligibility Indicator-------------
            # df.loc[df["HSNorSAC"].apply(lambda x:str(x)[0:2] == 99),"EligibilityIndicator"] = "IS"
            # df.loc[df["HSNorSAC"].apply(lambda x:str(x)[0:2] != 99),"EligibilityIndicator"] = "IG"
            # df.loc[df["ReverseChargeFlag"] == "Y","EligibilityIndicator"] = "IS"
            # df.loc[df["SupplierGSTIN"].isin(ListOfIneligibleGSTINs),"EligibilityIndicator"] = "NO"

            # ---------3.9.For Invalid Tax Rate-------------
            #df["IntegratedTaxRate"] = df["IntegratedTaxRate"].round(2)
            #df.loc[~df["IntegratedTaxRate"].isin([0,28,18,12,5,0.0,28.0,18.0,12.0,5.0]),"Flash Remark"] = df["Flash Remark"] + "| Invalid Integrated Tax Rate "

            #df["CentralTaxRate"] = df["CentralTaxRate"].round(2)
            #df.loc[~df["CentralTaxRate"].isin([0,14,9,6,2.5,0.0,14.0,9.0,6.0]),"Flash Remark"] = df["Flash Remark"] + "| Invalid CentralTaxRate"

            #df["StateUTTaxRate"] = df["StateUTTaxRate"].round(2)
            #df.loc[~df["StateUTTaxRate"].isin([0,14,9,6,2.5,0.0,14.0,9.0,6.0]),"Flash Remark"] = df["Flash Remark"] + "| Invalid StateUTTaxRate"

            # ---------3.10.Invoice Value-------------

            # ---------3.11.Available taxes-----------

            # df.loc[((df["AvailableIntegratedTaxAmount"] != df["IntegratedTaxAmount"]) & (df["EligibilityIndicator"] != "NO")),"EY Remark"] = df["EY Remark"] + "| Mismatch in IntegratedTaxAmount & Available IntegratedTaxAmount"
            # df.loc[((df["AvailableCentralTaxAmount"] != df["CentralTaxAmount"]) & (df["EligibilityIndicator"] != "NO")),"EY Remark"] = df["EY Remark"] + "| Mismatch in CentralTaxAmount & Available CentralTaxAmount"
            # df.loc[((df["AvailableStateUTTaxAmount"] != df["StateUTTaxAmount"]) & (df["EligibilityIndicator"] != "NO")),"EY Remark"] = df["EY Remark"] + "| Mismatch in StateUTTaxAmount & Available StateUTTaxAmount"

            # -----------------------------------Reverse Charge Flag--------------------------
            #df.loc[(df["ReverseChargeFlag"] == "Y") & (df["SupplierGSTIN"] != "NA"),"DocumentType"] = "INV"

            # ---------3.3.Replacing special characters in Document Numbers,Original doc No-------------
            df["DocumentNumber"] = df["DocumentNumber"].apply(
                lambda x: str(x).replace(".", ""))
            df["DocumentNumber"] = df["DocumentNumber"].apply(
                lambda x: str(x).replace("&", ""))
            df["DocumentNumber"] = df["DocumentNumber"].apply(
                lambda x: str(x).replace(" ", ""))
            df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                lambda x: str(x).replace(".", ""))
            df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                lambda x: str(x).replace("&", ""))
            df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                lambda x: str(x).replace(" ", ""))
            # df["DocumentNumber"]=df["DocumentNumber"].apply(lambda x:str(x).replace("\\","/"))
            # df["DocumentNumber"]=df["DocumentNumber"].apply(lambda x:str(x).replace("00:00:00",""))
            # df["DocumentNumber"]=df["DocumentNumber"].apply(lambda x:str(x).split("Dt")[0])

            # ---------3.12.Import of goods------------------------
            '''
            df.loc[((df["SupplyType"] == "IMPG") & ((df["PortCode"] == "NA") | (df["BillOfEntry"] == "NA") | (df["BillOfEntryDate"] == "NA"))),"EY Remark"] = df["EY Remark"]+""
            df.loc[((df["SupplyType"] == "IMPG") & ((df["PortCode"] == "NA") | (df["BillOfEntry"] == "NA") | (df["BillOfEntryDate"] == "NA"))),"Flash Remark"] = df["Flash Remark"]+"| Port code/Bill of entry details missing"
            '''

            # ---------2.10.Import Of Goods-------------

            # ---------2.14.No supplier GSTIN-------------

            #df.loc[(df["SupplierGSTIN"] == "NA") & (df["EY Remark"] != "Import of goods") & (df["EY Remark"] != "Import of services") & (df['EY Remark'].str.contains("RCM")==False),'EY Remark'] = df["EY Remark"]+"| Ignore"
            df.loc[(df["SupplierGSTIN"] == "NA"),
                   'Flash Remark'] = df["Flash Remark"]+"| No supplier GSTIN"

            #df.loc[((df["SupplierGSTIN"] != "NA") & (df["SupplierGSTIN"].str.len() != 15) & (df["SupplierGSTIN"] != "URP")),'EY Remark'] = df["EY Remark"]+""
            #df.loc[((df["SupplierGSTIN"] != "NA") & (df["SupplierGSTIN"].str.len() != 15) & (df["SupplierGSTIN"] != "URP")),'Flash Remark'] = df["Flash Remark"]+"| SupplierGSTIN is not proper"

            # ----------IGST,CSGST,SGST - Invalid Taxes applied-------------

            try:

                df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] == df["POS"]) & (
                    df["IntegratedTaxAmount"] != 0) & ((df["CentralTaxAmount"] == 0) & (df["StateUTTaxAmount"] == 0)), 'Flash Remark'] = df["Flash Remark"]+"| IGST cannot be charged"

                #df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] != df["POS"]) & (df["IntegratedTaxAmount"] == 0) & ((df["CentralTaxAmount"] != 0) & (df["StateUTTaxAmount"] != 0)) ,'Information Error'] = df["Information Error"]+"| Why CGST+SGST charged instead of IGST"
            except:
                try:
                    df["SupplierGSTIN"] = df["SupplierGSTIN"].apply(
                        lambda x: str(x) if x != "NA" else x)
                    df["RecipientGSTIN"] = df["RecipientGSTIN"].apply(
                        lambda x: str(x) if x != "NA" else x)
                    df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] == df["POS"]) & (
                        df["IntegratedTaxAmount"] != 0) & ((df["CentralTaxAmount"] == 0) & (df["StateUTTaxAmount"] == 0)), 'Flash Remark'] = df["Flash Remark"]+"| IGST cannot be charged"

                    #df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] != df["RecipientGSTIN"].str[0:2]) & (df["IntegratedTaxAmount"] == 0) & ((df["CentralTaxAmount"] != 0) & (df["StateUTTaxAmount"] != 0)) ,'Information Error'] = df["Information Error"]+"| Why CGST+SGST charged instead of IGST"
                except:
                    pass

            # ---------New - CR & CAN - Pending ------------------------------------------------------------------------------------

            # ---------------------In case of Inter-State Supply, CGST & SGST/UTGST Tax Rate and/or Tax Amount cannot be applied----
            # ---------------------In case of Intra-State Supply, IGST Tax Rate and and/or Tax Amount cannot be applied---
            # ---------------In case of intra-State Supply, CGST & SGST cannot be zero----------------
            # -------------------In case of inter-State Supply, IGST cannot be zero
            df.loc[(df["SupplierGSTIN"].str[0:2] != df["POS"]) & ((df["CentralTaxRate"] != 0) & (df["StateUTTaxRate"] != 0)) & (
                df["POS"] != "NA") & (df["SupplierGSTIN"] != "NA"), 'Flash Remark'] = df["Flash Remark"]+"| Inter-state Supply_No CGST/SGST"
            df.loc[(df["SupplierGSTIN"].str[0:2] == df["POS"]) & (df["IntegratedTaxRate"] != 0) & (df["POS"] != "NA") & (
                df["SupplierGSTIN"] != "NA"), 'Flash Remark'] = df["Flash Remark"]+"| IGST_Rt cannot appear in Intra-state supply"

            df["Temp"] = df["SupplierGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)

            df.loc[(df["SupplierGSTIN"].str[0:2] == df["POS"]) & ((df["CentralTaxAmount"] == 0) & (df["StateUTTaxAmount"] == 0)) & (df["SupplyType"].isin(Non_zero_supply_type)) & (
                df["POS"] != "NA") & (df["SupplierGSTIN"] != "NA") & (df["Temp"] == "ValidGSTIN"), 'Flash Remark'] = df["Flash Remark"]+"| C&S cannot be 0"
            df.loc[(df["SupplierGSTIN"].str[0:2] != df["POS"]) & (df["IntegratedTaxAmount"] == 0) & (df["SupplyType"].isin(Non_zero_supply_type)) & (
                df["POS"] != "NA") & (df["SupplierGSTIN"] != "NA") & (df["Temp"] == "ValidGSTIN"), 'Flash Remark'] = df["Flash Remark"]+"| IGST cannot be 0"

            # ---------------------------------Document Number Length Exceeding 16 Digits------------------------
            df.loc[df["DocumentNumber"].str.len(
            ) > 16, "Flash Remark"] = df["Flash Remark"] + "| DocumentNumber exceeds 16 digits"
            # ---------Original Document Number Length Exceeding 16 Digits-------------
            df.loc[df["OriginalDocumentNumber"].str.len() > 16, "Flash Remark"] = df["Flash Remark"] + \
                "| OriginalDocumentNumber exceeds 16 digits"
            #progress['value'] = 90

            # ---------------------------------Invalid Tax Rates----------------------------
            # try:
            #    df["Temp"] = ((df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"]) / (df["TaxableValue"])* 100).round()
            # except:
            #    df["Temp"] = "NA"

            #df["Temp"] = df["Temp"].round()

            #df["Temp2"] = df["StateUTTaxRate"] + df["CentralTaxRate"] + df["IntegratedTaxRate"]
            #df["Temp2"] = df["Temp2"].round()

            #df.loc[((df["Temp"] != df["Temp2"]) & (df["TaxableValue"] != 0)) ,"EY Remark"] =  df["EY Remark"] + ""
            #df.loc[((df["Temp"] != df["Temp2"]) & (df["TaxableValue"] != 0)) ,"Flash Remark"] =  df["Flash Remark"] + "| Incorrect GST Tax Rates"

            #df.drop('Temp', axis=1, inplace=True)
            #df.drop('Temp2', axis=1, inplace=True)

            # ---------2.3.Tax amount is zero ------------------------------------------------------------------------------------
            try:

                df.loc[(df["InvoiceValue"] == 0) & (df["SupplyType"].isin(
                    Non_zero_supply_type)), 'Flash Remark'] = df["Flash Remark"] + "| InvoiceValue is zero"

                df.loc[(df["IntegratedTaxAmount"]+df["CentralTaxAmount"]+df["StateUTTaxAmount"] == 0) & (
                    df["SupplyType"].isin(Non_zero_supply_type)), 'Flash Remark'] = df["Flash Remark"] + "| Tax amount is zero"
            except:
                pass

            # ------------------Invalid Multiple Supply type Combination------------

            lst = list(outward_supp_doctype_combo["Document Type"].map(
                str)+outward_supp_doctype_combo["Supply Type"].map(str))
            lst1 = list(outward_supp_doctype_combo["Supply Type"].map(str))
            lst2 = list(outward_supp_doctype_combo["Document Type"].map(str))
            df["combo"] = df["DocumentType"].map(str)+df["SupplyType"].map(str)
            df.loc[(~df["combo"].isin(lst)) & (df["SupplyType"].isin(lst1)) & (df["DocumentType"].isin(
                lst2)), "Flash Remark"] = df["Flash Remark"]+"| Invalid Document and Supply type Combination"
            df.drop("combo", axis=1, inplace=True)

            # ---------------Same Document cannot have multiple Supply type---------

            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["ReturnPeriod"].map(str)
            DocumentSeries = df.groupby('Key_ID')['SupplyType'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["SupplyType"]
                                      > 1]["Key_ID"].tolist()
            lst1 = list(outward_supp_doctype_combo["Supply Type"].map(str))

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Multiple_Supply type"

            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ---------------Original Document & Revised Document cannot be reported in same tax period for the same Document Number---------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["ReturnPeriod"].map(str)
            DocumentSeries = df.groupby('Key_ID')['DocumentNumber'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["DocumentNumber"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i) & (df["Temp"] == "ValidDate"), 'Flash Remark'] + "| Single_doc-INV+RNV"

            df.drop('Temp', axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # --------------Original document for this cancelled record was not reported in this tax period--------
            lst = df["DocumentNumber"].unique()
            df.loc[(df["SupplyType"] == "CAN") & (df["OriginalDocumentNumber"] != "NA") & (~df["OriginalDocumentNumber"].isin(
                lst)), "Flash Remark"] = df["Flash Remark"] + "| Incorrect tax period for cancelled doc"

            # --------------Document Date cannot be prior to Original Document Date-----------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["Temp1"] = df["OriginalDocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            try:
                df.loc[(df["Temp"] == "ValidDate") & (df["Temp1"] == "ValidDate") & (df["DocumentDate"] <
                                                                                     df["OriginalDocumentDate"]), "Flash Remark"] = df["Flash Remark"]+"|  Doc_Dt cannot be prior to Org_Doc_Dt"
            except:
                pass
            try:
                df.loc[(df["Temp"] == "ValidDate") & (df["Temp1"] == "ValidDate") & (df["DocumentDate"] > str(datetime.now()).split()[
                    0]), "Flash Remark"] = df["Flash Remark"]+" | Document Date cannot be a future date beyond the current Date"

            except:
                pass
            df.drop('Temp', axis=1, inplace=True)
            df.drop('Temp1', axis=1, inplace=True)

            # ------------------Single CR / RCR Document cannot have both positive and negative amounts-------------

            df["Temp"] = df["TaxableValue"].apply(lambda x: -1 if x < 0 else 1)
            LineSeries = df.groupby('DocumentNumber')['Temp'].sum()
            newdf = LineSeries.to_frame()
            newdf.reset_index(inplace=True)

            DocumentSeries = df.groupby('DocumentNumber')[
                'DocumentNumber'].count()
            newdf1 = DocumentSeries.to_frame()
            newdf1.columns = ["DocumentNumber1"]

            newdf1.reset_index(inplace=True)
            lst = ["CR", "RCR"]

            for i in range(len(newdf)):
                if abs(newdf["Temp"][i]) != newdf1["DocumentNumber1"][i]:

                    df.loc[(df["DocumentNumber"] == newdf["DocumentNumber"][i]) & (df["DocumentType"].isin(
                        lst)), 'Flash Remark'] = df['Flash Remark'] + "| Single CR_with +ve and -ve values"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------------------------------------------Only Outward--------------------------------------------------
            # ----------------Invalid BillToState------------------------

            # ----------------Invalid ShippingBillNumber------------------------
            df["ShippingBillNumber"].fillna("NA", inplace=True)
            try:
                df["ShippingBillNumber"] = df["ShippingBillNumber"].apply(
                    lambda x: str(int(x)) if x != "NA" else x)
            except:
                try:
                    df["ShippingBillNumber"] = df["ShippingBillNumber"].apply(
                        lambda x: str(x) if x != "NA" else x)
                except:
                    pass

            df.loc[(df["ShippingBillNumber"].str.len() != 7) & (df["ShippingBillNumber"]
                                                                != "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_SB_No"

            # ----------------For supply type DTA, HSN is mandatory and HSN should start from 99-----------
            df.loc[(df["SupplyType"] == "DTA") & ((df["HSNorSAC"].str[0:2] != "99") | (
                df["HSNorSAC"] == "NA")), "Flash Remark"] = df["Flash Remark"] + "| SEZtoDTA_Chapter 99 HSN required"

            # ----------------In case Supply type is SEZ/EXPT/DXP/DTA the CGST and SGST can't be charged-----------
            df.loc[((df["SupplyType"] == "DTA") | (df["SupplyType"] == "SEZ") | (df["SupplyType"] == "EXPT")) & ((df["CentralTaxAmount"] != 0) | (
                df["StateUTTaxAmount"] != 0)), "Flash Remark"] = df["Flash Remark"] + "| Invalid GSTAmt_Supply type SEZ/EXPT/DTA"

            # ---------------Invalid TCSFlag------------------
            df.loc[((df["TCSFlag"] != "N") & (df["TCSFlag"] != "Y")) & (
                df["TCSFlag"] != "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_TCS_Flag"

            # ---------------In case,Supply Type is EXPWT the Tax Amount needs to be zero------------------
            try:
                df.loc[((df["SupplyType"] == "EXPWT") & (df["CentralTaxAmount"]+df["IntegratedTaxAmount"] +
                        df["StateUTTaxAmount"] != 0)), "Flash Remark"] = df["Flash Remark"] + "| EXPWT_Nil IGST"

                # ---------------Tax Rate and Tax Amount should be BLANK for invoices on which reverse charge is applicable------------------
                df.loc[(df["ReverseChargeFlag"] == "Y") & ((df["CentralTaxAmount"]+df["IntegratedTaxAmount"]+df["StateUTTaxAmount"] != 0) | (df["CentralTaxRate"] +
                                                                                                                                             df["StateUTTaxRate"]+df["IntegratedTaxRate"] != 0)), "Flash Remark"] = df["Flash Remark"] + "| In RCM, Tax_Rt and Tax_Amt must be 0"
            except:
                pass

            if option_list[0] == 0:
                df["DocumentNumber"] = df["Unaltered_Document_Number"]

            # -----------------Invalid ExportDuty--------------------------
            df["Temp"] = df["ExportDuty"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_ExportDuty"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid FOB--------------------------
            df["Temp"] = df["FOB"].apply(checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_FOB"
            df.drop('Temp', axis=1, inplace=True)

            # --------------Deriving PAN----------------
            PAN = " "

            lst = df["SupplierGSTIN"].unique()
            for x in lst:
                if len(str(x)) == 15:
                    PAN = x[3:13]
                    break

            # -----------------------------------------------------Only Outward--------------------------------------------------

            lst1 = df["Flash Remark"].unique()
            lst_all = df["Flash Remark"]
            df_lst = pd.DataFrame()
            lst3 = set()
            try:
                for lst in lst1:
                    lst2 = (lst.split("|", maxsplit=50))
                    for lst in lst2:
                        lst3.add(lst)
            except:
                lst3.add("")

            lst4 = []
            for x in lst3:
                count = 0
                for item in lst_all:
                    if item.find(x) != -1:
                        count = count+1
                lst4.append(count)

            df_lst["Unique_Errors_list"] = pd.Series(list(lst3))
            df_lst["Count_of_errors"] = pd.Series(list(lst4))
            # df_lst["Unique_Errors_list"].dropna(inplace=True)
            i = df_lst[df_lst["Unique_Errors_list"] == "NA"].index
            df_lst.drop(i, inplace=True)

            # ---------2.17(ii).Remaining transactions-------------
            df["EY Remark"].fillna("NA", inplace=True)
            df.loc[df["EY Remark"] == "NA", "EY Remark"] = "Other Purchases"

            df.replace("NA", "", inplace=True)
            df.replace("NN", "NA", inplace=True)

            df.drop("EY Remark", inplace=True, axis=1)
            df["Flash Remark"] = df["Flash Remark"].apply(
                lambda x: str(x).replace("NA", ""))

            # print("Saving To Database...")
            # ---Extracting Filename from Full Input File Path------

            #df.to_excel(writer, sheet_name='Sheet1',index=None)
            #df_lst.to_excel(writer, sheet_name='List_of_errors',index=None)

            df.to_sql("SalesRegisterFlash", panwisedb,
                      if_exists="append", index=False)
            df.to_sql("SalesRegisterFlashError", panwisedb,
                      if_exists="append", index=False)

    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'

    c = panwisedb.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS Summary_Totals (Particulars TEXT Primary Key,Count INTEGER, Taxable_Value INTEGER,CGST INTEGER,SGST INTEGER,IGST INTEGER,Cess INTEGER)")
    query = "select * FROM Summary_Totals"

    current = pd.read_sql_query(query, panwisedb)
    current.drop(current[current['Particulars'] ==
                 "SR - Flash"].index, inplace=True)
    query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM SalesRegisterFlash"
    upd = pd.read_sql_query(query, panwisedb)

    upd['Particulars'] = "SR - Flash"
    upd = upd.rename(columns={'SUM(TaxableValue)': 'Taxable_Value', 'SUM(CentralTaxAmount)': 'CGST', 'SUM(StateUTTaxAmount)': 'SGST',
                     'SUM(IntegratedTaxAmount)': 'IGST', 'COUNT(SupplierGSTIN)': 'Count', 'COUNT(CessAmountAdvalorem)': 'Cess'})
    neworder = ['Particulars', 'Count',
                'Taxable_Value', 'CGST', 'SGST', 'IGST', 'Cess']
    upd = upd.reindex(columns=neworder)
    upd = upd.fillna(0)
    upd = upd.append(current)
    upd.to_sql("Summary_Totals", panwisedb, if_exists="replace", index=False)
    query = "select * FROM Summary_Totals"

    current = pd.read_sql_query(query, panwisedb)
    current.drop(current[current['Particulars'] ==
                 "SR - Flash Error"].index, inplace=True)
    query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM SalesRegisterFlashError"
    upd = pd.read_sql_query(query, panwisedb)

    upd['Particulars'] = "SR - Flash Error"
    upd = upd.rename(columns={'SUM(TaxableValue)': 'Taxable_Value', 'SUM(CentralTaxAmount)': 'CGST', 'SUM(StateUTTaxAmount)': 'SGST',
                     'SUM(IntegratedTaxAmount)': 'IGST', 'COUNT(SupplierGSTIN)': 'Count', 'COUNT(CessAmountAdvalorem)': 'Cess'})
    neworder = ['Particulars', 'Count',
                'Taxable_Value', 'CGST', 'SGST', 'IGST', 'Cess']
    upd = upd.reindex(columns=neworder)
    upd = upd.fillna(0)
    upd = upd.append(current)

    upd.to_sql("Summary_Totals", panwisedb, if_exists="replace", index=False)

 # outward flash/sr flash

# def seqRun1(option_list, clientPAN, current_user):
def seqRun1(is_local, parameters):

    starttime = time.time()
    option_list = ["False", "False", "False", "False", "False", "False", "False"]

    # ---------------------Reading Input Inward SAP Dump File------------------------------
    dir_path = os.path.dirname(os.path.realpath(__file__))
    MASTER_path = dir_path+"\\"+"Flash"+"\\"+"Master"+"\\"
    Annexure_path = dir_path+"\\"+"Flash"+"\\"+"Annexure"+"\\"

    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]

    path = dir_path + '/Client-Details' if is_local else dir_path + '/Azure Blob GL Recon'
    path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'


    panwisedb = sqlite3.connect(path, timeout=10)
    cur = panwisedb.cursor()

    try:
        sql = "DROP TABLE PurchaseRegisterFlash"
        cur.execute(sql)
        panwisedb.commit()
    except:
        pass
    query = "select * FROM PurchaseRegisterDigi"
    hsn_goods = pd.read_excel(
        r""+MASTER_path+"Master file - HSN Code for goods.xlsx")
    hsn_services = pd.read_excel(
        r""+MASTER_path+"Master file - HSN codes for services.xlsx")
    uqc = pd.read_excel(r""+MASTER_path+"UOM Master.xlsx")
    port_code = pd.read_excel(
        r""+MASTER_path+"Service codes_state code and port code master.xlsx", sheet_name="Port Code")
    inward_doc_type = pd.read_excel(
        r""+MASTER_path+"Document type Supply type master (Inward).xlsx", sheet_name="Document type master")
    inward_supp_type = pd.read_excel(
        r""+MASTER_path+"Document type Supply type master (Inward).xlsx", sheet_name="Supply type master")
    inward_supp_doctype_combo = pd.read_excel(
        r""+MASTER_path+"Supplytype-DOCTYPE-Combination.xlsx", sheet_name="Inward")

    annexure = pd.read_excel(r""+Annexure_path+"DigiGST_inward_format.xlsx")
    for df in pd.read_sql_query(query, panwisedb, chunksize=10000):
        try:
            processed = df[~df['Unaltered_Document_Number'].isna()] 
            df.drop(df[~df['Unaltered_Document_Number'].isna()].index, inplace=True)
        except:
            traceback.print_exc()

    # ------------------Checking Format and changing accordingly-----------------
        ListOfFilecolumns = df.columns
        flag = checkInputFileColumns1(ListOfFilecolumns)
        if(flag == 1):
            #messagebox.showwarning("warning","Error: File not in DigiGST Format",icon="warning")
            # print("Process Terminated...Please upload again")
            return False
        else:
            # ------------------------Adding required columns -----------------
            # df.columns=annexure.columns
            ## print("Applying Validations...")
            df["EY Remark"] = "NA"
            df["Unaltered_Document_Number"] = df["DocumentNumber"]
            df.replace("NA", "NN", inplace=True)
            df["Flash Remark"] = "NA"

            def Return_period_change(df, key):
                if key:
                    # ----------------Changing Return Period-----------------
                    def checkReturnPeriodFormat(x):
                        regex = '(0[1-9]|10|11|12)20[0-9]{2}$'
                        try:
                            if(re.search(regex, str(x))):
                                return "Valid"
                            else:
                                return "Invalid"
                        except:
                            return "Invalid"

                    try:
                        df["ReturnPeriod"] = df["ReturnPeriod"].apply(
                            lambda x: "0"+str(x) if len(x) == 5 else str(x))
                    except:
                        traceback.print_exc()
                    df["Temp"] = df["ReturnPeriod"].apply(
                        checkReturnPeriodFormat)
                    #df.loc[df["Temp"]=="Invalid","Flash Remark"] = df["Flash Remark"] + "| Invalid_Ret_Pd"

                    try:
                        df.loc[(df["ReturnPeriod"].str.len() != 6) & (
                            df["Temp"] == "Invalid"), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Ret_Pd"
                    except:
                        traceback.print_exc()
                    df.drop('Temp', axis=1, inplace=True)
                else:
                    # -----------------Invalid ReturnPeriod--------------------------
                    def checkReturnPeriodFormat(x):
                        regex = '(0[1-9]|10|11|12)20[0-9]{2}$'
                        try:
                            if(re.search(regex, str(x))):
                                return "Valid"
                            else:
                                return "Invalid"
                        except:
                            return "Invalid"

                    df["Temp"] = df["ReturnPeriod"].apply(
                        checkReturnPeriodFormat)
                    try:
                        df.loc[(df["ReturnPeriod"].str.len() != 6) & (
                            df["Temp"] == "Invalid"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Ret_Pd"
                    except:
                        traceback.print_exc()
                    df.drop('Temp', axis=1, inplace=True)

                return df

            def valid_date(datestring):

                if str(datestring).find("-") != -1:

                    try:

                        if len(str(datestring).split("-")[0]) == 4:

                            try:
                                int(str(datestring).split("-")[0])
                                if int(str(datestring).split("-")[1]) <= 12:
                                    return "ValidDate"
                                else:
                                    return "InvalidDate"

                            except:
                                return "InvalidDate"
                        else:

                            return "InvalidDate"
                    except:
                        return "InvalidDate"
                if str(datestring).find("//") != -1:

                    try:

                        if len(str(datestring).split("//")[0]) == 4:

                            try:
                                int(str(datestring).split("//")[0])
                                if int(str(datestring).split("//")[1]) <= 12:
                                    return "ValidDate"
                                else:
                                    return "InvalidDate"
                            except:
                                return "InvalidDate"
                        else:

                            return "InvalidDate"
                    except:
                        return "InvalidDate"

                if str(datestring).find(".") != -1:

                    try:

                        if len(str(datestring).split(".")[0]) == 4:

                            try:
                                int(str(datestring).split(".")[0])
                                if int(str(datestring).split(".")[1]) <= 12:
                                    return "ValidDate"
                                else:
                                    return "InvalidDate"
                            except:
                                return "InvalidDate"
                        else:

                            return "InvalidDate"
                    except:
                        return "InvalidDate"

            def Invalid_Document_Number(df, key):
                lst = ["!", "@", "#", ".", "$", "^",
                       "&", "*", "(", ")", "\\", " "]
                lst1 = ["CR", "DR", "RNV"]
                try:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x))
                except:
                    try:
                        df["DocumentNumber"] = df["DocumentNumber"].apply(
                            lambda x: int(x))
                    except:
                        traceback.print_exc()

                if key:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace(".", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("&", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace(" ", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("^", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("@", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("!", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("#", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("$", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("%", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("*", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("\\", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace(")", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("(", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("?", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("_", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("<", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace(">", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("{", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("}", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("[", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("]", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("~", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace(":", ""))
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).replace("=", ""))
                    try:
                        df.loc[(df["DocumentNumber"].str.len() > 16) | (
                            df["DocumentNumber"] == "NA"), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Doc_No"
                    except:
                        traceback.print_exc()
                else:

                    def checkSpecialCharacter(x):
                        regex = re.compile('[@_!#$%^&*()<>?\|}{~:]= .')
                        try:
                            if(regex.search(str(x))):

                                return "Invalid"

                            if (str(x).find("?") != -1):
                                return "Invalid"
                            else:
                                return "Valid"

                        except:
                            return "Invalid"

                    df["Temp"] = df["DocumentNumber"].apply(
                        checkSpecialCharacter)
                    try:
                        df.loc[(df["Temp"] == "Invalid") | (df["DocumentNumber"].str.len() > 16) | (
                            df["DocumentNumber"] == "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Doc_No"
                    except:
                        traceback.print_exc()
                    df.drop('Temp', axis=1, inplace=True)

                return df

            def Invalid_Original_document_Number(df, key):
                lst = ["!", "@", "#", ".", "$", "^",
                       "&", "*", "(", ")", "\\", " "]
                lst1 = ["CR", "DR", "RNV"]
                try:
                    df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                        lambda x: str(x))
                except:
                    try:
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: int(x))
                    except:
                        traceback.print_exc()

                if key:
                    try:
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(".", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("&", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(" ", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("!", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("@", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("#", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("$", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("%", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("^", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("*", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("(", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(")", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("\\", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("?", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("_", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("<", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(">", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("{", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("}", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("[", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("]", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("~", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace(":", ""))
                        df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                            lambda x: str(x).replace("=", ""))

                    except:
                        traceback.print_exc()

                    try:
                        df.loc[((df["OriginalDocumentNumber"].str.len() > 16) | (df["OriginalDocumentNumber"] == "NA")) & (
                            df["DocumentType"].isin(lst1)), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    except:
                        traceback.print_exc()

                else:
                    lst1 = ["CR", "DR", "RNV"]

                    def checkSpecialCharacter(x):
                        regex = re.compile('[@_!#$%^&*()<>?\|}{~:]= .')
                        if(regex.search(str(x))):

                            return "Invalid"
                        else:
                            return "Valid"

                    df["Temp"] = df["OriginalDocumentNumber"].apply(
                        checkSpecialCharacter)
                    #df.loc[(df["Temp"]=="Invalid") & (df["DocumentType"].isin(lst1))  ,"Flash Remark"] = df["Flash Remark"] + "| Invalid_Org_Doc_No"
                    #df.loc[(df["OriginalDocumentNumber"]=="NA") & (df["DocumentType"].isin(lst1)),"Flash Remark"]=df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    try:
                        df.loc[((df["OriginalDocumentNumber"].str.len() > 16) | (df["Temp"] == "Invalid")) & (
                            df["DocumentType"].isin(lst1)), "Flash Remark"] = df["Flash Remark"]+"| Invalid_Org_Doc_No"
                    except:
                        traceback.print_exc()
                    df.drop('Temp', axis=1, inplace=True)
                return df

            # _______________Line Number Check and change______________________

            def Line_number_change(df, key):
                if key:

                    df["LineNumber"] = df.index+1

                else:
                    try:
                        df["LineNumber"] = df["LineNumber"].apply(
                            lambda x: int(x) if x != "NA" else x)
                    except:
                        traceback.print_exc()

                    df.loc[(df["LineNumber"] == "NA"),
                           "Flash Remark"] = df["Flash Remark"]+"| Invalid_Line_No"

                    try:

                        LineSeries = df.groupby('DocumentNumber')[
                            'LineNumber'].nunique()
                        newdf = LineSeries.to_frame()
                        newdf.reset_index(inplace=True)

                        DocumentSeries = df.groupby('DocumentNumber')[
                            'DocumentNumber'].count()
                        newdf1 = DocumentSeries.to_frame()
                        newdf1.columns = ["DocumentNumber1"]

                        newdf1.reset_index(inplace=True)

                        for i in range(len(newdf)):
                            if newdf["LineNumber"][i] != newdf1["DocumentNumber1"][i]:
                                df.loc[(df["DocumentNumber"] == newdf["DocumentNumber"][i]),
                                       'Flash Remark'] = df['Flash Remark'] + "| Invalid_Line_No"
                    except:
                        traceback.print_exc()

                    df.reset_index(drop=True, inplace=True)

                return df

            # ------------------------Invalid POS check and chnage----------------

            def POS_change(df, key):
                # ----------------Invalid POS------------------------------
                if key:
                    df["POS"] = df["POS"].apply(
                        lambda x: "0"+str(x) if len(str(x)) == 1 else str(x))
                    listof_POS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["POS"].isin(
                        listof_POS), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"
                    lst = ["EXPWT", "EXPT"]
                    df.loc[(df["POS"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"

                else:
                    listof_POS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "97"]
                    df.loc[~df["POS"].isin(
                        listof_POS), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"
                    lst = ["EXPWT", "EXPT"]
                    df.loc[(df["POS"] == "97") & (~df["SupplyType"].isin(
                        lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS"

                return df

            def Date_check_change(df, key):
                # -----------------------------------Invalid Document Date------------------------------

                if key:

                    def all_format_change(x):

                        if (str(x).split()[0].find(".") != -1):

                            if (len(str(x).split()[0].split(".")[0]) == 2) & (len(str(x).split()[0].split(".")[2]) == 4):

                                return (str(x).split()[0].split(".")[2]+"-"+str(x).split()[0].split(".")[1]+"-"+str(x).split()[0].split(".")[0])
                            else:
                                return (str(x).split()[0].split(".")[0]+"-"+str(x).split()[0].split(".")[1]+"-"+str(x).split()[0].split(".")[2])

                        if (str(x).split()[0].find("-") != -1):

                            if (len(str(x).split()[0].split("-")[0]) == 2) & (len(str(x).split()[0].split("-")[1]) == 2):

                                return (str(x).split()[0].split("-")[2]+"-"+str(x).split()[0].split("-")[1]+"-"+str(x).split()[0].split("-")[0])
                            else:
                                return (str(x).split()[0].split("-")[0]+"-"+str(x).split()[0].split("-")[1]+"-"+str(x).split()[0].split("-")[2])

                        else:
                            return x

                    def Chnage_Date_format(x):
                        try:

                            date = datetime.strptime(x, '%Y-%m-%d')

                            return(str(date).split()[0])
                        except:

                            return all_format_change(x)

                    df["DocumentDate"] = df["DocumentDate"].apply(
                        lambda x: Chnage_Date_format(x))
                    #df["OriginalDocumentDate"]=df["OriginalDocumentDate"].apply(lambda x:Chnage_Date_format(x))
                    #df["PurchaseVoucherDate"]=df["PurchaseVoucherDate"].apply(lambda x:Chnage_Date_format(x))
                    #df["PaymentDate"]=df["PaymentDate"].apply(lambda x:Chnage_Date_format(x))
                    #df["BillOfEntryDate"]=df["BillOfEntryDate"].apply(lambda x:Chnage_Date_format(x))

                else:

                    df["Temp"] = df["DocumentDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[(df["Temp"] == "InvalidDate") | (df["Temp"] == "NA"),
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Doc_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["DocumentDate"] = df["DocumentDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)
                    # # ---------------------------------End Invalid Document Dates---------------------------

                    # -----------------------------------Invalid OriginalDocumentDate------------------------------
                    df["OriginalDocumentDate"] = np.where(df["OriginalDocumentDate"] == 'NaT', "NA", df["OriginalDocumentDate"])
                    df["Temp"] = df["OriginalDocumentDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Org_Doc_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)
                    # # ---------------------------------End Invalid OriginalDocument Dates---------------------------

                    # -----------------------------------Invalid PurchaseVoucherDate------------------------------
                    df["PurchaseVoucherDate"] = np.where(df["PurchaseVoucherDate"] == 'NaT', "NA", df["PurchaseVoucherDate"])
                    df["Temp"] = df["PurchaseVoucherDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Pur_Vou_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["PurchaseVoucherDate"] = df["PurchaseVoucherDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)

                    # # ---------------------------------End Invalid PurchaseVoucherDate Dates---------------------------

                    # -----------------------------------Invalid BillOfEntryDate------------------------------
                    df["BillOfEntryDate"] = np.where(df["BillOfEntryDate"] == 'NaT', "NA", df["BillOfEntryDate"])
                    df["Temp"] = df["BillOfEntryDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_BOE_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)

                    df["BillOfEntryDate"] = df["BillOfEntryDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)

                    # # ---------------------------------End Invalid BillOfEntryDate Dates---------------------------
                    # -----------------------------------Invalid PaymentDate------------------------------
                    df["PaymentDate"] = np.where(df["PaymentDate"] == 'NaT', "NA", df["PaymentDate"])
                    df["Temp"] = df["PaymentDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[df["Temp"] == "InvalidDate",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Pt_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                    # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["PaymentDate"] = df["PaymentDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)
                # # ---------------------------------End Invalid PaymentDate Dates---------------------------

                # # ---------------------------------End Invalid ContractDate ---------------------------
                    df["ContractDate"] = np.where(df["ContractDate"] == 'NaT', "NA", df["ContractDate"])
                    df["Temp"] = df["ContractDate"].apply(
                        lambda x: valid_date(x) if x != "NA" else x)
                    df.loc[(df["Temp"] == "InvalidDate"),
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cont_Dt"
                    df2 = df[df["Temp"] == "InvalidDate"]
                    df1 = df[(df["Temp"] == "ValidDate")
                             | (df["Temp"] == "NA")]

                # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------
                    frames = [df1, df2]
                    df = pd.concat(frames)
                    df.drop('Temp', axis=1, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    df["ContractDate"] = df["ContractDate"].apply(
                        lambda x: str(x).split()[0] if x != "NA" else x)

                return df

            # -------------------------itc VALIDATIONS---------------------------

            def checkPositiveandNegativeValuesFormat(x):
                regex = '^-?(0|[1-9]\d*)(\.\d+)?$'
                if(re.search(regex, str(x)) or x == "NA"):

                    return "Valid"
                else:

                    return "Invalid"

            def ITC_validations(df, key):
                if key:

                    df["Temp"] = df["AvailableIGST"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Available_IGST"
                    df.drop('Temp', axis=1, inplace=True)
                    df["Temp"] = df["AvailableCGST"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Available_CGST"
                    df.drop('Temp', axis=1, inplace=True)
                    df["Temp"] = df["AvailableSGST"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Available_SGST"
                    df.drop('Temp', axis=1, inplace=True)
                    df["Temp"] = df["AvailableCess"].apply(
                        checkPositiveandNegativeValuesFormat)
                    df.loc[df["Temp"] == "Invalid",
                           "Flash Remark"] = df["Flash Remark"] + "| Invalid_Available_Cess"
                    df.drop('Temp', axis=1, inplace=True)
                    try:

                        df["Temp"] = df["AvailableSGST"].map(
                            int)+df["AvailableIGST"].map(int)+df["AvailableCGST"].map(int)
                        # df["Temp1"] = df["IntegratedTaxAmount"].map(int)+df["CentralTaxAmount"].map(
                        #     int)+df["StateUTTaxAmount"].map(int)+df["CessAmountSpecific"].map(int)+df["CessAmountAdvalorem"].map(str)
                        df["Temp1"] = df["IntegratedTaxAmount"].map(int)+df["CentralTaxAmount"].map(
                            int)+df["StateUTTaxAmount"].map(int)+df["CessAmountSpecific"].map(int)+df["CessAmountAdvalorem"].map(int)

                        #df.loc[(df["Temp"]>0) & (df["EligibilityIndicator"]=="NO"),"Flash Remark"]=df["Flash Remark"]+"| Invalid Available ITC_Eligibility is NO"

                        # df["Temp"]=df["AvailableIGST"]+df["AvailableCGST"]+df["AvailableSGST"]
                        df.loc[df["Temp"] > df["Temp1"],
                               "Flash Remark"] = df["Flash Remark"]+"| TotAvailTax > TotalTax"
                        df.loc[(df["Temp"] > 0) & (df["EligibilityIndicator"] == "NO"),
                               "Flash Remark"] = df["Flash Remark"]+"| Avail_ITC_nomatch_EligiIndicator"
                        df.drop("Temp", inplace=True, axis=1)
                        df.drop("Temp1", inplace=True, axis=1)

                    except:
                        try:
                            df.drop("Temp", inplace=True, axis=1)
                            df.drop("Temp1", inplace=True, axis=1)

                        except:
                            traceback.print_exc()

                # --------------Receipient should not claim ITC when POS and his State code are diffferent-----------
                    try:
                        df["Temp"] = df["AvailableSGST"] + \
                            df["AvailableIGST"]+df["AvailableCGST"]

                        df.loc[(df["POS"] != "NA") & (df["POS"] != df["RecipientGSTIN"].apply(lambda x:str(x)[0:2])) & (df["RecipientGSTIN"] != "NA") & (
                            df["Temp"] != 0), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS and state code_NotAvail ITC"
                        df.drop("Temp", axis=1, inplace=True)
                    except:
                        try:
                            df.drop("Temp", axis=1, inplace=True)
                        except:
                            traceback.print_exc()

                else:
                    try:

                        df["Temp"] = df["StateUTTaxAmount"].map(
                            int)+df["IntegratedTaxAmount"].map(int)+df["CentralTaxAmount"].map(int)

                        #df.loc[(df["Temp"]>0) & (df["EligibilityIndicator"]=="NO"),"Flash Remark"]=df["Flash Remark"]+"| Invalid Available ITC_Eligibility is NO"

                        # df["Temp"]=df["AvailableIGST"]+df["AvailableCGST"]+df["AvailableSGST"]

                        df.loc[(df["Temp"] > 0) & (df["EligibilityIndicator"] == "NO"),
                               "Flash Remark"] = df["Flash Remark"]+"| ITC_nomatch_EligiIndicator"
                        df.drop("Temp", inplace=True, axis=1)

                    except:
                        try:
                            df.drop("Temp", inplace=True, axis=1)

                        except:
                            traceback.print_exc()

                # --------------Receipient should not claim ITC when POS and his State code are diffferent-----------
                    try:
                        df["Temp"] = df["StateUTTaxAmount"] + \
                            df["IntegratedTaxAmount"]+df["CentralTaxAmount"]

                        df.loc[(df["POS"] != "NA") & (df["POS"] != df["RecipientGSTIN"].apply(lambda x:str(x)[0:2])) & (df["RecipientGSTIN"] != "NA") & (
                            df["Temp"] != 0), "Flash Remark"] = df["Flash Remark"] + "| Invalid_POS and state code_NotAvail ITC"
                        df.drop("Temp", axis=1, inplace=True)
                    except:
                        try:
                            df.drop("Temp", axis=1, inplace=True)
                        except:
                            traceback.print_exc()

                return df

            # ---------------------------Invoice value Validations----------------------------

            def Invoice_value_validaions(df, *key):

                if key[0][2]:
                    try:

                        df["InvoiceValue"] = df["TaxableValue"] + df["IntegratedTaxAmount"]+df["CentralTaxAmount"] + \
                            df["StateUTTaxAmount"]+df["CessAmountAdvalorem"] + \
                            df["CessAmountSpecific"]
                    except:
                        traceback.print_exc()

                if key[0][1]:
                    # --------------Line Number Check--------------------
                    try:
                        ## print(df["Flash Remark"])
                        df.loc[(df["InvoiceValue"] != (df["TaxableValue"] + df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"] +
                                df["CessAmountAdvalorem"]+df["CessAmountSpecific"])), "Flash Remark"] = str(df["Flash Remark"].map(str))+"| Invalid_Inv_Val"
                    except:
                        traceback.print_exc()
                if key[0][0]:
                    # ------------Invoice Level Check------------------
                    try:
                        df["Temp"] = df["TaxableValue"]+df["IntegratedTaxAmount"]+df["CentralTaxAmount"] + \
                            df["StateUTTaxAmount"]+df["CessAmountAdvalorem"] + \
                            df["CessAmountSpecific"]
                        LineSeries = df.groupby('DocumentNumber')[
                            'InvoiceValue', 'Temp'].sum()
                        newdf = LineSeries
                        newdf.reset_index(inplace=True)
                        for i in range(len(newdf)):

                            if newdf["InvoiceValue"][i] != newdf["Temp"][i]:

                                df.loc[(df["DocumentNumber"] == newdf["DocumentNumber"][i]),
                                       'Flash Remark'] = df['Flash Remark'] + "| Invalid_Inv_Val"

                        df.drop('Temp', axis=1, inplace=True)
                    except:
                        traceback.print_exc()
                return df
            # ----------------------Replacing all values where blank is there with 0 -------------------------
            df["Quantity"].fillna(0, inplace=True)

            df["IntegratedTaxAmount"].fillna(0, inplace=True)
            df["StateUTTaxAmount"].fillna(0, inplace=True)
            df["CentralTaxAmount"].fillna(0, inplace=True)

            df["StateUTTaxRate"].fillna(0, inplace=True)
            df["IntegratedTaxRate"].fillna(0, inplace=True)
            df["CentralTaxRate"].fillna(0, inplace=True)
            df["CessRateSpecific"].fillna(0, inplace=True)
            df["CessRateAdvalorem"].fillna(0, inplace=True)

            df["AvailableCGST"].fillna(0, inplace=True)
            df["AvailableIGST"].fillna(0, inplace=True)
            df["AvailableSGST"].fillna(0, inplace=True)
            df["AvailableCess"].fillna(0, inplace=True)
            df["CessAmountAdvalorem"].fillna(0, inplace=True)
            df["CessAmountSpecific"].fillna(0, inplace=True)

            df["InvoiceValue"].fillna(0, inplace=True)
            df["TaxableValue"].fillna(0, inplace=True)

            def TAX_valuation(x):
                try:
                    a = float(x)
                    return "Valid"
                except:
                    return "Invalid"

            # df["tax_check"]="NA"
            '''
            df["tax_check"]=df["IntegratedTaxAmount"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["tax_check"]=df["CentralTaxAmount"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["tax_check"]=df["StateUTTaxAmount"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["tax_check"]=df["InvoiceValue"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["tax_check"]=df["TaxableValue"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["tax_check"]=df["CessAmountSpecific"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["tax_check"]=df["CessAmountAdvalorem"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)

            df["rate_check"]="NA"
            df["rate_check"]=df["IntegratedTaxRate"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["rate_check"]=df["CentralTaxRate"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["rate_check"]=df["StateUTTaxRate"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["rate_check"]=df["CessRateSpecific"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)
            df["rate_check"]=df["CessRateAdvalorem"].apply(lambda x:TAX_valuation(x) if x=="NA" else x)

            df["Return_check"]="NA"


            df["Return_check"]=df["IntegratedTaxRate"].apply(lambda x:checkReturnPeriodFormat(x) if x=="NA" else x)
            '''

            df.fillna("NA", inplace=True)
            try:
                df["DocumentType"] = df["DocumentType"].apply(
                    lambda x: str(x).upper())
                df["SupplyType"] = df["SupplyType"].apply(
                    lambda x: str(x).upper())
            except:
                traceback.print_exc()

            # ---------------Calling Option Functions----------------------
            Non_zero_supply_type = ["TAX", "EXPT", "ISD",
                                    "ISDIE", "ISD8", "ISIE8", "SOA", "LGAS", "DXP"]

            df = Return_period_change(df, option_list[4])
            df = Invalid_Document_Number(df, option_list[0])

            df = Invalid_Original_document_Number(df, option_list[0])
            df = Line_number_change(df, option_list[2])
            df = POS_change(df, option_list[3])

            df = Date_check_change(df, option_list[1])
            df = ITC_validations(df, option_list[5])

            df = Invoice_value_validaions(df, option_list[6])

            df = Date_check_change(df, 0)
            # ----------------Changing Date columns to required format-----------------
            #df["DocumentDate"] = df["DocumentDate"].apply(lambda x:str(x).split()[0])
            #df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(lambda x:str(x).split()[0])
            #df["PurchaseVoucherDate"] = df["PurchaseVoucherDate"].apply(lambda x:str(x).split()[0])
            #df["BillOfEntryDate"] = df["BillOfEntryDate"].apply(lambda x:str(x).split()[0])
            #df["PaymentDate"] = df["PaymentDate"].apply(lambda x:str(x).split()[0])

            # ----------------Changing DocumentNumber,OriginaldocumentNumber to required Text format-----------------
            #df["DocumentNumber"] = df["DocumentNumber"].astype(str)
            #df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].astype(str)
            df["POS"] = df["POS"].astype(str)

            # df["EY Comments"] = np.nan

            # ------------------------Similar document number having multiple return period------------------------------------------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['ReturnPeriod'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ReturnPeriod"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate") & (
                    df["DocumentType"] != "SLF"), 'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Multiple_Ret_Pd"

            df["Key_ID1"] = df["DocumentNumber"].map(str)+df["RecipientGSTIN"].map(str)+df["SupplierGSTIN"].map(
                str)+df["DocumentType"].map(str)+df["FY"].map(str)+df["SupplierName"].map(str)

            DocumentSeries = df.groupby('Key_ID1')['ReturnPeriod'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ReturnPeriod"]
                                      > 1]["Key_ID1"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID1"] == i) & (df["Temp"] == "ValidDate") & (df["DocumentType"] == "SLF"),
                       'Flash Remark'] = df['Flash Remark'] + "|  Similar DocumentNumber cannot have multiple Return Period"
            df.drop('Key_ID1', axis=1, inplace=True)

            df.drop('Temp', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple document dates------------------------------------------------

            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['DocumentDate'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["DocumentDate"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                try:
                    df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"), 'Flash Remark'] = df.loc[(
                        df["Key_ID"] == i) & (df["Temp"] == "ValidDate"), 'Flash Remark'] + "| Single_doc-Multiple_Doc_Dt"
                except:
                    traceback.print_exc()
            df.drop('Temp', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple customer GSTIN------------------------------------------------
            #df["Temp"] = df["DocumentDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['SupplierGSTIN'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["SupplierGSTIN"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(df["Key_ID"] == i),
                                                                     'Flash Remark'] + "| Similar DocumentNumber cannot have multiple SupplierGSTINs"

            #df.drop('Temp', axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple RecipientGSTIN GSTIN------------------------------------------------
            #df["Temp"] = df["DocumentDate"].apply(lambda x : valid_date(x) if x!= "NA" else x)
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['RecipientGSTIN'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["RecipientGSTIN"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "| Single_doc-Multiple_Rec_GSTIN"

            #df.drop('Temp', axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple Original Customer,Supplier GSTIN------------------------------------------------
            try:
                DocumentSeries = df.groupby('DocumentNumber')[
                    'OriginalSupplierGSTIN'].nunique()
                newdf = DocumentSeries.to_frame()
                newdf.reset_index(inplace=True)

                ListOfDocumentNos = newdf[newdf["OriginalSupplierGSTIN"]
                                          > 1]["DocumentNumber"].tolist()

                for i in ListOfDocumentNos:
                    df.loc[df["DocumentNumber"] == i, 'Flash Remark'] = df.loc[df["DocumentNumber"] ==
                                                                               i, 'Flash Remark'] + "| Similar DocumentNumber cannot have multiple SupplierGSTINs"
            except:
                traceback.print_exc()
            # ------------------------Similar document number having multiple PORT Codes------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)

            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['PortCode'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["PortCode"] > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "| Single_doc-Multiple_Portcode"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple  BillOFEntry------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)

            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['BillOfEntry'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["BillOfEntry"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[df["Key_ID"] == i, 'Flash Remark'] = df.loc[df["Key_ID"]
                                                                   == i, 'Flash Remark'] + "| Single_doc-Multiple_BoE"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ------------------------Similar document number having multiple  BillOFEntryDate------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)

            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby('Key_ID')['BillOfEntryDate'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["BillOfEntryDate"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "|  Single_doc-Multiple_BoEDt"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple POS------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            df["Key_ID1"] = df["DocumentNumber"].map(str)+df["RecipientGSTIN"].map(str)+df["SupplierGSTIN"].map(
                str)+df["DocumentType"].map(str)+df["FY"].map(str)+df["SupplierName"].map(str)
            DocumentSeries = df.groupby('Key_ID')['POS'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["POS"] > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID1"] == i) & (df["DocumentType"] != "SLF"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Different_POS"

            df["Key_ID1"] = df["DocumentNumber"].map(str)+df["RecipientGSTIN"].map(str)+df["SupplierGSTIN"].map(
                str)+df["DocumentType"].map(str)+df["FY"].map(str)+df["SupplierName"].map(str)

            DocumentSeries = df.groupby('Key_ID1')['POS'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["POS"] > 1]["Key_ID1"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID1"] == i) & (df["DocumentType"] == "SLF"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Different_POS"
            df.drop('Key_ID1', axis=1, inplace=True)

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # ------------------------Similar document number having multiple ReverseChargeFlag------------------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            DocumentSeries = df.groupby(
                'Key_ID')['ReverseChargeFlag'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["ReverseChargeFlag"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i), 'Flash Remark'] + "| Single_doc-Different_RCM"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            # --------------------------Document Numbers (CR or DR) having multiple OriginalReferences--------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)

            ListOfDocNos = df[((df["DocumentType"] == "CR") | (
                df["DocumentType"] == "DR"))]["Key_ID"].unique()
            for i in ListOfDocNos:
                if(df[df["Key_ID"] == i].count()[0] > 1):
                    df2 = df[df["Key_ID"] == i]
                    if(len(df2["OriginalDocumentNumber"].unique()) > 1):
                        df.loc[df["Key_ID"] == i, 'Flash Remark'] = df.loc[df["Key_ID"]
                                                                           == i, 'Flash Remark'] + "| Single_doc-Multiple_OrgDoc"
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)

            # -----------------------------------
            # --------------------------Document Numbers (CR or DR) having multiple OriginalSupplierGSTIN--------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)

            ListOfDocNos = df[((df["DocumentType"] == "CR") | (
                df["DocumentType"] == "DR"))]["Key_ID"].unique()
            for i in ListOfDocNos:
                if(df[df["Key_ID"] == i].count()[0] > 1):
                    df2 = df[df["Key_ID"] == i]
                    if(len(df2["OriginalSupplierGSTIN"].unique()) > 1):
                        df.loc[df["Key_ID"] == i, 'Flash Remark'] = df['Flash Remark'] + \
                            "| Single_doc-Multiple_Sup_GSTIN"
            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            # --------------------------Document Numbers (CR or DR) having multiple OriginalReferences--------------------------------------
            df["FY"] = df["DocumentDate"].apply(
                lambda x: x[0:4] if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["DocumentType"].map(str)+df["FY"].map(str)
            ListOfDocNos = df[((df["DocumentType"] == "CR") | (
                df["DocumentType"] == "DR"))]["Key_ID"].unique()
            for i in ListOfDocNos:
                if(df[df["Key_ID"] == i].count()[0] > 1):
                    df2 = df[df["Key_ID"] == i]
                    if(len(df2["OriginalDocumentDate"].unique()) > 1):
                        df.loc[df["Key_ID"] == i, 'Flash Remark'] = df.loc[df["Key_ID"]
                                                                           == i, 'Flash Remark'] + "| Single_doc-Multiple_Org_date"

            df.drop('Key_ID', axis=1, inplace=True)
            df.drop("FY", axis=1, inplace=True)
            # -----------------Applying Validations-------------------------

            def checkPositiveandNegativeValuesFormat(x):
                regex = '^-?(0|[1-9]\d*)(\.\d+)?$'
                if(re.search(regex, str(x)) or x == "NA"):
                    return "Valid"
                else:
                    return "Invalid"

            def checkPositiveValuesFormat(x):
                regex = '^(0|[1-9]\d*)(\.\d+)?$'
                if(re.search(regex, str(x)) or x == "NA"):
                    return "Valid"
                else:
                    return "Invalid"
            # ----------------Invalid_Taxable_Val-------------------------
            df["Temp"] = df["TaxableValue"].apply(
                checkPositiveandNegativeValuesFormat)
            #df.loc[(df["TaxableValue"]<=0) & (df["DocumentType"]!="CR"),"Flash Remark"] = df["Flash Remark"] + "| Invalid_Taxable_Val"

            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Taxable_Val"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid_IGST_Amt-------------------
            df["Temp"] = df["IntegratedTaxAmount"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_IGST_Amt"
            df["Information Error"] = ""
            try:
                df["Temp"] = df["TaxableValue"]*df["IntegratedTaxRate"]/100
                df.loc[abs(df["Temp"]-df["IntegratedTaxAmount"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"|  Incorrect IGST Amount"
                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    traceback.print_exc()

            # -----------------Invalid Invalid_CGST_Amt-------------------
            df["Temp"] = df["CentralTaxAmount"].apply(
                checkPositiveandNegativeValuesFormat)

            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_CGST_Amt"
            try:
                df["Temp"] = df["TaxableValue"]*df["CentralTaxRate"]/100

                df.loc[abs(df["Temp"]-df["CentralTaxAmount"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"| Incorrect CGST Amount"

                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    traceback.print_exc()

            # -----------------Invalid StateUTTaxAmount-------------------
            df["Temp"] = df["StateUTTaxAmount"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_SGST/UTGST_Amt"
            try:
                df["Temp"] = df["TaxableValue"]*df["StateUTTaxRate"]/100

                df.loc[abs(df["Temp"]-df["StateUTTaxAmount"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"|  Incorrect SGST Amount"
                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    traceback.print_exc()

            # -----------------Invalid CessAmountAdvalorem-------------------

            df["Temp"] = df["CessAmountAdvalorem"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cess_Amt"
            try:
                df["Temp"] = df["TaxableValue"]*df["CessRateAdvalorem"]/100

                df.loc[abs(df["Temp"]-df["CessAmountAdvalorem"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"| Incorrect CessAmountAdvalorem"
                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    traceback.print_exc()
            # -----------------Invalid CessAmountSpecific--------------------
            df["Temp"] = df["CessAmountSpecific"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid", "Flash Remark"] = df["Flash Remark"] + \
                "| Invalid CessAmountSpecific"
            try:
                df["Temp"] = df["TaxableValue"]*df["CessRateSpecific"]/100

                df.loc[abs(df["Temp"]-df["CessAmountSpecific"]) > 0.01,
                       "Information Error"] = df["Information Error"]+"| Incorrect CessAmountSpecific"
                df.drop('Temp', axis=1, inplace=True)
            except:
                try:
                    df.drop('Temp', axis=1, inplace=True)
                except:
                    traceback.print_exc()
            # -----------------Invalid InvoiceValue--------------------------
            df["Temp"] = df["InvoiceValue"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Invoice_Val"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid ContractValue--------------------------

            df["Temp"] = df["ContractValue"].apply(
                checkPositiveandNegativeValuesFormat)
            #df.loc[(df["TaxableValue"]<=0) & (df["DocumentType"]!="CR"),"Flash Remark"] = df["Flash Remark"] + "| Invalid_Taxable_Val"

            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cont_Val"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid CIFValue--------------------------
            df["Temp"] = df["CIFValue"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_CIF_Val"
            lst = ["IMPG", "SEZG"]
            df.loc[(df["CIFValue"] == "NA") & (df["DocumentType"].isin(lst)),
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_CIF_Val"
            df.drop('Temp', axis=1, inplace=True)
            # -----------------Invalid CustomDuty--------------------------
            df["Temp"] = df["CustomDuty"].apply(
                checkPositiveandNegativeValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_CustomDuty"
            df.drop('Temp', axis=1, inplace=True)

            # -----------------Invalid Quantity------------------------------
            df["Temp"] = df["Quantity"].apply(checkPositiveValuesFormat)
            df.loc[df["Temp"] == "Invalid",
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid_Qty"
            df.drop('Temp', axis=1, inplace=True)

            # ---------For Invalid IGSTRate------------------------
            '''

            try:
                df["Temp"] =  (df["IntegratedTaxAmount"]) / ((df["TaxableValue"]))*100
            except:
                df["Temp"] = 0
            df["Temp"] = df["Temp"].apply(lambda x: round(x,2))
            df["IntegratedTaxRate"] = df["IntegratedTaxRate"].apply(lambda x: round(x,2))
            df.loc[df["IntegratedTaxRate"]-df["Temp"]>0.1,"Flash Remark"] = df["Flash Remark"] + "| Invalid_IGST_Rt"
            '''
            df.loc[~df["IntegratedTaxRate"].isin(
                [0, 0.1, 3, 5, 12, 18, 28, 0.0, 3.0, 5.0, 12.0, 18.0, 28.0]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_IGST_Rt"
            #df.drop('Temp', axis=1, inplace=True)

            # ---------For Invalid SGSTRate--------------------
            '''
            try:
                df["Temp"] =  (df["StateUTTaxAmount"]) / ((df["TaxableValue"]))*100
            except:
                df["Temp"] = 0
            df["Temp"] = df["Temp"].apply(lambda x: round(x,2))
            df["StateUTTaxRate"] = df["StateUTTaxRate"].apply(lambda x: round(x,2))
            df.loc[(df["StateUTTaxRate"]-df["Temp"]>0.1),"Flash Remark"] = df["Flash Remark"] + "| Invalid StateUTTaxRate"
            '''
            df.loc[~df["StateUTTaxRate"].isin(
                [0, 0.05, 1.5, 2.5, 6, 9, 14, 6.0, 9.0, 14.0]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_SGST/UT_Rt"
            #df.drop('Temp', axis=1, inplace=True)
            # ---------For Invalid CGSTRate-------------------
            '''

            try:
                df["Temp"] =  (df["CentralTaxAmount"]) / ((df["TaxableValue"]))*100
            except:
                df["Temp"] = 0
            df["Temp"] = df["Temp"].apply(lambda x: round(x,2))
            df["CentralTaxRate"] = df["CentralTaxRate"].apply(lambda x: round(x,2))

            df.loc[(df["CentralTaxRate"]-df["Temp"]>0.1),"Flash Remark"] = df["Flash Remark"] + "| Invalid_CGST_Rt"
            '''
            df.loc[~df["CentralTaxRate"].isin(
                [0, 0.05, 1.5, 2.5, 6, 9, 14, 6.0, 9.0, 14.0]), "Flash Remark"] = df["Flash Remark"] + "| Invalid_CGST_Rt"
            #df.drop('Temp', axis=1, inplace=True)

            # ----------------Invalid Pre-GST Flag-----------------

            lst = ["Y", "N", "NA", "Yes", "No", "NO", "YES"]
            df.loc[~df["CRDRPreGST"].isin(lst), "Flash Remark"] = df.loc[~df["CRDRPreGST"].isin(
                lst), "Flash Remark"]+"| Invalid_Pre_GST_flag"
            # ----------------Invalid Supplier GSTIN-----------------

            def GSTIN_check(x):
                regex = "(\d{2})(\D{5})(\d{4})(\D{1})(\w{3})"
                if re.search(regex, str(x)):
                    listof_POS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
                                  "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37"]
                    gstin_2 = str(x)[0:2]
                    if ((gstin_2) not in listof_POS):
                        return "InvalidGSTIN"
                    else:
                        return "ValidGSTIN"

                else:
                    return "InvalidGSTIN"

            df["Temp"] = df["SupplierGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            df.loc[df["Temp"] == "InvalidGSTIN",
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Supplier GSTIN"
            try:
                df.loc[(df["Temp"] != "InvalidGSTIN") & (df["SupplierGSTIN"].str.len() != 15) & (
                    df["SupplierGSTIN"] != "NA"), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Supplier GSTIN"
            except:
                traceback.print_exc()

            # ----------------Invalid Customer or RecipientGSTIN GSTIN-----------------

            df["Temp"] = df["RecipientGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            df.loc[df["Temp"] == "InvalidGSTIN",
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Recipient GSTIN"

            try:
                df.loc[(df["RecipientGSTIN"].str.len() != 15) & (df["Temp"] != "InvalidGSTIN"),
                       "Flash Remark"] = df["Flash Remark"] + "| Invalid_Recipient GSTIN"
            except:
                traceback.print_exc()

            # ----------------Invalid OriginalSupplierGSTIN-----------------
            df["Temp"] = df["OriginalSupplierGSTIN"].apply(
                lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            df.loc[df["Temp"] == "InvalidGSTIN",
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid OriginalSupplierGSTIN"
            try:
                df.loc[(df["OriginalSupplierGSTIN"].str.len() != 15) & (df["OriginalSupplierGSTIN"] != "NA") & (
                    df["Temp"] != "InvalidGSTIN"), "Flash Remark"] = df["Flash Remark"] + "| Invalid OriginalSupplierGSTIN"
            except:
                traceback.print_exc()

            # ----------------Invalid CessRate------------------------------
            listof_CessRate = ["NA", "0", "60", "12", "71", "65", "61", "21", "5", "12.5", "72",
                               "17", "11", "290", "49", "160", "142", "20", "204", "96", "89", "15", "1", "3", "15"]
            df["CessRateSpecific"] = df["CessRateSpecific"].apply(
                lambda x: str(x).split(".")[0] if x != "NA" else x)
            df.loc[~df["CessRateSpecific"].isin(
                listof_CessRate), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Cess_Rt"

            # ----------------Invalid ReverseChargeFlag-----------------
            lst = ["Y", "N", "NA", "y", "n"]
            df.loc[~df["ReverseChargeFlag"].isin(
                lst), "Flash Remark"] = df["Flash Remark"] + "| Invalid_RCMFlag"

            # ----------------Invalid ITC Flag-----------------

            lst = ["T1", "T2", "NA", "T3", "T4"]
            df.loc[~df["ITCReversalIdentifier"].isin(
                lst), "Flash Remark"] = df["Flash Remark"] + "| Invalid_ITCFlag"

            # ----------------Invalid Bill Of Entry-----------------
            lst = ["IMPG", "SEZG"]
            try:
                df["BillOfEntry"] = df["BillOfEntry"].apply(
                    lambda x: str(x).split(".")[0])

            except:
                df["BillOfEntry"] = df["BillOfEntry"].apply(lambda x: str(x))

            try:
                df.loc[(df["BillOfEntry"] != "NA") & (df["BillOfEntry"].str.len(
                ) != 7), "Flash Remark"] = df["Flash Remark"] + "| Invalid Length of BillOfEntry Number "
            except:
                traceback.print_exc()
            df.loc[(df["BillOfEntry"] == "NA") & (df["SupplyType"].isin(
                lst)), "Flash Remark"] = df["Flash Remark"] + "| Invalid_BOE"

            # --------------------------------Supplier and Customer GSTIN are same---------------------------------------------------------
            df.loc[df["SupplierGSTIN"] == df["RecipientGSTIN"],
                   'Flash Remark'] = df["Flash Remark"]+"| Supplier and Recipient GSTIN are same"

            # ----------------Invalid ITC Eligibility Indicator-----------------
            lst = ["IS", "IG", "NO", "CG"]
            df.loc[~df["EligibilityIndicator"].isin(
                lst), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Eligibilitiy_Indicator"

            # -------------CGST and SGST Rates cannot be different-------------
            df.loc[(df["CentralTaxRate"] != df["StateUTTaxRate"]),
                   "Flash Remark"] = df["Flash Remark"] + "| CGST_rate≠SGST_rate"

            # -------------CGST and SGST Amounts cannot be different-------------
            df.loc[(df["CentralTaxAmount"] != df["StateUTTaxAmount"]),
                   "Flash Remark"] = df["Flash Remark"] + "| CGST≠SGST"

            # ------------Invalid document Number length-------------------------
            try:
                df.loc[df["DocumentNumber"].str.len() > 16, "Flash Remark"] = df["Flash Remark"] + \
                    "| Invalid DocumnetNumber-DocumentNumber exceeds 16 digits"
            except:
                try:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).split(".")[0])
                except:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x))

                df.loc[df["DocumentNumber"].str.len() > 16, "Flash Remark"] = df["Flash Remark"] + \
                    "| Invalid DocumnetNumber-DocumentNumber exceeds 16 digits"

            # ------------Invalid document Number length-------------------------
            try:
                df.loc[(df["OriginalDocumentNumber"].str.len() > 16) & ((df["OriginalDocumentNumber"] != "NA")),
                       "Flash Remark"] = df["Flash Remark"] + "| Invalid OriginalDocumentNumber-OriginalDocumentNumberDocumentNumber exceeds 16 digits"
            except:
                df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                    lambda x: str(x).split(".")[0])
                df.loc[df["OriginalDocumentNumber"].str.len() > 16, "Flash Remark"] = df["Flash Remark"] + \
                    "| Invalid OriginalDocumentNumber-OriginalDocumentNumber exceeds 16 digits"

            # ------------Original Document number or Original Document Date is missing-------------
            lst1 = ["CR", "DR"]
            df.loc[(df["DocumentType"].isin(lst1)) & (df["OriginalDocumentNumber"]
                                                      == "NA"), 'Flash Remark'] = df["Flash Remark"]+"| Invalid_Org_Doc_No"
            df.loc[(df["DocumentType"].isin(lst1)) & (df["OriginalDocumentDate"]
                                                      == "NA"), 'Flash Remark'] = df["Flash Remark"]+"| Invalid_Org_Doc_Dt"

            # ------------Taxable Value should be 0 or BLANK for Exempt, Non-GST or NIL rated supply------------
            #df.loc[(((df["SupplyType"] =="EXT") | (df["SupplyType"] =="NON") | (df["SupplyType"] =="NSY")) & (df["TaxableValue"]!=0)),"Flash Remark"] = df["Flash Remark"] + "| Exempt/Non-GST cannot have taxable value"

            # ------------Taxable Value should not be 0 or BLANK for NIL rated supply-------------
            df.loc[((df["SupplyType"] == "NIL") & (df["TaxableValue"] == 0)),
                   "Flash Remark"] = df["Flash Remark"] + "| NIL rated cannot have taxable value"

            # ------------Invoice Value should not be 0 or BLANK for Exempt, Non-GST or NIL rated supply-------------
            df.loc[(((df["SupplyType"] == "EXT") | (df["SupplyType"] == "NON") | (df["SupplyType"] == "NSY") | (df["SupplyType"] == "NIL")) & (
                df["InvoiceValue"] == 0)), "Flash Remark"] = df["Flash Remark"] + "| For Exempt/Non-GST/NIL, Inv_Val cannot be 0"

            # ---------For Invalid CommonSupplyIndicator-------------
            lst = ["YL", "NL", "Y", "N", "NA"]
            df.loc[~df["CommonSupplyIndicator"].isin(
                lst), "Flash Remark"] = df["Flash Remark"] + "| Invalid_Common_Ind"

            # ----------8 digit HSN is required for export/import------------------------

            # ----------In case of SEZ, CGST and SGST cannot be applied-----------------------------
            df.loc[(((df["SupplyType"] == "SEZG") | (df["SupplyType"] == "SEZS")) & (
                (df["CentralTaxAmount"]+df["StateUTTaxAmount"]) != 0)), "Flash Remark"] = df["Flash Remark"] + "| SEZ_Nil CGST/SGST"

            # ---------------------------------------
            df.loc[(((df["SupplyType"] == "NIL") | (df["SupplyType"] == "EXT") | (df["SupplyType"] == "NON") | (df["SupplyType"] == "NSY")) & (
                (df["CentralTaxAmount"]+df["StateUTTaxAmount"]) != 0)), "Flash Remark"] = df["Flash Remark"] + "| Capital Goods flag cannot be there for services"

            # ----------In case of Imports, CGST and SGST cannot be applied-----------------------------
            df.loc[(((df["SupplyType"] == "IMPG") | (df["SupplyType"] == "IMPS")) & ((df["CentralTaxAmount"] +
                    df["StateUTTaxAmount"]) != 0)), "Flash Remark"] = df["Flash Remark"] + "| Imports_NoCGST+SGSTallowed"

            # ----------Capital Goods flag cannot be there for services-----------------------------
            #df["HSNorSAC"]=df["HSNorSAC"].apply(lambda x:str(x).split(".")[0] if x!="NA" else x)
            # # print((df["HSNorSAC"].str[0:2]))
            try:
                try:
                    df["HSNorSAC"] = df["HSNorSAC"].apply(
                        lambda x: str(x).split(".")[0])
                except:
                    df["HSNorSAC"] = df["HSNorSAC"].apply(lambda x: str(x))

                df.loc[((df["HSNorSAC"].str[0:2] == "99") & (df["EligibilityIndicator"] ==
                        "CG")), "Flash Remark"] = df["Flash Remark"] + "| CGFlag_for services"
            except:
                df["HSNorSAC"] = df["HSNorSAC"].apply(
                    lambda x: str(x) if x != "NA" else x)
                df.loc[((df["HSNorSAC"].str[0:2] == "99") & (df["EligibilityIndicator"] ==
                        "CG")), "Flash Remark"] = df["Flash Remark"] + "| CGFlag_for services"

            # ----------Input flag cannot be there for services-----------------------------Input services flag cannot be there for goods
            try:
                df.loc[((df["HSNorSAC"].str[0:2] != "99") & (df["HSNorSAC"].str[0:2] != "NA") & (
                    df["EligibilityIndicator"] == "IS")), "Flash Remark"] = df["Flash Remark"] + "| Incorrect flag in case of goods"
            except:
                df["HSNorSAC"] = df["HSNorSAC"].apply(
                    lambda x: str(x) if x != "NA" else x)
                df.loc[((df["HSNorSAC"].str[0:2] != "99") & (df["HSNorSAC"].str[0:2] != "NA") & (
                    df["EligibilityIndicator"] == "IS")), "Flash Remark"] = df["Flash Remark"] + "| Incorrect flag in case of goods"

            # ----------Input flag cannot be there for goods-----------------------------
            try:

                df.loc[(df["HSNorSAC"].str[0:2] == "99") & (df["EligibilityIndicator"] == "IG"),
                       "Flash Remark"] = df["Flash Remark"] + "| Incorrect flag in case of services"
            except:
                df["HSNorSAC"] = df["HSNorSAC"].apply(
                    lambda x: str(x) if x != "NA" else x)
                df.loc[(df["HSNorSAC"].str[0:2] == "99") & (df["EligibilityIndicator"] == "IG"),
                       "Flash Remark"] = df["Flash Remark"] + "| Incorrect flag in case of services"
            try:
                # ----------Tax Rate and Tax Amount can not to be 0 or BLANK in case Supply Type is TAX or SEZ or DTA or DXP--------------
                df.loc[(((df["SupplyType"] == "TAX") | (df["SupplyType"] == "SEZ") | (df["SupplyType"] == "DTA") | (df["SupplyType"] == "DXP")) & ((df["StateUTTaxAmount"]+df["CentralTaxAmount"]+df["IntegratedTaxAmount"]
                        == 0) & (df["StateUTTaxRate"]+df["CentralTaxRate"]+df["IntegratedTaxRate"] == 0))), "Flash Remark"] = df["Flash Remark"] + "| TaxRate/Tax Amt is NIL_SupplyType_TAXorSEZorDTAorDXP_"

                # ----------Tax Rate and Tax Amount need to be 0 or BLANK in case Supply Type is NIL or NON or EXT or COM--------------
                df.loc[(((df["SupplyType"] == "NIL") | (df["SupplyType"] == "NON") | (df["SupplyType"] == "EXT") | (df["SupplyType"] == "COM")) & ((df["StateUTTaxAmount"]+df["CentralTaxAmount"] +
                        df["IntegratedTaxAmount"] != 0) & (df["StateUTTaxRate"]+df["CentralTaxRate"]+df["IntegratedTaxRate"] != 0))), "Flash Remark"] = df["Flash Remark"] + "| Tax_shud_be_0"

                # ----------TaxableValue/TaxAmount cannot be negative as document types other than CR or RCR or RFV--------------
                df.loc[(((df["DocumentType"] != "CR") & (df["DocumentType"] != "RCR") & (df["DocumentType"] != "RFV")) & ((df["TaxableValue"] < 0) | (df["StateUTTaxAmount"] < 0) | (df["CentralTaxAmount"] < 0) | (
                    df["IntegratedTaxAmount"] < 0) | (df["InvoiceValue"] < 0) | (df["CessAmountSpecific"] < 0) | (df["CessAmountAdvalorem"] < 0))), "Flash Remark"] = df["Flash Remark"] + "| Values cannot be negative"
            except:
                traceback.print_exc()

            # -------------------Original document to be reported in case of RFV---------------
            df.loc[(df["DocumentType"] == "RFV") & (df["OriginalDocumentNumber"] == "NA"),
                   "Flash Remark"] = df["Flash Remark"] + "| Missing Original details"

            # ----------Reverse charge flag cannot be Y, when services are procured from SEZ--------------
            df.loc[(df["SupplyType"] == "SEZS") & (df["ReverseChargeFlag"] == "Y"),
                   "Flash Remark"] = df["Flash Remark"] + "| Invalid RCFlag_SEZ supply"

            # ----------Tax amount cannot be blank in case value in Cess amount is available--------------
            try:
                df.loc[(df["CessAmountSpecific"]+df["CessRateAdvalorem"] != 0) & (df["TaxableValue"]
                                                                                  == 0), "Flash Remark"] = df["Flash Remark"] + "| Invalid TaxAmt_CessAmt available"
            except:
                traceback.print_exc()

            #df["Flash Remark"]=df["Flash Remark"].apply(lambda x:x.replace("NA|",""))

            # df.replace("NA","",inplace=True)

            # _____________________MAster Validations__________________________

            # -------------Unit of Measurement-----------------------
            UQC = []
            for x in list(uqc["UQC"]):
                UQC.append(x)
                UQC.append(str(x).split("-")[0])

            df.loc[~df["UnitOfMeasurement"].isin(
                UQC), "Flash Remark"] = df["Flash Remark"]+"| Invalid_UOM"

            #uqc["UQC"]=uqc["UQC"].apply(lambda x:str(x).split("-")[0])

            # UQC=list(uqc["UQC"])

            #df.loc[~df["UnitOfMeasurement"].isin(UQC),"Flash Remark"]=df["Flash Remark"]+"| Invalid_UOM"

            #df["UnitOfMeasurement"]=df["UnitOfMeasurement"].apply(lambda x:"OTH" if x=="NA" else x)

            if len(df[df["PortCode"] == "NA"]) != len(df):
                PCODE = list(port_code["Code"])
                df.loc[(~df["PortCode"].isin(PCODE)) & (df["PortCode"] != "NA"),
                       "Flash Remark"] = df["Flash Remark"]+"| Invalid_Portcode"

            doc_type = inward_doc_type["Corresponding Document type"]
            supp_type = inward_supp_type["Corresponding supply type"]

            # ---------------------Document Type and Supply Type validations from Masters------------------------------

            df.loc[(~df["DocumentType"].isin(doc_type)),
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Doc_Type"

            df.loc[(~df["SupplyType"].isin(supp_type)),
                   "Flash Remark"] = df["Flash Remark"]+"| Invalid_Supply_Type"

            # ---------------------HSN Validation from Masters------------------------------

            hsn_g = hsn_goods["HSN Code"]

            hsn_s = hsn_services["Service Code (Tariff)"]
            hsn_list = []

            for x in hsn_g:
                hsn_list.append(str(x))
                hsn_list.append(str(x)[0:6])
                hsn_list.append(str(x)[0:4])
                hsn_list.append(str(x)[0:2])
            for x in hsn_s:
                hsn_list.append(str(x))
                hsn_list.append(str(x)[0:4])
                hsn_list.append(str(x)[0:2])

            # ------------8 digit HSN is required for export/import of Goods----------
            # ------------------  Invalid HSN or SAC--------------------------------------

            if len(df[df["HSNorSAC"] == "NA"]) != len(df):
                #      df.loc[(~df["HSNorSAC"].isin(hsn_list)) | (df["HSNorSAC"]=="NA") ,"Flash Remark"]=df["Flash Remark"]+"| Invalid_HSN"
                df.loc[(~df["HSNorSAC"].isin(hsn_list)),
                       "Flash Remark"] = df["Flash Remark"]+"| Invalid_HSN"

                lst = ["EXPWT", "EXPT"]
                try:

                    df.loc[(df["SupplyType"].isin(lst)) & (df["HSNorSAC"].str[0:2] != "99") & (df["HSNorSAC"].str.len(
                    ) != 8), "Information Error"] = df["Information Error"] + "| Export_Import-8 digit HSN"
                except:
                    df["HSNorSAC"] = df["HSNorSAC"].apply(
                        lambda x: str(x) if x != "NA" else x)
                    df.loc[(df["SupplyType"].isin(lst)) & (df["HSNorSAC"].str[0:2] != "99") & (df["HSNorSAC"].str.len(
                    ) != 8), "Information Error"] = df["Information Error"] + "| Export_Import-8 digit HSN"

            # ----------------------Replacing all values where blank is there with 0 -------------------------
            df["Quantity"].fillna(0, inplace=True)

            df["IntegratedTaxAmount"].fillna(0, inplace=True)
            df["StateUTTaxAmount"].fillna(0, inplace=True)
            df["CentralTaxAmount"].fillna(0, inplace=True)

            df["StateUTTaxRate"].fillna(0, inplace=True)
            df["IntegratedTaxRate"].fillna(0, inplace=True)
            df["CentralTaxRate"].fillna(0, inplace=True)

            df["AvailableCGST"].fillna(0, inplace=True)
            df["AvailableIGST"].fillna(0, inplace=True)
            df["AvailableSGST"].fillna(0, inplace=True)

            df["InvoiceValue"].fillna(0, inplace=True)
            df["TaxableValue"].fillna(0, inplace=True)
            #progress['value'] = 30

            df.fillna("NA", inplace=True)

            # ----------------Changing Date columns to required format-----------------
            #df["DocumentDate"] = df["DocumentDate"].apply(lambda x:str(x).split()[0])
            #df["OriginalDocumentDate"] = df["OriginalDocumentDate"].apply(lambda x:str(x).split()[0])
            #df["PurchaseVoucherDate"] = df["PurchaseVoucherDate"].apply(lambda x:str(x).split()[0])
            #df["BillOfEntryDate"] = df["BillOfEntryDate"].apply(lambda x:str(x).split()[0])
            #df["PaymentDate"] = df["PaymentDate"].apply(lambda x:str(x).split()[0])

            # ----------------Changing DocumentNumber,OriginaldocumentNumber to required Text format-----------------
            #df["DocumentNumber"] = df["DocumentNumber"].astype(str)
            #df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].astype(str)
            df["POS"] = df["POS"].astype(str)
            #progress['value'] = 40

            # ----------------Original invoice prior to FY 19-20-----------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["DocumentDate"].fillna("NA", inplace=True)
            #AfterMonthYearCompare = InputFinancialYear*100+9
            #PriorMonthYearToCompare = InputFinancialYear*100+4
            # PreviousMonthYearToCompare=(InputFinancialYear-1)*100+4

            df["OriginalDocumentDate"].fillna("NA", inplace=True)

            def FY1(x):
                try:
                    if x != "NA":
                        try:
                            return(int(int(x)/10000))
                        except:
                            return 0
                except:
                    return 0

            df["temp1"] = df["ReturnPeriod"].apply(lambda x: FY1(x))
            df["temp"] = ""
            for i in range(len(df)):
                if(df["DocumentType"][i] == "CR") & (df["Temp"][i] == "ValidDate"):
                    if(df["OriginalDocumentDate"][i] != "NA"):
                        try:
                            df["temp"][i] = int(df["OriginalDocumentDate"][i].split(
                                "-")[0] + df["OriginalDocumentDate"][i].split("-")[1])
                        except:
                            try:
                                df["temp"][i] = int(df["OriginalDocumentDate"][i].split(
                                    "//")[0] + df["OriginalDocumentDate"][i].split("//")[1])
                            except:
                                try:
                                    df["temp"][i] = int(df["OriginalDocumentDate"][i].split(
                                        ".")[0] + df["OriginalDocumentDate"][i].split(".")[1])
                                except:
                                    df["temp"][i] = 0
                        # if((df["temp"][i] > AfterMonthYearCompare) & (df["temp"][i] <PriorMonthYearToCompare) & (df["temp1"][i]>9)):

                            #df["Flash Remark"] = df["Flash Remark"][i]+ "| CN reorted after sep of next FY"
                        # if((df["temp"][i] > AfterMonthYearCompare) & (df["temp"][i] < PreviousMonthYearToCompare) & (df["temp1"][i]<=9)):

                            #df["Flash Remark"] = df["Flash Remark"][i]+ "| CN reorted after sep of next FY"
            df.drop('temp', axis=1, inplace=True)
            df.drop('Temp', axis=1, inplace=True)

            # ----------------2.7.Invoice are dated prior FY 19-20----------------------------------------------------------------------

            def transform(x):
                if str(x).find("-") != -1:
                    try:
                        return(int(str(x).split("-")[0]+str(x).split("-")[1]))
                    except:
                        return(0)
                else:
                    try:
                        return(int(str(x).split("/")[0]+str(x).split("/")[1]))
                    except:
                        return(0)

            df["temp"] = df["DocumentDate"].apply(lambda x: transform(x))

            #df["EY Remark"] = df["EY Remark"] + df["temp"].apply(lambda x:"| Pending" if x < PriorMonthYearToCompare else "")

            #df.loc[(df["temp"]<PriorMonthYearToCompare) & (df["temp1"]>9),"Flash Remark"]=df["Flash Remark"]+"| Invoice pretaining to Previous FY"

            #df.loc[(df["temp"]<PreviousMonthYearToCompare) & (df["temp1"]<=9),"Flash Remark"]=df["Flash Remark"]+"| Invoice pretaining to Previous FY"

            #df["Flash Remark"] = df["Flash Remark"]+ df["temp"].apply(lambda x:"| Invoice pretaining to Previous FY" if x < PriorMonthYearToCompare and x!=0 else "")
            df.drop('temp', axis=1, inplace=True)
            df.drop('temp1', axis=1, inplace=True)
            #progress['value'] = 50

            # ---------2.8.Invoice are dated further to the return period---------------------------------------------------------------

            def transform(x):

                if str(x).find("-") != -1:
                    try:
                        return(int(str(x).split("-")[0]+str(x).split("-")[1]))
                    except:
                        return(0)
                if str(x).find("//") != -1:
                    try:
                        return(int(str(x).split("/")[0]+str(x).split("/")[1]))
                    except:
                        return(0)
                if str(x).find(".") != -1:
                    try:
                        return(int(str(x).split(".")[0]+str(x).split(".")[1]))
                    except:
                        return(0)
                else:
                    return(0)

            def transform1(x):
                try:
                    if len(str(x)) == 5:
                        return("0"+str(x))
                    else:
                        return(str(x))
                except:
                    return(0)
            try:
                df["ReturnPeriodTemp"] = df["ReturnPeriod"].apply(
                    lambda x: transform1(x))
                df["ReturnPeriodTemp"] = df["ReturnPeriodTemp"].apply(lambda x: str(
                    x)[2:6]) + df["ReturnPeriodTemp"].apply(lambda x: str(x)[0:2])
                df["ReturnPeriodTemp"] = df["ReturnPeriodTemp"].apply(
                    lambda x: int(x) if x != "NA" else 0)
                df["DocumentDateTemp"] = df["DocumentDate"].apply(
                    lambda x: transform(x))
                current_date = transform(str(datetime.now()).split()[0])
                #df.loc[df["ReturnPeriodTemp"] < df["DocumentDateTemp"],'EY Remark'] = df["EY Remark"]+"| Pending"
                df.loc[(df["ReturnPeriodTemp"] < df["DocumentDateTemp"]) & (df["ReturnPeriodTemp"] != 0) & (
                    df["DocumentDateTemp"] != 0), 'Flash Remark'] = df["Flash Remark"]+"| Inv beyond Ret_Pd"
                df.loc[(df["ReturnPeriodTemp"] > current_date) & (df["ReturnPeriodTemp"] != 0),
                       'Flash Remark'] = df["Flash Remark"]+"| TaxPeriod_beyond current date"

                df.drop("ReturnPeriodTemp", axis=1, inplace=True)
                df.drop("DocumentDateTemp", axis=1, inplace=True)
            except:
                try:
                    df.drop("ReturnPeriodTemp", axis=1, inplace=True)
                    df.drop("DocumentDateTemp", axis=1, inplace=True)
                except:
                    traceback.print_exc()

            # ---------------------------------Concatenating two dataframes of Invalid & Valid Document Dates---------------------------

            #progress['value'] = 60

            # -------------------------Cancelled Documents--------------------------
            dff1 = df[df["SupplyType"] == "CAN"]
            l2 = dff1['DocumentNumber'].unique().tolist()

            df["TaxableValue"] = df["TaxableValue"].apply(pd.to_numeric)

            for j in l2:
                aa = df.loc[(df['DocumentNumber'] == j) & (
                    df['TaxableValue'] < 0), 'TaxableValue'].unique()
                for ii in aa:
                    count1 = df.loc[(df['DocumentNumber'] == j) & (
                        df['TaxableValue'] == abs(ii)), 'Flash Remark'].count()
                    count = df.loc[(df['DocumentNumber'] == j) & (
                        df['TaxableValue'] == ii), 'Flash Remark'].count()
                    if count1 >= count:
                        df.loc[(df['DocumentNumber'] == j) & (df['TaxableValue'] == ii),
                               'Flash Remark'] = df['Flash Remark']+'| Cancelled Documents'
                        df.loc[(df['DocumentNumber'] == j) & (
                            df['TaxableValue'] == ii), 'EY Remark'] = 'Pending'
                        for i in range(len(df)):
                            if (df['DocumentNumber'].iloc[i] == j) & (df['TaxableValue'].iloc[i] == abs(ii)) & (count != 0):
                                df['Flash Remark'].iloc[i] = df['Flash Remark'][i] + \
                                    '| Cancelled Documents'
                                df['EY Remark'].iloc[i] = 'Pending'
                                count = count-1
                    if ((count1 < count) & (count1 != 0)):
                        df.loc[(df['DocumentNumber'] == j) & (df['TaxableValue'] == abs(
                            ii)), 'Flash Remark'] = df['Flash Remark']+'| Cancelled Documents'
                        df.loc[(df['DocumentNumber'] == j) & (
                            df['TaxableValue'] == abs(ii)), 'EY Remark'] = 'Pending'
                        for i in range(len(df)):
                            if (df['DocumentNumber'].iloc[i] == j) & (df['TaxableValue'].iloc[i] == (ii)) & (count1 != 0):
                                df['Flash Remark'].iloc[i] = df['Flash Remark'][i] + \
                                    '| Cancelled Documents'
                                df['EY Remark'].iloc[i] = 'Pending'
                                count1 = count1-1

            #      # print("Done with Cancelled Document Validation.....")
            # ------------------------Similar document number having multiple dates------------------------------------------------

            # ----------------Similar Original Document number having multiple original Document Dates------------------------------------------------
            OriginalDocSeries = df.groupby('OriginalDocumentNumber')[
                'OriginalDocumentDate'].nunique()
            newdf = OriginalDocSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfOriginalDocNo = newdf[newdf["OriginalDocumentDate"]
                                        > 1]["OriginalDocumentNumber"].tolist()

            for i in ListOfOriginalDocNo:
                df.loc[df["OriginalDocumentNumber"] == i,
                       'EY Remark'] = df.loc[df["OriginalDocumentNumber"] == i, 'EY Remark'] + "| Pending"
                df.loc[df["OriginalDocumentNumber"] == i, 'Flash Remark'] = df.loc[df["OriginalDocumentNumber"]
                                                                                   == i, 'Flash Remark'] + "| Similar OriginalDocumentNumber having multiple dates"

        #     # print("Done with Original Document having Similar Original DocDates ...Validation.....")

            # ----------------Adding initial zero if length of the below columns is 1 and to required Text format: POS, ReturnPeriod-----------------
            #df["POS"] = df["POS"].apply(lambda x:"0"+x if len(x) == 1 else x)
            #df["ReturnPeriod"] = df["ReturnPeriod"].apply(lambda x:"0"+x if len(x) == 5 else x)

            # -------------------------------------Originial invoice date and document number is missing-------------
            #df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & ((df["OriginalDocumentNumber"] == "NA") | (df["OriginalDocumentDate"] == "NA"))) ,'EY Remark'] = df["EY Remark"]+"| Pending"
            #df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & ((df["OriginalDocumentNumber"] == "NA") | (df["OriginalDocumentDate"] == "NA"))) ,'Flash Remark'] = df["Flash Remark"]+"| Original document number/date is missing"

            df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & (
                (df["DocumentNumber"] == "NA") | (df["DocumentDate"] == "NA"))), 'EY Remark'] = df["EY Remark"]+"| Pending"
            df.loc[(((df["DocumentType"] == "CR") | (df["DocumentType"] == "DR")) & ((df["DocumentNumber"] == "NA") | (
                df["DocumentDate"] == "NA"))), 'Flash Remark'] = df["Flash Remark"]+"| Document number/date is missing"

            # --------------------------------Supplier and Customer GSTIN are same---------------------------------------------------------
            #df.loc[df["SupplierGSTIN"] == df["RecipientGSTIN"] ,'EY Remark'] = df["EY Remark"]+"| Pending"
            #df.loc[df["SupplierGSTIN"] == df["RecipientGSTIN"] ,'Flash Remark'] = df["Flash Remark"]+"| Supplier and Recipient GSTIN are same"

            # --------------------------------Bill of Entry details/Port code to be provided.-------------------------------------------
            df.loc[(df["PortCode"] == "NA") & (df["BillOfEntry"] == "NA") & (
                df["SupplyType"] == "IMPG"), 'Flash Remark'] = df["Flash Remark"]+"| Invalid_BOE/Portcode"

            # -----------------------------POS is different from recipient GSTIN--------------------------------------------------------------------------
            df.loc[(df["POS"] != "NA") & (df["POS"] != df["RecipientGSTIN"].apply(
                lambda x:str(x)[0:2])), "EY Remark"] = df["EY Remark"] + "| Pending"
            #df.loc[(df["POS"] != "NA") & (df["POS"] != df["RecipientGSTIN"].apply(lambda x:str(x)[0:2])), "Information Error"] = df["Information Error"] + "| POS is different from recipient GSTIN - please confirm the recipient state"
            #df.loc[df["POS"] == "NA", "Flash Remark"] = df["Flash Remark"] + "| POS is blank- Kindly confirm"

            # ---------------------------Remaining Transactions:Stock Transfer--------------------------------------------------------------------------
            df.loc[df["SupplierGSTIN"].apply(lambda x:str(x)[3:10]) == df["RecipientGSTIN"].apply(
                lambda x:str(x)[3:10]), 'EY Remark'] = df["EY Remark"] + "| Stock Transfer"
            try:
                # --------------------------SupplyType1==EXT--------------------------------------------------------------------------
                df.loc[((df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                        df["StateUTTaxAmount"] == 0)), "EY Remark"] = df["EY Remark"] + "| Ignore"
                df.loc[((df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                        df["StateUTTaxAmount"] == 0)), "Flash Remark"] = df["Flash Remark"] + "| Exempt"
                #progress['value'] = 80

                df.loc[(df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                                                      df["StateUTTaxAmount"] != 0), "EY Remark"] = df["EY Remark"] + "| Pending"
                df.loc[(df["SupplyType"] == "EXT") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"]
                                                      != 0), "Flash Remark"] = df["Flash Remark"] + "| Tax charged in exempt category supplies"

                # --------------------------SupplyType=NIL----------------------------
                df.loc[(df["SupplyType"] == "NIL") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                                                      df["StateUTTaxAmount"] == 0), "EY Remark"] = df["EY Remark"] + "| Ignore"
                df.loc[(df["SupplyType"] == "NIL") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                                                      df["StateUTTaxAmount"] == 0), "Flash Remark"] = df["Flash Remark"] + "| Nil rated"

                df.loc[(df["SupplyType"] == "NIL") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                                                      df["StateUTTaxAmount"] != 0), "EY Remark"] = df["EY Remark"] + "| Pending"
                df.loc[(df["SupplyType"] == "NIL") & (df["IntegratedTaxAmount"] + df["CentralTaxAmount"] +
                                                      df["StateUTTaxAmount"] != 0), "Flash Remark"] = df["EY Remark"] + "| Tax charged in NIL rated inward supplies"
            except:
                traceback.print_exc()

            # ---------3.4.Eligibility Indicator-------------
            # df.loc[df["HSNorSAC"].apply(lambda x:str(x)[0:2] == 99),"EligibilityIndicator"] = "IS"
            # df.loc[df["HSNorSAC"].apply(lambda x:str(x)[0:2] != 99),"EligibilityIndicator"] = "IG"
            # df.loc[df["ReverseChargeFlag"] == "Y","EligibilityIndicator"] = "IS"
            # df.loc[df["SupplierGSTIN"].isin(ListOfIneligibleGSTINs),"EligibilityIndicator"] = "NO"

            # ---------3.9.For Invalid Tax Rate-------------
            #df["IntegratedTaxRate"] = df["IntegratedTaxRate"].round(2)
            #df.loc[~df["IntegratedTaxRate"].isin([0,28,18,12,5,0.0,28.0,18.0,12.0,5.0]),"Flash Remark"] = df["Flash Remark"] + "| Invalid Integrated Tax Rate "

            #df["CentralTaxRate"] = df["CentralTaxRate"].round(2)
            #df.loc[~df["CentralTaxRate"].isin([0,14,9,6,2.5,0.0,14.0,9.0,6.0]),"Flash Remark"] = df["Flash Remark"] + "| Invalid_CGST_Rt"

            #df["StateUTTaxRate"] = df["StateUTTaxRate"].round(2)
            #df.loc[~df["StateUTTaxRate"].isin([0,14,9,6,2.5,0.0,14.0,9.0,6.0]),"Flash Remark"] = df["Flash Remark"] + "| Invalid StateUTTaxRate"

            # ---------3.10.Invoice Value-------------

            # ---------3.11.Available taxes-----------

            # df.loc[((df["AvailableIntegratedTaxAmount"] != df["IntegratedTaxAmount"]) & (df["EligibilityIndicator"] != "NO")),"EY Remark"] = df["EY Remark"] + "| Mismatch in IntegratedTaxAmount & Available IntegratedTaxAmount"
            # df.loc[((df["AvailableCentralTaxAmount"] != df["CentralTaxAmount"]) & (df["EligibilityIndicator"] != "NO")),"EY Remark"] = df["EY Remark"] + "| Mismatch in CentralTaxAmount & Available CentralTaxAmount"
            # df.loc[((df["AvailableStateUTTaxAmount"] != df["StateUTTaxAmount"]) & (df["EligibilityIndicator"] != "NO")),"EY Remark"] = df["EY Remark"] + "| Mismatch in StateUTTaxAmount & Available StateUTTaxAmount"

            # -----------------------------------Reverse Charge Flag--------------------------
            #df.loc[(df["ReverseChargeFlag"] == "Y") & (df["SupplierGSTIN"] != "NA"),"DocumentType"] = "INV"

            # ---------3.3.Replacing special characters in Document Numbers,Original doc No-------------
            df["DocumentNumber"] = df["DocumentNumber"].apply(
                lambda x: str(x).replace(".", ""))
            df["DocumentNumber"] = df["DocumentNumber"].apply(
                lambda x: str(x).replace("&", ""))
            df["DocumentNumber"] = df["DocumentNumber"].apply(
                lambda x: str(x).replace(" ", ""))
            df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                lambda x: str(x).replace(".", ""))
            df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                lambda x: str(x).replace("&", ""))
            df["OriginalDocumentNumber"] = df["OriginalDocumentNumber"].apply(
                lambda x: str(x).replace(" ", ""))
            # df["DocumentNumber"]=df["DocumentNumber"].apply(lambda x:str(x).replace("\\","/"))
            # df["DocumentNumber"]=df["DocumentNumber"].apply(lambda x:str(x).replace("00:00:00",""))
            # df["DocumentNumber"]=df["DocumentNumber"].apply(lambda x:str(x).split("Dt")[0])

            # ---------3.12.Import of goods------------------------
            #df.loc[((df["SupplyType"] == "IMPG") & ((df["PortCode"] == "NA") | (df["BillOfEntry"] == "NA") | (df["BillOfEntryDate"] == "NA"))),"EY Remark"] = df["EY Remark"]+"| Pending"
            #df.loc[((df["SupplyType"] == "IMPG") & ((df["PortCode"] == "NA") | (df["BillOfEntry"] == "NA") | (df["BillOfEntryDate"] == "NA"))),"Flash Remark"] = df["Flash Remark"]+"| Port code/Bill of entry details missing"

            # ---------2.11.Import Of Services----------------------
            df.loc[(df["DocumentType"] == "SLF") & (df["SupplyType"] == "IMPS") & (
                df["SupplierGSTIN"] == "NA") & (df["ReverseChargeFlag"] == "Y"), 'EY Remark'] = "Import of services"

            # ---------2.10.Import Of Goods-------------
            df.loc[(df["DocumentType"] == "INV") & (df["SupplyType"] == "IMPG") & (
                df["SupplierGSTIN"] == "NA"), 'EY Remark'] = "Import of goods"

            # ---------2.14.No supplier GSTIN-------------
            lst = ["IMPG", "IMPS"]
            lst1 = ["SLF"]
            #df.loc[(df["SupplierGSTIN"] == "NA") & (df["EY Remark"] != "Import of goods") & (df["EY Remark"] != "Import of services") & (df['EY Remark'].str.contains("RCM")==False),'EY Remark'] = df["EY Remark"]+"| Ignore"
            try:
                df.loc[(df["SupplierGSTIN"] == "NA") & (~df["SupplyType"].isin(lst)) & (
                    ~df["DocumentType"].isin(lst1)), 'Flash Remark'] = df["Flash Remark"]+"| No supplier GSTIN"
            except:
                traceback.print_exc()

            #df.loc[((df["SupplierGSTIN"] != "NA") & (df["SupplierGSTIN"].str.len() != 15) & (df["SupplierGSTIN"] != "URP")),'EY Remark'] = df["EY Remark"]+"| Pending"
            #df.loc[((df["SupplierGSTIN"] != "NA") & (df["SupplierGSTIN"].str.len() != 15) & (df["SupplierGSTIN"] != "URP")),'Flash Remark'] = df["Flash Remark"]+"| SupplierGSTIN is not proper"

            # ----------IGST,CSGST,SGST - Invalid Taxes applied-------------
            try:
                df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] == df["POS"]) & (
                    df["IntegratedTaxAmount"] != 0) & ((df["CentralTaxAmount"] == 0) & (df["StateUTTaxAmount"] == 0)), 'Flash Remark'] = df["Flash Remark"]+"| IGST cannot be charged"

                #df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] != df["POS"]) & (df["IntegratedTaxAmount"] == 0) & ((df["CentralTaxAmount"] != 0) & (df["StateUTTaxAmount"] != 0)) ,'Information Error'] = df["Information Error"]+"| Why CGST+SGST charged instead of IGST"
            except:
                try:
                    df["SupplierGSTIN"] = df["SupplierGSTIN"].apply(
                        lambda x: str(x) if x != "NA" else x)
                    df["RecipientGSTIN"] = df["RecipientGSTIN"].apply(
                        lambda x: str(x) if x != "NA" else x)
                    df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] == df["POS"]) & (
                        df["IntegratedTaxAmount"] != 0) & ((df["CentralTaxAmount"] == 0) & (df["StateUTTaxAmount"] == 0)), 'Flash Remark'] = df["Flash Remark"]+"| IGST cannot be charged"

                    #df.loc[(df["SupplierGSTIN"].str.len() == 15) & (df["RecipientGSTIN"].str.len() == 15) & (df["SupplierGSTIN"].str[0:2] != df["RecipientGSTIN"].str[0:2]) & (df["IntegratedTaxAmount"] == 0) & ((df["CentralTaxAmount"] != 0) & (df["StateUTTaxAmount"] != 0)) ,'Information Error'] = df["Information Error"]+"| Why CGST+SGST charged instead of IGST"
                except:
                    traceback.print_exc()

            # ---------New - CR & CAN - Pending ------------------------------------------------------------------------------------
            df.loc[((df["SupplyType"] == "CAN") & (df["DocumentType"] ==
                    "CR")), 'EY Remark'] = df["EY Remark"] + "| Pending"

            # ---------------------In case of Inter-State Supply, CGST & SGST/UTGST Tax Rate and/or Tax Amount cannot be applied----
            # ---------------------In case of Intra-State Supply, IGST Tax Rate and and/or Tax Amount cannot be applied---
            # ---------------In case of intra-State Supply, CGST & SGST cannot be zero----------------
            # -------------------In case of inter-State Supply, IGST cannot be zero
            try:
                df["Temp"] = df["SupplierGSTIN"].apply(
                    lambda x: GSTIN_check(str(x)) if x != "NA" else x)
            except:
                traceback.print_exc()
            try:
                df.loc[(df['Temp'] != "InvalidGSTIN") & (df["SupplierGSTIN"].str[0:2] != df["POS"]) & (df["SupplierGSTIN"] != "NA") & (df["POS"] != "NA") & (
                    (df["CentralTaxRate"] != 0) & (df["StateUTTaxRate"] != 0)), 'Flash Remark'] = df["Flash Remark"]+"| C&S cannot be charged"
            except:
                traceback.print_exc()
            try:
                df.loc[(df['Temp'] != "InvalidGSTIN") & (df["SupplierGSTIN"].str[0:2] == df["POS"]) & (df["SupplierGSTIN"] != "NA") & (
                    df["POS"] != "NA") & (df["IntegratedTaxRate"] != 0), 'Flash Remark'] = df["Flash Remark"]+"| IGST_Rt cannot appear in Intra-state supply"
            except:
                traceback.print_exc()
            try:
                df.loc[(df['Temp'] != "InvalidGSTIN") & (df["SupplierGSTIN"].str[0:2] == df["POS"]) & (df["SupplierGSTIN"] != "NA") & (df["POS"] != "NA") & ((df["CentralTaxAmount"] == 0) & (
                    df["StateUTTaxAmount"] == 0)) & ((df["CentralTaxRate"] == 0) & (df["StateUTTaxRate"] != 0)), 'Flash Remark'] = df["Flash Remark"]+"| C&S cannot be 0"
            except:
                traceback.print_exc()
            try:
                df.loc[(df['Temp'] != "InvalidGSTIN") & (df["SupplierGSTIN"].str[0:2] != df["POS"]) & (df["SupplierGSTIN"] != "NA") & (df["POS"] != "NA") & (
                    (df["IntegratedTaxAmount"] == 0) & (df["IntegratedTaxRate"] == 0)), 'Flash Remark'] = df["Flash Remark"]+"| IGST cannot be 0"
            except:
                traceback.print_exc()

            df.drop("Temp", axis=1, inplace=True)

            # ---------------------------------Document Number Length Exceeding 16 Digits------------------------
            try:
                df.loc[df["DocumentNumber"].str.len(
                ) > 16, "Flash Remark"] = df["Flash Remark"] + "| DocumentNumber exceeds 16 digits"
            except:
                try:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x).split(".")[0])
                except:
                    df["DocumentNumber"] = df["DocumentNumber"].apply(
                        lambda x: str(x))

                df.loc[df["DocumentNumber"].str.len(
                ) > 16, "Flash Remark"] = df["Flash Remark"] + "| DocumentNumber exceeds 16 digits"
            # ---------Original Document Number Length Exceeding 16 Digits-------------
            try:
                df.loc[df["OriginalDocumentNumber"].str.len(
                ) > 16, "Flash Remark"] = df["Flash Remark"] + "| OriginalDocumentNumber exceeds 16 digits"
            except:
                traceback.print_exc()
            #progress['value'] = 90

            # ---------------------------------Invalid Tax Rates----------------------------
            # try:
            #    df["Temp"] = ((df["IntegratedTaxAmount"] + df["CentralTaxAmount"] + df["StateUTTaxAmount"]) / (df["TaxableValue"])* 100).round()
            # except:
            #    df["Temp"] = "NA"

            #df["Temp"] = df["Temp"].round()

            #df["Temp2"] = df["StateUTTaxRate"] + df["CentralTaxRate"] + df["IntegratedTaxRate"]
            #df["Temp2"] = df["Temp2"].round()

            #df.loc[((df["Temp"] != df["Temp2"]) & (df["TaxableValue"] != 0)) ,"EY Remark"] =  df["EY Remark"] + "| Pending"
            #df.loc[((df["Temp"] != df["Temp2"]) & (df["TaxableValue"] != 0)) ,"Flash Remark"] =  df["Flash Remark"] + "| Incorrect GST Tax Rates"

            #df.drop('Temp', axis=1, inplace=True)
            #df.drop('Temp2', axis=1, inplace=True)

            # ---------2.3.Tax amount is zero ------------------------------------------------------------------------------------
            try:
                df.loc[(df["InvoiceValue"] == 0) & (df["SupplyType"].isin(
                    Non_zero_supply_type)), 'Flash Remark'] = df["Flash Remark"] + "| InvoiceValue is zero"
                df.loc[df["IntegratedTaxAmount"]+df["CentralTaxAmount"] +
                       df["StateUTTaxAmount"] == 0, 'EY Remark'] = df["EY Remark"] + "| Ignore"

                df.loc[df["IntegratedTaxAmount"]+df["CentralTaxAmount"]+df["StateUTTaxAmount"]
                       == 0, 'Flash Remark'] = df["Flash Remark"] + "| Tax amount is zero"
            except:
                traceback.print_exc()

            # ------------------Invalid Multiple Supply type Combination------------

            lst = list(inward_supp_doctype_combo["Document Type"].map(
                str)+inward_supp_doctype_combo["Supply Type"].map(str))
            lst1 = list(inward_supp_doctype_combo["Supply Type"].map(str))
            lst2 = list(inward_supp_doctype_combo["Document Type"].map(str))
            df["combo"] = df["DocumentType"].map(str)+df["SupplyType"].map(str)
            df.loc[(~df["combo"].isin(lst)) & (df["SupplyType"].isin(lst1)) & (df["DocumentType"].isin(
                lst2)), "Flash Remark"] = df["Flash Remark"]+"| Invalid Document and Supply type Combination"
            df.drop("combo", axis=1, inplace=True)

            # ---------------Same Document cannot have multiple Supply type---------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)

            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["ReturnPeriod"].map(str)
            DocumentSeries = df.groupby('Key_ID')['SupplyType'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["SupplyType"]
                                      > 1]["Key_ID"].tolist()
            lst1 = list(inward_supp_doctype_combo["Supply Type"].map(str))

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"),
                       'Flash Remark'] = df['Flash Remark'] + "| Single_doc-Multiple_Supply type"

            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # ---------------Original Document & Revised Document cannot be reported in same tax period for the same Document Number---------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["Key_ID"] = df["DocumentNumber"].map(
                str)+df["SupplierGSTIN"].map(str)+df["ReturnPeriod"].map(str)
            DocumentSeries = df.groupby('Key_ID')['DocumentNumber'].nunique()
            newdf = DocumentSeries.to_frame()
            newdf.reset_index(inplace=True)

            ListOfDocumentNos = newdf[newdf["DocumentNumber"]
                                      > 1]["Key_ID"].tolist()

            for i in ListOfDocumentNos:
                df.loc[(df["Key_ID"] == i) & (df["Temp"] == "ValidDate"), 'Flash Remark'] = df.loc[(
                    df["Key_ID"] == i) & (df["Temp"] == "ValidDate"), 'Flash Remark'] + "| Single_doc-INV+RNV"

            df.drop('Temp', axis=1, inplace=True)
            df.drop('Key_ID', axis=1, inplace=True)
            df.reset_index(drop=True, inplace=True)

            # --------------Original document for this cancelled record was not reported in this tax period--------
            lst = df["DocumentNumber"].unique()
            df.loc[(df["SupplyType"] == "CAN") & (df["OriginalDocumentNumber"] != "NA") & (~df["OriginalDocumentNumber"].isin(
                lst)), "Flash Remark"] = df["Flash Remark"] + "| Incorrect tax period for cancelled doc"

            # --------------Document Date cannot be prior to Original Document Date-----------------
            df["Temp"] = df["DocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            df["Temp1"] = df["OriginalDocumentDate"].apply(
                lambda x: valid_date(x) if x != "NA" else x)
            try:
                df.loc[(df["Temp"] == "ValidDate") & (df["Temp1"] == "ValidDate") & (df["DocumentDate"] <
                                                                                     df["OriginalDocumentDate"]), "Flash Remark"] = df["Flash Remark"] + "|  Doc_Dt cannot be prior to Org_Doc_Dt"
                df.loc[(df["Temp"] == "ValidDate") & (df["Temp1"] == "ValidDate") & (df["DocumentDate"] > str(
                    datetime.now()).split()[0]), "Flash Remark"] = df["Flash Remark"] + "|  Doc_Dt cannot be prior to Org_Doc_Dt"

            except:
                traceback.print_exc()
            df.drop('Temp', axis=1, inplace=True)
            df.drop('Temp1', axis=1, inplace=True)

            # ------------------Single CR / RCR Document cannot have both positive and negative amounts-------------

            df["Temp"] = df["TaxableValue"].apply(lambda x: -1 if x < 0 else 1)
            LineSeries = df.groupby('DocumentNumber')['Temp'].sum()
            newdf = LineSeries.to_frame()
            newdf.reset_index(inplace=True)

            DocumentSeries = df.groupby('DocumentNumber')[
                'DocumentNumber'].count()
            newdf1 = DocumentSeries.to_frame()
            newdf1.columns = ["DocumentNumber1"]

            newdf1.reset_index(inplace=True)
            lst = ["CR", "RCR"]

            for i in range(len(newdf)):
                if abs(newdf["Temp"][i]) != newdf1["DocumentNumber1"][i]:

                    df.loc[(df["DocumentNumber"] == newdf["DocumentNumber"][i]) & (df["DocumentType"].isin(lst)) & (
                        df["SupplierGSTIN"] != "NA"), 'Flash Remark'] = df['Flash Remark'] + "| Single CR_with +ve and -ve values"
            df.drop('Temp', axis=1, inplace=True)
            # ---------2.17(ii).Remaining transactions-------------
            df["EY Remark"].fillna("NA", inplace=True)
            df.loc[df["EY Remark"] == "NA", "EY Remark"] = "Other Purchases"
            if option_list[0] == 0:
                df["DocumentNumber"] = df["Unaltered_Document_Number"]

            df.drop("EY Remark", inplace=True, axis=1)
            # df.drop("tax_check",inplace=True,axis=1)
            # df.drop("rate_check",inplace=True,axis=1)

            # --------------Deriving PAN----------------

            lst = df["RecipientGSTIN"].unique()
            PAN = ""
            for x in lst:
                if len(str(x)) == 15:
                    PAN = x[3:13]
                    break

            # ---------------Creating list of errors applied---------------

            lst1 = df["Flash Remark"].unique()
            lst_all = list(df["Flash Remark"])
            df_lst = pd.DataFrame()
            lst3 = set()
            try:
                for lst in lst1:
                    lst2 = (lst.split("|", maxsplit=50))
                    for lst in lst2:
                        lst3.add(lst)
            except:
                lst3.add("")

            lst4 = []
            for x in lst3:
                count = 0
                for item in lst_all:
                    if item.find(x) != -1:
                        count = count+1
                lst4.append(count)

            df_lst["Unique_Errors_list"] = pd.Series(list(lst3))
            df_lst["Count_of_errors"] = pd.Series(list(lst4))
            df_lst["Unique_Errors_list"].dropna(inplace=True)
            i = df_lst[df_lst["Unique_Errors_list"] == "NA"].index
            df_lst.drop(i, inplace=True)
            df.replace("NA", "", inplace=True)
            df.replace("NN", "NA", inplace=True)
            df["Flash Remark"] = df["Flash Remark"].apply(
                lambda x: str(x).replace("NA", ""))
            # print(df.columns)
            # print(processed.columns)
            df=processed.append(df,ignore_index=True)
            # print("Saving To Database...")
            
            # print(f'Seqrun1 shape : {df.shape}')
            df.to_sql("PurchaseRegisterFlash", panwisedb,
                      if_exists="append", index=False)
    try:
        cur = panwisedb.cursor()
        sql = "DROP TABLE PurchaseRegisterDigi"
        cur.execute(sql)
        panwisedb.commit()
        sql = "ALTER TABLE PurchaseRegisterFlash RENAME TO PurchaseRegisterDigi"
        cur.execute(sql)
        panwisedb.commit()
    except:
        traceback.print_exc()
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'

    c = panwisedb.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS Summary_Totals (Particulars TEXT Primary Key,Count INTEGER, Taxable_Value INTEGER,CGST INTEGER,SGST INTEGER,IGST INTEGER,Cess INTEGER)")
    query = "select * FROM Summary_Totals"
    
    current = pd.read_sql_query(query, panwisedb)
    current.drop(current[current['Particulars']=="PR - Flash"].index, inplace=True)
    query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigi"
    upd = pd.read_sql_query(query, panwisedb)
    
    upd['Particulars'] = "PR - Flash"
    upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
    neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
    upd=upd.reindex(columns=neworder)        
    upd=upd.fillna(0)
    upd = upd.append(current)
    upd.to_sql("Summary_Totals", panwisedb, if_exists="replace",index=False)
    query = "select * FROM Summary_Totals"
    
    current = pd.read_sql_query(query, panwisedb)
    current.drop(current[current['Particulars']=="PR - Flash Error"].index, inplace=True)
    query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigi"
    upd = pd.read_sql_query(query, panwisedb)
    
    upd['Particulars'] = "PR - Flash Error"
    upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
    neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
    upd=upd.reindex(columns=neworder)        
    upd=upd.fillna(0)
    upd = upd.append(current)
    
    upd.to_sql("Summary_Totals", panwisedb, if_exists="replace",index=False)

# validate pr
# @router.post('/valpr', response_class = Response)
def valpr(clientPAN, current_user):
    starttime = time.time()
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'

    panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)
    i = 1

    folder_name = 'Duplicates dropped'
    try:
        os.mkdir(dir_path +'\\'+ 'PR2A\\'+ folder_name)
    except:
        pass

    i=1
    
    sql = "DROP TABLE purchaseregister_temp"
    cur = panwisedb.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS purchaseregister_temp (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
    cur.execute(sql)
    panwisedb.commit()
    sql = "DROP TABLE PurchaseRegisterDigiConsotemp"
    cur = panwisedb.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS PurchaseRegisterDigiConsotemp (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
    cur.execute(sql)
    panwisedb.commit()
    cur.close()
    

    query = "select * FROM PurchaseRegisterDigi where `GSTR 9` ==''"
    df = pd.DataFrame()
    cdf2 = pd.DataFrame()
    cdf5 = pd.DataFrame()
    date2 = datetime(2099,3,31)
    date3 = datetime(2000,3,31)
                
    for cdf in pd.read_sql_query(query, panwisedb,chunksize=10000):
        # print('Reading from PurchaseRegisterDigi')
        # print(cdf.shape)
        # print(cdf.columns)
        processed = cdf[cdf['GSTR 9']!='']
        # print(processed)
        cdf.drop(cdf[cdf['GSTR 9']!=''].index, inplace=True)
        # print(cdf.shape)
    
        cdf['PurchaseVoucherDate'] = pd.to_datetime(
            cdf['PurchaseVoucherDate'], dayfirst=True)
        cdf['DocumentDate'] = pd.to_datetime(
            cdf['DocumentDate'], dayfirst=True)
        try:
            cdf['ReturnPeriod_temp'] = pd.to_datetime(cdf['ReturnPeriod'], dayfirst=True, format='%m%Y')
        except:
            cdf['ReturnPeriod_temp'] = pd.to_datetime(cdf['ReturnPeriod'], dayfirst=True, format='%Y%m')
        #cdf['Original Invoice Date']= pd.to_datetime(cdf['Original Invoice Date'],dayfirst=True)
        date = datetime(2018, 3, 31)
        date1 = datetime(2018, 9, 30)
        date4 = datetime(2019, 3, 31)
        #cdf['Key']= cdf['FI/Accounting Document Number'].map(str)+cdf['Billing Date'].map(str)+cdf['Bus Plc GSTIN'].map(str)
        #cdf['Key'] = cdf['Key'].str.replace(r'\.0','')
        cdf['TaxableValue'] = cdf['TaxableValue'].fillna(0.00)

        cdf['CentralTaxAmount'] = cdf['CentralTaxAmount'].fillna(0.00)
        cdf['StateUTTaxAmount'] = cdf['StateUTTaxAmount'].fillna(0.00)
        cdf['IntegratedTaxAmount'] = cdf['IntegratedTaxAmount'].fillna(0.00)
        cdf['CessAmountSpecific'] = cdf['CessAmountSpecific'].fillna(0.00)
        cdf['CentralTaxRate'] = cdf['CentralTaxRate'].fillna(0.00)
        cdf['StateUTTaxRate'] = cdf['StateUTTaxRate'].fillna(0.00)
        cdf['IntegratedTaxRate'] = cdf['IntegratedTaxRate'].fillna(0.00)
        cdf['CessRateSpecific'] = cdf['CessRateSpecific'].fillna(0.00)
        cdf['Tax_Amount_Reg'] = 0
        cdf['Tax_Rate_Reg'] = 0
        cdf['TaxableValue'] = cdf['TaxableValue'].astype(float).round(2)
        cdf['CentralTaxAmount'] = cdf['CentralTaxAmount'].astype(
            float).round(2)
        cdf['StateUTTaxAmount'] = cdf['StateUTTaxAmount'].astype(
            float).round(2)
        cdf['IntegratedTaxAmount'] = cdf['IntegratedTaxAmount'].astype(
            float).round(2)
        cdf['CessAmountSpecific'] = cdf['CessAmountSpecific'].astype(
            float).round(2)

        cdf['Tax_Rate_Reg'] = np.where((cdf['Tax_Rate_Reg'] == 0) | (cdf['Tax_Rate_Reg'].isna(
        )), cdf['CentralTaxRate'] + cdf['StateUTTaxRate'] + cdf['IntegratedTaxRate'], cdf['Tax_Rate_Reg'])
        cdf['Tax_Rate_Reg'] = pd.to_numeric(
            abs(cdf['Tax_Rate_Reg']), errors='coerce')

        cdf['Tax_Amount_Reg'] = np.where((cdf['Tax_Amount_Reg'] == 0) | (cdf['Tax_Amount_Reg'].isna(
        )), cdf['CentralTaxAmount']+cdf['StateUTTaxAmount']+cdf['IntegratedTaxAmount'], cdf['Tax_Amount_Reg'])

        cdf['Rate Check'] = np.where(
            abs(cdf['TaxableValue']*cdf['Tax_Rate_Reg']/100-cdf['Tax_Amount_Reg']) > 1, 0, 1)
        cdf['Rate Check'] = np.where((cdf['Tax_Rate_Reg'] == 0) | (cdf['Tax_Rate_Reg'] == 0.1) | (cdf['Tax_Rate_Reg'] == 3) | (
            cdf['Tax_Rate_Reg'] == 5) | (cdf['Tax_Rate_Reg'] == 12) | (cdf['Tax_Rate_Reg'] == 18) | (cdf['Tax_Rate_Reg'] == 28), cdf['Rate Check'], 0)

        cdf['Rate Check'] = cdf['Rate Check'].astype(int)

        cdf['POS Check'] = np.where(np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) == cdf['RecipientGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount'] == 0), 1, np.where(np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) == cdf['RecipientGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount']), 0, np.where(
            np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) != cdf['RecipientGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount'] == 0), 0, np.where(np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) != cdf['RecipientGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount']), 1, "NA"))))
        cdf['POS Check'] = np.where(np.logical_and(
            cdf['POS Check'] == "NA", cdf['Tax_Amount_Reg'] == 0), 1, cdf['POS Check'])
        cdf['Period'] = np.where(cdf['DocumentDate'] <= date, "2017-18",
                                 np.where(cdf['DocumentDate'] <= date4, "2018-19", "2019-20"))
        cdf['Filing Period'] = cdf['DocumentDate'].apply(
            lambda x: x.strftime('%m&Y') if x else "")
        cdf['FY'] = cdf['DocumentDate'].map(
            lambda x: x.year if x.month > 3 else x.year-1)

        # added by mayur
        # cdf['Late_reporting_Check'] = np.where(cdf['ReturnPeriod'].dt.year == cdf['FY'], 1, np.where(np.logical_and(cdf['ReturnPeriod'].dt.year != cdf['FY'], cdf['ReturnPeriod'] < pd.to_datetime('09-30-'+ (cdf['FY'] +1).map(str))), 1, 0))
        cdf['Late_reporting_Check'] = np.where(cdf['ReturnPeriod_temp'].dt.year == cdf['FY'], 1, np.where(np.logical_and(
        cdf['ReturnPeriod_temp'].dt.year != cdf['FY'], cdf['ReturnPeriod_temp'] < pd.to_datetime('09-30-' + (cdf['FY'] + 1).map(str))), 1, 0))
        del cdf['ReturnPeriod_temp']

        cdf['GSTIN Valid'] = cdf['SupplierGSTIN'].str.match(
            r'[0-9]{2}[a-zA-Z]{5}[0-9]{4}[a-zA-Z]{1}[1-9A-Za-z]{1}[Z]{1}[0-9a-zA-Z]{1}')
        #cdf['Credit Note'] = np.where(np.logical_and(np.logical_or(cdf['Doc Type']=="C",cdf['Doc Type']=="CM"),np.logical_and(cdf['Original Invoice Date']<=date,cdf['Billing Date']>date1)),0,1)
        #cdf['Credit Note'] = np.where(np.logical_or(cdf['Doc Type']=="C",cdf['Doc Type']=="CM"),np.where(cdf['Original Invoice No.'].isna(),0,cdf['Credit Note']),cdf['Credit Note'])

        #cdf['Export days condition'] = np.where((cdf['Shipping Bill Date']-cdf['Billing Date'])>timedelta(days=90),0,1)

        cdf['HSN'] = np.where(cdf['HSNorSAC'].isna(), 0, 1)
        cdf['DocumentType1'] = "b2b"
        cdf['DocumentType1'] = np.where(cdf['DocumentType'].map(
            str) == "CRN", "cdn", cdf['DocumentType1'])
        cdf['DocumentType1'] = np.where(cdf['DocumentType'].map(
            str) == "CR", "cdn", cdf['DocumentType1'])
        cdf['DocumentType1'] = np.where(cdf['DocumentType'].map(
            str) == "CREDIT NOTE", "cdn", cdf['DocumentType1'])
        cdf['DocumentType1'] = np.where(cdf['DocumentType'].map(
            str) == "DBN", "dbn", cdf['DocumentType1'])
        cdf['DocumentType1'] = np.where(cdf['DocumentType'].map(
            str) == "DN", "dbn", cdf['DocumentType1'])
        cdf['DocumentType1'] = np.where(cdf['DocumentType'].map(
            str) == "DEBIT NOTE", "dbn", cdf['DocumentType1'])

        cdf['TaxableValue'] = np.where(np.logical_and(
            cdf['DocumentType1'] == "cdn", cdf['TaxableValue'] > 0), cdf['TaxableValue']*-1, cdf['TaxableValue'])
        cdf['CentralTaxAmount'] = np.where(np.logical_and(
            cdf['DocumentType1'] == "cdn", cdf['CentralTaxAmount'] > 0), cdf['CentralTaxAmount']*-1, cdf['CentralTaxAmount'])
        cdf['StateUTTaxAmount'] = np.where(np.logical_and(
            cdf['DocumentType1'] == "cdn", cdf['StateUTTaxAmount'] > 0), cdf['StateUTTaxAmount']*-1, cdf['StateUTTaxAmount'])
        cdf['IntegratedTaxAmount'] = np.where(np.logical_and(
            cdf['DocumentType1'] == "cdn", cdf['IntegratedTaxAmount'] > 0), cdf['IntegratedTaxAmount']*-1, cdf['IntegratedTaxAmount'])
        cdf['CessAmountSpecific'] = np.where(np.logical_and(
            cdf['DocumentType1'] == "cdn", cdf['CessAmountSpecific'] > 0), cdf['CessAmountSpecific']*-1, cdf['CessAmountSpecific'])

        cdf['consokey'] = cdf['DocumentNumber']+cdf['DocumentType1'] + cdf['SupplierGSTIN']+cdf['RecipientGSTIN']+cdf['FY'].map(str)
        # print(cdf.shape)
        cdf1 = cdf.pivot_table(index=['consokey'], values=['TaxableValue', 'CentralTaxAmount', 'StateUTTaxAmount','IntegratedTaxAmount', 'CessAmountSpecific', 'Quantity', 'Tax_Amount_Reg'], aggfunc=sum, dropna=False)
        # print(cdf1.shape)
        cdf1.to_csv("consokeypre.csv")
        # cdf1.to_csv("C:\\Upload documents\\pr2b timer test abbott\\consokeypre_cdf1.csv", if_exists='replace')
        
        cdf2 = cdf2.append(cdf1)
        # print(cdf2.shape)
        cdf2 = cdf2.groupby(['consokey']).agg('sum')
        # print(cdf2.shape)
        # cdf2.to_csv("consokey.csv")
        # cdf2.to_csv("C:\\Upload documents\\pr2b timer test abbott\\consokey_cdf2.csv", if_exists='replace')
        cdf['GSTR 9'] = "Unknown"
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "TAX", cdf['ReverseChargeFlag'] != "Y"), "Table 6B", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(cdf['ReverseChargeFlag'] == "Y", (
            cdf['SupplierGSTIN'] == "") | (cdf['SupplierGSTIN'] == 0)), "Table 6C", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(cdf['ReverseChargeFlag'] == "Y", (
            cdf['SupplierGSTIN'] != "") | (cdf['SupplierGSTIN'] != 0)), "Table 6D", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "IMPG", "Table 6E", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "IMPS", "Table 6F", cdf['GSTR 9'])

    #              cdf['GSTR 9'] = np.where(np.logical_and(cdf['Doc Type']== "E", cdf['Tax_Amount_Reg']),"Table 4C",cdf['GSTR 9'])
    #             cdf['GSTR 9'] = np.where(np.logical_and(cdf['Doc Type']== "E", cdf['Tax_Amount_Reg']==0),"Table 5A",cdf['GSTR 9'])

    #            cdf['GSTR 9'] = np.where(np.logical_and(cdf['Doc Type'] == "CM", cdf['Tax_Amount_Reg']<0),"Table 4I",cdf['GSTR 9'])
    #           cdf['GSTR 9'] = np.where(np.logical_and(cdf['Doc Type'] == "DM", cdf['Tax_Amount_Reg']>0),"Table 4J",cdf['GSTR 9'])
    #          cdf['GSTR 9'] = np.where(np.logical_and(cdf['Doc Type'] == "C", cdf['Tax_Amount_Reg']<0),"Table 4I",cdf['GSTR 9'])
    #         cdf['GSTR 9'] = np.where(np.logical_and(cdf['Doc Type'] == "C", cdf['Tax_Amount_Reg']>0),"Table 4J",cdf['GSTR 9'])
    #        cdf['GSTR 9'] = np.where(np.logical_and(cdf['GSTIN/UIN']=="",cdf['Doc Type']== "I"),"Table 4A",cdf['GSTR 9'])
    #       cdf['GSTR 9'] = np.where(np.logical_and(cdf['Taxable Amount']==0,cdf['Qty']==0),"Ignore - Taxable Value 0",cdf['GSTR 9'])
    #      cdf['GSTR 9'] = np.where(np.logical_and(cdf['Taxable Amount']==0,cdf['Qty']!=0),np.where(np.logical_or(cdf['Billing Type']=="ZS53",cdf['Billing Type']=="ZS34"),"Sample Sales","Ignore - Taxable Value 0"),cdf['GSTR 9'])
    #     cdf['GSTR 9'] = np.where(cdf['GSTIN/UIN']==cdf['Bus Plc GSTIN'],"Ignore - Intrastate Stock Transfer",cdf['GSTR 9'])
        # print(cdf)
        # print(cdf.columns)
        
        # print(processed.columns)
        cdf5 = cdf5.append(cdf,ignore_index=True)
        # print(cdf5)
        cdf.to_sql("purchaseregister_temp", panwisedb,if_exists="append", index=False)
        

    query = "select * FROM PurchaseRegisterDigi where `GSTR 9` !=''"
    oldrows = pd.read_sql_query(query, panwisedb)
    print(oldrows.columns)
    print(cdf5.columns)
    if len(cdf5)>0:
        cdf5.to_sql("PurchaseRegisterDigi", panwisedb,if_exists="replace", index=False)
        oldrows.to_sql("PurchaseRegisterDigi", panwisedb,if_exists="append", index=False)
    
    else :
        oldrows.to_sql("PurchaseRegisterDigi", panwisedb,if_exists="replace", index=False)
        

    try:
        query = "select * FROM matchconfig"
        config = pd.read_sql_query(query, panwisedb)
        rowdata = dataframe_to_rows(config, index=False, header=False)
        rowdata1 = dataframe_to_rows(config, index=False, header=False)
        rowdata2 = dataframe_to_rows(config, index=False, header=False)
        matchcon = 1
    except:
        matchcon = 0
        config = []
    
    try:
        query = "select * FROM matchconfig_2b"
        config_2b = pd.read_sql_query(query, panwisedb)
        rowdata_2b = dataframe_to_rows(config_2b, index=False, header=False)
        matchcon_2b = 1
    except:
        matchcon_2b = 0
        config_2b = []
    if len(cdf5)>0:
        query = "SELECT * FROM purchaseregister_temp GROUP BY consokey ORDER BY DocumentNumber ASC"
        j = 1
        for df in pd.read_sql(query, panwisedb, chunksize=10000):
            # print(df.columns)
            # print(df)
            processed1 = df[df['Key1']!=""]
            # print(processed1)
            df.drop(df[df['Key1']!=""].index, inplace=True)
            
            df.drop(columns=['TaxableValue', 'CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount',
                    'CessAmountSpecific', 'Quantity', 'Tax_Amount_Reg'], axis=1, inplace=True)
            df = df[(~df['consokey'].duplicated()) | df['consokey'].isna()]
            # df.to_csv("left.csv")
            df = pd.merge(df, cdf2, on='consokey', how='left')
            # df.to_csv("merged.csv")
            df['TaxableValue'] = df['TaxableValue'].round(2)
            df['DocumentNumber'] = df['DocumentNumber'].str.lstrip('0')

            # print("Read User Defined Keys")

            duplicate = df[df.duplicated()]
            df.drop_duplicates(keep='first', inplace=True)
            duplicate['Reason'] = "Duplicate"
            duplicate1 = df.drop(df[~df['DocumentNumber'].isnull()].index)
            duplicate1['Reason'] = "Doc No. Blank"
            df.drop(df[df['DocumentNumber'].isnull()].index, inplace=True)
            duplicate2 = df.drop(df[~df['DocumentNumber'].isna()].index)
            duplicate2['Reason'] = "Doc No. Blank"

            df.drop(df[df['DocumentNumber'].isna()].index, inplace=True)
            droppedpr = duplicate.append(duplicate1).append(duplicate2)
            save_duplicate_files(droppedpr, 'PR_validated - Duplicate Drops', current_user)

            df['DocumentNumber'] = df['DocumentNumber'].astype(
                str, copy=True, errors='ignore')
            # cdf['DocumentNumber']=cdf['DocumentNumber'].str.encode('utf-8')
            df['DocumentNumber'] = df['DocumentNumber'].map(str).str.upper()

            df.reset_index(inplace=True)

            # # print(df)
            # cdf.to_csv(r"C:\GAPS\Reports\Dump1.csv",index=False)
            df['first1'] = df['DocumentNumber'].apply(lambda x: re.search(
                r'[\\/*?:."<(-)>|]', x).start() if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
            df['last1'] = df['DocumentNumber'].apply(lambda x: re.search(
                r'[\\/*?:."<(-)>|]', x[::-1]).start() if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
            #df['last1']= len(df['DocumentNo']) - df['last1']
            # # print(df['first1'])
            df['first1'] = df['first1'].fillna(
                df['DocumentNumber'].str.len()).astype(int)
            df['last1'] = df['last1'].fillna(
                df['DocumentNumber'].str.len()).astype(int)
            df['DocNoBfrSpl'] = [DocumentNumber[:first1]
                                for DocumentNumber, first1 in zip(df.DocumentNumber, df.first1)]
            df['DocNoAftrSpl'] = [DocumentNumber[-last1:]
                                for DocumentNumber, last1 in zip(df.DocumentNumber, df.last1)]
            df['DocNoWOSplChar'] = df['DocumentNumber'].replace(
                '[^a-zA-Z0-9]', '', regex=True)
            df['DocNoNumeric'] = df['DocumentNumber'].str.replace(
                r"[a-zA-Z]", '').replace('\W', '').replace('_', '').replace('/', '').replace("\\", '')
            # df['DocNoBfrSpl'] =
            df['DocNo'] = df['DocumentNumber']
            # df['DocNoAftrSpl'] =
            df['SuppGSTIN'] = df['SupplierGSTIN']
            df['SuppPAN'] = df['SupplierGSTIN'].str[2:12].map(str)
            df['ResGSTIN'] = df['RecipientGSTIN']
            df['ResPAN'] = df['RecipientGSTIN'].str[2:12].map(str)
            df['DocDate'] = pd.to_datetime(df['DocumentDate'], dayfirst=True).apply(
                lambda x: x.strftime('%d%m%Y') if x else "")
            df['MMYYYY'] = pd.to_datetime(df['DocumentDate'], dayfirst=True).apply(
                lambda x: x.strftime('%m%Y') if x else "")
            df['FY'] = pd.to_datetime(df['DocumentDate'],dayfirst=True).apply(lambda x: x.strftime('%Y') if x else "")
            df['TaxVal'] = df['TaxableValue'].round(2)
            df['InvVal'] = df['InvoiceValue'].map(float).round(2)
            df['CGST'] = df['CentralTaxAmount'].round(2)
            df['SGST'] = df['StateUTTaxAmount'].round(2)
            df['IGST'] = df['IntegratedTaxAmount'].round(2)
            df['CessAmountSpecific'] = df['CessAmountSpecific'].round(2)

            df['GST'] = df['CentralTaxAmount'].round(2) + df['StateUTTaxAmount'].round(
                2) + df['IntegratedTaxAmount'].round(2) + df['CessAmountSpecific'].round(2)
            df['POS'] = df['POS'].apply(pd.to_numeric)
            df['POS'] = df['POS'].map(str)
            df['POS'] == df["POS"].replace('.0', '')

            df['RCM'] = df['ReverseChargeFlag']
            df['DocType'] = df['DocumentType1']
            df['NA'] = ""

            df['Key1'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNumber'].map(str)+df['RecipientGSTIN'].map(str)+df['DocDate'].map(str)+df['TaxableValue'].round(2).map(str)+df['CentralTaxAmount'].round(
                2).map(str)+df['StateUTTaxAmount'].round(2).map(str)+df['IntegratedTaxAmount'].round(2).map(str)+df['CessAmountSpecific'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
            df['Key2'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(str)+df['RecipientGSTIN'].map(str)+df['DocDate'].map(str)+df['TaxableValue'].round(2).map(str)+df['CentralTaxAmount'].round(
                2).map(str)+df['StateUTTaxAmount'].round(2).map(str)+df['IntegratedTaxAmount'].round(2).map(str)+df['CessAmountSpecific'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
            df['Key3'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNumber'].map(
                str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key4'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key5'] = df['SupplierGSTIN'].map(
                str)+df['DocumentNumber'].map(str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key6'] = df['SupplierGSTIN'].map(
                str)+df['DocNoWOSplChar'].map(str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key7'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['RecipientGSTIN'].map(str)+df['DocDate'].map(str)+df['TaxableValue'].round(2).map(str)+df['CentralTaxAmount'].round(
                2).map(str)+df['StateUTTaxAmount'].round(2).map(str)+df['IntegratedTaxAmount'].round(2).map(str)+df['CessAmountSpecific'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
            df['Key8'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
                str)+df['TaxableValue'].round(2).map(str)+df['Tax_Amount_Reg'].round(2).map(str)+df['DocType'].map(str)
            df['Key9'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
                str)+df['TaxableValue'].round(2).map(str)+df['Tax_Amount_Reg'].round(2).map(str)+df['DocType'].map(str)
            df['Key10'] = df['FY'].map(str) + df['SuppPAN'].map(str) + df['DocNoWOSplChar'].map(
                str)+df['ResPAN'].map(str)+df['Tax_Amount_Reg'].round(2).map(str)+df['DocType'].map(str)

            df['Key11'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNumber'].map(
                str)+df['RecipientGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
            df['Key12'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                str)+df['RecipientGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
            df['Key13'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNumber'].map(
                str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key14'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key15'] = df['SupplierGSTIN'].map(
                str)+df['DocumentNumber'].map(str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key16'] = df['SupplierGSTIN'].map(
                str)+df['DocNoWOSplChar'].map(str)+df['RecipientGSTIN'].map(str)+df['DocType'].map(str)
            df['Key19'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(
                str)+df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)
            df['Key20'] = df['FY'].map(str) + df['SuppPAN'].map(
                str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocType'].map(str)

            df['Key17'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['RecipientGSTIN'].map(
                str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
            df['Key18'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(
                str)+df['DocDate'].map(str)+df['DocType'].map(str)

            # print("Saved PR Default Keys")
            l = len(config)
            n = 21
            try:
                for row in rowdata:
                    key = "Key" + str(n) + "_2a"
                    df[key] = df[row[1]].map(str) + df[row[2]].map(str) + df[row[3]].map(str) + df[row[4]].map(
                        str) + df[row[5]].map(str) + df[row[6]].map(str) + df[row[7]].map(str) + df[row[8]].map(str)
                    n += 1
                    ## print("prkey ",row[0])
            except:
                pass

            n = 21
            try:
                for row in rowdata_2b:
                    key = "Key" + str(n) + "_2b"
                    df[key] = df[row[1]].map(str) + df[row[2]].map(str) + df[row[3]].map(str) + df[row[4]].map(
                        str) + df[row[5]].map(str) + df[row[6]].map(str) + df[row[7]].map(str) + df[row[8]].map(str)
                    n += 1
                    ## print("prkey ",row[0])
            except:
                pass

            # print("Saved User Defined Keys")
            # print(df.columns)
            # print(processed1.columns)
            # df1=processed1.append(df,ignore_index=True)
            df.to_sql("PurchaseRegisterDigiConso",panwisedb, if_exists="append", index=False)
            #df.to_csv("GAPS\Purchase Register - Full Conso Processed.csv",index=False, header=j, mode='a')
            j = 0
    
    try:
        cur = panwisedb.cursor()
        
        sql = "DROP TABLE purchaseregister_temp"
        try:
            cur.execute(sql)
            panwisedb.commit()
            # print('Dropped PurchaseRegisterDigi')
        except:
            traceback.print_exc()

    except:
        traceback.print_exc()

    c = panwisedb.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS Summary_Totals (Particulars TEXT Primary Key,Count INTEGER, Taxable_Value INTEGER,CGST INTEGER,SGST INTEGER,IGST INTEGER,Cess INTEGER)")
    query = "select * FROM Summary_Totals"
    
    current = pd.read_sql_query(query, panwisedb)
    current.drop(current[current['Particulars']=="Purchase Register - Conso"].index, inplace=True)
    query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigiConso"
    upd = pd.read_sql_query(query, panwisedb)
    
    upd['Particulars'] = "Purchase Register - Conso"
    upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
    neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
    upd=upd.reindex(columns=neworder)
    upd=upd.fillna(0)
    upd = upd.append(current)
    
    upd.to_sql("Summary_Totals", panwisedb, if_exists="replace",index=False)
    c = panwisedb.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS Summary_Totals (Particulars TEXT Primary Key,Count INTEGER, Taxable_Value INTEGER,CGST INTEGER,SGST INTEGER,IGST INTEGER,Cess INTEGER)")
    query = "select * FROM Summary_Totals"
    
    current = pd.read_sql_query(query, panwisedb)
    current.drop(current[current['Particulars']=="Purchase Register - Processed"].index, inplace=True)
    query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM PurchaseRegisterDigi"
    upd = pd.read_sql_query(query, panwisedb)
    
    upd['Particulars'] = "Purchase Register - Processed"
    upd=upd.rename(columns={'SUM(TaxableValue)':'Taxable_Value','SUM(CentralTaxAmount)':'CGST','SUM(StateUTTaxAmount)':'SGST','SUM(IntegratedTaxAmount)':'IGST','COUNT(SupplierGSTIN)':'Count','COUNT(CessAmountAdvalorem)':'Cess'})
    neworder = ['Particulars','Count','Taxable_Value','CGST','SGST','IGST','Cess']
    upd=upd.reindex(columns=neworder)        
    upd=upd.fillna(0)
    upd = upd.append(current)
    
    upd.to_sql("Summary_Totals", panwisedb, if_exists="replace",index=False)

# validate sr
# @router.post('/valsr', response_class = Response)
def valsr(is_local, parameters):
    # is_local true if localhost else false if Azure Blob Storage
    # parameters will be clientPAN and current_user if localhost else temp_db_name
    dir_path = os.path.dirname(os.path.realpath(__file__))
    folder_name = '' if is_local else parameters[1]
    # path = dir_path + '/Client-Details' if is_local else dir_path + folder_name
    path = os.path.join(dir_path, 'Client-Details') if is_local else os.path.join(dir_path, folder_name)

    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]

    # panwisedb_path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'
    panwisedb_path = os.path.join(path, current_user, clientPAN, f'{clientPAN}.db') if is_local else os.path.join(path, f'{temp_db_name}.db')

    panwisedb = sqlite3.connect(panwisedb_path, timeout=10)
    i = 1

    sql = "DROP TABLE Sales_Register_temp"
    cur = panwisedb.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Sales_Register_temp (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
    cur.execute(sql)
    panwisedb.commit()
    cur.close()

    # query = "select * FROM SalesRegisterFlash" # commented for skipping the seqRun function
    query = "select * FROM SalesRegisterDigi"
    df = pd.DataFrame()
    cdf2 = pd.DataFrame()
    date2 = datetime(2099, 3, 31)
    date3 = datetime(2000, 3, 31)

    for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000):

        cdf.drop(cdf[cdf['SupplierGSTIN'].isna()].index, inplace=True)
        cdf['UnitOfMeasurement'] = np.where(
            cdf['UnitOfMeasurement'].isna(), "NotAvailable", cdf['UnitOfMeasurement'])
        cdf['OriginalDocumentDate'] = np.where(
            cdf['OriginalDocumentDate'].isna(), date2, cdf['OriginalDocumentDate'])
        cdf['ShippingBillDate'] = np.where(
            cdf['ShippingBillDate'].isna(), date3, cdf['ShippingBillDate'])
        cdf['OriginalDocumentDate'] = np.where(
            cdf['OriginalDocumentDate'].isnull(), date2, cdf['OriginalDocumentDate'])
        cdf['ShippingBillDate'] = np.where(
            cdf['ShippingBillDate'].isnull(), date3, cdf['ShippingBillDate'])

        try:
            cdf['OriginalDocumentDate'] = cdf['OriginalDocumentDate'].str.replace(
                r'00.00.0000', '31.03.2099')
            cdf['ShippingBillDate'] = cdf['ShippingBillDate'].str.replace(
                r'00.00.0000', '31.03.2000')
        except:
            pass
        finally:
            pass
        # cdf.to_csv(r"C:\CAM\Reports\SRtemp.csv")
        cdf['DocumentDate'] = pd.to_datetime(
            cdf['DocumentDate'], dayfirst=True)
        cdf['ShippingBillDate'] = pd.to_datetime(
            cdf['ShippingBillDate'], dayfirst=True)

        cdf['OriginalDocumentDate'] = pd.to_datetime(
            cdf['OriginalDocumentDate'], dayfirst=True)
        date = datetime(2018, 3, 31)
        date1 = datetime(2018, 9, 30)
        date4 = datetime(2019, 3, 31)

        cdf = cdf.round(2)

        # # print(cdf)

        cdf['TaxableValue'] = cdf['TaxableValue'].astype(float).round(2)
        cdf['StateUTTaxAmount'] = cdf['StateUTTaxAmount'].astype(float).round(2)
        cdf['CentralTaxAmount'] = cdf['CentralTaxAmount'].astype(float).round(2)
        cdf['IntegratedTaxAmount'] = cdf['IntegratedTaxAmount'].astype(float).round(2)
        cdf['CessAmountSpecific'] = cdf['CessAmountSpecific'].astype(float).round(2)
        
        # cdf[['CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount', 'CessAmountSpecific']] = cdf[['CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount', 'CessAmountSpecific']].applymap(lambda x: round(float(x.replace(',','')),2) if isinstance(x, str) else round(float(x), 2))
        cdf[['CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount', 'CessAmountSpecific']] = cdf[['CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount', 'CessAmountSpecific']].applymap(
        lambda x: round(float(x.replace(',', '')), 2) if isinstance(x, str) and x else (round(float(x), 2)))
        
        cdf['Tax_Amount_Reg'] = 0

        cdf['Tax_Rate_Reg'] = 0
        cdf['CentralTaxRate'] = cdf['CentralTaxRate'].astype(float).round(2)
        cdf['StateUTTaxRate'] = cdf['StateUTTaxRate'].astype(float).round(2)
        cdf['IntegratedTaxRate'] = cdf['IntegratedTaxRate'].astype(float).round(2)

        cdf['TaxableValue'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['TaxableValue'] > 0), cdf['TaxableValue']*-1, cdf['TaxableValue'])
        cdf['CentralTaxAmount'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['CentralTaxAmount'] > 0), cdf['CentralTaxAmount']*-1, cdf['CentralTaxAmount'])
        cdf['StateUTTaxAmount'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['StateUTTaxAmount'] > 0), cdf['StateUTTaxAmount']*-1, cdf['StateUTTaxAmount'])
        cdf['IntegratedTaxAmount'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['IntegratedTaxAmount'] > 0), cdf['IntegratedTaxAmount']*-1, cdf['IntegratedTaxAmount'])
        cdf['CessAmountSpecific'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['CessAmountSpecific'] > 0), cdf['CessAmountSpecific']*-1, cdf['CessAmountSpecific'])

        cdf['Tax_Rate_Reg'] = np.where((cdf['Tax_Rate_Reg'] == 0) | (cdf['Tax_Rate_Reg'].isna(
        )), cdf['CentralTaxRate'] + cdf['StateUTTaxRate'] + cdf['IntegratedTaxRate'], cdf['Tax_Rate_Reg'])
        cdf['Tax_Rate_Reg'] = pd.to_numeric(
            abs(cdf['Tax_Rate_Reg']), errors='coerce')

        #cdf['Tax Amount'] = cdf['Tax Amount'].fillna(0)
        cdf['Tax_Amount_Reg'] = np.where((cdf['Tax_Amount_Reg'] == 0) | (cdf['Tax_Amount_Reg'].isna(
        )), cdf['CentralTaxAmount']+cdf['StateUTTaxAmount']+cdf['IntegratedTaxAmount'], cdf['Tax_Amount_Reg'])
        cdf['Rate Check'] = np.where(
            abs(cdf['TaxableValue']*cdf['Tax_Rate_Reg']/100-cdf['Tax_Amount_Reg']) > 1, 0, 1)
        cdf['Rate Check'] = np.where((cdf['Tax_Rate_Reg'] == 0) | (cdf['Tax_Rate_Reg'] == 0.1) | (cdf['Tax_Rate_Reg'] == 3) | (
            cdf['Tax_Rate_Reg'] == 5) | (cdf['Tax_Rate_Reg'] == 12) | (cdf['Tax_Rate_Reg'] == 18) | (cdf['Tax_Rate_Reg'] == 28), cdf['Rate Check'], 0)
        cdf['Rate Check'] = cdf['Rate Check'].astype(int)
        cdf['POS Check'] = np.where(np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) == cdf['CustomerGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount'] == 0), 1, np.where(np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) == cdf['CustomerGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount']), 0, np.where(
            np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) != cdf['CustomerGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount'] == 0), 0, np.where(np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) != cdf['CustomerGSTIN'].str[0:2].map(str), cdf['IntegratedTaxAmount']), 1, "NA"))))
        cdf['POS Check'] = np.where(np.logical_and(
            cdf['POS Check'] == "NA", cdf['Tax_Amount_Reg'] == 0), 1, cdf['POS Check'])
        cdf['POS Check'] = np.where(
            cdf['CustomerGSTIN'].isna(), "NA", cdf['POS Check'])
        cdf['Period'] = np.where(cdf['DocumentDate'] <= date, "2017-18",
                                 np.where(cdf['DocumentDate'] <= date4, "2018-19", "2019-20"))
        cdf['Filing Period'] = cdf['DocumentDate'].apply(
            lambda x: x.strftime('%m%Y') if x else "")
        cdf['GSTIN Valid'] = cdf['CustomerGSTIN'].str.match(
            r'[0-9]{2}[a-zA-Z]{5}[0-9]{4}[a-zA-Z]{1}[1-9A-Za-z]{1}[Z]{1}[0-9a-zA-Z]{1}')
        cdf['Credit Note'] = np.where(np.logical_and(cdf['DocumentType'] == "CN", np.logical_and(
            cdf['OriginalDocumentDate'] <= date, cdf['DocumentDate'] > date1)), 0, 1)
        cdf['Export days condition'] = np.where(
            (cdf['ShippingBillDate']-cdf['DocumentDate']) > timedelta(days=90), 0, 1)
        cdf['HSN'] = np.where(cdf['HSNorSAC'].isna(), 0, 1)
        cdf1 = cdf.pivot_table(index=['DocumentNumber'], values=['TaxableValue', 'CentralTaxAmount', 'StateUTTaxAmount',
                               'IntegratedTaxAmount', 'CessAmountSpecific', 'Quantity', 'Tax_Amount_Reg'], aggfunc=sum, dropna=False)
        cdf2 = cdf2.append(cdf1)
        cdf2 = cdf2.groupby(['DocumentNumber']).agg('sum')

        cdf['GSTR 9'] = "Unknown"

        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "TAX", cdf['Tax_Amount_Reg'] > 0), "Table 4B", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "TAX", cdf['Tax_Amount_Reg'] < 0), "Table 4I", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(cdf['SupplyType'] == "TAX", np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0")), "Table 4A", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "EXP", cdf['Tax_Amount_Reg'] == 0), "Table 5A", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "EXP", cdf['Tax_Amount_Reg'] > 0), "Table 4C", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "EXP", cdf['Tax_Amount_Reg'] < 0), "Table 4I", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(cdf['SupplyType'] == "EXP", np.logical_and(
            cdf['Tax_Amount_Reg'] == 0, cdf['TaxableValue'] < 0)), "Table 5H", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "SEZ", cdf['Tax_Amount_Reg'] > 0), "Table 4D", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "SEZ", cdf['Tax_Amount_Reg'] < 0), "Table 4I", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['SupplyType'] == "SEZ", cdf['Tax_Amount_Reg'] == 0), "Table 5B", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "NIL", "Table 5E", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "ADV", "Table 4F", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "RCM", "Table 4G", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['Tax_Amount_Reg'] == 0), "Table 5H", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['DocumentType'] == "CN", cdf['Tax_Amount_Reg'] < 0), "Table 4I", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(cdf['DocumentType'] == "CN", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == 0), cdf['SupplyType'] == "TAX")), "Table 4A", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['DocumentType'] == "DN", cdf['Tax_Amount_Reg'] == 0), "Table 5I", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(
            cdf['DocumentType'] == "DN", cdf['Tax_Amount_Reg'] > 0), "Table 4J", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(np.logical_and(cdf['DocumentType'] == "DN", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == 0), cdf['SupplyType'] == "TAX")), "Table 4A", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(np.logical_and(cdf['TaxableValue'] == 0, cdf['Quantity']
                                 == 0), "Ignore - Taxable Value and Quantity 0", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "Ignore", "Ignore", cdf['GSTR 9'])
        cdf['GSTR 9'] = np.where(
            cdf['SupplyType'] == "ORC", "Table 5C", cdf['GSTR 9'])

        cdf['GSTR 9'] = np.where(cdf['CustomerGSTIN'] == cdf['SupplierGSTIN'],
                                 "Ignore - Intrastate Stock Transfer", cdf['GSTR 9'])

        # GSTR1 LOGICS START

        cdf['Category'] = "Unknown"
        cdf['Sub_Category'] = "Unknown"

        # 1,2,3
        cdf['Category'] = np.where((cdf['DocumentType'] == "INV") & (cdf['SupplyType'] == "TAX") | (
            cdf["SupplyType"] == "DTA") & (cdf['CustomerGSTIN']), "B2B", cdf['Category'])

        cdf['Sub_Category'] = np.where((cdf['Category'] == "B2B") & (
            cdf['ReverseChargeFlag'] == "N") & (cdf['TCSFlag'] == "N"), "4A", cdf['Sub_Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "B2B") & (
            cdf['ReverseChargeFlag'] == "Y"), "4B", cdf['Sub_Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "B2B") & (
            cdf['TCSFlag'] == "Y"), "4C", cdf['Sub_Category'])

        # 4
        cdf['Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "TAX", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) != cdf['POS'], np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000))))), "B2CL", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "TAX", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['SupplierGSTIN'].str[0:2].map(str) != cdf['POS'], np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000))))), "5A", cdf['Sub_Category'])

        # 4.2
        cdf['Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "DTA", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000)))), "B2CL", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "DTA", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000)))), "5A", cdf['Sub_Category'])

        # 5
        cdf['Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "TAX", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['TCSFlag'] == "Y", np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000))))), "B2CL", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "TAX", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['TCSFlag'] == "Y", np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000))))), "5B", cdf['Sub_Category'])

        # 5.2
        cdf['Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "DTA", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['TCSFlag'] == "Y", np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000))))), "B2CL", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(cdf['SupplyType'] == "DTA", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == 0), np.logical_and(cdf['TCSFlag'] == "Y", np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['InvoiceValue'] > 250000))))), "5B", cdf['Sub_Category'])

        # 6
        cdf['Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_or(
            cdf['SupplyType'] == "EXPT", cdf['SupplyType'] == "EXPWT")), "EXP", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_or(
            cdf['SupplyType'] == "EXPT", cdf['SupplyType'] == "EXPWT")), "6A", cdf['Sub_Category'])

        # 7
        cdf['Category'] = np.where(np.logical_and(
            cdf['DocumentType'] == "INV", cdf['SupplyType'] == "SEZ"), "EXP", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(
            cdf['DocumentType'] == "INV", cdf['SupplyType'] == "SEZ"), "6B", cdf['Sub_Category'])

        # 8
        cdf['Category'] = np.where(np.logical_and(
            cdf['DocumentType'] == "INV", cdf['SupplyType'] == "DXP"), "EXP", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(
            cdf['DocumentType'] == "INV", cdf['SupplyType'] == "DXP"), "6C", cdf['Sub_Category'])

        # 9
        cdf['Category'] = np.where(np.logical_and((cdf['DocumentType'] == "INV") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['CentralTaxAmount'] > 0, np.logical_or(cdf['StateUTTaxAmount'] > 0, cdf['IntegratedTaxAmount'] > 0))))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf['DocumentType'] == "INV") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['CentralTaxAmount'] > 0, np.logical_or(cdf['StateUTTaxAmount'] > 0, cdf['IntegratedTaxAmount'] > 0))))), "7A(1)", cdf['Sub_Category'])

        # 10
        cdf['Category'] = np.where(np.logical_and((cdf['DocumentType'] == "INV") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['CentralTaxAmount'] > 0, np.logical_and(np.logical_or(cdf['StateUTTaxAmount'] > 0, cdf['IntegratedTaxAmount'] > 0), cdf['TCSFlag'] == "Y"))))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf['DocumentType'] == "INV") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['CentralTaxAmount'] > 0, np.logical_and(np.logical_or(cdf['StateUTTaxAmount'] > 0, cdf['IntegratedTaxAmount'] > 0), cdf['TCSFlag'] == "Y"))))), "7A(2)", cdf['Sub_Category'])

        # 11
        cdf['Category'] = np.where(np.logical_and((cdf['DocumentType'] == "INV") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['CentralTaxAmount'] > 0, np.logical_or(cdf['StateUTTaxAmount'] > 0, cdf['IntegratedTaxAmount'] > 0))))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf['DocumentType'] == "INV") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['CentralTaxAmount'] > 0, np.logical_or(cdf['StateUTTaxAmount'] > 0, cdf['IntegratedTaxAmount'] > 0))))), "7B(1)", cdf['Sub_Category'])

        # 11.2
        cdf['Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(
            np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, cdf['IntegratedTaxAmount'] > 0)))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, cdf['IntegratedTaxAmount'] > 0)))), "7B(1)", cdf['Sub_Category'])

        # 11.3
        cdf['Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "DTA", np.logical_and(
            np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, cdf['IntegratedTaxAmount'] > 0)))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "DTA", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, cdf['IntegratedTaxAmount'] > 0)))), "7B(1)", cdf['Sub_Category'])

        # 12
        cdf['Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(np.logical_or(cdf["SupplyType"] == "TAX", cdf["SupplyType"] == "DTA"), np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['TCSFlag'] == "Y"))))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and(cdf['DocumentType'] == "INV", np.logical_and(np.logical_or(cdf["SupplyType"] == "TAX", cdf["SupplyType"] == "DTA"), np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['TCSFlag'] == "Y"))))), "7B(2)", cdf['Sub_Category'])

        # 12.2
        cdf['Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['TCSFlag'] == "Y"))))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "TAX", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['TCSFlag'] == "Y"))))), "7B(2)", cdf['Sub_Category'])

        # 12.3
        cdf['Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "DTA", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['TCSFlag'] == "Y"))))), "B2CS", cdf['Category'])
        cdf['Sub_Category'] = np.where(np.logical_and((cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "CR"), np.logical_and(cdf["SupplyType"] == "DTA", np.logical_and(np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), np.logical_and(cdf['InvoiceValue'] <= 250000, np.logical_and(cdf['IntegratedTaxAmount'] > 0, cdf['TCSFlag'] == "Y"))))), "7B(2)", cdf['Sub_Category'])

        # 13, 14, 15, 16
        # NIL, 8A, 8B, 8C, 8D
        cdf['Category'] = np.where(((cdf['DocumentType'] == "INV") | (cdf['DocumentType'] == "RNV") | (cdf['DocumentType'] == "CR") | (cdf['DocumentType'] == "RCR") | (
            cdf['DocumentType'] == "DR") | (cdf['DocumentType'] == "RDR")) & ((cdf['SupplyType'] == "EXT") | (cdf['SupplyType'] == "NON") | (cdf['SupplyType'] == "NIL")), "NIL", cdf['Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "NIL") & (cdf['CustomerGSTIN']) & (
            cdf['CustomerGSTIN'].str[0:2] != cdf['SupplierGSTIN'].str[0:2]), "8A", cdf['Sub_Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "NIL") & (cdf['CustomerGSTIN']) & (
            cdf['CustomerGSTIN'].str[0:2] == cdf['SupplierGSTIN'].str[0:2]), "8B", cdf['Sub_Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "NIL") & np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == "0") & (cdf['POS'] != cdf['SupplierGSTIN'].str[:2]), "8C", cdf['Sub_Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "NIL") & np.logical_or(cdf['CustomerGSTIN'].isna(
        ), cdf['CustomerGSTIN'] == "0") & (cdf['POS'] == cdf['SupplierGSTIN'].str[:2]), "8D", cdf['Sub_Category'])

        # 17
        cdf['Category'] = np.where((cdf['DocumentType'] == "RNV") & (cdf['SupplyType'] == "TAX") | (
            cdf["SupplyType"] == "DTA") & np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"), "B2BA", cdf['Category'])
        cdf['Category'] = np.where((cdf['DocumentType'] == "RNV") & (cdf['SupplyType'] == "TAX") | (cdf["SupplyType"] == "DTA") & np.logical_or(
            cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0") & (cdf['IntegratedTaxAmount'] > 0) & (cdf['InvoiceValue'] > 250000), "B2CLA", cdf['Category'])
        cdf['Category'] = np.where((cdf['DocumentType'] == "RNV") & (cdf['SupplyType'] == "EXPWT") | (
            cdf["SupplyType"] == "EXP") | (cdf["SupplyType"] == "SEZ") | (cdf["SupplyType"] == "DXP"), "EXPA", cdf['Category'])
        cdf['Sub_Category'] = np.where((cdf['Category'] == "B2BA") | (
            cdf['Category'] == "B2CLA") | (cdf['Category'] == "EXPA"), "9A", cdf['Sub_Category'])

        # 18
        cdf['Category'] = np.where(((cdf["DocumentType"] == "RFV")
                                    & ((cdf["SupplyType"] == "TAX") | (cdf["SupplyType"] == "DTA"))
                                    & np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
                                   | (((cdf["DocumentType"] == "CR") | (cdf["DocumentType"] == "DR"))
                                      & (cdf["SupplyType"] == "TAX") & np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0")
                                      & (cdf['POS'] != cdf['SupplierGSTIN'].str[:2])
                                      & (cdf["IntegratedTaxAmount"] > 0)
                                      & (cdf["InvoiceValue"] > 250000)
                                      ) | (((cdf["DocumentType"] == "CR") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "RFV"))
                                           & (cdf["SupplyType"] == "TAX")
                                           & (cdf["CustomerGSTIN"])
                                           )
                                   | (((cdf["DocumentType"] == "CR") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "RFV"))
                                      & ((cdf["SupplyType"] == "DXP") | (cdf["SupplyType"] == "EXPT") | (cdf["SupplyType"] == "DTA"))
                                      & (cdf["IntegratedTaxAmount"] > 0)
                                      ) | (((cdf["DocumentType"] == "CR") | (cdf["DocumentType"] == "DR") | (cdf["DocumentType"] == "RFV"))
                                           & ((cdf["SupplyType"] == "SEZ") | (cdf["SupplyType"] == "EXPWT"))
                                           ), "CR/DR/RFV", cdf['Category'])

        cdf['Sub_Category'] = np.where(
            (cdf['Category'] == "CR/DR/RFV"), "9B", cdf['Sub_Category'])

        # 19
        cdf['Category'] = np.where(
            ((cdf['DocumentType'] == "ARFV")
             & ((cdf['SupplyType'] == "TAX") | (cdf['SupplyType'] == "DTA"))
             & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0")))
            |
            (((cdf['DocumentType'] == "RCR") | (cdf['DocumentType'] == "RDR") | (cdf['DocumentType'] == "ARFV"))
             & (cdf['SupplyType'] == "TAX")
             & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
             & (cdf['POS'] != cdf['SupplierGSTIN'].str[0:2])
             & (cdf['IntegratedTaxAmount'] > 0))
            |
            (((cdf['DocumentType'] == "RCR") | (cdf['DocumentType'] == "RDR") | (cdf['DocumentType'] == "ARFV"))
             & ((cdf['SupplyType'] == "TAX") | (cdf['SupplyType'] == "DTA"))
             & (cdf['CustomerGSTIN'])),
            "RCR/RDR/ARFV",
            cdf['Category']
        )

        cdf['Sub_Category'] = np.where(
            (cdf['Category'] == "RCR/RDR/ARFV"), "9C", cdf['Sub_Category'])

        # 20

        cdf['Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0")),
            "B2CSA",
            cdf['Category']
        )

        cdf['Sub_Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0")),
            "10A",
            cdf['Sub_Category']
        )

        # 21
        cdf['Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
            & (cdf['TCSFlag'] == "Y"),
            "B2CSA",
            cdf['Category']
        )

        cdf['Sub_Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
            & (cdf['TCSFlag'] == "Y"),
            "10A(1)",
            cdf['Sub_Category']
        )

        # 22

        cdf['Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
            & (cdf["InvoiceValue"] <= 250000),
            "B2CSA",
            cdf['Category']
        )

        cdf['Sub_Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
            & (cdf["InvoiceValue"] <= 250000),
            "10B",
            cdf['Sub_Category']
        )

        # 23
        cdf['Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
            & (cdf['TCSFlag'] == "Y")
            & (cdf["InvoiceValue"] <= 250000),
            "B2CSA",
            cdf['Category']
        )

        cdf['Sub_Category'] = np.where(
            ((cdf['DocumentType'] == "RNV") | (cdf['DocumentType']
             == "RDR") | (cdf['DocumentType'] == "RCR"))
            & (cdf['SupplyType'] == "TAX")
            & (np.logical_or(cdf['CustomerGSTIN'].isna(), cdf['CustomerGSTIN'] == "0"))
            & (cdf['TCSFlag'] == "Y")
            & (cdf["InvoiceValue"] <= 250000),
            "10B(1)",
            cdf['Sub_Category']
        )

        # GSTR1 LOGICS END

        cdf.to_sql("Sales_Register_temp", panwisedb,
                   if_exists="append", index=False)
    cdf2 = cdf2.reset_index(level=0)
    # cdf2.to_csv(r"C:\CAM\Reports\SRcpivot.csv")
    dtype_dict = {
        'DocumentNumber':'str',
        'TaxableValue':'float',
        'CentralTaxAmount':'float',
        'StateUTTaxAmount':'float',
        'IntegratedTaxAmount':'float',
        'CessAmountSpecific':'float',
        'Quantity':'float',
        'Tax_Amount_Reg':'float'
    }
    if i != 0:

        sql = "DROP TABLE Sales_Register_Processed"
        cur = panwisedb.cursor()
        cur.execute(
            "CREATE TABLE IF NOT EXISTS Sales_Register_Processed (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
        cur.execute(sql)
        panwisedb.commit()
        cur.close()
        sql = "DROP TABLE Sales_Register_Consolidated"
        cur = panwisedb.cursor()
        cur.execute(
            "CREATE TABLE IF NOT EXISTS Sales_Register_Consolidated (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
        cur.execute(sql)
        panwisedb.commit()
        cur.close()
        #   j=1
        query = "SELECT * FROM Sales_Register_temp ORDER BY DocumentNumber ASC"
        for df in pd.read_sql(query, panwisedb, chunksize=10000):
            df.to_sql("Sales_Register_Processed", panwisedb,
                      if_exists="append", index=False)
            #df.to_csv(r"C:\CAM\Reports\1 Sales Register - Loaded and Processed.csv",index=False, header=j, mode='a')
    #        j=0
        query = "SELECT * FROM Sales_Register_temp GROUP BY DocumentNumber ORDER BY DocumentNumber ASC"
        # j=1
        for df in pd.read_sql(query, panwisedb, chunksize=10000):
            df.drop(columns=['TaxableValue', 'CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount',
                    'CessAmountSpecific', 'Quantity', 'Tax_Amount_Reg'], axis=1, inplace=True)
            df = df[(~df['DocumentNumber'].duplicated())
                    | df['DocumentNumber'].isna()]

            for column in ['DocumentNumber','TaxableValue', 'CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount',
                    'CessAmountSpecific', 'Quantity', 'Tax_Amount_Reg']:
                if column not in cdf2.columns:
                    cdf2[column] = 0.0 if dtype_dict.get(column) == 'float' else ''

            df = pd.merge(df, cdf2, on='DocumentNumber', how='left')
            # df = pd.merge(df, cdf2.reset_index(drop=True), on='DocumentNumber', how='left')
            df['TaxableValue'] = df['TaxableValue'].round(2)
            df['DocumentDate'] = pd.to_datetime(
                df['DocumentDate'], dayfirst=True)
            df['ShippingBillDate'] = pd.to_datetime(
                df['ShippingBillDate'], dayfirst=True)

            df['OriginalDocumentDate'] = pd.to_datetime(
                df['OriginalDocumentDate'], dayfirst=True)
            df.to_sql("Sales_Register_Consolidated", panwisedb,
                      if_exists="append", index=False)
            #df.to_csv(r"C:\CAM\Reports\1 Sales Register - Full Conso Processed.csv",index=False, header=j, mode='a')
            # j=0

        c = panwisedb.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS Summary_Totals (Particulars TEXT Primary Key,Count INTEGER, Taxable_Value INTEGER,CGST INTEGER,SGST INTEGER,IGST INTEGER,Cess INTEGER)")
        query = "select * FROM Summary_Totals"

        current = pd.read_sql_query(query, panwisedb)
        current.drop(current[current['Particulars'] ==
                     "Sales Register - Conso"].index, inplace=True)
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM Sales_Register_Consolidated"
        upd = pd.read_sql_query(query, panwisedb)

        upd['Particulars'] = "Sales Register - Conso"
        upd = upd.rename(columns={'SUM(TaxableValue)': 'Taxable_Value', 'SUM(CentralTaxAmount)': 'CGST', 'SUM(StateUTTaxAmount)': 'SGST',
                         'SUM(IntegratedTaxAmount)': 'IGST', 'COUNT(SupplierGSTIN)': 'Count', 'COUNT(CessAmountAdvalorem)': 'Cess'})
        neworder = ['Particulars', 'Count',
                    'Taxable_Value', 'CGST', 'SGST', 'IGST', 'Cess']
        upd = upd.reindex(columns=neworder)
        upd = upd.fillna(0)
        upd = upd.append(current)

        upd.to_sql("Summary_Totals", panwisedb,
                   if_exists="replace", index=False)
        c = panwisedb.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS Summary_Totals (Particulars TEXT Primary Key,Count INTEGER, Taxable_Value INTEGER,CGST INTEGER,SGST INTEGER,IGST INTEGER,Cess INTEGER)")
        query = "select * FROM Summary_Totals"

        current = pd.read_sql_query(query, panwisedb)
        current.drop(current[current['Particulars'] ==
                     "Sales Register - Processed"].index, inplace=True)
        query = "select SUM(TaxableValue),COUNT(SupplierGSTIN),SUM(CentralTaxAmount),SUM(StateUTTaxAmount),SUM(IntegratedTaxAmount),SUM(CessAmountAdvalorem) FROM Sales_Register_Processed"
        upd = pd.read_sql_query(query, panwisedb)

        upd['Particulars'] = "Sales Register - Processed"
        upd = upd.rename(columns={'SUM(TaxableValue)': 'Taxable_Value', 'SUM(CentralTaxAmount)': 'CGST', 'SUM(StateUTTaxAmount)': 'SGST',
                         'SUM(IntegratedTaxAmount)': 'IGST', 'COUNT(SupplierGSTIN)': 'Count', 'COUNT(CessAmountAdvalorem)': 'Cess'})
        neworder = ['Particulars', 'Count',
                    'Taxable_Value', 'CGST', 'SGST', 'IGST', 'Cess']
        upd = upd.reindex(columns=neworder)
        upd = upd.fillna(0)
        upd = upd.append(current)

        upd.to_sql("Summary_Totals", panwisedb, if_exists="replace", index=False)

def gstr1sanitize():
    starttime = time.time()
    # print("started 1 sanitization")
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'

    panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
    date = datetime(2017, 6, 30)
    query = "select * from GSTR_1"
    df = pd.read_sql_query(query, panwisedb)
    df['DocumentDate'] = pd.to_datetime(df['DocumentDate'], dayfirst=True)
    df['Category'] = ''
    df['Category'] = df['GSTR-1 category'].map(str)
    df['Category'] = np.where(df['Supply Type'].map(
        str) != "None", df['Category']+"-" + df['Supply Type'].map(str), df['Category'])
    df['Category'] = np.where(df['ReverseCharge'].map(
        str) == "Y", df['Category']+"-RCM", df['Category'])

    df['Category'] = np.where(df['DocumentDate'] <=
                              date, df['Category']+"-PreGST", df['Category'])
    df['Category'] = np.where(np.logical_and(
        df['GSTR-1 category'].map(str) == "AT", df['POS'] == 97), df['Category']+"-EXP", df['Category'])
    df['Category'] = np.where(np.logical_and(df['GSTR-1 category'].map(
        str) == "ATA", df['POS'] == 97), df['Category']+"-EXP", df['Category'])
    df['Category'] = np.where(np.logical_and(df['GSTR-1 category'].map(
        str) == "TXPD", df['POS'] == 97), df['Category']+"-EXP", df['Category'])
    df['Category'] = np.where(np.logical_and(df['GSTR-1 category'].map(
        str) == "TXPDA", df['POS'] == 97), df['Category']+"-EXP", df['Category'])

    df['cdnpos'] = df['DocumentNo'].map(
        str) + df['SupplierGSTIN'].map(str) + df['DocumentDate'].map(str)
    cdnpos = df.groupby('cdnpos', as_index=False).agg(
        {'POS': 'first', 'GSTR-1 category': 'first', 'Supply Type': 'first'})
    #cdnpos = cdnpos[np.logical_and(cdnpos['Supply Type']!="C",cdnpos['Supply Type']!="D")]
    df['cdnpos'] = df['Original Invoice Number'].map(
        str) + df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)

    df = pd.merge(df, cdnpos, on='cdnpos', how='left')
    # # print(cdf1.columns)
    df['POS_x'] = np.where(df['POS_x'].map(
        str) == "", df['POS_y'], df['POS_x'])
    df = df.rename(columns={'POS_x': 'POS', 'GSTR-1 category_x': 'GSTR-1 category', 'Supply Type_x': 'Supply Type',
                   'GSTR-1 category_y': 'GSTR-1 original category', 'Supply Type_y': 'Supply Type Original'})
    df.drop(columns=['POS_y'], inplace=True)
    #df['Category'] = np.where(df['POS']==97,df['Category']+"-Exp",df['Category'])
    # print(df['GSTR-1 original category'])
    # print(df['Supply Type Original'])
    df['Category'] = np.where(np.logical_and(df['GSTR-1 category'] != df['GSTR-1 original category'], ~
                              df['GSTR-1 original category'].isna()), df['Category']+"-"+df['GSTR-1 original category'], df['Category'])
    df['Category'] = np.where(np.logical_and(df['Supply Type'] != df['Supply Type Original'], ~
                              df['Supply Type Original'].isna()), df['Category']+"-"+df['Supply Type Original'], df['Category'])

    df.to_sql("GSTR_1", panwisedb, if_exists="replace", index=False)
    # tcdf1=tcdf.copy()
    #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)
    df.to_csv(r"GAPS\GSTR1-with category.csv", index=False)
    return "OK"

@router.get('/gstr2asanitize', response_class = Response)
def gstr2asanitize():
    starttime = time.time()
    # print("started 2a key generation")
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'

    panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)

    query = "select * from GSTR_2A"
    df = pd.read_sql_query(query, panwisedb)

    duplicate = df[df.duplicated()]
    duplicate['Reason'] = "Duplicate"
    df.drop_duplicates(keep='first', inplace=True)
    #cdf1['FY'] = pd.to_datetime(cdf1['DocumentDate'],dayfirst=True,errors='coerce')
    #cdf1['FY'] = np.where(cdf1['FY'] != "NaT", cdf1['FY'].map(str).str[0:4],"")

    duplicate1 = df[df.duplicated(subset=[
                                  'CFS', 'SupplierGSTIN', 'DocumentNo', 'FY', 'GSTR-2A category', 'Supply Type'])]
    duplicate1['Reason'] = "Duplicate"
    df.drop_duplicates(subset=['CFS', 'SupplierGSTIN', 'DocumentNo', 'FY',
                       'GSTR-2A category', 'Supply Type'], keep='first', inplace=True)

    df.drop(df[np.logical_or(df['GSTR-2A category'] == "impg",
            df['GSTR-2A category'] == "tds")].index, inplace=True)
    df.drop(df[df['GSTR-2A category'] == "tdsa"].index, inplace=True)
    try:
        df.drop(columns=['level_0'], inplace=True)
    except:
        pass
    df.reset_index(inplace=True)
    # # print(df.columns)

    df['cdnpos'] = df['DocumentNo'].map(
        str) + df['SupplierGSTIN'].map(str) + df['DocumentDate'].map(str)
    cdnpos = df.groupby('cdnpos', as_index=False).agg(
        {'POS': 'first', 'GSTR-2A category': 'first'})
    cdnpos = cdnpos[cdnpos['GSTR-2A category'] == "b2b"]
    df['cdnpos'] = df['Original Invoice Number'].map(
        str) + df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)

    df = pd.merge(df, cdnpos, on='cdnpos', how='left')
    # # print(cdf1.columns)
    df['POS_x'] = np.where(df['POS_x'].map(
        str) == "nan", df['POS_y'], df['POS_x'])
    df = df.rename(
        columns={'POS_x': 'POS', 'GSTR-2A category_x': 'GSTR-2A category'})
    df.drop(columns=['POS_y', 'GSTR-2A category_y'], inplace=True)

    #cdf1.drop(cdf1[~cdf1['CustomerGSTIN'].isin(gstins)].index, inplace=True)
    df['Dropcfs'] = ""
    df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2A category'] == "b2ba", df['CFS'] == "Y"), df['Original Invoice Number'].map(
        str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2A category'].str[0:3], df['Dropcfs'])
    df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2A category'] == "cdna", df['CFS'] == "Y"), df['Original Invoice Number'].map(
        str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2A category'].str[0:3], df['Dropcfs'])
    df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2A category'] == "cdna", df['CFS'] == "Y"), np.where(df['Dropcfs'] == "", df['DocumentNo'].map(
        str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2A category'], df['Dropcfs']), df['Dropcfs'])

    # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
    #cdf1['Dropcfs'] = np.where(np.logical_and(cdf1['GSTR-2A category']=="cdna",cdf1['Dropcfs'] ==""),np.where(cdf1['CFS']=="Y",cdf1['DocumentNo'].map(str)+cdf1['SupplierGSTIN'].map(str)+cdf1['DocumentDate'].map(str)+cdf1['GSTR-2A category'].str[0:2],cdf1['Dropcfs'] ),cdf1['Dropcfs'] )

    cfsdrop = df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
        str)+df['DocumentDate'].map(str)+df['GSTR-2A category']).isin(df['Dropcfs']), df['Dropcfs'] == "")]
    cfsdrop['Reason'] = "Amended"

    df.drop(df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str) +
            df['GSTR-2A category']).isin(df['Dropcfs']), df['Dropcfs'] == "")].index, inplace=True)

    #dropcfsn = df[df['CFS']=="N"]
    #df.drop(df[df['CFS']=="N"].index, inplace=True)
    #dropcfsn['Reason']="Not Filed"

    # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)

    duplicate1 = duplicate.append(cfsdrop)
    # duplicate2=duplicate1.append(dropcfsn)
    # duplicate1.to_csv("PR2A\Dropped 2A.csv", index=False)
    df['DocumentNo'] = df['DocumentNo'].astype(
        str).str.replace('\.0', '', regex=True)
    # cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
    #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)
    # # print(cdf1.dtypes)
    try:
        df.drop(columns=['level_0'], inplace=True)
    except:
        pass
    df.reset_index(inplace=True)
    df['first1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x).start(
    ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
    df['last1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x[::-1]).start(
    ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
    #df['last1']= len(df['DocumentNo']) - df['last1']
    # # print(df['first1'])
    df['first1'] = df['first1'].fillna(df['DocumentNo'].str.len()).astype(int)
    df['last1'] = df['last1'].fillna(df['DocumentNo'].str.len()).astype(int)
    df['DocNoBfrSpl'] = [DocumentNo[:first1]
                         for DocumentNo, first1 in zip(df.DocumentNo, df.first1)]
    df['DocNoAftrSpl'] = [DocumentNo[-last1:]
                          for DocumentNo, last1 in zip(df.DocumentNo, df.last1)]
    df['DocNoWOSplChar'] = df['DocumentNo'].str.replace(
        '[^a-zA-Z0-9]', '', regex=True)
    df['DocNoNumeric'] = df['DocumentNo'].str.replace(
        r"[a-zA-Z]", '').replace('\W', '').replace('_', '').replace('/', '').replace("\\", '')
    # df['DocNoBfrSpl'] =
    df['DocNo'] = df['DocumentNo']
    # df['DocNoAftrSpl'] =
    df['SuppGSTIN'] = df['SupplierGSTIN']
    df['SuppPAN'] = df['SupplierGSTIN'].str[2:12].map(str)
    df['ResGSTIN'] = df['CustomerGSTIN']
    df['ResPAN'] = df['CustomerGSTIN'].str[2:12].map(str)
    df['DocDate'] = pd.to_datetime(df['DocumentDate'], dayfirst=True).apply(
        lambda x: x.strftime('%d%m%Y') if x else "")
    df['MMYYYY'] = pd.to_datetime(df['DocumentDate'], dayfirst=True).apply(
        lambda x: x.strftime('%m%Y') if x else "")
    #df['FY'] = pd.to_datetime(df['DocumentDate'],dayfirst=True).apply(lambda x: x.strftime('%Y') if x else "")
    df['TaxVal'] = df['Taxable Value'].round(2)
    df['InvVal'] = df['InvoiceValue'].round(2)
    df['CGST'] = df['Central Tax Amount'].round(2)
    df['SGST'] = df['State/UT Tax Amount'].round(2)
    df['IGST'] = df['IGST Amount'].round(2)
    df['CessAmount'] = df['CessAmount'].round(2)
    df['GST'] = df['Central Tax Amount'].round(
        2) + df['State/UT Tax Amount'].round(2) + df['IGST Amount'].round(2)+df['CessAmount'].round(2)
    df["POS"].fillna(0, inplace=True)
    # # print(df['POS'])

    df['POS'] = df['POS'].map(int)

    df['RCM'] = df['ReverseCharge']
    df['DocType'] = np.where(df['Supply Type'] == "C", "cdn", np.where(
        df['Supply Type'] == "D", "dbn", "b2b"))

    df['NA'] = ""

    df['Key1'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
        str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key2'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
        str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key3'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key4'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key5'] = df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key6'] = df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key7'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(str)+df['Central Tax Amount'].round(
        2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key8'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
        str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    df['Key9'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
        str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    df['Key10'] = df['FY'].map(str) + df['SuppPAN'].map(str) + df['DocNoWOSplChar'].map(
        str)+df['ResPAN'].map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    ## print("Saved 2A Default Keys")
    df['Key11'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key12'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key13'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key14'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key15'] = df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key16'] = df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key19'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str) + \
        df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)
    df['Key20'] = df['FY'].map(str) + df['SuppPAN'].map(
        str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocType'].map(str)

    df['Key17'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(
        str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key18'] = df['FY'].map(str) + df['SuppPAN'].map(str) + \
        df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)

    try:
        query = "select * FROM matchconfig"
        config = pd.read_sql_query(query, panwisedb)
        rowdata1 = dataframe_to_rows(config, index=False, header=False)
        matchcon = 1
    except:
        matchcon = 0
        config = []

    l = len(config)
    n = 11

    try:
        for row1 in rowdata1:

            key = "Key" + str(n)
            df[key] = df[row1[1]].map(str) + df[row1[2]].map(str) + df[row1[3]].map(str) + df[row1[4]].map(
                str) + df[row1[5]].map(str) + df[row1[6]].map(str) + df[row1[7]].map(str) + df[row1[8]].map(str)
            ## print("prkey ",row[0])
            n += 1
    except:
        pass
    ## print("Saved User Defined Keys")
    df.to_sql("GSTR_2A", panwisedb, if_exists="replace", index=False)
    # tcdf1=tcdf.copy()
    #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)
    df.to_csv(r"GAPS\GSTR2a-with keys.csv", index=False)
    return "OK"

# @router.get('/gstr2bsanitize', response_class = Response)
def gstr2bsanitize(clientPAN, current_user):
    starttime = time.time()
    # print("started 2b key generation")
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'

    panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)

    query = "select * from GSTR_2B"
    df = pd.read_sql_query(query, panwisedb)

    duplicate = df[df.duplicated()]
    duplicate['Reason'] = "Duplicate"
    df.drop_duplicates(keep='first', inplace=True)
    #cdf1['FY'] = pd.to_datetime(cdf1['DocumentDate'],dayfirst=True,errors='coerce')
    #cdf1['FY'] = np.where(cdf1['FY'] != "NaT", cdf1['FY'].map(str).str[0:4],"")

    duplicate1 = df[df.duplicated(subset=[
                                  'SupplierGSTIN', 'DocumentNo', 'FY', 'GSTR-2B category', 'Supply Type'])]
    duplicate1['Reason'] = "Duplicate"
    df.drop_duplicates(subset=['SupplierGSTIN', 'DocumentNo', 'FY',
                       'GSTR-2B category', 'Supply Type'], keep='first', inplace=True)

    # df.drop(df[np.logical_or(df['GSTR-2B category'] == "impg",
    #         df['GSTR-2B category'] == "tds")].index, inplace=True)
    # df.drop(df[df['GSTR-2B category'] == "tdsa"].index, inplace=True)
    try:
        df.drop(columns=['level_0'], inplace=True)
    except:
        pass
    df.reset_index(inplace=True)
    # # print(df.columns)

    df['cdnpos'] = df['DocumentNo'].map(
        str) + df['SupplierGSTIN'].map(str) + df['DocumentDate'].map(str)
    cdnpos = df.groupby('cdnpos', as_index=False).agg(
        {'POS': 'first', 'GSTR-2B category': 'first'})
    cdnpos = cdnpos[cdnpos['GSTR-2B category'] == "b2b"]
    df['cdnpos'] = df['Original Invoice Number'].map(
        str) + df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)

    df = pd.merge(df, cdnpos, on='cdnpos', how='left')
    # # print(cdf1.columns)
    df['POS_x'] = np.where(df['POS_x'].map(
        str) == "nan", df['POS_y'], df['POS_x'])
    df = df.rename(
        columns={'POS_x': 'POS', 'GSTR-2B category_x': 'GSTR-2B category'})
    df.drop(columns=['POS_y', 'GSTR-2B category_y'], inplace=True)

    #cdf1.drop(cdf1[~cdf1['CustomerGSTIN'].isin(gstins)].index, inplace=True)
    df['Dropcfs'] = ""
    df['Dropcfs'] = np.where(df['GSTR-2B category'] == "b2ba", df['Original Invoice Number'].map(
        str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2B category'].str[0:3], df['Dropcfs'])
    df['Dropcfs'] = np.where(df['GSTR-2B category'] == "cdna", df['Original Invoice Number'].map(
        str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2B category'].str[0:3], df['Dropcfs'])
    df['Dropcfs'] = np.where(df['GSTR-2B category'] == "cdna", np.where(df['Dropcfs'] == "", df['DocumentNo'].map(
        str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category'], df['Dropcfs']), df['Dropcfs'])

    # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
    #cdf1['Dropcfs'] = np.where(np.logical_and(cdf1['GSTR-2A category']=="cdna",cdf1['Dropcfs'] ==""),np.where(cdf1['CFS']=="Y",cdf1['DocumentNo'].map(str)+cdf1['SupplierGSTIN'].map(str)+cdf1['DocumentDate'].map(str)+cdf1['GSTR-2A category'].str[0:2],cdf1['Dropcfs'] ),cdf1['Dropcfs'] )

    cfsdrop = df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
        str)+df['DocumentDate'].map(str)+df['GSTR-2B category']).isin(df['Dropcfs']), df['Dropcfs'] == "")]
    cfsdrop['Reason'] = "Amended"

    df.drop(df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str) +
            df['GSTR-2B category']).isin(df['Dropcfs']), df['Dropcfs'] == "")].index, inplace=True)

    #dropcfsn = df[df['CFS']=="N"]
    #df.drop(df[df['CFS']=="N"].index, inplace=True)
    #dropcfsn['Reason']="Not Filed"

    # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)

    duplicate1 = duplicate.append(cfsdrop)
    # duplicate2=duplicate1.append(dropcfsn)
    # duplicate1.to_csv("PR2B\Dropped 2B.csv", index=False)
    df['DocumentNo'] = df['DocumentNo'].astype(
        str).str.replace('\.0', '', regex=True)
    # cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
    #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)
    # # print(cdf1.dtypes)
    try:
        df.drop(columns=['level_0'], inplace=True)
    except:
        pass
    df.reset_index(inplace=True)
    df['first1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x).start(
    ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
    df['last1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x[::-1]).start(
    ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
    #df['last1']= len(df['DocumentNo']) - df['last1']
    # # print(df['first1'])
    df['first1'] = df['first1'].fillna(df['DocumentNo'].str.len()).astype(int)
    df['last1'] = df['last1'].fillna(df['DocumentNo'].str.len()).astype(int)
    df['DocNoBfrSpl'] = [DocumentNo[:first1]
                         for DocumentNo, first1 in zip(df.DocumentNo, df.first1)]
    df['DocNoAftrSpl'] = [DocumentNo[-last1:]
                          for DocumentNo, last1 in zip(df.DocumentNo, df.last1)]
    df['DocNoWOSplChar'] = df['DocumentNo'].str.replace(
        '[^a-zA-Z0-9]', '', regex=True)
    df['DocNoNumeric'] = df['DocumentNo'].str.replace(
        r"[a-zA-Z]", '').replace('\W', '').replace('_', '').replace('/', '').replace("\\", '')
    # df['DocNoBfrSpl'] =
    df['DocNo'] = df['DocumentNo']
    # df['DocNoAftrSpl'] =
    df['SuppGSTIN'] = df['SupplierGSTIN']
    df['SuppPAN'] = df['SupplierGSTIN'].str[2:12].map(str)
    df['ResGSTIN'] = df['CustomerGSTIN']
    df['ResPAN'] = df['CustomerGSTIN'].str[2:12].map(str)
    df['DocDate'] = pd.to_datetime(df['DocumentDate'], dayfirst=True).apply(
        lambda x: x.strftime('%d%m%Y') if x else "")
    df['MMYYYY'] = pd.to_datetime(df['DocumentDate'], dayfirst=True).apply(
        lambda x: x.strftime('%m%Y') if x else "")
    #df['FY'] = pd.to_datetime(df['DocumentDate'],dayfirst=True).apply(lambda x: x.strftime('%Y') if x else "")
    df['TaxVal'] = df['Taxable Value'].round(2)
    df['InvVal'] = df['InvoiceValue'].round(2)
    df['CGST'] = df['Central Tax Amount'].round(2)
    df['SGST'] = df['State/UT Tax Amount'].round(2)
    df['IGST'] = df['IGST Amount'].round(2)
    df['CessAmount'] = df['CessAmount'].round(2)
    df['GST'] = df['Central Tax Amount'].round(
        2) + df['State/UT Tax Amount'].round(2) + df['IGST Amount'].round(2)+df['CessAmount'].round(2)
    df["POS"].fillna(0, inplace=True)
    # # print(df['POS'])

    df['POS'] = df['POS'].map(int)

    df['RCM'] = df['ReverseCharge']
    df['DocType'] = np.where(df['Supply Type'] == "C", "cdn", np.where(
        df['Supply Type'] == "D", "dbn", "b2b"))

    df['NA'] = ""

    df['Key1'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
        str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key2'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
        str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key3'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key4'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key5'] = df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key6'] = df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key7'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(str)+df['Central Tax Amount'].round(
        2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key8'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
        str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    df['Key9'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
        str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    df['Key10'] = df['FY'].map(str) + df['SuppPAN'].map(str) + df['DocNoWOSplChar'].map(
        str)+df['ResPAN'].map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    ## print("Saved 2A Default Keys")
    df['Key11'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key12'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key13'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key14'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key15'] = df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key16'] = df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key19'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str) + \
        df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)
    df['Key20'] = df['FY'].map(str) + df['SuppPAN'].map(
        str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocType'].map(str)

    df['Key17'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(
        str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key18'] = df['FY'].map(str) + df['SuppPAN'].map(str) + \
        df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)

    try:
        query = "select * FROM matchconfig_2b"
        config = pd.read_sql_query(query, panwisedb)
        rowdata1 = dataframe_to_rows(config, index=False, header=False)
        matchcon = 1
    except:
        matchcon = 0
        config = []

    l = len(config)
    n = 11

    try:
        for row1 in rowdata1:

            key = "Key" + str(n)
            df[key] = df[row1[1]].map(str) + df[row1[2]].map(str) + df[row1[3]].map(str) + df[row1[4]].map(
                str) + df[row1[5]].map(str) + df[row1[6]].map(str) + df[row1[7]].map(str) + df[row1[8]].map(str)
            ## print("prkey ",row[0])
            n += 1
    except:
        pass
    ## print("Saved User Defined Keys")
    df.to_sql("GSTR_2B", panwisedb, if_exists="replace", index=False)
    # tcdf1=tcdf.copy()
    #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)
    #df.to_csv(r"GAPS\GSTR2a-with keys.csv", index=False)
    return "OK"

def valgldump(is_local, parameters):
    # dir_path = os.path.dirname(os.path.realpath(__file__))
    # path = dir_path + '/Client-Details'
    # panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)

    dir_path = os.path.dirname(os.path.realpath(__file__))
    folder_name = '' if is_local else parameters[1]
    # path = dir_path + '/Client-Details' if is_local else dir_path + folder_name
    path = os.path.join(dir_path, 'Client-Details') if is_local else os.path.join(dir_path, folder_name)
    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]
    # panwisedb_path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'
    panwisedb_path = os.path.join(path, current_user, clientPAN, clientPAN + '.db') if is_local else os.path.join(path, temp_db_name + '.db')
    panwisedb = sqlite3.connect(panwisedb_path, timeout=10)

    try:
        query = "select `GSTIN_TB`, `G/L_GL`, `Account_Name`, `Account_Name_Main_Heading`, `IND AS Sub Heading` from TrialBalanceDigi"
        tb = pd.read_sql_query(query, panwisedb)
    except:
        pass
    # tb = trial_balance
    # tb.drop(columns=['Opening Balance','Closing Balance','Debit','Credit'], axis =1, inplace = True)
    query = "SELECT * FROM GLCodeDigi"
    glc = pd.read_sql_query(query, panwisedb)
    # glc = glcode_master
    query = "SELECT * FROM TaxCodeDigi"
    tcm = pd.read_sql_query(query, panwisedb)
    # tcm = taxcode_master
    # cdf2 = pd.merge(cdf2,cdf3,on='Transaction_Type_GL',how='left') 
    query = "SELECT * FROM DocTypeDigi"
    dtm = pd.read_sql_query(query, panwisedb)
    # dtm = doctype_master
    sql = 'DROP TABLE gld'
    cur = panwisedb.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS gld (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
    cur.execute(sql)
    panwisedb.commit()

    
    query = "SELECT * FROM GLDump ORDER BY Accounting_Document_Number_GL ASC"
    df = pd.DataFrame()
    df2 = pd.DataFrame()
    date = datetime(2018,3,31)
    date1 = datetime(2019,3,31)
    
    cdf3=pd.DataFrame()        
    for cdf2 in pd.read_sql_query(query, panwisedb, chunksize=10000):
        try:
            cdf2 = pd.merge(cdf2,tb,on='G/L_GL',how='left')
        except:
            pass
        cdf2 = pd.merge(cdf2,tcm,on='Transaction_Type_GL',how='left')
        cdf2 = pd.merge(cdf2,dtm,left_on='Document_Type_GL',right_on='Document_Type',how='left')
        cdf2['Posting_Date_GL']= pd.to_datetime(cdf2['Posting_Date_GL'],dayfirst=True) 
        cdf2['Document_Date_GL']= pd.to_datetime(cdf2['Document_Date_GL'],dayfirst=True) 
    
        cdf2['Period'] = np.where(cdf2['Document_Date_GL'] <= date,"2017-18",np.where(cdf2['Document_Date_GL'] <= date1,"2018-19","2019-20"))
        
        #cdf2.to_csv(r"C:\CAM\Reports\GL Dump Report1.csv",index=False)
        
        #cdf2.drop(cdf2[np.logical_or(cdf2['Posting_Date_GL']<=date,cdf2['Posting_Date_GL']>date1)].index, inplace=True)
        #cdf2.to_csv(r"C:\CAM\Reports\GL Dump Report2.csv",index=False)
        
        try:
            cdf2['GL_Type'] = "GL Code Missing in Master"
            cdf2.drop(cdf2[cdf2['G/L_GL']==''].index, inplace=True)
        except:
            pass
        cdf2.drop(cdf2[cdf2['G/L_GL'].isna()].index, inplace=True)

        #cdf2['GL_Type'] = np.where((cdf2['G/L_GL'].isin(glc['CGST_Output'])) |(cdf2['G/L_GL'].isin(glc['SGST_Output']))|(cdf2['G/L_GL'].isin(glc['IGST_Output']))|(cdf2['G/L_GL'].isin(glc['GST_Output']))|(cdf2['G/L_GL'].isin(glc['Input Tax GLs']))|(cdf2['G/L_GL'].isin(glc['RCM GLs'])),"Tax",cdf2['GL_Type']) 
        cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['Revenue GLs']),"Revenue",cdf2['GL_Type'])
        cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['Taxable Advance (Liability) GLs']),"Advance",cdf2['GL_Type'])
        cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['Forex Gls Part of revenue']),"Forex",cdf2['GL_Type'])
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['CGST_Output'].map(str)),"CGST_GL",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['SGST_Output'].map(str)),"SGST_GL",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['IGST_Output'].map(str)),"IGST_GL",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['UGST_Output'].map(str)),"SGST_GL",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['CGST_Input']),"CGST_GL_Input",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['SGST_Input']),"SGST_GL_Input",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['IGST_Input']),"IGST_GL_Input",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['GL_Type'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['UGST_Input']),"SGST_GL_Input",cdf2['GL_Type'])
        except:
            pass
        try:
            cdf2['Tax GL'] = np.where((cdf2['G/L_GL'].map(str).isin(glc['CGST_Output'].map(str))) |(cdf2['G/L_GL'].map(str).isin(glc['SGST_Output'].map(str)))|(cdf2['G/L_GL'].map(str).isin(glc['IGST_Output'].map(str)))|(cdf2['G/L_GL'].map(str).isin(glc['UGST_Output'].map(str))),1,0) 
        except:
            cdf2['Tax GL'] = 0
        cdf2['Revenue GL'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['Revenue GLs']),1,0) 
        cdf2['Advance GL'] = np.where(cdf2['G/L_GL'].map(str).isin(glc['Taxable Advance (Liability) GLs']),1,0)
        cdf2['Key']=cdf2['GL_Type'].map(str)+cdf2['Accounting_Document_Number_GL'].map(str)
        try:
            cdf21 = cdf2.pivot_table(index=['Accounting_Document_Number_GL'], values=['Tax GL','Revenue GL','Advance GL'], aggfunc=sum,dropna=False)
            cdf3=cdf3.append(cdf21)
        except:
            pass
        try:
            cdf3 = cdf3.groupby(['Accounting_Document_Number_GL']).agg('sum')
        except:
            pass
        #cdf3['Accounting_Document_Number_GL'] = cdf3['Accounting_Document_Number_GL'].astype(str)
        
        df1 = cdf2.pivot_table(index=['Accounting_Document_Number_GL'], columns=['GL_Type'], values=['Amount_GL'],aggfunc=sum,dropna=False)
        df1= df1.droplevel(0, axis=1)
        df2 = df2.append(df1)
        df2 = df2.groupby(['Accounting_Document_Number_GL']).agg('sum')
        #df2['Accounting_Document_Number_GL'] = df2['Accounting_Document_Number_GL'].astype(str)
        cdf2.to_sql("gld", panwisedb, if_exists="append",index=False)
    sql = "DROP TABLE GL_Dump_Line_Level"
    cur.execute("CREATE TABLE IF NOT EXISTS GL_Dump_Line_Level (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
    cur.execute(sql)
    panwisedb.commit()
    sql = "DROP TABLE GL_Dump_Consolidated"
    cur.execute("CREATE TABLE IF NOT EXISTS GL_Dump_Consolidated (Particulars TEXT PRIMARY KEY, GSTIN TEXT, From_Date TEXT, To_Date TEXT)")
    cur.execute(sql)
    panwisedb.commit()
    j=1
    query = "SELECT * FROM gld ORDER BY Accounting_Document_Number_GL ASC"
    for df in pd.read_sql(query, panwisedb, chunksize=10000):
        df=df[df['Accounting_Document_Number_GL'].isnull() | ~df[df['Accounting_Document_Number_GL'].notnull()].duplicated(subset='Accounting_Document_Number_GL',keep='first')]
        df.drop(df[df['Accounting_Document_Number_GL'].isna()].index, inplace=True)
        df.drop(df[df['Accounting_Document_Number_GL']==''].index, inplace=True)
        df.to_sql("GL_Dump_Line_Level", panwisedb, if_exists="append",index=False)
        #df.to_csv(r"C:GAPSReports1 Sales Register - Loaded and Processed.csv",index=False, header=j, mode='a')
    df2.index=df2.index.astype(str)
    cdf3.index=cdf3.index.astype(str)
            
    query = "SELECT * FROM GL_Dump_Line_Level GROUP BY Key ORDER BY Accounting_Document_Number_GL ASC"
    j=1
    for df in pd.read_sql(query, panwisedb, chunksize=10000):
        try:
            df.drop(columns=['Tax GL','Revenue GL','Advance GL'], axis =1, inplace = True)
        except:
            pass
        df=df[df['Accounting_Document_Number_GL'].isnull() | ~df[df['Accounting_Document_Number_GL'].notnull()].duplicated(subset='Accounting_Document_Number_GL',keep='first')]
        
        
        df['Accounting_Document_Number_GL'] = df['Accounting_Document_Number_GL'].astype(str)
        try:
            df = pd.merge(df,cdf3,on='Accounting_Document_Number_GL',how='left')
        except:
            pass
        try:
            df = pd.merge(df,df2,on='Accounting_Document_Number_GL',how='left')
        except:
            pass
        try:
            df['Amount_GL'] = df.apply(lambda row: row[row['GL_Type']], axis=1)
        except:
            pass
        df['Tax GL'] = np.where(df['Tax GL'] >=1,1,0) 
        df['Revenue GL'] = np.where(df['Revenue GL']>=1,3,0) 
        df['Advance GL'] = np.where(df['Advance GL']>=1,7,0) 
        df['Category'] = df['Tax GL']+df['Revenue GL']+df['Advance GL']
        #df['Amount_GL']=df['Amount_GL'].round(2)
        df['Category'] = np.where(df['Category']==3,"Document Number of Revenue, not in Tax",df['Category'])
        df['Category'] = np.where(df['Category']=="7","Document Number of Advance, not in Tax",df['Category'])
        df['Category'] = np.where(df['Category']=="1","Document Number of Tax, not in Revenue and Advance",df['Category'])
        df['Category'] = np.where(df['Category']=="4","Common Document Numbers in Revenue and Tax entries",df['Category'])
        df['Category'] = np.where(df['Category']=="8","Common Document Numbers in Advance and Tax entries",df['Category'])
        df['Category'] = np.where(df['Category']=="10","Document Number in both Advance and Revenue, not in Tax",df['Category'])
        df['Category'] = np.where(df['Category']=="0","Null",df['Category'])
        df['Category'] = np.where(df['Category'].isna(),"Null",df['Category'])
        #print (df.columns)
        
        #df['CGST_GL'] = np.where(df['Category']!="Common Document Numbers in Revenue and Tax entries",np.where(df['GL_Type']!="CGST_GL",0,df['CGST_GL']),df['CGST_GL'])
        #df['SGST_GL'] = np.where(df['Category']!="Common Document Numbers in Revenue and Tax entries",np.where(df['GL_Type']!="SGST_GL",0,df['SGST_GL']),df['SGST_GL'])
        #df['IGST_GL'] = np.where(df['Category']!="Common Document Numbers in Revenue and Tax entries",np.where(df['GL_Type']!="IGST_GL",0,df['IGST_GL']),df['IGST_GL'])
        #df['RCM_GL'] = np.where(df['Category']!="Common Document Numbers in Revenue and Tax entries",np.where(df['GL_Type']!="RCM_GL",0,df['RCM_GL']),df['RCM_GL'])
        #df.drop(df[np.logical_and(df['Category']=="Common Document Numbers in Revenue and Tax entries",df['GL_Type']!="Revenue")].index, inplace=True)
        df['Key'] = np.where(df['Category']=="Document Number of Tax, not in Revenue and Advance",df['Key'],df['Reference_GL'])
        df['Accounting_Document_Number_Reg'] = df['Accounting_Document_Number_GL']
        #df.reset_index().drop_duplicates(subset='Accounting_Document_Number_GL',keep='first', inplace= True)
        ## print(df)   
        #df1=df[df['Accounting_Document_Number_GL'].isnull() | ~df[df['Accounting_Document_Number_GL'].notnull()].duplicated(subset='Accounting_Document_Number_GL',keep='first')]
        #df=df.reindex()
        df['GSTR 9C'] = " "
        df['GSTR 9C'] = np.where(df['GL_Type']=="Revenue","Table 5A",df['GSTR 9C'])
        df['GSTR 9C'] = np.where(df['GL_Type']=="Forex","Table 5N",df['GSTR 9C'])
        df['GSTR 9C'] = np.where(np.logical_and(df['Document_Type_MS']=="INV",df['Tax_Rate_MS']==0),"Table 7C",df['GSTR 9C'])
        df['GSTR 9C'] = np.where(df['Document_Type_MS']=="NIL","Table 7B",df['GSTR 9C'])
        df['GSTR 9C'] = np.where(df['Document_Type_MS']=="EXP","Table 7C",df['GSTR 9C'])
        
        #df.to_csv(r"C:\GAPS\Reports\GL Dump - Full Conso Processed.csv",index=False, header=j, mode='a')
        df.to_sql("GL_Dump_Consolidated", panwisedb, if_exists="append",index=False)
        j=0

# @router.post('/triggerval', response_class = Response)
def validatepr(clientPAN, current_user):
    starttime = datetime.now()
    print(clientPAN, current_user)
    router.message = "Flash Validation Started......"
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = dir_path + '/Client-Details'
    panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)
    router.message = "Running Flash validation on Inward Data...."
    lst = ["False", "False", "False", "False", "False", "False", "False"]

    queuedb = sqlite3.connect('queue.db')
    cursor = queuedb.cursor()

    try:
        seqRun1(lst, clientPAN, current_user) # inward flash/pr flash
    except:
        traceback.print_exc()
    router.message = "Running GAPS validation on Inward Data...."
    # try:
    #         valpr(clientPAN, current_user) # inward validation/validate purchase register
    #         cursor.execute(f'''update queue set prvalidationendtime = '{datetime.now()}', prvalidationendtimetaken = '{datetime.now()-starttime}' where clientPAN = '{clientPAN}' and username = '{current_user}';''')
    #         queuedb.commit()
    #     except:
    #         traceback.print_exc()

# def validate2b(clientPAN, current_user):
#     starttime = datetime.now()
#     print(clientPAN, current_user)
#     dir_path = os.path.dirname(os.path.realpath(__file__))
#     path = dir_path + '/Client-Details'
#     panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)
#     queuedb = sqlite3.connect('queue.db')
#     cursor = queuedb.cursor()
#     try:
#         query = "select Key1 from GSTR_2B"
#         doclist2b = pd.read_sql_query(query, panwisedb)
#     except:
#         doclist2b = pd.DataFrame(columns=['NoKey'])
#     if "NoKey" in doclist2b.columns:
#         router.message = "Sanitizing GSTR 2B"
#         try:
#             gstr2bsanitize(clientPAN, current_user)
#         except:
#             traceback.print_exc()
#     elif any(doclist2b['Key1'].isna()) or any(doclist2b['Key1'] == ""):
#         router.message = "Sanitizing GSTR 2B"
#         try:
#             gstr2bsanitize(clientPAN, current_user)
#         except:
#             traceback.print_exc()
    # cursor.execute(f'''update queue set "2bvalidationendtime" = '{datetime.now()}', "2bvalidationendtimetaken" = '{datetime.now()-starttime}' where clientPAN = '{clientPAN}' and username = '{current_user}';''')
    # queuedb.commit()

def gstr2ajson(files, clientPAN, filename1, current_user):
    starttime = time.time()
    # current_user = current_user.get('current_user')
    path = dir_path + '/Client-Details'
    panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)
    try:
        query = "select DocKey from GSTR_2A"
        doclist = pd.read_sql_query(query, panwisedb)
    except:
        doclist = pd.DataFrame(columns=['DocKey'])
    df = pd.DataFrame()
    csv_data = StringIO()
    csv_writer = csv.writer(csv_data)

    #workbook = xlsxwriter.Workbook(r"GAPS\Gstr2a.xlsx")
    #worksheet_b2b = workbook.add_worksheet('Sheet1')
    global row
    row = 0
    csv_writer.writerow(['TaxPeriod', 'SupplierGSTIN', 'DocumentNo',
                         'DocumentDate', 'CustomerGSTIN', 'POS', 'ReverseCharge',
                         'Taxable Value', 'IGST Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'CessAmount',
                         'InvoiceValue', 'Tax Rate', 'Original Invoice Number', 'Original Invoice Date', 'Supply Type',
                         'GSTR-2A category', 'Filing Date', 'Filing Period', 'GSTR 3B Filed', 'CFS', 'IRN', 'IRN Gen-Date',
                         'Differential %', 'Cancelation Date', 'Delinking Flag'])

    # # print(file)
    # # print(myzip.namelist()[0])

    # # print(myfile)
    parsed_json = json.loads(files)
    # # print(parsed_json.keys())
    # dict_keys(['gstin', 'cdn', 'b2b', 'fp'])
    # # print(parsed_json['fp'])
    doc_type_list = ['b2b', 'b2ba']
    for doc_type in doc_type_list:
        if doc_type in parsed_json:
            # # print(doc_type)
            for supplier in parsed_json[doc_type]:
                # # print(supplier)
                for inv in supplier['inv']:
                    # # print(inv)
                    inv_values = calc_inv_value(inv['itms'])
                    # # print(inv)
                    # # print(doc_type)
                    if doc_type == 'b2ba':
                        oinum, oidt, b2ba = (
                            inv['oinum'], inv['oidt'], inv['inv_typ'])
                    else:
                        oinum, oidt, b2ba = (
                            '', '', inv['inv_typ'] if 'inv_typ' in inv else '')
                    row += 1
                    csv_writer.writerow([parsed_json['fp'],
                                         supplier['ctin'],
                                         inv['inum'], inv['idt'], parsed_json['gstin'], inv['pos'] if 'pos' in inv else '',
                                         inv['rchrg'], inv_values['txval'], inv_values['iamt'], inv_values['camt'],
                                         inv_values['samt'], inv_values['csamt'], inv['val'], ','.join(
                        str(x) for x in inv_values['tax_rate']),
                        oinum, oidt, b2ba, doc_type, supplier['fldtr1'] if 'fldtr1' in supplier else '',
                        supplier['flprdr1'] if 'flprdr1' in supplier else '', supplier['cfs3b'] if 'cfs3b' in supplier else '', supplier['cfs'],
                        inv['irn'] if 'irn' in inv else '', inv['irngendate'] if 'irngendate' in inv else '',
                        inv['diff_percent'] if 'diff_percent' in inv else '', supplier['dtcancel'] if 'dtcancel' in supplier else '',
                        inv['d_flag'] if 'd_flag' in inv else ''])

    doc_type_list = ['cdn', 'cdna']
    for doc_type in doc_type_list:
        if doc_type in parsed_json:
            for supplier in parsed_json[doc_type]:
                for inv in supplier['nt']:
                    inv_values = calc_inv_value(inv['itms'])

                    if doc_type == 'cdna':
                        oinum, oidt, b2ba = (
                            inv['ont_num'], inv['ont_dt'], inv['ntty'])
                    else:
                        oinum, oidt, b2ba = (
                            inv['inum'] if 'inum' in inv else '', inv['idt'] if 'idt' in inv else '', inv['ntty'])
                    row += 1
                    if b2ba == "D":
                        csv_writer.writerow([parsed_json['fp'],
                                             supplier['ctin'],
                                             inv['nt_num'], inv['nt_dt'], parsed_json['gstin'], inv['pos'] if 'pos' in inv else '',
                                             inv['rchrg'] if 'rchrg' in inv else '', inv_values['txval'], inv_values['iamt'], inv_values['camt'],
                                             inv_values['samt'], inv_values['csamt'], inv['val'], ','.join(
                            str(x) for x in inv_values['tax_rate']),
                            oinum, oidt, b2ba, doc_type, supplier['fldtr1'] if 'fldtr1' in supplier else '',
                            supplier['flprdr1'] if 'flprdr1' in supplier else '', supplier[
                                'cfs3b'] if 'cfs3b' in supplier else '', supplier['cfs'],
                            inv['irn'] if 'irn' in inv else '', inv['irngendate'] if 'irngendate' in inv else '',
                            inv['diff_percent'] if 'diff_percent' in inv else '', supplier['dtcancel'] if 'dtcancel' in supplier else '',
                            inv['d_flag'] if 'd_flag' in inv else ''])
                    if b2ba == "C":
                        csv_writer.writerow([parsed_json['fp'],
                                             supplier['ctin'],
                                             inv['nt_num'], inv['nt_dt'], parsed_json['gstin'], inv['pos'] if 'pos' in inv else '',
                                             inv['rchrg'] if 'rchrg' in inv else '', -
                                             inv_values['txval'], -
                                             inv_values['iamt'], -
                                             inv_values['camt'],
                                             -inv_values['samt'], -inv_values['csamt'], -
                                             inv['val'], ','.join(
                                                 str(x) for x in inv_values['tax_rate']),
                                             oinum, oidt, b2ba, doc_type, supplier[
                                                 'fldtr1'] if 'fldtr1' in supplier else '',
                                             supplier['flprdr1'] if 'flprdr1' in supplier else '', supplier[
                            'cfs3b'] if 'cfs3b' in supplier else '', supplier['cfs'],
                            inv['irn'] if 'irn' in inv else '', inv['irngendate'] if 'irngendate' in inv else '',
                            inv['diff_percent'] if 'diff_percent' in inv else '', supplier['dtcancel'] if 'dtcancel' in supplier else '',
                            inv['d_flag'] if 'd_flag' in inv else ''])
    doc_type_list = ['tds', 'tdsa']
    for doc_type in doc_type_list:
        if doc_type in parsed_json:
            # # print(doc_type)
            for inv in parsed_json[doc_type]:
                # # print(supplier)

                # # print(inv)
                # # print(doc_type)

                row += 1
                csv_writer.writerow([parsed_json['fp'],
                                     inv['gstin_deductor'],
                                     inv['deductor_name'], '', parsed_json['gstin'], '',
                                     '', inv['amt_ded'], inv['iamt'], inv['camt'],
                                     inv['samt'], '', inv['amt_ded'], '', '', '', '', doc_type, '', inv['month'], '', '', '', '', '', '', ''])
    doc_type_list = ['impg', 'impga']
    for doc_type in doc_type_list:
        if doc_type in parsed_json:
            # # print(doc_type)
            for inv in parsed_json[doc_type]:
                # # print(supplier)

                # # print(inv)
                # # print(doc_type)

                row += 1
                csv_writer.writerow([parsed_json['fp'],
                                     inv['sgstin'],
                                     inv['benum'], inv['bedt'], parsed_json['gstin'], '',
                                     '', inv['txval'], inv['iamt'], '',
                                     '', inv['csamt'], '', '', '', '', inv['portcd'], doc_type, inv['refdt'], '', '', '', '', '', '', '', ''])

    date = datetime(2018, 3, 31)
    date1 = datetime(2019, 3, 31)
    # workbook.close()
    csv_data.seek(0)
    # # print(csv_data)
    df = pd.read_csv(csv_data, encoding='utf-8')

    df.drop(index=clientPAN_check(df, 'CustomerGSTIN'), inplace=True)
    if df.shape[0] != 0:
        df = df.reset_index(drop=True)
    # # print(csv_data)

    df['DocumentDate'] = pd.to_datetime(df['DocumentDate'], dayfirst=True)
    df['FY'] = df['DocumentDate'].map(
        lambda x: x.year if x.month > 3 else x.year-1)
    df['FY'] = np.where(df['FY'].isna(), "nan", df['FY'])
    df['FY'] = np.where(df['FY'].map(str) != "nan",
                        df['FY'].map(str).str[0:4], df['FY'])
    df['Original Invoice Date'] = pd.to_datetime(
        df['Original Invoice Date'], dayfirst=True)

    pd.to_numeric(df['POS'])
    df['POS'] = df['POS'].map(str)
    df['POS'] = df["POS"].apply(lambda x: x.replace(r'.0', '') if isinstance(x, str) else x)
    # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                
    df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
    # print(df['POS'].unique())

    df['Supplier Name'] = ''
    df = df[['TaxPeriod', 'SupplierGSTIN', 'Supplier Name', 'DocumentNo', 'DocumentDate',
       'CustomerGSTIN', 'POS', 'ReverseCharge', 'Taxable Value', 'IGST Amount',
       'Central Tax Amount', 'State/UT Tax Amount', 'CessAmount',
       'InvoiceValue', 'Tax Rate', 'Original Invoice Number',
       'Original Invoice Date', 'Supply Type', 'GSTR-2A category',
       'Filing Date', 'Filing Period', 'GSTR 3B Filed', 'CFS', 'IRN',
       'IRN Gen-Date', 'Differential %', 'Cancelation Date', 'Delinking Flag',
       'FY']]

    df['DocumentNo'] = df['DocumentNo'].map(str).str.upper()
    try:
        df['DocumentNo'] = df['DocumentNo'].str.lstrip('0')
    except:
        pass

    try:
        df['Original Invoice Number'] = df['Original Invoice Number'].str.lstrip(
            '0')
    except:
        pass
    df['Tax'] = df['State/UT Tax Amount'] + \
        df['Central Tax Amount']+df['IGST Amount']
    df['SourceFile'] = str(filename1)
    df['UploadTime'] = datetime.now()
    df['DocKey'] = df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
        str)+df['FY'].map(str)+df['Supply Type'].map(str)+df['GSTR-2A category'].map(str)
    totalcount = len(df)
    try:
        drop = df[df['DocKey'].isin(doclist['DocKey'])]
        drop.to_sql("GSTR_2Adropped", panwisedb,
                    if_exists="append", index=False)
        df.drop(df[df['DocKey'].isin(doclist['DocKey'])].index, inplace=True)
    except:
        pass
    processedcount = len(df)
    droppedcount = totalcount - processedcount

    # tcdf1=tcdf.copy()
    #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)
    df.to_sql("GSTR_2A", panwisedb, if_exists="append", index=False)
    return totalcount, processedcount, droppedcount

def gstr2bjson(files, clientPAN, filename1, current_user):

    starttime = time.time()
    # username = current_user.get('username')
    path = dir_path + '/Client-Details'
    panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)
    try:
        query = "select DocKey from GSTR_2B"
        doclist = pd.read_sql_query(query, panwisedb)
    except:
        doclist = pd.DataFrame(columns=['DocKey'])

    df = pd.DataFrame()
    csv_data = StringIO()
    csv_writer = csv.writer(csv_data)

    #workbook = xlsxwriter.Workbook(r"GAPS\Gstr2a.xlsx")
    #worksheet_b2b = workbook.add_worksheet('Sheet1')
    global row
    row = 0
    csv_writer.writerow(['TaxPeriod', 'SupplierGSTIN', 'Supplier Name','DocumentNo',
                         'DocumentDate', 'CustomerGSTIN', 'POS', 'ReverseCharge',
                         'Taxable Value', 'IGST Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'CessAmount', 'InvoiceValue', 'Tax Rate', 'Original Invoice Number', 'Original Invoice Date', 'Supply Type', 'GSTR-2B category', 'Filing Date', 'Filing Period', 'GSTR 3B Filed'])

    parsed_json = json.loads(files)
    # # print(parsed_json.keys())
    # dict_keys(['gstin', 'cdn', 'b2b', 'fp'])
    # # print(parsed_json['fp'])
    doc_type_list = ['b2b', 'b2ba']
    for doc_type in doc_type_list:
        if doc_type in parsed_json['data']['docdata']:
            # # print(doc_type)
            for supplier in parsed_json['data']['docdata'][doc_type]:
                # # print(supplier)
                for inv in supplier['inv']:
                    # # print(inv)
                    inv_values = calc_inv_value1(inv['items'])
                    # # print(inv)
                    # # print(doc_type)
                    if doc_type == 'b2ba':
                        oinum, oidt, b2ba = (
                            inv['oinum'], inv['oidt'], inv['typ'])
                    else:
                        oinum, oidt, b2ba = (
                            '', '', inv['typ'] if 'typ' in inv else '')
                    row += 1
                    csv_writer.writerow([parsed_json['data']['rtnprd'],
                                         supplier['ctin'],supplier['trdnm'],
                                         inv['inum'], inv['dt'], parsed_json['data']['gstin'], inv['pos'] if 'pos' in inv else '',
                                         inv['rev'] if 'rev' in inv else '', inv_values['txval'], inv_values['igst'], inv_values['cgst'],
                                         inv_values['sgst'], inv_values['cess'], inv['val'], ','.join(str(x) for x in inv_values['tax_rate']), oinum, oidt, b2ba, doc_type, supplier['fldtr1'] if 'fldtr1' in supplier else '', supplier['supfildt'] if 'supfildt' in supplier else '', supplier['supprd'] if 'supprd' in supplier else ''])

    doc_type_list = ['cdnr', 'cdnra']
    for doc_type in doc_type_list:
        if doc_type in parsed_json['data']['docdata']:
            for supplier in parsed_json['data']['docdata'][doc_type]:
                for inv in supplier['nt']:
                    inv_values = calc_inv_value1(inv['items'])

                    if doc_type == 'cdnra':
                        oinum, oidt, b2ba = (
                            inv['ontnum'], inv['ontdt'], inv['typ'])
                    else:
                        oinum, oidt, b2ba = (
                            inv['inum'] if 'inum' in inv else '', inv['idt'] if 'idt' in inv else '', inv['typ'])
                    row += 1
                    if b2ba == "D":
                        csv_writer.writerow([parsed_json['data']['rtnprd'],
                                             supplier['ctin'],supplier['trdnm'],
                                             inv['ntnum'], inv['dt'], parsed_json['data']['gstin'], inv['pos'] if 'pos' in inv else '',
                                             inv['rev'] if 'rev' in inv else '', inv_values['txval'], inv_values['igst'], inv_values['cgst'],
                                             inv_values['sgst'], inv_values['cess'], inv['val'], ','.join(str(x) for x in inv_values['tax_rate']), oinum, oidt, b2ba, doc_type, supplier['fldtr1'] if 'fldtr1' in supplier else '', supplier['flprdr1'] if 'flprdr1' in supplier else '', supplier['cfs3b'] if 'cfs3b' in supplier else ''])
                    if b2ba == "C":
                        csv_writer.writerow([parsed_json['data']['rtnprd'],
                                             supplier['ctin'],supplier['trdnm'],
                                             inv['ntnum'], inv['dt'], parsed_json['data']['gstin'], inv['pos'] if 'pos' in inv else '',
                                             inv['rev'] if 'rev' in inv else '', -inv_values['txval'], -
                                             inv_values['igst'], -
                                             inv_values['cgst'],
                                             -inv_values['sgst'], -inv_values['cess'], -inv['val'], ','.join(str(x) for x in inv_values['tax_rate']), oinum, oidt, b2ba, doc_type])
    
    date = datetime(2018, 3, 31)
    date1 = datetime(2019, 3, 31)
    # workbook.close()
    csv_data.seek(0)
    # # print(csv_data)
    df = pd.read_csv(csv_data, encoding='utf-8')

    df.drop(index=clientPAN_check(df, 'CustomerGSTIN'), inplace=True)
    if df.shape[0] != 0:
        df = df.reset_index(drop=True)

    # aligning columns
    df['DocumentDate'] = pd.to_datetime(df['DocumentDate'], dayfirst=True)
    df['FY'] = df['DocumentDate'].map(
        lambda x: x.year if x.month > 3 else x.year-1).map(str)
    df['Differential %'] = ''
    df['IRN Gen-Date'] = ''

    duplicate = df[df.duplicated()]
    duplicate['Reason'] = "Duplicate"
    df.drop_duplicates(keep='first', inplace=True)
    duplicate.to_csv("GAPS\GSTR2B - Duplicate Drops.csv", index=False)
    df = df.groupby(['SupplierGSTIN', 'DocumentNo', 'FY', 'GSTR-2B category', 'Supply Type'], as_index=False).agg({'TaxPeriod': 'first', 'CustomerGSTIN': 'first', 'Supply Type': 'first', 'DocumentNo': 'first', 'DocumentDate': 'first', 'Original Invoice Number': 'first', 'Original Invoice Date': 'first', 'SupplierGSTIN': 'first', 'POS': 'first',
    'Taxable Value': 'sum', 'Tax Rate': 'sum', 'IGST Amount': 'sum', 'Central Tax Amount': 'sum', 'State/UT Tax Amount': 'sum', 'CessAmount': 'sum', 'InvoiceValue': 'sum', 'ReverseCharge': 'first', 'Differential %': 'first', 'Filing Date': 'first', 'Filing Period': 'first', 'IRN Gen-Date': 'first', 'GSTR-2B category': 'first'})
    df.drop(columns=['FY'], inplace=True)
    df = df.reindex()
    
    df = df[['TaxPeriod', 'SupplierGSTIN', 'Supply Type', 'DocumentNo', 'DocumentDate', 
    'Original Invoice Number', 'Original Invoice Date', 
    'CustomerGSTIN', 'POS', 'Taxable Value', 'Tax Rate', 'IGST Amount',
    'Central Tax Amount', 'State/UT Tax Amount', 'CessAmount',
    'InvoiceValue', 'ReverseCharge', 'Differential %', 'Filing Date', 'Filing Period', 'IRN Gen-Date', 'GSTR-2B category']]

    neworder = ['TaxPeriod', 'CustomerGSTIN', 'Supply Type', 'DocumentNo', 'DocumentDate', 'Original Invoice Number', 'Original Invoice Date', 'SupplierGSTIN', 'POS', 'Taxable Value', 'Tax Rate', 'IGST Amount', 'Central Tax Amount', 
    'State/UT Tax Amount', 'CessAmount', 'InvoiceValue', 'ReverseCharge', 'Differential %', 'Filing Date', 'Filing Period', 'IRN Gen-Date', 'GSTR-2B category']
    df.columns = neworder
    # # print(csv_data)

    df['Taxable Value'] = np.where(np.logical_and(
                df['Supply Type'] == "cdn", df['Taxable Value'] > 0), df['Taxable Value']*-1, df['Taxable Value'])
    df['Central Tax Amount'] = np.where(np.logical_and(
        df['Supply Type'] == "cdn", df['Central Tax Amount'] > 0), df['Central Tax Amount']*-1, df['Central Tax Amount'])
    df['State/UT Tax Amount'] = np.where(np.logical_and(df['Supply Type'] == "cdn",
                                            df['State/UT Tax Amount'] > 0), df['State/UT Tax Amount']*-1, df['State/UT Tax Amount'])
    df['IGST Amount'] = np.where(np.logical_and(
        df['Supply Type'] == "cdn", df['IGST Amount'] > 0), df['IGST Amount']*-1, df['IGST Amount'])
    df['CessAmount'] = np.where(np.logical_and(
        df['Supply Type'] == "cdn", df['CessAmount'] > 0), df['CessAmount']*-1, df['CessAmount'])

    #df['Original Invoice Date']= pd.to_datetime(df['Original Invoice Date'],dayfirst=True)
    df['DocumentNo'] = df['DocumentNo'].map(str).str.upper()
    df['Original Invoice Number'] = df['Original Invoice Number'].map(str).str.upper()

    try:
        df['DocumentNo'] = df['DocumentNo'].str.lstrip('0')
    except:
        pass

    try:
        df['Original Invoice Number'] = df['Original Invoice Number'].str.lstrip(
            '0')
    except:
        pass
    df['FY'] = df['DocumentDate'].map(
        lambda x: x.year if x.month > 3 else x.year-1)

    df['Tax'] = df['State/UT Tax Amount'] + \
        df['Central Tax Amount']+df['IGST Amount'] + df['CessAmount']
    df['SourceFile'] = str(filename1)
    df['UploadTime'] = datetime.now()
    df['DocKey'] = df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
        str)+df['FY'].map(str)+df['Supply Type'].map(str)+df['GSTR-2B category'].map(str)
    totalcount = len(df)
    try:
        drop = df[df['DocKey'].isin(doclist['DocKey'])]
        drop.to_sql("GSTR_2Bdropped", panwisedb,
                    if_exists="append", index=False)
        df.drop(df[df['DocKey'].isin(
            doclist['DocKey'])].index, inplace=True)
    except:
        pass
    processedcount = len(df)
    droppedcount = totalcount - processedcount

    # # print(df)
    duplicate = df[df.duplicated()]
    duplicate['Reason'] = "Duplicate"
    df.drop_duplicates(keep='first', inplace=True)
    #cdf1['FY'] = pd.to_datetime(cdf1['DocumentDate'],dayfirst=True,errors='coerce')
    #cdf1['FY'] = np.where(cdf1['FY'] != "NaT", cdf1['FY'].map(str).str[0:4],"")

    duplicate1 = df[df.duplicated(subset=[
                                    'SupplierGSTIN', 'DocumentNo', 'FY', 'GSTR-2B category', 'Supply Type'])]
    duplicate1['Reason'] = "Duplicate"
    df.drop_duplicates(subset=['SupplierGSTIN', 'DocumentNo', 'FY',
                        'GSTR-2B category', 'Supply Type'], keep='first', inplace=True)

    df.drop(df[np.logical_or(df['GSTR-2B category'] == "impg",
            df['GSTR-2B category'] == "tds")].index, inplace=True)
    df.drop(df[df['GSTR-2B category'] == "tdsa"].index, inplace=True)
    df.reset_index(inplace=True)
    # print(df.columns)
    df['cdnpos'] = df['DocumentNo'].map(
        str) + df['SupplierGSTIN'].map(str) + df['DocumentDate'].map(str)
    cdnpos = df.groupby('cdnpos', as_index=False).agg(
        {'POS': 'first', 'GSTR-2B category': 'first'})
    cdnpos = cdnpos[cdnpos['GSTR-2B category'] == "b2b"]
    df['cdnpos'] = df['Original Invoice Number'].map(
        str) + df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)

    df = pd.merge(df, cdnpos, on='cdnpos', how='left')
    # # print(cdf1.columns)
    df['POS_x'] = np.where(df['POS_x'].map(
        str) == "nan", df['POS_y'], df['POS_x'])
    df = df.rename(
        columns={'POS_x': 'POS', 'GSTR-2B category_x': 'GSTR-2B category'})
    df.drop(columns=['POS_y', 'GSTR-2B category_y'], inplace=True)
    #cdf1.drop(cdf1[~cdf1['CustomerGSTIN'].isin(gstins)].index, inplace=True)

    # df['Dropcfs'] = ""
    # df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2B category'] == "b2ba", df['CFS'] == "Y"), df['Original Invoice Number'].map(
    #     str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2B category'].str[0:3]+df['Supply Type'].map(str), df['Dropcfs'])
    # df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2B category'] == "cdna", df['CFS'] == "Y"), df['Original Invoice Number'].map(
    #     str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2B category'].str[0:3]+df['Supply Type'].map(str), df['Dropcfs'])
    # df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2B category'] == "cdna", df['CFS'] == "Y"), np.where(df['Dropcfs'] == "", df['DocumentNo'].map(
    #     str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category']+df['Supply Type'].map(str), df['Dropcfs']), df['Dropcfs'])
    # cfsdrop = df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category'] +
    #                             df['Supply Type'].map(str)).isin(df['Dropcfs']), np.logical_and(df['Dropcfs'] == "", df['Supply Type'].map(str) == "R"))]
    # cfsdrop['Reason'] = "Amended"
    # df.drop(df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category'] +
    #         df['Supply Type'].map(str)).isin(df['Dropcfs']), np.logical_and(df['Dropcfs'] == "", df['Supply Type'].map(str) == "R"))].index, inplace=True)
    # duplicate1 = duplicate.append(cfsdrop)

    # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
    #cdf1['Dropcfs'] = np.where(np.logical_and(cdf1['GSTR-2B category']=="cdna",cdf1['Dropcfs'] ==""),np.where(cdf1['CFS']=="Y",cdf1['DocumentNo'].map(str)+cdf1['SupplierGSTIN'].map(str)+cdf1['DocumentDate'].map(str)+cdf1['GSTR-2B category'].str[0:2],cdf1['Dropcfs'] ),cdf1['Dropcfs'] )

    #dropcfsn = df[df['CFS']=="N"]
    #df.drop(df[df['CFS']=="N"].index, inplace=True)
    #dropcfsn['Reason']="Not Filed"

    # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)

    # duplicate2=duplicate1.append(dropcfsn)
    # duplicate1.to_csv("PR2A\Dropped 2A.csv", index=False)
    duplicate.to_csv("PR2A\Dropped 2B.csv", index=False)
    df['DocumentNo'] = df['DocumentNo'].astype(
        str).str.replace('\.0', '', regex=True)
    # cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
    #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)
    # # print(cdf1.dtypes)

    df.reset_index(inplace=True)
    df['first1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x).start(
    ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
    df['last1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x[::-1]).start(
    ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
    #df['last1']= len(df['DocumentNo']) - df['last1']
    # # print(df['first1'])
    df['first1'] = df['first1'].fillna(
        df['DocumentNo'].str.len()).astype(int)
    df['last1'] = df['last1'].fillna(
        df['DocumentNo'].str.len()).astype(int)
    df['DocNoBfrSpl'] = [DocumentNo[:first1]
                            for DocumentNo, first1 in zip(df.DocumentNo, df.first1)]
    df['DocNoAftrSpl'] = [DocumentNo[-last1:]
                            for DocumentNo, last1 in zip(df.DocumentNo, df.last1)]
    df['DocNoWOSplChar'] = df['DocumentNo'].str.replace(
        '[^a-zA-Z0-9]', '', regex=True)
    df['DocNoNumeric'] = df['DocumentNo'].str.replace(
        r"[a-zA-Z]", '').replace('\W', '').replace('_', '').replace('/', '').replace("\\", '')
    # df['DocNoBfrSpl'] =
    df['DocNo'] = df['DocumentNo']
    # df['DocNoAftrSpl'] =
    df['SuppGSTIN'] = df['SupplierGSTIN']
    df['SuppPAN'] = df['SupplierGSTIN'].str[2:12].map(str)
    df['ResGSTIN'] = df['CustomerGSTIN']
    df['ResPAN'] = df['CustomerGSTIN'].str[2:12].map(str)
    df['DocDate'] = df['DocumentDate'].apply(
        lambda x: x.strftime('%d%m%Y') if x else "")
    df['MMYYYY'] = df['DocumentDate'].apply(
        lambda x: x.strftime('%m%Y') if x else "")
    #df['FY'] = pd.to_datetime(df['DocumentDate'],dayfirst=True).apply(lambda x: x.strftime('%Y') if x else "")
    df['TaxVal'] = df['Taxable Value'].round(2)
    df['InvVal'] = df['InvoiceValue'].round(2)
    df['CGST'] = df['Central Tax Amount'].round(2)
    df['SGST'] = df['State/UT Tax Amount'].round(2)
    df['IGST'] = df['IGST Amount'].round(2)
    df['GST'] = df['Central Tax Amount'].round(2) + df['State/UT Tax Amount'].round(
        2) + df['IGST Amount'].round(2) + df['CessAmount'].round(2)

    pd.to_numeric(df['POS'])
    df['POS'] = df['POS'].map(str)
    df['POS'] = df["POS"].apply(lambda x: x.replace(r'.0', '') if isinstance(x, str) else x)
    # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                
    df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
    # print(df['POS'].unique())

    df['RCM'] = df['ReverseCharge']
    df['DocType'] = np.where(df['Supply Type'] == "C", "cdn", np.where(
        df['Supply Type'] == "D", "dbn", "b2b"))
    df['NA'] = ""

    df['Key1'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
        str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key2'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
        str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key3'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key4'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key5'] = df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key6'] = df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key7'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(str)+df['Central Tax Amount'].round(
        2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key8'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
        str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    df['Key9'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
        str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
    df['Key10'] = df['FY'].map(str) + df['SuppPAN'].map(str) + df['DocNoWOSplChar'].map(
        str)+df['ResPAN'].map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)

    df['Key11'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key12'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
        str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key13'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key14'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key15'] = df['SupplierGSTIN'].map(
        str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key16'] = df['SupplierGSTIN'].map(
        str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
    df['Key19'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(
        str)+df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)
    df['Key20'] = df['FY'].map(str) + df['SuppPAN'].map(
        str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocType'].map(str)

    df['Key17'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(
        str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
    df['Key18'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(
        str)+df['DocDate'].map(str)+df['DocType'].map(str)
    # print("Saved 2B Default Keys")
    try:
        query = "select * FROM matchconfig"
        config = pd.read_sql_query(query, panwisedb)
        rowdata = dataframe_to_rows(config, index=False, header=False)
        rowdata1 = dataframe_to_rows(config, index=False, header=False)
        rowdata2 = dataframe_to_rows(config, index=False, header=False)
        matchcon = 1
    except:
        matchcon = 0
        config = []
    l = len(config)
    n = 11

    try:
        for row1 in rowdata1:

            key = "Key" + str(n)
            df[key] = df[row1[1]].map(str) + df[row1[2]].map(str) + df[row1[3]].map(str) + df[row1[4]].map(
                str) + df[row1[5]].map(str) + df[row1[6]].map(str) + df[row1[7]].map(str) + df[row1[8]].map(str)
            ## print("prkey ",row[0])
            n += 1
    except:
        pass
    # print("Saved User Defined Keys")
    df.to_sql("GSTR_2B", panwisedb, if_exists="append", index=False)

    # tcdf1=tcdf.copy()
    #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)

    # df.to_csv(r"GAPS\GSTR2b.csv", index=False)

    
    # df['Original Invoice Date'] = pd.to_datetime(
    #     df['Original Invoice Date'], dayfirst=True)
    # try:
    #     df['DocumentNo'] = df['DocumentNo'].str.lstrip('0')
    # except:
    #     pass

    # try:
    #     df['Original Invoice Number'] = df['Original Invoice Number'].str.lstrip(
    #         '0')
    # except:
    #     pass

    # df['Tax'] = df['State/UT Tax Amount'] + \
    #     df['Central Tax Amount']+df['IGST Amount']
    # df['SourceFile'] = str(filename1)
    # df['UploadTime'] = datetime.now()
    # df['DocKey'] = df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
    #     str)+df['FY'].map(str)+df['Supply Type'].map(str)+df['GSTR-2B category'].map(str)
    # totalcount = len(df)
    # try:
    #     drop = df[df['DocKey'].isin(doclist['DocKey'])]
    #     drop.to_sql("GSTR_2Bdropped", panwisedb,
    #                 if_exists="append", index=False)
    #     df.drop(df[df['DocKey'].isin(doclist['DocKey'])].index, inplace=True)
    # except:
    #     pass
    # processedcount = len(df)
    # droppedcount = totalcount - processedcount

    # df.to_csv(r"GAPS\GSTR2b.csv", index=False)
    # df.to_sql("GSTR_2B", panwisedb, if_exists="append", index=False)
    return totalcount, processedcount, droppedcount

# def process_file(lst, filedata, clientPAN, current_user):
def process_file(is_local, parameters):
    # is_local true if localhost else false if Azure Blob Storage
    # parameters will be clientPAN and current_user if localhost else temp_db_name
    dir_path = os.path.dirname(os.path.realpath(__file__))
    folder_name = '' if is_local else parameters[3]
    path = dir_path + '/Client-Details' if is_local else dir_path + folder_name

    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]
    # filedata = parameters[2] if is_local else parameters[1] # file path if is_local else dataframe
    filedata = parameters[2] if is_local else parameters[2] # file path if is_local else sas_url

    panwisedb_path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'

    # path = dir_path + '/Client-Details'
    upldtype = ""

    lst = ["False", "False", "False", "False", "False", "False", "False"]

    # folder_name = 'Duplicates dropped'
    try:
        os.mkdir(dir_path + '\\' + 'PR2A')
        os.mkdir(dir_path +'\\'+ 'PR2A\\'+ 'Duplicates dropped')
    except:
        pass
    try:
        os.mkdir(dir_path + '\\' + 'PR2B')
        os.mkdir(dir_path +'\\'+ 'PR2B\\'+ 'Duplicates dropped')
    except:
        pass

    # from io import BytesIO
    # filedata = BytesIO(content.encode())
    filename1 = filedata.split('/')[-1] if is_local else parameters[1]
    filetype = os.path.splitext(filename1)[1][1:].lower()

    # panwisedb = sqlite3.connect(f'{path}/{current_user}/{clientPAN}/{clientPAN}.db', timeout=10)
    panwisedb = sqlite3.connect(panwisedb_path, timeout=10)
    # try:
    #     if filetype != 'pdf':
    #         filename = filename.read().decode("utf-8")
    # except:
    #     pass

    srcol59 = ['SourceIdentifier', 'SourceFileName', 'GLAccountCode', 'Division', 'SubDivision', 'ProfitCentre1', 'ProfitCentre2', 'PlantCode', 'ReturnPeriod', 'SupplierGSTIN', 'DocumentType', 'SupplyType', 'DocumentNumber', 'DocumentDate', 'OriginalDocumentNumber', 'OriginalDocumentDate', 'CRDRPreGST', 'LineNumber', 'CustomerGSTIN', 'UINorComposition', 'OriginalCustomerGSTIN', 'CustomerName', 'CustomerCode', 'BillToState', 'ShipToState', 'POS', 'PortCode', 'ShippingBillNumber', 'ShippingBillDate', 'FOB', 'ExportDuty', 'HSNorSAC',
            'ProductCode', 'ProductDescription', 'CategoryOfProduct', 'UnitOfMeasurement', 'Quantity', 'TaxableValue', 'IntegratedTaxRate', 'IntegratedTaxAmount', 'CentralTaxRate', 'CentralTaxAmount', 'StateUTTaxRate', 'StateUTTaxAmount', 'CessRateAdvalorem', 'CessAmountAdvalorem', 'CessRateSpecific', 'CessAmountSpecific', 'InvoiceValue', 'ReverseChargeFlag', 'TCSFlag', 'eComGSTIN', 'ITCFlag', 'ReasonForCreditDebitNote', 'AccountingVoucherNumber', 'AccountingVoucherDate', 'Userdefinedfield1', 'Userdefinedfield2', 'Userdefinedfield3']
    glazure = ['Supplier GSTIN','Assignment','Document Number','Document Type','Document Date','Posting Date','Posting Key','Amount in local currency','Local Currency','Clearing Document','Cost Center','Business Unit','Segment','Reference','Invoice reference','Account Type','Account','Customer','Vendor','Name 1','Tax code','Profit Center','Text','Clearing date','Plant','Reference Key 1','Reference Key 3','Payment reference','Reference Transact.','Reference Date','Reference Key','Reference Key 2','Purchasing Document','Order','Year/month','Sales Document Item','Trading Partner','Payment date','Baseline Payment Dte','Terms of Payment','Credit Control Area','Clearing item','Transaction Type','Entry Date']
    # srazure = ['SourceIdentifier','SourceFileName','GLAccountCode','Division','SubDivision','ProfitCentre1','ProfitCentre2','PlantCode','ReturnPeriod','SupplierGSTIN','DocumentType','SupplyType','DocumentNumber','DocumentDate','OriginalDocumentNumber','OriginalDocumentDate','CRDRPreGST','LineNumber','CustomerGSTIN','UINorComposition','OriginalCustomerGSTIN','CustomerName','CustomerCode','BillToState','ShipToState','POS','PortCode','ShippingBillNumber','ShippingBillDate','FOB','ExportDuty','HSNorSAC','ProductCode','ProductDescription','CategoryOfProduct','UnitOfMeasurement','Quantity','TaxableValue','IntegratedTaxRate','IntegratedTaxAmount','CentralTaxRate','CentralTaxAmount','StateUTTaxRate','StateUTTaxAmount','CessRateAdvalorem','CessAmountAdvalorem','CessRateSpecific','CessAmountSpecific','InvoiceValue','ReverseChargeFlag','TCSFlag','eComGSTIN','ITCFlag','ReasonForCreditDebitNote','AccountingVoucherNumber','AccountingVoucherDate','Userdefinedfield1','Userdefinedfield2','Userdefinedfield3']
    srazure = ['SourceIdentifier', 'SourceFileName', 'GLAccountCode', 'Division','SubDivision', 'ProfitCentre1', 'ProfitCentre2', 'PlantCode','ReturnPeriod', 'SupplierGSTIN', 'DocumentType', 'SupplyType','DocumentNumber', 'DocumentDate', 'OriginalDocumentNumber','OriginalDocumentDate', 'CRDRPreGST', 'LineNumber', 'CustomerGSTIN','UINorComposition', 'OriginalCustomerGSTIN', 'CustomerName','CustomerCode', 'BillToState', 'ShipToState', 'POS', 'PortCode','ShippingBillNumber', 'ShippingBillDate', 'FOB', 'ExportDuty','HSNorSAC', 'Product Code', 'Product Description','Category of Product', 'UnitOfMeasurement', 'Quantity', 'TaxableValue','IntegratedTaxRate', 'IntegratedTaxAmount', 'CentralTaxRate','CentralTaxAmount', 'StateUTTaxRate', 'StateUTTaxAmount','CessRateAdvalorem', 'CessAmountAdvalorem', 'CessRateSpecific','CessAmountSpecific', 'InvoiceValue', 'ReverseChargeFlag', 'TCSFLAG','ECOMMGSTIN', 'ITCFLAG', 'ReasonForCreditDebitNote','AccountingVoucherNumber', 'AccountingVoucherDate', 'Userdefinedfield1','Userdefinedfield2', 'Userdefinedfield3']
    srsap = ['UOM_Reg','Original_Invoice_Date_Reg','Shipping_Bill_Date_Reg','Posting_Date_Reg','Accounting_Document_Number_Reg','CGST_Rate_Reg','SGST_Rate_Reg','IGST_Rate_Reg','Cess_Rate_Reg','Tax_Rate_Reg','CGST_Amount_Reg','SGST_Amount_Reg','IGST_Amount_Reg','Cess_Amount_Reg','Tax_Amount_Reg','Taxable_Amount_Reg','Supplier_GSTIN_Reg','Customer_GSTIN_Reg','Document_Type_Reg','Supply_Type_Reg','POS_Reg','Qty_Reg','Invoice_No_Reg','Invoice_Date_Reg','HSN_Reg','Customer_Name_Reg','Description_Reg']
    prcol65 = ['SourceIdentifier', 'SourceFileName', 'GLAccountCode', 'Division', 'SubDivision', 'ProfitCentre1', 'ProfitCentre2', 'PlantCode', 'ReturnPeriod', 'RecipientGSTIN', 'DocumentType', 'SupplyType', 'DocumentNumber', 'DocumentDate', 'OriginalDocumentNumber', 'OriginalDocumentDate', 'CRDRPreGST', 'LineNumber', 'SupplierGSTIN', 'OriginalSupplierGSTIN', 'SupplierName', 'SupplierCode', 'POS', 'PortCode', 'BillOfEntry', 'BillOfEntryDate', 'CIFValue', 'CustomDuty', 'HSNorSAC', 'ItemCode', 'ItemDescription', 'CategoryOfItem', 'UnitOfMeasurement', 'Quantity', 'TaxableValue', 'IntegratedTaxRate',
            'IntegratedTaxAmount', 'CentralTaxRate', 'CentralTaxAmount', 'StateUTTaxRate', 'StateUTTaxAmount', 'CessRateAdvalorem', 'CessAmountAdvalorem', 'CessRateSpecific', 'CessAmountSpecific', 'InvoiceValue', 'ReverseChargeFlag', 'EligibilityIndicator', 'CommonSupplyIndicator', 'AvailableIGST', 'AvailableCGST', 'AvailableSGST', 'AvailableCess', 'ITCReversalIdentifier', 'ReasonForCreditDebitNote', 'PurchaseVoucherNumber', 'PurchaseVoucherDate', 'PaymentVoucherNumber', 'PaymentDate', 'ContractNumber', 'ContractDate', 'ContractValue', 'Userdefinedfield1', 'Userdefinedfield2', 'Userdefinedfield3']
    digi2a = ['IRN', 'Counter Party Return Status', 'Return Period', 'Recipent GSTIN', 'Document Type', 'Document Number', 'Document Date', 'Original Document Number', 'Original Document Date', 'Invoice Number', 'Invoice Date', 'CR/DR Pre-GST', 'Line Number', 'Supplier GSTIN', 'Supplier Name', 'POS', 'State Name', 'ITC Eligible', 'Taxable Value ', 'Tax Rate',
            'Integrated Tax Amount', 'Central Tax Amount', 'StateUT TaxAmount', 'Cess Amount', 'Invoice Value', 'Reverse Charge Flag', 'Differential Percentage', 'Record Type', 'DeLinking Flag', 'CFS_GSTR3B', 'CancellationDt', 'GSTR1_FilingDt', 'GSTR1_FilingPeriod', 'OrgInvAmendmentPeriod', 'OrgInvAmendmentType', 'ReferenceDt', 'PortCode', 'Source Type', 'Generation Date']
    digi2b = ['Return Period', 'Recipient GSTIN', 'Supplier GSTIN', 'Supplier Name',
    'Document Type', 'Supply Type', 'Document Number', 'Document Date',
    'Taxable Value ', 'Tax Rate', 'IGST Amount', 'CGST Amount',
    'SGST Amount', 'CESS Amount', 'Invoice Value', 'POS', 'State Name',
    'Line Number', 'BOE-ReferenceDate(ICEGATE)', 'BOE-Received Date(GSTN)',
    'PortCode', 'Bill Of Entry Number', 'Bill Of Entry Date',
    'BOE-Amendment Flag', 'Original Document Number',
    'Original Document Date', 'Original Document Type',
    'Original Invoice Number', 'Original Invoice Date', '2B-GenerationDate',
    'GSTR-1/5/6 Filing Period', 'GSTR-1/5/6 Filing Date',
    'Differential Percentage', 'Reverse Charge Flag', 'ITC Availability',
    'Reason for ITC Unavailability', 'Source Type', 'Generation Date',
    'IRN']
    prshort = ['Sl.No', ' Name of Vendor ', 'Supplier GSTIN', 'HSN ', 'Return Period', 'Availment period', 'Posting Date ', 'DocumentNumber', 'Document Date ', 'Document amount  ', 'Taxable Value ', ' CGST ',
            ' SGST ', ' IGST ', 'Cess', ' Total GST ', 'POS', 'Recepient GSTIN', 'Document Type', 'Eligibility Indicator', 'Original Document Number', 'Original Document Date', 'RCM', 'EY1', 'EY2', 'EY3', 'EY4', 'EY5']
    gstinmaster = ['BP_GL', 'GSTIN']
    doctypemaster = ['Document_Type', 'Document_Type_MS']
    supplytypemaster = ['Supply_Type_Reg', 'Supply_Type_MS']
    glcodemaster = ['CGST_Output','SGST_Output','IGST_Output','UGST_Output','Compensation Cess GL Code','Kerala Cess GL Code','Revenue GLs','Expense GL','Forex Gls Part of revenue','Taxable Advance (Liability) GLs','Non-Taxable Advance (Liability) GLs','Cross-Charge and Stock transfer GLs','Unbilled Revenue GLs','Bank A/C GLs','CGST_Input','SGST_Input','IGST_Input','UGST_Input']
    glcodemastersap = ['CGST Tax GL Codes','SGST Tax GL Code','IGST Tax GL Code','UGST Tax GL Code','Compensation Cess GL Code','Kerala Cess GL Code','Revenue GLs','Forex Gls Part of revenue','Taxable Advance (Liability) GLs','Non-Taxable Advance (Liability) GLs','Cross-Charge and Stock transfer GLs','Unbilled Revenue GLs','Bank A/C GLs','Input Tax GLs','Fixed Asset GLs']
    trialbalance = ['GSTIN_TB', 'G/L_GL', 'Account_Name', 'Account_Name_Main_Heading',
                    'IND AS Sub Heading', 'Opening Balance', 'Debit', 'Credit', 'Closing Balance']
    taxcodemaster = ['Transaction_Type_GL', 'Tax_Code_Description_MS',
                    'Tax_Type_MS', 'Eligibility_MS', 'Tax_Rate_MS']
    # gldump = ['G/L_GL', 'Reference_GL', 'Accounting_Document_Number_GL', 'Document_Type_GL', 'Posting_Date_GL', 'Document_Date_GL', 'Clearing_Document_Number_GL', 'Amount_GL',
    #         'Year/month_GL', 'Transaction_Type_GL', 'BP_GL', 'Text_GL', 'Customer_Code_GL', 'Plant_GL', 'Company_Code_GL', 'Entry_Date_GL', 'Period_GL', 'Vendor_Code_GL', 'Offset_Account_GL']
    gldump = ['G/L_GL', 'Reference_GL', 'Accounting_Document_Number_GL', 'Document_Type_GL', 'Posting_Date_GL', 'Document_Date_GL', 'Clearing_Document_Number_GL', 'Amount_GL',
            'Year/month_GL', 'Transaction_Type_GL', 'BP_GL', 'Text_GL', 'Customer_Code_GL', 'Plant_GL', 'Company_Code_GL', 'Entry_Date_GL', 'Period_GL', 'Vendor_Code_GL', 'Offset_Account_GL']
    # gldumpsap = ['Cleared/open items symbol','G/L Account','Assignment','Order','Reference','Document Number','Document Type','Posting Date','Document Date','Posting Key','Clearing Document','Clearing date','Document currency','Amount in doc. curr.','Amount in local currency','Cost Center','Year/month','Local Currency','Tax code','Profit Center','Text','Offsett.account type','Offsetting acct no.','Offset account Name',]
    firc_col = ['Supplier GSTIN', 'Invoice Type', 'Customer Name', 'Export Invoice No.',
                'Export Invoice Date', 'Taxable Value GSTR1 (FC)',
                'Export Exchange Rate', 'Export Exchange Currency', 'Goods/Services',
                'Shipping Bill No.', 'Shipping Bill Date', 'Port Code',
                'Export General Manifest No.', 'Export General Manifest Date',
                'FOB VALUE', 'FIRC No.', 'FIRC Date', 'FIRC Amount INR',
                'FIRC Amount (FC)', 'FIRC Exchange Rate', 'EY1', 'EY2']
    duedatemaster = ['TaxPeriod', 'ReturnType', 'Due date']
    returnfilingmaster = ['GSTIN', 'TaxPeriod', 'ReturnType', 'FilingDate', 'ArnNo', 'Status',
                        'Legal Name of Business', 'Trade Name', 'GSTIN Status', 'Date of registration', 'Registration type']
    
    gldump_mapping = ['Base Headers', 'Input File Headers']

    totalcount = 0
    processedcount = 0
    droppedcount = 0
    if filetype == "csv":
        headers = pd.read_csv(filedata, nrows=1, encoding='utf-8')
        columns = headers.columns
        length = len(columns)
        
        try:
            is_gl_dump = True
            df = pd.read_sql_query("select * from GLDumpMapping", panwisedb)
            df = df.transpose()
            df.columns = df.iloc[0]
            neworder= list(df)
            df.columns = df.iloc[1]
            colop = list(df)
            for i in colop:
                if i in columns:
                    pass
                elif i in [None, '']:
                    pass
                else:
                    is_gl_dump = False
                    break
                    # return "Invalid Column Headers"
            if is_gl_dump:
                columns = neworder
            else:
                pass
        except:
            pass

        if length == 65:
            # print("PR")
            if set(columns) == set(prcol65):
                upldtype = "Purchase Register - 65 Columns"
        if length == 28:
            if set(columns) == set(prshort):
                upldtype = "Purchase Register - 28 Columns"
        if length == 59:
            if set(columns) == set(srcol59):
                upldtype = "Sales Register - 59 Columns"
        if length == 39:
            if set(columns) == set(digi2a):
                upldtype = "GSTR 2A"
                # # print("2A")
    
            if set(columns) == set(digi2b):
                upldtype = "GSTR 2B"
        if set(columns) == set(gldump): # or set(columns) == set(gldumpsap):
            upldtype = "GL Dump"
        if set(columns) == set(srazure):
            upldtype = "Sales Register Azure"
        if set(columns) == set(gldump_mapping):
            try:
                df = pd.read_csv(filedata, encoding='utf-8')
                df.to_sql("GLDumpMapping", panwisedb, if_exists="replace", index=False)
            except:
                pass
            return
        if set(columns) == set(glazure):
            upldtype = "GL Dump Azure"
        if set(columns) == set(srsap):
            upldtype = "Sales Register SAP"
        if set(columns) == set(gstinmaster):
            upldtype = "GSTIN"
        if set(columns) == set(doctypemaster):
            upldtype = "Doc Type"
        if set(columns) == set(supplytypemaster):
            upldtype = "Supply Type"
        if set(columns) == set(glcodemaster) or set(columns) == set(glcodemastersap):
            upldtype = "GL Code"
        if set(columns) == set(trialbalance):
            upldtype = "TB"
        if set(columns) == set(taxcodemaster):
            upldtype = "Tax Code"
        if set(columns) == set(['Electronic Credit Ledger']):
            # print("Electronic Credit Ledger")
            upldtype = "Electronic Credit Ledger"
        if set(columns) == set(['Electronic Cash Ledger']):
            # print("Electronic Cash Ledger")
            upldtype = "Electronic Cash Ledger"
        if set(columns) == set(firc_col):
            # print("FIRC")
            upldtype = "FIRC"
        if set(columns) == set(duedatemaster):
            upldtype = "Due Date Master"
        if set(columns) == set(returnfilingmaster):
            upldtype = "Return Filing Master"

    if filetype == "xlsx":
        # if is_local:
        #     headers = open(filedata)
        #     headers = openpyxl.load_workbook(BytesIO(b64decode(headers.read())),read_only=True,data_only=True)
        #     ws = headers[headers.sheetnames[0]]
        #     data = ws.values
        #     columns = list(next(data)[0:])
        # else:
        #     headers = urllib.request.urlopen(filedata).read()
        #     headers = openpyxl.load_workbook(filename = BytesIO(headers))
        #     ws = headers[headers.sheetnames[0]]
        #     data = ws.values
        #     columns = list(next(data)[0:])

        headers = pd.read_excel(filedata, nrows=1)
        columns = headers.columns

        try:
            is_gl_dump = True
            df = pd.read_sql_query("select * from GLDumpMapping", panwisedb)
            df = df.transpose()
            df.columns = df.iloc[0]
            neworder= list(df)
            df.columns = df.iloc[1]
            colop = list(df)
            for i in colop:
                if i in columns:
                    pass
                elif i in [None, '']:
                    pass
                else:
                    is_gl_dump = False
                    break
                    # return "Invalid Column Headers"
            if is_gl_dump:
                columns = neworder
            else:
                pass
        except:
            pass
        
        # headers = pd.read_excel(filedata, nrows=1)
        # # print(headers)
        # columns = headers.columns
        length = len(columns)
        if length == 65:
            if set(columns) == set(prcol65):
                upldtype = "Purchase Register - 65 Columns"
                # # print("PR")
        if length == 28:
            if set(columns) == set(prshort):
                upldtype = "Purchase Register - 28 Columns"
                # print("PR28")
        if length == 59:
            if set(columns) == set(srcol59):
                upldtype = "Sales Register - 59 Columns"
                # print("SR")
        if length == 39:
            if set(columns) == set(digi2a):
                upldtype = "GSTR 2A"
            if set(columns) == set(digi2b):
                upldtype = "GSTR 2B"
                # # print("2A")
        if set(columns) == set(gldump): # or set(columns) == set(gldumpsap):
            upldtype = "GL Dump"
        if set(columns) == set(srazure):
            upldtype = "Sales Register Azure"
        if set(columns) == set(glazure):
            upldtype = "GL Dump Azure"
        if set(columns) == set(srsap):
            upldtype = "Sales Register SAP"
        if set(columns) == set(gldump_mapping):
            try:
                df = pd.read_excel(filedata)
                df.to_sql("GLDumpMapping", panwisedb, if_exists="replace", index=False)
            except:
                pass
            return
        if set(columns) == set(gstinmaster):
            upldtype = "GSTIN"
        if set(columns) == set(doctypemaster):
            upldtype = "Doc Type"
        if set(columns) == set(supplytypemaster):
            upldtype = "Supply Type"
        if set(columns) == set(glcodemaster) or set(columns) == set(glcodemastersap):
            upldtype = "GL Code"
        if set(columns) == set(trialbalance):
            upldtype = "TB"
        if set(columns) == set(taxcodemaster):
            upldtype = "Tax Code"
        if columns[0] == 'Electronic Credit Ledger':
            # print("Electronic Credit Ledger")
            upldtype = "Electronic Credit Ledger"

    if filetype == "zip":
        with zipfile.ZipFile(filedata) as myzip:
            # # print(myzip.namelist()[0])
            if myzip.namelist()[0][-4:] == 'json':
                with myzip.open(myzip.namelist()[0]) as myfile:
                    parsed_json = json.loads(myfile.read().decode('utf-8'))

                    if "data" in parsed_json.keys():
                        # # print("GSTR-1")
                        upldtype = "GSTR 2B"

                    elif "filing_typ" in parsed_json.keys():
                        # # print("GSTR-1")
                        upldtype = "GSTR 1"
                    else:
                        # # print("GSTR-2A")
                        upldtype = "GSTR 2A"

    if filetype == "json":
        content = open(filedata, 'r')
        content = content.read()
        parsed_json = json.loads(content)
        if "data" in parsed_json.keys():
            # # print("GSTR-1")
            upldtype = "GSTR 2B"
        elif "filing_typ" in parsed_json.keys():
            # print("GSTR-1")
            upldtype = "GSTR 1"
        elif 'refundRsn' in parsed_json and parsed_json['refundRsn'] == "EXPWOP":
            upldtype = "Statement3"
        else:
            # # print("GSTR-2A")
            upldtype = "GSTR 2A"

    if filetype == "pdf":
        # print("UPLOADED PDF FOR GSTR1/3B")
        # filename.save(os.path.join('temp/', filename1))
        doc = fitz.open(stream=filedata, filetype="pdf")
        page1 = doc.loadPage(0)
        page1text = page1.getText("text")
        pdf_document_type = page1text.split('\n')[0]
        doc.close()

        if(pdf_document_type == "Form GSTR-1"):
            upldtype = "GSTR1"

            filepath = f"{dir_path}\\temp\\{filename1}"

            import subprocess
            # JAR FILE
            subprocess.run(['java','-jar',f'{dir_path}\\tabula-1.0.5-jar-with-dependencies.jar','--format','JSON','--pages','all','--outfile',f"{dir_path}\\temp\\{filename1}.json",'--lattice',filepath], capture_output=True, text=True).stdout

            with open(f"{dir_path}\\temp\\{filename1}.json", 'r') as file:
                raw_data = file.read()
            extracted_data = json.loads(raw_data)
            def remove_additional_keys(d):
                if not isinstance(d, (dict, list)):
                    return d
                if isinstance(d, list):
                    return [remove_additional_keys(v) for v in d]
                return {k: remove_additional_keys(v) for k, v in d.items()
                        if k not in {'top', 'left', 'width','height','right', 'bottom' }}
            filtered_data = [remove_additional_keys(sample['data']) for sample in extracted_data]
            final_data = []
            for data_item in filtered_data:
                one_item = []
                for i in data_item:
                    one_item.append([f['text'] for f in i])
                if len(one_item) > 0:
                    df = pd.DataFrame(one_item[1:], columns=one_item[0])
                else:
                    continue
                final_data.append(df)
            
            masterDF = pd.DataFrame(columns=['GSTIN#', 'Year', 'Month', 'Description', 'NumberOfRecords','TotalInvoiceValue','Taxable', 'IGST', 'CGST', 'SGST', 'CESS'])

            description = [
                    "year","GSTIN","B2B Invoices","B2C Invoices","Credit / Debit Notes (Registered)",
                    "Credit / Debit Notes (Registered) data","Credit / Debit Notes (Unregistered)","Exports Invoices",
                    "B2C (Others)","Nil rated, exempted and non GST outward supplies","Tax Liability (Advances Received)",
                    "Adjustment of Advances","HSN-wise summary of outward supplies"," Documents Issued","Amended B2B Invoices", 
                    "Amended B2C (Large) Invoices","Amended Credit/Debit Notes (Registered)","Amended Credit/Debit Notes (Unregistered)",
                    "Amended Exports Invoices","Amended B2C(Others)",
                    "Amended Tax Liability (Advance Received)"," Amendment of Adjustment of Advances"
                ]

            count = 0
            for df in final_data:
                row = df.columns.to_list()
                row = pd.to_numeric(row, errors='ignore', downcast='integer')
                if row[0] != "No. of Records" and len(row) == 7:
                    additionaldf = pd.DataFrame(columns=['No. of Records', 'Total Invoice value', 'Total Taxable value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT  Tax', 'Total Cess'], data=[row])
                    final_data[count-1] = additionaldf
                count += 1
            date_col = final_data[0].columns.values.tolist()
            gstin_col = final_data[1].columns.values.tolist()

            for i in range(2, len(final_data)):
                cols = final_data[i].columns.values.tolist()
                data = {
                    'GSTIN#': gstin_col[1],
                    'Year': date_col[1],
                    'Month': final_data[0][date_col[1]].to_string(index=False),
                    'Description': description[i]
                }
                # print(data)

                if 'No. of Records' in cols:
                    data['NumberOfRecords'] = pd.to_numeric(final_data[i]['No. of Records'].to_string(index=False), errors='ignore', downcast="integer")
                else:
                    data['NumberOfRecords'] = None

                if 'Total Invoice value' in cols:
                    data['TotalInvoiceValue'] = pd.to_numeric(final_data[i]['Total Invoice value'].to_string(index=False), errors='ignore')
                else:
                    data['TotalInvoiceValue'] = None

                if 'Total Taxable value' in cols:
                    data['Taxable'] = pd.to_numeric(final_data[i]['Total Taxable value'].to_string(index=False), errors='ignore')
                else:
                    data['Taxable'] = None

                if 'Total Integrated Tax' in cols:
                    data['IGST'] = pd.to_numeric(final_data[i]['Total Integrated Tax'].to_string(index=False), errors='ignore')
                else:
                    data['IGST'] = None

                if 'Total Central Tax' in cols:
                    data['CGST'] = pd.to_numeric(final_data[i]['Total Central Tax'].to_string(index=False), errors='ignore')
                else:
                    data['CGST'] = None

                if 'Total State/UT  Tax' in cols:
                    data['SGST'] = pd.to_numeric(final_data[i]['Total State/UT  Tax'].to_string(index=False), errors='ignore')
                else:
                    data['SGST'] = None

                if 'Total Cess' in cols:
                    data['CESS'] = pd.to_numeric(final_data[i]['Total Cess'].to_string(index=False), errors='ignore')
                else:
                    data['CESS'] = None

                masterDF = masterDF.append(data, ignore_index=True)
            masterDF = masterDF.replace('NaN', np.nan)
            # masterDF.drop(masterDF[masterDF['Description'] == 'Credit / Debit Notes (Registered) data'].index, inplace=True)
            masterDF.to_sql("GSTR1",panwisedb,if_exists="append",index=False)
            panwisedb.commit()
            if(os.path.exists(dir_path + "/temp/" + filename1)):
                os.remove(dir_path + "/temp/" + filename1)
            if(os.path.exists(dir_path + "/temp/" + filename1 + ".json")):
                os.remove(dir_path + "/temp/" + filename1 + ".json")

        elif (pdf_document_type == "Form GSTR-3B"):
            upldtype = "GSTR3B"
            
            filepath = f"{dir_path}\\temp\\{filename1}"

            import subprocess
            subprocess.run(['java','-jar',f'{dir_path}\\tabula-1.0.5-jar-with-dependencies.jar','--format','JSON','--pages','all','--outfile',f"{dir_path}\\temp\\{filename1}.json",'--lattice',filepath], capture_output=True, text=True).stdout

            with open(f"{dir_path}\\temp\\{filename1}.json", 'r') as file:
                raw_data = file.read()
            extracted_data = json.loads(raw_data)
            def remove_additional_keys(d):
                if not isinstance(d, (dict, list)):
                    return d
                if isinstance(d, list):
                    return [remove_additional_keys(v) for v in d]
                return {k: remove_additional_keys(v) for k, v in d.items()
                        if k not in {'top', 'left', 'width','height','right', 'bottom' }}
            filtered_data = [remove_additional_keys(sample['data']) for sample in extracted_data]
            final_data = []
            for data_item in filtered_data:
                one_item = []
                for i in data_item:
                    one_item.append([f['text'] for f in i])
                if len(one_item) > 0:
                    df = pd.DataFrame(one_item[1:], columns=one_item[0])
                else:
                    df = pd.DataFrame()
                if df.empty:
                    continue
                final_data.append(df)
            masterDF = pd.DataFrame(columns=['GSTIN#', 'Year', 'Month', 'Description', 'Taxable', 'IGST', 'CGST', 'SGST', 'CESS', 'Cash','InterStateSupplies','IntraStateSupplies','Interest', 'Late', 'GST3B-Tag'])

            description = ["year","GSTIN","3.1 ","3.2 ","4. ","5 ","5.1 ","6.1 ", "6.2 "]

            gstin_col = final_data[1].columns.values.tolist()
            date_col = final_data[0].columns.values.tolist()

            for i in range(2, len(final_data)):

                cols = final_data[i].columns.values.tolist()
                final_data[i] = final_data[i].replace({"": np.nan}).dropna(how='all')
                
                if ('Nature of Supplies' not in cols or 'Details' not in cols) and ('Unnamed: 1' in cols or '5  Values of exempt, nil-rated and non-GST inward supplies' in cols or '5.1 Interest and Late fee' in cols):
                    header_row = 0
                    final_data[i].columns = final_data[i].iloc[header_row]
                    final_data[i] = final_data[i][1:]
                if '6.1 Payment of tax' in cols:
                    final_data[i].columns = ['Description','Total tax payable','Integrated Tax','Central Tax','State/UT Tax','Cess','Cash','Interest','Late']
                    final_data[i] = final_data[i].iloc[2:,:]
                if 'Tax paid through ITC' in cols:
                    final_data[i].columns = ['Description','Total tax payable','Integrated Tax','Central Tax','State/UT Tax','Cess','Cash','Interest','Late']
                    final_data[i] = final_data[i].iloc[1:,:]
                
                final_data[i].rename(columns=lambda x: x.replace('\r', ' ').replace('(?)', '').strip(), inplace=True)
                cols = final_data[i].columns.values.tolist()

                for index, row in final_data[i].iterrows():
                    data = {
                        'GSTIN#': gstin_col[1],
                        'Year': date_col[1],
                        'Month': final_data[0][date_col[1]].to_string(index=False),
                    }

                    if 'Nature of Supplies' in cols:
                        print(i)
                        data['Description'] = description[i] + row['Nature of Supplies']
                    
                    elif 'Details' in cols:
                        data['Description'] = description[i] + row['Details']
                        
                    elif 'Description' in cols:
                        data['Description'] = description[i] + str(row['Description'])

                    if 'Total Taxable Value' in cols:
                        row['Total Taxable Value'] = pd.to_numeric(row['Total Taxable Value'], errors='coerce')
                        data['Taxable'] = row['Total Taxable Value']

                    if 'Integrated Tax' in cols:
                        row['Integrated Tax'] = pd.to_numeric(row['Integrated Tax'], errors='coerce')
                        data['IGST'] = row['Integrated Tax']

                    if 'Central Tax' in cols:
                        row['Central Tax'] = pd.to_numeric(row['Central Tax'], errors='coerce')
                        data['CGST'] = row['Central Tax']
                    
                    if 'State/UT Tax' in cols:
                        row['State/UT Tax'] = pd.to_numeric(row['State/UT Tax'], errors='coerce')
                        data['SGST'] = row['State/UT Tax']

                    if 'Cess' in cols:
                        row['Cess'] = pd.to_numeric(row['Cess'], errors='coerce')
                        data['CESS'] = row['Cess']
                    
                    if 'Inter- State supplies' in cols:
                        row['Inter- State supplies'] = pd.to_numeric(row['Inter- State supplies'], errors='coerce')
                        data['InterStateSupplies'] = row['Inter- State supplies']
                    
                    if 'Intra- State supplies' in cols:
                        row['Intra- State supplies'] = pd.to_numeric(row['Intra- State supplies'], errors='coerce')
                        data['IntraStateSupplies'] = row['Intra- State supplies']
                        
                    if 'Cash' in cols:
                        data['Cash'] = row['Cash']
                    
                    if 'Interest' in cols:
                        data['Interest'] = row['Interest']
                    
                    if 'Late' in cols:
                        data['Late'] = row['Late']

                    masterDF = masterDF.append(data, ignore_index=True)

            masterDF['Description'] = masterDF['Description'].str.replace('\r', ' ')
            masterDF['Month'] = masterDF['Month'].str.replace(' ', '')
            masterDF.to_sql("GSTR3B", panwisedb, if_exists="append", index=False)
            panwisedb.commit()
        
            if(os.path.exists(dir_path + "/temp/" + filename1)):
                os.remove(dir_path + "/temp/" + filename1)
            if(os.path.exists(dir_path + "/temp/" + filename1 + ".json")):
                os.remove(dir_path + "/temp/" + filename1 + ".json")

    mapping = "off"
    # filedata = BytesIO(content.encode())
    if upldtype == 'Sales Register - 59 Columns' or upldtype == "Sales Register SAP":
        # print('Sales Register being uploaded')
        try:
            try:
                query = "select DocKey from SalesRegisterDigi"
                doclist = pd.read_sql_query(query, panwisedb)
            except:
                doclist = pd.DataFrame(columns=['DocKey'])

            if upldtype == 'Sales Register - 59 Columns':
                if mapping == 'off' and filetype == 'csv':
                    df = pd.read_csv(filedata, dtype={'SourceIdentifier': 'str', 'SourceFileName': 'str', 'GLAccountCode': 'str', 'Division': 'str', 'SubDivision': 'str', 'ProfitCentre1': 'str', 'ProfitCentre2': 'str', 'PlantCode': 'str', 'ReturnPeriod': 'str', 'SupplierGSTIN': 'str', 'DocumentType': 'str', 'SupplyType': 'str', 'DocumentNumber': 'str', 'DocumentDate': 'str', 'OriginalDocumentNumber': 'str', 'OriginalDocumentDate': 'str', 'CRDRPreGST': 'str', 'LineNumber': 'str', 'CustomerGSTIN': 'str', 'UINorComposition': 'str', 'OriginalCustomerGSTIN': 'str', 'CustomerName': 'str', 'CustomerCode': 'str', 'BillToState': 'str', 'ShipToState': 'str', 'POS': 'str', 'PortCode': 'str', 'ShippingBillNumber': 'str', 'ShippingBillDate': 'str', 'FOB': 'float', 'ExportDuty': 'float', 'HSNorSAC': 'str', 'ProductCode': 'str', 'ProductDescription': 'str',
                                    'CategoryOfProduct': 'str', 'UnitOfMeasurement': 'str', 'Quantity': 'float', 'TaxableValue': 'float', 'IntegratedTaxRate': 'float', 'IntegratedTaxAmount': 'float', 'CentralTaxRate': 'float', 'CentralTaxAmount': 'float', 'StateUTTaxRate': 'float', 'StateUTTaxAmount': 'float', 'CessRateAdvalorem': 'float', 'CessAmountAdvalorem': 'float', 'CessRateSpecific': 'float', 'CessAmountSpecific': 'float', 'InvoiceValue': 'float', 'ReverseChargeFlag': 'str', 'TCSFlag': 'str', 'eComGSTIN': 'str', 'ITCFlag': 'str', 'ReasonForCreditDebitNote': 'str', 'AccountingVoucherNumber': 'str', 'AccountingVoucherDate': 'str', 'Userdefinedfield1': 'str', 'Userdefinedfield2': 'str', 'Userdefinedfield3': 'str'}, parse_dates=['DocumentDate', 'OriginalDocumentDate', 'ShippingBillDate', 'AccountingVoucherDate'], encoding='utf-8', thousands=',')

                elif mapping == 'off' and filetype == 'xlsx':
                    filedata = open(filedata)
                    df = pd.read_excel(BytesIO(b64decode(filedata.read())), engine='openpyxl', dtype={'SourceIdentifier': 'str', 'SourceFileName': 'str', 'GLAccountCode': 'str', 'Division': 'str', 'SubDivision': 'str', 'ProfitCentre1': 'str', 'ProfitCentre2': 'str', 'PlantCode': 'str', 'ReturnPeriod': 'str', 'SupplierGSTIN': 'str', 'DocumentType': 'str', 'SupplyType': 'str', 'DocumentNumber': 'str', 'DocumentDate': 'str', 'OriginalDocumentNumber': 'str', 'OriginalDocumentDate': 'str', 'CRDRPreGST': 'str', 'LineNumber': 'str', 'CustomerGSTIN': 'str', 'UINorComposition': 'str', 'OriginalCustomerGSTIN': 'str', 'CustomerName': 'str', 'CustomerCode': 'str', 'BillToState': 'str', 'ShipToState': 'str', 'POS': 'str', 'PortCode': 'str', 'ShippingBillNumber': 'str', 'ShippingBillDate': 'str', 'FOB': 'float', 'ExportDuty': 'float', 'HSNorSAC': 'str', 'ProductCode': 'str',
                                    'ProductDescription': 'str', 'CategoryOfProduct': 'str', 'UnitOfMeasurement': 'str', 'Quantity': 'float', 'TaxableValue': 'float', 'IntegratedTaxRate': 'float', 'IntegratedTaxAmount': 'float', 'CentralTaxRate': 'float', 'CentralTaxAmount': 'float', 'StateUTTaxRate': 'float', 'StateUTTaxAmount': 'float', 'CessRateAdvalorem': 'float', 'CessAmountAdvalorem': 'float', 'CessRateSpecific': 'float', 'CessAmountSpecific': 'float', 'InvoiceValue': 'float', 'ReverseChargeFlag': 'str', 'TCSFlag': 'str', 'eComGSTIN': 'str', 'ITCFlag': 'str', 'ReasonForCreditDebitNote': 'str', 'AccountingVoucherNumber': 'str', 'AccountingVoucherDate': 'str', 'Userdefinedfield1': 'str', 'Userdefinedfield2': 'str', 'Userdefinedfield3': 'str'}, parse_dates=['DocumentDate', 'OriginalDocumentDate', 'ShippingBillDate', 'AccountingVoucherDate'], thousands=',')
            elif upldtype == "Sales Register SAP":
                parse_dates_list=['Original_Invoice_Date_Reg','Shipping_Bill_Date_Reg','Posting_Date_Reg','Invoice_Date_Reg']
                dtype_dict = {
                    'UOM_Reg':'str',
                    'Accounting_Document_Number_Reg':'str',
                    'CGST_Rate_Reg':'float',
                    'SGST_Rate_Reg':'float',
                    'IGST_Rate_Reg':'float',
                    'Cess_Rate_Reg':'float',
                    'Tax_Rate_Reg':'float',
                    'CGST_Amount_Reg':'float',
                    'SGST_Amount_Reg':'float',
                    'IGST_Amount_Reg':'float',
                    'Cess_Amount_Reg':'float',
                    'Tax_Amount_Reg':'float',
                    'Taxable_Amount_Reg':'float',
                    'Supplier_GSTIN_Reg':'str',
                    'Customer_GSTIN_Reg':'str',
                    'Document_Type_Reg':'str',
                    'Supply_Type_Reg':'str',
                    'POS_Reg':'str',
                    'Qty_Reg':'float',
                    'Invoice_No_Reg':'str',
                    'HSN_Reg':'str',
                    'Customer_Name_Reg':'str',
                    'Description_Reg':'str'
                }
                if mapping == 'off' and filetype == 'csv':
                    df = pd.read_csv(filedata, dtype=dtype_dict, parse_dates=parse_dates_list, encoding='utf-8', thousands=',')
                elif mapping == 'off' and filetype == 'xlsx':
                    df = pd.read_excel(BytesIO(b64decode(filedata.read())), engine='openpyxl', dtype=dtype_dict, parse_dates=parse_dates_list, thousands=',')
            
            if set(columns) == set(srsap):
                df['SourceIdentifier'] = ''
                df['SourceFileName'] = ''
                df['GLAccountCode'] = ''
                df['Division'] = ''
                df['SubDivision'] = ''
                df['ProfitCentre1'] = ''
                df['ProfitCentre2'] = ''
                df['PlantCode'] = ''
                df['ReturnPeriod'] = ''
                df['SupplierGSTIN'] =df['Supplier_GSTIN_Reg']
                df['DocumentType'] =df['Document_Type_Reg']
                df['SupplyType'] =df['Supply_Type_Reg']
                df['DocumentNumber'] =df['Invoice_No_Reg']
                df['DocumentDate'] =df['Invoice_Date_Reg']
                df['OriginalDocumentNumber'] =df['Accounting_Document_Number_Reg']
                df['OriginalDocumentDate'] =df['Original_Invoice_Date_Reg']
                df['CRDRPreGST'] = ''
                df['LineNumber'] = ''
                df['CustomerGSTIN'] = df['Customer_GSTIN_Reg']
                df['UINorComposition'] = ''
                df['OriginalCustomerGSTIN'] = ''
                df['CustomerName'] = df['Customer_Name_Reg']
                df['CustomerCode'] = ''
                df['BillToState'] = ''
                df['ShipToState'] = ''
                df['POS'] = df['POS_Reg']
                df['PortCode'] = ''
                df['ShippingBillNumber'] = ''
                df['ShippingBillDate'] = df['Shipping_Bill_Date_Reg']
                df['FOB'] = ''
                df['ExportDuty'] = ''
                df['HSNorSAC'] = df['HSN_Reg']
                df['ProductCode'] = ''
                df['ProductDescription'] = df['Description_Reg']
                df['CategoryOfProduct'] = ''
                df['UnitOfMeasurement'] = df['UOM_Reg']
                df['Quantity'] = df['Qty_Reg']
                df['TaxableValue'] = df['Tax_Amount_Reg']
                df['IntegratedTaxRate'] = df['IGST_Rate_Reg']
                df['IntegratedTaxAmount'] = df['IGST_Amount_Reg']
                df['CentralTaxRate'] = df['CGST_Rate_Reg']
                df['CentralTaxAmount'] = df['CGST_Amount_Reg']
                df['StateUTTaxRate'] = df['SGST_Rate_Reg']
                df['StateUTTaxAmount'] = df['SGST_Amount_Reg']
                df['CessRateAdvalorem'] = df['Cess_Rate_Reg']
                df['CessAmountAdvalorem'] = df['Cess_Amount_Reg']
                df['CessRateSpecific'] = ''
                df['CessAmountSpecific'] = ''
                df['InvoiceValue'] = df['Taxable_Amount_Reg']
                df['ReverseChargeFlag'] = ''
                df['TCSFlag'] = ''
                df['eComGSTIN'] = ''
                df['ITCFlag'] = ''
                df['ReasonForCreditDebitNote'] = ''
                df['AccountingVoucherNumber'] = ''
                df['AccountingVoucherDate'] = ''
                df['Userdefinedfield1'] = ''
                df['Userdefinedfield2'] = ''
                df['Userdefinedfield3'] = ''
            
            df.drop(index=clientPAN_check(df, 'SupplierGSTIN'), inplace=True)
            if df.shape[0] != 0:
                df = df.reset_index(drop=True)

            pd.to_numeric(df['POS'])
            df['POS'] = df['POS'].map(str)
            df['POS'] = df["POS"].apply(lambda x: x.replace(
                r'.0', '') if isinstance(x, str) else x)
            
            df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
            
            df['GST'] = df['CentralTaxAmount'].round(2) + df['StateUTTaxAmount'].round(
                2) + df['IntegratedTaxAmount'].round(2) + df['CessAmountAdvalorem'].round(2)
            df['SourceFile'] = str(filename1)
            df['UploadTime'] = datetime.now()
            df['TotalTaxRate'] = df['CentralTaxRate'] + \
                df['StateUTTaxRate']+df['IntegratedTaxRate']
            df['DocumentDate'] = pd.to_datetime(
                df['DocumentDate'], dayfirst=True)
            df['FY'] = df['DocumentDate'].map(
                lambda x: x.year if x.month > 3 else x.year-1)
            df['DocKey'] = df['DocumentNumber'].map(str)+df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(
                str)+df['FY'].map(str)+df['DocumentType'].map(str) + df['TotalTaxRate'].map(str)
            totalcount = len(df)
            try:
                drop = df[df['DocKey'].isin(doclist['DocKey'])]
                drop.to_sql("SalesRegisterdropped", panwisedb,
                            if_exists="append", index=False)
                df.drop(df[df['DocKey'].isin(
                    doclist['DocKey'])].index, inplace=True)
            except:
                pass
            processedcount = len(df)
            droppedcount = totalcount - processedcount

            df[srcol59].to_sql("SalesRegisterDigi", panwisedb,
                    if_exists="append", index=False)
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype in ['Purchase Register - 65 Columns', 'Purchase Register - 28 Columns']:
        if upldtype == 'Purchase Register - 65 Columns':
            try:
                # print('Purchase Register being uploaded')
                try:
                    query = "select DocKey from PurchaseRegisterDigi"
                    doclist = pd.read_sql_query(query, panwisedb)
                except:
                    doclist = pd.DataFrame(columns=['DocKey'])
                # # print(doclist)
                if mapping == 'off' and filetype == 'csv':
                    totalcount = 0
                    processedcount = 0
                    for df in pd.read_csv(filedata, dtype={'SourceIdentifier': 'str', 'SourceFileName': 'str', 'GLAccountCode': 'str', 'Division': 'str', 'SubDivision': 'str', 'ProfitCentre1': 'str', 'ProfitCentre2': 'str', 'PlantCode': 'str', 'ReturnPeriod': 'str', 'RecipientGSTIN': 'str', 'DocumentType': 'str', 'SupplyType': 'str', 'DocumentNumber': 'str', 'DocumentDate': 'str', 'OriginalDocumentNumber': 'str', 'OriginalDocumentDate': 'str', 'CRDRPreGST': 'str', 'LineNumber': 'str', 'SupplierGSTIN': 'str', 'OriginalSupplierGSTIN': 'str', 'SupplierName': 'str', 'SupplierCode': 'str', 'POS': 'str', 'PortCode': 'str', 'BillOfEntry': 'str', 'BillOfEntryDate': 'str', 'CIFValue': 'float', 'CustomDuty': 'float', 'HSNorSAC': 'str', 'ItemCode': 'str', 'ItemDescription': 'str', 'CategoryOfItem': 'str', 'UnitOfMeasurement': 'str', 'Quantity': 'float', 'TaxableValue': 'float', 'IntegratedTaxRate': 'float', 'IntegratedTaxAmount': 'float', 'CentralTaxRate': 'float', 'CentralTaxAmount': 'float', 'StateUTTaxRate': 'float', 'StateUTTaxAmount': 'float', 'CessRateAdvalorem': 'float', 'CessAmountAdvalorem': 'float', 'CessRateSpecific': 'float', 'CessAmountSpecific': 'float', 'InvoiceValue': 'float', 'ReverseChargeFlag': 'str', 'EligibilityIndicator': 'str', 'CommonSupplyIndicator': 'str', 'AvailableIGST': 'float', 'AvailableCGST': 'float', 'AvailableSGST': 'float', 'AvailableCess': 'float', 'ITCReversalIdentifier': 'str', 'ReasonForCreditDebitNote': 'str', 'PurchaseVoucherNumber': 'str', 'PurchaseVoucherDate': 'str', 'PaymentVoucherNumber': 'str', 'PaymentDate': 'str', 'ContractNumber': 'str', 'ContractDate': 'str', 'ContractValue': 'float', 'Userdefinedfield1': 'str', 'Userdefinedfield2': 'str', 'Userdefinedfield3': 'str'}, parse_dates=['DocumentDate', 'OriginalDocumentDate', 'BillOfEntryDate', 'PurchaseVoucherDate', 'PaymentDate', 'ContractDate'],infer_datetime_format=True, chunksize=10000, encoding='utf-8', thousands=','):
                        df.drop(index=clientPAN_check(df, 'RecipientGSTIN'), inplace=True)
                        if df.shape[0] != 0:
                            df = df.reset_index(drop=True)

                        pd.to_numeric(df['POS'])
                        df['POS'] = df['POS'].map(str)
                        df['POS'] = df["POS"].apply(lambda x: x.replace(
                            r'.0', '') if isinstance(x, str) else x)
                        # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                        
                        df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
                        # print(df['POS'].unique())
                        
                        df['GST'] = df['CentralTaxAmount'].map(float).round(2) + df['StateUTTaxAmount'].map(float).round(
                            2) + df['IntegratedTaxAmount'].map(float).round(2) + df['CessAmountAdvalorem'].map(float).round(2)
                        #df['DocumentDate']= pd.to_datetime(df['DocumentDate'],dayfirst=True)
                        # # print(df['DocumentDate'])
                        # # print(df['FY'])
                        df['SourceFile'] = str(filename1)
                        df['UploadTime'] = datetime.now()
                        df['DocumentNumber'] = np.where(df['DocumentNumber'].map(
                            str).str[0:1] == "'", df['DocumentNumber'].map(str).str[1:], df['DocumentNumber'].map(str))

                        df['TotalTaxRate'] = df['CentralTaxRate'] + \
                            df['StateUTTaxRate']+df['IntegratedTaxRate']
                        df['DocumentDate'] = pd.to_datetime(
                            df['DocumentDate'], dayfirst=True)
                        df['FY'] = df['DocumentDate'].map(
                            lambda x: x.year if x.month > 3 else x.year-1)
                        df['DocKey'] = df['DocumentNumber'].map(str)+df['SupplierGSTIN'].map(str)+df['RecipientGSTIN'].map(
                            str)+df['FY'].map(str)+df['DocumentType'].map(str) + df['TotalTaxRate'].map(str)
                        totalcount += len(df)
                        try:
                            drop = df[df['DocKey'].isin(doclist['DocKey'])]
                            drop.to_sql("PurchaseRegisterdropped", panwisedb,
                                        if_exists="append", index=False)
                            df.drop(df[df['DocKey'].isin(
                                doclist['DocKey'])].index, inplace=True)
                        except:
                            pass
                        processedcount += len(df)

                        df.to_sql("PurchaseRegisterDigi", panwisedb,
                                if_exists="append", index=False)
                    droppedcount = totalcount - processedcount

                elif mapping == 'off' and filetype == 'xlsx':
                    filedata = open(filedata)
                    df = pd.read_excel(BytesIO(b64decode(filedata.read())), dtype={'SourceIdentifier': 'str', 'SourceFileName': 'str', 'GLAccountCode': 'str', 'Division': 'str', 'SubDivision': 'str', 'ProfitCentre1': 'str', 'ProfitCentre2': 'str', 'PlantCode': 'str', 'ReturnPeriod': 'str', 'RecipientGSTIN': 'str', 'DocumentType': 'str', 'SupplyType': 'str', 'DocumentNumber': 'str', 'DocumentDate': 'str', 'OriginalDocumentNumber': 'str', 'OriginalDocumentDate': 'str', 'CRDRPreGST': 'str', 'LineNumber': 'str', 'SupplierGSTIN': 'str', 'OriginalSupplierGSTIN': 'str', 'SupplierName': 'str', 'SupplierCode': 'str', 'POS': 'str', 'PortCode': 'str', 'BillOfEntry': 'str', 'BillOfEntryDate': 'str', 'CIFValue': 'float', 'CustomDuty': 'float', 'HSNorSAC': 'str', 'ItemCode': 'str', 'ItemDescription': 'str', 'CategoryOfItem': 'str', 'UnitOfMeasurement': 'str', 'Quantity': 'float', 'TaxableValue': 'float', 'IntegratedTaxRate': 'float', 'IntegratedTaxAmount': 'float',
                                    'CentralTaxRate': 'float', 'CentralTaxAmount': 'float', 'StateUTTaxRate': 'float', 'StateUTTaxAmount': 'float', 'CessRateAdvalorem': 'float', 'CessAmountAdvalorem': 'float', 'CessRateSpecific': 'float', 'CessAmountSpecific': 'float', 'InvoiceValue': 'float', 'ReverseChargeFlag': 'str', 'EligibilityIndicator': 'str', 'CommonSupplyIndicator': 'str', 'AvailableIGST': 'float', 'AvailableCGST': 'float', 'AvailableSGST': 'float', 'AvailableCess': 'float', 'ITCReversalIdentifier': 'str', 'ReasonForCreditDebitNote': 'str', 'PurchaseVoucherNumber': 'str', 'PurchaseVoucherDate': 'str', 'PaymentVoucherNumber': 'str', 'PaymentDate': 'str', 'ContractNumber': 'str', 'ContractDate': 'str', 'ContractValue': 'float', 'Userdefinedfield1': 'str', 'Userdefinedfield2': 'str', 'Userdefinedfield3': 'str'}, parse_dates=['DocumentDate', 'OriginalDocumentDate', 'BillOfEntryDate', 'PurchaseVoucherDate', 'PaymentDate', 'ContractDate'], thousands=',', engine='openpyxl')
                    
                    df.drop(index=clientPAN_check(df, 'RecipientGSTIN'), inplace=True)
                    if df.shape[0] != 0:
                        df = df.reset_index(drop=True)

                    pd.to_numeric(df['POS'])
                    df['POS'] = df['POS'].map(str)
                    df['POS'] = df["POS"].apply(lambda x: x.replace(
                        r'.0', '') if isinstance(x, str) else x)
                    # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                    
                    df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
                    # print(df['POS'].unique())
                    
                    df['GST'] = df['CentralTaxAmount'].round(2) + df['StateUTTaxAmount'].round(
                        2) + df['IntegratedTaxAmount'].round(2) + df['CessAmountAdvalorem'].round(2)
                    # # print(df['FY'])
                    df['SourceFile'] = str(filename1)
                    df['UploadTime'] = datetime.now()
                    df['DocumentNumber'] = np.where(df['DocumentNumber'].map(
                        str).str[0:1] == "'", df['DocumentNumber'].map(str).str[1:], df['DocumentNumber'].map(str))

                    df['TotalTaxRate'] = df['CentralTaxRate'] + \
                        df['StateUTTaxRate']+df['IntegratedTaxRate']
                    df['DocumentDate'] = pd.to_datetime(
                        df['DocumentDate'], dayfirst=True)
                    df['FY'] = df['DocumentDate'].map(
                        lambda x: x.year if x.month > 3 else x.year-1)
                    df['DocKey'] = df['DocumentNumber'].map(str)+df['SupplierGSTIN'].map(str)+df['RecipientGSTIN'].map(
                        str)+df['FY'].map(str)+df['DocumentType'].map(str) + df['TotalTaxRate'].map(str)
                    totalcount = len(df)
                    try:
                        drop = df[df['DocKey'].isin(doclist['DocKey'])]
                        drop.to_sql("PurchaseRegisterdropped", panwisedb,
                                    if_exists="append", index=False)
                        df.drop(df[df['DocKey'].isin(
                            doclist['DocKey'])].index, inplace=True)
                    except:
                        pass
                    processedcount = len(df)
                    droppedcount = totalcount - processedcount
                    df.to_sql("PurchaseRegisterDigi", panwisedb,
                            if_exists="append", index=False)
                    panwisedb.commit()
            except:
                logging.error(f"process_file : {upldtype}", exc_info=True)

        elif upldtype == 'Purchase Register - 28 Columns':
            try:
                # print('Purchase Register being uploaded')
                databasecols = ['SupplierName', 'SupplierGSTIN', 'HSNorSAC', 'ReturnPeriod', 'PurchaseVoucherDate', 'DocumentNumber', 'DocumentDate', 'InvoiceValue', 'TaxableValue', 'CentralTaxAmount', 'StateUTTaxAmount', 'IntegratedTaxAmount', 'CessAmountAdvalorem',
                                'POS', 'RecipientGSTIN', 'DocumentType', 'EligibilityIndicator', 'OriginalDocumentNumber', 'OriginalDocumentDate', 'ReverseChargeFlag', 'Userdefinedfield1', 'Userdefinedfield2', 'Userdefinedfield3', 'Userdefinedfield4', 'Userdefinedfield5']
                try:
                    query = "select DocKey from PurchaseRegisterDigi"
                    doclist = pd.read_sql_query(query, panwisedb)
                except:
                    doclist = pd.DataFrame(columns=['DocKey'])
                if mapping == 'off' and filetype == 'csv':

                    for df in pd.read_csv(filedata, infer_datetime_format=True, dayfirst=True, dtype={'Sl.No': 'str', ' Name of Vendor ': 'str', 'Supplier GSTIN': 'str', 'HSN ': 'str', 'Return Period': 'str', 'Availment Month': 'str', 'Posting Date ': 'str', 'DocumentNumber': 'str', 'Document Date ': 'str', 'Document amount  ': 'float', 'Taxable Value ': 'float', ' CGST ': 'float', ' SGST ': 'float', ' IGST ': 'float', 'Cess': 'float', ' Total GST ': 'float', 'POS': 'str', 'Recepient GSTIN': 'str', 'Document Type': 'str', 'Eligibility Indicator': 'str', 'Original Document Number': 'str', 'Original Document Date': 'str', 'RCM': 'str', 'EY1': 'str', 'EY2': 'str', 'EY3': 'str', 'EY4': 'str', 'EY5': 'str'}, parse_dates=['Posting Date ', 'Document Date ', 'Original Document Date'], chunksize=10000, encoding='utf-8', thousands=','):
                        df.drop(index=clientPAN_check(df, 'Recepient GSTIN'), inplace=True)
                        if df.shape[0] != 0:
                            df = df.reset_index(drop=True)

                        neworder = [' Name of Vendor ', 'Supplier GSTIN', 'HSN ', 'Return Period', 'Posting Date ', 'DocumentNumber', 'Document Date ', 'Document amount  ', 'Taxable Value ', ' CGST ', ' SGST ',
                                    ' IGST ', 'Cess', 'POS', 'Recepient GSTIN', 'Document Type', 'Eligibility Indicator', 'Original Document Number', 'Original Document Date', 'RCM', 'EY1', 'EY2', 'EY3', 'EY4', 'EY5']
                        df = df.reindex(columns=neworder)
                        df.columns = databasecols

                        pd.to_numeric(df['POS'])
                        df['POS'] = df['POS'].map(str)
                        df['POS'] = df["POS"].apply(lambda x: x.replace(
                            r'.0', '') if isinstance(x, str) else x)
                        # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                        
                        df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
                        # print(df['POS'].unique())

                        df['GST'] = df['CentralTaxAmount'].round(2) + df['StateUTTaxAmount'].round(2) + df['IntegratedTaxAmount'].round(2) + df['CessAmountAdvalorem'].round(2)
                        # # print(df['DocumentDate'])
                        # # print(df['FY'])
                        df['DocumentNumber'] = np.where(df['DocumentNumber'].map(str).str[0:1] == "'", df['DocumentNumber'].map(str).str[1:], df['DocumentNumber'].map(str))

                        df['SourceFile'] = str(filename1)
                        df['UploadTime'] = datetime.now()
                        #df['DocumentDate']= pd.to_datetime(df['DocumentDate'],dayfirst=True, format='%d/%m/%Y', errors='ignore')
                        df['FY'] = df['DocumentDate'].map(
                            lambda x: x.year if x.month > 3 else x.year-1)
                        df['DocKey'] = df['DocumentNumber'].map(str)+df['SupplierGSTIN'].map(str)+df['RecipientGSTIN'].map(
                            str)+df['FY'].map(str)+df['DocumentType'].map(str) + df['GST'].map(str)
                        totalcount += len(df)
                        try:
                            drop = df[df['DocKey'].isin(doclist['DocKey'])]
                            drop.to_sql("PurchaseRegisterdropped", panwisedb,
                                        if_exists="append", index=False)
                            df.drop(df[df['DocKey'].isin(
                                doclist['DocKey'])].index, inplace=True)
                        except:
                            pass
                        processedcount += len(df)

                        df.to_sql("PurchaseRegisterDigi", panwisedb,
                                if_exists="append", index=False)
                    droppedcount = totalcount - processedcount
                    df.to_csv("GAPS\PR.csv", index=False)
                elif mapping == 'off' and filetype == 'xlsx':
                    filedata = open(filedata)
                    # for df in pd.read_excel(filedata, infer_datetime_format=True, dayfirst=True, dtype={'Sl.No': 'str', ' Name of Vendor ': 'str', 'Supplier GSTIN': 'str', 'HSN ': 'str', 'Return Period': 'str', 'Availment Month': 'str', 'Posting Date ': 'str', 'DocumentNumber': 'str', 'Document Date ': 'str', 'Document amount  ': 'float', 'Taxable Value ': 'float', ' CGST ': 'float', ' SGST ': 'float', ' IGST ': 'float', 'Cess': 'float', ' Total GST ': 'float', 'POS': 'str', 'Recepient GSTIN': 'str', 'Document Type': 'str', 'Eligibility Indicator': 'str', 'Original Document Number': 'str', 'Original Document Date': 'str', 'RCM': 'str', 'EY1': 'str', 'EY2': 'str', 'EY3': 'str', 'EY4': 'str', 'EY5': 'str'}, parse_dates=['Posting Date ', 'Document Date ', 'Original Document Date'], chunksize=10000, encoding='utf-8', thousands=','):
                    df = pd.read_excel(BytesIO(b64decode(filedata.read())), dtype={'Sl.No': 'str', ' Name of Vendor ': 'str', 'Supplier GSTIN': 'str', 'HSN ': 'str', 'Return Period': 'str', 'Availment Month': 'str', 'Posting Date ': 'str', 'DocumentNumber': 'str', 'Document Date ': 'str', 'Document amount  ': 'float', 'Taxable Value ': 'float', ' CGST ': 'float', ' SGST ': 'float', ' IGST ': 'float', 'Cess': 'float', ' Total GST ': 'float', 'POS': 'str', 'Recepient GSTIN': 'str', 'Document Type': 'str', 'Eligibility Indicator': 'str', 'Original Document Number': 'str', 'Original Document Date': 'str', 'RCM': 'str', 'EY1': 'str', 'EY2': 'str', 'EY3': 'str', 'EY4': 'str', 'EY5': 'str'}, parse_dates=['Posting Date ', 'Document Date ', 'Original Document Date'], thousands=',', engine='openpyxl')

                    df.drop(index=clientPAN_check(df, 'Recepient GSTIN'), inplace=True)
                    if df.shape[0] != 0:
                        df = df.reset_index(drop=True)
                    
                    neworder = [' Name of Vendor ', 'Supplier GSTIN', 'HSN ', 'Return Period', 'Posting Date ', 'DocumentNumber', 'Document Date ', 'Document amount  ', 'Taxable Value ', ' CGST ',
                                ' SGST ', ' IGST ', 'Cess', 'POS', 'Recepient GSTIN', 'Document Type', 'Eligibility Indicator', 'Original Document Number', 'Original Document Date', 'RCM', 'EY1', 'EY2', 'EY3']
                    df = df.reindex(columns=neworder)
                    df.columns = databasecols

                    pd.to_numeric(df['POS'])
                    df['POS'] = df['POS'].map(str)
                    df['POS'] = df["POS"].apply(lambda x: x.replace(
                        r'.0', '') if isinstance(x, str) else x)
                    # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                    df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
                    # print(df['POS'].unique())

                    df['GST'] = df['CentralTaxAmount'].round(2) + df['StateUTTaxAmount'].round(
                        2) + df['IntegratedTaxAmount'].round(2) + df['CessAmountAdvalorem'].round(2)

                    #df['DocumentDate']= pd.to_datetime(df['DocumentDate'],dayfirst=True)
                    # # print(df['DocumentDate'])
                    # # print(df['FY'])
                    df['SourceFile'] = str(filename1)
                    df['UploadTime'] = datetime.now()
                    df['DocumentNumber'] = np.where(df['DocumentNumber'].map(
                        str).str[0:1] == "'", df['DocumentNumber'].map(str).str[1:], df['DocumentNumber'].map(str))

                    df['TotalTaxRate'] = df['CentralTaxRate'] + \
                        df['StateUTTaxRate']+df['IntegratedTaxRate']
                    #df['DocumentDate']= pd.to_datetime(df['DocumentDate'],dayfirst=True, format='%d/%m/%Y', errors='ignore')
                    df['FY'] = df['DocumentDate'].map(
                        lambda x: x.year if x.month > 3 else x.year-1)
                    df['DocKey'] = df['DocumentNumber'].map(str)+df['SupplierGSTIN'].map(str)+df['RecipientGSTIN'].map(
                        str)+df['FY'].map(str)+df['DocumentType'].map(str) + df['GST'].map(str)
                    totalcount += len(df)
                    try:
                        drop = df[df['DocKey'].isin(doclist['DocKey'])]
                        drop.to_sql("PurchaseRegisterdropped", panwisedb,
                                    if_exists="append", index=False)
                        df.drop(df[df['DocKey'].isin(
                            doclist['DocKey'])].index, inplace=True)
                    except:
                        pass
                    processedcount += len(df)

                    df.to_sql("PurchaseRegisterDigi", panwisedb,
                                if_exists="append", index=False)
                    droppedcount = totalcount - processedcount
            except:
                logging.error(f"process_file : {upldtype}", exc_info=True)

        # seqRun1(lst, clientPAN, current_user) # inward flash/pr flash
        seqRun1(True, [clientPAN, current_user]) # inward flash/pr flash
        valpr(clientPAN, current_user) # inward validation/validate purchase register

    if upldtype == 'GSTR 2A':
        try:
            ## print('GSTR-2A being uploaded')
            try:
                query = "select DocKey from GSTR_2A"
                doclist = pd.read_sql_query(query, panwisedb)
            except:
                doclist = pd.DataFrame(columns=['DocKey'])
            if filetype == 'zip':

                totalcount, processedcount, droppedcount = gstr2a(content, clientPAN, filename1, current_user)
                df1 = pd.DataFrame(columns=['File Name', 'Type', 'Extention', 'Time Stamp', 'Count', 'Uploaded', 'Dropped'], data=[
                                [filename1, upldtype, filetype, datetime.now(), totalcount, processedcount, droppedcount]])
                df1.to_sql("filelist", panwisedb, if_exists="append", index=False)

            if filetype == 'json':

                totalcount, processedcount, droppedcount = gstr2ajson(content, clientPAN, filename1, current_user)
                df1 = pd.DataFrame(columns=['File Name', 'Type', 'Extention', 'Time Stamp', 'Count', 'Uploaded', 'Dropped'], data=[
                                [filename1, upldtype, filetype, datetime.now(), totalcount, processedcount, droppedcount]])
                df1.to_sql("filelist", panwisedb, if_exists="append", index=False)

            if filetype == 'csv' or filetype == 'xlsx':
                if filetype == 'csv':
                    df = pd.read_csv(filedata, dtype={'Counter Party Return Status': 'str', 'Return Period': 'str', 'Recipent GSTIN': 'str', 'Document Type': 'str', 'Document Number': 'str', 'Document Date': 'str', 'Original Document Number': 'str', 'Original Document Date': 'str', 'Supplier GSTIN': 'str', 'Supplier Name' : 'str', 'POS': 'str', 'Taxable Value ': 'float', 'Tax Rate': 'float', 'Integrated Tax Amount': 'float', 'Central Tax Amount': 'float', 'StateUT TaxAmount': 'float',
                                    'Cess Amount': 'float', 'Invoice Value': 'float', 'Reverse Charge Flag': 'str', 'Differential Percentage': 'str', 'DeLinking Flag': 'str', 'CFS_GSTR3B': 'str', 'CancellationDt': 'str', 'GSTR1_FilingDt': 'str', 'GSTR1_FilingPeriod': 'str', 'Generation Date': 'str', 'IRN': 'str'}, parse_dates=['Document Date', 'Original Document Date', 'CancellationDt', 'GSTR1_FilingDt', 'Generation Date'], encoding='utf-8', thousands=',')
                    #df.drop(columns=['CR/DR Pre-GST','Line Number','Supplier Name','State Name','ITC Eligible','Record Type','OrgInvAmendmentPeriod','OrgInvAmendmentType','Source Type'],inplace=True)

                if filetype == 'xlsx':
                    filedata = open(filedata)
                    df = pd.read_excel(BytesIO(b64decode(filedata.read())), engine='openpyxl', dtype={'Counter Party Return Status': 'str', 'Return Period': 'str', 'Recipent GSTIN': 'str', 'Document Type': 'str', 'Document Number': 'str', 'Document Date': 'str', 'Original Document Number': 'str', 'Original Document Date': 'str', 'Supplier GSTIN': 'str', 'Supplier Name' : 'str', 'POS': 'str', 'Taxable Value ': 'float', 'Tax Rate': 'float', 'Integrated Tax Amount': 'float', 'Central Tax Amount': 'float', 'StateUT TaxAmount': 'float',
                                    'Cess Amount': 'float', 'Invoice Value': 'float', 'Reverse Charge Flag': 'str', 'Differential Percentage': 'str', 'DeLinking Flag': 'str', 'CFS_GSTR3B': 'str', 'CancellationDt': 'str', 'GSTR1_FilingDt': 'str', 'GSTR1_FilingPeriod': 'str', 'Generation Date': 'str', 'IRN': 'str'}, parse_dates=['Document Date', 'Original Document Date', 'CancellationDt', 'GSTR1_FilingDt', 'Generation Date'], thousands=',')

                df.drop(index=clientPAN_check(df, 'Recipent GSTIN'), inplace=True)
                if df.shape[0] != 0:
                    df = df.reset_index(drop=True)

                df['Document Date'] = pd.to_datetime(
                    df['Document Date'], dayfirst=True)
                df['Invoice Date'] = pd.to_datetime(
                    df['Invoice Date'], dayfirst=True)
                df['Document Number'] = np.where(df['Document Number'].map(
                    str).str[0:1] == "'", df['Document Number'].map(str).str[1:], df['Document Number'].map(str))
                df['Original Document Number'] = np.where(df['Original Document Number'].map(
                    str).str[0:1] == "'", df['Original Document Number'].map(str).str[1:], df['Original Document Number'].map(str))

                df['Original Document Number'] = np.where(df['Original Document Number'].map(
                    str) == "nan", "", df['Original Document Number'])

                df['Invoice Number'] = np.where(df['Invoice Number'].map(
                    str).str[0:1] == "'", df['Invoice Number'].map(str).str[1:], df['Invoice Number'].map(str))

                df['Invoice Number'] = np.where(df['Invoice Number'].map(
                    str) == "nan", "", df['Invoice Number'])

                df['GSTR1_FilingDt'] = pd.to_datetime(
                    df['GSTR1_FilingDt'], dayfirst=True)
                df['GSTR-2A category'] = ""
                df['GSTR-2A category'] = np.where(df['Document Type']
                                                == "C", "cdn", df['GSTR-2A category'])
                df['GSTR-2A category'] = np.where(df['Document Type']
                                                == "D", "dbn", df['GSTR-2A category'])
                df['GSTR-2A category'] = np.where(df['Document Type']
                                                == "R", "b2b", df['GSTR-2A category'])
                df['GSTR-2A category'] = np.where(np.logical_and(df['Document Type'] == "R",
                                                df['Original Document Number'].str.len() > 0), "b2ba", df['GSTR-2A category'])
                df['GSTR-2A category'] = np.where(np.logical_and(df['Document Type'] == "C", np.logical_and(
                    df['Original Document Number'].str.len() > 0, df['Invoice Number'].str.len() > 0)), "cdna", df['GSTR-2A category'])
                df['GSTR-2A category'] = np.where(np.logical_and(df['Document Type'] == "D", np.logical_and(
                    df['Original Document Number'].str.len() > 0, df['Invoice Number'].str.len() > 0)), "dbna", df['GSTR-2A category'])

                df['GSTR-2A category'] = np.where(df['Document Type']
                                                == "IMPG", "impg", df['GSTR-2A category'])
                df['Document Number'] = np.where(
                    df['Document Type'] == "IMPG", df['Invoice Number'], df['Document Number'])
                df['Document Date'] = np.where(
                    df['Document Type'] == "IMPG", df['Invoice Date'], df['Document Date'])
                df['GSTR1_FilingDt'] = np.where(
                    df['Document Type'] == "IMPG", df['ReferenceDt'], df['GSTR1_FilingDt'].map(str))
                df['Supplier GSTIN'] = np.where(
                    df['Document Type'] == "IMPG", "Dummy", df['Supplier GSTIN'])

                df['Document Type'] = np.where(
                    df['Document Type'] == "IMPG", df['PortCode'], df['Document Type'])
                df.drop(columns=['CR/DR Pre-GST', 'Line Number', 'State Name', 'ITC Eligible', 'Invoice Date', 'Invoice Number',
                        'ReferenceDt', 'Record Type', 'PortCode', 'OrgInvAmendmentPeriod', 'OrgInvAmendmentType', 'Source Type'], inplace=True)
                # # print(df.columns)
                df['Document Date'] = pd.to_datetime(
                    df['Document Date'], dayfirst=True)
                df['FY'] = df['Document Date'].map(
                    lambda x: x.year if x.month > 3 else x.year-1)
                duplicate = df[df.duplicated()]
                duplicate['Reason'] = "Duplicate"
                df.drop_duplicates(keep='first', inplace=True)
                duplicate.to_csv("GAPS\GSTR2A - Duplicate Drops.csv", index=False)
                df = df.groupby(['Counter Party Return Status', 'Supplier GSTIN', 'Document Number', 'FY', 'GSTR-2A category', 'Document Type'], as_index=False).agg({'Counter Party Return Status': 'first', 'Return Period': 'first', 'Recipent GSTIN': 'first', 'Document Type': 'first', 'Document Number': 'first', 'Document Date': 'first', 'Original Document Number': 'first', 'Original Document Date': 'first', 'Supplier GSTIN': 'first', 'Supplier Name': 'first', 'POS': 'first',
                                                                                                                                                                    'Taxable Value ': 'sum', 'Tax Rate': 'sum', 'Integrated Tax Amount': 'sum', 'Central Tax Amount': 'sum', 'StateUT TaxAmount': 'sum', 'Cess Amount': 'sum', 'Invoice Value': 'sum', 'Reverse Charge Flag': 'first', 'Differential Percentage': 'first', 'DeLinking Flag': 'first', 'CFS_GSTR3B': 'first', 'CancellationDt': 'first', 'GSTR1_FilingDt': 'first', 'GSTR1_FilingPeriod': 'first', 'Generation Date': 'first', 'IRN': 'first', 'GSTR-2A category': 'first'})
                df.drop(columns=['FY'], inplace=True)
                df = df.reindex()
                # # print(df)
                neworder = ['CFS', 'TaxPeriod', 'CustomerGSTIN', 'Supply Type', 'DocumentNo', 'DocumentDate', 'Original Invoice Number', 'Original Invoice Date', 'SupplierGSTIN', 'Supplier Name', 'POS', 'Taxable Value', 'Tax Rate', 'IGST Amount', 'Central Tax Amount',
                            'State/UT Tax Amount', 'CessAmount', 'InvoiceValue', 'ReverseCharge', 'Differential %', 'Delinking Flag', 'GSTR 3B Filed', 'Cancelation Date', 'Filing Date', 'Filing Period', 'IRN Gen-Date', 'IRN', 'GSTR-2A category']
                df.columns = neworder
                df['Taxable Value'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['Taxable Value'] > 0), df['Taxable Value']*-1, df['Taxable Value'])
                df['Central Tax Amount'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['Central Tax Amount'] > 0), df['Central Tax Amount']*-1, df['Central Tax Amount'])
                df['State/UT Tax Amount'] = np.where(np.logical_and(df['Supply Type'] == "cdn",
                                                    df['State/UT Tax Amount'] > 0), df['State/UT Tax Amount']*-1, df['State/UT Tax Amount'])
                df['IGST Amount'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['IGST Amount'] > 0), df['IGST Amount']*-1, df['IGST Amount'])
                df['CessAmount'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['CessAmount'] > 0), df['CessAmount']*-1, df['CessAmount'])

                #df['Original Invoice Date']= pd.to_datetime(df['Original Invoice Date'],dayfirst=True)
                df['DocumentNo'] = df['DocumentNo'].map(str).str.upper()
                df['Original Invoice Number'] = df['Original Invoice Number'].str.upper()

                try:
                    df['DocumentNo'] = df['DocumentNo'].str.lstrip('0')
                except:
                    pass

                try:
                    df['Original Invoice Number'] = df['Original Invoice Number'].str.lstrip(
                        '0')
                except:
                    pass
                df['FY'] = df['DocumentDate'].map(
                    lambda x: x.year if x.month > 3 else x.year-1)

                df['Tax'] = df['State/UT Tax Amount'] + \
                    df['Central Tax Amount']+df['IGST Amount'] + df['CessAmount']
                df['SourceFile'] = str(filename1)
                df['UploadTime'] = datetime.now()
                df['DocKey'] = df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
                    str)+df['FY'].map(str)+df['Supply Type'].map(str)+df['GSTR-2A category'].map(str)
                totalcount = len(df)
                try:
                    drop = df[df['DocKey'].isin(doclist['DocKey'])]
                    drop.to_sql("GSTR_2Adropped", panwisedb,
                                if_exists="append", index=False)
                    df.drop(df[df['DocKey'].isin(
                        doclist['DocKey'])].index, inplace=True)
                except:
                    pass
                processedcount = len(df)
                droppedcount = totalcount - processedcount

                duplicate = df[df.duplicated()]
                duplicate['Reason'] = "Duplicate"
                df.drop_duplicates(keep='first', inplace=True)
                #cdf1['FY'] = pd.to_datetime(cdf1['DocumentDate'],dayfirst=True,errors='coerce')
                #cdf1['FY'] = np.where(cdf1['FY'] != "NaT", cdf1['FY'].map(str).str[0:4],"")

                duplicate1 = df[df.duplicated(subset=[
                                            'CFS', 'SupplierGSTIN', 'DocumentNo', 'FY', 'GSTR-2A category', 'Supply Type'])]
                duplicate1['Reason'] = "Duplicate"
                df.drop_duplicates(subset=['CFS', 'SupplierGSTIN', 'DocumentNo', 'FY',
                                'GSTR-2A category', 'Supply Type'], keep='first', inplace=True)

                df.drop(df[np.logical_or(df['GSTR-2A category'] == "impg",
                        df['GSTR-2A category'] == "tds")].index, inplace=True)
                df.drop(df[df['GSTR-2A category'] == "tdsa"].index, inplace=True)
                df.reset_index(inplace=True)

                df['cdnpos'] = df['DocumentNo'].map(
                    str) + df['SupplierGSTIN'].map(str) + df['DocumentDate'].map(str)
                cdnpos = df.groupby('cdnpos', as_index=False).agg(
                    {'POS': 'first', 'GSTR-2A category': 'first'})
                cdnpos = cdnpos[cdnpos['GSTR-2A category'] == "b2b"]
                df['cdnpos'] = df['Original Invoice Number'].map(
                    str) + df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)

                df = pd.merge(df, cdnpos, on='cdnpos', how='left')
                df['POS_x'] = np.where(df['POS_x'].map(
                    str) == "nan", df['POS_y'], df['POS_x'])
                df = df.rename(
                    columns={'POS_x': 'POS', 'GSTR-2A category_x': 'GSTR-2A category'})
                df.drop(columns=['POS_y', 'GSTR-2A category_y'], inplace=True)
                #cdf1.drop(cdf1[~cdf1['CustomerGSTIN'].isin(gstins)].index, inplace=True)
                df['Dropcfs'] = ""
                df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2A category'] == "b2ba", df['CFS'] == "Y"), df['Original Invoice Number'].map(
                    str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2A category'].str[0:3]+df['Supply Type'].map(str), df['Dropcfs'])
                df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2A category'] == "cdna", df['CFS'] == "Y"), df['Original Invoice Number'].map(
                    str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2A category'].str[0:3]+df['Supply Type'].map(str), df['Dropcfs'])
                df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2A category'] == "cdna", df['CFS'] == "Y"), np.where(df['Dropcfs'] == "", df['DocumentNo'].map(
                    str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2A category']+df['Supply Type'].map(str), df['Dropcfs']), df['Dropcfs'])

                # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
                #cdf1['Dropcfs'] = np.where(np.logical_and(cdf1['GSTR-2A category']=="cdna",cdf1['Dropcfs'] ==""),np.where(cdf1['CFS']=="Y",cdf1['DocumentNo'].map(str)+cdf1['SupplierGSTIN'].map(str)+cdf1['DocumentDate'].map(str)+cdf1['GSTR-2A category'].str[0:2],cdf1['Dropcfs'] ),cdf1['Dropcfs'] )

                cfsdrop = df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2A category'] +
                                            df['Supply Type'].map(str)).isin(df['Dropcfs']), np.logical_and(df['Dropcfs'] == "", df['Supply Type'].map(str) == "R"))]
                cfsdrop['Reason'] = "Amended"

                df.drop(df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2A category'] +
                        df['Supply Type'].map(str)).isin(df['Dropcfs']), np.logical_and(df['Dropcfs'] == "", df['Supply Type'].map(str) == "R"))].index, inplace=True)

                #dropcfsn = df[df['CFS']=="N"]
                #df.drop(df[df['CFS']=="N"].index, inplace=True)
                #dropcfsn['Reason']="Not Filed"

                # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)

                duplicate1 = duplicate.append(cfsdrop)
                # duplicate2=duplicate1.append(dropcfsn)
                # duplicate1.to_csv("PR2A\Dropped 2A.csv", index=False)
                df['DocumentNo'] = df['DocumentNo'].astype(
                    str).str.replace('\.0', '', regex=True)
                # cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
                #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)

                df.reset_index(inplace=True)
                df['first1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x).start(
                ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
                df['last1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x[::-1]).start(
                ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
                #df['last1']= len(df['DocumentNo']) - df['last1']
                # # print(df['first1'])
                df['first1'] = df['first1'].fillna(
                    df['DocumentNo'].str.len()).astype(int)
                df['last1'] = df['last1'].fillna(
                    df['DocumentNo'].str.len()).astype(int)
                df['DocNoBfrSpl'] = [DocumentNo[:first1]
                                    for DocumentNo, first1 in zip(df.DocumentNo, df.first1)]
                df['DocNoAftrSpl'] = [DocumentNo[-last1:]
                                    for DocumentNo, last1 in zip(df.DocumentNo, df.last1)]
                df['DocNoWOSplChar'] = df['DocumentNo'].str.replace(
                    '[^a-zA-Z0-9]', '', regex=True)
                df['DocNoNumeric'] = df['DocumentNo'].str.replace(
                    r"[a-zA-Z]", '').replace('\W', '').replace('_', '').replace('/', '').replace("\\", '')
                # df['DocNoBfrSpl'] =
                df['DocNo'] = df['DocumentNo']
                # df['DocNoAftrSpl'] =
                df['SuppGSTIN'] = df['SupplierGSTIN']
                df['SuppPAN'] = df['SupplierGSTIN'].str[2:12].map(str)
                df['ResGSTIN'] = df['CustomerGSTIN']
                df['ResPAN'] = df['CustomerGSTIN'].str[2:12].map(str)
                df['DocDate'] = df['DocumentDate'].apply(
                    lambda x: x.strftime('%d%m%Y') if x else "")
                df['MMYYYY'] = df['DocumentDate'].apply(
                    lambda x: x.strftime('%m%Y') if x else "")
                #df['FY'] = pd.to_datetime(df['DocumentDate'],dayfirst=True).apply(lambda x: x.strftime('%Y') if x else "")
                df['TaxVal'] = df['Taxable Value'].round(2)
                df['InvVal'] = df['InvoiceValue'].round(2)
                df['CGST'] = df['Central Tax Amount'].round(2)
                df['SGST'] = df['State/UT Tax Amount'].round(2)
                df['IGST'] = df['IGST Amount'].round(2)
                df['GST'] = df['Central Tax Amount'].round(2) + df['State/UT Tax Amount'].round(
                    2) + df['IGST Amount'].round(2) + df['CessAmount'].round(2)

                pd.to_numeric(df['POS'])
                df['POS'] = df['POS'].map(str)
                df['POS'] = df["POS"].apply(lambda x: x.replace(
                    r'.0', '') if isinstance(x, str) else x)
                # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                
                df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)

                df['RCM'] = df['ReverseCharge']
                df['DocType'] = np.where(df['Supply Type'] == "C", "cdn", np.where(
                    df['Supply Type'] == "D", "dbn", "b2b"))
                df['NA'] = ""

                df['Key1'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
                    str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key2'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
                    str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key3'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key4'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key5'] = df['SupplierGSTIN'].map(
                    str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key6'] = df['SupplierGSTIN'].map(
                    str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key7'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(str)+df['Central Tax Amount'].round(
                    2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key8'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
                    str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
                df['Key9'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
                    str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
                df['Key10'] = df['FY'].map(str) + df['SuppPAN'].map(str) + df['DocNoWOSplChar'].map(
                    str)+df['ResPAN'].map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)

                df['Key11'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key12'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key13'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
                    str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key14'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
                    str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key15'] = df['SupplierGSTIN'].map(
                    str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key16'] = df['SupplierGSTIN'].map(
                    str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key19'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(
                    str)+df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)
                df['Key20'] = df['FY'].map(str) + df['SuppPAN'].map(
                    str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocType'].map(str)

                df['Key17'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(
                    str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key18'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(
                    str)+df['DocDate'].map(str)+df['DocType'].map(str)
                try:
                    query = "select * FROM matchconfig"
                    config = pd.read_sql_query(query, panwisedb)
                    rowdata = dataframe_to_rows(config, index=False, header=False)
                    rowdata1 = dataframe_to_rows(config, index=False, header=False)
                    rowdata2 = dataframe_to_rows(config, index=False, header=False)
                    matchcon = 1
                except:
                    matchcon = 0
                    config = []
                l = len(config)
                n = 11

                try:
                    for row1 in rowdata1:

                        key = "Key" + str(n)
                        df[key] = df[row1[1]].map(str) + df[row1[2]].map(str) + df[row1[3]].map(str) + df[row1[4]].map(
                            str) + df[row1[5]].map(str) + df[row1[6]].map(str) + df[row1[7]].map(str) + df[row1[8]].map(str)
                        n += 1
                except:
                    pass
                df.to_sql("GSTR_2A", panwisedb, if_exists="append", index=False)
                # tcdf1=tcdf.copy()
                #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)

                df.to_csv(r"GAPS\GSTR2a.csv", index=False)

            try:
                query = "select Key1 from GSTR_2A"
                doclist2a = pd.read_sql_query(query, panwisedb)
            except:
                doclist2a = pd.DataFrame(columns=['NoKey'])
            if "NoKey" in doclist2a.columns:
                router.message = "Sanitizing GSTR 2A"
                gstr2asanitize()
            elif any(doclist2a['Key1'].isna()) or any(doclist2a['Key1'] == ""):
                router.message = "Sanitizing GSTR 2A"
                gstr2asanitize()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'GSTR 1':
        try:
            # print('GSTR1')
            # print(filetype)
            # print('GSTR 1 being uploaded')
            if filetype == 'zip':
                print('GSTR1 zip json')
                totalcount, processedcount, droppedcount = gstr1(
                    filedata, clientPAN, filename1, current_user)
            elif filetype == 'json':
                totalcount, processedcount, droppedcount = gstr1(
                    filedata, clientPAN, filename1, current_user)
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'GSTR 2B':
        try:
            # print('GSTR-2B being uploaded')

            if filetype == 'zip':
                totalcount, processedcount, droppedcount = gstr2b(
                    filedata, clientPAN, filename1, current_user)
            if filetype == 'json':
                totalcount, processedcount, droppedcount = gstr2bjson(content, clientPAN, filename1, current_user)
            if filetype == 'csv' or filetype == 'xlsx':
                # print('Gstr2b csv identified')
                filedata = open(filedata)
                if filetype == 'xlsx':
                    # print('GSTR 2B - 39 Columns XLSX')
                    df = pd.read_excel(BytesIO(b64decode(filedata.read())), engine='openpyxl', dtype={
                        'Return Period':'str',
                        'Recipient GSTIN':'str',
                        'Supplier GSTIN':'str',
                        'Supplier Name':'str',
                        'Document Type':'str',
                        'Supply Type':'str',
                        'Document Number':'str',
                        'Document Date':'str',
                        'Taxable Value ':'float',
                        'Tax Rate':'float',
                        'IGST Amount':'float',
                        'CGST Amount':'float',
                        'SGST Amount':'float',
                        'CESS Amount':'float',
                        'Invoice Value':'float',
                        'POS':'str',
                        'State Name':'str',
                        'Line Number':'str',
                        'BOE-ReferenceDate(ICEGATE)':'str',
                        'BOE-Received Date(GSTN)':'str',
                        'PortCode':'str',
                        'Bill Of Entry Number':'str',
                        'Bill Of Entry Date':'str',
                        'BOE-Amendment Flag':'str',
                        'Original Document Number':'str',
                        'Original Document Date':'str',
                        'Original Document Type':'str',
                        'Original Invoice Number':'str',
                        'Original Invoice Date':'str',
                        '2B-GenerationDate':'str',
                        'GSTR-1/5/6 Filing Period':'str',
                        'GSTR-1/5/6 Filing Date':'str',
                        'Differential Percentage':'str',
                        'Reverse Charge Flag':'str',
                        'ITC Availability':'str',
                        'Reason for ITC Unavailability':'str',
                        'Source Type':'str',
                        'Generation Date':'str',
                        'IRN':'str'
                    }, parse_dates=['Document Date', 'Original Document Date', 'GSTR-1/5/6 Filing Date', 'Original Invoice Date', '2B-GenerationDate'], thousands=',')
                    
                if filetype == 'csv':
                    df = pd.read_csv(filedata, dtype={
                        'Return Period':'str',
                        'Recipient GSTIN':'str',
                        'Supplier GSTIN':'str',
                        'Supplier Name':'str',
                        'Document Type':'str',
                        'Supply Type':'str',
                        'Document Number':'str',
                        'Document Date':'str',
                        'Taxable Value ':'float',
                        'Tax Rate':'float',
                        'IGST Amount':'float',
                        'CGST Amount':'float',
                        'SGST Amount':'float',
                        'CESS Amount':'float',
                        'Invoice Value':'float',
                        'POS':'str',
                        'State Name':'str',
                        'Line Number':'str',
                        'BOE-ReferenceDate(ICEGATE)':'str',
                        'BOE-Received Date(GSTN)':'str',
                        'PortCode':'str',
                        'Bill Of Entry Number':'str',
                        'Bill Of Entry Date':'str',
                        'BOE-Amendment Flag':'str',
                        'Original Document Number':'str',
                        'Original Document Date':'str',
                        'Original Document Type':'str',
                        'Original Invoice Number':'str',
                        'Original Invoice Date':'str',
                        '2B-GenerationDate':'str',
                        'GSTR-1/5/6 Filing Period':'str',
                        'GSTR-1/5/6 Filing Date':'str',
                        'Differential Percentage':'str',
                        'Reverse Charge Flag':'str',
                        'ITC Availability':'str',
                        'Reason for ITC Unavailability':'str',
                        'Source Type':'str',
                        'Generation Date':'str',
                        'IRN':'str'
                        
                    }, parse_dates=['Document Date', 'Original Document Date', 'GSTR-1/5/6 Filing Date', 'Original Invoice Date', '2B-GenerationDate'], encoding='utf-8', thousands=',')

                df.drop(index=clientPAN_check(df, 'Recipient GSTIN'), inplace=True)
                if df.shape[0] != 0:
                    df = df.reset_index(drop=True)

                df['Document Date'] = pd.to_datetime(
                    df['Document Date'], dayfirst=True)
                df['Original Invoice Date'] = pd.to_datetime(
                    df['Original Invoice Date'], dayfirst=True)
                df['Document Number'] = np.where(df['Document Number'].map(
                    str).str[0:1] == "'", df['Document Number'].map(str).str[1:], df['Document Number'].map(str))
                df['Original Document Number'] = np.where(df['Original Document Number'].map(
                    str).str[0:1] == "'", df['Original Document Number'].map(str).str[1:], df['Original Document Number'].map(str))
                df['Original Document Number'] = np.where(df['Original Document Number'].map(
                    str) == "nan", "", df['Original Document Number'])
                df['Original Invoice Number'] = np.where(df['Original Invoice Number'].map(
                    str).str[0:1] == "'", df['Original Invoice Number'].map(str).str[1:], df['Original Invoice Number'].map(str))
                df['Original Invoice Number'] = np.where(df['Original Invoice Number'].map(
                    str) == "nan", "", df['Original Invoice Number'])
                df['GSTR-1/5/6 Filing Date'] = pd.to_datetime(
                    df['GSTR-1/5/6 Filing Date'], dayfirst=True)
                df['GSTR-2B category'] = ""
                df['GSTR-2B category'] = np.where(df['Document Type']
                                                == "C", "cdn", df['GSTR-2B category'])
                df['GSTR-2B category'] = np.where(df['Document Type']
                                                == "D", "dbn", df['GSTR-2B category'])
                df['GSTR-2B category'] = np.where(df['Document Type']
                                                == "R", "b2b", df['GSTR-2B category'])
                df['GSTR-2B category'] = np.where(np.logical_and(df['Document Type'] == "R",
                                                df['Original Document Number'].str.len() > 0), "b2ba", df['GSTR-2B category'])
                df['GSTR-2B category'] = np.where(np.logical_and(df['Document Type'] == "C", np.logical_and(
                    df['Original Document Number'].str.len() > 0, df['Original Invoice Number'].str.len() > 0)), "cdna", df['GSTR-2B category'])
                df['GSTR-2B category'] = np.where(np.logical_and(df['Document Type'] == "D", np.logical_and(
                    df['Original Document Number'].str.len() > 0, df['Original Invoice Number'].str.len() > 0)), "dbna", df['GSTR-2B category'])
                df['GSTR-2B category'] = np.where(df['Document Type']
                                                == "IMPG", "impg", df['GSTR-2B category'])
                df['Document Number'] = np.where(
                    df['Document Type'] == "IMPG", df['Original Invoice Number'], df['Document Number'])
                df['Document Date'] = np.where(
                    df['Document Type'] == "IMPG", df['Original Invoice Date'], df['Document Date'])
                df['GSTR-1/5/6 Filing Date'] = np.where(
                    df['Document Type'] == "IMPG", df['BOE-ReferenceDate(ICEGATE)'], df['GSTR-1/5/6 Filing Date'].map(str))
                df['Supplier GSTIN'] = np.where(
                    df['Document Type'] == "IMPG", "Dummy", df['Supplier GSTIN'])
                df['Document Type'] = np.where(
                    df['Document Type'] == "IMPG", df['PortCode'], df['Document Type'])
                # df.drop(columns=['Line Number', 'Supplier Name', 'State Name', 'ITC Availability', 'BOE-ReferenceDate(ICEGATE)', 'PortCode'], inplace=True)
                # # print(df.columns)
                df['Document Date'] = pd.to_datetime(
                    df['Document Date'], dayfirst=True)
                df['FY'] = df['Document Date'].map(
                    lambda x: x.year if x.month > 3 else x.year-1)
                duplicate = df[df.duplicated()]
                duplicate['Reason'] = "Duplicate"
                df.drop_duplicates(keep='first', inplace=True)
                duplicate.to_csv("GAPS\GSTR2B - Duplicate Drops.csv", index=False)
                df = df.groupby(['Supplier GSTIN','Supplier Name', 'Document Number', 'FY', 'GSTR-2B category', 'Document Type'], as_index=False).agg({'Return Period': 'first', 'Recipient GSTIN': 'first', 'Document Type': 'first', 'Document Number': 'first', 'Document Date': 'first', 'Original Invoice Number': 'first', 'Original Invoice Date': 'first', 'Supplier GSTIN': 'first', 'POS': 'first',
                'Taxable Value ': 'sum', 'Tax Rate': 'first', 'IGST Amount': 'sum', 'CGST Amount': 'sum', 'SGST Amount': 'sum', 'CESS Amount': 'sum', 'Invoice Value': 'sum', 'Reverse Charge Flag': 'first', 'Differential Percentage': 'first', 'GSTR-1/5/6 Filing Date': 'first', 'GSTR-1/5/6 Filing Period': 'first', '2B-GenerationDate': 'first', 'GSTR-2B category': 'first','ITC Availability':'first'})
                df.drop(columns=['FY'], inplace=True)
                df = df.reindex()
                df = df[['Return Period', 'Recipient GSTIN','Document Type', 'Document Number', 'Document Date', 'Original Invoice Number', 'Original Invoice Date', 'Supplier GSTIN','Supplier Name', 'POS', 'Taxable Value ', 'Tax Rate', 'IGST Amount', 'CGST Amount', 
                'SGST Amount', 'CESS Amount', 'Invoice Value', 'Reverse Charge Flag', 'Differential Percentage', 'GSTR-1/5/6 Filing Date', 'GSTR-1/5/6 Filing Period', '2B-GenerationDate', 'GSTR-2B category','ITC Availability']]
                neworder = ['TaxPeriod', 'CustomerGSTIN', 'Supply Type', 'DocumentNo', 'DocumentDate', 'Original Invoice Number', 'Original Invoice Date', 'SupplierGSTIN','SupplierName', 'POS', 'Taxable Value', 'Tax Rate', 'IGST Amount', 'Central Tax Amount', 
                'State/UT Tax Amount', 'CessAmount', 'InvoiceValue', 'ReverseCharge', 'Differential %', 'Filing Date', 'Filing Period', '2B Gen-Date', 'GSTR-2B category','ITC Availability']
                # # print(df)
                df.columns = neworder
                df['Taxable Value'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['Taxable Value'] > 0), df['Taxable Value']*-1, df['Taxable Value'])
                df['Central Tax Amount'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['Central Tax Amount'] > 0), df['Central Tax Amount']*-1, df['Central Tax Amount'])
                df['State/UT Tax Amount'] = np.where(np.logical_and(df['Supply Type'] == "cdn",
                                                    df['State/UT Tax Amount'] > 0), df['State/UT Tax Amount']*-1, df['State/UT Tax Amount'])
                df['IGST Amount'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['IGST Amount'] > 0), df['IGST Amount']*-1, df['IGST Amount'])
                df['CessAmount'] = np.where(np.logical_and(
                    df['Supply Type'] == "cdn", df['CessAmount'] > 0), df['CessAmount']*-1, df['CessAmount'])

                #df['Original Invoice Date']= pd.to_datetime(df['Original Invoice Date'],dayfirst=True)
                df['DocumentNo'] = df['DocumentNo'].map(str).str.upper()
                df['Original Invoice Number'] = df['Original Invoice Number'].map(str).str.upper()
                
                # df['Original Invoice Number'] = np.where(df['Original Invoice Number'].map(str) != '', df['Original Invoice Number'].map(str).str.upper(),df['Original Invoice Number'])

                try:
                    df['DocumentNo'] = df['DocumentNo'].str.lstrip('0')
                except:
                    pass

                try:
                    df['Original Invoice Number'] = df['Original Invoice Number'].str.lstrip(
                        '0')
                except:
                    pass
                df['FY'] = df['DocumentDate'].map(
                    lambda x: x.year if x.month > 3 else x.year-1)

                df['Tax'] = df['State/UT Tax Amount'] + \
                    df['Central Tax Amount']+df['IGST Amount'] + df['CessAmount']
                df['SourceFile'] = str(filename1)
                df['UploadTime'] = datetime.now()
                df['DocKey'] = df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(
                    str)+df['FY'].map(str)+df['Supply Type'].map(str)+df['GSTR-2B category'].map(str)
                totalcount = len(df)
                try:
                    drop = df[df['DocKey'].isin(doclist['DocKey'])]
                    drop.to_sql("GSTR_2Bdropped", panwisedb,
                                if_exists="append", index=False)
                    df.drop(df[df['DocKey'].isin(
                        doclist['DocKey'])].index, inplace=True)
                except:
                    pass
                processedcount = len(df)
                droppedcount = totalcount - processedcount

                # # print(df)
                duplicate = df[df.duplicated()]
                duplicate['Reason'] = "Duplicate"
                df.drop_duplicates(keep='first', inplace=True)
                #cdf1['FY'] = pd.to_datetime(cdf1['DocumentDate'],dayfirst=True,errors='coerce')
                #cdf1['FY'] = np.where(cdf1['FY'] != "NaT", cdf1['FY'].map(str).str[0:4],"")

                duplicate1 = df[df.duplicated(subset=[
                                            'SupplierGSTIN','SupplierName', 'DocumentNo', 'FY', 'GSTR-2B category', 'Supply Type'])]
                duplicate1['Reason'] = "Duplicate"
                df.drop_duplicates(subset=['SupplierGSTIN', 'DocumentNo', 'FY',
                                'GSTR-2B category', 'Supply Type'], keep='first', inplace=True)

                # df.drop(df[np.logical_or(df['GSTR-2B category'] == "impg",
                #         df['GSTR-2B category'] == "tds")].index, inplace=True)
                # df.drop(df[df['GSTR-2B category'] == "tdsa"].index, inplace=True)
                # df.reset_index(inplace=True)
                # print(df.columns)
                df['cdnpos'] = df['DocumentNo'].map(
                    str) + df['SupplierGSTIN'].map(str) + df['DocumentDate'].map(str)
                cdnpos = df.groupby('cdnpos', as_index=False).agg(
                    {'POS': 'first', 'GSTR-2B category': 'first'})
                cdnpos = cdnpos[cdnpos['GSTR-2B category'] == "b2b"]
                df['cdnpos'] = df['Original Invoice Number'].map(
                    str) + df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)

                df = pd.merge(df, cdnpos, on='cdnpos', how='left')
                # # print(cdf1.columns)
                df['POS_x'] = np.where(df['POS_x'].map(
                    str) == "nan", df['POS_y'], df['POS_x'])
                df = df.rename(
                    columns={'POS_x': 'POS', 'GSTR-2B category_x': 'GSTR-2B category'})
                df.drop(columns=['POS_y', 'GSTR-2B category_y'], inplace=True)
                #cdf1.drop(cdf1[~cdf1['CustomerGSTIN'].isin(gstins)].index, inplace=True)

                # df['Dropcfs'] = ""
                # df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2B category'] == "b2ba", df['CFS'] == "Y"), df['Original Invoice Number'].map(
                #     str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2B category'].str[0:3]+df['Supply Type'].map(str), df['Dropcfs'])
                # df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2B category'] == "cdna", df['CFS'] == "Y"), df['Original Invoice Number'].map(
                #     str)+df['SupplierGSTIN'].map(str)+df['Original Invoice Date'].map(str)+df['GSTR-2B category'].str[0:3]+df['Supply Type'].map(str), df['Dropcfs'])
                # df['Dropcfs'] = np.where(np.logical_and(df['GSTR-2B category'] == "cdna", df['CFS'] == "Y"), np.where(df['Dropcfs'] == "", df['DocumentNo'].map(
                #     str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category']+df['Supply Type'].map(str), df['Dropcfs']), df['Dropcfs'])
                # cfsdrop = df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category'] +
                #                             df['Supply Type'].map(str)).isin(df['Dropcfs']), np.logical_and(df['Dropcfs'] == "", df['Supply Type'].map(str) == "R"))]
                # cfsdrop['Reason'] = "Amended"
                # df.drop(df[np.logical_and((df['DocumentNo'].map(str)+df['SupplierGSTIN'].map(str)+df['DocumentDate'].map(str)+df['GSTR-2B category'] +
                #         df['Supply Type'].map(str)).isin(df['Dropcfs']), np.logical_and(df['Dropcfs'] == "", df['Supply Type'].map(str) == "R"))].index, inplace=True)
                # duplicate1 = duplicate.append(cfsdrop)

                # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
                #cdf1['Dropcfs'] = np.where(np.logical_and(cdf1['GSTR-2B category']=="cdna",cdf1['Dropcfs'] ==""),np.where(cdf1['CFS']=="Y",cdf1['DocumentNo'].map(str)+cdf1['SupplierGSTIN'].map(str)+cdf1['DocumentDate'].map(str)+cdf1['GSTR-2B category'].str[0:2],cdf1['Dropcfs'] ),cdf1['Dropcfs'] )

                #dropcfsn = df[df['CFS']=="N"]
                #df.drop(df[df['CFS']=="N"].index, inplace=True)
                #dropcfsn['Reason']="Not Filed"

                # cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)

                # duplicate2=duplicate1.append(dropcfsn)
                # duplicate1.to_csv("PR2A\Dropped 2A.csv", index=False)
                # duplicate.to_csv("PR2A\Dropped 2B.csv", index=False)
                df['DocumentNo'] = df['DocumentNo'].astype(
                    str).str.replace('\.0', '', regex=True)
                # cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
                #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)
                # # print(cdf1.dtypes)

                df.reset_index(inplace=True)
                df['first1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x).start(
                ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
                df['last1'] = df['DocumentNo'].apply(lambda x: re.search(r'[\\/*?:."<(-)>|]', x[::-1]).start(
                ) if re.search(r'[\\/*?:."<(-)>|]', x) else re.search(r'[\\/*?:."<(-)>|]', x))
                #df['last1']= len(df['DocumentNo']) - df['last1']
                # # print(df['first1'])
                df['first1'] = df['first1'].fillna(
                    df['DocumentNo'].str.len()).astype(int)
                df['last1'] = df['last1'].fillna(
                    df['DocumentNo'].str.len()).astype(int)
                df['DocNoBfrSpl'] = [DocumentNo[:first1]
                                    for DocumentNo, first1 in zip(df.DocumentNo, df.first1)]
                df['DocNoAftrSpl'] = [DocumentNo[-last1:]
                                    for DocumentNo, last1 in zip(df.DocumentNo, df.last1)]
                df['DocNoWOSplChar'] = df['DocumentNo'].str.replace(
                    '[^a-zA-Z0-9]', '', regex=True)
                df['DocNoNumeric'] = df['DocumentNo'].str.replace(
                    r"[a-zA-Z]", '').replace('\W', '').replace('_', '').replace('/', '').replace("\\", '')
                # df['DocNoBfrSpl'] =
                df['DocNo'] = df['DocumentNo']
                # df['DocNoAftrSpl'] =
                df['SuppGSTIN'] = df['SupplierGSTIN']
                df['SuppPAN'] = df['SupplierGSTIN'].str[2:12].map(str)
                df['ResGSTIN'] = df['CustomerGSTIN']
                df['ResPAN'] = df['CustomerGSTIN'].str[2:12].map(str)
                df['DocDate'] = df['DocumentDate'].apply(
                    lambda x: x.strftime('%d%m%Y') if x else "")
                df['MMYYYY'] = df['DocumentDate'].apply(
                    lambda x: x.strftime('%m%Y') if x else "")
                #df['FY'] = pd.to_datetime(df['DocumentDate'],dayfirst=True).apply(lambda x: x.strftime('%Y') if x else "")
                df['TaxVal'] = df['Taxable Value'].round(2)
                df['InvVal'] = df['InvoiceValue'].round(2)
                df['CGST'] = df['Central Tax Amount'].round(2)
                df['SGST'] = df['State/UT Tax Amount'].round(2)
                df['IGST'] = df['IGST Amount'].round(2)
                df['GST'] = df['Central Tax Amount'].round(2) + df['State/UT Tax Amount'].round(
                    2) + df['IGST Amount'].round(2) + df['CessAmount'].round(2)
                
                pd.to_numeric(df['POS'])
                df['POS'] = df['POS'].map(str)
                # df['POS'] == df["POS"].apply(lambda x: x.replace(
                #     r'.0', '') if isinstance(x, str) else x)
                df['POS'] = df["POS"].apply(lambda x: x.replace(
                    r'.0', '') if isinstance(x, str) else x)
                # df['POS'] = df["POS"].apply(lambda x: '0' + str(x) if isinstance(x, str) and len(x) == 1 else x)
                df['POS'] = df["POS"].apply(lambda x: x[-1] if isinstance(x, str) and len(x) == 2 and x[0] == '0' else x)
                # print(df['POS'].unique())

                df['RCM'] = df['ReverseCharge']
                df['DocType'] = np.where(df['Supply Type'] == "C", "cdn", np.where(
                    df['Supply Type'] == "D", "dbn", "b2b"))
                df['NA'] = ""

                df['Key1'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
                    str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key2'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(
                    str)+df['Central Tax Amount'].round(2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key3'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key4'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key5'] = df['SupplierGSTIN'].map(
                    str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key6'] = df['SupplierGSTIN'].map(
                    str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key7'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['Taxable Value'].round(2).map(str)+df['Central Tax Amount'].round(
                    2).map(str)+df['State/UT Tax Amount'].round(2).map(str)+df['IGST Amount'].round(2).map(str)+df['CessAmount'].round(2).map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key8'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
                    str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
                df['Key9'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocDate'].map(
                    str)+df['Taxable Value'].round(2).map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)
                df['Key10'] = df['FY'].map(str) + df['SuppPAN'].map(str) + df['DocNoWOSplChar'].map(
                    str)+df['ResPAN'].map(str)+df['GST'].round(2).map(str)+df['DocType'].map(str)

                df['Key11'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocumentNo'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key12'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['DocNoWOSplChar'].map(
                    str)+df['CustomerGSTIN'].map(str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key13'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
                    str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key14'] = df['FY'].map(str) + df['SupplierGSTIN'].map(
                    str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key15'] = df['SupplierGSTIN'].map(
                    str)+df['DocumentNo'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key16'] = df['SupplierGSTIN'].map(
                    str)+df['DocNoWOSplChar'].map(str)+df['CustomerGSTIN'].map(str)+df['DocType'].map(str)
                df['Key19'] = df['FY'].map(str) + df['DocNoWOSplChar'].map(
                    str)+df['ResPAN'].map(str)+df['DocDate'].map(str)+df['DocType'].map(str)
                df['Key20'] = df['FY'].map(str) + df['SuppPAN'].map(
                    str) + df['DocNoWOSplChar'].map(str)+df['ResPAN'].map(str)+df['DocType'].map(str)

                df['Key17'] = df['FY'].map(str) + df['SupplierGSTIN'].map(str)+df['CustomerGSTIN'].map(
                    str)+df['DocDate'].map(str)+df['POS'].map(str)+df['DocType'].map(str)
                df['Key18'] = df['FY'].map(str) + df['SuppPAN'].map(str)+df['ResPAN'].map(
                    str)+df['DocDate'].map(str)+df['DocType'].map(str)
                # print("Saved 2B Default Keys")
                try:
                    query = "select * FROM matchconfig"
                    config = pd.read_sql_query(query, panwisedb)
                    rowdata = dataframe_to_rows(config, index=False, header=False)
                    rowdata1 = dataframe_to_rows(config, index=False, header=False)
                    rowdata2 = dataframe_to_rows(config, index=False, header=False)
                    matchcon = 1
                except:
                    matchcon = 0
                    config = []
                l = len(config)
                n = 11

                try:
                    for row1 in rowdata1:

                        key = "Key" + str(n)
                        df[key] = df[row1[1]].map(str) + df[row1[2]].map(str) + df[row1[3]].map(str) + df[row1[4]].map(
                            str) + df[row1[5]].map(str) + df[row1[6]].map(str) + df[row1[7]].map(str) + df[row1[8]].map(str)
                        ## print("prkey ",row[0])
                        n += 1
                except:
                    pass
                # print("Saved User Defined Keys")
                df.to_sql("GSTR_2B", panwisedb, if_exists="append", index=False)
                # tcdf1=tcdf.copy()
                #tcdf1.drop(tcdf[tcdf['DocumentDate']<=date1].index, inplace=True)

                # df.to_csv(r"GAPS\GSTR2b.csv", index=False)

            try:
                query = "select Key1 from GSTR_2B"
                doclist2b = pd.read_sql_query(query, panwisedb)
            except:
                doclist2b = pd.DataFrame(columns=['NoKey'])
            if "NoKey" in doclist2b.columns:
                router.message = "Sanitizing GSTR 2B"
                try:
                    gstr2bsanitize()
                except:
                    pass
            elif any(doclist2b['Key1'].isna()) or any(doclist2b['Key1'] == ""):
                router.message = "Sanitizing GSTR 2B"
                try:
                    gstr2bsanitize()
                except:
                    pass
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    # current_user = current_user.get('current_user')
    path = dir_path + '/Client-Details'

    if upldtype == 'GL Dump':
        try:
            # print('GL Dump being uploaded')
            try:
                query = "SELECT * FROM GLCodeDigi"
                glc = pd.read_sql_query(query, panwisedb)
                # glc = glcode_master
            except:
                # flash('GL Code Master not Uploaded', "error")
                pass
            try:
                query = "SELECT * FROM GSTINDigi"
                bp = pd.read_sql_query(query, panwisedb)
                # bp = gstin_master
            except:
                # flash('GSTIN Master not Uploaded', "error")
                pass
            
            if filetype == "csv":
                tcdf = pd.DataFrame()
                for chunk in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    tcdf = pd.concat([tcdf, chunk])
            elif filetype == "xlsx":
                tcdf = pd.read_excel(filedata, engine='openpyxl')

            try:
                odf = tcdf.reindex(columns=colop)
                odf.columns = neworder
            except:
                odf = tcdf
            
            # if set(columns) == set(gldumpsap):
            #     odf['G/L_GL'] = odf['G/L Account']
            #     odf['Reference_GL'] = odf['Reference']
            #     odf['Accounting_Document_Number_GL'] = odf['Document Number']
            #     odf['Document_Type_GL'] = odf['Document Type']
            #     odf['Document_Date_GL'] = odf['Document Date']
            #     odf['Amount_GL'] = odf['Amount in local currency']
            #     odJf['Text_GL'] = odf['Text']
            #     odf['Clearing_Document_Number_GL'] = odf['Clearing Document']
            #     odf['Plant_GL'] = ''
            #     odf['Posting_Date_GL'] = odf['Posting Date']
            #     odf['Company_Code_GL'] = ''
            #     odf['Customer_Code_GL'] = odf['Offsetting acct no.']
            #     odf['Vendor_Code_GL'] =odf['Offsetting acct no.']
            #     odf['Entry_Date_GL'] = odf['Clearing date']
            #     odf['BP_GL'] = odf['Profit Center']
            #     odf['Offset_Account_GL'] = odf['Offsetting acct no.']
            #     odf['Transaction_Type_GL'] = odf['Tax code']
            #     odf['Period_GL'] = ''
            #     odf['Year/month_GL'] = odf['Year/month']
                
            #convert_dict = {'SourceIdentifier':'str', 'SourceFileName':'str', 'GLAccountCode':'str', 'Division':'str', 'SubDivision':'str', 'ProfitCentre1':'str', 'ProfitCentre2':'str', 'PlantCode':'str', 'ReturnPeriod':'str', 'RecipientGSTIN':'str', 'DocumentType':'str', 'SupplyType':'str', 'DocumentNumber':'str', 'DocumentDate':'str', 'OriginalDocumentNumber':'str', 'OriginalDocumentDate':'str', 'CRDRPreGST':'str', 'LineNumber':'str', 'SupplierGSTIN':'str', 'OriginalSupplierGSTIN':'str', 'SupplierName':'str', 'SupplierCode':'str', 'POS':'str', 'PortCode':'str', 'BillOfEntry':'str', 'BillOfEntryDate':'str', 'CIFValue':'float', 'CustomDuty':'float', 'HSNorSAC':'str', 'ItemCode':'str', 'ItemDescription':'str', 'CategoryOfItem':'str', 'UnitOfMeasurement':'str', 'Quantity':'float', 'TaxableValue':'float', 'IntegratedTaxRate':'float', 'IntegratedTaxAmount':'float', 'CentralTaxRate':'float', 'CentralTaxAmount':'float', 'StateUTTaxRate':'float', 'StateUTTaxAmount':'float', 'CessRateAdvalorem':'float', 'CessAmountAdvalorem':'float', 'CessRateSpecific':'float', 'CessAmountSpecific':'float', 'InvoiceValue':'float', 'ReverseChargeFlag':'str', 'EligibilityIndicator':'str', 'CommonSupplyIndicator':'str', 'AvailableIGST':'float', 'AvailableCGST':'float', 'AvailableSGST':'float', 'AvailableCess':'float', 'ITCReversalIdentifier':'str', 'ReasonForCreditDebitNote':'str', 'PurchaseVoucherNumber':'str', 'PurchaseVoucherDate':'str', 'PaymentVoucherNumber':'str', 'PaymentDate':'str', 'ContractNumber':'str', 'ContractDate':'str', 'ContractValue':'float', 'Userdefinedfield1':'str', 'Userdefinedfield2':'str', 'Userdefinedfield3':'str'}
            odf['Amount_GL'] = odf['Amount_GL'].apply(
                lambda x: x.strip() if isinstance(x, str) else x)
            odf['Amount_GL'] = odf['Amount_GL'].apply(
                lambda x: x.replace(r',', '') if isinstance(x, str) else x)
            #odf['Amount_GL']=odf['Amount_GL'].apply(lambda x: x.replace(r'-', '') if isinstance(x,str) else x)
            odf['Amount_GL'] = odf['Amount_GL'].apply(
                lambda x: 0 if x == '' else x)
            odf['Amount_GL'] = odf['Amount_GL'].replace(r's+', 0, regex=True)
            odf['Amount_GL'] = odf['Amount_GL'].fillna(0)
            odf['Amount_GL'] = odf['Amount_GL'].astype(float).round(2)

            try:
                odf['Posting_Date_GL'] = pd.to_datetime(
                    (odf['Posting_Date_GL'] - 25569) * 86400.0, unit='s', errors='ignore')
            except:
                pass
            try:
                odf['Document_Date_GL'] = pd.to_datetime(
                    (odf['Document_Date_GL'] - 25569) * 86400.0, unit='s', errors='ignore')
            except:
                pass
            try:

                odf['Entry_Date_GL'] = pd.to_datetime(
                    (odf['Entry_Date_GL'] - 25569) * 86400.0, unit='s', errors='ignore')
            except:
                pass
            try:
                odf['Posting_Date_GL'] = pd.to_datetime(
                    odf['Posting_Date_GL'], errors='ignore')
            except:
                pass
            try:
                odf['Document_Date_GL'] = pd.to_datetime(
                    odf['Document_Date_GL'], errors='ignore')
            except:
                pass
            try:
                odf['Entry_Date_GL'] = pd.to_datetime(
                    odf['Entry_Date_GL'], errors='ignore')
            except:
                pass
            # odf['Amount_GL']=odf['Amount_GL']*-1
            #odf['Source_Reg'] = os.path.basename(f).split('.')[0]
            try:
                # odf['Nature_GL'] = np.where(odf['G/L_GL'].map(str).isin(glc['Revenue GLs']), "Output", "Input")
                odf['Nature_GL'] = np.where(odf['G/L_GL'].isin(glc['Input Tax GLs']),"Input","Output")
            except:
                pass

            odf = odf.reindex()
            # # print(odf)
            #odf.drop(odf[odf['Accounting_Document_Number_GL'].isna()], inplace=True)

            odf = odf.round(2)
            odf['Reference_GL'] = odf['Reference_GL'].astype(str)
            i = 0

            odf['Reference_GL'] = odf['Reference_GL'].astype(str)

            odf['Amount_GL'] = odf['Amount_GL'].astype(float)
            try:
                odf['Posting_Date_GL'] = pd.to_datetime(
                    (odf['Posting_Date_GL'] - 25569) * 86400.0, unit='s', errors='ignore')
            except:
                pass
            try:
                odf['Document_Date_GL'] = pd.to_datetime(
                    (odf['Document_Date_GL'] - 25569) * 86400.0, unit='s', errors='ignore')
            except:
                pass
            try:

                odf['Entry_Date_GL'] = pd.to_datetime(
                    (odf['Entry_Date_GL'] - 25569) * 86400.0, unit='s', errors='ignore')
            except:
                pass
            try:
                odf['Posting_Date_GL'] = pd.to_datetime(
                    odf['Posting_Date_GL'], errors='ignore')
            except:
                pass
            try:
                odf['Document_Date_GL'] = pd.to_datetime(
                    odf['Document_Date_GL'], errors='ignore')
            except:
                pass
            try:
                odf['Entry_Date_GL'] = pd.to_datetime(
                    odf['Entry_Date_GL'], errors='ignore')
            except:
                pass

            try:
                odf['BP_GL'] = odf['BP_GL'].astype(str)
                bp['BP_GL'] = bp['BP_GL'].astype(str)
                odf = pd.merge(odf, bp, on='BP_GL', how='left')
            except:
                pass

            try:
                odf['GL_Type'] = ""
                #odf = pd.merge(odf,tcm,on='Transaction_Type_GL',how='left')
                #odf = pd.merge(odf,dtm,left_on='Document_Type_GL',right_on='Document_Type',how='left')
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['Revenue GLs']), "Revenue", odf['GL_Type'])
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['Taxable Advance (Liability) GLs']), "Advance", odf['GL_Type'])
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['Forex Gls Part of revenue']), "Forex", odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['CGST_Output']), "CGST_GL", odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['SGST_Output']), "SGST_GL", odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['IGST_Output']), "IGST_GL", odf['GL_Type'])
            except:
                pass
            
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['UGST_Output']), "UGST_GL", odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['CGST_Input']),"CGST_GL_Input",odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['SGST_Input']),"SGST_GL_Input",odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['IGST_Input']),"IGST_GL_Input",odf['GL_Type'])
            except:
                pass
            try:
                odf['GL_Type'] = np.where(odf['G/L_GL'].map(str).isin(glc['UGST_Input']),"SGST_GL_Input",odf['GL_Type'])
            except:
                pass
            
            odf['Amount_GL'] = odf['Amount_GL'].astype(float).round(2)
            #tcdf = tcdf.astype(convert_dict,errors='ignore')
            odf.to_sql("GLDump", panwisedb, if_exists="append", index=False)
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "GL Dump Azure" or upldtype == "Sales Register Azure":
        try:
            database_table_name = 'GLDump' if upldtype == "GL Dump Azure" else 'SalesRegisterDigi'
            if filetype == 'csv':
                for odf in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    odf.to_sql(database_table_name, panwisedb, if_exists="append", index=False)
            elif filetype == 'xlsx':
                if is_local:
                    headers = open(filedata)
                    headers = openpyxl.load_workbook(BytesIO(b64decode(headers.read())),read_only=True,data_only=True)
                else:
                    headers = urllib.request.urlopen(filedata).read()
                    headers = openpyxl.load_workbook(filename = BytesIO(headers))
                ws = headers[headers.sheetnames[0]]
                data = ws.values
                df = pd.DataFrame(data, columns=columns)
                df.to_sql(database_table_name, panwisedb, if_exists="append", index=False)
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'GSTIN':
        # print('GSTIN being uploaded')
        try:
            if filetype == 'csv':
                for df in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    df.to_sql("GSTINDigi", panwisedb, if_exists="append", index=False)
            else:
                try:
                    headers = open(filedata)
                    headers = openpyxl.load_workbook(BytesIO(b64decode(headers.read())),read_only=True,data_only=True)
                except:
                    headers = pd.read_excel(filedata)
                headers.to_sql("GSTINDigi", panwisedb, if_exists="append", index=False)
            c = panwisedb.cursor()
            c.execute("UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'GSTIN'")
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'Supply Type':
        # print('Supply Type being uploaded')
        try:
            if filetype == 'csv':
                for df in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    df.to_sql("SupplyTypeDigi", panwisedb, if_exists="replace", index=False)
            else:
                try:
                    headers = open(filedata)
                    headers = openpyxl.load_workbook(BytesIO(b64decode(headers.read())),read_only=True,data_only=True)
                except:
                    headers = pd.read_excel(filedata)
                headers.to_sql("SupplyTypeDigi", panwisedb, if_exists="append", index=False)
            c = panwisedb.cursor()
            c.execute("UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'Supply Type'")
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'Doc Type':
        # print('Document Type being uploaded')
        try:
            if filetype == 'csv':
                for df in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    df.to_sql("DocTypeDigi", panwisedb, if_exists="replace", index=False)
            else:
                try:
                    headers = open(filedata)
                    headers = openpyxl.load_workbook(BytesIO(b64decode(headers.read())),read_only=True,data_only=True)
                except:
                    headers = pd.read_excel(filedata)
                headers.to_sql("DocTypeDigi", panwisedb, if_exists="append", index=False)
            c = panwisedb.cursor()
            c.execute(
                "UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'Doc Type'")
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'GL Code':
        try:
            if filetype == 'csv':
                df = pd.read_csv(filedata, encoding='utf-8')
            else:
                df = pd.read_excel(filedata)
            if set(columns) == set(glcodemastersap):
                df['CGST_Output'] = df['CGST Tax GL Codes']
                df['SGST_Output'] = df['SGST Tax GL Code']
                df['IGST_Output'] = df['IGST Tax GL Code']
                df['UGST_Output'] = df['UGST Tax GL Code']
                df['Expense GL'] = ''
                df['CGST_Input'] = ''
                df['SGST_Input'] = ''
                df['IGST_Input'] = ''
                df['UGST_Input'] = ''
            df = df.fillna('')
            df = df.applymap(lambda x: str(x).split('.')[0] if '.' in str(x) else x)
            df.to_sql("GLCodeDigi", panwisedb,if_exists="replace", index=False)
            c = panwisedb.cursor()
            c.execute("UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'GL Code'")
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'Tax Code':
        # print('Tax Code being uploaded')
        try:
            if filetype == 'csv':
                for df in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    df.to_sql("TaxCodeDigi", panwisedb, if_exists="replace", index=False)
            else:
                try:
                    headers = open(filedata)
                    headers = openpyxl.load_workbook(BytesIO(b64decode(headers.read())),read_only=True,data_only=True)
                except:
                    headers = pd.read_excel(filedata)
                headers.to_sql("TaxCodeDigi", panwisedb, if_exists="append", index=False)
            c = panwisedb.cursor()
            c.execute("UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'Tax Code'")
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == 'TB':
        try:
            # print('Trial Balance being uploaded')
            try:
                for df in pd.read_csv(filedata, chunksize=10000, encoding='utf-8'):
                    df.to_sql("TrialBalanceDigi", panwisedb,
                            if_exists="replace", index=False)
                c = panwisedb.cursor()
                c.execute(
                    "UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'TB'")
                panwisedb.commit()

            except:
                for df in pd.read_excel(filedata, chunksize=10000, encoding='utf-8'):
                    df.to_sql("TrialBalanceDigi", panwisedb,
                            if_exists="replace", index=False)
                c = panwisedb.cursor()
                c.execute(
                    "UPDATE masters SET Status='Loaded' WHERE masters.Particulars = 'TB'")
                panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "Electronic Credit Ledger":
        try:
            if filetype == "csv":
                df = pd.read_csv(filedata, encoding='utf-8',
                                header=None, skiprows=8, index_col=False)
                credit_header = ['Srno', 'Date', 'Reference_No', 'Tax_Period', 'Description', 'Transaction_Type', 'Credit_Debit_Integrated', 'Credit_Debit_Central',
                                'Credit_Debit_State', 'Credit_Debit_Cess', 'Credit_Debit_Total', 'Balance_Integrated', 'Balance_Central', 'Balance_State', 'Balance_Cess', 'Balance_Total']
                df = df.iloc[:, 0:16]
                df.columns = credit_header
                #df.to_csv('creditLedger.csv', index=False, mode='a')
                df.to_sql("CreditLedger", panwisedb,
                        if_exists="append", index=False)
            elif filetype == "xlsx":
                df = pd.read_excel(filedata, header=None,
                                skiprows=8, index_col=False)
                credit_header = ['Srno', 'Date', 'Reference_No', 'Tax_Period', 'Description', 'Transaction_Type', 'Credit_Debit_Integrated', 'Credit_Debit_Central',
                                'Credit_Debit_State', 'Credit_Debit_Cess', 'Credit_Debit_Total', 'Balance_Integrated', 'Balance_Central', 'Balance_State', 'Balance_Cess', 'Balance_Total']
                df = df.iloc[:, 0:16]
                df.columns = credit_header
                df.to_sql("CreditLedger", panwisedb,
                        if_exists="append", index=False)
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "Electronic Cash Ledger":
        try:
            df = pd.read_csv(filedata, encoding='utf-8',
                            header=None, skiprows=8)
            cash_ledger_header = ['Srno', 'Date', 'Time_of_Deposit', 'Reporting_Date_by_Bank',
                                'Reference_No', 'Tax_Period', 'Description', 'Transaction_Type',

                                'Credit_Debit_Integrated_Tax', 'Credit_Debit_Integrated_Interest', 'Credit_Debit_Integrated_Penalty',
                                'Credit_Debit_Integrated_Fee', 'Credit_Debit_Integrated_Others', 'Credit_Debit_Integrated_Total',

                                'Credit_Debit_Central_Tax', 'Credit_Debit_Central_Interest', 'Credit_Debit_Central_Penalty',
                                'Credit_Debit_Central_Fee', 'Credit_Debit_Central_Others', 'Credit_Debit_Central_Total',

                                'Credit_Debit_State_Tax', 'Credit_Debit_State_Interest', 'Credit_Debit_State_Penalty',
                                'Credit_Debit_State_Fee', 'Credit_Debit_State_Others', 'Credit_Debit_State_Total',

                                'Credit_Debit_Cess_Tax', 'Credit_Debit_Cess_Interest', 'Credit_Debit_Cess_Penalty',
                                'Credit_Debit_Cess_Fee', 'Credit_Debit_Cess_Others', 'Credit_Debit_Cess_Total',

                                'Balance_Integrated_Tax', 'Balance_Integrated_Interest', 'Balance_Integrated_Penalty',
                                'Balance_Integrated_Fee', 'Balance_Integrated_Others', 'Balance_Integrated_Total',

                                'Balance_Central_Tax', 'Balance_Central_Interest', 'Balance_Central_Penalty',
                                'Balance_Central_Fee', 'Balance_Central_Others', 'Balance_Central_Total',

                                'Balance_State_Tax', 'Balance_State_Interest', 'Balance_State_Penalty',
                                'Balance_State_Fee', 'Balance_State_Others', 'Balance_State_Total',

                                'Balance_Cess_Tax', 'Balance_Cess_Interest', 'Balance_Cess_Penalty',
                                'Balance_Cess_Fee', 'Balance_Cess_Others', 'Balance_Cess_Total'
                                ]
            df = df.iloc[:, 0:56]
            df.columns = cash_ledger_header
            df.to_csv('cashLedger.csv', index=False, mode='a')
            df.to_sql("CashLedger", panwisedb, if_exists="append", index=False)
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "FIRC":
        try:
            df = pd.read_csv(filedata, parse_dates=['Export Invoice Date', 'Shipping Bill Date',
                                                            'Export General Manifest Date', 'FIRC Date'], encoding='utf-8')
            modified_columns = ['SupplierGSTIN', 'InvoiceType', 'CustomerName', 'ExportInvoiceNo',
                                'ExportInvoiceDate', 'TaxableValueGSTR1(FC)',
                                'ExportExchangeRate', 'ExportExchangeCurrency', 'Goods/Services',
                                'ShippingBillNo.', 'ShippingBillDate', 'PortCode',
                                'ExportGeneralManifestNo', 'Export General Manifest Date',
                                'FOBVALUE', 'FIRCNo', 'FIRCDate', 'FIRCAmountINR',
                                'FIRCAmount(FC)', 'FIRCExchangeRate', 'EY1', 'EY2']
            # print("Print Columns of FIRC dataframe", df.columns, len(df.columns))
            # print("Print Columns of FIRC dataframe", modified_columns, len(modified_columns))
            df.columns = modified_columns
            # print(df.head(5))
            df.to_sql("FIRC", panwisedb, if_exists="append", index=False)
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "Statement3":
        try:
            parsed_json = json.loads(filename)
            new_dict = {}
            for entry in parsed_json['stmt03']:
                if 'sno' in entry.keys():
                    if entry['sno'] in new_dict.keys():
                        new_dict[entry['sno']].update(entry)
                    else:
                        new_dict[entry['sno']] = entry
            sorted_dict = {i: new_dict[i] for i in sorted(new_dict)}
            df = pd.json_normalize(sorted_dict.values())
            df.to_sql("Statement3", panwisedb, if_exists="append", index=False)
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "Due Date Master":
        try:
            df = pd.read_csv(filedata, parse_dates=[
                            'Due date'], encoding='utf-8', dayfirst=True)
            df['Due date'] = pd.to_datetime(df['Due date'], dayfirst=True)

            df = df.pivot(index='TaxPeriod', columns='ReturnType', values='Due date').rename_axis(
                None, axis=1).reset_index(drop=False)
            df.to_sql("duedate", panwisedb, if_exists="replace", index=False)
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "Return Filing Master":
        try:
            df = pd.read_csv(filedata, parse_dates=[
                            'FilingDate', 'Date of registration'], encoding='utf-8', dayfirst=True)
            df['Date of registration'] = pd.to_datetime(
                df['Date of registration'], dayfirst=True)
            df['FilingDate'] = pd.to_datetime(df['FilingDate'], dayfirst=True)

            df = df.groupby(['GSTIN', 'TaxPeriod', 'GSTIN Status', 'Status', 'ReturnType']).FilingDate.first().unstack().rename_axis(None, axis=1).reset_index(drop=False)
            df['test'] = df['GSTR1'].dt.quarter
            g = df.groupby(by=['GSTIN', df['GSTR1'].dt.year])['test'].agg(Count='count', Sum='sum') # .size().to_frame('test' = 'sum').reset_index() # .reset_index(drop=True)
            g.index = g.index.map(lambda x: x[0] + str(x[1]))
            merged = pd.merge(df, g, left_on=df['GSTIN'] + df['GSTR1'].dt.year.map(str), right_on=g.index, how='left')
            del merged['key_0']
            merged['M/Q'] = np.where(np.logical_and(merged['Sum'] == 10, merged['Count'] == 4), 'Q', 'M')


            merged.to_sql("returnfiling", panwisedb, if_exists="append", index=False)
            panwisedb.commit()
        except:
            logging.error(f"process_file : {upldtype}", exc_info=True)

    if upldtype == "":
        return

    df = pd.DataFrame(columns=['File Name', 'Type', 'Extention', 'Time Stamp', 'Count', 'Uploaded', 'Dropped'], data=[[filename1, upldtype, filetype, datetime.now(), totalcount, processedcount, droppedcount]])

    df.to_sql("filelist", panwisedb, if_exists="append", index=False)
    panwisedb.commit()
    panwisedb.close()

def triggerval(clientPAN, current_user, record, starttime, uid):
    print(clientPAN, current_user)
    router.message = "Flash Validation Started......"
    dir_path = os.path.dirname(os.path.realpath(__file__))
    # path = dir_path + '\Client-Details'
    path = os.path.join(dir_path, 'Client-Details')
    router.message = "Running Flash validation on Inward Data...."
    # lst = ["False", "False", "False", "False", "False", "False", "False"]

    queuedb = sqlite3.connect('queue.db')
    cursor = queuedb.cursor()

    # queue_files_path = path + f'/{current_user}/{clientPAN}/Queue_files/'
    queue_files_path = os.path.join(path, current_user, clientPAN, 'Queue_files')

    # for file_path in os.listdir(queue_files_path):
    # number_of_records = pd.read_sql_query(f'''select clientPAN from queue where uid = "{uid}";''', queuedb).shape[0]
    file_path = os.listdir(queue_files_path)[0]
    try:
        validation_starttime = str(datetime.now())
        # process_file(lst, queue_files_path +'/'+ file_path, clientPAN, current_user)
        process_file(True, [clientPAN, current_user, queue_files_path +'/'+ file_path])
        updatecounters(clientPAN, current_user)
        # uploadsummary(clientPAN, current_user)
        uploadsummary(True, [clientPAN, current_user])
        if record:
            cursor.execute(f'''update queue set status="done" where uid = "{uid}";''')
            df = pd.DataFrame({
                'uid': [uid],
                'username' : [current_user],
                'clientPAN' : [clientPAN],
                'activity' : ['validating'],
                'status': ['done'],
                'size' : [str(len(os.listdir(queue_files_path)) -1)],
                'recdtime': [''],
                'starttime': [validation_starttime],
                'endtime': [''],
                'timetaken': [''],
                'recon_type': [''],
                'recon_parameters': [''],
                'report_parameters' : ['']
            })
            df.to_sql('queue', queuedb, index=False, if_exists='append')
        else:
            number_of_files = len(os.listdir(queue_files_path))
            update_query = f'''update queue set activity = "validating", status = "{'done' if number_of_files > 1 else 'all done'}", size = "{str(number_of_files -1)}", endtime = "{str(datetime.now())}", timetaken = "{'' if number_of_files > 1 else str(datetime.now() - datetime.strptime(starttime, '%Y-%m-%d %H:%M:%S.%f'))}" where uid = "{uid}" and activity = "validating";''' # if len(os.listdir(queue_files_path)) > 1 else f'''update queue set activity = "validated", status = "all done" where uid = "{uid}";'''
            cursor.execute(update_query)
        try:
            os.remove(queue_files_path +'/'+ file_path)
        except:
            traceback.print_exc()
    except:
        traceback.print_exc()
        cursor.execute(f'''update queue set activity = "validating", status = "failed", size = "{str(number_of_files -1)}", endtime = "{str(datetime.now())}", timetaken = "{'' if number_of_files > 1 else str(datetime.now() - datetime.strptime(starttime, '%Y-%m-%d %H:%M:%S.%f'))}" where uid = "{uid}" and activity = "validating";''')
    queuedb.commit()
    cursor.close()
    queuedb.close()
    print('Done')

def exportgaps(is_local, parameters):
    # is_local true if localhost else false if Azure Blob Storage
    # parameters will be clientPAN and current_user if localhost else temp_db_name
    dir_path = os.path.dirname(os.path.realpath(__file__))
    folder_name = '' if is_local else parameters[5]
    # path = dir_path + '/Client-Details' if is_local else dir_path + folder_name
    path = os.path.join(dir_path, 'Client-Details') if is_local else os.path.join(dir_path, folder_name)

    clientPAN = parameters[0] if is_local else ''
    current_user = parameters[1] if is_local else ''
    temp_db_name = '' if is_local else parameters[0]

    # panwisedb_path = f'{path}/{current_user}/{clientPAN}/{clientPAN}.db' if is_local else f'{path}/{temp_db_name}.db'
    panwisedb_path = os.path.join(path, current_user, clientPAN, f'{clientPAN}.db') if is_local else os.path.join(path, f'{temp_db_name}.db')

    reportlist = parameters[2 if is_local else 1]
    fromdate = parameters[3 if is_local else 2]
    todate = parameters[4 if is_local else 3]
    gstinlist = parameters[5 if is_local else 4]
    gstinlist = tuple(gstinlist)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    save_dir = os.path.join(base_dir, "GAPS", "Reports", clientPAN) if is_local else os.path.join(base_dir, "GAPS", "Reports", temp_db_name)
    if not os.path.exists(save_dir):
        os.makedirs(save_dir, exist_ok=True)

    for report in reportlist:
        if report in ('Sales_Register_Processed', 'Sales_Register_Consolidated', 'PurchaseRegisterDigi', 'PurchaseRegisterDigiConso') :# != "on":
            dir_path = os.path.dirname(os.path.realpath(__file__))
            path = dir_path + '/Client-Details'

            panwisedb = sqlite3.connect(
                path + '/' + clientPAN+'.db', timeout=10)

            c = panwisedb.cursor()
            c.execute("SELECT name FROM sqlite_master WHERE type='table'")

            tables = c.fetchall()

            time1 = datetime.now()
            timestamp = time1.strftime("%d-%m-%Y-%H-%M")

            gstinlist = tuple(gstinlist)
            # print(gstinlist)
            i = 1
            if fromdate[0] == "" or todate[0] == "":
                query = "select * FROM " + str(report)
                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000):
                    try:
                        cdf.drop(cdf[~cdf['RecipientGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    except:
                        cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    #cdf = cdf.loc[(cdf['Rate Check'] == 0) | (cdf['POS Check'] == "0") | (cdf['GSTIN Valid'] == 0) | (cdf['Credit Note'] == 0) | (cdf['Export days condition']==0) | (cdf['HSN'] ==0) | (cdf['Period']=="2017-18")]
                    cdf.to_csv(save_dir + '/' + report + timestamp +
                                ".csv", index=False, header=i, mode='a')
                    i = 0
            else:
                query = "select * FROM " + \
                    str(report) + " where DocumentDate >=? AND DocumentDate <=?"

                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000, params=(fromdate[0], todate[0],)):
                    try:
                        cdf.drop(cdf[~cdf['RecipientGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    except:
                        cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    #cdf = cdf.loc[(cdf['Rate Check'] == 0) | (cdf['POS Check'] == "0") | (cdf['GSTIN Valid'] == 0) | (cdf['Credit Note'] == 0) | (cdf['Export days condition']==0) | (cdf['HSN'] ==0) | (cdf['Period']=="2017-18")]
                    cdf.to_csv(save_dir + '/' + report + timestamp +
                                ".csv", index=False, header=i, mode='a')
                    i = 0

        if report == 'Purchase_Register_Error':
            dir_path = os.path.dirname(os.path.realpath(__file__))
            # path = dir_path + '/Client-Details'
            path = os.path.join(dir_path, 'Client-Details')

            # panwisedb = sqlite3.connect(
            #     path + '/' + clientPAN+'.db', timeout=10)
            panwisedb = sqlite3.connect(os.path.join(path, f'{clientPAN}.db'), timeout=10)

            time1 = datetime.now()
            timestamp = time1.strftime("%d-%m-%Y-%H-%M")

            gstinlist = tuple(gstinlist)
            # print(gstinlist)
            i = 1
            if fromdate[0] == "" or todate[0] == "":
                query = "select * FROM PurchaseRegisterDigi WHERE `Flash Remark` IS NOT NULL OR `Flash Remark` != ''"
                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000):
                    try:
                        cdf.drop(cdf[~cdf['RecipientGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    except:
                        cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    #cdf = cdf.loc[(cdf['Rate Check'] == 0) | (cdf['POS Check'] == "0") | (cdf['GSTIN Valid'] == 0) | (cdf['Credit Note'] == 0) | (cdf['Export days condition']==0) | (cdf['HSN'] ==0) | (cdf['Period']=="2017-18")]
                    # cdf.to_csv(save_dir + '/' + report + timestamp +
                    #             ".csv", index=False, header=i, mode='a')
                    cdf.to_csv(os.path.join(save_dir, f'{report}{timestamp}.csv'), index=False, header=i, mode='a')
                    i = 0
            else:
                query = "select * FROM  where DocumentDate >=? AND DocumentDate <=? AND (`Flash Remark` IS NOT NULL OR `Flash Remark` = '')"

                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000, params=(fromdate[0], todate[0],)):
                    try:
                        cdf.drop(cdf[~cdf['RecipientGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    except:
                        cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    #cdf = cdf.loc[(cdf['Rate Check'] == 0) | (cdf['POS Check'] == "0") | (cdf['GSTIN Valid'] == 0) | (cdf['Credit Note'] == 0) | (cdf['Export days condition']==0) | (cdf['HSN'] ==0) | (cdf['Period']=="2017-18")]
                    # cdf.to_csv(save_dir + '/' + report + timestamp +
                    #             ".csv", index=False, header=i, mode='a')
                    cdf.to_csv(os.path.join(save_dir, f'{report}{timestamp}.csv'), index=False, header=i, mode='a')
                    i = 0

        if report == 'Sales_Register_Error':

            dir_path = os.path.dirname(os.path.realpath(__file__))
            # path = dir_path + '/Client-Details'
            path = os.path.join(dir_path, 'Client-Details')

            # panwisedb = sqlite3.connect(
            #     path + '/' + clientPAN+'.db', timeout=10)
            panwisedb = sqlite3.connect(os.path.join(path, f'{clientPAN}.db'), timeout=10)

            time1 = datetime.now()
            timestamp = time1.strftime("%d-%m-%Y-%H-%M")

            gstinlist = tuple(gstinlist)
            # print(gstinlist)
            i = 1
            if fromdate[0] == "" or todate[0] == "":
                query = "select * FROM Sales_Register_Processed WHERE `Flash Remark` IS NOT NULL AND `Flash Remark` != ''"
                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000):
                    cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(gstinlist)].index, inplace=True)
                    #cdf = cdf.loc[(cdf['Rate Check'] == 0) | (cdf['POS Check'] == "0") | (cdf['GSTIN Valid'] == 0) | (cdf['Credit Note'] == 0) | (cdf['Export days condition']==0) | (cdf['HSN'] ==0) | (cdf['Period']=="2017-18")]
                    # cdf.to_csv(save_dir + '/' + report + timestamp +".csv", index=False, header=i, mode='a')
                    cdf.to_csv(os.path.join(save_dir, f'{report}{timestamp}.csv'), index=False, header=i, mode='a')
                    i = 0
            else:
                query = "select * FROM Sales_Register_Processed where DocumentDate >=? AND DocumentDate <=? AND (`Flash Remark` IS NOT NULL OR `Flash Remark` != '')"

                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000, params=(fromdate[0], todate[0],)):
                    try:
                        cdf.drop(cdf[~cdf['RecipientGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    except:
                        cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(
                            gstinlist)].index, inplace=True)
                    #cdf = cdf.loc[(cdf['Rate Check'] == 0) | (cdf['POS Check'] == "0") | (cdf['GSTIN Valid'] == 0) | (cdf['Credit Note'] == 0) | (cdf['Export days condition']==0) | (cdf['HSN'] ==0) | (cdf['Period']=="2017-18")]
                    # cdf.to_csv(save_dir + '/' + report + timestamp +
                    #             ".csv", index=False, header=i, mode='a')
                    cdf.to_csv(os.path.join(save_dir, f'{report}{timestamp}.csv'), index=False, header=i, mode='a')
                    i = 0
    
        if report == 'srgl':
            try:
                dir_path = os.path.dirname(os.path.realpath(__file__))
                # path = dir_path + '/Client-Details'
                path = os.path.join(dir_path, 'Client-Details')
                time1 = datetime.now()
                timestamp = time1.strftime("%d-%m-%Y-%H-%M")
                # panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
                panwisedb = sqlite3.connect(panwisedb_path, timeout=10)
                cur = panwisedb.cursor()
                
                query = "SELECT * FROM GL_Dump_Consolidated"
                # cur.execute(query)
                # cdf=pd.DataFrame(cur.fetchall()) 
                # cdf.columns=[x[0] for x in cur.description]
                cdf = pd.read_sql_query(query, panwisedb)
                #cdf = cdf.drop_duplicates(subset=['Reference_GL'],keep='first')
                cdf.drop(cdf[~cdf['GSTIN'].isin(gstinlist)].index, inplace=True)
                
                cdf.drop(cdf[cdf['Nature_GL'] == "Input"].index, inplace=True)
                cdf['Accounting_Document_Number_Reg']=cdf['Accounting_Document_Number_Reg'].astype(str).str.replace('\.0', '', regex=True)
                #gstr2=cdf[cdf['Category'] == "Document Number of Tax, not in Revenue and Advance"] 
                # print(cdf['Accounting_Document_Number_Reg'])
                #cdf.drop(cdf[cdf['Category'] == "Document Number of Tax, not in Revenue and Advance"].index, inplace=True)
                #cdf.reset_index(inplace = True, drop = True)
                
                ## print(cdf.dtypes)
                query = "SELECT * FROM Sales_Register_Consolidated"
                # cur.execute(query)
                # cdf1=pd.DataFrame(cur.fetchall()) 
                # cdf1.columns=[x[0] for x in cur.description]
                cdf1 = pd.read_sql_query(query, panwisedb)
                #cdf1 = cdf1.drop_duplicates(subset=['DocumentNo'],keep='first')
                
                cdf1['AccountingVoucherNumber']=cdf1['AccountingVoucherNumber'].astype(str).str.replace('\.0', '', regex=True)
                cdf1.drop(cdf1[~cdf1['SupplierGSTIN'].isin(gstinlist)].index, inplace=True)
                # # print(cdf['AccountingVoucherNumber'])
                
                cdf2 = pd.merge(cdf,cdf1,how='outer',left_on='Accounting_Document_Number_Reg', right_on='AccountingVoucherNumber',indicator=True)
                
                cdf2.drop(cdf2.columns[cdf2.columns.str.contains('level')], axis=1, inplace=True)
                cdf2.drop(cdf2.columns[cdf2.columns.str.contains('index')], axis=1, inplace=True)
                
                cdf2.reset_index(inplace = True, drop = True)
                #cdf2=cdf2[cdf2['DocumentNo'].isnull() | ~cdf2[cdf2['DocumentNo'].notnull()].duplicated(subset='DocumentNo',keep='first')]
                cdf2=cdf2[cdf2['DocumentNumber'].isnull() | ~cdf2[cdf2['DocumentNumber'].notnull()].duplicated(subset='DocumentNumber',keep='first')]
                cdf2=cdf2[cdf2['Accounting_Document_Number_Reg'].isnull() | ~cdf2[cdf2['Accounting_Document_Number_Reg'].notnull()].duplicated(subset='Accounting_Document_Number_Reg',keep='first')]
            
                #cdf2['SR_vs_GL'] = ""
                #cdf2['SR_vs_GL'] = np.where(np.logical_and(np.logical_and(~cdf2['Accounting_Document_Number_GL'].isna(),~cdf2['Accounting_Document_Number_Reg'].isna()),cdf2['DocumentNo'].isna()),"Present in SR and GL",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL'] = np.where(np.logical_and(np.logical_and(~cdf2['Invoice_No_Reg'].isna(),~cdf2['DocumentNo'].isna()),cdf2['Accounting_Document_Number_GL'].isna()),"Present in SR and G1",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL'] = np.where(np.logical_and(np.logical_and(~cdf2['Invoice_No_Reg'].isna(),~cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_GL'].isna()),"Present in All",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Invoice_No_Reg'].isna(),cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_GL'].isna()),"Present in GL only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Accounting_Document_Number_GL'].isna(),cdf2['Accounting_Document_Number_Reg']=="nan"),~cdf2['DocumentNo'].isna()),"Present in GSTR 1 only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Accounting_Document_Number_GL'].isna(),cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_Reg'].isna()),"Present in SR only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['Amounts Match / Mismatch'] = np.where((cdf2['Difference in Taxable Amount']+cdf2['Difference in CGST Amount'] +cdf2['Difference in SGST Amount']+cdf2['Difference in IGST Amount']+cdf2['Difference in Cess Amount'])<1,"Match","Mis-Match")
                try:
                    try:
                        cdf2['TaxableValue'].fillna(0, inplace=True)
                        cdf2['Difference in Taxable Amount'] = abs(abs(cdf2['TaxableValue'].map(float)) - abs(cdf2['Revenue'].map(float)))
                    except:
                        cdf2['Difference in Taxable Amount'] = abs(abs(cdf2['TaxableValue'].map(float)) - 0)
                    try:
                        cdf2['IntegratedTaxAmount'].fillna(0, inplace=True)
                        cdf2['Difference in IGST Amount'] = abs(abs(cdf2['IntegratedTaxAmount'].map(float)) - abs(cdf2['IGST_GL'].map(float)))
                    except:
                        cdf2['Difference in IGST Amount'] = abs(abs(cdf2['IntegratedTaxAmount'].map(float)) - 0)
                    try:
                        cdf2['CentralTaxAmount'].fillna(0, inplace=True)
                        cdf2['Difference in CGST Amount'] = abs(abs(cdf2['CentralTaxAmount'].map(float)) - abs(cdf2['CGST_GL'].map(float)))
                    except:
                        cdf2['Difference in CGST Amount'] = abs(abs(cdf2['CentralTaxAmount'].map(float)) - 0)
                    try:
                        cdf2['StateUTTaxAmount'].fillna(0, inplace=True)
                        cdf2['Difference in SGST Amount'] = abs(abs(cdf2['StateUTTaxAmount'].map(float)) - abs(cdf2['SGST_GL'].map(float)))
                    except:
                        cdf2['Difference in SGST Amount'] = abs(abs(cdf2['StateUTTaxAmount'].map(float)) - 0)
                except:
                    pass

                try:
                    Diff_Taxable = cdf2['Difference in Taxable Amount'] if 'Difference in Taxable Amount' in cdf2.columns else 0
                    Diff_CSGT = cdf2['Difference in CGST Amount'] if 'Difference in CGST Amount' in cdf2.columns else 0 
                    Diff_SGST = cdf2['Difference in SGST Amount'] if 'Difference in SGST Amount' in cdf2.columns else 0
                    Diff_IGST = cdf2['Difference in IGST Amount'] if 'Difference in IGST Amount' in cdf2.columns else 0
                    s_check = Diff_Taxable + Diff_CSGT + Diff_SGST + Diff_IGST
                except:
                    s_check = 0
                
                cdf2['_merge'] = np.where(cdf2['_merge']=="both","Mis-Match",cdf2['_merge'])
                cdf2['Comment'] = ''
                # Iterate through each row in DataFrame to apply the conditions
                for index, row in cdf2.iterrows():
                    if (row['_merge']=="Mis-Match") and ((row['Difference in Taxable Amount'] + row['Difference in CGST Amount'] + row['Difference in SGST Amount'] + row['Difference in IGST Amount']) > 1):
                        comment = 'Value Mis-Match in'
                        comment += ' Taxable Amount,' if row['Difference in Taxable Amount'] != 0 else ''
                        comment += ' CGST Amount,' if row['Difference in CGST Amount'] != 0 else ''
                        comment += ' SGST Amount,' if row['Difference in SGST Amount'] != 0 else ''
                        comment += ' IGST Amount' if row['Difference in IGST Amount'] != 0 else ''
                    else:
                        comment = row['_merge']
                    # Update the comment in the DataFrame
                    cdf2.at[index, 'Comment'] = comment
                cdf2['_merge'] = np.where(np.logical_and(cdf2['_merge']=="Mis-Match",s_check<1),"Match",cdf2['Comment'])
                cdf2.drop(columns=['Comment'], inplace=True)
                cdf2['_merge'] = np.where(cdf2['_merge']=="left_only","Present only in GL",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="right_only","Present only in SR",cdf2['_merge'])
                cdf2 = cdf2.rename(columns={'_merge':'SR_vs_GL'})        

                if is_local:
                    # cdf2.to_csv(save_dir + '/' + report + timestamp + ".csv",index=False)
                    cdf2.to_csv(os.path.join(save_dir, f'{report}{timestamp}.csv'), index=False)
                    cdf2.to_sql("SRvsGL",panwisedb, if_exists="replace",index=False)
                else:
                    return cdf2
                
                # filename = r"C:\GAPS\Reports\SR vs GL.csv"
                # try:
                #     self.format1(filename)
            except:
                traceback.print_exc()
                pass
    
        if report == 'srg1':
            try:
                dir_path = os.path.dirname(os.path.realpath(__file__))
                path = dir_path + '/Client-Details'
                time1 = datetime.now()
                timestamp = time1.strftime("%d-%m-%Y-%H-%M")
                panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
                cur = panwisedb.cursor()
                
                query = "SELECT * FROM Sales_Register_Consolidated"
                cur.execute(query)
                cdf=pd.DataFrame(cur.fetchall()) 
                cdf.columns=[x[0] for x in cur.description]
                cdf = cdf.drop_duplicates(subset=['DocumentNumber'],keep='first')
                cdf.drop(cdf[~cdf['SupplierGSTIN'].isin(gstinlist)].index, inplace=True)
                
                cdf['DocumentNumber']=cdf['DocumentNumber'].astype(str,copy=True,errors='ignore')
                
                cdf.reset_index(inplace = True)
                
                query = "SELECT * FROM GSTR_1"
                cur.execute(query)
                cdf1=pd.DataFrame(cur.fetchall()) 
                cdf1.columns=[x[0] for x in cur.description]
                cdf1 = cdf1.drop_duplicates(subset=['DocumentNo'],keep='first')
                cdf1.drop(cdf1[~cdf1['SupplierGSTIN'].isin(gstinlist)].index, inplace=True)
                
                cdf1['Original Invoice Number']= cdf1['Original Invoice Number'].astype(str).str.replace('\.0', '', regex=True)
                

                cdf1['DocumentNo']= cdf1['DocumentNo'].astype(str).str.replace('\.0', '', regex=True)
                #cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
                #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)        
                # print(cdf1.dtypes)
                cdf1.drop(cdf1[((cdf1['GSTR-1 category']=="B2CS")|(cdf1['GSTR-1 category']=="B2CSA"))].index, inplace=True)
                cdf1.reset_index(inplace = True)

                
                cdf2 = pd.merge(cdf,cdf1,how='outer',left_on=['DocumentNumber'], right_on=['DocumentNo'],indicator=True)
                
                cdf2.reset_index(inplace = True)
                cdf2=cdf2[cdf2['DocumentNo'].isnull() | ~cdf2[cdf2['DocumentNo'].notnull()].duplicated(subset='DocumentNo',keep='first')]
                cdf2=cdf2[cdf2['DocumentNumber'].isnull() | ~cdf2[cdf2['DocumentNumber'].notnull()].duplicated(subset='DocumentNumber',keep='first')]
                cdf2['Difference in Taxable Amount'] = abs(cdf2['TaxableValue'].map(float) - cdf2['Taxable Value'].map(float))
                cdf2['Difference in CGST Amount'] = abs(cdf2['CentralTaxAmount'].map(float) - cdf2['Central Tax Amount'].map(float))
                cdf2['Difference in SGST Amount'] = abs(cdf2['StateUTTaxRate'].map(float) - cdf2['State/UT Tax Amount'].map(float))
                cdf2['Difference in IGST Amount'] = abs(cdf2['IntegratedTaxAmount'].map(float) - cdf2['IGST Amount'].map(float))
                cdf2['Difference in Cess Amount'] = abs(cdf2['CessAmountAdvalorem'].map(float) - cdf2['CessAmount'].map(float))
                cdf2['_merge'] = np.where(cdf2['_merge']=="both","Mis-Match",cdf2['_merge'])
                cdf2['_merge'] = np.where(np.logical_and(cdf2['_merge']=="Mis-Match",(cdf2['Difference in Taxable Amount']+cdf2['Difference in CGST Amount'] +cdf2['Difference in SGST Amount']+cdf2['Difference in IGST Amount']+cdf2['Difference in Cess Amount'])<1),"Match",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="left_only","Present only in SR",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="right_only","Present only in GSTR 1",cdf2['_merge'])
                cdf2 = cdf2.rename(columns={'_merge':'SR_vs_G1'})
        
                cdf2.to_csv(save_dir + '/' + report + timestamp + ".csv",index=False)
                cdf2.to_sql("SRvsG1",panwisedb, if_exists="replace",index=False)

            except:
                pass

        if report == 'srg1gl':
            try:
                dir_path = os.path.dirname(os.path.realpath(__file__))
                path = dir_path + '/Client-Details'
                time1 = datetime.now()
                timestamp = time1.strftime("%d-%m-%Y-%H-%M")
                panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
                cur = panwisedb.cursor()
                
                query = "SELECT * FROM GL_Dump_Consolidated"
                cur.execute(query)
                cdf=pd.DataFrame(cur.fetchall()) 
                cdf.columns=[x[0] for x in cur.description]
                #cdf = cdf.drop_duplicates(subset=['Reference_GL'],keep='first')
                cdf.drop(cdf[~cdf['GSTIN'].isin(gstinlist)].index, inplace=True)
                

                cdf['Accounting_Document_Number_Reg']=cdf['Accounting_Document_Number_Reg'].astype(str).str.replace('\.0', '', regex=True)
                gstr2=cdf[cdf['Category'] == "Document Number of Tax, not in Revenue and Advance"] 
                
                cdf.drop(cdf[cdf['Category'] == "Document Number of Tax, not in Revenue and Advance"].index, inplace=True)
                cdf.reset_index(inplace = True, drop = True)
                
                ## print(cdf.dtypes)
                query = "SELECT * FROM SRvsG1"
                cur.execute(query)
                cdf1=pd.DataFrame(cur.fetchall()) 
                cdf1.columns=[x[0] for x in cur.description]
                #cdf1 = cdf1.drop_duplicates(subset=['DocumentNo'],keep='first')
                cdf1.drop(cdf1[~cdf1['SupplierGSTIN_x'].isin(gstinlist) & ~cdf1['SupplierGSTIN_y'].isin(gstinlist)].index, inplace=True)
                

                cdf1['AccountingVoucherNumber']=cdf1['AccountingVoucherNumber'].astype(str).str.replace('\.0', '', regex=True)     
                gstr1=cdf1[cdf1['SR_vs_G1'] == "Present only in GSTR 1"] 
                cdf1.drop(cdf1[cdf1['SR_vs_G1'] == "Present only in GSTR 1"].index, inplace=True)
                
                cdf1.reset_index(inplace = True, drop = True)
                
                # print(cdf.dtypes)
                # print(cdf1.dtypes)
                # print(gstr1.dtypes)
                # print(gstr2.dtypes)
                #df = df.drop_duplicates(subset=['Column1','Column2'],keep='first')
                #cdf1.to_csv(r"C:\GAPS\Reports\Dump2.csv",index=False)
                #cdf2=pd.DataFrame()
                #neworder = cdf.columns + cdf1.columns
                #cdf2=cdf2.reindex(columns=neworder)
                
                cdf2 = pd.merge(cdf,cdf1,how='outer',left_on='Accounting_Document_Number_Reg', right_on='AccountingVoucherNumber',indicator=True)
                
                cdf2=cdf2.append(gstr1)
                cdf2=cdf2.append(gstr2)
                cdf2.drop(cdf2.columns[cdf2.columns.str.contains('level')], axis=1, inplace=True)
                cdf2.drop(cdf2.columns[cdf2.columns.str.contains('index')], axis=1, inplace=True)
                
                cdf2.reset_index(inplace = True, drop = True)
                cdf2=cdf2[cdf2['DocumentNo'].isnull() | ~cdf2[cdf2['DocumentNo'].notnull()].duplicated(subset='DocumentNo',keep='first')]
                cdf2=cdf2[cdf2['DocumentNumber'].isnull() | ~cdf2[cdf2['DocumentNumber'].notnull()].duplicated(subset='DocumentNumber',keep='first')]
                
                cdf2['SR_vs_GL_vs_G1'] = ""
                cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(~cdf2['Accounting_Document_Number_GL'].isna(),~cdf2['Accounting_Document_Number_Reg'].isna()),cdf2['DocumentNo'].isna()),"Present in SR and GL",cdf2['SR_vs_GL_vs_G1'])
                cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(~cdf2['DocumentNumber'].isna(),~cdf2['DocumentNo'].isna()),cdf2['Accounting_Document_Number_GL'].isna()),"Present in SR and G1",cdf2['SR_vs_GL_vs_G1'])
                cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(~cdf2['DocumentNumber'].isna(),~cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_GL'].isna()),"Present in All",cdf2['SR_vs_GL_vs_G1'])
                cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['DocumentNumber'].isna(),cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_GL'].isna()),"Present in GL only",cdf2['SR_vs_GL_vs_G1'])
                cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Accounting_Document_Number_GL'].isna(),cdf2['Accounting_Document_Number_Reg']=="nan"),~cdf2['DocumentNo'].isna()),"Present in GSTR 1 only",cdf2['SR_vs_GL_vs_G1'])
                cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Accounting_Document_Number_GL'].isna(),cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_Reg'].isna()),"Present in SR only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['Amounts Match / Mismatch'] = np.where((cdf2['Difference in Taxable Amount']+cdf2['Difference in CGST Amount'] +cdf2['Difference in SGST Amount']+cdf2['Difference in IGST Amount']+cdf2['Difference in Cess Amount'])<1,"Match","Mis-Match")
                
                cdf2['Difference in Taxable Amount'] = np.where(np.logical_and(cdf2['TaxableValue']==cdf2['Revenue'],cdf2['Revenue']==cdf2['Taxable Value']),0,1)
                cdf2['Difference in CGST Amount'] = np.where(np.logical_and(cdf2['CentralTaxAmount']==cdf2['CGST_GL'],cdf2['CGST_GL']==cdf2['Central Tax Amount']),0,1)
                cdf2['Difference in SGST Amount'] = np.where(np.logical_and(cdf2['StateUTTaxRate']==cdf2['SGST_GL'],cdf2['SGST_GL']==cdf2['State/UT Tax Amount']),0,1)
                cdf2['Difference in IGST Amount'] = np.where(np.logical_and(cdf2['IntegratedTaxAmount']==cdf2['IGST_GL'],cdf2['IGST_GL']==cdf2['IGST Amount']),0,1)
                cdf2['Amounts Match / Mismatch'] = np.where((cdf2['Difference in Taxable Amount']+cdf2['Difference in CGST Amount'] +cdf2['Difference in SGST Amount']+cdf2['Difference in IGST Amount'])<1,"Match","Mis-Match")
                
            # cdf2.reset_index(inplace = True)
                #cdf2=cdf2[cdf2['DocumentNo'].isnull() | ~cdf2[cdf2['DocumentNo'].notnull()].duplicated(subset='DocumentNo',keep='first')]
                #cdf2=cdf2[cdf2['Reference_GL'].isnull() | ~cdf2[cdf2['Reference_GL'].notnull()].duplicated(subset='Reference_GL',keep='first')]
                
        
                cdf2.to_csv(save_dir + '/' + report + timestamp + ".csv",index=False)
                
                cdf2.to_sql("SRvsG1vsGL",panwisedb, if_exists="replace",index=False)
                # filename = r"C:\GAPS\Reports\SR vs GL vs G1.csv"
                # try:
                #     self.format1(filename)
                # finally:
                #     QMessageBox.information(QMessageBox(),'Successful','Sales Register vs GSTR 1 Vs GL Dump Analysis Completed')
            except:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # print(exc_type, exc_tb.tb_lineno)
        
        if report == 'prgl':
            try:
                dir_path = os.path.dirname(os.path.realpath(__file__))
                path = dir_path + '/Client-Details'
                time1 = datetime.now()
                timestamp = time1.strftime("%d-%m-%Y-%H-%M")
                panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
                cur = panwisedb.cursor()

                query = "SELECT * FROM GL_Dump_Consolidated"
                cur.execute(query)
                cdf=pd.DataFrame(cur.fetchall()) 
                cdf.columns=[x[0] for x in cur.description]
                #cdf = cdf.drop_duplicates(subset=['Reference_GL'],keep='first')
                cdf.drop(cdf[cdf['Nature_GL'] == "Output"].index, inplace=True)

                cdf['Accounting_Document_Number_Reg']=cdf['Accounting_Document_Number_Reg'].astype(str).str.replace('\.0', '', regex=True)
                #gstr2=cdf[cdf['Category'] == "Document Number of Tax, not in Revenue and Advance"] 
                cdf.drop(cdf[~cdf['GSTIN'].isin(gstinlist)].index, inplace=True)
                
                #cdf.drop(cdf[cdf['Category'] == "Document Number of Tax, not in Revenue and Advance"].index, inplace=True)
                #cdf.reset_index(inplace = True, drop = True)
                
                ## print(cdf.dtypes)
                query = "SELECT * FROM PurchaseRegisterDigiConso"
                cur.execute(query)
                cdf1=pd.DataFrame(cur.fetchall()) 
                cdf1.columns=[x[0] for x in cur.description]
                #cdf1 = cdf1.drop_duplicates(subset=['DocumentNo'],keep='first')
                

                cdf1['PaymentVoucherNumber']=cdf1['PaymentVoucherNumber'].astype(str).str.replace('\.0', '', regex=True)        
                cdf1.drop(cdf1[~cdf1['RecipientGSTIN'].isin(gstinlist)].index, inplace=True)
                
                cdf2 = pd.merge(cdf,cdf1,how='outer',left_on='Accounting_Document_Number_Reg',right_on='PaymentVoucherNumber',indicator=True)
                
                cdf2.drop(cdf2.columns[cdf2.columns.str.contains('level')], axis=1, inplace=True)
                cdf2.drop(cdf2.columns[cdf2.columns.str.contains('index')], axis=1, inplace=True)
                
                cdf2.reset_index(inplace = True, drop = True)
                #cdf2=cdf2[cdf2['DocumentNo'].isnull() | ~cdf2[cdf2['DocumentNo'].notnull()].duplicated(subset='DocumentNo',keep='first')]
                cdf2=cdf2[cdf2['DocumentNumber'].isnull() | ~cdf2[cdf2['DocumentNumber'].notnull()].duplicated(subset='DocumentNumber',keep='first')]
                
                #cdf2['SR_vs_GL'] = ""
                #cdf2['SR_vs_GL'] = np.where(np.logical_and(np.logical_and(~cdf2['Accounting_Document_Number_GL'].isna(),~cdf2['Accounting_Document_Number_Reg'].isna()),cdf2['DocumentNo'].isna()),"Present in SR and GL",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL'] = np.where(np.logical_and(np.logical_and(~cdf2['Invoice_No_Reg'].isna(),~cdf2['DocumentNo'].isna()),cdf2['Accounting_Document_Number_GL'].isna()),"Present in SR and G1",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL'] = np.where(np.logical_and(np.logical_and(~cdf2['Invoice_No_Reg'].isna(),~cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_GL'].isna()),"Present in All",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Invoice_No_Reg'].isna(),cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_GL'].isna()),"Present in GL only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Accounting_Document_Number_GL'].isna(),cdf2['Accounting_Document_Number_Reg']=="nan"),~cdf2['DocumentNo'].isna()),"Present in GSTR 1 only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['SR_vs_GL_vs_G1'] = np.where(np.logical_and(np.logical_and(cdf2['Accounting_Document_Number_GL'].isna(),cdf2['DocumentNo'].isna()),~cdf2['Accounting_Document_Number_Reg'].isna()),"Present in SR only",cdf2['SR_vs_GL_vs_G1'])
                #cdf2['Amounts Match / Mismatch'] = np.where((cdf2['Difference in Taxable Amount']+cdf2['Difference in CGST Amount'] +cdf2['Difference in SGST Amount']+cdf2['Difference in IGST Amount']+cdf2['Difference in Cess Amount'])<1,"Match","Mis-Match")
                try:
                    cdf2['Difference in Taxable Amount'] = abs(abs(cdf2['TaxVal'].map(float)) - abs(cdf2['Revenue'].map(float)))
        
                except:
                    pass
                try:    
                    cdf2['Difference in CGST Amount'] = abs(abs(cdf2['CGST'].map(float)) - abs(cdf2['CGST_GL'].map(float)))
                    cdf2['Difference in SGST Amount'] = abs(abs(cdf2['SGST'].map(float)) - abs(cdf2['SGST_GL'].map(float)))
                    cdf2['Difference in IGST Amount'] = abs(abs(cdf2['IGST'].map(float)) - abs(cdf2['IGST_GL'].map(float)))
                except:
                    pass
                
                cdf2['_merge'] = np.where(cdf2['_merge']=="both","Mis-Match",cdf2['_merge'])
                cdf2['_merge'] = np.where(np.logical_and(cdf2['_merge']=="Mis-Match",(cdf2['Difference in CGST Amount'] if 'Difference in CGST Amount' in cdf2.columns else 0 +cdf2['Difference in SGST Amount']if 'Difference in SGST Amount' in cdf2.columns else 0+cdf2['Difference in IGST Amount']if 'Difference in IGST Amount' in cdf2.columns else 0)<1),"Match",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="left_only","Present only in GL",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="right_only","Present only in PR",cdf2['_merge'])
                cdf2 = cdf2.rename(columns={'_merge':'PR_vs_GL'})        
        
                cdf2.to_csv(save_dir + '/' + report + timestamp + ".csv", index=False)
                cdf2.to_sql("PRvsGL",panwisedb, if_exists="replace",index=False)
                # filename = r"C:\GAPS\Reports\PR vs GL.csv"
                # try:
                #     self.format1(filename)
                # finally:
                #     QMessageBox.information(QMessageBox(),'Successful','Purchase Register vs GL Dump Analysis Completed')
            except:
                pass  

        if report == 'pr2a':
            try:
                dir_path = os.path.dirname(os.path.realpath(__file__))
                path = dir_path + '/Client-Details'
                time1 = datetime.now()
                timestamp = time1.strftime("%d-%m-%Y-%H-%M")
                panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
                cur = panwisedb.cursor()
                
                query = "SELECT * FROM PurchaseRegisterDigiConso"
                cur.execute(query)
                cdf=pd.DataFrame(cur.fetchall()) 
                cdf.columns=[x[0] for x in cur.description]
                cdf = cdf.drop_duplicates(subset=['DocumentNumber'],keep='first')
                cdf.drop(cdf[~cdf['RecipientGSTIN'].isin(gstinlist)].index, inplace=True)
                
                cdf['DocumentNumber']=cdf['DocumentNumber'].astype(str,copy=True,errors='ignore')
                #cdf['Invoice_No_Reg']=cdf['Invoice_No_Reg'].str.encode('utf-8')
                
                cdf.reset_index(inplace = True)
                # print(cdf.dtypes)
                #cdf.to_csv(r"C:\GAPS\Reports\Dump1.csv",index=False)
                
                query = "SELECT * FROM GSTR_2A"
                cur.execute(query)
                cdf1=pd.DataFrame(cur.fetchall()) 
                cdf1.columns=[x[0] for x in cur.description]
                cdf1 = cdf1.drop_duplicates(subset=['DocumentNo'],keep='first')
                cdf1.drop(cdf1[~cdf1['CustomerGSTIN'].isin(gstinlist)].index, inplace=True)
                
                #cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
                

                #cdf1['DocumentNo']=cdf1['DocumentNo'].round(0)
                cdf1['DocumentNo']= cdf1['DocumentNo'].astype(str).str.replace('\.0', '', regex=True)
                #cdf1['DocumentNo']=cdf1['DocumentNo'].str.encode('utf-8')
                #cdf1['DocumentNo']= cdf1['DocumentNo'].str.replace(".0","",1)        
                # print(cdf1.dtypes)
                # cdf1.reset_index(inplace = True)
                
                cdf2 = pd.merge(cdf,cdf1,how='outer',on='Key1',indicator=True)
                
                cdf2.reset_index(inplace = True)
                cdf2=cdf2[cdf2['DocumentNo'].isnull() | ~cdf2[cdf2['DocumentNo'].notnull()].duplicated(subset='DocumentNo',keep='first')]
                cdf2=cdf2[cdf2['DocumentNumber'].isnull() | ~cdf2[cdf2['DocumentNumber'].notnull()].duplicated(subset='DocumentNumber',keep='first')]
                cdf2['Difference in Taxable Amount'] = abs(cdf2['TaxableValue'].map(float) - cdf2['Taxable Value'].map(float))
                cdf2['Difference in CGST Amount'] = abs(cdf2['AvailableCGST'].map(float) - cdf2['Central Tax Amount'].map(float))
                cdf2['Difference in SGST Amount'] = abs(cdf2['AvailableSGST'].map(float) - cdf2['State/UT Tax Amount'].map(float))
                cdf2['Difference in IGST Amount'] = abs(cdf2['AvailableIGST'].map(float) - cdf2['IGST Amount'].map(float))
                cdf2['Difference in Cess Amount'] = abs(cdf2['AvailableCess'].map(float) - cdf2['CessAmountAdvalorem'].map(float))
                cdf2['_merge'] = np.where(cdf2['_merge']=="both","Mis-Match",cdf2['_merge'])
                cdf2['_merge'] = np.where(np.logical_and(cdf2['_merge']=="Mis-Match",(cdf2['Difference in Taxable Amount']+cdf2['Difference in CGST Amount'] +cdf2['Difference in SGST Amount']+cdf2['Difference in IGST Amount']+cdf2['Difference in Cess Amount'])<1),"Match",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="left_only","Present only in PR",cdf2['_merge'])
                cdf2['_merge'] = np.where(cdf2['_merge']=="right_only","Present only in GSTR 2A",cdf2['_merge'])
                cdf2 = cdf2.rename(columns={'_merge':'PR_vs_2A'})
        
                cdf2.to_csv(save_dir + '/' + report + timestamp + ".csv",index=False)
                cdf2.to_sql("PRvs2A",panwisedb, if_exists="replace",index=False)
                # filename = r"C:\GAPS\Reports\PR vs 2A.csv"
            except:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                # print(exc_type, exc_tb.tb_lineno)
        
        if report == 'gldumpprocessed':
            try:
                dir_path = os.path.dirname(os.path.realpath(__file__))
                path = dir_path + '/Client-Details'
                time1 = datetime.now()
                timestamp = time1.strftime("%d-%m-%Y-%H-%M")
                panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
                
                query = "SELECT * FROM GL_Dump_Consolidated"
                i = 1
                for cdf in pd.read_sql_query(query, panwisedb, chunksize=10000):
                    cdf.drop(cdf[~cdf['GSTIN'].isin(gstinlist)].index, inplace=True)
                    cdf.to_csv(save_dir + '/' + report + timestamp + ".csv",index=False, header=i, mode='a')
                    i = 0

            except:
                pass  
        if report == 'g1json':
            dir_path = os.path.dirname(os.path.realpath(__file__))
            path = dir_path + '/Client-Details'
            time1 = datetime.now()
            timestamp = time1.strftime("%d-%m-%Y-%H-%M")
            panwisedb = sqlite3.connect(path + '/' + clientPAN+'.db', timeout=10)
            
            try:
                # print('GSTR1->JSON')
                SupplierGSTIN = pd.read_sql_query('SELECT DISTINCT SupplierGSTIN FROM Sales_Register_Processed', panwisedb)['SupplierGSTIN'].to_list()
                for gstin in SupplierGSTIN:
                    # sr_to_json(gstin)
                    sheet_data = pd.read_sql_query(f"SELECT * FROM Sales_Register_Processed WHERE SupplierGSTIN == '{gstin}'", panwisedb)
                    sheet_details = OrderedDict()
                    # Date = datetime.strptime(sheet_data[0][3], '%d-%b-%Y')
                    # sheet_details['fp'] = str(Date.strftime("%m%Y"))
                    sheet_details['gstin'] = str(gstin)
                    sheet_details['hash'] = 'hash'
                    sheet_details['version'] = 'GST2.2.6'
                    gstins = sheet_data['CustomerGSTIN'].to_list()
                    # for row in sheet_data['CustomerGSTIN']:
                    #     gstins.append(row[0]) 
                    # # print(gstins)
                    data_list = []
                    gstin_index = 0
                    index = 0
                    while(index < len(sheet_data)):
                        gstin_details = OrderedDict()
                        # gstin_details_list = sheet_data[index]
                        gstin_details_list = sheet_data.iloc[index, :]
                        invoice_list=[]
                        while True:
                            invoice = {} 
                            gstin_details['ctin'] = gstin_details_list['CustomerGSTIN']
                            invoice['inum'] = gstin_details_list['DocumentNumber']
                            invoice_date = datetime.strptime(gstin_details_list['DocumentDate'], "%Y-%m-%d %H:%M:%S")
                            # gstin_details_list['DocumentDate']
                            invoice['idt'] = datetime.strftime(invoice_date,  '%d-%m-%Y')
                            # invoice['idt'] = gstin_details_list['DocumentDate'].apply(lambda x: datetime.strftime(x,  '%d-%m-%Y'))
                            invoice['val'] = gstin_details_list['InvoiceValue']
                            invoice['pos'] = gstin_details_list['POS']
                            invoice['rchrg'] = gstin_details_list['ReverseChargeFlag']
                            invoice['inv_typ'] = gstin_details_list['DocumentType']
                            num_list = []
                            num = OrderedDict()
                            num['num'] = 1201
                            items_details = OrderedDict()
                            items_details['txval'] = gstin_details_list['TaxableValue']
                            items_details['rt'] = int(gstin_details_list['TotalTaxRate'])
                            items_details['camt'] = ''
                            items_details['samt'] = ''
                            items_details['csamt'] = 0
                            num['itm_det'] = items_details
                            num_list.append(num)
                            invoice['itms'] = num_list
                            invoice_list.append(invoice)
                            gstin_index += 1
                            if ( gstin_index >= len(gstins)):
                                break
                            if (gstins[gstin_index-1] == gstins[gstin_index]):
                                index += 1
                                gstin_details = OrderedDict()
                                gstin_details_list = sheet_data.iloc[index, :]
                            else:
                                break
                        index += 1
                        gstin_details['inv'] = invoice_list
                        sheet_details['b2b'] = data_list
                        data_list.append(gstin_details)

                    json_data = json.dumps(sheet_details)
                    with open(save_dir + '/' + report + gstin + timestamp + '.json',  'w') as json_file:
                        json_file.write(json_data)
            except:
                pass
    return {"success": True}