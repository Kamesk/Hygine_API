from bs4 import BeautifulSoup
import requests
import pandas as pd
from requests_kerberos import HTTPKerberosAuth, OPTIONAL
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from requests.packages.urllib3 import disable_warnings
from pandas import read_excel, read_csv, Series, DataFrame, isnull
import time
from email.header import Header
import os
import shutil
import re
import getpass
import openpyxl
import tkinter.messagebox
from lxml import etree
import numpy as np
import math
from email.mime.text import MIMEText
from smtplib import SMTP
import os
import smtplib
from email.mime.multipart import MIMEMultipart
import datetime
from io import StringIO
import pickle
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
from gevent import monkey
# monkey.patch_all(thread=False)
import gevent,time,requests
from gevent.queue import Queue
import sys


disable_warnings(InsecureRequestWarning)
auth = HTTPKerberosAuth(mutual_authentication=OPTIONAL)


def getHeaders(data=None, contentType=None, contentCharset=None):
    """
    Returns a dictionary of HTTP/1.1 headers to send to the Fluxo service.

    If data is specified, this adds the Content-Length, Content-Type, and
    Date headers to the request."""
    headers = {
        'Accept': "*/*",
        'Accept-Charset': "UTF-8",
        'Connection': "close",
        'Host': "ticket-api-test.amazon.com",
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                      "(KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36"
    }
    if data is not None:
        if contentType is None:
            contentType = "application/x-www-form-urlencoded"
        if contentCharset is None:
            contentCharset = "UTF-8"

        headers['Content-Length'] = str(len(data))
        headers['Content-Type'] = (contentType + "; charset=" +
                                   contentCharset)
        headers['Date'] = time.strftime("%a, %d %b %Y %H:%M:%S GMT",
                                        time.gmtime())

    return headers

def get_mwinit_cookie():
    #MidwayConfigDir1 = os.path.join(os.path.expanduser("~"), ".midway")
    #MidwayCookieJarFile1 = os.path.join(MidwayConfigDir1, "cookie1")
    MidwayConfigDir = os.path.join(os.path.expanduser("~"), ".midway")
    MidwayCookieJarFile = os.path.join(MidwayConfigDir, "cookie")
    fields = []
    '''
    try:
        keyfile = open(MidwayCookieJarFile, "r")
    except:
        shutil.copyfile(MidwayCookieJarFile1, MidwayCookieJarFile)
        keyfile = open(MidwayCookieJarFile, "r")
    '''
    keyfile = open(MidwayCookieJarFile, "r")
    for line in keyfile:
        # parse the record into fields (separated by whitespace)
        fields = line.split()
        if len(fields) != 0:
            # get the yubi session token and expire time
            if fields[0] == "#HttpOnly_midway-auth.amazon.com":
                session_token = fields[6].replace("\n", "")
                expires = fields[4]
            # get the user who generated the session token
            elif fields[0] == "midway-auth.amazon.com":
                username = fields[6].replace("\n", "")
    keyfile.close()
    # make sure the session token hasn't expired
    if time.gmtime() > time.gmtime(int(expires)):
        os.remove(MidwayCookieJarFile)
        raise SystemError("Your Midway token has expired. Run mwinit to renew")
    # construct the cookie value required by calls to k2
    try:
        cookie = {"username": username, "session": session_token}

    except:
        raise SystemError("Your Midway token has expired. Run mwinit to renew")
    return cookie

# ASIN, customer complaint, shipping FC
def get_info_from_tt(ticket):
    try:
        tempcookie = get_mwinit_cookie()
    except:
        tempcookie = get_mwinit_cookie()

    fluxo_auth = ("flx-JP-Andon", "JP-Andon-prod")
    response = requests.get("https://ticket-api.amazon.com/tickets/" + ticket,
                            auth=fluxo_auth, proxies=None, cookies=tempcookie,
                            cert=None, timeout=None, verify=False, headers=getHeaders())
    if response.status_code == 404:
        details = 'Invaild case id'

    else:
        data = response.content.decode('utf-8')

        soup = BeautifulSoup(data.replace('&amp;', "&"), 'xml')
        #print(soup)

        if soup.details is not None:
            details = soup.details.string
        else:
            details = ''

    return details

def identify_asin(details):
    keyword_list = ['ASIN:','ASIN Suppressed:','[ASIN]','Asin Detected:','ASIN :']
    asin = 'none'
    for keyword in keyword_list:
        try:
            m = re.search("({})()*[^\n]*".format(keyword), details)
            asin_line = str(m.group())
            asin_line_list = asin_line.split(keyword[-1])
            asin_filter = asin_line_list[len(asin_line_list) - 1].strip()
            if len(asin_filter) == 10:
                asin = asin_filter
                break
        except:
            pass
    return asin

def get_description(details):
    if "description:" in details:
        if "wizardLog" in details:
            start_index = str(details).index("description:")
            end_index = str(details).index("wizardLog")
            description = details[start_index:end_index]
            description = description.replace("wizardLog", "")
            return description
        else:
            start_index = str(details).index("description:")
            description = details[start_index:]
            return description
    elif "[Detail]" in details:
        start_index = str(details).index("[Detail]")
        description = details[start_index:]
        return description
    else:
        description = ""
        return description

def get_shipping_fc(details):
    fc_list = []
    try:
        m = re.search("(Shipping FC:)()*[^\n]*", details)
        str_shipping_fc = str(m.group())
        fc_list_line = str_shipping_fc.split(":")[-1].strip()
        if len(fc_list_line) == 4:
            fc_list.append(fc_list_line)
        elif len(fc_list_line) > 4:
            if "," in fc_list_line:
                fc_list = fc_list_line.split(",")
            else:
                fc_list = fc_list_line.split(" ")
        else:
            pass

    except:
        pass
    return fc_list

def get_inventory_condition(asin):
    try:
        tempcookie = get_mwinit_cookie()
    except:
        tempcookie = get_mwinit_cookie()

    # https://alaska-eu.amazon.com/index.html?viewtype=summaryview&use_scrollbars=&fnsku_simple=B079K9VMTT&marketplaceid=3&merchantid=9&AvailData=Get+Availability+Data
    status_code = ''
    times = 0
    inventory_dic = {}
    while status_code != 200 and times < 3:
        try:
            url = 'https://alaska-na.amazon.com/index.html?viewtype=summaryview&use_scrollbars=&fnsku_simple=B017Y8QMQQ&marketplaceid=1&merchantid=1&AvailData=Get+Availability+Data'.format(asin)
            res = requests.get(url,auth=auth, proxies=None,cert=None, timeout=None, verify=False, cookies=tempcookie)
            status_code = res.status_code
            selector = etree.HTML(res.content.decode())
            tr_list = selector.xpath('//*[@id="view"]/table[2]/tr')
            # print(tr_list)
            if len(tr_list) == 4:
                pass
            else:
                for tr in tr_list[2:-2]:
                    fc = tr.xpath("string(./td[1])").strip()
                    in_building = tr.xpath("string(./td[2])").strip()
                    if int(in_building) != 0:
                        inventory_dic[fc] = int(in_building)
            break
        except:
            time.sleep(2)
            times = times + 1
    return inventory_dic

def create_ticket(data: dict) -> (int, str):
    try:
        tempcookie = get_mwinit_cookie()
    except:
        tempcookie = get_mwinit_cookie()

    payload = {"status": "Assigned",
               "impact": "5",
               "requester_login": getpass.getuser(),
               "requester_name": getpass.getuser()}
    payload.update(data)
    session = requests.session()
    fluxo_auth = ("flx-JP-Andon", "JP-Andon-prod")
    response = session.post("https://ticket-api.amazon.com/tickets",
                                 data=payload, headers=getHeaders(),
                                 auth=fluxo_auth, proxies=None,
                                 cert=None, timeout=None, verify=False, cookies=tempcookie)
    return response

def cut_bincheck_tt(final_initial_fc:dict, asin:str, andon_ticket:str, customer_complaint:str):

    fluxo_auth = ("flx-JP-Andon", "JP-Andon-prod")
    short_description = "{} Bin Check for FnSku:{}　Andon TT：{}".format(
        list(final_initial_fc.keys()), asin, andon_ticket)

    fc_comment = ''
    for fc,units in final_initial_fc.items():
        fc_comment = fc_comment + str(fc) + ' ' + str(units) + '\n'

    details = '''

Retail ICQAチームの皆さん

お疲れ様です。

Issue：{}

確認依頼：在庫品に同問題があるかご確認お願い致します。

ASIN: {}

Retail FC 在庫数：
{}

どうぞ宜しくお願い致します。

RBS 
    '''.format(customer_complaint, asin, fc_comment)

    print(short_description)
    print('---------------------')
    print(details)

    cut_bincheck = create_ticket(
                                 data={
                    "short_description": short_description,
                    "details": details,
                    "category": "RBS",
                    "type": "China-QA",
                    "item": "Automation",
                    "assigned_group": "DEQA-Automation",
                    "assigned_individual": "yayou"}
    )
    times = 0
    status_code = cut_bincheck.status_code
    while times < 3 and status_code != 200:
        if status_code == 200:
            break
        time.sleep(5)
        times += 1

    ticket_number = None
    # https://ticket-api.amazon.com/tickets/0528240527
    ticketLocationRegEx = re.compile(
        r"http(?:s?)://[a-zA-Z0-9][-a-zA-Z0-9]*\." +
        r"(?:[-a-zA-Z0-9][-a-zA-Z0-9]*\.)*" +
        r"amazon\.com/tickets/([A-Z]?[a-z0-9\-]+) *")

    # Success is a 2xx response with a valid Location header.
    location = cut_bincheck.headers.get("Location")
    if location is not None:
        m = ticketLocationRegEx.match(location)
        if m is not None:
            ticket_number = m.group(1)

    return ticket_number

def update_info_to_andon(ticket,info):
    global tempcookie
    try:
        tempcookie = get_mwinit_cookie()
    except:
        tempcookie = get_mwinit_cookie()

    # data = {"run_as_user": "yayou"}
    data = {}
    data.update(info)
    fluxo_auth = ("flx-JP-Andon", "JP-Andon-prod")
    # print("https://ticket-api.amazon.com/tickets/" + str(ticket))
    response = requests.put("https://ticket-api.amazon.com/tickets/" + str(ticket),
                            data=data,
                            auth=fluxo_auth, proxies=None, cookies=tempcookie,
                            cert=None, timeout=None, verify=False, headers=getHeaders())
    return response.status_code

def retry_put(ticket,info):
    times = 0
    stas_code = 0
    while times < 3 and stas_code != 200:
        stas_code = update_info_to_andon(ticket,info)
        if stas_code == 200:
            break
        time.sleep(5)
        times += 1
    return stas_code

def consolidate_functions(ticket):
    details = get_info_from_tt(ticket)
    asin = identify_asin(details)
    customer_complaint = get_description(details)
    shipping_fc = get_shipping_fc(details)

    print(shipping_fc)
    print('----------------')

    final_initial_fc = {}

    inventory_dic = get_inventory_condition(asin)
    print(inventory_dic)
    print('----------------')

    if inventory_dic != {}:
        for fc,units in inventory_dic.items():
            if fc in shipping_fc:
                final_initial_fc[fc] = units
            else:
                pass

    print(final_initial_fc)
    print('----------------')

    if len(final_initial_fc) == 0:
        if len(inventory_dic) == 0:
            correspondence = 'Bin Check Result:\nThere is no inventory in any FC.'
        else:
            correspondence = 'Bin Check Result:\nThere is no inventory in shipping FC.\nRecent Inventory:\n' + str(
                inventory_dic)
    else:
        bin_check_number = cut_bincheck_tt(final_initial_fc, asin, ticket, customer_complaint)
        correspondence = 'Initialed Bin Check to:\n'+str(final_initial_fc)+'\nhttps://tt.amazon.com/' + str(bin_check_number)

    retry_put(ticket, {'correspondence': correspondence})
    # retry_put(ticket, {'work_log': correspondence})

if __name__ == '__main__':
    ticket = '0528240527'
    if len(ticket) != 10:
        ticket = '0' + ticket

    consolidate_functions(ticket)