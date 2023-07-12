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
        assigned_group = 'Invaild case id'

    else:
        data = response.content.decode('utf-8')

        soup = BeautifulSoup(data.replace('&amp;', "&"), 'xml')
        #print(soup)

        if soup.details is not None:
            details = soup.details.string
        else:
            details = ''

        if soup.assigned_group is not None:
            assigned_group = soup.assigned_group.string
        else:
            assigned_group = ''

    return details, assigned_group

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

def get_pior_ticket(assigned_group:str, asin:str, ticket:str):
    search_link = 'https://tt.amazon.com/search?category=RBS&type=China-QA&item=Automation&assigned_group={}&status=Assigned%3BResearching%3BWork+In+Progress%3BPending%3BResolved%3BClosed&impact=&assigned_individual=&requester_login=&login_name=&cc_email=&phrase_search_text={}&keyword_bq={}&exact_bq=&or_bq1=&or_bq2=&or_bq3=&exclude_bq=&create_date=&modified_date=&tags=&case_type=&building_id=&search=Search%21'.format(
        assigned_group, asin, asin)
    try:
        tempcookie = get_mwinit_cookie()
    except:
        tempcookie = get_mwinit_cookie()

    pior_dic = {}

    auth = HTTPKerberosAuth(mutual_authentication=OPTIONAL)
    session = requests.session()
    response = session.get(search_link,
                           auth=auth, proxies=None,
                           cert=None, timeout=None, verify=False, cookies=tempcookie)
    selector_ = etree.HTML(response.content.decode())
    tables = selector_.xpath("//table[@id='search_results']")
    df_ = pd.read_html(etree.tostring(tables[0]))[0] if len(tables) != 0 else None
    # df_.to_excel("./repeat.xlsx", index=False)
    if type(df_) == DataFrame and list(df_['Case ID']) != []:
        research_tt_list = list(df_['Case ID'])
        status_list = list(df_['Status'])
        for num in range(len(research_tt_list)):
            if "0" + str(research_tt_list[num]) != ticket and status_list[num] != 'Resolved' and status_list[num] != 'Closed':
                pior_dic["0" + str(research_tt_list[num])] = status_list[num]

    return pior_dic

if __name__ == '__main__':
    ticket = '0528727416'

    print('Ticket: '+ticket)
    details, assigned_group = get_info_from_tt(ticket)
    print('Assigned_group: '+assigned_group)
    asin = identify_asin(details)
    print('ASIN: '+asin)
    prior_dic = get_pior_ticket(assigned_group, asin, ticket)
    print('Current open TT: '+str(prior_dic))
    if len(prior_dic) != 0:
        for prior_tt, status in prior_dic.items():
            print('Resolve duplicate: ' + prior_tt)
            retry_put(prior_tt, {'correspondence': 'Closed as duplicate with https://tt.amazon.com/' + ticket,
                                 'status': 'Resolved',
                                 'root_cause': 'Web Tech'})
            time.sleep(2)
            prior_details, prior_assigned_group = get_info_from_tt(prior_tt)
            time.sleep(2)
            print('Comment on original: ' + ticket)
            comment = 'Automatically resolved duplicate TT: https://tt.amazon.com/'+prior_tt + '\nDetails:\n' + prior_details
            retry_put(ticket, {'correspondence': comment})
            time.sleep(2)

    print('----------------------Finished----------------------')