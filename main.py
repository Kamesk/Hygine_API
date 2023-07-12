import requests
import os
import time
from bs4 import BeautifulSoup
import warnings
import re
from selenium import webdriver
from tkinter import messagebox
from requests_kerberos import HTTPKerberosAuth, OPTIONAL
import getpass


auth = HTTPKerberosAuth(mutual_authentication=OPTIONAL)
fluxo_auth = ("flx-JP-Andon", "JP-Andon-prod")


def get_mwinit_cookie():
    # MidwayConfigDir1 = os.path.join(os.path.expanduser("~"), ".midway")
    # MidwayCookieJarFile1 = os.path.join(MidwayConfigDir1, "cookie1")
    MidwayConfigDir = os.path.join(os.path.expanduser("~"), ".midway")
    # print(MidwayConfigDir)
    MidwayCookieJarFile = os.path.join(MidwayConfigDir, "cookie")
    # print(MidwayCookieJarFile)
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


def get_tempcookie():
    global tempcookie
    try:
        tempcookie = get_mwinit_cookie()
    except:
        tempcookie = get_mwinit_cookie()
    return tempcookie


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


def get_info_from_tickets(ticket):
    response = requests.get("https://ticket-api.amazon.com/tickets/" + ticket,
                            auth=fluxo_auth, proxies=None, cookies=tempcookie,
                            cert=None, timeout=None, verify=False, headers=getHeaders())
    if response.status_code == 404:
        details = 'Invaild case id'
        asin_info = 'Invaild case id'
    else:

        data = response.content.decode('utf-8')
        soup = BeautifulSoup(data.replace('&amp;', "&"), 'xml')

        if soup.details is not None:
            details = soup.details.string
        else:
            details = ''

        if soup.asin is not None:
            try:
                asin_info = soup.asin.string.strip()
            except:
                asin_info = ''
        else:
            asin_info = ''

    return details, asin_info


def get_asin(details, short_description):
    try:
        asin_text = re.search("(B0)[A-Z0-9]{8}|[0-9]{10}|[0-9]{9}[A-Z]{1}", details)
        str_asin = str(asin_text.group())
        return str_asin
    except:
        asin_text = re.search("(B0)[A-Z0-9]{8}|[0-9]{10}|[0-9]{9}[A-Z]{1}", short_description)
        str_asin = str(asin_text.group())
        return str_asin
    else:
        return ""


def get_info_from_details(details, asin_info):
    # shipping FC customer issue asin
    try:
        m = re.search("(Shipping FC:)()*[^\n]*", details)
        str_shipping_fc = str(m.group()).strip()
        fc_list = str_shipping_fc.split(":")[-1].strip().split(',')
    except:
        fc_list = []
    customer_issue = details[details.index('Additional description:'): details.index('wizardLog:')]
    asin = ''
    if asin_info in [None, '', '-'] or len(asin_info) < 10:
        try:
            m = re.search("(ASIN:)()*[^\n]*", details)
            str_asin = str(m.group()).strip()
            # print(str_asin)
            asin = str_asin.split(":")[-1].strip()
        except:
            asin = ''
    else:
        asin = asin_info.strip()
    return fc_list, customer_issue, asin


def isElementExist(driver, element):
    flag = True
    try:
        driver.find_element_by_xpath(element)
        return flag
    except:
        flag = False
        return flag


def get_inventory_data(asin):
    path = os.path.abspath(".")
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    driver = webdriver.Chrome(executable_path=path + r"\\chromedriver.exe",chrome_options=option)
    driver.implicitly_wait(15)
    driver.get('https://alaska-eu.amazon.com/index.html?viewtype=summaryview&use_scrollbars='
               '&fnsku_simple=' + asin + '&marketplaceid=3&merchantid=9&AvailData=Get+Availability+Data')
    try_time = 0
    while isElementExist(driver, "//div[@id='view']/table[2]") == False and try_time < 5:
        try_time += 1
        time.sleep(3)
    else:
        print('Error with Alaska!')
        return None, {}, 0
    # 操作alaska的表格，这个还是有一点价值的，网上查到的
    try:
        table = driver.find_element_by_xpath("//div[@id='view']/table[2]")
        table_rows = table.find_elements_by_tag_name('tr')  # 得到的是一个tr的list，一个元素应该是一整行，中间用空格隔开的
        table_list = []
        for rows in table_rows:
            table_list.append(str(rows.text).split(" "))  # 这应该是一个二位的列表，就是表里的每个元素都是一个列表
        table_list = table_list[2:-2]
        max_row = len(table_list) - 1  # 得到最后一行的下标索引
        max_fc = ""
        max_inventory = 0
        fc_list = {}
        i = 1

        if table_list[max_row][1] != 0:  # 判断有没有库存，就是总在库
            while i < max_row:  # 获得库存最多的库房，判断库房是因为有的库房不能发bincheckTT，要发邮件
                if int(table_list[i][1]) != 0:
                    fc_list[table_list[i][0]] = table_list[i][1]  # 因为我要FC和它所有的库存数，所以定义的是字典
                    if int(table_list[i][1]) > max_inventory:  # 取库存最多的FC
                        max_inventory = int(table_list[i][1])
                        max_fc = table_list[i][0]
                i += 1
            return max_fc, fc_list, table_list[max_row][1]  # 返回在库最多库房，有在库的库房及在库数，在库总数，方法可以返回多个值，实在tuple中的
        else:
            return None, {}, table_list[max_row][1]  # 没有在库返回的
    except:
        print('Error with Alaska!')
        return None, {}, 0


def create_ticket(data: dict) -> (int, str):
    payload = {"status": "Assigned",
               "impact": "5",
               "requester_login": getpass.getuser(),
               "requester_name": getpass.getuser()}
    payload.update(data)
    session = requests.session()
    response = session.post("https://ticket-api.amazon.com/tickets",
                                 data=payload, headers=getHeaders(),
                                 auth=fluxo_auth, proxies=None,
                                 cert=None, timeout=None, verify=False, cookies=tempcookie)
    return response


def cut_bincheck_tt(fc,asin,ticket,issue,fc_inventory_dic):
    cut_bincheck = create_ticket(data={
                    "short_description": "【%s】Bin Check for FnSku:%s Andon TT: %s"%(fc, asin, ticket),
                    "details": '''
Retail ICQAチームの皆さん

お疲れ様です。

Issue：%s

確認依頼：在庫品に同問題があるかご確認お願い致します。

ASIN: %s

Retail FC 在庫数：
FC: %s  在庫数: %s

どうぞ宜しくお願い致します。

RBS'''%(('\n' + issue),asin,fc,str(fc_inventory_dic[fc])),
                    "category": "RBS",
                    "type": "China-QA",
                    "item": "Automation",
                    "assigned_group": "DEQA-Automation",
                    "assigned_individual": "yingqin"}
    )
    times = 0
    status_code = cut_bincheck.status_code
    while times < 3 and status_code != 200:
        if status_code == 200:
            break
        time.sleep(5)
        times += 1
    ticket_number = None
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


def retry_put(ticket, correspondence):
    times = 0
    stas_code = 0
    while times < 3 and stas_code != 200:
        stas_code = update_correspondence(ticket, correspondence)
        if stas_code == 200:
            break
        time.sleep(5)
        times += 1
    return stas_code


def update_correspondence(ticket, correspondence):
    '''
    Update correspondence action
    :param ticket: each ticket
    :param correspondence: correspondence text
    :return: Y or response.status_code
    '''

    fluxo_auth = ("flx-JP-Andon", "JP-Andon-prod")
    # print("https://ticket-api.amazon.com/tickets/" + str(ticket))
    response = requests.put("https://ticket-api.amazon.com/tickets/" + str(ticket),
                            data={'correspondence': correspondence, "run_as_user": {getpass.getuser()}},
                            auth=fluxo_auth, proxies=None, cookies=tempcookie,
                            cert=None, timeout=None, verify=False, headers=getHeaders())
    return response.status_code


def run():
    warnings.filterwarnings("ignore")
    cmd = 'start mwinit -o'
    cmd_progress = False

    res = os.popen(cmd)
    res.read()
    tempcookie = get_tempcookie()
    ticket = '0528240527'
    details, asin_info = get_info_from_tickets(ticket)
    fc_list, customer_issue, asin = get_info_from_details(details, asin_info)
    max_fc, fc_inventory_dic, total_inventory = get_inventory_data(asin)
    tt_dic = {}
    # print(fc_inventory_dic)

    if total_inventory == 0:
        print('Asin %s has no inventory or error of Alaska!'%asin)
    else:
        correspondence = 'Inventory Information:\n\nInventory in Shipping FC:\n'
        is_shipping_fc_dic = {}
        for fc in fc_list:
            if fc in list(fc_inventory_dic.keys()):
                is_shipping_fc_dic[fc] = True
                tt_dic[fc] = 'https://tt.amazon.com/' + str(cut_bincheck_tt(fc, asin, ticket, customer_issue, fc_inventory_dic))
            else:
                is_shipping_fc_dic[fc] = False
        if tt_dic == {}:
            tt_dic[fc] = 'https://tt.amazon.com/' + str(
                cut_bincheck_tt(max_fc, asin, ticket, customer_issue, fc_inventory_dic))
        for k in is_shipping_fc_dic.keys():
            if is_shipping_fc_dic[k] == False:
                correspondence += k + '  :  ' + 'N' + '\n'
            else:
                correspondence += k + '  :  ' + 'Y' + '\n'
        correspondence += '\nFC inventory Count:\n'
        for k in fc_inventory_dic.keys():
            correspondence += k + '    ' + fc_inventory_dic[k] + '\n'
        retry_put(ticket, correspondence)

        correspondence = 'Bincheck TT:\n'
        for k in tt_dic.keys():
            correspondence += k + '  :  ' + tt_dic[k] + '\n'
        retry_put(ticket, correspondence)
        messagebox.showinfo('Message', 'Done')


if __name__=='__main__':
    run()