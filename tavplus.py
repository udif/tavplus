#!/usr/bin/env python3

import json
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.utils import datetime as openpyxl_datetime
from openpyxl.styles import numbers
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import warnings
import pickle
import os
import sys
import argparse
import datetime
#import pprint

#pp = pprint.PrettyPrinter(indent=4)

cards = {}
xactions = {}

def save_prev_file(name, ext):
    old = ("_old."+ext).join(name.rsplit("."+ext, 1))
    if os.path.isfile(name):
        if os.path.isfile(old):
            os.remove(old)
        os.rename(name, old)

def handle_buyme(id, code):
    s = requests.Session()
    bj = s.get(
        url='https://buyme.co.il/',
        headers = {'User-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'},
    )
    if bj.status_code != 200:
        print("Unexpected error while updating", id, code)
        sys.exit(1)
    id = str(id)
    bj = s.get(url='https://buyme.co.il/siteapi/voucherBalance',
        params={'serialNumber':id, 'expiryDate':code},
        headers = {'User-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'},
        verify=False)
    if bj.status_code != 200:
        print("Unexpected error while updating", id, code)
        sys.exit(1)
    b = json.loads(bj.text)
    cards[id] = b['value']
    creation_date = 'T'.join(b['voucher']['crspackage']['created_at'].split(' ')) + ".000"
    xactions[(creation_date, id)] = {'name': b['title'], 'deposit': True, 'sum': float(b['originalValue'])}
    for field in b['realizations']:
        d = 'T'.join(field['date'].split(' ')) + ".000"
        xactions[(d, id)] = {'name': field['redeemer'], 'deposit': False, 'sum': float(field['amount'])}

def handle_ybitan(id, code):
    bj = requests.post(url='https://tavplus.mltp.co.il/multipassapi/getbudget.php', data={'cardid':id}, verify=False)
    b = json.loads(bj.text)
    if b['ResultMessage'] != '' and b['ResultId'] == 0:
        cards[id] = int(b['UpdatedBugdet']) / 100.0
    bj = requests.post(url='https://tavplus.mltp.co.il/multipassapi/GetLastTransactions.php', data={'cardid':id}, verify=False)
    b = json.loads(bj.text)
    for field in b['data']:
        d = field['date']
        dep = field['LoadActualSum'] != ''
        xactions[(d, id)] = {'name': field['SupplierName'], 'deposit':dep, 'sum': field['LoadActualSum'] if dep else field['ApprovedSum']}

def handle_tav_zahav(id, code):
    id = str(id)
    bj = requests.post(url='https://www.shufersal.co.il/myshufersal/api/CardBalanceApi/GetCardBalanceAndTransactions', data={'cardNumber':str(id)+str(code)}, verify=False)
    b = json.loads(bj.text)
    if b['HasCard'] == True and b['InactiveCard'] == False:
        cards[id] = b['CurrentBalance']
    elif b['HasCard'] == False:
        return
    for field in b['Transactions']:
        d = field['DateObject']
        dep = field['ActivityType'] == 'deposit'
        am = field['Amount']
        xactions[(d, id)] = {'name': field['BusinessName'], 'deposit':dep, 'sum':str(am) if dep else str(-am)}

def detect_paytment_method(id, code):
    l1 = len(str(id))
    l2 = len(str(code))
    if (l1, l2) == (16,10):
        handle_buyme(id, code)
    elif (l1, l2) == (8, 4):
        handle_ybitan(id, code)
    elif (l1, l2) == (16, 3):
        handle_tav_zahav(id, code)
    else:
        print("Unknown ID:", id)
        return False
    return True

parser = argparse.ArgumentParser(description="""
This utility parses an XLSX file with a list of Prepaid cards.
You can then get a report of all your purchases saved to an output XLSX file.
Only new cards added to the input XLSX file are read from the internet.
Currently supported:
Yenot Bitan
Tav Hazahav (non-7215)
Buy Me
TODO:
Isracard
Max
""")
#parser.add_argument("-v", "--verbose", help="increase output verbosity", action="store_true", type=int)
#parser.add_argument("-q", "--quiet", action="store_true", help="quiet mode - only warn if all dupes or if more than one orig")
parser.add_argument("-f", "--data_file", metavar="picklefile", type=str, nargs=1, help="name of pickle file to store program state", default='cards_data.pickle')
parser.add_argument("-d", "--delete", metavar="card-ID", type=str, nargs=1, help="ID of card to remove")
parser.add_argument("-i", "--input", metavar="input-cards.xls", type=str, help="Name of XLSX file containing cards")
parser.add_argument("-o", "--output", metavar="output-transactions.xls", type=str, help="Name of XLSX file containing all transactions")
args = parser.parse_args()

if False:
    import logging

    # These two lines enable debugging at httplib level (requests->urllib3->http.client)
    # You will see the REQUEST, including HEADERS and DATA, and RESPONSE with HEADERS but without DATA.
    # The only thing missing will be the response.body which is not logged.
    try:
        import http.client as http_client
    except ImportError:
        # Python 2
        import httplib as http_client
    http_client.HTTPConnection.debuglevel = 1

    # You must initialize logging, otherwise you'll not see debug output.
    logging.basicConfig()
    logging.getLogger().setLevel(logging.DEBUG)
    requests_log = logging.getLogger("requests.packages.urllib3")
    requests_log.setLevel(logging.DEBUG)
    requests_log.propagate = True

try:
    with open(args.data_file, "rb") as f:
        cards = pickle.load(f)
        xactions = pickle.load(f)
        xls_file = pickle.load(f)
        output_xls_file = pickle.load(f)
except:
    pass

if args.input:
    xls_file = args.input
if args.output:
    output_xls_file = args.output

if args.delete:
    id = args.delete[0]
    if cards.get(id, None):
        del cards[id]
        print("Deleting", id)
    else:
        print("{} already deleted in cards[]".format(id))

if not xls_file:
    print("No input file defined")
    sys.exit(1)
if not output_xls_file:
    print("No output file defined")
    sys.exit(1)

wb = load_workbook(filename = xls_file, data_only=True)
sheet = wb.worksheets[0]

with warnings.catch_warnings():
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    for i in range(sheet.min_row, sheet.max_row+1):
        id = sheet.cell(row=i, column=1).value
        code = sheet.cell(row=i, column=2).value
        if type(code) is datetime.datetime:
            code = code.strftime("%Y-%m-%d")
        if id == args.delete:
            continue
        if cards.get(id, -1) == 0:
            continue
        if not detect_paytment_method(id, code):
            continue
        v =  cards.get(id, -1)
        if  v > 0:
            print("ID: {}-{} => {}".format(id, code,  cards[id]))
        elif v == 0:
            print('ID:', id)
        else:
            print("Unknown card:", id)

if args.delete:
    x2 = xactions.copy()
    for (d, id) in x2.keys():
        if id == args.delete[0]:
            print("Deleting", d, id)
            del xactions[(d, id)]
    x2 = None

save_prev_file(args.data_file, "pickle")
with open(args.data_file, "wb") as f:
    pickle.dump(cards, f)
    pickle.dump(xactions, f)
    pickle.dump(xls_file, f)
    pickle.dump(output_xls_file, f)
    print("Saved", len(xactions.keys()), "transactions")

wb = Workbook()
ws = wb.active
ws.title = "קניות"
ws.cell(row=1, column=1).value = "תאריך"
ws.cell(row=1, column=2).value = "שעה"
ws.cell(row=1, column=3).value = "מיקום"
ws.cell(row=1, column=4).value = "סכום"
ws.cell(row=1, column=5).value = "טעינה"
ws.cell(row=1, column=6).value = "כרטיס"
row = 2
for (d, id) in sorted(xactions.keys(), key=lambda item:item[0]):
    fields = xactions[(d, id)]
    dt = openpyxl_datetime.from_ISO8601(d)
    ws.cell(row=row, column=1).value = dt.date()
    ws.cell(row=row, column=2).value = dt.time()
    v = fields['sum']
    dep = fields['deposit']
    if not v or v == '':
        v = 0.0
        #print(d, id, fields)
    ws.cell(row=row, column=3).value = fields['name']
    ws.cell(row=row, column=4).value = float(v) if not dep else 0
    ws.cell(row=row, column=4).number_format = numbers.FORMAT_NUMBER_00
    ws.cell(row=row, column=5).value = float(v) if dep else 0
    ws.cell(row=row, column=5).number_format = numbers.FORMAT_NUMBER_00
    ws.cell(row=row, column=6).value = id
    row += 1
save_prev_file(output_xls_file, "xlsx")
wb.save(output_xls_file)

    

#print(sheet_ranges['D18'].value)
