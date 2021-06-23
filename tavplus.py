#!/usr/bin/env python3

import json
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.utils import datetime
from openpyxl.styles import numbers
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import warnings
import pickle
import os
import sys

data_file = 'cards_data.pickle'

def save_prev_file(name, ext):
    old = ("_old."+ext).join(name.rsplit("."+ext, 1))
    if os.path.isfile(name):
        if os.path.isfile(old):
            os.remove(old)
        os.rename(name, old)

cards = {}
xactions = {}
try:
    with open(data_file, "rb") as f:
        cards = pickle.load(f)
        xactions = pickle.load(f)
        xls_file = pickle.load(f)
        output_xls_file = pickle.load(f)
except:
    pass

if 'xls_file' not in vars() or not xls_file:
    if len(sys.argv) != 3:
        print("""
This utility parses an XLSX file with a list of Yenot Bitan cards.
You can then get a report of all your purchases saved to an output XLSX file.
For the initial run please add the input and output xlsx files as arguments.
Once these are setm, they are saved in a pickle file,
together with any data pulled from the Yenot bitan website.
Only new cards added to the input XLSX file are read from the internet.
""")
        exit()
    xls_file  = sys.argv[1]
if 'output_xls_file' not in vars() or not output_xls_file:
    output_xls_file = sys.argv[2]
wb = load_workbook(filename = xls_file, data_only=True)
sheet = wb.worksheets[0]

with warnings.catch_warnings():
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    for i in range(sheet.min_row, sheet.max_row+1):
        id = sheet.cell(row=i, column=1).value
        code = sheet.cell(row=i, column=2).value
        if cards.get(id, -1) == 0:
            continue
        l = len(str(id))
        if l == 8:
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
        elif l == 16:
            id = str(id)
            bj = requests.post(url='https://www.shufersal.co.il/myshufersal/api/CardBalanceApi/GetCardBalanceAndTransactions', data={'cardNumber':str(id)+str(code)}, verify=False)
            b = json.loads(bj.text)
            if b['HasCard'] == True and b['InactiveCard'] == False:
                cards[id] = b['CurrentBalance']
            elif b['HasCard'] == False:
                continue
            for field in b['Transactions']:
                d = field['DateObject']
                dep = field['ActivityType'] == 'deposit'
                am = field['Amount']
                xactions[(d, id)] = {'name': field['BusinessName'], 'deposit':dep, 'sum':str(am) if dep else str(-am)}
        else:
            print("Unknown ID:", id)
            continue
        v =  cards.get(id, -1)
        if  v > 0:
            print("ID: {}-{} => {}".format(id, code,  cards[id]))
        elif v == 0:
            print('ID:', id)
        else:
            print("Unknown card:", id)

save_prev_file(data_file, "pickle")
with open(data_file, "wb") as f:
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
for ((d, id), fields) in sorted(xactions.items()):
    dt = datetime.from_ISO8601(d)
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
