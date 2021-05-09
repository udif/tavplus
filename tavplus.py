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
        if cards.get(id, -1) == 0:
            continue
        bj = requests.post(url='https://tavplus.mltp.co.il/multipassapi/getbudget.php', data={'cardid':id}, verify=False)
        b = json.loads(bj.text)
        if b['ResultMessage'] != '' and b['ResultId'] == 0:
            cards[id] = int(b['UpdatedBugdet']) / 100.0
        if cards[id] > 0:
            print("ID:", id, '=>',  cards[id])
        else:
            print('ID:', id)
        bj = requests.post(url='https://tavplus.mltp.co.il/multipassapi/GetLastTransactions.php', data={'cardid':id}, verify=False)
        b = json.loads(bj.text)
        for field in b['data']:
            d = field.pop('date')
            xactions[(d, id)] = field

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
ws.cell(row=1, column=3).value = "סכום"
ws.cell(row=1, column=4).value = "כרטיס"
row = 2
for ((d, id), fields) in sorted(xactions.items()):
    dt = datetime.from_ISO8601(d)
    ws.cell(row=row, column=1).value = dt
    #ws.cell(row=row, column=2).value = d[1]
    v = fields['ApprovedSum']
    if not v or v == '':
        v = 0.0
        print(d, id, fields)
    ws.cell(row=row, column=3).value = float(v)
    ws.cell(row=row, column=3).number_format = numbers.FORMAT_NUMBER_00
    ws.cell(row=row, column=4).value = id
    row += 1
save_prev_file(output_xls_file, "xlsx")
wb.save(output_xls_file)

    

#print(sheet_ranges['D18'].value)
