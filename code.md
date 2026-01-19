# findchips_python_scrapy

import bs4, json, requests
import pandas as pd
from bs4 import BeautifulSoup
from decimal import Decimal
from os import listdir
from openpyxl import load_workbook

chrome_version = listdir(r'C:\Program Files\Google\Chrome\Application')[0].split('.')[0]
header = {'User-Agent': f'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{chrome_version}.0.0.0 Safari/537.36'}

df = pd.read_excel('TestBOM.xlsx')
mpn, mfg = df['MPN'].tolist(), df['MFG'].tolist()

final_result, reason= [], []
for i, (mpn, mfg) in enumerate(zip(mpn, mfg), 1):
    url = f'https://www.findchips.com/search/{mpn}'
    r = requests.get(url, headers=header)
    soup = BeautifulSoup(r.text, 'html.parser')

    agents = soup.select('h2[class=distributor-title]')
    tbodys = soup.select('tbody')

    allowed_tags, tbodys_flag, tbodys_lower = {'Authorized', 'Direct'}, [], []
    for j, (name_info, tbody) in enumerate(zip(agents, tbodys), 1):
        name = name_info.select_one('img').get('alt').replace(' logo', '')
        info = name_info.select_one('span.other-disti-details').text.split()

        if set(info).intersection(allowed_tags):
            trs = tbody.select('tr')
            trs_flag, trs_lower = [], []
            for k, tr in enumerate (trs, 1):
                td_pn = tr.select_one('a').text.strip()
                td_mfg = tr.select_one('td[class=td-mfg]').text.strip()
                if mpn == td_pn and mfg == td_mfg:
                    tr_data_price = tr.get('data-price').strip()
                    if tr_data_price != '[]':
                        tdp_list = json.loads(tr_data_price)
                        tr_lower = []
                        for ele in tdp_list:
                            unit_price = Decimal(ele[2])
                            tr_lower.append(unit_price)
                        tr_lowest_price = min(tr_lower)
                        trs_lower.append(tr_lowest_price)
                        trs_flag.append(True)
                    else:
                        trs_flag.append(False)
                        continue
                else:
                    trs_flag.append(False)
                    continue
            if any(trs_flag):
                tbodys_lower.append(min(trs_lower))
                tbodys_flag.append(True)
        else:
            continue
    if any(tbodys_flag):    
        final_lowest_price = min(tbodys_lower)
        if final_lowest_price == 0:
            final_result.append('none')
            reason.append('0 price exists')
        else:
            final_result.append(final_lowest_price)
            reason.append('normal')
        # print(f'{i}, final_lowest_price: {final_lowest_price}')
    else:
        final_lowest_price = 'none'
        final_result.append(final_lowest_price)
        reason.append('does not match REQ')
    print(f'{i}, final_lowest_price: {final_lowest_price}')
print(f'final_result: {final_result}')
print(f'reason: {reason}')

df_1 = pd.DataFrame(final_result)
df_2 = pd.DataFrame(reason)
file_name = 'TestBOM.xlsx'
book = load_workbook(file_name)
sheet = book['Sheet1']
start_row, start_col = 2, 5

for r_idx, (row_1, row_2) in enumerate(zip(df_1.values, df_2.values)):
    for c_idx, (value_1, value_2) in enumerate(zip(row_1, row_2)):
        sheet.cell(row=start_row + r_idx, column=start_col + c_idx, value=value_1)
        sheet.cell(row=start_row + r_idx, column=start_col + c_idx + 1, value=value_2)
book.save(file_name)
